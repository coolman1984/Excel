import * as React from "react";
import {
    DataEditor,
    GridCellKind,
    CompactSelection,
    type DataEditorRef,
    type DrawCellCallback,
    type DrawHeaderCallback,
    type FillPatternEventArgs,
    type GridCell,
    type GridColumn,
    type GridSelection,
    type Item,
    type Rectangle,
    type TextCell,
} from "@glideapps/glide-data-grid";
import "@glideapps/glide-data-grid/dist/index.css";
import { call } from "../bridge";
import { fitCellText, bodyRow, frozenCount } from "./display";
import { RowStore } from "../rowStore";
import { ROW_HEIGHT, FROZEN_LINE, FILTER_BLUE, scaledTheme } from "../theme";
import type { CellData, CellStyle, SheetInfo, Condition } from "../types";

// The grid (plan 2.6): row-number column, letters, selection, zoom, editing hooks,
// the fill handle and a custom drawCell for styled cells. Frozen rows are drawn
// inside the header.

/** Extra fields we hang on a TextCell so drawCell can paint Excel styles. */
export interface ExcelCell extends TextCell {
    excel?: CellStyle;
    excelKind?: number;
    isRowNumber?: boolean;
    filtered?: boolean;
}

export interface SelRange {
    r1: number;
    c1: number;
    r2: number;
    c2: number;
}

export interface GridHandle {
    refresh(): void;
    invalidateCells(cells: number[][]): void;
    scrollTo(viewRow: number, col?: number): void;
    setSelection(r1: number, c1: number, r2: number, c2: number): void;
    getSelection(): SelRange;
    store: RowStore;
    focus(): void;
}

interface Props {
    sheet: SheetInfo;
    viewRows: number;
    filterActive: boolean;
    filters: Record<number, Condition>;
    zoom: number;
    showGridlines: boolean;
    showHeadings: boolean;
    freeze: { rows: number; cols: number };
    onStoreReady: (store: RowStore) => void;
    onSelectionChange: (sel: SelRange) => void;
    onCellEdited: (viewRow: number, col: number, text: string) => void;
    onDeleteRanges: (ranges: number[][]) => void;
    onFill: (from: number[], to: number[]) => void;
    onHeaderMenu: (col: number, x: number, y: number) => void;
    onPaint: (ms: number) => void;
    /** Bumped by the parent whenever the view (filter/sort/open) changed. */
    viewKey: number;
}

const VISIBLE_MARGIN = 40; // roughly two visible screens above and below

/** Excel column letters for a 0-based column index. */
export function letters(col0: number): string {
    let n = col0 + 1;
    let s = "";
    while (n > 0) {
        const rem = (n - 1) % 26;
        s = String.fromCharCode(65 + rem) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

function alignToContent(align: number): "left" | "right" | "center" {
    return align === 1 ? "right" : align === 2 ? "center" : "left";
}

const loadingCell: GridCell = { kind: GridCellKind.Loading, allowOverlay: false };

function useSize(ref: React.RefObject<HTMLDivElement | null>): { width: number; height: number } {
    const [size, setSize] = React.useState({ width: 800, height: 400 });
    React.useEffect(() => {
        const el = ref.current;
        if (!el) return;
        const measure = () => setSize({ width: el.clientWidth, height: el.clientHeight });
        measure();
        const ro = new ResizeObserver(measure);
        ro.observe(el);
        return () => ro.disconnect();
    }, [ref]);
    return size;
}

/** Paint styled text ourselves: Glide's default renderer resets the font. */
function paintText(ctx: CanvasRenderingContext2D, rect: Rectangle, cell: ExcelCell, theme: { baseFontStyle: string; fontFamily: string; textDark: string }) {
    const st = cell.excel;
    ctx.save();
    ctx.beginPath(); ctx.rect(rect.x, rect.y, rect.width, rect.height); ctx.clip();
    if (st?.fill) { ctx.fillStyle = st.fill; ctx.fillRect(rect.x, rect.y, rect.width, rect.height); }
    ctx.font = `${st?.italic ? "italic " : ""}${st?.bold ? "bold " : ""}${theme.baseFontStyle} ${theme.fontFamily}`;
    ctx.fillStyle = st?.fontColor ?? cell.themeOverride?.textDark ?? theme.textDark;
    ctx.textBaseline = "middle";
    ctx.textAlign = cell.contentAlign ?? "left";
    const text = fitCellText(cell.displayData, cell.excelKind ?? 2, rect.width, s => ctx.measureText(s).width, st?.numberFormat);
    const x = cell.contentAlign === "right" ? rect.x + rect.width - 4 : cell.contentAlign === "center" ? rect.x + rect.width / 2 : rect.x + 4;
    ctx.fillText(text, x, rect.y + rect.height / 2);
    if (st?.underline) {
        const w = ctx.measureText(text).width;
        const left = cell.contentAlign === "right" ? x - w : cell.contentAlign === "center" ? x - w / 2 : x;
        ctx.fillRect(left, rect.y + rect.height - 3, w, 1);
    }
    if (st?.border && st.border !== "none") {
        ctx.strokeStyle = "#000"; ctx.lineWidth = 1; ctx.beginPath();
        if (st.border === "bottom") { ctx.moveTo(rect.x, rect.y + rect.height - .5); ctx.lineTo(rect.x + rect.width, rect.y + rect.height - .5); }
        else ctx.rect(rect.x + .5, rect.y + .5, rect.width - 1, rect.height - 1);
        ctx.stroke();
    }
    ctx.restore();
}

export const Grid = React.forwardRef<GridHandle, Props>(function Grid(props, ref) {
    const {
        sheet,
        viewRows,
        filterActive,
        zoom,
        showGridlines,
        showHeadings,
        freeze,
        onSelectionChange,
        onCellEdited,
        onDeleteRanges,
        onFill,
        onHeaderMenu,
        onPaint,
    } = props;

    const containerRef = React.useRef<HTMLDivElement | null>(null);
    const editorRef = React.useRef<DataEditorRef | null>(null);
    const colsRef = React.useRef(1);
    const lastVisibleRef = React.useRef<Rectangle>({ x: 0, y: 0, width: 0, height: 20 });
    const pendingScrollRef = React.useRef<{ viewRow: number; t0: number } | null>(null);
    const damageCells = React.useRef<(x: number, y: number, w: number, h: number) => void>(() => {});
    const { width, height } = useSize(containerRef);

    const store = React.useMemo(
        () =>
            new RowStore(call, (from, count) => {
                damageCells.current(0, from, colsRef.current, count);
                if (from === 0) redrawHeader(v => v + 1);
                const p = pendingScrollRef.current;
                if (p && from <= p.viewRow && p.viewRow < from + count) {
                    pendingScrollRef.current = null;
                    requestAnimationFrame(() => onPaint(performance.now() - p.t0));
                }
            }),
        [onPaint]
    );

    React.useEffect(() => {
        props.onStoreReady(store);
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [store]);

    const z = zoom / 100;
    const rowHeight = Math.round(ROW_HEIGHT * z);
    const headerBand = Math.round(ROW_HEIGHT * z);
    const frozen = frozenCount(freeze.rows, viewRows);
    const headerHeight = headerBand * (1 + frozen);
    const [headerVersion, redrawHeader] = React.useState(0);

    const columns: GridColumn[] = React.useMemo(() => {
        const dataCols: GridColumn[] = [];
        for (let k = 0; k < sheet.cols; k++) {
            dataCols.push({
                id: "c" + k,
                title: letters(sheet.firstCol - 1 + k),
                width: Math.round((sheet.colWidthsPx[k] ?? 64) * z),
                hasMenu: filterActive,
            } as GridColumn);
        }
        const rn: GridColumn = { id: "__rn", title: "", width: Math.round(52 * z) } as GridColumn;
        const all = [rn, ...dataCols];
        colsRef.current = all.length;
        return all;
    }, [sheet, z, filterActive]);

    damageCells.current = (x, y, w, h) => {
        const cells: { cell: Item }[] = [];
        const x1 = Math.min(x + w, colsRef.current);
        if (frozen > 0 && y < frozen) for (let c = x; c < x1; c++) cells.push({ cell: [c, -1] });
        for (let r = Math.max(y, frozen); r < y + h; r++) for (let c = x; c < x1; c++) cells.push({ cell: [c, r - frozen] });
        editorRef.current?.updateCells(cells);
    };

    const refresh = React.useCallback(() => {
        store.clear();
        const v = lastVisibleRef.current;
        store.ensure(0, frozen - 1);
        store.ensure(v.y + frozen - VISIBLE_MARGIN, v.y + frozen + v.height + VISIBLE_MARGIN);
        damageCells.current(0, v.y, colsRef.current, Math.max(v.height, 40));
    }, [store, frozen]);

    React.useEffect(() => {
        store.setSheet(sheet.index, viewRows);
        refresh();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [sheet.index, viewRows, props.viewKey]);

    const [sel, setSel] = React.useState<GridSelection>({
        columns: CompactSelection.empty(),
        rows: CompactSelection.empty(),
    });
    const selRef = React.useRef<SelRange>({ r1: 0, c1: 0, r2: 0, c2: 0 });

    const computeRange = React.useCallback(
        (s: GridSelection): SelRange => {
            const cur = s.current;
            if (cur) {
                const r = cur.range;
                return { r1: bodyRow(r.y, frozen), c1: Math.max(0, r.x - 1), r2: bodyRow(r.y + r.height - 1, frozen), c2: Math.max(0, r.x + r.width - 2) };
            }
            const rowsArr = s.rows.toArray();
            const colsArr = s.columns.toArray();
            if (rowsArr.length > 0) return { r1: bodyRow(rowsArr[0], frozen), c1: 0, r2: bodyRow(rowsArr[rowsArr.length - 1], frozen), c2: sheet.cols - 1 };
            if (colsArr.length > 0)
                return { r1: 0, c1: colsArr[0] - 1, r2: Math.max(0, viewRows - 1), c2: colsArr[colsArr.length - 1] - 1 };
            return { r1: 0, c1: 0, r2: 0, c2: 0 };
        },
        [sheet.cols, viewRows, frozen]
    );

    const onGridSelectionChange = React.useCallback(
        (s: GridSelection) => {
            setSel(s);
            selRef.current = computeRange(s);
            onSelectionChange(selRef.current);
        },
        [computeRange, onSelectionChange]
    );

    React.useImperativeHandle(
        ref,
        (): GridHandle => ({
            refresh,
            invalidateCells: (cells: number[][]) => {
                if (!cells || cells.length === 0) return;
                const rows = new Set<number>();
                for (const c of cells) rows.add(c[0]);
                store.invalidateRows(rows);
                for (const r of rows) store.ensure(r, r);
                for (const r of rows) damageCells.current(0, r, colsRef.current, 1);
            },
            scrollTo: (viewRow: number, col = 1) => {
                const pendingStart = performance.now();
                pendingScrollRef.current = { viewRow, t0: pendingStart };
                store.ensure(viewRow, viewRow + 40);
                editorRef.current?.scrollTo(col, Math.max(0, viewRow - frozen), "vertical");
                if (store.getRow(viewRow)) requestAnimationFrame(() => {
                    pendingScrollRef.current = null; onPaint(performance.now() - (pendingStart));
                });
            },
            setSelection: (r1, c1, r2, c2) => {
                const x = Math.min(c1, c2) + 1;
                const y = Math.max(0, Math.min(r1, r2) - frozen);
                const w = Math.abs(c2 - c1) + 1;
                const next: GridSelection = {
                    columns: CompactSelection.empty(), rows: CompactSelection.empty(),
                    current: Math.max(r1, r2) < frozen ? undefined : { cell: [x, y], range: { x, y, width: w, height: Math.max(1, Math.max(r1, r2) - Math.max(frozen, Math.min(r1, r2)) + 1) }, rangeStack: [] },
                };
                setSel(next);
                selRef.current = { r1, c1, r2, c2 };
                onSelectionChange(selRef.current); redrawHeader(v => v + 1);
            },
            getSelection: () => selRef.current,
            store,
            focus: () => editorRef.current?.focus(),
        }),
        [refresh, store, computeRange, onSelectionChange, frozen, onPaint]
    );

    const toExcelCell = React.useCallback((cd: CellData): ExcelCell => {
        const st = cd.style;
        const align = st?.align ?? cd.align;
        const cell: ExcelCell = {
            kind: GridCellKind.Text,
            data: cd.formula ?? cd.text,
            displayData: cd.warn ? cd.text + " \u26A0" : cd.text,
            allowOverlay: true,
            contentAlign: alignToContent(align),
            excel: st,
            excelKind: cd.kind,
            themeOverride: cd.color ? { textDark: cd.color } : undefined,
        };
        return cell;
    }, []);

    const rowNumberCell = React.useCallback(
        (rowNumber: number): ExcelCell => ({
            kind: GridCellKind.Text,
            data: String(rowNumber),
            displayData: String(rowNumber),
            allowOverlay: false,
            contentAlign: "right",
            isRowNumber: true,
            filtered: filterActive,
            themeOverride: { textDark: filterActive ? FILTER_BLUE : "#444444", bgCell: "#F3F3F3" },
        }),
        [filterActive]
    );

    const getCellContent = React.useCallback(
        (cell: Item): GridCell => {
            const [col, body] = cell;
            const row = bodyRow(body, frozen);
            const wire = store.getRow(row);
            if (wire === undefined) {
                store.ensure(row, row);
                return loadingCell;
            }
            if (col === 0) return rowNumberCell(wire.rowNumber);
            const cd = wire.cells[col - 1];
            if (cd === undefined) return loadingCell;
            return toExcelCell(cd);
        },
        [store, rowNumberCell, toExcelCell, frozen]
    );

    const drawCell = React.useCallback<DrawCellCallback>((args, done) => {
        const cell = args.cell as ExcelCell;
        if (cell.kind !== GridCellKind.Text) { done(); return; }
        paintText(args.ctx, args.rect, cell, args.theme);
    }, []);

    const drawHeader = React.useCallback<DrawHeaderCallback>(
        (a, _drawContent) => {
            const { ctx, rect, column } = a;
            const t = a.theme as any;
            ctx.fillStyle = t.bgHeader;
            ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
            const band = headerBand;
            if (showHeadings) {
                ctx.save();
                ctx.fillStyle = t.textHeader;
                ctx.font = t.headerFontFull ?? `${t.headerFontStyle} ${t.fontFamily}`;
                ctx.textAlign = "center";
                ctx.textBaseline = "middle";
                ctx.fillText(String(column.title ?? ""), rect.x + rect.width / 2, rect.y + band / 2);
                ctx.restore();
            }
            for (let i = 0; i < frozen; i++) {
                const top = rect.y + band + i * rowHeight;
                ctx.fillStyle = i % 2 === 0 ? t.bgCell : t.bgCellMedium;
                ctx.fillRect(rect.x, top, rect.width, rowHeight);
                const wire = store.getRow(i);
                const cd = wire?.cells[a.columnIndex - 1];
                const cell = a.columnIndex === 0 && wire ? rowNumberCell(wire.rowNumber) : cd ? toExcelCell(cd) : undefined;
                if (cell) paintText(ctx, { x: rect.x, y: top, width: rect.width, height: rowHeight }, cell, a.theme);
                const selection = selRef.current;
                if (i >= selection.r1 && i <= selection.r2 && a.columnIndex - 1 >= selection.c1 && a.columnIndex - 1 <= selection.c2) {
                    ctx.save(); ctx.strokeStyle = "#107C41"; ctx.strokeRect(rect.x + 1, top + 1, rect.width - 2, rowHeight - 2); ctx.restore();
                }
            }

            ctx.save();
            if (freeze.rows > 0) {
                ctx.strokeStyle = FROZEN_LINE;
                ctx.beginPath();
                ctx.moveTo(rect.x, rect.y + band - 0.5);
                ctx.lineTo(rect.x + rect.width, rect.y + band - 0.5);
                ctx.stroke();
            }
            ctx.strokeStyle = t.borderColor;
            ctx.beginPath();
            ctx.moveTo(rect.x, rect.y + rect.height - 0.5);
            ctx.lineTo(rect.x + rect.width, rect.y + rect.height - 0.5);
            ctx.stroke();
            ctx.restore();
            void _drawContent;
        },
        [headerBand, rowHeight, frozen, showHeadings, store, rowNumberCell, toExcelCell, headerVersion]
    );

    const getCellsForSelection = React.useCallback(
        (range: Rectangle): GridCell[][] => {
            const emptyCell: TextCell = { kind: GridCellKind.Text, data: "", displayData: "", allowOverlay: false };
            const out: GridCell[][] = [];
            for (let y = range.y; y < range.y + range.height; y++) {
                const wire = store.getRow(bodyRow(y, frozen));
                const rowCells: GridCell[] = [];
                for (let x = range.x; x < range.x + range.width; x++) {
                    if (wire === undefined) {
                        rowCells.push(emptyCell);
                    } else if (x === 0) {
                        rowCells.push(rowNumberCell(wire.rowNumber));
                    } else {
                        const cd = wire.cells[x - 1];
                        rowCells.push(cd ? toExcelCell(cd) : emptyCell);
                    }
                }
                out.push(rowCells);
            }
            return out;
        },
        [store, rowNumberCell, toExcelCell, frozen]
    );

    const handleCellEdited = React.useCallback(
        (cell: Item, newValue: GridCell) => {
            const [col, row] = cell;
            if (col <= 0) return;
            onCellEdited(bodyRow(row, frozen), col - 1, (newValue as TextCell).data);
        },
        [onCellEdited, frozen]
    );

    const handleDelete = React.useCallback(
        (selection: GridSelection) => {
            const r = selection.current ? computeRange(selection) : selRef.current;
            onDeleteRanges([[r.r1, r.c1, r.r2, r.c2]]);
            return false;
        },
        [computeRange, onDeleteRanges]
    );

    const handleFill = React.useCallback(
        (e: FillPatternEventArgs) => {
            const src = e.patternSource;
            const dst = e.fillDestination;
            const from = [bodyRow(src.y, frozen), Math.max(0, src.x - 1), bodyRow(src.y + src.height - 1, frozen), Math.max(0, src.x + src.width - 2)];
            const to = [bodyRow(dst.y, frozen), Math.max(0, dst.x - 1), bodyRow(dst.y + dst.height - 1, frozen), Math.max(0, dst.x + dst.width - 2)];
            onFill(from, to);
            e.preventDefault();
        },
        [onFill, frozen]
    );

    const handleCellClicked = React.useCallback(
        (cell: Item) => {
            if (cell[0] === 0) {
                const row = cell[1];
                const next: GridSelection = {
                    columns: CompactSelection.empty(),
                    rows: CompactSelection.empty().add(row),
                    current: undefined,
                };
                setSel(next);
                selRef.current = computeRange(next);
                onSelectionChange(selRef.current);
            }
        },
        [computeRange, onSelectionChange]
    );

    const handleHeaderMenu = React.useCallback(
        (col: number, bounds: Rectangle) => {
            if (col > 0) onHeaderMenu(col - 1, bounds.x, bounds.y);
        },
        [onHeaderMenu]
    );

    const onVisible = React.useCallback(
        (range: Rectangle) => {
            lastVisibleRef.current = range;
            store.ensure(0, frozen - 1);
            store.ensure(range.y + frozen, range.y + frozen + range.height);
            store.ensure(range.y + frozen - VISIBLE_MARGIN, range.y + frozen + range.height + VISIBLE_MARGIN);
        },
        [store, frozen]
    );

    const theme = React.useMemo(() => {
        const t = scaledTheme(zoom);
        if (!showGridlines) {
            t.borderColor = "transparent";
            t.horizontalBorderColor = "transparent";
        }
        return t;
    // Glide's blit check ignores drawHeader identity. A new theme forces the frozen band to repaint.
    }, [zoom, showGridlines, headerVersion]);

    return (
        <div className="gridHost" ref={containerRef}>
            <DataEditor
                ref={editorRef}
                width={width}
                height={height}
                columns={columns}
                rows={Math.max(0, viewRows - frozen)}
                rowHeight={rowHeight}
                headerHeight={headerHeight}
                freezeColumns={1 + freeze.cols}
                rowMarkers="none"
                theme={theme}
                gridSelection={sel}
                onGridSelectionChange={onGridSelectionChange}
                getCellContent={getCellContent}
                getCellsForSelection={getCellsForSelection}
                onVisibleRegionChanged={onVisible}
                onCellEdited={handleCellEdited}
                onDelete={handleDelete}
                onFillPattern={handleFill}
                onPaste={() => false}
                onCellClicked={handleCellClicked}
                onHeaderClicked={(col, e) => {
                    const y = e.localEventY;
                    if (y >= headerBand && col > 0) {
                        const row = Math.min(frozen - 1, Math.floor((y - headerBand) / rowHeight));
                        selRef.current = { r1: row, c1: col - 1, r2: row, c2: col - 1 };
                        setSel({ columns: CompactSelection.empty(), rows: CompactSelection.empty() });
                        onSelectionChange(selRef.current); redrawHeader(v => v + 1);
                    }
                }}
                onHeaderMenuClick={handleHeaderMenu}
                drawCell={drawCell}
                drawHeader={drawHeader}
                smoothScrollX
                smoothScrollY
                rangeSelect="multi-rect"
                columnSelect="multi"
                rowSelect="multi"
                fillHandle
                keybindings={{ copy: false, cut: false, paste: false, delete: false, selectAll: false,
                    goToFirstCell: false, goToLastCell: false, goToFirstRow: false, goToLastRow: false }}
            />
        </div>
    );
});






