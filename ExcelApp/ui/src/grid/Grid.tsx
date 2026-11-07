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
import { RowStore } from "../rowStore";
import { ROW_HEIGHT, FROZEN_LINE, FILTER_BLUE, scaledTheme } from "../theme";
import type { CellData, CellStyle, SheetInfo, Condition } from "../types";

// The grid (plan 2.6): row-number column, letters, selection, zoom, editing hooks,
// the fill handle and a custom drawCell for styled cells. Frozen rows are drawn
// inside the header.

/** Extra fields we hang on a TextCell so drawCell can paint Excel styles. */
export interface ExcelCell extends TextCell {
    excel?: CellStyle;
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

const VISIBLE_MARGIN = 400; // two screens of 200-row blocks above and below

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
    const headerHeight = headerBand * (1 + freeze.rows);

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
        for (let r = y; r < y + h; r++) for (let c = x; c < x1; c++) cells.push({ cell: [c, r] });
        editorRef.current?.updateCells(cells);
    };

    const refresh = React.useCallback(() => {
        store.clear();
        const v = lastVisibleRef.current;
        store.ensure(v.y - VISIBLE_MARGIN, v.y + v.height + VISIBLE_MARGIN);
        damageCells.current(0, v.y, colsRef.current, Math.max(v.height, 40));
    }, [store]);

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
                return { r1: r.y, c1: Math.max(0, r.x - 1), r2: r.y + r.height - 1, c2: Math.max(0, r.x + r.width - 2) };
            }
            const rowsArr = s.rows.toArray();
            const colsArr = s.columns.toArray();
            if (rowsArr.length > 0) return { r1: rowsArr[0], c1: 0, r2: rowsArr[rowsArr.length - 1], c2: sheet.cols - 1 };
            if (colsArr.length > 0)
                return { r1: 0, c1: colsArr[0] - 1, r2: Math.max(0, viewRows - 1), c2: colsArr[colsArr.length - 1] - 1 };
            return { r1: 0, c1: 0, r2: 0, c2: 0 };
        },
        [sheet.cols, viewRows]
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
                for (const r of rows) damageCells.current(0, r, colsRef.current, 1);
            },
            scrollTo: (viewRow: number, col = 1) => {
                pendingScrollRef.current = { viewRow, t0: performance.now() };
                editorRef.current?.scrollTo(col, viewRow, "vertical");
            },
            setSelection: (r1, c1, r2, c2) => {
                const x = Math.min(c1, c2) + 1;
                const y = Math.min(r1, r2);
                const w = Math.abs(c2 - c1) + 1;
                const h = Math.abs(r2 - r1) + 1;
                const next: GridSelection = {
                    columns: CompactSelection.empty(),
                    rows: CompactSelection.empty(),
                    current: { cell: [x, y], range: { x, y, width: w, height: h }, rangeStack: [] },
                };
                setSel(next);
                selRef.current = computeRange(next);
                onSelectionChange(selRef.current);
            },
            getSelection: () => selRef.current,
            store,
            focus: () => editorRef.current?.focus(),
        }),
        [refresh, store, computeRange, onSelectionChange]
    );

    const toExcelCell = React.useCallback((cd: CellData): ExcelCell => {
        const st = cd.style;
        const align = st?.align ?? cd.align;
        const cell: ExcelCell = {
            kind: GridCellKind.Text,
            data: cd.text,
            displayData: cd.warn ? cd.text + " \u26A0" : cd.text,
            allowOverlay: true,
            contentAlign: alignToContent(align),
            excel: st,
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
            const [col, row] = cell;
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
        [store, rowNumberCell, toExcelCell]
    );

    const drawCell = React.useCallback<DrawCellCallback>((args, done) => {
        const cell = args.cell as ExcelCell;
        const st = cell.excel;
        if (!st) {
            done();
            return;
        }
        const { ctx, rect } = args;
        if (st.fill) {
            ctx.fillStyle = st.fill;
            ctx.fillRect(rect.x, rect.y, rect.width, rect.height);
        }
        if (st.border && st.border !== "none") {
            ctx.save();
            ctx.strokeStyle = "#000000";
            ctx.lineWidth = 1;
            ctx.beginPath();
            if (st.border === "bottom") {
                ctx.moveTo(rect.x, rect.y + rect.height - 0.5);
                ctx.lineTo(rect.x + rect.width, rect.y + rect.height - 0.5);
            } else {
                ctx.rect(rect.x + 0.5, rect.y + 0.5, rect.width - 1, rect.height - 1);
            }
            ctx.stroke();
            ctx.restore();
        }
        const prevFont = ctx.font;
        const baseFont = (args.theme as any).baseFontFull ?? "";
        ctx.font = `${st.italic ? "italic " : ""}${st.bold ? "bold " : ""}${baseFont}`;
        const prevTheme = args.theme;
        if (st.fontColor) args.theme = { ...args.theme, textDark: st.fontColor };
        done();
        args.theme = prevTheme;
        ctx.font = prevFont;
        if (st.underline) {
            ctx.save();
            ctx.fillStyle = st.fontColor ?? "#000000";
            const w = Math.min(rect.width - 8, ctx.measureText(cell.displayData).width);
            ctx.fillRect(rect.x + 4, rect.y + rect.height - 3, w, 1);
            ctx.restore();
        }
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
            for (let i = 0; i < freeze.rows; i++) {
                const top = rect.y + band + i * rowHeight;
                ctx.fillStyle = i % 2 === 0 ? t.bgCell : t.bgCellMedium;
                ctx.fillRect(rect.x, top, rect.width, rowHeight);
                if (a.columnIndex > 0) {
                    const cd = store.getRow(i)?.cells[a.columnIndex - 1];
                    if (cd) {
                        const align = cd.style?.align ?? cd.align;
                        ctx.save();
                        ctx.fillStyle = cd.color ?? t.textDark;
                        ctx.font = t.baseFontFull ?? `${t.baseFontStyle} ${t.fontFamily}`;
                        ctx.textBaseline = "middle";
                        if (align === 1) {
                            ctx.textAlign = "right";
                            ctx.fillText(cd.text, rect.x + rect.width - 4, top + rowHeight / 2);
                        } else if (align === 2) {
                            ctx.textAlign = "center";
                            ctx.fillText(cd.text, rect.x + rect.width / 2, top + rowHeight / 2);
                        } else {
                            ctx.textAlign = "left";
                            ctx.fillText(cd.text, rect.x + 4, top + rowHeight / 2);
                        }
                        ctx.restore();
                    }
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
        [headerBand, rowHeight, freeze.rows, showHeadings, store]
    );

    const getCellsForSelection = React.useCallback(
        (range: Rectangle): GridCell[][] => {
            const emptyCell: TextCell = { kind: GridCellKind.Text, data: "", displayData: "", allowOverlay: false };
            const out: GridCell[][] = [];
            for (let y = range.y; y < range.y + range.height; y++) {
                const wire = store.getRow(y);
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
        [store, rowNumberCell, toExcelCell]
    );

    const handleCellEdited = React.useCallback(
        (cell: Item, newValue: GridCell) => {
            const [col, row] = cell;
            if (col <= 0) return;
            onCellEdited(row, col - 1, (newValue as TextCell).data);
        },
        [onCellEdited]
    );

    const handleDelete = React.useCallback(
        (selection: GridSelection) => {
            const r = computeRange(selection);
            onDeleteRanges([[r.r1, r.c1, r.r2, r.c2]]);
            return false;
        },
        [computeRange, onDeleteRanges]
    );

    const handleFill = React.useCallback(
        (e: FillPatternEventArgs) => {
            const src = e.patternSource;
            const dst = e.fillDestination;
            const from = [src.y, Math.max(0, src.x - 1), src.y + src.height - 1, Math.max(0, src.x + src.width - 2)];
            const to = [dst.y, Math.max(0, dst.x - 1), dst.y + dst.height - 1, Math.max(0, dst.x + dst.width - 2)];
            onFill(from, to);
            e.preventDefault();
        },
        [onFill]
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
            store.ensure(range.y - VISIBLE_MARGIN, range.y + range.height + VISIBLE_MARGIN);
        },
        [store]
    );

    const theme = React.useMemo(() => {
        const t = scaledTheme(zoom);
        if (!showGridlines) {
            t.borderColor = "transparent";
            t.horizontalBorderColor = "transparent";
        }
        return t;
    }, [zoom, showGridlines]);

    return (
        <div className="gridHost" ref={containerRef}>
            <DataEditor
                ref={editorRef}
                width={width}
                height={height}
                columns={columns}
                rows={viewRows}
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
                onHeaderMenuClick={handleHeaderMenu}
                drawCell={drawCell}
                drawHeader={drawHeader}
                smoothScrollX
                smoothScrollY
                rangeSelect="multi-rect"
                columnSelect="multi"
                rowSelect="multi"
                fillHandle
                keybindings={{ copy: false, cut: false, paste: false, delete: false, selectAll: false }}
            />
        </div>
    );
});






