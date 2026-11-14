import * as React from "react";
import { call, on as onEvent, isMock } from "./bridge";
import { PaintSignal } from "./paintSignal";
import { resolveCommand } from "./keys";
import { Grid, type GridHandle, type SelRange } from "./grid/Grid";
import { Ribbon, type TabName } from "./ribbon/Ribbon";
import { FileBackstage, type FileInfo } from "./ribbon/FileBackstage";
import { FormulaBar } from "./bars/FormulaBar";
import { SheetTabs } from "./bars/SheetTabs";
import { StatusBar } from "./bars/StatusBar";
import { Dialogs } from "./dialogHost";
import { increaseDecimals, decreaseDecimals, PERCENT_CODE, COMMA_CODE, GENERAL } from "./numberFormat";
import { SHORTCUT_FORMATS, fillRange, todayInput, timeInput } from "./shortcuts";
import { AgentPanel } from "./agent/AgentPanel";
import type { WorkbookInfo, SheetInfo, Stats, CellStyle, Condition, SortKey } from "./types";

type AppState = "empty" | "opening" | "ready";

/** Excel column letters "A" -> 1 (1-based). */
function colFromLetters(s: string): number {
    let n = 0;
    for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
}

/** Column letters for a 1-based column number (1 -> "A"). */
// The Data mode is loaded the first time it is opened, so the editor starts with a smaller bundle.
const DataApp = React.lazy(() => import("./data/DataApp").then((m) => ({ default: m.DataApp })));

function lettersFromCol(n: number): string {
    let s = "";
    while (n > 0) {
        const rem = (n - 1) % 26;
        s = String.fromCharCode(65 + rem) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

export function App(): React.ReactElement {
    const [mode, setMode] = React.useState<"editor" | "data">(() => new URLSearchParams(window.location.search).get("mode") === "data" ? "data" : "editor");
    const [state, setState] = React.useState<AppState>("empty");
    const [opening, setOpening] = React.useState<{ phase: string; percent: number }>({ phase: "Opening", percent: 0 });
    const [workbook, setWorkbook] = React.useState<WorkbookInfo | null>(null);
    const [sheetIndex, setSheetIndex] = React.useState(0);
    const [viewRows, setViewRows] = React.useState(0);
    const [filterActive, setFilterActive] = React.useState(false);
    const [filters, setFilters] = React.useState<Record<number, Condition>>({});
    const [filterMode, setFilterMode] = React.useState(false);
    const [sort, setSort] = React.useState<SortKey[]>([]);
    const [selection, setSelection] = React.useState<SelRange>({ r1: 0, c1: 0, r2: 0, c2: 0 });
    const [cellRaw, setCellRaw] = React.useState("");
    const [cellAddress, setCellAddress] = React.useState("A1");
    const [cellEndAddress, setCellEndAddress] = React.useState("A1");
    const [zoom, setZoom] = React.useState(100);
    const [gridlines, setGridlines] = React.useState(true);
    const [headings, setHeadings] = React.useState(true);
    const [formulaBarVisible, setFormulaBarVisible] = React.useState(true);
    const [freeze, setFreeze] = React.useState({ rows: 0, cols: 0 });
    const [status, setStatus] = React.useState("Ready");
    const [recovery, setRecovery] = React.useState<{ cells: number; savedAt: string; sameFile: boolean } | null>(null);
    const [stats, setStats] = React.useState<Stats | null>(null);
    const [ribbonTab, setRibbonTab] = React.useState<TabName>("Home");
    const [ribbonCollapsed, setRibbonCollapsed] = React.useState(false);
    const [recent, setRecent] = React.useState<string[]>([]);
    const [dirty, setDirty] = React.useState(false);
    const [viewKey, setViewKey] = React.useState(0);
    const [dialog, setDialog] = React.useState<{ name: string; props: any } | null>(null);
    const [fileInfo, setFileInfo] = React.useState<FileInfo | null>(null);
    const [paintMs, setPaintMs] = React.useState(0);
    const [numFmt, setNumFmt] = React.useState(GENERAL);
    const [fmtFlags, setFmtFlags] = React.useState({ bold: false, italic: false, underline: false });

    const gridRef = React.useRef<GridHandle | null>(null);
    const activeCell = React.useMemo(() => ({ row: selection.r1, col: selection.c1 }), [selection.r1, selection.c1]);
    const sheet: SheetInfo | null = workbook?.sheets[sheetIndex] ?? null;

    // Keep the latest values for the window.__app self-test hooks.
    const live = React.useRef<any>({});
    live.current = { mode, state, selection, sheetIndex, sheet, zoom, freeze, viewRows, filterActive, dirty, fileName: workbook?.name, workbook };

    const loadSheet = React.useCallback(async (index: number) => {
        setStatus("Reading...");
        try {
            const info: any = await call("sheet", { sheet: index });
            setViewRows(info.viewRows ?? 0);
            setFilterActive(!!info.filterActive);
            setFilters(info.filters ?? {});
            setSort(info.sort ?? []);
        } finally {
            setViewKey((k) => k + 1);
            setStatus("Ready");
        }
    }, []);

    const sheetRef = React.useRef(0);
    sheetRef.current = sheetIndex;
    const [agentOpen, setAgentOpen] = React.useState(false);

    React.useEffect(() => {
        const offs: (() => void)[] = [];
        // An AI agent changed the workbook through ExcelApp's MCP tools: read the sheet, the visible cells and the saved state again.
        offs.push(onEvent("agentChanged", () => {
            void loadSheet(sheetRef.current).catch(() => {});
            gridRef.current?.refresh();
            call<any>("workbook").then((w) => setDirty(!!w?.dirty)).catch(() => {});
        }));
        offs.push(
            onEvent("opening", (e) => {
                setState("opening");
                setRecovery(null);
                setOpening({ phase: e.phase, percent: e.percent });
                setStatus(e.phase + "...");
            })
        );
        offs.push(
            onEvent("opened", (e) => {
                const wb: WorkbookInfo = e.workbook;
                setWorkbook(wb);
                setState("ready");
                setDirty(!!wb.dirty);
                const index = e.keep ? Math.min(live.current.sheetIndex ?? 0, wb.sheets.length - 1) : ((wb as any).currentSheet ?? wb.sheets.findIndex(s => !s.hidden));
                const nextSheet = Math.max(0, index);
                setSheetIndex(nextSheet);
                setFreeze({ rows: wb.sheets[nextSheet]?.frozenRows ?? 0, cols: wb.sheets[nextSheet]?.frozenCols ?? 0 });
                if (!e.keep) setSelection({ r1: 0, c1: 0, r2: 0, c2: 0 });
                setFileInfo({
                    name: wb.name,
                    path: wb.path,
                    sizeBytes: 0,
                    lastRead: new Date().toLocaleTimeString(),
                    cacheBytes: 0,
                    nasaProtected: true,
                });
                setStatus("Ready");
                void loadSheet(nextSheet).then(async () => {
                    const info: any = await call("info");
                    setFileInfo({ name: wb.name, path: wb.path, sizeBytes: info.size, lastRead: info.readTime, cacheBytes: info.cacheBytes, nasaProtected: !!info.nasca });
                }).catch(() => {});
            })
        );
        offs.push(onEvent("recovery", (e) => setRecovery(e?.available ? { cells: e.cells ?? 0, savedAt: e.savedAt ?? "", sameFile: !!e.sameFile } : null)));
        offs.push(onEvent("openFailed", (e) => {
            setState("empty");
            setStatus("Could not open the file: " + (e.error ?? ""));
        }));
        offs.push(onEvent("closeRequested", (e) => {
            setDialog({ name: "ClosePrompt", props: { fileName: e.fileName, onAnswer: (action: string) => {
                setDialog(null);
                void call("closeAnswer", { action });
            } } });
        }));
        offs.push(onEvent("saving", (e) => setStatus(e.phase + "...")));
        offs.push(onEvent("saved", () => {
            setDialog(null);
            setDirty(false);
            setStatus("Ready");
            void loadSheet(live.current.sheetIndex ?? 0);
        }));
        offs.push(onEvent("saveFailed", (e) => {
            setDialog(null);
            setStatus("Save failed: " + (e.error ?? ""));
        }));
        offs.push(onEvent("memory", (e) => {
            if ((e.commitPercent ?? 0) > 85) setStatus("Memory " + e.commitPercent + "% used");
        }));
        offs.push(onEvent("fileChanged", () => setStatus("This file changed on disk. Read it again to see the changes.")));
        return () => offs.forEach((o) => o());
    }, [loadSheet]);

    React.useEffect(() => {
        (async () => {
            try {
                const r: any = await call("ready");
                setRecent(r?.recent ?? []);
                if (isMock()) await call("open", { path: "mock://syn-mock.xlsx" });
            } catch {
                /* the host is not ready yet */
            }
        })();
    }, []);

    React.useEffect(() => {
        if (state !== "ready") return;
        let cancelled = false;
        (async () => {
            try {
                const d: any = await call("cellDetail", { sheet: sheetIndex, row: activeCell.row, col: activeCell.col });
                const end: any = selection.r1 === selection.r2 && selection.c1 === selection.c2 ? d : await call("cellDetail", { sheet: sheetIndex, row: selection.r2, col: selection.c2 });
                if (!cancelled) {
                    setCellEndAddress(end.address ?? d.address);
                    setCellRaw(d.raw ?? "");
                    setCellAddress(d.address ?? "A1");
                    setFmtFlags({ bold: !!d.style?.bold, italic: !!d.style?.italic, underline: !!d.style?.underline });
                    setNumFmt(d.format ?? GENERAL);
                }
            } catch {
                /* ignore transient errors */
            }
        })();
        return () => {
            cancelled = true;
        };
    }, [state, sheetIndex, activeCell.row, activeCell.col, selection.r2, selection.c2, viewKey]);

    React.useEffect(() => {
        if (state !== "ready") {
            setStats(null);
            return;
        }
        const cells = (selection.r2 - selection.r1 + 1) * (selection.c2 - selection.c1 + 1);
        const handle = setTimeout(async () => {
            if (cells > 2_000_000) {
                setStats({ count: cells, numbers: 0, sum: 0, average: 0, min: 0, max: 0 });
                return;
            }
            try {
                const s: Stats = await call("stats", {
                    sheet: sheetIndex,
                    ranges: [[selection.r1, selection.c1, selection.r2, selection.c2]],
                });
                setStats(s);
            } catch {
                setStats(null);
            }
        }, 150);
        return () => clearTimeout(handle);
    }, [state, sheetIndex, selection]);

    const refreshView = React.useCallback(() => setViewKey((k) => k + 1), []);

    const applyChanged = React.useCallback((res: { changed?: number[][]; recalc?: boolean } | null | undefined) => {
        // With formulas on the sheet any cell can show a new result, so read the visible rows again.
        if (res?.recalc) { gridRef.current?.refresh(); return; }
        const changed = res?.changed;
        if (changed && changed.length > 0) gridRef.current?.invalidateCells(changed);
    }, []);

    const confirmLeave = React.useCallback((next: () => Promise<void>) => {
        if (!live.current.dirty) { void next().catch((e: Error) => setStatus(e.message)); return; }
        setDialog({ name: "ClosePrompt", props: { fileName: live.current.fileName, onAnswer: async (action: string) => {
            setDialog(null);
            if (action === "cancel") return;
            try {
                if (action === "save") {
                    const result: any = await call("save", { saveAs: false });
                    if (!result?.saved) return;
                }
                await next();
            } catch (e) { setStatus((e as Error).message); }
        } } });
    }, []);

    const openPath = React.useCallback(async (path: string) => {
        confirmLeave(async () => { await call("open", { path, discard: true }); });
    }, [confirmLeave]);

    const openBrowse = React.useCallback(async () => {
        const r: any = await call("pickFile");
        if (r?.path) await openPath(r.path);
    }, [openPath]);

    const doSave = React.useCallback(
        (saveAs: boolean) => {
            if (!sheet) return;
            setDialog({ name: "SaveProgress", props: { onDone: () => setDialog(null) } });
            void call("save", { saveAs }).then((r: any) => {
                if (r?.cancelled) setDialog(null);
            }).catch((e: Error) => { setDialog(null); setStatus(e.message); });
        },
        [sheet]
    );

    const goToAddress = React.useCallback(
        async (addr: string) => {
            if (!sheet) return;
            const m = /^\s*([A-Za-z]+)\s*(\d+)\s*(?::\s*([A-Za-z]+)\s*(\d+))?\s*$/.exec(addr);
            if (!m) return;
            const c1 = colFromLetters(m[1]) - sheet.firstCol;
            const r1 = Number(m[2]) - sheet.firstRow;
            const c2 = m[3] ? colFromLetters(m[3]) - sheet.firstCol : c1;
            const r2 = m[4] ? Number(m[4]) - sheet.firstRow : r1;
            if (c1 < 0 || c2 < 0 || c1 >= sheet.cols || c2 >= sheet.cols) return;
            const a: any = await call("locate", { sheet: sheetIndex, row: r1 });
            const b: any = await call("locate", { sheet: sheetIndex, row: r2 });
            if (a.view < 0 || b.view < 0) { setStatus("This cell is hidden by a filter"); return; }
            gridRef.current?.setSelection(a.view, c1, b.view, c2);
            gridRef.current?.scrollTo(a.view, c1 + 1);
        },
        [sheet, sheetIndex]
    );

    const selectSheet = React.useCallback(
        (index: number) => {
            setSheetIndex(index);
            const next = live.current.workbook?.sheets?.[index];
            if (next) setFreeze({ rows: next.frozenRows ?? 0, cols: next.frozenCols ?? 0 });
            setSelection({ r1: 0, c1: 0, r2: 0, c2: 0 });
            void loadSheet(index);
        },
        [loadSheet]
    );

    const applyFreeze = React.useCallback(
        (rows: number, cols: number) => {
            const r = Math.max(0, Math.min(10, rows));
            setFreeze({ rows: r, cols: Math.max(0, cols) });
            void call("setFreeze", { sheet: sheetIndex, rows: r, cols });
        },
        [sheetIndex]
    );

    const setZoomPct = React.useCallback((pct: number) => {
        setZoom(Math.max(50, Math.min(200, Math.round(pct))));
    }, []);

    const applyStyle = React.useCallback(
        async (style: CellStyle) => {
            if (!sheet) return;
            const ranges = [[selection.r1, selection.c1, selection.r2, selection.c2]];
            const res: any = await call("setStyle", { sheet: sheetIndex, ranges, style });
            applyChanged(res);
            setDirty(true);
        },
        [sheet, selection, sheetIndex, applyChanged]
    );

    const commitCell = React.useCallback(
        async (viewRow: number, col: number, text: string) => {
            let res: any;
            try {
                res = await call("setCell", { sheet: sheetIndex, row: viewRow, col, input: text });
            } catch (e) {
                // Excel refuses a formula it cannot read; say why and leave the cell as it was.
                const msg = (e as Error)?.message ?? String(e);
                setStatus(msg);
                setDialog({ name: "Message", props: { title: "Microsoft Excel", text: msg, onClose: () => setDialog(null) } });
                return;
            }
            applyChanged(res);
            setDirty(true);
            setCellRaw(text);
        },
        [sheetIndex, applyChanged]
    );

    const clearRanges = React.useCallback(
        async (ranges: number[][], what = "contents") => {
            if (!sheet) return;
            const res: any = await call("clearCells", { sheet: sheetIndex, ranges, what });
            applyChanged(res);
            setDirty(true);
        },
        [sheet, sheetIndex, applyChanged]
    );

    const doFill = React.useCallback(
        async (from: number[], to: number[], copy = false) => {
            // copy: Ctrl+D / Ctrl+R copy exactly; the fill handle continues a series
            const res: any = await call("fill", { sheet: sheetIndex, from, to, copy });
            applyChanged(res);
            setDirty(true);
        },
        [sheetIndex, applyChanged]
    );

    const rangesNow = React.useCallback((): number[][] => [[selection.r1, selection.c1, selection.r2, selection.c2]], [selection]);

    const doCopy = React.useCallback(() => {
        if (sheet) void call("copy", { sheet: sheetIndex, ranges: rangesNow() });
    }, [sheet, sheetIndex, rangesNow]);

    const doCut = React.useCallback(() => {
        if (sheet) void call("cut", { sheet: sheetIndex, ranges: rangesNow() });
    }, [sheet, sheetIndex, rangesNow]);

    const doPaste = React.useCallback(async () => {
        if (!sheet) return;
        await call("paste", { sheet: sheetIndex, row: activeCell.row, col: activeCell.col });
        refreshView();
        setDirty(true);
    }, [sheet, sheetIndex, activeCell, refreshView]);

    const doUndo = React.useCallback(async () => {
        if (!sheet) return;
        const r: any = await call("undo", { sheet: sheetIndex });
        setDirty(!!r.dirty);
        if (r.sheet >= 0 && r.sheet !== sheetIndex) selectSheet(r.sheet);
        refreshView();
    }, [sheet, sheetIndex, refreshView, selectSheet]);

    const doRedo = React.useCallback(async () => {
        if (!sheet) return;
        const r: any = await call("redo", { sheet: sheetIndex });
        setDirty(!!r.dirty);
        if (r.sheet >= 0 && r.sheet !== sheetIndex) selectSheet(r.sheet);
        refreshView();
    }, [sheet, sheetIndex, refreshView, selectSheet]);

    const toggleBold = React.useCallback(() => {
        const v = !fmtFlags.bold;
        setFmtFlags((f) => ({ ...f, bold: v }));
        void applyStyle({ bold: v });
    }, [fmtFlags.bold, applyStyle]);
    const toggleItalic = React.useCallback(() => {
        const v = !fmtFlags.italic;
        setFmtFlags((f) => ({ ...f, italic: v }));
        void applyStyle({ italic: v });
    }, [fmtFlags.italic, applyStyle]);
    const toggleUnderline = React.useCallback(() => {
        const v = !fmtFlags.underline;
        setFmtFlags((f) => ({ ...f, underline: v }));
        void applyStyle({ underline: v });
    }, [fmtFlags.underline, applyStyle]);

    const onNumberFormat = React.useCallback(
        (code: string) => {
            setNumFmt(code);
            void applyStyle({ numberFormat: code });
        },
        [applyStyle]
    );
    const onPercent = React.useCallback(() => {
        setNumFmt(PERCENT_CODE);
        void applyStyle({ numberFormat: PERCENT_CODE });
    }, [applyStyle]);
    const onComma = React.useCallback(() => {
        setNumFmt(COMMA_CODE);
        void applyStyle({ numberFormat: COMMA_CODE });
    }, [applyStyle]);
    const onIncrease = React.useCallback(() => {
        const c = increaseDecimals(numFmt);
        setNumFmt(c);
        void applyStyle({ numberFormat: c });
    }, [numFmt, applyStyle]);
    const onDecrease = React.useCallback(() => {
        const c = decreaseDecimals(numFmt);
        setNumFmt(c);
        void applyStyle({ numberFormat: c });
    }, [numFmt, applyStyle]);

    const onClearAll = React.useCallback(() => void clearRanges(rangesNow(), "all"), [clearRanges, rangesNow]);
    const onClearContents = React.useCallback(() => void clearRanges(rangesNow()), [clearRanges, rangesNow]);
    const onClearFormats = React.useCallback(() => void clearRanges(rangesNow(), "formats"), [clearRanges, rangesNow]);

    const numberAbove = React.useCallback(
        (row: number, col: number): boolean => gridRef.current?.store.getRow(row - 1)?.cells[col]?.kind === 1,
        []
    );

    const onAutoSum = React.useCallback(() => {
        if (!sheet) return;
        const col = activeCell.col;
        const row = activeCell.row;
        const L = (c: number) => lettersFromCol(c + sheet.firstCol);
        let top = row;
        while (top - 1 >= 0 && numberAbove(top, col)) top--;
        if (top < row) {
            void commitCell(row, col, `=SUM(${L(col)}${top + sheet.firstRow}:${L(col)}${row - 1 + sheet.firstRow})`);
            return;
        }
        let left = col;
        while (left - 1 >= 0 && gridRef.current?.store.getRow(row)?.cells[left - 1]?.kind === 1) left--;
        if (left < col) void commitCell(row, col, `=SUM(${L(left)}${row + sheet.firstRow}:${L(col - 1)}${row + sheet.firstRow})`);
    }, [sheet, activeCell, numberAbove, commitCell]);

    const onFormatCells = React.useCallback(() => {
        if (sheet) setDialog({ name: "FormatCells", props: { code: numFmt, onApply: onNumberFormat, onClose: () => setDialog(null) } });
    }, [sheet, numFmt, onNumberFormat]);

    const doSort = React.useCallback(
        async (col: number, desc: boolean) => {
            const res: any = await call("setSort", { sheet: sheetIndex, keys: col >= 0 ? [{ col, desc }] : [] });
            setSort(col >= 0 ? [{ col, desc }] : []);
            setViewRows(res.viewRows ?? viewRows);
            refreshView();
        },
        [sheetIndex, refreshView, viewRows]
    );
    const onSortAZ = React.useCallback(() => void doSort(activeCell.col, false), [doSort, activeCell.col]);
    const onSortZA = React.useCallback(() => void doSort(activeCell.col, true), [doSort, activeCell.col]);

    const onClearFilter = React.useCallback(async () => {
        const res: any = await call("clearFilters", { sheet: sheetIndex });
        setFilterActive(false);
        setViewRows(res.viewRows ?? viewRows);
        refreshView();
    }, [sheetIndex, refreshView, viewRows]);

    const openFilterMenu = React.useCallback(
        (col: number, x: number, y: number) => {
            if (!sheet) return;
            setDialog({
                name: "FilterMenu",
                props: {
                    sheet: sheetIndex,
                    col,
                    letter: lettersFromCol(col + sheet.firstCol),
                    anchor: { x, y },
                    onClose: () => {
                        setDialog(null);
                        void loadSheet(sheetIndex);
                    },
                },
            });
        },
        [sheet, sheetIndex, loadSheet]
    );

    const openSortDialog = React.useCallback(() => {
        if (!sheet) return;
        setDialog({
            name: "SortDialog",
            props: {
                sheet: sheetIndex,
                colCount: sheet.cols,
                letter: (c: number) => lettersFromCol(c + sheet.firstCol),
                onClose: () => {
                    setDialog(null);
                    void loadSheet(sheetIndex);
                },
            },
        });
    }, [sheet, sheetIndex, loadSheet]);

    const dlgFind = React.useCallback(
        (initialTab: "find" | "replace") => {
            setDialog({
                name: "FindReplace",
                props: {
                    sheet: sheetIndex,
                    initialTab,
                    onChanged: (r: any) => { setDirty(!!r.dirty); refreshView(); },
                    onSelect: async (s: number, r: number, c: number) => {
                        const loc: any = await call("locate", { sheet: s, row: r });
                        if (loc.view < 0) { setStatus("This cell is hidden by a filter"); return; }
                        await loadSheet(s); setSheetIndex(s);
                        const next = live.current.workbook?.sheets?.[s];
                        if (next) setFreeze({ rows: next.frozenRows ?? 0, cols: next.frozenCols ?? 0 });
                        requestAnimationFrame(() => {
                            gridRef.current?.setSelection(loc.view, c, loc.view, c);
                            gridRef.current?.scrollTo(loc.view, c + 1);
                        });
                    },
                    onClose: () => setDialog(null),
                },
            });
        },
        [sheetIndex, loadSheet]
    );

    const dlgGoTo = React.useCallback(() => {
        setDialog({ name: "GoTo", props: { onGo: goToAddress, onClose: () => setDialog(null) } });
    }, [goToAddress]);

    const doEdge = React.useCallback(
        async (cmd: string) => {
            if (!sheet) return;
            const dir = cmd === "edgeUp" ? "up" : cmd === "edgeDown" ? "down" : cmd === "edgeLeft" ? "left" : "right";
            const wire = gridRef.current?.store.getRow(activeCell.row);
            const realRow = wire ? wire.rowNumber - sheet.firstRow : activeCell.row;
            try {
                const res: any = await call("edge", { sheet: sheetIndex, row: realRow, col: activeCell.col, direction: dir });
                const loc: any = await call("locate", { sheet: sheetIndex, row: res.row });
                const vr = loc.view >= 0 ? loc.view : activeCell.row;
                gridRef.current?.setSelection(vr, res.col, vr, res.col);
                gridRef.current?.scrollTo(vr, res.col + 1);
            } catch {
                /* ignore */
            }
        },
        [sheet, activeCell, sheetIndex]
    );

    const runFilter = React.useCallback(
        async (col: number, values: string[]) => {
            const res: any = await call("setFilter", { sheet: sheetIndex, col, condition: { kind: "values", include: values, blanks: false } });
            setFilterActive(!!res.filterActive);
            setViewRows(res.viewRows ?? viewRows);
            refreshView();
        },
        [sheetIndex, viewRows, refreshView]
    );

    const runStyle = React.useCallback(
        async (range: any, style: CellStyle) => {
            const ranges = Array.isArray(range) ? [range] : rangesNow();
            const res: any = await call("setStyle", { sheet: sheetIndex, ranges, style });
            applyChanged(res);
        },
        [sheetIndex, rangesNow, applyChanged]
    );

    const dispatch = React.useCallback(
        (cmd: string): boolean => {
            switch (cmd) {
                case "save": doSave(false); return true;
                case "open": void openBrowse(); return true;
                case "find": dlgFind("find"); return true;
                case "replace": dlgFind("replace"); return true;
                case "goTo": dlgGoTo(); return true;
                case "copy": doCopy(); return true;
                case "cut": doCut(); return true;
                case "paste": void doPaste(); return true;
                case "undo": void doUndo(); return true;
                case "redo": void doRedo(); return true;
                case "bold": toggleBold(); return true;
                case "italic": toggleItalic(); return true;
                case "underline": toggleUnderline(); return true;
                case "filterToggle": setFilterMode((f) => !f); return true;
                case "clear": void clearRanges(rangesNow()); return true;
                case "openFilter": if (sheet) openFilterMenu(activeCell.col, 120, 120); return true;
                case "firstCell": gridRef.current?.setSelection(0, 0, 0, 0); gridRef.current?.scrollTo(0, 1); return true;
                case "lastCell": if (sheet) { const r = Math.max(0, viewRows - 1), c = Math.max(0, sheet.cols - 1); gridRef.current?.setSelection(r, c, r, c); gridRef.current?.scrollTo(r, c + 1); } return true;
                case "selectAll": if (sheet) gridRef.current?.setSelection(0, 0, Math.max(0, viewRows - 1), sheet.cols - 1); return true;
                case "prevSheet": selectSheet(Math.max(0, sheetIndex - 1)); return true;
                case "nextSheet": selectSheet(Math.min((workbook?.sheets.length ?? 1) - 1, sheetIndex + 1)); return true;
                case "zoomIn": setZoomPct(zoom + 10); return true;
                case "zoomOut": setZoomPct(zoom - 10); return true;
                case "fillDown":
                case "fillRight": {
                    const fr = sheet ? fillRange(cmd === "fillDown" ? "down" : "right", selection) : null;
                    if (fr) void doFill(fr.from, fr.to, true).catch((e: Error) => setStatus(e.message));
                    return true;
                }
                case "insertDate": if (sheet) void commitCell(activeCell.row, activeCell.col, todayInput(new Date())); return true;
                case "insertTime": if (sheet) void commitCell(activeCell.row, activeCell.col, timeInput(new Date())); return true;
                case "autoSum": onAutoSum(); return true;
                case "formatCells": onFormatCells(); return true;
                case "selectColumn": if (sheet) gridRef.current?.setSelection(0, selection.c1, Math.max(0, viewRows - 1), selection.c2); return true;
                case "selectRow": if (sheet) gridRef.current?.setSelection(selection.r1, 0, selection.r2, Math.max(0, sheet.cols - 1)); return true;
                case "edgeUp":
                case "edgeDown":
                case "edgeLeft":
                case "edgeRight":
                    void doEdge(cmd);
                    return true;
                default:
                    if (cmd in SHORTCUT_FORMATS) {
                        if (sheet) onNumberFormat(SHORTCUT_FORMATS[cmd]);
                        return true;
                    }
                    return false;
            }
        },
        [doSave, openBrowse, dlgFind, dlgGoTo, doCopy, doCut, doPaste, doUndo, doRedo, toggleBold, toggleItalic, toggleUnderline, clearRanges, rangesNow, sheet, openFilterMenu, activeCell, viewRows, selectSheet, sheetIndex, workbook, zoom, setZoomPct, doEdge, selection, doFill, commitCell, onAutoSum, onFormatCells, onNumberFormat]
    );

    const dispatchRef = React.useRef(dispatch);
    dispatchRef.current = dispatch;
    const paintRef = React.useRef(new PaintSignal());
    const apiRef = React.useRef<any>({});
    apiRef.current = { applyFreeze, setZoomPct, runFilter, runStyle, commitCell };

    React.useEffect(() => {
        const onKey = (e: KeyboardEvent) => {
            if (live.current.mode !== "editor" || live.current.state !== "ready") return;
            const tag = (e.target as HTMLElement)?.tagName;
            const inField = tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT";
            const cmd = resolveCommand(e);
            if (!cmd) return;
            if (inField && cmd !== "save") return;
            if (dispatchRef.current(cmd)) e.preventDefault();
        };
        window.addEventListener("keydown", onKey);
        return () => window.removeEventListener("keydown", onKey);
    }, []);

    React.useEffect(() => {
        const host: any = {
            state: () => live.current.state,
            hasRows: (row = 0) => !!gridRef.current?.store.getRow(row),
            scrollTo: (v: number) => gridRef.current?.scrollTo(v),
            run: (action: string, ...args: any[]) => {
                const api = apiRef.current;
                if (action === "freeze") api.applyFreeze(args[0] ?? 0, args[1] ?? 0);
                else if (action === "zoom") api.setZoomPct(args[0] ?? 100);
                else if (action === "filter") return api.runFilter(args[0], args[1] ?? []);
                else if (action === "style") return api.runStyle(args[0], args[1] ?? {});
                else if (action === "edit") return api.commitCell(args[0], args[1], args[2]);
                else if (action === "select") gridRef.current?.setSelection(args[0], args[1], args[2], args[3]);
            },
        };
        Object.defineProperty(host, "lastPaintMs", { get: () => paintRef.current.lastPaintMs, configurable: true });
        Object.defineProperty(host, "paintToken", { get: () => paintRef.current.token, configurable: true });
        (window as any).__app = host;
        return () => {
            delete (window as any).__app;
        };
    }, []);

    const handlePaint = React.useCallback((ms: number) => {
        paintRef.current.record(ms);
        setPaintMs(Math.round(ms));
    }, []);
    const onStoreReady = React.useCallback(() => {}, []);

    const displayAddress = React.useMemo(() => {
        if (!sheet) return "A1";
        if (selection.r1 !== selection.r2 || selection.c1 !== selection.c2) {
            return cellAddress + ":" + cellEndAddress;
        }
        return cellAddress;
    }, [sheet, selection, cellAddress, cellEndAddress]);

    const filterInfo = filterActive && sheet ? `${viewRows.toLocaleString()} of ${sheet.rows.toLocaleString()} records found` : null;

    return (
        <div className="appRoot">
            <div className="titleRow">
                <span className="appIcon" aria-hidden>
                    {"\u25A6"}
                </span>
                <span className="fileName">{workbook ? workbook.name + (dirty ? " *" : "") : "ExcelApp"}</span>
                <div className="mode-switch" role="group" aria-label="Application mode">
                    <button type="button" className={mode === "editor" ? "active" : ""} aria-pressed={mode === "editor"} onClick={() => setMode("editor")}>Editor</button>
                    <button type="button" className={mode === "data" ? "active" : ""} aria-pressed={mode === "data"} onClick={() => setMode("data")}>Data</button>
                </div>
                <button type="button" className={"agentToggle" + (agentOpen ? " active" : "")} aria-pressed={agentOpen} onClick={() => setAgentOpen((o) => !o)} data-testid="agent-toggle" title="Ask Claude Code or Codex to work in ExcelApp">
                    ✦ Agent
                </button>
                <span className="autoSavePill">AutoSave Off</span>
            </div>
            <div style={{ display: mode === "editor" ? "contents" : "none" }}>
            <Ribbon
                activeTab={ribbonTab}
                onTab={(t) => setRibbonTab(t)}
                collapsed={ribbonCollapsed}
                onToggleCollapse={() => setRibbonCollapsed((c) => !c)}
                enabled={state === "ready"}
                onPaste={() => void doPaste()}
                onCut={doCut}
                onCopy={doCopy}
                fontName="Aptos"
                fontSize="11"
                onBold={toggleBold}
                onItalic={toggleItalic}
                onUnderline={toggleUnderline}
                onStyle={applyStyle}
                onBorders={(k) => void applyStyle({ border: k })}
                onAlign={(a) => void applyStyle({ align: a })}
                onWrap={() => {}}
                onNumberFormat={onNumberFormat}
                onPercent={onPercent}
                onComma={onComma}
                onIncrease={onIncrease}
                onDecrease={onDecrease}
                onAutoSum={onAutoSum}
                onFormatCells={onFormatCells}
                onClearAll={onClearAll}
                onClearContents={onClearContents}
                onClearFormats={onClearFormats}
                onSortAZ={onSortAZ}
                onSortZA={onSortZA}
                onSortDialog={openSortDialog}
                onFilterToggle={() => setFilterMode((f) => !f)}
                onClearFilter={onClearFilter}
                onReapply={refreshView}
                filterActive={filterMode}
                onFind={() => dlgFind("find")}
                onReplace={() => dlgFind("replace")}
                onGoTo={dlgGoTo}
                gridlines={gridlines}
                headings={headings}
                formulaBar={formulaBarVisible}
                onToggleGridlines={() => setGridlines((g) => !g)}
                onToggleHeadings={() => setHeadings((h) => !h)}
                onToggleFormulaBar={() => setFormulaBarVisible((v) => !v)}
                onZoomDialog={() => setZoomPct(zoom === 100 ? 150 : 100)}
                onZoom100={() => setZoomPct(100)}
                onFreezePanes={() => applyFreeze(activeCell.row, activeCell.col)}
                onFreezeFirstColumn={() => applyFreeze(freeze.rows, 1)}
                onFreezeTopRow={() => applyFreeze(1, freeze.cols)}
                onUnfreeze={() => applyFreeze(0, 0)}
            />

            {ribbonTab === "File" ? (
                <FileBackstage
                    info={fileInfo}
                    recent={recent}
                    onOpenBrowse={() => void openBrowse()}
                    onOpenPath={(p) => void openPath(p)}
                    onSave={() => {
                        setRibbonTab("Home");
                        doSave(false);
                    }}
                    onSaveAs={() => {
                        setRibbonTab("Home");
                        doSave(true);
                    }}
                    onClose={() => {
                        setRibbonTab("Home");
                        confirmLeave(async () => {
                            await call("close", { discard: true });
                            setWorkbook(null); setState("empty"); setDirty(false);
                        });
                    }}
                    onBack={() => setRibbonTab("Home")}
                />
            ) : state === "ready" && sheet ? (
                <>
                    {recovery && (
                        <div className="recovery-bar" role="alert">
                            <span>
                                {recovery.cells} unsaved change{recovery.cells === 1 ? "" : "s"} from {recovery.savedAt} can be restored.
                                {!recovery.sameFile && " The file changed on disk since then; check the restored cells."}
                            </span>
                            <button onClick={async () => {
                                try {
                                    const r: any = await call("recoverEdits");
                                    setRecovery(null); setDirty(!!r.dirty); refreshView();
                                    setStatus(r.restored + " change" + (r.restored === 1 ? "" : "s") + " restored");
                                } catch (err: any) { setRecovery(null); setStatus(String(err?.message ?? err)); }
                            }}>Restore</button>
                            <button onClick={async () => { await call("discardRecovery").catch(() => {}); setRecovery(null); }}>Discard</button>
                        </div>
                    )}
                    <FormulaBar
                        address={displayAddress}
                        raw={cellRaw}
                        visible={formulaBarVisible}
                        onNameSubmit={goToAddress}
                        onRawCommit={(text) => void commitCell(activeCell.row, activeCell.col, text)}
                    />
                    <div className="gridArea">
                        <Grid
                            ref={gridRef}
                            sheet={sheet}
                            viewRows={viewRows}
                            filterActive={filterActive}
                            filters={filters}
                            zoom={zoom}
                            showGridlines={gridlines}
                            showHeadings={headings}
                            freeze={freeze}
                            onStoreReady={onStoreReady}
                            onSelectionChange={setSelection}
                            onCellEdited={commitCell}
                            onDeleteRanges={clearRanges}
                            onFill={doFill}
                            onHeaderMenu={openFilterMenu}
                            onPaint={handlePaint}
                            viewKey={viewKey}
                        />
                    </div>
                    <SheetTabs sheets={workbook?.sheets ?? []} activeIndex={sheetIndex} onSelect={selectSheet} onUnhide={selectSheet} />
                    <StatusBar status={status} filterInfo={filterInfo} stats={stats} zoom={zoom} onZoom={setZoomPct} />
                </>
            ) : state === "opening" ? (
                <div className="stateScreen">
                    <div className="stateCard">
                        <div className="stateTitle">Opening...</div>
                        <div className="stateSub">
                            {opening.phase} {opening.percent}%
                        </div>
                        <div className="progressBar">
                            <div className="progressFill" style={{ width: opening.percent + "%" }} />
                        </div>
                    </div>
                </div>
            ) : (
                <div className="stateScreen">
                    <div className="stateCard">
                        <div className="stateTitle">Open a file</div>
                        <div className="stateSub">Read a workbook to start. The first opening waits for Excel.</div>
                        <button type="button" className="primaryBtn" onClick={() => void openBrowse()} data-testid="open-file">
                            Open a file
                        </button>
                        <ul className="recentList">
                            {recent.length === 0 && <li className="recentEmpty">No recent files</li>}
                            {recent.map((p) => (
                                <li key={p}>
                                    <button type="button" className="recentItem" onClick={() => void openPath(p)}>
                                        {p}
                                    </button>
                                </li>
                            ))}
                        </ul>
                    </div>
                </div>
            )}

            </div>
            {mode === "data" && <React.Suspense fallback={<div className="data-loading" role="status">Loading Data...</div>}><DataApp /></React.Suspense>}
            {agentOpen && <AgentPanel onClose={() => setAgentOpen(false)} />}
            {dialog !== null &&
                React.createElement(Dialogs[dialog.name as keyof typeof Dialogs] as React.ComponentType<any>, dialog.props)}
            <span className="paintMs" data-testid="paintMs" style={{ display: "none" }}>
                {paintMs}
            </span>
        </div>
    );
}







