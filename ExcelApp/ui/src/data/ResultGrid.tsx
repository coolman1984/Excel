import * as React from "react";
import { DataEditor, GridCellKind, type GridCell, type Item } from "@glideapps/glide-data-grid";
import "@glideapps/glide-data-grid/dist/index.css";
import { call } from "../bridge";
import type { LineageCell, ResultInfo, ResultPage } from "./contracts";

interface Props {
    result: ResultInfo;
    onSort: (column: string) => void;
    onDrill: (result: ResultInfo) => void;
    onAction: (action: "export" | "pin" | "view" | "alert") => void;
    onEdit?: (row: number, column: string, value: string) => void;
    height?: number;
    editedCells?: ReadonlySet<string>;
}

export function isCurrentPage(token: number, generation: number, expected: string, received: string): boolean {
    return token === generation && expected === received;
}
export async function requestCurrentPage(request: Promise<ResultPage>, token: number, generation: () => number, resultId: string, accept: (page: ResultPage) => void): Promise<void> {
    const page = await request;
    if (isCurrentPage(token, generation(), resultId, page.resultId)) accept(page);
}

export function ResultGrid({ result, onSort, onDrill, onAction, onEdit, height = 460, editedCells }: Props): React.ReactElement {
    const alt = React.useRef(false);
    React.useEffect(() => { const key = (e: KeyboardEvent) => { alt.current = e.altKey; }; const clear = () => { alt.current = false; }; window.addEventListener("keydown", key); window.addEventListener("keyup", key); window.addEventListener("blur", clear); return () => { window.removeEventListener("keydown", key); window.removeEventListener("keyup", key); window.removeEventListener("blur", clear); }; }, []);
    const generation = React.useRef(0);
    const blocks = React.useRef(new Map<number, string[][]>());
    const pending = React.useRef(new Set<number>());
    const [revision, setRevision] = React.useState(0);
    const [lineage, setLineage] = React.useState<LineageCell | null>(null);
    const [menu, setMenu] = React.useState(false);
    React.useEffect(() => {
        generation.current++; blocks.current.clear(); pending.current.clear(); setLineage(null); setRevision(v => v + 1);
        return () => { generation.current++; };
    }, [result.resultId]);
    const ensure = React.useCallback((first: number, last: number) => {
        for (let block = Math.max(0, Math.floor(first / 200)); block <= Math.floor(Math.min(result.rows - 1, last) / 200); block++) {
            if (blocks.current.has(block) || pending.current.has(block)) continue;
            pending.current.add(block);
            const token = generation.current;
            void requestCurrentPage(call<ResultPage>("data.page", { resultId: result.resultId, from: block * 200, count: 200 }), token, () => generation.current, result.resultId, page => {
                blocks.current.set(block, page.rows); setRevision(v => v + 1);
            }).catch(() => { /* A later viewport update may retry the block. */ }).finally(() => { if (generation.current === token) pending.current.delete(block); });
        }
    }, [result.resultId, result.rows]);
    React.useEffect(() => { ensure(0, 199); }, [ensure]);
    const getCell = React.useCallback(([col, row]: Item): GridCell => {
        const values = blocks.current.get(Math.floor(row / 200))?.[row % 200];
        if (!values) return { kind: GridCellKind.Loading, allowOverlay: false };
        const text = values[col] ?? "";
        const statusIndex = result.columns.findIndex(c => c.name === "_status");
        const status = statusIndex < 0 ? "" : values[statusIndex]?.toLowerCase();
        const edited = editedCells?.has(`${row}:${result.columns[col].name}`);
        const bgCell = edited ? "#fff0b8" : status === "added" ? "#e4f5e8" : status === "removed" ? "#fde9e9" : status === "changed" ? "#fff3d8" : undefined;
        return { kind: GridCellKind.Text, data: text, displayData: text, allowOverlay: !!onEdit, readonly: !onEdit, themeOverride: bgCell ? { bgCell } : undefined };
    }, [result.columns, revision, onEdit, editedCells]);
    return <div className="result-grid"><div className="result-toolbar">{result.rows.toLocaleString()} rows · {result.ms} ms <button onClick={() => setMenu(v => !v)}>Actions ▾</button>{menu && <span>{(["export", "pin", "view", "alert"] as const).map(a => <button key={a} onClick={() => { setMenu(false); onAction(a); }}>{{ export: "Export…", pin: "Pin to Home", view: "Save as view", alert: "Make alert" }[a]}</button>)}</span>}</div>
        <DataEditor width="100%" height={height} columns={result.columns.map(c => ({ title: c.name, width: 160 }))} rows={result.rows} getCellContent={getCell} onVisibleRegionChanged={r => ensure(r.y, r.y + r.height + 40)} onHeaderClicked={col => onSort(result.columns[col].name)} onCellActivated={([, row]) => { if (result.groups) void call<ResultInfo>("data.drill", { resultId: result.resultId, row }).then(onDrill); }} onCellClicked={([col, row]) => { if (alt.current && !result.groups) void call<LineageCell>("data.lineage", { resultId: result.resultId, row, col }).then(setLineage); }} onCellEdited={([col, row], cell) => { if (onEdit && cell.kind === GridCellKind.Text) onEdit(row, result.columns[col].name, cell.data); }} onCellContextMenu={(_, e) => { e.preventDefault(); setMenu(true); }} />
        {lineage && <div className="lineage-card"><button onClick={() => setLineage(null)}>×</button><strong>Source of this value</strong><p>{lineage.file} · {lineage.sheet}!{lineage.cell}</p><p>Version {lineage.version} · {lineage.loadedAt}</p></div>}
    </div>;
}



