// The dashboard designer: tiles (a KPI, a chart or a table, each from a SQL query on the data store) on a 12-column grid.
// View shows the numbers; Design adds tiles, moves them by dragging, resizes them from the corner and edits them.
import * as React from "react";
import { call, on } from "../bridge";
import type { Dashboard, Tile, TileValue, TileData } from "./contracts";
import { ChartSvg, type ChartModel } from "../charts/ChartSvg";
import { fullNumber } from "../charts/scale";

const COLS = 12, ROW = 84, GAP = 14;
const slug = (s: string) => s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40) || "dashboard";

/** A tile's rows as a chart: the first text column holds the categories, the number columns are series. */
export function tileChart(data: TileData, type: string): ChartModel {
    const cat = data.columns.findIndex((c) => c.type === "text" || c.type === "date");
    const nums = data.columns.map((c, i) => ({ c, i })).filter((x) => x.c.type === "number" && x.i !== cat);
    return {
        type, categories: data.rows.map((r, i) => (cat >= 0 ? String(r[cat] ?? "") : String(i + 1))),
        series: nums.map(({ c, i }) => ({ name: c.name, values: data.rows.map((r) => (typeof r[i] === "number" ? (r[i] as number) : null)) })),
    };
}

const placed = (t: Tile) => t.x != null && t.y != null && t.x >= 0 && t.y >= 0;
/** The first free row under the tiles, for a new tile. */
export function nextRow(tiles: Tile[]): number { return layout(tiles).reduce((m, t) => Math.max(m, t.y! + (t.h ?? 2)), 0); }
/** Tiles with a place keep it; tiles without one (older dashboards) flow under them, four to a row. */
export function layout(tiles: Tile[]): Tile[] {
    let row = tiles.filter(placed).reduce((m, t) => Math.max(m, t.y! + (t.h ?? 2)), 0), col = 0, tallest = 0;
    return tiles.map((t) => {
        if (placed(t)) return t;
        const w = Math.min(COLS, t.w ?? 3), h = t.h ?? 2;
        if (col + w > COLS) { row += tallest; col = 0; tallest = 0; }
        const at = { ...t, x: col, y: row, w, h }; col += w; tallest = Math.max(tallest, h);
        return at;
    });
}

/** After a tile moves or grows, any tile it covers moves down below it, and so on: tiles never overlap. */
export function settle(tiles: Tile[], movedId: string): Tile[] {
    const out = layout(tiles).map((t) => ({ ...t }));
    const box = (t: Tile) => ({ x: t.x ?? 0, y: t.y ?? 0, w: t.w ?? 3, h: t.h ?? 2 });
    const overlap = (a: Tile, b: Tile) => { const p = box(a), q = box(b); return p.x < q.x + q.w && q.x < p.x + p.w && p.y < q.y + q.h && q.y < p.y + p.h; };
    const first = out.find((t) => t.id === movedId); if (!first) return out;
    const queue: Tile[] = [first];
    for (let guard = 0; queue.length && guard < 10000; guard++) {
        const f = queue.shift()!;
        for (const t of out) if (t !== f && overlap(f, t)) { t.y = box(f).y + box(f).h; queue.push(t); }
    }
    return out;
}

function formatValue(v: number | null, format: string): string {
    if (v == null) return "—";
    if (format === "percent") return (v * 100).toLocaleString(undefined, { maximumFractionDigits: 1 }) + "%";
    if (format === "currency") return v.toLocaleString(undefined, { style: "currency", currency: "USD", maximumFractionDigits: 0 });
    return Math.abs(v) >= 100000 ? new Intl.NumberFormat(undefined, { notation: "compact", maximumFractionDigits: 1 }).format(v) : v.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function TileBody({ tile, value, width, height }: { tile: Tile; value?: TileValue; width: number; height: number }): React.ReactElement {
    if (!value) return <div className="dash-muted">Calculating…</div>;
    if (value.error) return <div className="dash-error" role="alert">{value.error}</div>;
    if ((tile.kind ?? "kpi") === "kpi") return (
        <div className="dash-kpi">
            <div className="dash-kpi-value">{formatValue(value.value, tile.format)}</div>
            {value.deltaPct != null && <div className={"dash-delta " + (value.deltaPct >= 0 ? "up" : "down")}>{value.deltaPct >= 0 ? "▲" : "▼"} {Math.abs(value.deltaPct).toFixed(1)}%</div>}
        </div>);
    if (!value.data) return <div className="dash-muted">No rows.</div>;
    if (tile.kind === "chart") return <ChartSvg model={tileChart(value.data, tile.chartType ?? "column")} width={Math.max(120, width)} height={Math.max(80, height)} />;
    return (
        <div className="dash-table"><table><thead><tr>{value.data.columns.map((c) => <th key={c.name}>{c.name}</th>)}</tr></thead>
            <tbody>{value.data.rows.slice(0, 100).map((r, i) => <tr key={i}>{r.map((v, k) => <td key={k} className={typeof v === "number" ? "num" : ""}>{typeof v === "number" ? fullNumber(v) : String(v ?? "")}</td>)}</tr>)}</tbody></table></div>);
}

function TileEditor({ tile, onSave, onDelete, onClose }: { tile: Tile; onSave: (t: Tile) => void; onDelete: () => void; onClose: () => void }): React.ReactElement {
    const [t, setT] = React.useState<Tile>(tile);
    return (
        <div className="dash-modal" role="dialog" aria-label="Edit tile" onMouseDown={onClose}>
            <div className="dash-modal-box" onMouseDown={(e) => e.stopPropagation()}>
                <h2>Edit tile</h2>
                <label>Title <input value={t.title} onChange={(e) => setT({ ...t, title: e.target.value })} autoFocus /></label>
                <div className="dash-row">
                    <label>Kind <select value={t.kind ?? "kpi"} onChange={(e) => setT({ ...t, kind: e.target.value as Tile["kind"] })}><option value="kpi">KPI number</option><option value="chart">Chart</option><option value="table">Table</option></select></label>
                    {t.kind === "chart" && <label>Chart <select value={t.chartType ?? "column"} onChange={(e) => setT({ ...t, chartType: e.target.value })}>{["column", "bar", "line", "area", "pie"].map((c) => <option key={c}>{c}</option>)}</select></label>}
                    {(t.kind ?? "kpi") === "kpi" && <label>Format <select value={t.format || "number"} onChange={(e) => setT({ ...t, format: e.target.value })}><option value="number">Number</option><option value="currency">Currency</option><option value="percent">Percent</option></select></label>}
                </div>
                <label>Query (SQL)
                    <textarea rows={5} spellCheck={false} value={t.sql ?? ""} placeholder={(t.kind ?? "kpi") === "kpi" ? "SELECT SUM(Sales) FROM sales" : "SELECT Region, SUM(Sales) AS Sales FROM sales GROUP BY Region ORDER BY Sales DESC"} onChange={(e) => setT({ ...t, sql: e.target.value })} />
                </label>
                <p className="dash-muted">{(t.kind ?? "kpi") === "kpi" ? "A KPI shows the first number of the answer." : "A chart uses the first text column as categories and the number columns as series."} The agent can write these queries for you.</p>
                <div className="dash-row dash-buttons">
                    <button type="button" className="danger" onClick={onDelete}>Delete tile</button>
                    <span style={{ flex: 1 }} />
                    <button type="button" onClick={onClose}>Cancel</button>
                    <button type="button" className="primary" onClick={() => onSave({ ...t, sql: t.sql?.trim() || null })}>Save</button>
                </div>
            </div>
        </div>
    );
}

export function DashboardDesigner(): React.ReactElement {
    const [boards, setBoards] = React.useState<Dashboard[]>([]);
    const [id, setId] = React.useState("");
    const [values, setValues] = React.useState<Record<string, TileValue>>({});
    const [design, setDesign] = React.useState(false);
    const [editing, setEditing] = React.useState<Tile | null>(null);
    const [message, setMessage] = React.useState<string | null>(null);
    const [width, setWidth] = React.useState(1000);
    const gridRef = React.useRef<HTMLDivElement | null>(null);
    const found = boards.find((b) => b.id === id) ?? boards[0];
    const board = React.useMemo(() => (found ? { ...found, tiles: layout(found.tiles) } : undefined), [found]);

    const loadBoards = React.useCallback(async () => { const r = await call<any>("data.dashboard.list"); setBoards(r.dashboards ?? []); }, []);
    const loadValues = React.useCallback(async (bid?: string) => {
        if (!bid) return;
        try { const r = await call<any>("data.dashboard.values", { id: bid }); setValues(Object.fromEntries((r.values ?? []).map((v: TileValue) => [v.tileId, v]))); } catch (e: any) { setMessage(e?.message ?? String(e)); }
    }, []);
    React.useEffect(() => { void loadBoards(); }, [loadBoards]);
    React.useEffect(() => { void loadValues(board?.id); return on("data.loaded", () => void loadValues(board?.id)); }, [board?.id, loadValues]);
    React.useEffect(() => {
        const el = gridRef.current; if (!el) return;
        const ro = new ResizeObserver(() => setWidth(el.clientWidth)); ro.observe(el); setWidth(el.clientWidth);
        return () => ro.disconnect();
    }, [board?.id]);

    const cell = (width - GAP * (COLS - 1)) / COLS;
    const save = async (next: Dashboard) => {
        setBoards((all) => all.some((b) => b.id === next.id) ? all.map((b) => (b.id === next.id ? next : b)) : [...all, next]);
        try { await call("data.dashboard.save", { dashboard: next }); await loadValues(next.id); } catch (e: any) { setMessage(e?.message ?? String(e)); }
    };
    const create = async () => {
        const name = window.prompt("Dashboard name", "Overview"); if (!name) return;
        const d: Dashboard = { id: `${slug(name)}-${Date.now().toString(36)}`, name, tiles: [] };
        await save(d); setId(d.id); setDesign(true);
    };
    const addTile = (kind: "kpi" | "chart" | "table") => {
        if (!board) return;
        const t: Tile = { id: Date.now().toString(36), title: kind === "kpi" ? "New number" : kind === "chart" ? "New chart" : "New table", spec: { source: { kind: "table", id: "" } } as any, measure: "rows", format: "number",
            kind, chartType: kind === "chart" ? "column" : null, sql: null, x: 0, y: nextRow(board.tiles), w: kind === "kpi" ? 3 : kind === "chart" ? 6 : 12, h: kind === "kpi" ? 2 : 4 };
        void save({ ...board, tiles: [...board.tiles, t] }); setEditing(t);
    };
    const updateTile = (t: Tile) => board && void save({ ...board, tiles: settle(board.tiles.map((x) => (x.id === t.id ? t : x)), t.id) });
    const drag = (e: React.MouseEvent, t: Tile, resize: boolean) => {
        if (!design || (e.target as HTMLElement).closest("button")) return;
        e.preventDefault();
        const sx = e.clientX, sy = e.clientY, start = { x: t.x ?? 0, y: t.y ?? 0, w: t.w ?? 3, h: t.h ?? 2 };
        let last = start;
        const move = (m: MouseEvent) => {
            const dx = Math.round((m.clientX - sx) / (cell + GAP)), dy = Math.round((m.clientY - sy) / (ROW + GAP));
            last = resize ? { ...start, w: Math.max(2, Math.min(COLS - start.x, start.w + dx)), h: Math.max(1, start.h + dy) }
                : { ...start, x: Math.max(0, Math.min(COLS - start.w, start.x + dx)), y: Math.max(0, start.y + dy) };
            setBoards((all) => all.map((b) => (b.id === board!.id ? { ...b, tiles: settle(b.tiles.map((x) => (x.id === t.id ? { ...x, ...last } : x)), t.id) } : b)));
        };
        const up = () => {
            window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up);
            if (last !== start) void save({ ...board!, tiles: settle(board!.tiles.map((x) => (x.id === t.id ? { ...x, ...last } : x)), t.id) });
        };
        window.addEventListener("mousemove", move); window.addEventListener("mouseup", up);
    };

    return (
        <div className="dash">
            <div className="section-heading">
                <div><h1>Dashboards</h1><p>Your numbers at a glance. They update when new data loads.</p></div>
                <div className="dash-actions">
                    <select aria-label="Dashboard" value={board?.id ?? ""} onChange={(e) => setId(e.target.value)}>{boards.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}</select>
                    <button type="button" onClick={() => void create()}>New dashboard</button>
                    {board && <button type="button" className={design ? "primary" : ""} aria-pressed={design} onClick={() => setDesign((d) => !d)}>{design ? "Done" : "Design"}</button>}
                    {board && <button type="button" onClick={() => void loadValues(board.id)}>Refresh</button>}
                </div>
            </div>
            {message && <p role="alert" className="dash-error" onClick={() => setMessage(null)}>{message}</p>}
            {design && board && (
                <div className="dash-toolbar" role="toolbar" aria-label="Add a tile">
                    <span>Add:</span>
                    <button type="button" onClick={() => addTile("kpi")}>+ KPI</button>
                    <button type="button" onClick={() => addTile("chart")}>+ Chart</button>
                    <button type="button" onClick={() => addTile("table")}>+ Table</button>
                    <span className="dash-muted">Drag a tile to move it, drag its corner to resize, click ✎ to edit.</span>
                </div>
            )}
            {!board && <div className="home-empty"><p>No dashboard yet. Make one, or ask the agent to build one from your data.</p><button type="button" onClick={() => void create()}>New dashboard</button></div>}
            {board && (
                <div className={"dash-grid" + (design ? " designing" : "")} ref={gridRef} style={{ gridTemplateColumns: `repeat(${COLS}, 1fr)`, gridAutoRows: ROW, gap: GAP }}>
                    {board.tiles.map((t) => {
                        const w = t.w ?? 3, h = t.h ?? 2;
                        return (
                            <article key={t.id} className={"dash-tile kind-" + (t.kind ?? "kpi")} data-testid="dash-tile" style={{ gridColumn: `${(t.x ?? 0) + 1} / span ${w}`, gridRow: `${(t.y ?? 0) + 1} / span ${h}` }}
                                onMouseDown={(e) => drag(e, t, false)}>
                                <header><h2>{t.title}</h2>{design && <><button type="button" aria-label={`Edit ${t.title}`} onClick={() => setEditing(t)}>✎</button><button type="button" aria-label={`Remove ${t.title}`} onClick={() => void save({ ...board, tiles: board.tiles.filter((x) => x.id !== t.id) })}>×</button></>}</header>
                                <div className="dash-tile-body"><TileBody tile={t} value={values[t.id]} width={w * cell + (w - 1) * GAP - 30} height={h * ROW + (h - 1) * GAP - 52} /></div>
                                {design && <div className="dash-resize" onMouseDown={(e) => { e.stopPropagation(); drag(e, t, true); }} aria-hidden />}
                            </article>
                        );
                    })}
                </div>
            )}
            {editing && board && <TileEditor tile={editing} onClose={() => setEditing(null)} onDelete={() => { void save({ ...board, tiles: board.tiles.filter((x) => x.id !== editing.id) }); setEditing(null); }}
                onSave={(t) => { updateTile(t); setEditing(null); }} />}
        </div>
    );
}
