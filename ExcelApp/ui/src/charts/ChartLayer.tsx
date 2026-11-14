// Chart cards floating over the sheet: drag by the header, resize from the corner, change type and title, show the
// numbers as a table, delete. Every change goes to the host (chart.update / chart.delete), so undo and save see it.
import * as React from "react";
import { call } from "../bridge";
import { ChartSvg, ChartTable, type ChartModel } from "./ChartSvg";

export interface ChartWire extends ChartModel { id: string; sheet: number; title: string | null; x: number; y: number; w: number; h: number; range: string; truncated: boolean }

const TYPES: [string, string][] = [["column", "Column"], ["bar", "Bar"], ["line", "Line"], ["area", "Area"], ["pie", "Pie"], ["scatter", "Scatter"]];

function Card({ c, onChanged }: { c: ChartWire; onChanged: () => void }): React.ReactElement {
    const [box, setBox] = React.useState({ x: c.x, y: c.y, w: c.w, h: c.h });
    const [table, setTable] = React.useState(false);
    const [title, setTitle] = React.useState(c.title ?? "");
    const [editing, setEditing] = React.useState(false);
    React.useEffect(() => { setBox({ x: c.x, y: c.y, w: c.w, h: c.h }); setTitle(c.title ?? ""); }, [c.x, c.y, c.w, c.h, c.title]);
    const update = (args: object) => { void call("chart.update", { id: c.id, ...args }).then(onChanged).catch(() => onChanged()); };
    const drag = (e: React.MouseEvent, resize: boolean) => {
        if ((e.target as HTMLElement).closest("input,select,button")) return;
        e.preventDefault();
        const sx = e.clientX, sy = e.clientY, start = { ...box };
        let last = start;
        const move = (m: MouseEvent) => {
            last = resize
                ? { ...start, w: Math.max(260, start.w + m.clientX - sx), h: Math.max(190, start.h + m.clientY - sy) }
                : { ...start, x: Math.max(0, start.x + m.clientX - sx), y: Math.max(0, start.y + m.clientY - sy) };
            setBox(last);
        };
        const up = () => { window.removeEventListener("mousemove", move); window.removeEventListener("mouseup", up); if (last !== start) update(last); };
        window.addEventListener("mousemove", move); window.addEventListener("mouseup", up);
    };
    const shown = title.trim() || (c.series.length === 1 ? c.series[0].name : "Chart");
    return (
        <section className="chartCard" style={{ left: box.x, top: box.y, width: box.w, height: box.h }} data-testid="chart-card" aria-label={`Chart: ${shown}`}>
            <header className="chartHead" onMouseDown={(e) => drag(e, false)}>
                {editing
                    ? <input className="chartTitle" value={title} placeholder={shown} aria-label="Chart title" autoFocus
                        onChange={(e) => setTitle(e.target.value)}
                        onBlur={() => { setEditing(false); if ((title.trim() || null) !== c.title) update({ title: title.trim() || null }); }}
                        onKeyDown={(e) => { if (e.key === "Enter") (e.target as HTMLInputElement).blur(); if (e.key === "Escape") { setTitle(c.title ?? ""); setEditing(false); } }} />
                    : <span className="chartTitleText" title="Double-click to rename; drag to move" onDoubleClick={() => setEditing(true)}>{shown}</span>}
                <select value={c.type} aria-label="Chart type" onChange={(e) => update({ type: e.target.value })}>
                    {TYPES.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                </select>
                <button type="button" className={table ? "on" : ""} onClick={() => setTable((t) => !t)} title="Show the numbers as a table">Data</button>
                <button type="button" aria-label="Delete chart" title="Delete chart" onClick={() => void call("chart.delete", { id: c.id }).then(onChanged)}>×</button>
            </header>
            <div className="chartBody">
                {table ? <ChartTable model={c} /> : <ChartSvg model={c} width={box.w - 16} height={box.h - 64} />}
            </div>
            <footer className="chartFoot">{c.range}{c.truncated ? " · showing the first 1,000 rows and 8 series" : ""}</footer>
            <div className="chartResize" onMouseDown={(e) => drag(e, true)} aria-hidden />
        </section>
    );
}

export function ChartLayer({ charts, onChanged }: { charts: ChartWire[]; onChanged: () => void }): React.ReactElement | null {
    if (charts.length === 0) return null;
    return <div className="chartLayer">{charts.map((c) => <Card key={c.id} c={c} onChanged={onChanged} />)}</div>;
}
