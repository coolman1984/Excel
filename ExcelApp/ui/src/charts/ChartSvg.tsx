// Hand-drawn SVG charts (no chart package): column, bar, line, area, pie, scatter. Thin marks, rounded data ends on the
// baseline, 2 px lines, a recessive grid, one value axis, a legend only for two or more series, and a hover tooltip.
import * as React from "react";
import { SERIES_COLORS, niceTicks, shortNumber, fullNumber, labelStep, pieSlices } from "./scale";

export interface ChartSeries { name: string; values: (number | null)[] }
export interface ChartModel { type: string; categories: string[]; series: ChartSeries[]; xValues?: (number | null)[] | null }

const INK = "#0b0b0b", INK2 = "#52514e", MUTED = "#8a8985", GRID = "#e6e5e1", SURFACE = "#fcfcfb";
const color = (i: number) => SERIES_COLORS[i % SERIES_COLORS.length];

/** A bar with 4 px rounded corners at its data end only (top for positive, bottom for negative). */
function barPath(x: number, y0: number, w: number, y1: number, horizontal = false): string {
    const r = Math.min(4, w / 2, Math.abs(y1 - y0));
    if (!horizontal) {
        const up = y1 < y0, top = Math.min(y0, y1), bot = Math.max(y0, y1);
        return up
            ? `M${x},${bot}V${top + r}Q${x},${top} ${x + r},${top}H${x + w - r}Q${x + w},${top} ${x + w},${top + r}V${bot}Z`
            : `M${x},${top}V${bot - r}Q${x},${bot} ${x + r},${bot}H${x + w - r}Q${x + w},${bot} ${x + w},${bot - r}V${top}Z`;
    }
    // horizontal: x = baseline position, y0..y1 along x; w = thickness, passed as y
    const right = y1 > y0, l = Math.min(y0, y1), rr = Math.max(y0, y1);
    return right
        ? `M${l},${x}H${rr - r}Q${rr},${x} ${rr},${x + r}V${x + w - r}Q${rr},${x + w} ${rr - r},${x + w}H${l}Z`
        : `M${rr},${x}H${l + r}Q${l},${x} ${l},${x + r}V${x + w - r}Q${l},${x + w} ${l + r},${x + w}H${rr}Z`;
}

export function ChartSvg({ model, width, height }: { model: ChartModel; width: number; height: number }): React.ReactElement {
    const [hover, setHover] = React.useState<{ x: number; y: number; title: string; rows: { name: string; value: string; color: string }[] } | null>(null);
    const { type, categories } = model;
    const series = type === "scatter" ? model.series.slice(0, 3) : model.series;   // more than three scatter colours cannot be told apart
    const legend = type === "pie" || series.length >= 2;
    const legendItems = type === "pie" ? pieSlices(categories, series[0]?.values ?? []).map((s) => s.name) : series.map((s) => s.name);
    const top = legend ? 30 : 10, right = 14, bottom = type === "pie" ? 10 : 34;
    const allValues = series.flatMap((s) => s.values).filter((v): v is number => v != null);
    const empty = allValues.length === 0;

    let body: React.ReactNode = null;
    const tip = (e: React.MouseEvent, title: string, rows: { name: string; value: string; color: string }[]) => {
        const box = (e.currentTarget as SVGElement).ownerSVGElement?.getBoundingClientRect() ?? (e.currentTarget as SVGElement).getBoundingClientRect();
        setHover({ x: e.clientX - box.left, y: e.clientY - box.top, title, rows });
    };

    if (empty) {
        body = <text x={width / 2} y={height / 2} textAnchor="middle" fill={MUTED} fontSize={13}>No numbers in this range</text>;
    } else if (type === "pie") {
        const slices = pieSlices(categories, series[0]?.values ?? []), total = slices.reduce((s, x) => s + x.value, 0);
        const cx = width / 2, cy = top + (height - top - bottom) / 2, rad = Math.max(10, Math.min(width, height - top - bottom) / 2 - 8);
        let a0 = -Math.PI / 2;
        body = slices.map((s, i) => {
            const a1 = a0 + (s.value / total) * Math.PI * 2, large = a1 - a0 > Math.PI ? 1 : 0;
            const p = slices.length === 1
                ? `M${cx},${cy - rad}A${rad},${rad} 0 1 1 ${cx - 0.01},${cy - rad}Z`
                : `M${cx},${cy}L${cx + rad * Math.cos(a0)},${cy + rad * Math.sin(a0)}A${rad},${rad} 0 ${large} 1 ${cx + rad * Math.cos(a1)},${cy + rad * Math.sin(a1)}Z`;
            a0 = a1;
            const pct = ((s.value / total) * 100).toFixed(1) + "%";
            return <path key={i} d={p} fill={color(i)} stroke={SURFACE} strokeWidth={2} onMouseMove={(e) => tip(e, s.name, [{ name: series[0].name, value: `${fullNumber(s.value)} (${pct})`, color: color(i) }])} onMouseLeave={() => setHover(null)} />;
        });
    } else {
        const horizontal = type === "bar", scatter = type === "scatter";
        const zeroBased = type === "column" || type === "bar" || type === "area";
        const lo = Math.min(...allValues, zeroBased ? 0 : Infinity), hi = Math.max(...allValues, zeroBased ? 0 : -Infinity);
        const ticks = niceTicks(lo, hi);
        const vmin = ticks[0], vmax = ticks[ticks.length - 1];
        const xs = scatter ? (model.xValues ?? categories.map((_, i) => i + 1)).map((v) => (v == null ? null : v)) : null;
        const xNums = xs?.filter((v): v is number => v != null) ?? [];
        const xticks = scatter ? niceTicks(Math.min(...xNums), Math.max(...xNums)) : [];
        const valueLabelW = Math.max(...ticks.map((t) => shortNumber(t).length)) * 7 + 10;
        const catLabelW = horizontal ? Math.min(140, Math.max(...categories.map((c) => c.length), 1) * 7 + 10) : 0;
        const left = horizontal ? catLabelW : valueLabelW;
        const pw = Math.max(10, width - left - right), ph = Math.max(10, height - top - bottom);
        const vpos = (v: number) => (horizontal ? left + ((v - vmin) / (vmax - vmin)) * pw : top + ph - ((v - vmin) / (vmax - vmin)) * ph);
        const n = categories.length, band = (horizontal ? ph : pw) / Math.max(1, n);
        const cpos = (i: number) => (horizontal ? top + i * band : left + i * band);
        const xpos = (v: number) => left + ((v - xticks[0]) / (xticks[xticks.length - 1] - xticks[0] || 1)) * pw;
        const step = labelStep(n, horizontal ? ph : pw, horizontal ? 18 : 64);
        const grid = ticks.map((t) => {
            const p = vpos(t);
            return horizontal
                ? <g key={"g" + t}><line x1={p} x2={p} y1={top} y2={top + ph} stroke={t === 0 ? MUTED : GRID} /><text x={p} y={top + ph + 16} textAnchor="middle" fill={INK2} fontSize={11}>{shortNumber(t)}</text></g>
                : <g key={"g" + t}><line x1={left} x2={left + pw} y1={p} y2={p} stroke={t === 0 ? MUTED : GRID} /><text x={left - 6} y={p + 4} textAnchor="end" fill={INK2} fontSize={11}>{shortNumber(t)}</text></g>;
        });
        const catLabels = scatter
            ? xticks.map((t) => <text key={"x" + t} x={xpos(t)} y={top + ph + 16} textAnchor="middle" fill={INK2} fontSize={11}>{shortNumber(t)}</text>)
            : categories.map((c, i) => i % step !== 0 ? null : horizontal
                ? <text key={"c" + i} x={left - 6} y={cpos(i) + band / 2 + 4} textAnchor="end" fill={INK2} fontSize={11}>{c.length > 18 ? c.slice(0, 17) + "…" : c}</text>
                : <text key={"c" + i} x={cpos(i) + band / 2} y={top + ph + 16} textAnchor="middle" fill={INK2} fontSize={11}>{c.length > 12 ? c.slice(0, 11) + "…" : c}</text>);
        const rowsAt = (i: number) => series.map((s, k) => ({ name: s.name, value: fullNumber(s.values[i]), color: color(k) }));
        let marks: React.ReactNode;
        if (type === "column" || horizontal) {
            const inner = Math.max(1, band * 0.72), bw = Math.max(1, (inner - 2 * (series.length - 1)) / series.length), z = vpos(0);
            marks = series.map((s, k) => s.values.map((v, i) => {
                if (v == null) return null;
                const at = cpos(i) + (band - inner) / 2 + k * (bw + 2);
                return <path key={k + ":" + i} d={barPath(at, z, bw, vpos(v), horizontal)} fill={color(k)} />;
            }));
        } else if (scatter) {
            marks = series.map((s, k) => s.values.map((v, i) => v == null || xs![i] == null ? null :
                <circle key={k + ":" + i} cx={xpos(xs![i]!)} cy={vpos(v)} r={4} fill={color(k)} stroke={SURFACE} strokeWidth={2}
                    onMouseMove={(e) => tip(e, `x ${fullNumber(xs![i])}`, [{ name: s.name, value: fullNumber(v), color: color(k) }])} onMouseLeave={() => setHover(null)} />));
        } else {
            marks = series.map((s, k) => {
                const pts = s.values.map((v, i) => (v == null ? null : [cpos(i) + band / 2, vpos(v)] as [number, number]));
                const runs: [number, number][][] = []; let cur: [number, number][] = [];
                for (const p of pts) { if (p) cur.push(p); else if (cur.length) { runs.push(cur); cur = []; } }
                if (cur.length) runs.push(cur);
                return <g key={"s" + k}>
                    {type === "area" && runs.map((r, j) => <path key={"a" + j} d={`M${r[0][0]},${vpos(0)}L${r.map((p) => p.join(",")).join("L")}L${r[r.length - 1][0]},${vpos(0)}Z`} fill={color(k)} opacity={0.16} />)}
                    {runs.map((r, j) => <polyline key={"l" + j} points={r.map((p) => p.join(",")).join(" ")} fill="none" stroke={color(k)} strokeWidth={2} strokeLinejoin="round" strokeLinecap="round" />)}
                    {n <= 40 && pts.map((p, i) => p && <circle key={"m" + i} cx={p[0]} cy={p[1]} r={4} fill={color(k)} stroke={SURFACE} strokeWidth={2} />)}
                </g>;
            });
        }
        const hit = !scatter && categories.map((c, i) => horizontal
            ? <rect key={"h" + i} x={left} y={cpos(i)} width={pw} height={band} fill="transparent" onMouseMove={(e) => tip(e, c, rowsAt(i))} onMouseLeave={() => setHover(null)} />
            : <rect key={"h" + i} x={cpos(i)} y={top} width={band} height={ph} fill="transparent" onMouseMove={(e) => tip(e, c, rowsAt(i))} onMouseLeave={() => setHover(null)} />);
        body = <>{grid}{catLabels}{marks}{hit}</>;
    }

    return (
        <div className="chartSvgWrap" style={{ position: "relative", width, height }}>
            <svg width={width} height={height} role="img" aria-label={`${type} chart of ${series.map((s) => s.name).join(", ")}`} style={{ display: "block", fontFamily: "inherit" }}>
                {legend && legendItems.slice(0, 8).map((name, i, all) => {
                    const slot = Math.max(60, Math.min(160, (width - 20) / all.length));
                    return <g key={"lg" + i} transform={`translate(${10 + i * slot},10)`}>
                        <rect width={10} height={10} rx={2} fill={color(i)} />
                        <text x={14} y={9} fontSize={11} fill={INK2}>{name.length > Math.floor((slot - 20) / 6.5) ? name.slice(0, Math.max(3, Math.floor((slot - 20) / 6.5) - 1)) + "…" : name}</text>
                    </g>;
                })}
                {body}
            </svg>
            {hover && (
                <div className="chartTip" style={{ left: Math.min(hover.x + 12, width - 170), top: Math.max(0, hover.y - 10) }} role="status">
                    <div className="chartTipTitle">{hover.title}</div>
                    {hover.rows.map((r, i) => <div key={i} className="chartTipRow"><span className="chartTipKey" style={{ background: r.color }} />{r.name}<b>{r.value}</b></div>)}
                </div>
            )}
        </div>
    );
}

/** The chart's numbers as a table (for low-contrast colours, screen readers and checking values). */
export function ChartTable({ model }: { model: ChartModel }): React.ReactElement {
    return (
        <div className="chartTable">
            <table>
                <thead><tr><th>{model.type === "scatter" ? "x" : ""}</th>{model.series.map((s, i) => <th key={i}>{s.name}</th>)}</tr></thead>
                <tbody>{model.categories.slice(0, 500).map((c, i) => (
                    <tr key={i}><td>{model.type === "scatter" && model.xValues ? fullNumber(model.xValues[i]) : c}</td>{model.series.map((s, k) => <td key={k}>{fullNumber(s.values[i])}</td>)}</tr>
                ))}</tbody>
            </table>
        </div>
    );
}
export { INK };
