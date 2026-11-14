// Axis arithmetic for the hand-drawn charts: round tick steps (1, 2, 5 x 10^n) that always include zero for bars and areas.
export const SERIES_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"];

export function niceTicks(min: number, max: number, target = 5): number[] {
    if (!isFinite(min) || !isFinite(max)) return [0, 1];
    if (min === max) { if (min === 0) return [0, 1]; const pad = Math.abs(min) * 0.1; min -= pad; max += pad; }
    const raw = (max - min) / Math.max(1, target);
    const mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= raw) ?? 10 * mag;
    const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
    const ticks: number[] = [];
    for (let v = lo; v <= hi + step / 2; v += step) ticks.push(Math.abs(v) < step / 1e6 ? 0 : +v.toPrecision(12));
    return ticks;
}

const compact = new Intl.NumberFormat(undefined, { notation: "compact", maximumFractionDigits: 1 });
const plain = new Intl.NumberFormat(undefined, { maximumFractionDigits: 2 });
/** Axis labels short (12K), tooltips exact to two decimals. */
export const shortNumber = (v: number) => (Math.abs(v) >= 10000 ? compact.format(v) : plain.format(v));
export const fullNumber = (v: number | null | undefined) => (v == null ? "—" : plain.format(v));

/** Show every k-th category label so labels about `px` wide do not collide in `width`. */
export function labelStep(count: number, width: number, px = 64): number {
    return Math.max(1, Math.ceil((count * px) / Math.max(1, width)));
}

/** Pie slices: the largest seven and "Other" for the rest, positive values only. */
export function pieSlices(categories: string[], values: (number | null)[]): { name: string; value: number }[] {
    const items = categories.map((name, i) => ({ name, value: values[i] ?? 0 })).filter((x) => x.value > 0).sort((a, b) => b.value - a.value);
    if (items.length <= 8) return items;
    const head = items.slice(0, 7), rest = items.slice(7).reduce((s, x) => s + x.value, 0);
    return [...head, { name: "Other", value: rest }];
}
