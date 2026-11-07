// The Excel colour palette (plan step 9): the 10 theme colours with 5 tints each,
// plus the 10 standard colours. Hex values are the public Office default theme.

/** Blend `hex` toward white by `t` (0..1). Returns "#RRGGBB". */
export function tint(hex: string, t: number): string {
    const n = parseInt(hex.slice(1), 16);
    const r = (n >> 16) & 0xff;
    const g = (n >> 8) & 0xff;
    const b = n & 0xff;
    const mix = (c: number) => Math.round(c + (255 - c) * t);
    const to2 = (v: number) => v.toString(16).padStart(2, "0").toUpperCase();
    return "#" + to2(mix(r)) + to2(mix(g)) + to2(mix(b));
}

export interface Swatch {
    hex: string;
    name: string;
}

/** One theme colour row: the base colour plus 5 lighter tints, Excel order. */
export interface ThemeRow {
    name: string;
    base: string;
    tints: string[]; // 5, lightest last
}

const THEME: { name: string; hex: string }[] = [
    { name: "White, Background 1", hex: "#FFFFFF" },
    { name: "Black, Text 1", hex: "#000000" },
    { name: "Light Gray, Background 2", hex: "#E7E6E6" },
    { name: "Dark Gray, Text 2", hex: "#44546A" },
    { name: "Blue, Accent 1", hex: "#4472C4" },
    { name: "Orange, Accent 2", hex: "#ED7D31" },
    { name: "Gray, Accent 3", hex: "#A5A5A5" },
    { name: "Gold, Accent 4", hex: "#FFC000" },
    { name: "Light Blue, Accent 5", hex: "#5B9BD5" },
    { name: "Green, Accent 6", hex: "#70AD47" },
];

/** Standard colours row (Excel's fixed 10). */
export const STANDARD: Swatch[] = [
    { hex: "#C00000", name: "Dark Red" },
    { hex: "#FF0000", name: "Red" },
    { hex: "#FFC000", name: "Orange" },
    { hex: "#FFFF00", name: "Yellow" },
    { hex: "#92D050", name: "Light Green" },
    { hex: "#00B050", name: "Green" },
    { hex: "#00B0F0", name: "Light Blue" },
    { hex: "#0070C0", name: "Blue" },
    { hex: "#002060", name: "Dark Blue" },
    { hex: "#7030A0", name: "Purple" },
];

/** The 10 theme rows, each with its base colour and 5 tints (blended toward white). */
export const THEME_ROWS: ThemeRow[] = THEME.map((c) => ({
    name: c.name,
    base: c.hex,
    tints: [0.2, 0.4, 0.6, 0.8, 0.9].map((t) => tint(c.hex, t)),
}));

/** Flat list of every base colour for "recent"/quick pick. */
export const BASE_COLOURS: string[] = THEME.map((c) => c.hex);
