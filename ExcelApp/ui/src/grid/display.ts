/** Fit the displayed value without clipping significant leading digits. */
export function fitCellText(text: string, kind: number, width: number, measure: (s: string) => number, numberFormat = "General"): string {
    const available = Math.max(0, width - 8);
    const fits = (candidate: string) => measure(candidate) <= available;
    if (kind !== 1 || fits(text)) return text;
    // Only General may sacrifice precision. Fixed formats (including dates/times)
    // must retain their requested representation or display overflow hashes.
    if (numberFormat.trim().toLowerCase() === "general" && /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(text)) {
        const value = Number(text);
        if (Number.isFinite(value)) {
            const decimals = (text.split(/[eE]/)[0].split(".")[1] ?? "").length;
            for (let places = Math.abs(value) >= 1e21 ? -1 : Math.min(15, decimals); places >= 0; places--) {
                const rounded = value.toFixed(places);
                // A nonzero value must not disappear into a rounded zero.
                if (Number(rounded) === 0 && value !== 0) continue;
                const candidate = places === 0 ? rounded : rounded.replace(/\.?0+$/, "");
                if (fits(candidate)) return candidate;
            }
            for (let places = 14; places >= 0; places--) {
                const [mantissa, exponent] = value.toExponential(places).split("e");
                const digits = places === 0 ? mantissa : mantissa.replace(/\.?0+$/, "");
                const candidate = `${digits}E${exponent[0]}${exponent.slice(1).padStart(2, "0")}`;
                if (fits(candidate)) return candidate;
            }
        }
    }
    return "#".repeat(Math.max(1, Math.floor(available / Math.max(1, measure("#")))));
}
export function frozenCount(requested: number, viewRows: number): number {
    return Math.max(0, Math.min(requested, viewRows));
}
export function bodyRow(row: number, frozen: number): number { return row + frozen; }
