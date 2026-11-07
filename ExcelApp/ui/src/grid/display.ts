/** Numeric overflow must never look like a different number. Text is clipped by canvas. */
export function fitCellText(text: string, kind: number, width: number, measure: (s: string) => number): string {
    const available = Math.max(0, width - 8);
    if (kind !== 1 || measure(text) <= available) return text;
    return "#".repeat(Math.max(1, Math.floor(available / Math.max(1, measure("#")))));
}
export function frozenCount(requested: number, viewRows: number): number {
    return Math.max(0, Math.min(requested, viewRows));
}
export function bodyRow(row: number, frozen: number): number { return row + frozen; }
