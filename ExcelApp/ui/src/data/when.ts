// Times from the host arrive as ISO text (often UTC, with fractions of a second). Pages show them in local time,
// to the minute, the same way everywhere: "2026-11-08 09:30". A date with no time stays a date.
const two = (n: number) => String(n).padStart(2, "0");

export function whenText(value: string | null | undefined, never = "Never"): string {
    if (!value) return never;
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return value;
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) return value;
    return `${d.getFullYear()}-${two(d.getMonth() + 1)}-${two(d.getDate())} ${two(d.getHours())}:${two(d.getMinutes())}`;
}
