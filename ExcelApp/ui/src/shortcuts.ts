// What the Excel shortcuts without a ribbon button do: the number formats of Ctrl+Shift+digit,
// the ranges of Ctrl+D / Ctrl+R, and the text Ctrl+; and Ctrl+Shift+; type into a cell.
import { GENERAL } from "./numberFormat";

/** Number formats applied by Ctrl+Shift+~ $ % ^ # @ !, as in Excel. */
export const SHORTCUT_FORMATS: Record<string, string> = {
    formatGeneral: GENERAL,
    formatCurrency: "[$$-409]#,##0.00",
    formatPercent: "0%",
    formatScientific: "0.00E+00",
    formatDate: "d-mmm-yy",
    formatTime: "h:mm AM/PM",
    formatNumber: "#,##0.00",
};

export interface Sel { r1: number; c1: number; r2: number; c2: number }

/**
 * Ctrl+D copies the top row of the selection down over it; with one row selected it copies the row above.
 * Ctrl+R does the same to the right. Null when there is nothing to copy from (row 1 or column A alone).
 */
export function fillRange(direction: "down" | "right", s: Sel): { from: number[]; to: number[] } | null {
    if (direction === "down") {
        if (s.r1 === s.r2) {
            if (s.r1 <= 0) return null;
            return { from: [s.r1 - 1, s.c1, s.r1 - 1, s.c2], to: [s.r1 - 1, s.c1, s.r2, s.c2] };
        }
        return { from: [s.r1, s.c1, s.r1, s.c2], to: [s.r1, s.c1, s.r2, s.c2] };
    }
    if (s.c1 === s.c2) {
        if (s.c1 <= 0) return null;
        return { from: [s.r1, s.c1 - 1, s.r2, s.c1 - 1], to: [s.r1, s.c1 - 1, s.r2, s.c2] };
    }
    return { from: [s.r1, s.c1, s.r2, s.c1], to: [s.r1, s.c1, s.r2, s.c2] };
}

const two = (n: number) => String(n).padStart(2, "0");

/** Ctrl+; : today's date as the cell input reads it (stored as a date, not a formula). */
export function todayInput(now: Date): string {
    return `${now.getFullYear()}-${two(now.getMonth() + 1)}-${two(now.getDate())}`;
}

/** Ctrl+Shift+; : the current time to the minute, as the cell input reads it. */
export function timeInput(now: Date): string {
    return `${two(now.getHours())}:${two(now.getMinutes())}`;
}
