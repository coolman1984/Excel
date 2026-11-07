// Number-format helpers (plan step 9). The chosen codes are fixed by the plan.
// Increase and decrease decimals edit the current code's decimal places; General
// becomes "0.0" on increase.

export const GENERAL = "General";

/** The format codes the Home → Number dropdown uses (plan step 9). */
export const NUMBER_FORMATS: { label: string; code: string }[] = [
    { label: "General", code: GENERAL },
    { label: "Number", code: "#,##0.00" },
    { label: "Currency", code: "[$$-409]#,##0.00" },
    { label: "Accounting", code: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)' },
    { label: "Short Date", code: "yyyy-mm-dd" },
    { label: "Long Date", code: "dddd, mmmm d, yyyy" },
    { label: "Time", code: "h:mm:ss AM/PM" },
    { label: "Percentage", code: "0.00%" },
    { label: "Fraction", code: "# ?/?" },
    { label: "Scientific", code: "0.00E+00" },
    { label: "Text", code: "@" },
];

/** `%` button applies the percentage format. */
export const PERCENT_CODE = "0%";
/** `,` button applies the thousands separator format. */
export const COMMA_CODE = "#,##0";

/**
 * Increase the number of decimal places in a format code (Increase Decimal button).
 * General → "0.0". Adds one "0" after the existing decimals, or inserts ".0" after
 * the first numeric run when there are none.
 */
export function increaseDecimals(code: string): string {
    if (!code || code === GENERAL) return "0.0";
    if (/\.([0#]+)/.test(code)) {
        return code.replace(/\.([0#]+)/, (_s, d: string) => "." + d + "0");
    }
    const run = code.match(/[0#,]+/);
    if (run && run.index !== undefined) {
        const at = run.index + run[0].length;
        return code.slice(0, at) + ".0" + code.slice(at);
    }
    return code;
}

/**
 * Decrease the number of decimal places in a format code (Decrease Decimal button).
 * General → "0". Removes one decimal, dropping the point when the last place goes.
 */
export function decreaseDecimals(code: string): string {
    if (!code || code === GENERAL) return "0";
    const m = code.match(/\.([0#]+)/);
    if (m) {
        if (m[1].length <= 1) {
            return code.replace(/\.([0#]+)/, "");
        }
        return code.replace(/\.([0#]+)/, (_s, d: string) => "." + d.slice(0, d.length - 1));
    }
    return code;
}
