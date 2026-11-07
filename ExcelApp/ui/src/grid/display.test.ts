import { describe, it, expect } from "vitest";
import { fitCellText, bodyRow, frozenCount } from "./display";
describe("Excel cell display", () => {
    const measure = (s: string) => s.length * 8;
    it("replaces overflowing numbers with hashes instead of losing leading digits", () => {
        expect(fitCellText("40416.259", 1, 64, measure, "0.000")).toBe("#######");
    });
    it("keeps numbers that fit and lets the canvas clip text", () => {
        expect(fitCellText("123", 1, 32, measure)).toBe("123");
        expect(fitCellText("long text", 2, 32, measure)).toBe("long text");
    });
    it("rounds General decimals to the most precise value that fits", () => {
        expect(fitCellText("24866.858", 1, 72, measure)).toBe("24866.86");
        expect(fitCellText("24866.858", 1, 48, measure, "General")).toBe("24867");
        expect(fitCellText("9.999", 1, 24, measure)).toBe("10");
    });
    it("uses scientific notation for General integers that cannot fit", () => {
        expect(fitCellText("12345678901", 1, 64, measure)).toBe("1.2E+10");
        expect(fitCellText("1234567890123456789012345", 1, 64, measure)).toBe("1.2E+24");
        expect(fitCellText("12345678901", 1, 48, measure)).toBe("1E+10");
        expect(fitCellText("12345678901", 1, 32, measure)).toBe("###");
    });
    it("keeps fixed number, percent, currency and date/time formats as hashes", () => {
        for (const [text, format] of [["24866.86", "0.00"], ["12.50%", "0.00%"], ["$123.45", "$0.00"], ["2026-11-07", "yyyy-mm-dd"], ["12:34:56", "hh:mm:ss"]]) {
            expect(fitCellText(text, 1, 48, measure, format)).toBe("#####");
        }
    });
    it("preserves negative signs when rounding and using scientific notation", () => {
        expect(fitCellText("-24866.858", 1, 72, measure)).toBe("-24866.9");
        expect(fitCellText("-24866.858", 1, 56, measure)).toBe("-24867");
        expect(fitCellText("-12345678901", 1, 72, measure)).toBe("-1.2E+10");
        expect(fitCellText("-0.000000001234", 1, 72, measure)).toBe("-1.2E-09");
    });
    it("never repeats frozen rows in the body, including a short filtered view", () => {
        expect(bodyRow(0, frozenCount(2, 100))).toBe(2);
        expect(frozenCount(10, 1)).toBe(1);
        expect(frozenCount(2, 0)).toBe(0);
    });
});
