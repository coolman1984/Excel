import { describe, it, expect } from "vitest";
import { fitCellText, bodyRow, frozenCount } from "./display";
describe("Excel cell display", () => {
    const measure = (s: string) => s.length * 8;
    it("replaces overflowing numbers with hashes instead of losing leading digits", () => {
        expect(fitCellText("40416.259", 1, 64, measure)).toBe("#######");
    });
    it("keeps numbers that fit and lets the canvas clip text", () => {
        expect(fitCellText("123", 1, 32, measure)).toBe("123");
        expect(fitCellText("long text", 2, 32, measure)).toBe("long text");
    });
    it("never repeats frozen rows in the body, including a short filtered view", () => {
        expect(bodyRow(0, frozenCount(2, 100))).toBe(2);
        expect(frozenCount(10, 1)).toBe(1);
        expect(frozenCount(2, 0)).toBe(0);
    });
});
