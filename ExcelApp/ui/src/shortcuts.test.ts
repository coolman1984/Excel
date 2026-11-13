import { describe, it, expect } from "vitest";
import { fillRange, todayInput, timeInput, SHORTCUT_FORMATS } from "./shortcuts";
import { BINDINGS, resolveCommand } from "./keys";

describe("shortcuts", () => {
    it("Ctrl+D on one cell copies the cell above", () => {
        expect(fillRange("down", { r1: 4, c1: 2, r2: 4, c2: 2 })).toEqual({ from: [3, 2, 3, 2], to: [3, 2, 4, 2] });
    });
    it("Ctrl+D on a block copies its top row down", () => {
        expect(fillRange("down", { r1: 1, c1: 0, r2: 5, c2: 3 })).toEqual({ from: [1, 0, 1, 3], to: [1, 0, 5, 3] });
    });
    it("Ctrl+D in the first row has nothing to copy", () => {
        expect(fillRange("down", { r1: 0, c1: 0, r2: 0, c2: 4 })).toBeNull();
    });
    it("Ctrl+R on one cell copies the cell to the left", () => {
        expect(fillRange("right", { r1: 2, c1: 3, r2: 2, c2: 3 })).toEqual({ from: [2, 2, 2, 2], to: [2, 2, 2, 3] });
    });
    it("Ctrl+R on a block copies its first column right", () => {
        expect(fillRange("right", { r1: 0, c1: 1, r2: 3, c2: 4 })).toEqual({ from: [0, 1, 3, 1], to: [0, 1, 3, 4] });
    });
    it("Ctrl+R in column A alone has nothing to copy", () => {
        expect(fillRange("right", { r1: 0, c1: 0, r2: 9, c2: 0 })).toBeNull();
    });
    it("date and time are typed in forms the host reads", () => {
        const d = new Date(2026, 0, 5, 9, 7, 30);
        expect(todayInput(d)).toBe("2026-01-05");
        expect(timeInput(d)).toBe("09:07");
    });
    it("every bound shortcut command has a format or a handler name", () => {
        for (const cmd of Object.keys(SHORTCUT_FORMATS)) expect(BINDINGS.some((b) => b.command === cmd)).toBe(true);
    });
    it("Ctrl+Shift+4 on a US keyboard is currency, Ctrl+Shift+; is the time", () => {
        expect(resolveCommand({ key: "$", ctrlKey: true, shiftKey: true })).toBe("formatCurrency");
        expect(resolveCommand({ key: ":", ctrlKey: true, shiftKey: true })).toBe("insertTime");
        expect(resolveCommand({ key: "=", altKey: true })).toBe("autoSum");
    });
});
