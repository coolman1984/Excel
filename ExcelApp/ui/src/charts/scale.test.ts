import { describe, it, expect } from "vitest";
import { niceTicks, labelStep, pieSlices, shortNumber } from "./scale";

describe("chart scale", () => {
    it("round steps that cover the data", () => {
        expect(niceTicks(0, 47)).toEqual([0, 10, 20, 30, 40, 50]);
        expect(niceTicks(-12, 30)[0]).toBeLessThanOrEqual(-12);
        const t = niceTicks(0.1, 0.9); expect(t[0]).toBe(0); expect(t[t.length - 1]).toBeGreaterThanOrEqual(0.9);
    });
    it("a flat series still gets an axis", () => {
        expect(niceTicks(5, 5).length).toBeGreaterThan(1);
        expect(niceTicks(0, 0)).toEqual([0, 1]);
    });
    it("labels thin out when crowded", () => {
        expect(labelStep(10, 800)).toBe(1);
        expect(labelStep(365, 600)).toBeGreaterThan(30);
    });
    it("pie keeps the seven largest and folds the rest into Other", () => {
        const s = pieSlices(Array.from({ length: 12 }, (_, i) => "c" + i), Array.from({ length: 12 }, (_, i) => i + 1));
        expect(s).toHaveLength(8); expect(s[0].value).toBe(12); expect(s[7]).toEqual({ name: "Other", value: 15 });
        expect(pieSlices(["a", "b"], [3, -1])).toEqual([{ name: "a", value: 3 }]);
    });
    it("short numbers on axes", () => {
        expect(shortNumber(1500)).toMatch(/1[,.]?5/);
        expect(shortNumber(250000)).toMatch(/250/);
    });
});
