import { describe, it, expect } from "vitest";
import { tileChart, nextRow, layout, settle } from "./DashboardDesigner";

describe("dashboard designer", () => {
    it("a tile's rows become a chart: text column as categories, number columns as series", () => {
        const m = tileChart({ columns: [{ name: "Region", type: "text" }, { name: "Sales", type: "number" }, { name: "Cost", type: "number" }], rows: [["North", 10, 4], ["South", 20, null]], rowCount: 2, truncated: false }, "column");
        expect(m.categories).toEqual(["North", "South"]);
        expect(m.series.map((s) => s.name)).toEqual(["Sales", "Cost"]);
        expect(m.series[1].values).toEqual([4, null]);
    });
    it("numbers only: categories are numbered", () => {
        expect(tileChart({ columns: [{ name: "N", type: "number" }], rows: [[1], [2]], rowCount: 2, truncated: false }, "line").categories).toEqual(["1", "2"]);
    });
    it("older tiles without a place flow four to a row under the placed ones", () => {
        const out = layout([{ id: "a", x: 0, y: 0, w: 6, h: 2 } as any, ...[1, 2, 3, 4, 5].map((i) => ({ id: "t" + i, x: -1, y: -1 } as any))]);
        expect(out.slice(1).map((t) => [t.x, t.y])).toEqual([[0, 2], [3, 2], [6, 2], [9, 2], [0, 4]]);
        expect(out[0]).toMatchObject({ x: 0, y: 0 });
    });
    it("a tile that grows pushes the tiles it covers down, in a chain", () => {
        const tiles = [{ id: "a", x: 0, y: 0, w: 6, h: 6 }, { id: "b", x: 0, y: 4, w: 12, h: 3 }, { id: "c", x: 0, y: 7, w: 3, h: 2 }, { id: "d", x: 6, y: 0, w: 6, h: 2 }] as any;
        const out = settle(tiles, "a");
        expect(out.find((t) => t.id === "b")).toMatchObject({ y: 6 });
        expect(out.find((t) => t.id === "c")).toMatchObject({ y: 9 });
        expect(out.find((t) => t.id === "d")).toMatchObject({ y: 0 });
    });
    it("a new tile goes under the others", () => {
        expect(nextRow([])).toBe(0);
        expect(nextRow([{ x: 0, y: 0, h: 2 } as any, { x: 3, y: 2, h: 4 } as any])).toBe(6);
    });
});
