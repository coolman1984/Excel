import { describe, it, expect } from "vitest";
import { BINDINGS, resolveCommand, eventFor, wheelZoom } from "./keys";

describe("keys", () => {
    it("maps every binding to its own command", () => {
        expect(BINDINGS.length).toBeGreaterThan(30);
        for (const b of BINDINGS) {
            expect(resolveCommand(eventFor(b)), b.label).toBe(b.command);
        }
    });

    it("resolves the named Excel shortcuts", () => {
        expect(resolveCommand({ key: "s", ctrlKey: true })).toBe("save");
        expect(resolveCommand({ key: "c", ctrlKey: true })).toBe("copy");
        expect(resolveCommand({ key: "v", ctrlKey: true })).toBe("paste");
        expect(resolveCommand({ key: "z", ctrlKey: true })).toBe("undo");
        expect(resolveCommand({ key: "ArrowUp", ctrlKey: true, shiftKey: true })).toBe("extendEdgeUp");
        expect(resolveCommand({ key: "ArrowDown", ctrlKey: true })).toBe("edgeDown");
        expect(resolveCommand({ key: "L", ctrlKey: true, shiftKey: true })).toBe("filterToggle");
        expect(resolveCommand({ key: "ArrowDown", altKey: true })).toBe("openFilter");
        expect(resolveCommand({ key: "Enter", shiftKey: true })).toBe("moveUp");
        expect(resolveCommand({ key: "Enter" })).toBe("moveDown");
        expect(resolveCommand({ key: "Tab" })).toBe("moveRight");
        expect(resolveCommand({ key: "Delete" })).toBe("clear");
        expect(resolveCommand({ key: "F2" })).toBe("edit");
        expect(resolveCommand({ key: "PageDown", ctrlKey: true })).toBe("nextSheet");
    });

    it("returns null for keys it does not own", () => {
        expect(resolveCommand({ key: "a" })).toBeNull();
        expect(resolveCommand({ key: "F5" })).toBeNull();
        expect(resolveCommand({ key: "q", ctrlKey: true })).toBeNull();
    });

    it("zooms on Ctrl + wheel", () => {
        expect(wheelZoom({ ctrlKey: true, deltaY: -100 })).toBe("zoomIn");
        expect(wheelZoom({ ctrlKey: true, deltaY: 100 })).toBe("zoomOut");
        expect(wheelZoom({ ctrlKey: false, deltaY: -100 })).toBeNull();
    });
});
