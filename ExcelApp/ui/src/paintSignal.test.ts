import { describe, expect, it } from "vitest";
import { PaintSignal } from "./paintSignal";

describe("jump paint completion signal", () => {
    it("distinguishes completed jumps with identical elapsed durations", () => {
        const signal = new PaintSignal();
        signal.record(47.89999999944121);
        const before = signal.token;
        signal.record(47.89999999944121);
        expect(signal.token).not.toBe(before);
        expect(signal.lastPaintMs).toBe(47.89999999944121);
    });
    it("recognizes a cached zero-duration completion before the next probe poll", () => {
        const signal = new PaintSignal();
        const before = signal.token;
        signal.record(0);
        expect(signal.token).not.toBe(before);
        const observed = signal.token;
        expect(signal.token).toBe(observed);
    });
});
