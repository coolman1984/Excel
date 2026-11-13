import { describe, it, expect } from "vitest";
import { whenText } from "./when";

describe("whenText", () => {
    it("shows an ISO time in local time to the minute", () => {
        const d = new Date(2026, 10, 8, 9, 30, 12, 345);
        expect(whenText(d.toISOString())).toBe("2026-11-08 09:30");
    });
    it("keeps a plain date and says Never for nothing", () => {
        expect(whenText("2026-11-07")).toBe("2026-11-07");
        expect(whenText(undefined)).toBe("Never");
        expect(whenText(null, "-")).toBe("-");
    });
    it("leaves text it cannot read as it is", () => {
        expect(whenText("yesterday")).toBe("yesterday");
    });
});
