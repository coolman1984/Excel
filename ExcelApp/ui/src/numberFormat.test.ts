import { describe, it, expect } from "vitest";
import { increaseDecimals, decreaseDecimals, GENERAL, NUMBER_FORMATS, PERCENT_CODE, COMMA_CODE } from "./numberFormat";

describe("number format buttons", () => {
    it("Increase Decimal: General becomes 0.0", () => {
        expect(increaseDecimals(GENERAL)).toBe("0.0");
    });

    it("Increase Decimal adds one place", () => {
        expect(increaseDecimals("0.00")).toBe("0.000");
        expect(increaseDecimals("0.00%")).toBe("0.000%");
        expect(increaseDecimals("0%")).toBe("0.0%");
        expect(increaseDecimals("#,##0")).toBe("#,##0.0");
        expect(increaseDecimals("#,##0.00")).toBe("#,##0.000");
    });

    it("Decrease Decimal removes one place", () => {
        expect(decreaseDecimals("0.000")).toBe("0.00");
        expect(decreaseDecimals("0.0%")).toBe("0%");
        expect(decreaseDecimals("0.00")).toBe("0.0");
        expect(decreaseDecimals("#,##0.00")).toBe("#,##0.0");
    });

    it("Decrease Decimal: General becomes 0 and an integer stays put", () => {
        expect(decreaseDecimals(GENERAL)).toBe("0");
        expect(decreaseDecimals("0")).toBe("0");
    });

    it("exposes the chosen codes", () => {
        const codes = NUMBER_FORMATS.map((f) => f.code);
        expect(codes).toContain("#,##0.00");
        expect(codes).toContain("[$$-409]#,##0.00");
        expect(codes).toContain("yyyy-mm-dd");
        expect(codes).toContain("0.00%");
        expect(codes).toContain("@");
        expect(PERCENT_CODE).toBe("0%");
        expect(COMMA_CODE).toBe("#,##0");
    });
});
