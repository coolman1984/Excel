import { describe, it, expect } from "vitest";
import { addEvent, addPrompt, ends } from "./transcript";

describe("agent transcript", () => {
    it("keeps the conversation in order", () => {
        let l = addPrompt([], "sum by region");
        l = addEvent(l, { kind: "status", text: "Claude Code is working..." });
        l = addEvent(l, { kind: "tool", text: "query: SELECT ..." });
        l = addEvent(l, { kind: "text", text: "North: 40" });
        l = addEvent(l, { kind: "done", text: "Done in 2 s." });
        expect(l.map((e) => e.kind)).toEqual(["prompt", "status", "tool", "text", "done"]);
        expect(l[0].role).toBe("user");
    });
    it("a new status replaces the previous one", () => {
        let l = addEvent([], { kind: "status", text: "Reconnecting 1/5" });
        l = addEvent(l, { kind: "status", text: "Reconnecting 2/5" });
        expect(l).toHaveLength(1);
        expect(l[0].text).toBe("Reconnecting 2/5");
    });
    it("drops session and empty events", () => {
        expect(addEvent([], { kind: "session", text: "abc" })).toHaveLength(0);
        expect(addEvent([], { kind: "text", text: "" })).toHaveLength(0);
    });
    it("done and error end the turn", () => {
        expect(ends({ kind: "done" }) && ends({ kind: "error" }) && !ends({ kind: "tool" })).toBe(true);
    });
});
