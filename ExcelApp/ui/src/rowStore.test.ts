import { describe, it, expect } from "vitest";
import { RowStore, BLOCK, MAX_BLOCKS } from "./rowStore";

// A fake `rows` command that answers any block immediately.
function fakeCall(log: any[]) {
    return (cmd: string, args: any) => {
        log.push({ cmd, from: args.from, count: args.count });
        const rows: any[] = [];
        for (let i = 0; i < args.count; i++) rows.push([args.from + i + 1, [["x", 2, 0]]]);
        return Promise.resolve({ from: args.from, rows });
    };
}

const flush = () => new Promise((r) => setTimeout(r, 0));

describe("RowStore", () => {
    it("loads 200-row blocks and reads them back", async () => {
        const log: any[] = [];
        const store = new RowStore(fakeCall(log), () => {});
        store.setSheet(0, 1000);
        store.ensure(0, 0);
        await flush();
        expect(log.length).toBe(1);
        expect(log[0].count).toBe(BLOCK);
        expect(store.getRow(0)).toBeTruthy();
        expect(store.getRow(0)!.rowNumber).toBe(1);
        expect(store.getRow(199)!.rowNumber).toBe(200);
        expect(store.blockCount).toBe(1);
    });

    it("coalesces requests: one in flight per block", async () => {
        const log: any[] = [];
        const store = new RowStore(fakeCall(log), () => {});
        store.setSheet(0, 1000);
        store.ensure(0, 50);
        store.ensure(0, 50); // second call before the first resolves
        store.ensure(10, 40);
        await flush();
        expect(log.length).toBe(1);
    });

    it("keeps at most 400 blocks (LRU) and evicts the oldest", async () => {
        expect(MAX_BLOCKS).toBe(400);
        const log: any[] = [];
        const store = new RowStore(fakeCall(log), () => {});
        store.setSheet(0, 100000);
        for (let b = 0; b <= MAX_BLOCKS; b++) store.ensureBlock(b); // one more than the cap
        await flush();
        await flush();
        expect(store.blockCount).toBe(MAX_BLOCKS);
        // Block 0 was the oldest and must be gone; the newest block survives.
        expect(store.getRow(0)).toBeUndefined();
        expect(store.getRow(MAX_BLOCKS * BLOCK)).toBeTruthy();
    });

    it("clears on setSheet and reports no block for a missing row", async () => {
        const log: any[] = [];
        const store = new RowStore(fakeCall(log), () => {});
        store.setSheet(0, 1000);
        store.ensure(0, 0);
        await flush();
        expect(store.blockCount).toBe(1);
        store.setSheet(1, 500);
        expect(store.blockCount).toBe(0);
        expect(store.getRow(0)).toBeUndefined();
    });
});


describe("view invalidation", () => {
    it("ignores an old response after switching sheets", async () => {
        const replies: ((value: any) => void)[] = [];
        const store = new RowStore(() => new Promise(resolve => replies.push(resolve)), () => {});
        store.setSheet(0, 1); store.ensure(0, 0);
        store.setSheet(1, 1); store.ensure(0, 0);
        replies[1]({ rows: [[42, [["new", 2, 0]]]] }); await flush();
        replies[0]({ rows: [[1, [["old", 2, 0]]]] }); await flush();
        expect(store.getRow(0)?.rowNumber).toBe(42);
    });
    it("reloads changed values and styles instead of repainting stale data", async () => {
        let text = "before";
        const store = new RowStore(async () => ({ rows: [[1, [[text, 1, 1, null, null, { bold: text === "after" }]]]] }), () => {});
        store.setSheet(0, 1); store.ensure(0, 0); await flush();
        text = "after"; store.invalidateRows([0]); store.ensure(0, 0); await flush();
        expect(store.getRow(0)?.cells[0].text).toBe("after");
        expect(store.getRow(0)?.cells[0].style?.bold).toBe(true);
    });
});


describe("wide sheets", () => {
    it("continues a block when the host caps each reply", async () => {
        const store = new RowStore(async (_cmd, a: any) => ({ rows: Array.from({ length: Math.min(3, a.count) }, (_, i) => [a.from + i + 1, [["x", 2, 0]]]) }), () => {});
        store.setSheet(0, 10); store.ensure(0, 0); await flush();
        expect(store.getRow(9)?.rowNumber).toBe(10);
    });
});
