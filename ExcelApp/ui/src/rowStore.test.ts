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
