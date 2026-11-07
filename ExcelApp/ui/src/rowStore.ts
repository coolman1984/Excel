import type { CellData, WireRow } from "./types";

// rowStore (plan 2.6): the grid asks for cells synchronously, so we keep blocks
// of 200 view rows in a Map, an LRU of 400 blocks, and fetch with `rows` for the
// visible range plus two screens above and below, coalescing one request per block.

/** Rows per block. */
export const BLOCK = 200;
/** Maximum number of blocks kept in memory (plan 2.6). */
export const MAX_BLOCKS = 400;

export type CallFn = (cmd: string, args?: object) => Promise<any>;

/** Decode one wire cell `[text, kind, align, color?, warn?, style?]`. */
export function decodeCell(arr: any[]): CellData {
    const c: CellData = { text: String(arr[0] ?? ""), kind: (arr[1] ?? 0) as CellData["kind"], align: (arr[2] ?? 0) as CellData["align"] };
    if (arr[3] != null) c.color = String(arr[3]);
    if (arr[4]) c.warn = true;
    if (arr[5] != null) c.style = arr[5];
    return c;
}

/**
 * Keeps decoded rows for the current sheet. Not React state: glide is told about
 * new data through the onBlock callback, which the grid turns into updateCells.
 */
export class RowStore {
    private blocks = new Map<number, WireRow[]>();
    private inflight = new Map<number, Promise<void>>();
    private generation = 0;
    private sheet = 0;
    private viewRows = 0;
    /** Number of blocks requested since the last clear (test/inspection helper). */
    public requests = 0;

    constructor(private call: CallFn, private onBlock: (fromRow: number, count: number) => void) {}

    /** Point the store at a sheet and total view row count; clears the cache. */
    setSheet(sheet: number, viewRows: number): void {
        this.sheet = sheet;
        this.viewRows = viewRows;
        this.clear();
    }

    get viewRowCount(): number {
        return this.viewRows;
    }

    get blockCount(): number {
        return this.blocks.size;
    }

    /** Drop every block (after filter, sort, undo, paste, opened, changed). */
    clear(): void {
        this.generation++;
        this.blocks.clear();
        this.inflight.clear();
        this.requests = 0;
    }

    /** Discard affected blocks and reject responses from the previous edit generation. */
    invalidateRows(rows: Iterable<number>): void {
        this.generation++;
        this.inflight.clear();
        for (const row of rows) this.blocks.delete(Math.floor(row / BLOCK));
    }

    /** The decoded row for a view row, or undefined when not loaded. */
    getRow(viewRow: number): WireRow | undefined {
        const b = Math.floor(viewRow / BLOCK);
        const blk = this.blocks.get(b);
        if (blk === undefined) return undefined;
        this.touch(b);
        return blk[viewRow - b * BLOCK];
    }

    /** Request every block covering [fromView, toView] (inclusive). */
    ensure(fromView: number, toView: number): void {
        if (this.viewRows <= 0) return;
        const first = Math.max(0, Math.floor(fromView / BLOCK));
        const last = Math.floor(Math.max(fromView, toView) / BLOCK);
        for (let b = first; b <= last; b++) this.ensureBlock(b);
    }

    /** Prime a single block (used by the frozen-row drawing and the self test). */
    ensureBlock(b: number): void {
        if (b < 0 || this.viewRows <= 0) return;
        const from = b * BLOCK;
        if (from >= this.viewRows) return;
        if (this.blocks.has(b)) {
            this.touch(b);
            return;
        }
        if (this.inflight.has(b)) return; // coalesce: one request per block in flight
        const count = Math.min(BLOCK, this.viewRows - from);
        this.requests += 1;
        const generation = this.generation;
        const p = (async () => {
            const rows: WireRow[] = [];
            while (rows.length < count) {
                const res = await this.call("rows", { sheet: this.sheet, from: from + rows.length, count: count - rows.length });
                if (generation !== this.generation) return;
                const batch = res?.rows ?? [];
                if (batch.length === 0) throw new Error("The host returned no rows.");
                rows.push(...batch.map((r: any[]) => ({ rowNumber: r[0], cells: (r[1] ?? []).map(decodeCell) })));
            }
            this.blocks.set(b, rows);
            this.touch(b); this.evict(); this.onBlock(from, rows.length);
        })().catch(() => {
            // The next visible-region request may retry a failed block.
        }).finally(() => {
            if (this.inflight.get(b) === p) this.inflight.delete(b);
        });
        this.inflight.set(b, p);
    }

    private touch(b: number): void {
        // Re-insert to move the block to the newest position (Map preserves order).
        const v = this.blocks.get(b);
        if (v === undefined) return;
        this.blocks.delete(b);
        this.blocks.set(b, v);
    }

    private evict(): void {
        while (this.blocks.size > MAX_BLOCKS) {
            const oldest = this.blocks.keys().next().value as number | undefined;
            if (oldest === undefined) break;
            this.blocks.delete(oldest);
        }
    }
}
