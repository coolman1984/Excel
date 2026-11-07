/** Paint timing and the signal observed by the native jump probe. */
export class PaintSignal {
    lastPaintMs = 0;
    private completions = 0;
    get token(): number { return this.completions; }
    record(ms: number): void {
        this.lastPaintMs = ms;
        this.completions++;
    }
}
