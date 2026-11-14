import { mockDataCall, mockDataOn } from "./data/mockData";
import type { CellData, CellStyle, Condition, SortKey, Stats, WorkbookInfo } from "./types";

// A made-up host so the UI can be built and driven in a plain browser (Chrome)
// without the C# WebView2 host. It answers every command of plan 2.3 plus `edge`
// and `locate` on a 1,000 x 26 sheet. No real workbook is ever involved.

interface MockSheet {
    name: string;
    hidden: boolean;
    rows: number;
    cols: number;
    base: CellData[][];
}

const NAMES = [
    "Orion", "Vega", "Lyra", "Draco", "Hydra", "Corvus", "Pavo", "Tucana", "Indus", "Volans",
    "Carina", "Dorado", "Fornax", "Lepus", "Mensa", "Norma", "Octans", "Reticulum", "Sculptor", "Vela",
];
const GROUPS = ["North", "South", "East", "West", "Central", "Coastal", "Inland", "Highland"];

function fmtMoney(v: number): string {
    return Math.abs(v).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function makeCell(r: number, c: number): CellData {
    const seed = (r * 31 + c * 17) >>> 0;
    switch (c) {
        case 0:
            return { text: String(r + 1), kind: 1, align: 1 };
        case 1:
            return { text: NAMES[(r * 7 + 3) % NAMES.length] + " " + (100 + (r % 90)), kind: 2, align: 0 };
        case 2: {
            const v = Math.round((((seed % 900) + 100) * 100) / 100);
            return { text: (v + r / 100).toFixed(2), kind: 1, align: 1 };
        }
        case 3: {
            const day = (r % 27) + 1;
            const month = (r % 12) + 1;
            return { text: `2026-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`, kind: 1, align: 1 };
        }
        case 4:
            return { text: r % 2 === 0 ? "TRUE" : "FALSE", kind: 3, align: 2 };
        case 5: {
            const v = ((seed % 9000) + 500) / 100;
            const neg = r % 5 === 0;
            return {
                text: (neg ? "-" : "") + "$" + fmtMoney(v),
                kind: 1,
                align: 1,
                color: neg ? "#C00000" : undefined,
            };
        }
        case 6:
            return r % 9 === 0 ? { text: "", kind: 0, align: 0 } : { text: GROUPS[r % GROUPS.length], kind: 2, align: 0 };
        case 7: {
            const v = ((r % 40) + 5) / 10;
            return { text: v.toFixed(1) + "%", kind: 1, align: 1 };
        }
        case 8: {
            const v = ((seed % 2000) - 1000) / 10;
            return { text: v.toFixed(1), kind: 1, align: 1 };
        }
        case 9:
            return r % 37 === 0 ? { text: "#N/A", kind: 4, align: 2 } : { text: "OK", kind: 2, align: 0 };
        case 10:
            return { text: (r % 3 === 0 ? "" : "PO-" + (1000 + (r % 500))), kind: r % 3 === 0 ? 0 : 2, align: 0 };
        default: {
            const letter = String.fromCharCode(65 + (c % 26));
            return { text: c % 4 === 0 ? String(r * c) : `${letter}${c}-${r}`, kind: c % 4 === 0 ? 1 : 2, align: c % 4 === 0 ? 1 : 0 };
        }
    }
}

function buildSheet(name: string, rows: number, cols: number, hidden = false): MockSheet {
    const base: CellData[][] = [];
    for (let r = 0; r < rows; r++) {
        const row: CellData[] = [];
        for (let c = 0; c < cols; c++) row.push(makeCell(r, c));
        base.push(row);
    }
    return { name, hidden, rows, cols, base };
}

interface EditRecord {
    kind: "value" | "style";
    key: string; // "r,c"
    prev: CellData | CellStyle | undefined;
    next: CellData | CellStyle | undefined;
}

export class MockHost {
    private sheets: MockSheet[] = [
        buildSheet("Sheet1", 1000, 26),
        buildSheet("Sheet2", 500, 12),
        buildSheet("Sheet3", 300, 8),
        buildSheet("Hidden1", 200, 6, true),
    ];
    private current = 0;
    private valueEdits = new Map<string, CellData>();
    private styleEdits = new Map<string, CellStyle>();
    private colWidths = new Map<string, number>();
    private freeze = new Map<number, { rows: number; cols: number }>();
    private filters = new Map<string, Condition>(); // key "sheet,col"
    private sortKeys: SortKey[] = [];
    private headerRow = true;
    private clipboard = "";
    private pendingCut: number[][] | null = null;
    private undoStack: EditRecord[][] = [];
    private redoStack: EditRecord[][] = [];
    private dirty = false;

    private listeners = new Map<string, Set<(e: any) => void>>();
    private viewCache: { sheet: number; map: number[] | null; active: boolean } | null = null;

    on(event: string, handler: (e: any) => void): () => void {
        if (event.startsWith("data.")) return mockDataOn(event, handler);
        let set = this.listeners.get(event);
        if (!set) {
            set = new Set();
            this.listeners.set(event, set);
        }
        set.add(handler);
        return () => set!.delete(handler);
    }

    // A pretend Claude Code for the browser preview: it reads two cells, writes one and answers.
    private agentTimers: ReturnType<typeof setTimeout>[] = [];
    private agentStart(args: any): any {
        const steps: [number, any][] = [
            [200, { kind: "status", text: "Claude Code is working..." }],
            [700, { kind: "tool", text: "app_status" }],
            [1200, { kind: "tool", text: "read_range: C2:C11" }],
            [1700, { kind: "tool", text: "write_range: C12" }],
            [2000, { kind: "text", text: `Done. I added the total of column C under the data in C12 (=SUM(C2:C11)). The change is not saved yet.\n\nYou asked: ${args?.prompt ?? ""}` }],
            [2100, { kind: "done", text: "Done in 1.9 s." }],
        ];
        this.agentTimers.forEach(clearTimeout);
        this.agentTimers = steps.map(([ms, e]) => setTimeout(() => {
            if (e.kind === "done") this.agentTimers = [];
            if (e.text?.startsWith("Done. I added")) { this.valueEdits.set(this.keyFor(11, 2), { text: "=SUM(C2:C11)", kind: 2, align: 0 }); this.emit("agentChanged", { cmd: "agent.write" }); }
            this.emit("agent.event", e);
        }, ms));
        return { started: true, agent: args?.agent ?? "claude", resumed: !args?.newChat };
    }

    private emit(event: string, payload: any): void {
        const set = this.listeners.get(event);
        if (set) for (const h of set) h({ event, ...payload });
    }

    private keyFor(r: number, c: number): string {
        return r + "," + c;
    }

    private get headerRows(): number {
        return this.headerRow ? 1 : 0;
    }

    private cellAt(r: number, c: number): CellData {
        const s = this.sheets[this.current];
        if (r < 0 || r >= s.rows || c < 0 || c >= s.cols) return { text: "", kind: 0, align: 0 };
        const edited = this.valueEdits.get(this.keyFor(r, c));
        const base = edited ?? s.base[r][c];
        const style = this.styleEdits.get(this.keyFor(r, c));
        return style ? { ...base, style: { ...base.style, ...style } } : base;
    }

    private activeFiltersFor(): { col: number; cond: Condition }[] {
        const out: { col: number; cond: Condition }[] = [];
        for (const [k, cond] of this.filters) {
            const [sheet, col] = k.split(",").map(Number);
            if (sheet === this.current) out.push({ col, cond });
        }
        return out;
    }

    private view(): { map: number[] | null; active: boolean } {
        if (this.viewCache && this.viewCache.sheet === this.current) return this.viewCache;
        const s = this.sheets[this.current];
        const active = this.activeFiltersFor().length > 0 || this.sortKeys.length > 0;
        let map: number[] | null = null;
        if (active) {
            const head: number[] = [];
            for (let r = 0; r < this.headerRows && r < s.rows; r++) head.push(r);
            const data: number[] = [];
            for (let r = this.headerRows; r < s.rows; r++) if (this.rowPasses(r)) data.push(r);
            if (this.sortKeys.length > 0) this.sortRows(data);
            map = head.concat(data);
        }
        this.viewCache = { sheet: this.current, map, active };
        return this.viewCache;
    }

    private invalidateView(): void {
        this.viewCache = null;
    }

    private currentViewRows(): number {
        const v = this.view();
        return v.map ? v.map.length : this.sheets[this.current].rows;
    }

    private rawNumber(r: number, c: number): number | null {
        const cell = this.cellAt(r, c);
        if (cell.kind !== 1) return null;
        const n = Number(cell.text.replace(/[$,%\s]/g, "").replace(/,/g, ""));
        return Number.isFinite(n) ? n : null;
    }

    private onePasses(r: number, col: number, cond: Condition): boolean {
        const cell = this.cellAt(r, col);
        const txt = cell.text;
        if (cond.kind === "values") {
            const blank = cell.kind === 0 || txt === "";
            if (blank) return !!cond.blanks;
            return (cond.include ?? []).includes(txt);
        }
        if (cond.kind === "number") {
            const n = this.rawNumber(r, col);
            if (n === null) return false;
            const a = cond.a ?? 0;
            switch (cond.op) {
                case "eq":
                    return n === a;
                case "ne":
                    return n !== a;
                case "gt":
                    return n > a;
                case "ge":
                    return n >= a;
                case "lt":
                    return n < a;
                case "le":
                    return n <= a;
                case "between":
                    return n >= a && n <= (cond.b ?? a);
                default:
                    return true; // top10 / above/below average not simulated precisely
            }
        }
        const v = cond.value ?? "";
        switch (cond.op) {
            case "eq":
                return txt === v;
            case "ne":
                return txt !== v;
            case "contains":
                return txt.indexOf(v) >= 0;
            case "notContains":
                return txt.indexOf(v) < 0;
            case "begins":
                return txt.startsWith(v);
            case "ends":
                return txt.endsWith(v);
            default:
                return true;
        }
    }

    private rowPasses(r: number): boolean {
        for (const { col, cond } of this.activeFiltersFor()) if (!this.onePasses(r, col, cond)) return false;
        return true;
    }

    private sortRows(rows: number[]): void {
        const keys = this.sortKeys;
        rows.sort((a, b) => {
            for (const k of keys) {
                const d = this.compareCells(a, b, k.col);
                if (d !== 0) return k.desc ? -d : d;
            }
            return 0;
        });
    }

    private compareCells(a: number, b: number, col: number): number {
        const ca = this.cellAt(a, col);
        const cb = this.cellAt(b, col);
        const rank = (x: CellData): number => (x.kind === 1 ? 0 : x.kind === 2 ? 1 : x.kind === 3 ? 2 : x.kind === 4 ? 3 : 4);
        const ra = rank(ca);
        const rb = rank(cb);
        if (ra !== rb) return ra - rb;
        if (ca.text === "" && cb.text === "") return 0;
        if (ca.text === "") return 1;
        if (cb.text === "") return -1;
        if (ca.kind === 1 && cb.kind === 1) return (this.rawNumber(a, col) ?? 0) - (this.rawNumber(b, col) ?? 0);
        return ca.text.localeCompare(cb.text, undefined, { sensitivity: "accent" });
    }

    private visibleSheets(): WorkbookInfo["sheets"] {
        return this.sheets.map((s, i) => ({
            index: i,
            name: s.name,
            hidden: s.hidden,
            rows: s.rows,
            cols: s.cols,
            firstRow: 1,
            firstCol: 1,
            colWidthsPx: Array.from({ length: s.cols }, (_v, c) => this.colWidths.get(i + "," + c) ?? 64),
            frozenRows: this.freeze.get(i)?.rows ?? 0,
            frozenCols: this.freeze.get(i)?.cols ?? 0,
        }));
    }

    private workbook(): WorkbookInfo {
        return { name: "syn-mock.xlsx", path: "mock://syn-mock.xlsx", sheets: this.visibleSheets(), dirty: this.dirty };
    }

    private letter(col0: number): string {
        let n = col0 + 1;
        let s = "";
        while (n > 0) {
            const rem = (n - 1) % 26;
            s = String.fromCharCode(65 + rem) + s;
            n = Math.floor((n - 1) / 26);
        }
        return s;
    }

    private applyEdits(records: EditRecord[], dir: "undo" | "redo"): void {
        for (const rec of records) {
            const value = dir === "undo" ? rec.prev : rec.next;
            if (rec.kind === "value") {
                if (value === undefined) this.valueEdits.delete(rec.key);
                else this.valueEdits.set(rec.key, value as CellData);
            } else {
                if (value === undefined) this.styleEdits.delete(rec.key);
                else this.styleEdits.set(rec.key, value as CellStyle);
            }
        }
    }

    private record(recs: EditRecord[]): void {
        if (recs.length > 0) {
            this.undoStack.push(recs);
            this.redoStack = [];
            this.dirty = true;
        }
    }

    async call(cmd: string, args: any = {}): Promise<any> {
        if (cmd.startsWith("data.")) return mockDataCall(cmd, args);
        switch (cmd) {
            case "ready":
                return { version: "mock-1", recent: [], approved: true };
            case "pickFile":
                return { path: null };
            case "workbook":
                return this.workbook();
            case "recent":
                return { recent: [] };
            case "open":
                await this.open(args.path);
                return { started: true };
            case "sheet": {
                this.current = args.sheet ?? 0;
                this.invalidateView();
                const v = this.view();
                return {
                    viewRows: v.map ? v.map.length : this.sheets[this.current].rows,
                    filterActive: this.activeFiltersFor().length > 0,
                    sort: this.sortKeys,
                    filters: Object.fromEntries(this.activeFiltersFor().map((f) => [f.col, f.cond])),
                };
            }
            case "rows":
                return this.rows(args);
            case "cellDetail":
                return this.cellDetail(args);
            case "setCell":
                return this.setCell(args);
            case "clearCells":
                return this.clearCells(args);
            case "setStyle":
                return this.setStyle(args);
            case "recoverEdits":
                return { restored: 3, dirty: true, canUndo: true };
            case "discardRecovery":
                return {};
            case "undo":
                return this.undo();
            case "redo":
                return this.redo();
            case "copy":
                return this.copy(args);
            case "cut":
                return this.cut(args);
            case "paste":
                return this.paste(args);
            case "fill":
                return this.fill(args);
            case "filterValues":
                return this.filterValues(args);
            case "setFilter":
                return this.setFilter(args);
            case "clearFilters":
                this.filters.clear();
                this.invalidateView();
                return { viewRows: this.currentViewRows(), filterActive: false };
            case "setSort":
                this.sortKeys = args.keys ?? [];
                this.invalidateView();
                return { viewRows: this.currentViewRows() };
            case "setHeaderRow":
                this.headerRow = !!args.hasHeader;
                this.invalidateView();
                return { viewRows: this.currentViewRows() };
            case "find":
                this.find(args);
                return { started: true };
            case "replace":
                return this.replace(args);
            case "stats":
                return this.stats(args);
            case "setColWidth":
                this.colWidths.set(args.sheet + "," + args.col, args.px);
                return { ok: true };
            case "setFreeze":
                this.freeze.set(args.sheet ?? this.current, { rows: args.rows ?? 0, cols: args.cols ?? 0 });
                return { ok: true };
            case "save":
                return this.save(args);
            case "close":
            case "openExternal":
            case "log":
                return { ok: true };
            case "agent.status":
                return { claude: true, codex: false, running: this.agentTimers.length > 0 ? "claude" : null };
            case "agent.start":
                return this.agentStart(args);
            case "agent.stop":
                this.agentTimers.forEach(clearTimeout); this.agentTimers = [];
                this.emit("agent.event", { kind: "done", text: "Stopped." });
                return { stopped: true };
            case "edge":
                return this.edge(args);
            case "locate":
                return this.locate(args);
            default:
                return { ok: true };
        }
    }

    private async open(path: string | null): Promise<void> {
        this.emit("opening", { phase: "Opening", percent: 0, seconds: 0 });
        await this.delay(60);
        this.emit("opening", { phase: "Reading file", percent: 45, seconds: 0.2 });
        await this.delay(80);
        if (!path) path = "mock://syn-mock.xlsx";
        this.emit("opening", { phase: "Reading file", percent: 90, seconds: 0.4 });
        await this.delay(60);
        this.current = 0;
        this.invalidateView();
        this.emit("opened", { workbook: this.workbook() });
        if (new URLSearchParams(globalThis.location?.search ?? "").get("recovery") === "1")
            this.emit("recovery", { available: true, cells: 3, savedAt: "2026-11-08 09:15", sameFile: true });
    }

    private delay(ms: number): Promise<void> {
        return new Promise((res) => setTimeout(res, ms));
    }

    private rows(args: any): any {
        const v = this.view();
        const total = v.map ? v.map.length : this.sheets[this.current].rows;
        const from = Math.max(0, args.from | 0);
        const count = Math.min(500, Math.max(0, args.count | 0));
        const out: any[] = [];
        for (let i = 0; i < count && from + i < total; i++) {
            const real = v.map ? v.map[from + i] : from + i;
            const row = this.sheets[this.current].base[real];
            out.push([real + 1, row.map((_c, c) => this.cellArray(real, c))]);
        }
        return { from, rows: out };
    }

    private cellArray(r: number, c: number): any[] {
        const cell = this.cellAt(r, c);
        const arr: any[] = [cell.text, cell.kind, cell.align];
        if (cell.color) arr[3] = cell.color;
        if (cell.warn) arr[4] = true;
        if (cell.style) arr[5] = cell.style;
        return arr;
    }

    private cellDetail(args: any): any {
        const v = this.view();
        const total = v.map ? v.map.length : this.sheets[this.current].rows;
        if (args.row < 0 || args.row >= total) return { address: "", raw: "", edited: false };
        const real = v.map ? v.map[args.row] : args.row;
        const cell = this.cellAt(real, args.col);
        const edited = this.valueEdits.has(this.keyFor(real, args.col)) || this.styleEdits.has(this.keyFor(real, args.col));
        return { address: this.letter(args.col) + (real + 1), raw: cell.text, edited };
    }

    private parseInput(input: string): CellData {
        const s = (input ?? "").trim();
        if (s === "") return { text: "", kind: 0, align: 0 };
        if (s.startsWith("=")) return { text: s, kind: 2, align: 0, style: { italic: true }, color: "#555555" };
        if (s.toUpperCase() === "TRUE") return { text: "TRUE", kind: 3, align: 2 };
        if (s.toUpperCase() === "FALSE") return { text: "FALSE", kind: 3, align: 2 };
        if (s.endsWith("%")) {
            const n = Number(s.slice(0, -1));
            if (Number.isFinite(n)) return { text: String(n / 100), kind: 1, align: 1 };
        }
        const n = Number(s.replace(/,/g, ""));
        if (Number.isFinite(n) && /^-?[\d.,]+$/.test(s)) return { text: String(n), kind: 1, align: 1 };
        return { text: s, kind: 2, align: 0 };
    }

    private setCell(args: any): any {
        const v = this.view();
        const real = v.map ? v.map[args.row] : args.row;
        const key = this.keyFor(real, args.col);
        const prev = this.valueEdits.get(key);
        const next = this.parseInput(args.input);
        this.valueEdits.set(key, next);
        this.record([{ kind: "value", key, prev, next }]);
        return { changed: [[args.row, args.col]], dirty: true };
    }

    private clearCells(args: any): any {
        const recs: EditRecord[] = [];
        const changed: number[][] = [];
        const v = this.view();
        for (const [r1, c1, r2, c2] of args.ranges ?? []) {
            for (let r = r1; r <= r2; r++) {
                const real = v.map ? v.map[r] : r;
                if (real === undefined) continue;
                for (let c = c1; c <= c2; c++) {
                    const key = this.keyFor(real, c);
                    const next = { text: "", kind: 0 as const, align: 0 as const };
                    recs.push({ kind: "value", key, prev: this.valueEdits.get(key), next });
                    this.valueEdits.set(key, next);
                    changed.push([r, c]);
                }
            }
        }
        this.record(recs);
        return { changed, dirty: true };
    }

    private setStyle(args: any): any {
        const recs: EditRecord[] = [];
        const changed: number[][] = [];
        const v = this.view();
        for (const [r1, c1, r2, c2] of args.ranges ?? []) {
            for (let r = r1; r <= r2; r++) {
                const real = v.map ? v.map[r] : r;
                if (real === undefined) continue;
                for (let c = c1; c <= c2; c++) {
                    const key = this.keyFor(real, c);
                    const prev = this.styleEdits.get(key);
                    const next = { ...prev, ...args.style };
                    recs.push({ kind: "style", key, prev, next });
                    this.styleEdits.set(key, next);
                    changed.push([r, c]);
                }
            }
        }
        this.record(recs);
        return { changed, dirty: true };
    }

    private undo(): any {
        const recs = this.undoStack.pop();
        if (!recs) return { changed: [], dirty: this.dirty, canUndo: false, canRedo: this.redoStack.length > 0 };
        this.applyEdits(recs, "undo");
        this.redoStack.push(recs);
        return {
            changed: recs.map((rec) => rec.key.split(",").map(Number)),
            dirty: this.undoStack.length > 0,
            canUndo: this.undoStack.length > 0,
            canRedo: true,
        };
    }

    private redo(): any {
        const recs = this.redoStack.pop();
        if (!recs) return { changed: [], dirty: this.dirty, canUndo: this.undoStack.length > 0, canRedo: false };
        this.applyEdits(recs, "redo");
        this.undoStack.push(recs);
        return { changed: recs.map((rec) => rec.key.split(",").map(Number)), dirty: true, canUndo: true, canRedo: this.redoStack.length > 0 };
    }

    private buildText(ranges: number[][]): { text: string; cells: number } {
        if (!ranges || ranges.length === 0) return { text: "", cells: 0 };
        const [r1, c1, r2, c2] = ranges[0];
        const v = this.view();
        const lines: string[] = [];
        let cells = 0;
        for (let r = r1; r <= r2; r++) {
            const real = v.map ? v.map[r] : r;
            const cols: string[] = [];
            for (let c = c1; c <= c2; c++) {
                cols.push(this.cellAt(real, c).text);
                cells++;
            }
            lines.push(cols.join("\t"));
        }
        return { text: lines.join("\n"), cells };
    }

    private copy(args: any): any {
        const { text, cells } = this.buildText(args.ranges);
        this.clipboard = text;
        return { cells };
    }

    private cut(args: any): any {
        const res = this.copy(args);
        this.pendingCut = args.ranges;
        return res;
    }

    private paste(args: any): any {
        const v = this.view();
        const lines = this.clipboard.split(/\r?\n/);
        const recs: EditRecord[] = [];
        const changed: number[][] = [];
        for (let i = 0; i < lines.length; i++) {
            const real = v.map ? v.map[args.row + i] : args.row + i;
            if (real === undefined) break;
            const cols = lines[i].split("\t");
            for (let j = 0; j < cols.length; j++) {
                const c = args.col + j;
                const key = this.keyFor(real, c);
                const next = this.parseInput(cols[j]);
                recs.push({ kind: "value", key, prev: this.valueEdits.get(key), next });
                this.valueEdits.set(key, next);
                changed.push([args.row + i, c]);
            }
        }
        if (this.pendingCut) {
            this.clearCells({ ranges: this.pendingCut });
            this.pendingCut = null;
        }
        this.record(recs);
        return { changed, dirty: true };
    }

    private fill(args: any): any {
        const [sr1, sc1, sr2, sc2] = args.from;
        const [tr1, tc1, tr2, tc2] = args.to;
        const v = this.view();
        const srcH = sr2 - sr1 + 1;
        const srcW = sc2 - sc1 + 1;
        const recs: EditRecord[] = [];
        const changed: number[][] = [];
        for (let r = tr1; r <= tr2; r++) {
            for (let c = tc1; c <= tc2; c++) {
                const sr = sr1 + ((r - tr1) % srcH);
                const sc = sc1 + ((c - tc1) % srcW);
                const sReal = v.map ? v.map[sr] : sr;
                const tReal = v.map ? v.map[r] : r;
                if (sReal === undefined || tReal === undefined) continue;
                const next = this.cellAt(sReal, sc);
                const key = this.keyFor(tReal, c);
                recs.push({ kind: "value", key, prev: this.valueEdits.get(key), next });
                this.valueEdits.set(key, next);
                changed.push([r, c]);
            }
        }
        this.record(recs);
        return { changed, dirty: true };
    }

    private filterValues(args: any): any {
        const counts = new Map<string, number>();
        let blanks = 0;
        const s = this.sheets[this.current];
        const search = args.search ? String(args.search).toLowerCase() : null;
        const otherConds = this.activeFiltersFor().filter((f) => f.col !== args.col);
        for (let r = this.headerRows; r < s.rows; r++) {
            let ok = true;
            for (const { col, cond } of otherConds) {
                if (!this.onePasses(r, col, cond)) {
                    ok = false;
                    break;
                }
            }
            if (!ok) continue;
            const cell = this.cellAt(r, args.col);
            if (cell.kind === 0 || cell.text === "") {
                blanks++;
                continue;
            }
            if (search && cell.text.toLowerCase().indexOf(search) < 0) continue;
            counts.set(cell.text, (counts.get(cell.text) ?? 0) + 1);
        }
        const all = Array.from(counts.entries()).map(([text, count]) => ({ text, count }));
        all.sort((a, b) => {
            const aNum = /^[\d.,$-]+$/.test(a.text);
            const bNum = /^[\d.,$-]+$/.test(b.text);
            if (aNum && bNum) return Number(a.text.replace(/[$,]/g, "")) - Number(b.text.replace(/[$,]/g, ""));
            if (aNum) return -1;
            if (bNum) return 1;
            return a.text.localeCompare(b.text, undefined, { sensitivity: "accent" });
        });
        return { values: all.slice(0, 10000), blanks, truncated: all.length > 10000 };
    }

    private setFilter(args: any): any {
        const key = this.current + "," + args.col;
        if (args.condition == null) this.filters.delete(key);
        else this.filters.set(key, args.condition);
        this.invalidateView();
        return { viewRows: this.currentViewRows(), filterActive: this.activeFiltersFor().length > 0 };
    }

    private find(args: any): void {
        const sheets = args.sheet >= 0 ? [args.sheet] : this.sheets.map((_s, i) => i);
        const q = args.matchCase ? String(args.query) : String(args.query).toLowerCase();
        const hits: any[][] = [];
        let total = 0;
        for (const si of sheets) {
            const s = this.sheets[si];
            for (let r = 0; r < s.rows; r++) {
                for (let c = 0; c < s.cols; c++) {
                    const t = s.base[r][c].text;
                    const hay = args.matchCase ? t : t.toLowerCase();
                    if (hay.indexOf(q) >= 0) {
                        total++;
                        hits.push([si, r, c, t]);
                        if (hits.length >= 200) this.emit("findHits", { hits: hits.splice(0, hits.length) });
                    }
                }
                this.emit("findProgress", { percent: Math.round((100 * (r + 1)) / s.rows) });
            }
        }
        if (hits.length) this.emit("findHits", { hits: hits.splice(0, hits.length) });
        this.emit("findDone", { total, capped: false });
    }

    private replace(args: any): any {
        const recs: EditRecord[] = [];
        let count = 0;
        const s = this.sheets[this.current];
        for (let r = 0; r < s.rows && !(!args.all && count >= 1); r++) {
            for (let c = 0; c < s.cols; c++) {
                const cell = this.cellAt(r, c);
                if (cell.kind !== 2) continue;
                const hay = args.matchCase ? cell.text : cell.text.toLowerCase();
                const q = args.matchCase ? args.query : String(args.query).toLowerCase();
                if (hay.indexOf(q) < 0) continue;
                const nextText = args.matchCase
                    ? cell.text.split(args.query).join(args.replacement)
                    : this.replaceInsensitive(cell.text, args.query, args.replacement);
                const key = this.keyFor(r, c);
                recs.push({ kind: "value", key, prev: this.valueEdits.get(key), next: { ...cell, text: nextText } });
                this.valueEdits.set(key, { ...cell, text: nextText });
                count++;
                if (!args.all) break;
            }
        }
        this.record(recs);
        return { changed: recs.map((rec) => rec.key.split(",").map(Number)), count };
    }

    private replaceInsensitive(text: string, q: string, replacement: string): string {
        const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
        return text.replace(re, replacement);
    }

    private stats(args: any): Stats {
        const v = this.view();
        let count = 0;
        let numbers = 0;
        let sum = 0;
        let min = Number.POSITIVE_INFINITY;
        let max = Number.NEGATIVE_INFINITY;
        for (const [r1, c1, r2, c2] of args.ranges ?? []) {
            for (let r = r1; r <= r2; r++) {
                const real = v.map ? v.map[r] : r;
                if (real === undefined) continue;
                for (let c = c1; c <= c2; c++) {
                    const cell = this.cellAt(real, c);
                    if (cell.kind === 0) continue;
                    count++;
                    const n = this.rawNumber(real, c);
                    if (n !== null) {
                        numbers++;
                        sum += n;
                        min = Math.min(min, n);
                        max = Math.max(max, n);
                    }
                }
            }
        }
        return { count, numbers, sum, average: numbers > 0 ? sum / numbers : 0, min: numbers > 0 ? min : 0, max: numbers > 0 ? max : 0 };
    }

    private async save(_args: any): Promise<any> {
        this.emit("saving", { phase: "Writing changes", percent: 40 });
        await this.delay(80);
        this.emit("saving", { phase: "Reading the saved file back", percent: 72 });
        await this.delay(80);
        this.dirty = false;
        this.emit("saved", { path: "mock://syn-mock.xlsx", seconds: 0.2 });
        this.emit("opened", { workbook: this.workbook() });
        return { started: true };
    }

    private edge(args: any): { row: number; col: number } {
        const s = this.sheets[this.current];
        let r = args.row;
        let c = args.col;
        const dir = args.direction as "up" | "down" | "left" | "right";
        const dr = dir === "up" ? -1 : dir === "down" ? 1 : 0;
        const dc = dir === "left" ? -1 : dir === "right" ? 1 : 0;
        let last = { r, c };
        while (r + dr >= 0 && r + dr < s.rows && c + dc >= 0 && c + dc < s.cols) {
            r += dr;
            c += dc;
            if (this.cellAt(r, c).text === "") break;
            last = { r, c };
        }
        return { row: last.r, col: last.c };
    }

    private locate(args: any): any {
        const v = this.view();
        if (!v.map) return { view: args.row };
        return { view: v.map.indexOf(args.row) };
    }
}






