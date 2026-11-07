// Plan 4 shared wire contract (lead-owned). Mirrors ExcelApp/Data/Contracts.cs and plan 4 section 4.
// Every "data.*" command answers JSON in camelCase; errors reject with a plain sentence.

export type ColType = "number" | "text" | "date" | "bool";
export type Granularity = "day" | "week" | "month";

export interface ColumnInfo { name: string; type: ColType }
export interface TableRef { kind: "report" | "table" | "version" | "sandbox"; id: string; asOf?: string; version?: number; truth?: "latest" | "locked" }
export interface TableInfo { kind: TableRef["kind"]; id: string; name: string; group: string; rows: number; columns: ColumnInfo[]; loadedAt: string }

/** A page of rows from a result (data.page). Cell text is already formatted for display. */
export interface ResultPage { resultId: string; from: number; rows: string[][]; kinds: number[][] }
/** A query result handle (data.run). Columns exclude the reserved "_" lineage columns. */
export interface ResultInfo { resultId: string; rows: number; columns: ColumnInfo[]; ms: number; groups?: boolean }

// ---- Query spec (lane B executes; the interface builds it) ----
export type Pred =
    | { op: "and" | "or"; items: Pred[] }
    | { op: "not"; item: Pred }
    | { op: "eq" | "ne" | "gt" | "ge" | "lt" | "le"; col: string; value: number | string | boolean }
    | { op: "between"; col: string; a: number | string; b: number | string }
    | { op: "in"; col: string; values: (string | number)[] }
    | { op: "contains" | "begins" | "ends"; col: string; value: string }
    | { op: "blank" | "notBlank"; col: string }
    | { op: "duplicate"; cols: string[] }                       // rows whose key appears more than once
    | { op: "colCompare"; col: string; cmp: "eq" | "ne" | "gt" | "ge" | "lt" | "le"; other: string; tolerancePct?: number };
export interface Computed { name: string; expr: string }        // expr grammar: plan 4 section 5.B.3
export interface Aggregate { fn: "sum" | "count" | "countDistinct" | "avg" | "min" | "max"; col?: string; as: string }
export interface JoinSpec { right: TableRef; on: { left: string; right: string }[]; kind: "inner" | "left" | "anti"; fuzzy?: "none" | "normalized" | "similar"; minScore?: number; prefix?: string }
export interface QuerySpec {
    source: TableRef;
    joins?: JoinSpec[];
    where?: Pred;
    computed?: Computed[];
    groupBy?: string[];
    aggregates?: Aggregate[];
    having?: Pred;
    sort?: { col: string; desc: boolean }[];
    limit?: number;
    columns?: string[];                                        // projection; empty = all
}

// ---- Ingestion ----
export interface WeekRule { firstDay: string; numbering: "iso" | "fiscal"; fiscalYearStart?: string }
export interface LockRule { workingDay: number; holidays: string[] }
export interface TypeFix { column: string; fix: "date" | "number" | "text" | "padZeros" | "trim" | "upper"; width?: number; dateFormat?: string }
export interface IngestRecipe {
    id: string; name: string; folders: string[]; patterns: string[]; headerSignature: string[]; headerRow: number; sheet: string;
    dateColumn: string; granularity: Granularity; week: WeekRule; keyColumns: string[]; typeFixes: TypeFix[];
    partialThreshold: number; truth: "latest" | "locked"; lock: LockRule; keepVersions: number; dependentViews: string[];
}
export interface CleanupItem { column: string; fix: string; cells: number; example: string }
export interface QualityFlag { kind: string; column: string; count: number; message: string; sampleRows: number[] }
export interface LoadRecord {
    id: string; recipeId: string; sourcePath: string; fileHash: string; contentHash: string; loadedAt: string; periods: string[];
    rows: number; added: number; removed: number; changed: number;
    status: "loaded" | "duplicate" | "quarantined" | "rejected" | "rolledBack" | "failed"; reason?: string;
    versions: Record<string, number>; cleanup?: { items: CleanupItem[] }; quality?: { flags: QualityFlag[] };
}
export interface VersionInfo { reportTypeId: string; period: string; version: number; loadId: string; loadedAt: string; rows: number; added: number; removed: number; changed: number; current: boolean; locked: boolean; postClose: boolean; note: string }
export interface DiffResult { resultId: string; added: number; removed: number; changed: number; changedColumns: Record<string, number> }

// ---- Saved objects (configuration; lane B) ----
export interface SavedView { id: string; name: string; spec: QuerySpec; exception: boolean; folder: string }
export interface Link { id: string; name: string; left: TableRef; right: TableRef; on: { left: string; right: string }[]; compare: { left: string; right: string; as: string }[]; fuzzy: "none" | "normalized" | "similar" }
export interface AlertRule { id: string; name: string; viewId?: string; spec?: QuerySpec; measure: string; op: "gt" | "ge" | "lt" | "le" | "eq" | "ne" | "changePct"; value: number; reportTypes: string[]; enabled: boolean }
export interface AlertEvent { id: string; ruleId: string; name: string; at: string; value: number; message: string }
export interface RecipeStep { kind: "query" | "export" | "pack"; spec?: QuerySpec; viewId?: string; export?: { target: string; template?: string; sheet: string; startCell: string } }
export interface ReportRecipe { id: string; name: string; steps: RecipeStep[]; schedule?: { kind: "daily" | "weekly" | "monthly" | "afterLoad"; time?: string; day?: number; reportType?: string }; enabled: boolean }
export interface Tile { id: string; title: string; spec: QuerySpec; measure: string; format: string; compareTo?: "previousPeriod" | "previousVersion" }
export interface Dashboard { id: string; name: string; tiles: Tile[] }
export interface TileValue { tileId: string; value: number | null; text: string; previous?: number | null; deltaPct?: number | null; asOf: string; error?: string }
export interface Sandbox { id: string; name: string; base: TableRef; edits: number; createdAt: string }
export interface LineageCell { file: string; sheet: string; cell: string; version: number; loadedAt: string; value: string }
export interface Question { text: string; spec?: QuerySpec; pattern?: string; explanation: string; ok: boolean }

/** Host events the interface listens to (bridge.on). */
export type DataEvent =
    | { event: "data.ingest"; stage: "noticed" | "recognized" | "compared" | "updated" | "notified" | "failed"; file: string; recipeId?: string; message: string }
    | { event: "data.loaded"; record: LoadRecord }
    | { event: "data.quarantine"; record: LoadRecord }
    | { event: "data.alert"; alert: AlertEvent }
    | { event: "data.recipeRun"; recipeId: string; ok: boolean; message: string }
    | { event: "data.notice"; message: string };
