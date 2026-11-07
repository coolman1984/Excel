// Shared wire types between the UI and the host (plan 2.3) and the mock host.
// These mirror the JSON the C# host sends; keep them in sync with the contract.

/** Cell kind: 0 empty 1 number 2 text 3 bool 4 error (plan 2.3). */
export type CellKindCode = 0 | 1 | 2 | 3 | 4;

/** Alignment: 0 left 1 right 2 centre (plan 2.3). */
export type AlignCode = 0 | 1 | 2;

/** A style delta as JSON (contract "Wire format additions"). */
export interface CellStyle {
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    fontColor?: string; // "#RRGGBB"
    fill?: string; // "#RRGGBB"
    align?: AlignCode;
    numberFormat?: string;
    border?: string; // "none" | "all" | "outside" | "bottom"
}

/** One cell as sent in `rows` (plan 2.3): [text, kind, align, color?, warn?, style?]. */
export interface CellData {
    text: string;
    kind: CellKindCode;
    align: AlignCode;
    color?: string;
    warn?: boolean;
    style?: CellStyle;
}

/** A view row: [rowNumber, [cell, ...]] in wire form, kept decoded here. */
export interface WireRow {
    rowNumber: number;
    cells: CellData[];
}

export interface SheetInfo {
    index: number;
    name: string;
    hidden: boolean;
    rows: number;
    cols: number;
    firstRow: number;
    firstCol: number;
    colWidthsPx: number[];
    frozenRows: number;
    frozenCols: number;
}

export interface WorkbookInfo {
    name: string;
    path: string;
    sheets: SheetInfo[];
    dirty: boolean;
}

export interface Condition {
    kind: "values" | "number" | "text";
    include?: string[];
    blanks?: boolean;
    op?: string;
    a?: number;
    b?: number;
    value?: string;
}

export interface SortKey {
    col: number;
    desc: boolean;
}

export interface Stats {
    count: number;
    numbers: number;
    sum: number;
    average: number;
    min: number;
    max: number;
}

export interface EdgeResult {
    row: number;
    col: number;
}
