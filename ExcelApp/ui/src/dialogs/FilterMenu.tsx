// Excel's AutoFilter menu. Values list, search, number/text filters, sort and clear.
// It talks to the host through bridge: filterValues, setFilter, setSort, clearFilters.
import { useEffect, useState, type CSSProperties } from "react";
import { call } from "../bridge";

const FONT = 'Aptos, "Segoe UI Variable", "Segoe UI", Calibri, sans-serif';
const ROW = 22;      // one value row, in pixels
const VISIBLE = 11;  // rows shown at once; the rest are virtualised

const backdrop: CSSProperties = { position: "fixed", inset: 0, zIndex: 1500 };
const menu: CSSProperties = {
    position: "fixed", width: 264, background: "#fff", border: "1px solid #D2D0CE", borderRadius: 2,
    boxShadow: "0 6px 20px rgba(0,0,0,0.28)", fontFamily: FONT, fontSize: 13, color: "#1b1b1b", zIndex: 1501,
};
const item: CSSProperties = { padding: "6px 12px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" };
const itemDisabled: CSSProperties = { ...item, color: "#A19F9D", cursor: "default" };
const sep: CSSProperties = { height: 1, background: "#E1E1E1", margin: "4px 0" };
const search: CSSProperties = {
    width: "100%", boxSizing: "border-box", padding: "6px 10px", border: "none",
    borderBottom: "1px solid #E1E1E1", outline: "none", fontFamily: FONT, fontSize: 13,
};
const list: CSSProperties = { height: ROW * VISIBLE, overflowY: "auto", borderBottom: "1px solid #E1E1E1" };
const vrow: CSSProperties = { display: "flex", alignItems: "center", gap: 6, height: ROW, boxSizing: "border-box", padding: "0 10px" };
const rowText: CSSProperties = { flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" };
const count: CSSProperties = { color: "#999", fontSize: 12 };
const custom: CSSProperties = { padding: "8px 12px", borderBottom: "1px solid #E1E1E1" };
const tin: CSSProperties = { width: 84, padding: "3px 6px", border: "1px solid #8A8886", borderRadius: 2, fontFamily: FONT, fontSize: 13, marginRight: 6 };
const footer: CSSProperties = { display: "flex", gap: 8, justifyContent: "flex-end", padding: "10px 12px" };
const btn: CSSProperties = {
    minWidth: 74, padding: "4px 14px", fontSize: 13, border: "1px solid #8A8886",
    background: "#F3F2F1", color: "#1b1b1b", borderRadius: 2, cursor: "pointer", fontFamily: FONT,
};
const primary: CSSProperties = { ...btn, background: "#107C41", borderColor: "#107C41", color: "#fff" };
const note: CSSProperties = { padding: "6px 12px", color: "#777", fontSize: 12 };

type ValueRow = { text: string; count: number };
type FilterValues = { values: ValueRow[]; blanks: number; truncated: boolean; kind?: string };

const NUMBER_OPS: [string, string][] = [
    ["Equals...", "eq"], ["Does Not Equal...", "ne"], ["Greater Than...", "gt"],
    ["Greater Than or Equal To...", "ge"], ["Less Than...", "lt"], ["Less Than or Equal To...", "le"],
    ["Between...", "between"], ["Top 10...", "top10"], ["Above Average", "aboveAverage"], ["Below Average", "belowAverage"],
];
const TEXT_OPS: [string, string][] = [
    ["Equals...", "eq"], ["Does Not Equal...", "ne"], ["Begins With...", "begins"],
    ["Ends With...", "ends"], ["Contains...", "contains"], ["Does Not Contain...", "notContains"],
];

/** AutoFilter menu anchored at the header cell. Sends one condition on OK, or a sort/clear. */
export default function FilterMenu(props: {
    sheet: number;
    col: number;
    letter: string;
    anchor: { x: number; y: number };
    onClose: () => void;
}) {
    const [values, setValues] = useState<ValueRow[]>([]);
    const [blanks, setBlanks] = useState(0);
    const [truncated, setTruncated] = useState(false);
    const [kind, setKind] = useState("");
    const [search, setSearch] = useState("");
    const [selected, setSelected] = useState<Set<string>>(new Set());
    const [includeBlanks, setIncludeBlanks] = useState(true);
    const [loading, setLoading] = useState(false);
    const [scrollTop, setScrollTop] = useState(0);
    const [submenu, setSubmenu] = useState<"" | "number" | "text">("");
    const [pendingOp, setPendingOp] = useState("");
    const [aText, setAText] = useState("");
    const [bText, setBText] = useState("");
    const [valueText, setValueText] = useState("");

    const load = async (searchText?: string) => {
        setLoading(true);
        try {
            const r = await call<FilterValues>("filterValues", { sheet: props.sheet, col: props.col, search: searchText });
            const vals: ValueRow[] = Array.isArray(r?.values) ? r.values : [];
            setValues(vals);
            setBlanks(typeof r?.blanks === "number" ? r.blanks : 0);
            setTruncated(!!r?.truncated);
            if (r?.kind) setKind(r.kind);
            if (searchText === undefined) {
                setSelected(new Set(vals.map((v) => v.text)));
                setIncludeBlanks(true);
            }
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        void load();
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    const total = values.length;
    const start = Math.max(0, Math.floor(scrollTop / ROW) - 2);
    const end = Math.min(total, start + VISIBLE + 6);
    const windowed = values.slice(start, end);
    const allChecked = total > 0 && selected.size === total && includeBlanks;
    const filterKind = kind === "number" ? "number" : "text";

    const toggle = (text: string, on: boolean) => {
        setSelected((prev) => {
            const next = new Set(prev);
            if (on) next.add(text); else next.delete(text);
            return next;
        });
    };

    const toggleAll = () => {
        if (allChecked) {
            setSelected(new Set());
            setIncludeBlanks(false);
        } else {
            setSelected(new Set(values.map((v) => v.text)));
            setIncludeBlanks(true);
        }
    };

    const onSearch = (v: string) => {
        setSearch(v);
        void load(v);
    };

    const sortBy = async (desc: boolean) => {
        await call("setSort", { sheet: props.sheet, keys: [{ col: props.col, desc }] });
        props.onClose();
    };

    const clearFilter = async () => {
        await call("setFilter", { sheet: props.sheet, col: props.col, condition: null });
        props.onClose();
    };

    const ok = async () => {
        const condition = { kind: "values", include: Array.from(selected), blanks: includeBlanks };
        await call("setFilter", { sheet: props.sheet, col: props.col, condition });
        props.onClose();
    };

    const chooseOp = (op: string) => {
        setPendingOp(op);
        setSubmenu(filterKind);
    };

    const applyCustom = async () => {
        if (!pendingOp) return;
        let condition: Record<string, unknown>;
        if (filterKind === "number") {
            const a = Number(aText) || 0;
            condition = { kind: "number", op: pendingOp, a: pendingOp === "top10" ? (a > 0 ? a : 10) : a };
            if (pendingOp === "between") condition.b = Number(bText) || 0;
        } else {
            condition = { kind: "text", op: pendingOp, value: valueText };
        }
        await call("setFilter", { sheet: props.sheet, col: props.col, condition });
        props.onClose();
    };

    return (
        <div style={backdrop} onMouseDown={props.onClose}>
            <div style={{ ...menu, left: props.anchor.x, top: props.anchor.y }} onMouseDown={(e) => e.stopPropagation()}>
                <div style={item} onClick={() => void sortBy(false)}>Sort A to Z</div>
                <div style={item} onClick={() => void sortBy(true)}>Sort Z to A</div>
                <div style={itemDisabled}>Sort by Color</div>
                <div style={sep} />
                <div style={item} onClick={() => void clearFilter()}>Clear Filter From &ldquo;{props.letter}&rdquo;</div>
                <div style={itemDisabled}>Filter by Color</div>
                <div style={sep} />
                <div style={item} onClick={() => setSubmenu(submenu === filterKind ? "" : filterKind)}>
                    {filterKind === "number" ? "Number Filters" : "Text Filters"}
                    <span>{submenu === filterKind ? "\u25be" : "\u25b8"}</span>
                </div>
                {submenu === filterKind && (
                    <div style={{ borderTop: "1px solid #E1E1E1", borderBottom: "1px solid #E1E1E1", background: "#FAFAFA" }}>
                        {(filterKind === "number" ? NUMBER_OPS : TEXT_OPS).map(([opLabel, op]) => (
                            <div key={op} style={{ ...item, paddingLeft: 24 }} onClick={() => chooseOp(op)}>{opLabel}</div>
                        ))}
                    </div>
                )}
                {pendingOp !== "" && (
                    <div style={custom}>
                        {filterKind === "number" ? (
                            pendingOp === "between" ? (
                                <div>
                                    <input style={tin} value={aText} onChange={(e) => setAText(e.target.value)} placeholder="from" autoFocus /> and
                                    <input style={tin} value={bText} onChange={(e) => setBText(e.target.value)} placeholder="to" />
                                </div>
                            ) : pendingOp === "top10" || pendingOp === "aboveAverage" || pendingOp === "belowAverage" ? (
                                <div style={{ color: "#555" }}>
                                    {pendingOp === "top10" ? "Top 10 items" : pendingOp === "aboveAverage" ? "Above average" : "Below average"}
                                </div>
                            ) : (
                                <input style={tin} value={aText} onChange={(e) => setAText(e.target.value)} placeholder="value" autoFocus />
                            )
                        ) : (
                            <input style={tin} value={valueText} onChange={(e) => setValueText(e.target.value)} placeholder="value" autoFocus />
                        )}
                        <div style={{ marginTop: 6 }}>
                            <button style={btn} onClick={() => void applyCustom()}>Apply</button>
                        </div>
                    </div>
                )}
                <div style={sep} />
                <input style={search} placeholder="Search" value={search} onChange={(e) => onSearch(e.target.value)} />
                <div style={list} onScroll={(e) => setScrollTop((e.target as HTMLDivElement).scrollTop)}>
                    <div style={{ paddingTop: start * ROW, paddingBottom: Math.max(0, total - end) * ROW }}>
                        <label style={vrow}>
                            <input type="checkbox" checked={allChecked} onChange={toggleAll} />
                            <span style={rowText}>(Select All)</span>
                        </label>
                        {blanks > 0 && (
                            <label style={vrow}>
                                <input type="checkbox" checked={includeBlanks} onChange={(e) => setIncludeBlanks(e.target.checked)} />
                                <span style={rowText}>(Blanks)</span>
                                <span style={count}>{blanks}</span>
                            </label>
                        )}
                        {windowed.map((v) => (
                            <label key={v.text} style={vrow}>
                                <input type="checkbox" checked={selected.has(v.text)} onChange={(e) => toggle(v.text, e.target.checked)} />
                                <span style={rowText} title={v.text}>{v.text}</span>
                                <span style={count}>{v.count}</span>
                            </label>
                        ))}
                        {!loading && total === 0 && <div style={note}>No values</div>}
                    </div>
                </div>
                {truncated && <div style={note}>Showing the first 10,000 values. Use the search box to narrow the list.</div>}
                <div style={footer}>
                    <button style={primary} onClick={() => void ok()}>OK</button>
                    <button style={btn} onClick={props.onClose}>Cancel</button>
                </div>
            </div>
        </div>
    );
}

