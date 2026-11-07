// Sort dialog with up to three keys, like Excel's. Sends setHeaderRow and setSort to the host.
import { useState, type CSSProperties } from "react";
import { call } from "../bridge";

const FONT = 'Aptos, "Segoe UI Variable", "Segoe UI", Calibri, sans-serif';

const overlay: CSSProperties = {
    position: "fixed", inset: 0, background: "rgba(0,0,0,0.12)", zIndex: 2000,
    display: "flex", alignItems: "center", justifyContent: "center", fontFamily: FONT,
};
const box: CSSProperties = {
    width: 460, background: "#fff", border: "1px solid #D2D0CE",
    boxShadow: "0 8px 24px rgba(0,0,0,0.3)", borderRadius: 2, overflow: "hidden",
};
const header: CSSProperties = { background: "#F3F3F3", color: "#1b1b1b", padding: "8px 14px", fontSize: 13, fontWeight: 600, borderBottom: "1px solid #E1E1E1" };
const body: CSSProperties = { padding: "16px 18px", fontSize: 13, color: "#1b1b1b" };
const grid: CSSProperties = { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, alignItems: "center" };
const head: CSSProperties = { color: "#444", fontSize: 12, paddingBottom: 4 };
const sel: CSSProperties = {
    width: "100%", padding: "4px 6px", fontSize: 13, fontFamily: FONT,
    border: "1px solid #8A8886", borderRadius: 2, background: "#fff", color: "#1b1b1b",
};
const check: CSSProperties = { display: "flex", gap: 6, alignItems: "center", marginTop: 14, color: "#1b1b1b" };
const row: CSSProperties = { display: "flex", gap: 8, justifyContent: "flex-end", padding: "0 16px 16px" };
const btn: CSSProperties = {
    minWidth: 82, padding: "5px 14px", fontSize: 13, border: "1px solid #8A8886",
    background: "#F3F2F1", color: "#1b1b1b", borderRadius: 2, cursor: "pointer", fontFamily: FONT,
};
const primary: CSSProperties = { ...btn, background: "#107C41", borderColor: "#107C41", color: "#fff" };

type SortKey = { col: number; desc: boolean };

/** Up to three sort keys. OK applies headers and sort, then closes. */
export default function SortDialog(props: {
    sheet: number;
    colCount: number;
    letter: (col: number) => string;
    onClose: () => void;
}) {
    const [keys, setKeys] = useState<SortKey[]>([
        { col: 0, desc: false },
        { col: -1, desc: false },
        { col: -1, desc: false },
    ]);
    const [hasHeaders, setHasHeaders] = useState(true);
    const [busy, setBusy] = useState(false);

    const setKey = (i: number, patch: Partial<SortKey>) =>
        setKeys((k) => k.map((x, j) => (j === i ? { ...x, ...patch } : x)));

    const ok = async () => {
        setBusy(true);
        try {
            await call("setHeaderRow", { sheet: props.sheet, hasHeader: hasHeaders });
            const active = keys.filter((k) => k.col >= 0).map((k) => ({ col: k.col, desc: k.desc }));
            await call("setSort", { sheet: props.sheet, keys: active });
        } finally {
            props.onClose();
        }
    };

    const cols: number[] = [];
    for (let c = 0; c < props.colCount; c++) cols.push(c);

    return (
        <div style={overlay} onMouseDown={props.onClose}>
            <div style={box} role="dialog" aria-modal="true" onMouseDown={(e) => e.stopPropagation()}>
                <div style={header}>Sort</div>
                <div style={body}>
                    <div style={grid}>
                        <div style={head}>Column</div>
                        <div style={head}>Sort On</div>
                        <div style={head}>Order</div>
                        {keys.map((k, i) => (
                            <div key={i} style={{ display: "contents" }}>
                                <select
                                    style={sel}
                                    value={k.col}
                                    onChange={(e) => setKey(i, { col: Number(e.target.value) })}
                                >
                                    {i > 0 && <option value={-1}>(none)</option>}
                                    {cols.map((c) => (
                                        <option key={c} value={c}>
                                            {props.letter(c)}
                                        </option>
                                    ))}
                                </select>
                                <select style={sel} value="values" disabled>
                                    <option value="values">Values</option>
                                </select>
                                <select
                                    style={sel}
                                    value={k.desc ? "desc" : "asc"}
                                    onChange={(e) => setKey(i, { desc: e.target.value === "desc" })}
                                >
                                    <option value="asc">A to Z</option>
                                    <option value="desc">Z to A</option>
                                </select>
                            </div>
                        ))}
                    </div>
                    <label style={check}>
                        <input type="checkbox" checked={hasHeaders} onChange={(e) => setHasHeaders(e.target.checked)} />
                        My data has headers
                    </label>
                </div>
                <div style={row}>
                    <button style={primary} disabled={busy} onClick={ok}>
                        OK
                    </button>
                    <button style={btn} onClick={props.onClose}>
                        Cancel
                    </button>
                </div>
            </div>
        </div>
    );
}
