// Find and Replace, modelled on Excel's dialog. Uses the find events and the replace command.
import { useEffect, useRef, useState, type CSSProperties } from "react";
import { call, on } from "../bridge";

const FONT = 'Aptos, "Segoe UI Variable", "Segoe UI", Calibri, sans-serif';

const overlay: CSSProperties = {
    position: "fixed", inset: 0, background: "rgba(0,0,0,0.12)", zIndex: 2000,
    display: "flex", alignItems: "center", justifyContent: "center", fontFamily: FONT,
};
const box: CSSProperties = {
    width: 540, background: "#fff", border: "1px solid #D2D0CE",
    boxShadow: "0 8px 24px rgba(0,0,0,0.3)", borderRadius: 2, overflow: "hidden",
};
const tabBar: CSSProperties = { display: "flex", alignItems: "stretch", background: "#F3F3F3", borderBottom: "1px solid #E1E1E1" };
const tab: CSSProperties = {
    padding: "8px 16px", fontSize: 13, border: "none", borderBottom: "2px solid transparent",
    background: "transparent", cursor: "pointer", color: "#444", fontFamily: FONT,
};
const tabActive: CSSProperties = { color: "#107C41", borderBottomColor: "#107C41", fontWeight: 600 };
const closeX: CSSProperties = { border: "none", background: "transparent", cursor: "pointer", fontSize: 16, color: "#666", padding: "0 12px" };
const body: CSSProperties = { padding: "14px 18px", fontSize: 13, color: "#1b1b1b" };
const label: CSSProperties = { display: "block", marginBottom: 4, color: "#444" };
const label2: CSSProperties = { display: "block", margin: "12px 0 4px", color: "#444" };
const input: CSSProperties = {
    width: "100%", boxSizing: "border-box", padding: "5px 8px", fontSize: 13, fontFamily: FONT,
    border: "1px solid #8A8886", borderRadius: 2, outlineColor: "#107C41",
};
const sel: CSSProperties = { padding: "3px 6px", fontSize: 13, fontFamily: FONT, border: "1px solid #8A8886", borderRadius: 2, background: "#fff" };
const checkRow: CSSProperties = { display: "flex", gap: 6, alignItems: "center", marginTop: 10, color: "#1b1b1b" };
const progress: CSSProperties = { marginTop: 10, color: "#0F703B", fontSize: 12 };
const hitList: CSSProperties = { marginTop: 10, height: 150, overflowY: "auto", border: "1px solid #E1E1E1", borderRadius: 2, background: "#fff" };
const empty: CSSProperties = { padding: 10, color: "#999" };
const hitRow: CSSProperties = { display: "flex", gap: 10, padding: "3px 10px", cursor: "pointer", borderBottom: "1px solid #F3F3F3" };
const hitCell: CSSProperties = { minWidth: 78, color: "#107C41", fontVariantNumeric: "tabular-nums" };
const hitText: CSSProperties = { color: "#1b1b1b", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" };
const row: CSSProperties = { display: "flex", gap: 8, justifyContent: "flex-end", padding: "0 16px 16px" };
const btn: CSSProperties = {
    minWidth: 84, padding: "5px 14px", fontSize: 13, border: "1px solid #8A8886",
    background: "#F3F2F1", color: "#1b1b1b", borderRadius: 2, cursor: "pointer", fontFamily: FONT,
};
const primary: CSSProperties = { ...btn, background: "#107C41", borderColor: "#107C41", color: "#fff" };

type Hit = [number, number, number, string]; // sheet, row, col, text

function colLetter(col: number): string {
    let n = col, s = "";
    do { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; } while (n >= 0);
    return s;
}

function tabStyle(active: boolean): CSSProperties {
    return active ? { ...tab, ...tabActive } : tab;
}

/** Find / Replace. Find All fills the list; a click on a hit asks the app to go there. */
export default function FindReplace(props: {
    sheet: number;
    initialTab: "find" | "replace";
    onSelect: (sheet: number, row: number, col: number) => void;
    onClose: () => void;
}) {
    const [tab, setTab] = useState<"find" | "replace">(props.initialTab);
    const [query, setQuery] = useState("");
    const [replacement, setReplacement] = useState("");
    const [within, setWithin] = useState<"sheet" | "workbook">("sheet");
    const [matchCase, setMatchCase] = useState(false);
    const [hits, setHits] = useState<Hit[]>([]);
    const [percent, setPercent] = useState(0);
    const [status, setStatus] = useState("");
    const [running, setRunning] = useState(false);
    const cursor = useRef(0);

    useEffect(() => {
        const offHits = on("findHits", (e: any) => {
            const batch: Hit[] = Array.isArray(e?.hits) ? e.hits : [];
            if (batch.length) setHits((prev) => prev.concat(batch));
        });
        const offProgress = on("findProgress", (e: any) => {
            setPercent(typeof e?.percent === "number" ? e.percent : 0);
        });
        const offDone = on("findDone", (e: any) => {
            const total = typeof e?.total === "number" ? e.total : 0;
            setRunning(false);
            setStatus(`${total} match${total === 1 ? "" : "es"}${e?.capped ? " (first batch shown)" : ""}`);
        });
        return () => { offHits(); offProgress(); offDone(); };
    }, []);

    const scope = within === "sheet" ? props.sheet : -1;

    const runFind = async () => {
        if (!query) return;
        setHits([]);
        cursor.current = 0;
        setPercent(0);
        setStatus("");
        setRunning(true);
        await call("find", { sheet: scope, query, matchCase });
    };

    const findNext = async () => {
        if (!query) return;
        if (hits.length === 0) {
            await runFind();
            return;
        }
        const h = hits[cursor.current % hits.length];
        cursor.current = (cursor.current + 1) % hits.length;
        props.onSelect(h[0], h[1], h[2]);
    };

    const replaceOne = async () => {
        if (!query) return;
        const h = hits[cursor.current] ?? hits[0];
        const at = h ? [h[1], h[2]] : undefined;
        const r = await call<{ changed: number; count: number }>("replace", { sheet: scope, query, replacement, all: false, at });
        setStatus(`Replaced ${r?.count ?? 0}`);
    };

    const replaceAll = async () => {
        if (!query) return;
        const r = await call<{ changed: number; count: number }>("replace", { sheet: scope, query, replacement, all: true });
        setStatus(`Replaced ${r?.count ?? 0}`);
    };

    return (
        <div style={overlay} onMouseDown={props.onClose}>
            <div style={box} role="dialog" aria-modal="true" onMouseDown={(e) => e.stopPropagation()}>
                <div style={tabBar}>
                    <button style={tabStyle(tab === "find")} onClick={() => setTab("find")}>Find</button>
                    <button style={tabStyle(tab === "replace")} onClick={() => setTab("replace")}>Replace</button>
                    <div style={{ flex: 1 }} />
                    <button style={closeX} onClick={props.onClose} aria-label="Close">&times;</button>
                </div>
                <div style={body}>
                    <label style={label}>Find what:</label>
                    <input
                        style={input}
                        value={query}
                        autoFocus
                        onChange={(e) => setQuery(e.target.value)}
                        onKeyDown={(e) => { if (e.key === "Enter") void findNext(); }}
                    />
                    {tab === "replace" && (
                        <>
                            <label style={label2}>Replace with:</label>
                            <input style={input} value={replacement} onChange={(e) => setReplacement(e.target.value)} />
                        </>
                    )}
                    <label style={checkRow}>
                        <input type="checkbox" checked={matchCase} onChange={(e) => setMatchCase(e.target.checked)} />
                        Match case
                    </label>
                    <div style={checkRow}>
                        <span>Within:</span>
                        <select style={sel} value={within} onChange={(e) => setWithin(e.target.value as "sheet" | "workbook")}>
                            <option value="sheet">Sheet</option>
                            <option value="workbook">Workbook</option>
                        </select>
                    </div>
                    {running && <div style={progress}>Searching... {percent}%</div>}
                    {!running && status && <div style={progress}>{status}</div>}
                    <div style={hitList}>
                        {hits.length === 0 && <div style={empty}>No results</div>}
                        {hits.map((h, i) => (
                            <div key={i} style={hitRow} onClick={() => props.onSelect(h[0], h[1], h[2])} title="Go to this cell">
                                <span style={hitCell}>{`${colLetter(h[2])}${h[1] + 1}`}</span>
                                <span style={hitText}>{h[3]}</span>
                            </div>
                        ))}
                    </div>
                </div>
                <div style={row}>
                    <button style={btn} onClick={() => void runFind()}>Find All</button>
                    <button style={btn} onClick={() => void findNext()}>Find Next</button>
                    {tab === "replace" && <button style={btn} onClick={() => void replaceOne()}>Replace</button>}
                    {tab === "replace" && <button style={primary} onClick={() => void replaceAll()}>Replace All</button>}
                    <button style={btn} onClick={props.onClose}>Close</button>
                </div>
            </div>
        </div>
    );
}

