// Format Cells (Ctrl+1), the Number tab: pick a category or type a format code; OK applies it to the selection.
import { useState, type CSSProperties } from "react";
import { NUMBER_FORMATS } from "../numberFormat";

const FONT = 'Aptos, "Segoe UI Variable", "Segoe UI", Calibri, sans-serif';
const overlay: CSSProperties = {
    position: "fixed", inset: 0, background: "rgba(0,0,0,0.12)", zIndex: 2000,
    display: "flex", alignItems: "center", justifyContent: "center", fontFamily: FONT,
};
const box: CSSProperties = {
    width: 420, background: "#fff", border: "1px solid #D2D0CE",
    boxShadow: "0 8px 24px rgba(0,0,0,0.3)", borderRadius: 2, overflow: "hidden",
};
const header: CSSProperties = { background: "#F3F3F3", color: "#1b1b1b", padding: "8px 14px", fontSize: 13, fontWeight: 600, borderBottom: "1px solid #E1E1E1" };
const body: CSSProperties = { padding: "14px 18px", fontSize: 13, color: "#1b1b1b", display: "flex", gap: 14 };
const list: CSSProperties = { width: 140, height: 220, overflowY: "auto", border: "1px solid #8A8886", borderRadius: 2, padding: 0, margin: 0, listStyle: "none" };
const item: CSSProperties = { padding: "3px 8px", cursor: "pointer" };
const chosen: CSSProperties = { ...item, background: "#C6EFCE", color: "#0B5A2B" };
const label: CSSProperties = { display: "block", marginBottom: 6, color: "#444" };
const input: CSSProperties = {
    width: "100%", boxSizing: "border-box", padding: "5px 8px", fontSize: 13, fontFamily: "Consolas, monospace",
    border: "1px solid #8A8886", borderRadius: 2, outlineColor: "#107C41",
};
const hint: CSSProperties = { marginTop: 8, color: "#777", fontSize: 12 };
const row: CSSProperties = { display: "flex", gap: 8, justifyContent: "flex-end", padding: "0 16px 16px" };
const btn: CSSProperties = {
    minWidth: 82, padding: "5px 14px", fontSize: 13, border: "1px solid #8A8886",
    background: "#F3F2F1", color: "#1b1b1b", borderRadius: 2, cursor: "pointer", fontFamily: FONT,
};
const primary: CSSProperties = { ...btn, background: "#107C41", borderColor: "#107C41", color: "#fff" };

export default function FormatCells(props: { code: string; onApply: (code: string) => void; onClose: () => void }) {
    const [code, setCode] = useState(props.code || "General");
    const match = NUMBER_FORMATS.find((f) => f.code === code);
    const ok = () => {
        const c = code.trim();
        if (c) props.onApply(c);
        props.onClose();
    };
    return (
        <div style={overlay} onMouseDown={props.onClose}>
            <div style={box} role="dialog" aria-modal="true" aria-label="Format Cells" onMouseDown={(e) => e.stopPropagation()}
                onKeyDown={(e) => { if (e.key === "Escape") props.onClose(); }}>
                <div style={header}>Format Cells</div>
                <div style={body}>
                    <ul style={list} role="listbox" aria-label="Category">
                        {NUMBER_FORMATS.map((f) => (
                            <li key={f.label} role="option" aria-selected={match === f} style={match === f ? chosen : item}
                                onClick={() => setCode(f.code)} onDoubleClick={() => { props.onApply(f.code); props.onClose(); }}>
                                {f.label}
                            </li>
                        ))}
                        <li role="option" aria-selected={!match} style={!match ? chosen : item}>Custom</li>
                    </ul>
                    <div style={{ flex: 1 }}>
                        <label style={label} htmlFor="format-code">Type:</label>
                        <input id="format-code" style={input} value={code} autoFocus spellCheck={false}
                            onChange={(e) => setCode(e.target.value)}
                            onKeyDown={(e) => { if (e.key === "Enter") ok(); }} />
                        <div style={hint}>Pick a category or type a format code, such as 0.000 or dd/mm/yyyy. The cells keep their values; only how they show changes.</div>
                    </div>
                </div>
                <div style={row}>
                    <button style={primary} onClick={ok}>OK</button>
                    <button style={btn} onClick={props.onClose}>Cancel</button>
                </div>
            </div>
        </div>
    );
}
