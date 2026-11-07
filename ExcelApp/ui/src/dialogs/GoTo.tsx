// Go To (Ctrl+G). Type a cell or range and press Enter; the app jumps and selects it.
import { useState, type CSSProperties } from "react";

const FONT = 'Aptos, "Segoe UI Variable", "Segoe UI", Calibri, sans-serif';

const overlay: CSSProperties = {
    position: "fixed", inset: 0, background: "rgba(0,0,0,0.12)", zIndex: 2000,
    display: "flex", alignItems: "center", justifyContent: "center", fontFamily: FONT,
};
const box: CSSProperties = {
    width: 340, background: "#fff", border: "1px solid #D2D0CE",
    boxShadow: "0 8px 24px rgba(0,0,0,0.3)", borderRadius: 2, overflow: "hidden",
};
const header: CSSProperties = { background: "#F3F3F3", color: "#1b1b1b", padding: "8px 14px", fontSize: 13, fontWeight: 600, borderBottom: "1px solid #E1E1E1" };
const body: CSSProperties = { padding: "16px 18px", fontSize: 13, color: "#1b1b1b" };
const label: CSSProperties = { display: "block", marginBottom: 6, color: "#444" };
const input: CSSProperties = {
    width: "100%", boxSizing: "border-box", padding: "5px 8px", fontSize: 13, fontFamily: FONT,
    border: "1px solid #8A8886", borderRadius: 2, outlineColor: "#107C41",
};
const hint: CSSProperties = { marginTop: 8, color: "#777", fontSize: 12 };
const row: CSSProperties = { display: "flex", gap: 8, justifyContent: "flex-end", padding: "0 16px 16px" };
const btn: CSSProperties = {
    minWidth: 82, padding: "5px 14px", fontSize: 13, border: "1px solid #8A8886",
    background: "#F3F2F1", color: "#1b1b1b", borderRadius: 2, cursor: "pointer", fontFamily: FONT,
};
const primary: CSSProperties = { ...btn, background: "#107C41", borderColor: "#107C41", color: "#fff" };

/** Go To dialog: an address box that hands the typed address back to the app. */
export default function GoTo(props: { onGo: (address: string) => void; onClose: () => void }) {
    const [address, setAddress] = useState("");

    const go = () => {
        const a = address.trim();
        if (a) props.onGo(a); else props.onClose();
    };

    return (
        <div style={overlay} onMouseDown={props.onClose}>
            <div style={box} role="dialog" aria-modal="true" onMouseDown={(e) => e.stopPropagation()}>
                <div style={header}>Go To</div>
                <div style={body}>
                    <label style={label} htmlFor="goto-address">
                        Go to:
                    </label>
                    <input
                        id="goto-address"
                        style={input}
                        value={address}
                        autoFocus
                        placeholder="B7 or A1:C9"
                        onChange={(e) => setAddress(e.target.value)}
                        onKeyDown={(e) => {
                            if (e.key === "Enter") go();
                            if (e.key === "Escape") props.onClose();
                        }}
                    />
                    <div style={hint}>Type a cell or a range, then press Enter.</div>
                </div>
                <div style={row}>
                    <button style={primary} onClick={go}>
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
