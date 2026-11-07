// Save prompt shown when the window is closed with unsaved changes. Excel's wording and buttons.
// Lane B owns every file in src/dialogs; this one calls nothing on the host: it just answers.
import type { CSSProperties } from "react";

const FONT = 'Aptos, "Segoe UI Variable", "Segoe UI", Calibri, sans-serif';

const overlay: CSSProperties = {
    position: "fixed", inset: 0, background: "rgba(0,0,0,0.15)", zIndex: 2000,
    display: "flex", alignItems: "center", justifyContent: "center", fontFamily: FONT,
};
const box: CSSProperties = {
    width: 400, background: "#fff", border: "1px solid #D2D0CE",
    boxShadow: "0 8px 24px rgba(0,0,0,0.3)", borderRadius: 2, overflow: "hidden",
};
const header: CSSProperties = { background: "#107C41", color: "#fff", padding: "8px 14px", fontSize: 13, fontWeight: 600 };
const body: CSSProperties = { padding: "20px 18px", fontSize: 13, color: "#1b1b1b", lineHeight: 1.4 };
const row: CSSProperties = { display: "flex", gap: 8, justifyContent: "flex-end", padding: "14px 16px 16px" };
const btn: CSSProperties = {
    minWidth: 88, padding: "5px 14px", fontSize: 13, border: "1px solid #8A8886",
    background: "#F3F2F1", color: "#1b1b1b", borderRadius: 2, cursor: "pointer", fontFamily: FONT,
};
const primary: CSSProperties = { ...btn, background: "#107C41", borderColor: "#107C41", color: "#fff" };

/** "Want to save your changes to X?" Save / Don't Save / Cancel. */
export default function ClosePrompt(props: {
    fileName: string;
    onAnswer: (action: "save" | "dontsave" | "cancel") => void;
}) {
    return (
        <div style={overlay} role="dialog" aria-modal="true">
            <div style={box}>
                <div style={header}>Excel</div>
                <div style={body}>Want to save your changes to &ldquo;{props.fileName}&rdquo;?</div>
                <div style={row}>
                    <button style={primary} onClick={() => props.onAnswer("save")}>
                        Save
                    </button>
                    <button style={btn} onClick={() => props.onAnswer("dontsave")}>
                        Don&apos;t Save
                    </button>
                    <button style={btn} onClick={() => props.onAnswer("cancel")} autoFocus>
                        Cancel
                    </button>
                </div>
            </div>
        </div>
    );
}
