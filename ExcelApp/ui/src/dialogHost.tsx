import * as React from "react";

// Lane B owns ExcelApp/ui/src/dialogs/**. This module imports those components by
// their contract names. import.meta.glob resolves them at build time and does NOT
// fail when the folder is empty, so the build passes before lane B lands; until
// then tiny placeholders are used (and say so on screen).

const mods = import.meta.glob("./dialogs/*.tsx", { eager: true }) as Record<string, any>;

const real: Record<string, React.ComponentType<any>> = {};
for (const path in mods) {
    const base = path.split("/").pop()!.replace(/\.tsx$/, "");
    const comp = mods[path]?.default;
    if (typeof comp === "function") real[base] = comp;
}

/** True once lane B's dialogs exist. */
export const hasRealDialogs = Object.keys(real).length > 0;

function Modal({ title, children, onClose }: { title: string; children: React.ReactNode; onClose: () => void }): React.ReactElement {
    return (
        <div className="dialogBackdrop" onClick={onClose}>
            <div className="dialog" onClick={(e) => e.stopPropagation()} data-testid="dialog">
                <div className="dialogTitle">
                    {title}
                    <button type="button" className="dialogClose" onClick={onClose} aria-label="Close">
                        &times;
                    </button>
                </div>
                <div className="dialogBody">{children}</div>
            </div>
        </div>
    );
}

function PlaceholderFilterMenu({ col, letter, onClose }: any): React.ReactElement {
    return (
        <Modal title={"Filter - column " + letter} onClose={onClose}>
            <p className="dialogNote">Column {col}. Waiting for the filter dialog (lane B).</p>
            <button type="button" onClick={onClose}>
                Close
            </button>
        </Modal>
    );
}

function PlaceholderSortDialog({ onClose }: any): React.ReactElement {
    return (
        <Modal title="Sort" onClose={onClose}>
            <p className="dialogNote">Sort by up to three keys. Waiting for the sort dialog (lane B).</p>
            <button type="button" onClick={onClose}>
                Close
            </button>
        </Modal>
    );
}

function PlaceholderFindReplace({ initialTab, onClose }: any): React.ReactElement {
    return (
        <Modal title={initialTab === "replace" ? "Find and Replace" : "Find"} onClose={onClose}>
            <p className="dialogNote">Waiting for the find dialog (lane B).</p>
            <button type="button" onClick={onClose}>
                Close
            </button>
        </Modal>
    );
}

function PlaceholderGoTo({ onGo, onClose }: any): React.ReactElement {
    const [addr, setAddr] = React.useState("A1");
    return (
        <Modal title="Go To" onClose={onClose}>
            <input value={addr} onChange={(e) => setAddr(e.target.value)} aria-label="Reference" />
            <button
                type="button"
                onClick={() => {
                    onGo(addr);
                    onClose();
                }}
            >
                Go To
            </button>
            <button type="button" onClick={onClose}>
                Cancel
            </button>
        </Modal>
    );
}

function PlaceholderSaveProgress({ onDone }: any): React.ReactElement {
    return (
        <Modal
            title="Saving"
            onClose={() => {
                /* cancel not available in the placeholder */
            }}
        >
            <p className="dialogNote">Saving through Excel... (placeholder; lane B)</p>
            <button type="button" onClick={() => onDone(true)}>
                Continue
            </button>
        </Modal>
    );
}

function PlaceholderClosePrompt({ fileName, onAnswer }: any): React.ReactElement {
    return (
        <Modal title="ExcelApp" onClose={() => onAnswer("cancel")}>
            <p>Want to save your changes to {fileName}?</p>
            <div className="dialogButtons">
                <button type="button" onClick={() => onAnswer("save")}>
                    Save
                </button>
                <button type="button" onClick={() => onAnswer("dontsave")}>
                    Don&apos;t Save
                </button>
                <button type="button" onClick={() => onAnswer("cancel")}>
                    Cancel
                </button>
            </div>
        </Modal>
    );
}

function MessageBox({ title, text, onClose }: { title?: string; text: string; onClose: () => void }): React.ReactElement {
    const ok = React.useRef<HTMLButtonElement | null>(null);
    React.useEffect(() => ok.current?.focus(), []);
    return (
        <Modal title={title ?? "ExcelApp"} onClose={onClose}>
            <p className="dialogMessage" data-testid="message-text">{text}</p>
            <div className="dialogButtons">
                <button type="button" ref={ok} onClick={onClose} onKeyDown={(e) => { if (e.key === "Enter" || e.key === "Escape") { e.preventDefault(); onClose(); } }}>
                    OK
                </button>
            </div>
        </Modal>
    );
}

export const Dialogs = {
    Message: MessageBox,
    FilterMenu: real.FilterMenu ?? PlaceholderFilterMenu,
    SortDialog: real.SortDialog ?? PlaceholderSortDialog,
    FindReplace: real.FindReplace ?? PlaceholderFindReplace,
    GoTo: real.GoTo ?? PlaceholderGoTo,
    SaveProgress: real.SaveProgress ?? PlaceholderSaveProgress,
    ClosePrompt: real.ClosePrompt ?? PlaceholderClosePrompt,
    FormatCells: real.FormatCells,
};
