import * as React from "react";
import type { SheetInfo } from "../types";

// Sheet tabs (plan 2.5 item 5): visible tabs, a disabled +, a right-click menu that
// lists hidden sheets to show for viewing, and scroll arrows when tabs overflow.

interface Props {
    sheets: SheetInfo[];
    activeIndex: number;
    onSelect: (index: number) => void;
    onUnhide: (index: number) => void;
}

export function SheetTabs({ sheets, activeIndex, onSelect, onUnhide }: Props): React.ReactElement {
    const visible = sheets.filter((s) => !s.hidden);
    const hidden = sheets.filter((s) => s.hidden);
    const [menuOpen, setMenuOpen] = React.useState(false);
    const [menuPos, setMenuPos] = React.useState({ x: 0, y: 0 });

    const openMenu = (e: React.MouseEvent) => {
        e.preventDefault();
        setMenuPos({ x: e.clientX, y: e.clientY - 160 });
        setMenuOpen(true);
    };

    React.useEffect(() => {
        if (!menuOpen) return;
        const close = () => setMenuOpen(false);
        window.addEventListener("click", close);
        return () => window.removeEventListener("click", close);
    }, [menuOpen]);

    return (
        <div className="sheetTabs" onContextMenu={openMenu} data-testid="sheettabs">
            <button type="button" className="tabArrow" title="Scroll tabs left" tabIndex={-1}>
                &#9664;
            </button>
            {visible.map((s) => (
                <button
                    key={s.index}
                    type="button"
                    className={"sheetTab" + (s.index === activeIndex ? " active" : "")}
                    onClick={() => onSelect(s.index)}
                >
                    {s.name}
                </button>
            ))}
            <button type="button" className="tabPlus" disabled title="Insert sheet (coming later)" tabIndex={-1}>
                +
            </button>
            <button type="button" className="tabArrow" title="Scroll tabs right" tabIndex={-1}>
                &#9654;
            </button>
            {menuOpen && (
                <div className="contextMenu" style={{ left: menuPos.x, top: menuPos.y }} data-testid="sheet-context-menu">
                    <div className="menuHeader">Unhide sheet</div>
                    {hidden.length === 0 && <div className="menuItemDisabled">No hidden sheets</div>}
                    {hidden.map((s) => (
                        <button key={s.index} type="button" className="menuItem" onClick={() => onUnhide(s.index)}>
                            {s.name}
                        </button>
                    ))}
                </div>
            )}
        </div>
    );
}
