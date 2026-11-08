import * as React from "react";
import { Icon } from "../icons";
import { THEME_ROWS, STANDARD } from "../palette";
import { NUMBER_FORMATS } from "../numberFormat";
import type { CellStyle } from "../types";

export type TabName = "File" | "Home" | "Data" | "View";

export interface RibbonProps {
    activeTab: TabName;
    onTab: (t: TabName) => void;
    collapsed: boolean;
    onToggleCollapse: () => void;
    enabled: boolean;

    onPaste: () => void;
    onCut: () => void;
    onCopy: () => void;

    fontName: string;
    fontSize: string;
    onBold: () => void;
    onItalic: () => void;
    onUnderline: () => void;
    onStyle: (style: CellStyle) => void;
    onBorders: (kind: "none" | "all" | "outside" | "bottom") => void;

    onAlign: (a: 0 | 1 | 2) => void;
    onWrap: () => void;

    onNumberFormat: (code: string) => void;
    onPercent: () => void;
    onComma: () => void;
    onIncrease: () => void;
    onDecrease: () => void;

    onAutoSum: () => void;
    onClearAll: () => void;
    onClearContents: () => void;
    onClearFormats: () => void;

    onSortAZ: () => void;
    onSortZA: () => void;
    onSortDialog: () => void;
    onFilterToggle: () => void;
    onClearFilter: () => void;
    onReapply: () => void;
    filterActive: boolean;

    onFind: () => void;
    onReplace: () => void;
    onGoTo: () => void;

    gridlines: boolean;
    headings: boolean;
    formulaBar: boolean;
    onToggleGridlines: () => void;
    onToggleHeadings: () => void;
    onToggleFormulaBar: () => void;
    onZoomDialog: () => void;
    onZoom100: () => void;
    onFreezePanes: () => void;
    onFreezeFirstColumn: () => void;
    onFreezeTopRow: () => void;
    onUnfreeze: () => void;
}

function RButton({
    label,
    icon,
    onClick,
    big,
    disabled,
    title,
    active,
    testid,
}: {
    label: string;
    icon?: string;
    onClick: () => void;
    big?: boolean;
    disabled?: boolean;
    title?: string;
    active?: boolean;
    testid?: string;
}): React.ReactElement {
    return (
        <button
            type="button"
            className={"rb" + (big ? " big" : "") + (active ? " active" : "")}
            onClick={onClick}
            disabled={disabled}
            title={title ?? label}
            data-testid={testid}
        >
            {icon && <Icon name={icon as never} />}
            <span className="rbLabel">{label}</span>
        </button>
    );
}

function Group({ label, children }: { label: string; children: React.ReactNode }): React.ReactElement {
    return (
        <div className="rgroup">
            <div className="rgroupBody">{children}</div>
            <div className="rgroupLabel">{label}</div>
        </div>
    );
}

function Menu({
    label,
    icon,
    big,
    children,
    testid,
}: {
    label: string;
    icon?: string;
    big?: boolean;
    children: (close: () => void) => React.ReactNode;
    testid?: string;
}): React.ReactElement {
    const [open, setOpen] = React.useState(false);
    const ref = React.useRef<HTMLDivElement | null>(null);
    React.useEffect(() => {
        if (!open) return;
        const h = (e: MouseEvent) => {
            if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
        };
        window.addEventListener("mousedown", h);
        return () => window.removeEventListener("mousedown", h);
    }, [open]);
    return (
        <div className="menuWrap" ref={ref}>
            <button type="button" className={"rb" + (big ? " big" : "")} onClick={() => setOpen((o) => !o)} data-testid={testid}>
                {icon && <Icon name={icon as never} />}
                <span className="rbLabel">{label}</span>
            </button>
            {open && (
                <div className="dropdown" role="menu">
                    {children(() => setOpen(false))}
                </div>
            )}
        </div>
    );
}

function PaletteGrid({ onPick, auto }: { onPick: (hex: string) => void; auto?: boolean }): React.ReactElement {
    return (
        <div className="palette">
            {auto && (
                <button type="button" className="paletteAuto" onClick={() => onPick("")}>
                    Automatic
                </button>
            )}
            {THEME_ROWS.map((row) => (
                <div className="paletteRow" key={row.name}>
                    <button type="button" className="swatch swatchBase" title={row.name} style={{ background: row.base }} onClick={() => onPick(row.base)} />
                    {row.tints.map((t, i) => (
                        <button key={i} type="button" className="swatch" style={{ background: t }} onClick={() => onPick(t)} />
                    ))}
                </div>
            ))}
            <div className="paletteRow">
                {STANDARD.map((s) => (
                    <button key={s.hex} type="button" className="swatch" title={s.name} style={{ background: s.hex }} onClick={() => onPick(s.hex)} />
                ))}
            </div>
        </div>
    );
}

export function Ribbon(props: RibbonProps): React.ReactElement {
    const { activeTab, onTab, collapsed, onToggleCollapse } = props;
    return (
        <div className="ribbon" data-testid="ribbon">
            <div className="tabStrip">
                {(["File", "Home", "Data", "View"] as TabName[]).map((t) => (
                    <button key={t} type="button" className={"tab" + (activeTab === t ? " active" : "")} onClick={() => onTab(t)} data-testid={"tab-" + t}>
                        {t}
                    </button>
                ))}
                <span className="tabStripSpacer" />
                <button type="button" className="collapseBtn" onClick={onToggleCollapse} title={collapsed ? "Show the ribbon" : "Hide the ribbon"} data-testid="ribbon-collapse">
                    {collapsed ? "\u25BE" : "\u25B4"}
                </button>
            </div>
            {!collapsed && activeTab !== "File" && (
                <div className="ribbonBody" data-testid="ribbon-body">
                    {activeTab === "Home" && <HomeTab {...props} />}
                    {activeTab === "Data" && <DataTab {...props} />}
                    {activeTab === "View" && <ViewTab {...props} />}
                </div>
            )}
        </div>
    );
}

function HomeTab(p: RibbonProps): React.ReactElement {
    const d = !p.enabled;
    return (
        <>
            <Group label="Clipboard">
                <RButton label="Paste" icon="paste" big disabled={d} onClick={p.onPaste} />
                <div className="rcol">
                    <RButton label="Cut" icon="cut" disabled={d} onClick={p.onCut} />
                    <RButton label="Copy" icon="copy" disabled={d} onClick={p.onCopy} />
                </div>
            </Group>

            <Group label="Font">
                <div className="rcol">
                    <input className="ribbonInput fontName" value={p.fontName} readOnly title="Sheet font" />
                    <div className="rrow">
                        <RButton label="B" title="Bold (Ctrl+B)" disabled={d} onClick={p.onBold} testid="btn-bold" />
                        <RButton label="I" title="Italic (Ctrl+I)" disabled={d} onClick={p.onItalic} />
                        <RButton label="U" title="Underline (Ctrl+U)" disabled={d} onClick={p.onUnderline} />
                        <Menu label="Borders" icon="borders" testid="menu-borders">
                            {(close) => (
                                <div className="menuList">
                                    {(["bottom", "all", "outside", "none"] as const).map((k) => (
                                        <button key={k} type="button" className="menuItem" onClick={() => { p.onBorders(k); close(); }}>
                                            {k === "bottom" ? "Bottom" : k === "all" ? "All Borders" : k === "outside" ? "Outside Borders" : "No Border"}
                                        </button>
                                    ))}
                                </div>
                            )}
                        </Menu>
                        <Menu label="Fill" icon="fill" testid="menu-fill">
                            {(close) => (
                                <PaletteGrid
                                    auto
                                    onPick={(hex) => {
                                        p.onStyle({ fill: hex === "" ? undefined : hex });
                                        close();
                                    }}
                                />
                            )}
                        </Menu>
                        <Menu label="Font colour" icon="fontColor" testid="menu-fontcolor">
                            {(close) => (
                                <PaletteGrid
                                    auto
                                    onPick={(hex) => {
                                        p.onStyle({ fontColor: hex === "" ? undefined : hex });
                                        close();
                                    }}
                                />
                            )}
                        </Menu>
                    </div>
                </div>
            </Group>

            <Group label="Alignment">
                <div className="rcol">
                    <div className="rrow">
                        <RButton label="Left" icon="alignLeft" title="Align left" disabled={d} onClick={() => p.onAlign(0)} testid="align-left" />
                        <RButton label="Centre" icon="alignCentre" title="Centre" disabled={d} onClick={() => p.onAlign(2)} />
                        <RButton label="Right" icon="alignRight" title="Align right" disabled={d} onClick={() => p.onAlign(1)} />
                    </div>
                    <div className="rrow">
                        <RButton label="Top" disabled title="Vertical alignment - coming later" onClick={p.onWrap} />
                        <RButton label="Middle" disabled title="Vertical alignment - coming later" onClick={p.onWrap} />
                        <RButton label="Bottom" disabled title="Vertical alignment - coming later" onClick={p.onWrap} />
                        <RButton label="Wrap" title="Wrap text - coming later" disabled onClick={p.onWrap} />
                    </div>
                </div>
            </Group>

            <Group label="Number">
                <div className="rcol">
                    <select
                        className="ribbonSelect"
                        disabled={d}
                        defaultValue=""
                        onChange={(e) => {
                            if (e.target.value) p.onNumberFormat(e.target.value);
                        }}
                        data-testid="number-format"
                    >
                        <option value="">Number format...</option>
                        {NUMBER_FORMATS.map((f) => (
                            <option key={f.label} value={f.code}>
                                {f.label}
                            </option>
                        ))}
                    </select>
                    <div className="rrow">
                        <RButton label="$" title="Currency" disabled={d} onClick={() => p.onNumberFormat("[$$-409]#,##0.00")} />
                        <RButton label="%" title="Percent" disabled={d} onClick={p.onPercent} />
                        <RButton label="," title="Comma style" disabled={d} onClick={p.onComma} />
                        <RButton label=".00+" title="Increase decimal" disabled={d} onClick={p.onIncrease} testid="inc-decimals" />
                        <RButton label=".0-" title="Decrease decimal" disabled={d} onClick={p.onDecrease} testid="dec-decimals" />
                    </div>
                </div>
            </Group>

            <Group label="Cells">
                <RButton label="Insert" disabled title="Insert cells - coming later" onClick={() => {}} />
                <RButton label="Delete" disabled title="Delete cells - coming later" onClick={() => {}} />
                <RButton label="Format" disabled title="Cell format - coming later" onClick={() => {}} />
            </Group>

            <Group label="Editing">
                <div className="rcol">
                    <RButton label="AutoSum" disabled={d} onClick={p.onAutoSum} testid="btn-autosum" />
                    <Menu label="Clear" testid="menu-clear">
                        {(close) => (
                            <div className="menuList">
                                <button type="button" className="menuItem" onClick={() => { p.onClearAll(); close(); }}>
                                    Clear All
                                </button>
                                <button type="button" className="menuItem" onClick={() => { p.onClearContents(); close(); }}>
                                    Clear Contents
                                </button>
                                <button type="button" className="menuItem" onClick={() => { p.onClearFormats(); close(); }}>
                                    Clear Formats
                                </button>
                            </div>
                        )}
                    </Menu>
                </div>
                <div className="rcol">
                    <Menu label="Sort & Filter" testid="menu-sortfilter">
                        {(close) => (
                            <div className="menuList">
                                <button type="button" className="menuItem" onClick={() => { p.onSortAZ(); close(); }}>
                                    Sort A to Z
                                </button>
                                <button type="button" className="menuItem" onClick={() => { p.onSortZA(); close(); }}>
                                    Sort Z to A
                                </button>
                                <button type="button" className="menuItem" onClick={() => { p.onSortDialog(); close(); }}>
                                    Custom Sort...
                                </button>
                                <button type="button" className="menuItem" onClick={() => { p.onFilterToggle(); close(); }}>
                                    {p.filterActive ? "Turn Filter Off" : "Filter"}
                                </button>
                            </div>
                        )}
                    </Menu>
                    <Menu label="Find & Select" icon="search" testid="menu-findselect">
                        {(close) => (
                            <div className="menuList">
                                <button type="button" className="menuItem" onClick={() => { p.onFind(); close(); }}>
                                    Find...
                                </button>
                                <button type="button" className="menuItem" onClick={() => { p.onReplace(); close(); }}>
                                    Replace...
                                </button>
                                <button type="button" className="menuItem" onClick={() => { p.onGoTo(); close(); }}>
                                    Go To...
                                </button>
                            </div>
                        )}
                    </Menu>
                </div>
            </Group>
        </>
    );
}

function DataTab(p: RibbonProps): React.ReactElement {
    const d = !p.enabled;
    return (
        <>
            <Group label="Sort & Filter">
                <div className="rcol">
                    <div className="rrow">
                        <RButton label="A to Z" icon="sort" disabled={d} onClick={p.onSortAZ} />
                        <RButton label="Z to A" icon="sort" disabled={d} onClick={p.onSortZA} />
                    </div>
                    <div className="rrow">
                        <RButton label="Sort" disabled={d} onClick={p.onSortDialog} testid="btn-sort" />
                        <RButton label="Filter" icon="filter" disabled={d} active={p.filterActive} onClick={p.onFilterToggle} testid="btn-filter" />
                        <RButton label="Clear" disabled={d} onClick={p.onClearFilter} />
                        <RButton label="Reapply" disabled={d} onClick={p.onReapply} />
                    </div>
                </div>
            </Group>
            <Group label="Data Tools">
                <RButton label="Remove Duplicates" disabled title="Remove Duplicates - coming later" onClick={() => {}} />
                <RButton label="Text to Columns" disabled title="Text to Columns - coming later" onClick={() => {}} />
            </Group>
        </>
    );
}

function ViewTab(p: RibbonProps): React.ReactElement {
    return (
        <>
            <Group label="Show">
                <div className="rcol">
                    <label className="checkRow">
                        <input type="checkbox" checked={p.gridlines} onChange={p.onToggleGridlines} data-testid="chk-gridlines" /> Gridlines
                    </label>
                    <label className="checkRow">
                        <input type="checkbox" checked={p.headings} onChange={p.onToggleHeadings} data-testid="chk-headings" /> Headings
                    </label>
                    <label className="checkRow">
                        <input type="checkbox" checked={p.formulaBar} onChange={p.onToggleFormulaBar} data-testid="chk-formulabar" /> Formula Bar
                    </label>
                </div>
            </Group>
            <Group label="Zoom">
                <RButton label="Zoom" icon="zoomIn" onClick={p.onZoomDialog} testid="btn-zoomdialog" />
                <RButton label="100%" onClick={p.onZoom100} testid="btn-zoom100" />
                <RButton label="Zoom to Selection" disabled title="Zoom to Selection - coming later" onClick={() => {}} />
            </Group>
            <Group label="Window">
                <Menu label="Freeze Panes" testid="menu-freeze">
                    {(close) => (
                        <div className="menuList">
                            <button type="button" className="menuItem" onClick={() => { p.onFreezePanes(); close(); }}>
                                Freeze Panes
                            </button>
                            <button type="button" className="menuItem" onClick={() => { p.onFreezeTopRow(); close(); }}>
                                Freeze Top Row
                            </button>
                            <button type="button" className="menuItem" onClick={() => { p.onFreezeFirstColumn(); close(); }}>
                                Freeze First Column
                            </button>
                            <button type="button" className="menuItem" onClick={() => { p.onUnfreeze(); close(); }}>
                                Unfreeze
                            </button>
                        </div>
                    )}
                </Menu>
            </Group>
        </>
    );
}





