import * as React from "react";

// File backstage (plan 2.5 item 2, File): a full-window page with a green left rail.

export interface FileInfo {
    name: string;
    path: string;
    sizeBytes: number;
    lastRead: string;
    cacheBytes: number;
    nasaProtected: boolean;
}

interface Props {
    info: FileInfo | null;
    recent: string[];
    onOpenBrowse: () => void;
    onOpenPath: (path: string) => void;
    onSave: () => void;
    onSaveAs: () => void;
    onClose: () => void;
    onBack: () => void;
}

type Section = "home" | "open" | "info";

function human(bytes: number): string {
    if (bytes <= 0) return "-";
    const units = ["B", "KB", "MB", "GB"];
    let n = bytes;
    let u = 0;
    while (n >= 1024 && u < units.length - 1) {
        n /= 1024;
        u++;
    }
    return n.toFixed(n < 10 && u > 0 ? 1 : 0) + " " + units[u];
}

export function FileBackstage({ info, recent, onOpenBrowse, onOpenPath, onSave, onSaveAs, onClose, onBack }: Props): React.ReactElement {
    const [section, setSection] = React.useState<Section>("home");
    return (
        <div className="backstage" data-testid="backstage">
            <div className="bsRail">
                <button type="button" className="bsBack" onClick={onBack} data-testid="bs-back">
                    &#8592; Back
                </button>
                <button type="button" className={"bsItem" + (section === "home" ? " active" : "")} onClick={() => setSection("home")}>
                    Home
                </button>
                <button type="button" className={"bsItem" + (section === "open" ? " active" : "")} onClick={() => setSection("open")}>
                    Open
                </button>
                <button type="button" className="bsItem" onClick={onSave} data-testid="bs-save">
                    Save
                </button>
                <button type="button" className="bsItem" onClick={onSaveAs} data-testid="bs-saveas">
                    Save As
                </button>
                <button type="button" className={"bsItem" + (section === "info" ? " active" : "")} onClick={() => setSection("info")}>
                    Info
                </button>
                <button type="button" className="bsItem" onClick={onClose} data-testid="bs-close">
                    Close
                </button>
            </div>
            <div className="bsBody">
                <h1 className="bsTitle">{section === "info" ? "Info" : section === "open" ? "Open" : "Home"}</h1>
                {section === "home" && (
                    <>
                        <h2 className="bsSub">Recent</h2>
                        <ul className="recentList" data-testid="recent-list">
                            {recent.length === 0 && <li className="recentEmpty">No recent files</li>}
                            {recent.map((p) => (
                                <li key={p}>
                                    <button type="button" className="recentItem" onClick={() => onOpenPath(p)}>
                                        {p}
                                    </button>
                                </li>
                            ))}
                        </ul>
                        <button type="button" className="bsPrimary" onClick={onOpenBrowse} data-testid="bs-browse">
                            Browse...
                        </button>
                    </>
                )}
                {section === "open" && (
                    <>
                        <button type="button" className="bsPrimary" onClick={onOpenBrowse} data-testid="bs-browse2">
                            Browse...
                        </button>
                        <ul className="recentList">
                            {recent.map((p) => (
                                <li key={p}>
                                    <button type="button" className="recentItem" onClick={() => onOpenPath(p)}>
                                        {p}
                                    </button>
                                </li>
                            ))}
                        </ul>
                    </>
                )}
                {section === "info" && (
                    <table className="infoTable" data-testid="info-table">
                        <tbody>
                            <tr>
                                <td>Name</td>
                                <td>{info?.name ?? "-"}</td>
                            </tr>
                            <tr>
                                <td>Path</td>
                                <td>{info?.path ?? "-"}</td>
                            </tr>
                            <tr>
                                <td>Size</td>
                                <td>{info ? human(info.sizeBytes) : "-"}</td>
                            </tr>
                            <tr>
                                <td>Last read</td>
                                <td>{info?.lastRead ?? "-"}</td>
                            </tr>
                            <tr>
                                <td>Cache size</td>
                                <td>{info ? human(info.cacheBytes) : "-"}</td>
                            </tr>
                            <tr>
                                <td>Protected by NASCA (from the first 8 bytes)</td>
                                <td>{info ? (info.nasaProtected ? "yes" : "no") : "-"}</td>
                            </tr>
                        </tbody>
                    </table>
                )}
            </div>
        </div>
    );
}
