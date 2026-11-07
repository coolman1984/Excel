import * as React from "react";

// Ribbon icons (plan 2.5 says "@fluentui/react-icons (MIT), regular 20 px").
// DEVIATION: `@fluentui/react-icons` (v2.0.345, ~30,000 files) makes `pnpm` hang
// for minutes on the link/extract step of that single package on this PC (proven
// over two fresh installs). We use these tiny inline SVG icons instead so the build
// is fast and reliable. They are drawn at the same size and keep the same API, so
// swapping back to Fluent icons later is a one-file change.

export type IconProps = { className?: string };

const S = ({ children, className }: { children: React.ReactNode; className?: string }): React.ReactElement => (
    <svg
        className={className ?? "ribbonIcon"}
        width="20"
        height="20"
        viewBox="0 0 20 20"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.4"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden="true"
    >
        {children}
    </svg>
);

const Glyphs: Record<string, () => React.ReactElement> = {
    paste: () => (
        <S>
            <rect x="5" y="3.5" width="10" height="13" rx="1" />
            <rect x="7.5" y="2" width="5" height="3" rx="0.8" fill="currentColor" />
            <line x1="7.5" y1="9" x2="12.5" y2="9" />
            <line x1="7.5" y1="12" x2="12.5" y2="12" />
        </S>
    ),
    cut: () => (
        <S>
            <circle cx="5.5" cy="14.5" r="2" />
            <circle cx="14.5" cy="14.5" r="2" />
            <line x1="7" y1="13" x2="13" y2="3" />
            <line x1="13" y1="13" x2="7" y2="3" />
        </S>
    ),
    copy: () => (
        <S>
            <rect x="6" y="6" width="10" height="11" rx="1" />
            <path d="M4 13V4a1 1 0 0 1 1-1h8" />
        </S>
    ),
    bold: () => (
        <S>
            <path d="M6 3.5h5a2.5 2.5 0 0 1 0 5H6z" />
            <path d="M6 8.5h5.5a2.75 2.75 0 0 1 0 5.5H6z" />
        </S>
    ),
    italic: () => (
        <S>
            <line x1="8.5" y1="3.5" x2="14.5" y2="3.5" />
            <line x1="5.5" y1="16.5" x2="11.5" y2="16.5" />
            <line x1="11" y1="3.5" x2="9" y2="16.5" />
        </S>
    ),
    underline: () => (
        <S>
            <path d="M6 3.5v6a4 4 0 0 0 8 0v-6" />
            <line x1="5" y1="16.5" x2="15" y2="16.5" />
        </S>
    ),
    alignLeft: () => (
        <S>
            <line x1="4" y1="6" x2="16" y2="6" />
            <line x1="4" y1="10" x2="12" y2="10" />
            <line x1="4" y1="14" x2="16" y2="14" />
        </S>
    ),
    alignCentre: () => (
        <S>
            <line x1="4" y1="6" x2="16" y2="6" />
            <line x1="6" y1="10" x2="14" y2="10" />
            <line x1="4" y1="14" x2="16" y2="14" />
        </S>
    ),
    alignRight: () => (
        <S>
            <line x1="4" y1="6" x2="16" y2="6" />
            <line x1="8" y1="10" x2="16" y2="10" />
            <line x1="4" y1="14" x2="16" y2="14" />
        </S>
    ),
    fill: () => (
        <S>
            <path d="M4 10.5 9.5 5l5 5-5.5 5.5z" fill="currentColor" stroke="none" />
            <path d="M15 12c1 1.4 1.6 2.2 1.6 3a1.6 1.6 0 0 1-3.2 0c0-.8.6-1.6 1.6-3z" fill="currentColor" stroke="none" />
        </S>
    ),
    fontColor: () => (
        <S>
            <path d="M6 15 10 4l4 11" />
            <line x1="7.2" y1="12" x2="12.8" y2="12" />
            <rect x="4" y="17" width="12" height="2" fill="currentColor" stroke="none" />
        </S>
    ),
    borders: () => (
        <S>
            <rect x="4" y="4" width="12" height="12" rx="0.5" />
            <line x1="4" y1="10" x2="16" y2="10" />
            <line x1="10" y1="4" x2="10" y2="16" />
        </S>
    ),
    sort: () => (
        <S>
            <line x1="7" y1="4" x2="7" y2="16" />
            <path d="M4.5 6.5 7 4l2.5 2.5" />
            <line x1="13" y1="4" x2="13" y2="16" />
            <path d="M10.5 13.5 13 16l2.5-2.5" />
        </S>
    ),
    filter: () => (
        <S>
            <path d="M3.5 5h13l-5 6v5l-3-1.5V11z" />
        </S>
    ),
    zoomIn: () => (
        <S>
            <circle cx="9" cy="9" r="5" />
            <line x1="12.8" y1="12.8" x2="17" y2="17" />
            <line x1="6.8" y1="9" x2="11.2" y2="9" />
            <line x1="9" y1="6.8" x2="9" y2="11.2" />
        </S>
    ),
    zoomOut: () => (
        <S>
            <circle cx="9" cy="9" r="5" />
            <line x1="12.8" y1="12.8" x2="17" y2="17" />
            <line x1="6.8" y1="9" x2="11.2" y2="9" />
        </S>
    ),
    document: () => (
        <S>
            <path d="M5 3h7l3 3v11H5z" />
            <path d="M12 3v3h3" />
        </S>
    ),
    save: () => (
        <S>
            <path d="M4 4h10l2 2v10H4z" />
            <rect x="7" y="4" width="6" height="4" />
            <rect x="6.5" y="11" width="7" height="5" />
        </S>
    ),
    search: () => (
        <S>
            <circle cx="9" cy="9" r="5" />
            <line x1="12.8" y1="12.8" x2="17" y2="17" />
        </S>
    ),
    undo: () => (
        <S>
            <path d="M7 7H13a4 4 0 0 1 0 8H8" />
            <path d="M9.5 4.5 7 7l2.5 2.5" />
        </S>
    ),
    redo: () => (
        <S>
            <path d="M13 7H7a4 4 0 0 0 0 8h5" />
            <path d="M10.5 4.5 13 7l-2.5 2.5" />
        </S>
    ),
};

export const Icons = Glyphs;

/** Render a named icon at 20 px. */
export function Icon({ name }: { name: keyof typeof Icons }): React.ReactElement {
    const G = Icons[name];
    if (!G) return <span className="iconFallback" />;
    return <G />;
}

