// Full Excel keyboard list (plan 2.5). Every binding maps to a command id.
// The grid's single keydown listener calls resolveCommand and dispatches; the
// navigation keys are also mapped here so the table is complete and testable.

export interface KeyEventLike {
    key: string;
    ctrlKey?: boolean;
    shiftKey?: boolean;
    altKey?: boolean;
    metaKey?: boolean;
}

export interface Binding {
    command: string;
    /** Normalised key. Letters are lowercase; named keys keep their DOM name. */
    key: string;
    ctrl?: boolean;
    shift?: boolean;
    alt?: boolean;
    label: string;
}

// Exactly one ordered list, most specific first: the resolver returns the first
// match, so modifier combinations must come before their un-modified forms.
export const BINDINGS: Binding[] = [
    // Ctrl + Shift
    { command: "extendEdgeUp", key: "ArrowUp", ctrl: true, shift: true, label: "Ctrl+Shift+Up" },
    { command: "extendEdgeDown", key: "ArrowDown", ctrl: true, shift: true, label: "Ctrl+Shift+Down" },
    { command: "extendEdgeLeft", key: "ArrowLeft", ctrl: true, shift: true, label: "Ctrl+Shift+Left" },
    { command: "extendEdgeRight", key: "ArrowRight", ctrl: true, shift: true, label: "Ctrl+Shift+Right" },
    { command: "filterToggle", key: "l", ctrl: true, shift: true, label: "Ctrl+Shift+L" },

    // Ctrl + arrows (jump to edge of data)
    { command: "edgeUp", key: "ArrowUp", ctrl: true, label: "Ctrl+Up" },
    { command: "edgeDown", key: "ArrowDown", ctrl: true, label: "Ctrl+Down" },
    { command: "edgeLeft", key: "ArrowLeft", ctrl: true, label: "Ctrl+Left" },
    { command: "edgeRight", key: "ArrowRight", ctrl: true, label: "Ctrl+Right" },

    // Shift + Enter / Shift + Tab (the negative of Enter / Tab)
    { command: "moveUp", key: "Enter", shift: true, label: "Shift+Enter" },
    { command: "moveLeft", key: "Tab", shift: true, label: "Shift+Tab" },

    // Ctrl + letters and keys
    { command: "selectAll", key: "a", ctrl: true, label: "Ctrl+A" },
    { command: "copy", key: "c", ctrl: true, label: "Ctrl+C" },
    { command: "find", key: "f", ctrl: true, label: "Ctrl+F" },
    { command: "goTo", key: "g", ctrl: true, label: "Ctrl+G" },
    { command: "replace", key: "h", ctrl: true, label: "Ctrl+H" },
    { command: "italic", key: "i", ctrl: true, label: "Ctrl+I" },
    { command: "open", key: "o", ctrl: true, label: "Ctrl+O" },
    { command: "save", key: "s", ctrl: true, label: "Ctrl+S" },
    { command: "underline", key: "u", ctrl: true, label: "Ctrl+U" },
    { command: "paste", key: "v", ctrl: true, label: "Ctrl+V" },
    { command: "cut", key: "x", ctrl: true, label: "Ctrl+X" },
    { command: "redo", key: "y", ctrl: true, label: "Ctrl+Y" },
    { command: "undo", key: "z", ctrl: true, label: "Ctrl+Z" },
    { command: "bold", key: "b", ctrl: true, label: "Ctrl+B" },
    { command: "firstCell", key: "Home", ctrl: true, label: "Ctrl+Home" },
    { command: "lastCell", key: "End", ctrl: true, label: "Ctrl+End" },
    { command: "prevSheet", key: "PageUp", ctrl: true, label: "Ctrl+Page Up" },
    { command: "nextSheet", key: "PageDown", ctrl: true, label: "Ctrl+Page Down" },

    // Alt + Down opens a header filter menu / dropdown.
    { command: "openFilter", key: "ArrowDown", alt: true, label: "Alt+Down" },

    // Shift + arrows extend the selection.
    { command: "extendUp", key: "ArrowUp", shift: true, label: "Shift+Up" },
    { command: "extendDown", key: "ArrowDown", shift: true, label: "Shift+Down" },
    { command: "extendLeft", key: "ArrowLeft", shift: true, label: "Shift+Left" },
    { command: "extendRight", key: "ArrowRight", shift: true, label: "Shift+Right" },

    // Plain movement and editing
    { command: "moveUp", key: "ArrowUp", label: "Up" },
    { command: "moveDown", key: "ArrowDown", label: "Down" },
    { command: "moveLeft", key: "ArrowLeft", label: "Left" },
    { command: "moveRight", key: "ArrowRight", label: "Right" },
    { command: "pageUp", key: "PageUp", label: "Page Up" },
    { command: "pageDown", key: "PageDown", label: "Page Down" },
    { command: "firstCol", key: "Home", label: "Home" },
    { command: "lastCol", key: "End", label: "End" },
    { command: "moveDown", key: "Enter", label: "Enter" },
    { command: "moveRight", key: "Tab", label: "Tab" },
    { command: "edit", key: "F2", label: "F2" },
    { command: "clear", key: "Delete", label: "Delete" },
    { command: "cancel", key: "Escape", label: "Esc" },
];

function matches(b: Binding, e: KeyEventLike): boolean {
    const raw = e.key;
    const key = raw.length === 1 ? raw.toLowerCase() : raw;
    if (key !== b.key) return false;
    const ctrl = !!(e.ctrlKey || e.metaKey);
    if ((b.ctrl ?? false) !== ctrl) return false;
    if ((b.shift ?? false) !== !!e.shiftKey) return false;
    if ((b.alt ?? false) !== !!e.altKey) return false;
    return true;
}

/** The command a keydown maps to, or null. */
export function resolveCommand(e: KeyEventLike): string | null {
    for (const b of BINDINGS) {
        if (matches(b, e)) return b.command;
    }
    return null;
}

/** Ctrl + mouse wheel zooming (plan 2.5). */
export function wheelZoom(e: { ctrlKey?: boolean; deltaY: number }): "zoomIn" | "zoomOut" | null {
    if (!e.ctrlKey) return null;
    return e.deltaY < 0 ? "zoomIn" : "zoomOut";
}

/** Build a synthetic keydown for a binding, used by the tests. */
export function eventFor(b: Binding): KeyEventLike {
    return { key: b.key, ctrlKey: b.ctrl, shiftKey: b.shift, altKey: b.alt };
}

