import * as React from "react";

// Formula bar (plan 2.5 item 3): the name box (C12 or C12:F40) and the formula box
// showing raw text. Enter commits, Esc cancels.

interface Props {
    address: string;
    raw: string;
    onNameSubmit: (address: string) => void;
    onRawCommit: (text: string) => void;
    visible: boolean;
}

export function FormulaBar({ address, raw, onNameSubmit, onRawCommit, visible }: Props): React.ReactElement | null {
    const [name, setName] = React.useState(address);
    const [text, setText] = React.useState(raw);
    const [editing, setEditing] = React.useState(false);

    React.useEffect(() => setName(address), [address]);
    React.useEffect(() => {
        if (!editing) setText(raw);
    }, [raw, editing]);

    if (!visible) return null;

    return (
        <div className="formulaBar" data-testid="formulabar">
            <input
                className="nameBox"
                value={name}
                aria-label="Name box"
                onChange={(e) => setName(e.target.value)}
                onKeyDown={(e) => {
                    if (e.key === "Enter") {
                        onNameSubmit(name);
                        (e.target as HTMLInputElement).blur();
                    }
                }}
            />
            <span className="fxLabel">fx</span>
            <input
                className="formulaBox"
                value={text}
                aria-label="Formula box"
                onChange={(e) => {
                    setEditing(true);
                    setText(e.target.value);
                }}
                onKeyDown={(e) => {
                    if (e.key === "Enter") {
                        onRawCommit(text);
                        setEditing(false);
                        (e.target as HTMLInputElement).blur();
                    } else if (e.key === "Escape") {
                        setEditing(false);
                        setText(raw);
                        (e.target as HTMLInputElement).blur();
                    }
                }}
                onBlur={() => setEditing(false)}
            />
        </div>
    );
}
