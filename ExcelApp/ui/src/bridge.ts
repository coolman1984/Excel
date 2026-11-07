import { MockHost } from "./mockHost";

// bridge (contract): call(cmd, args) -> Promise and on(event, handler) -> unsubscribe.
// With window.chrome.webview present this talks to the C# host; without it (a plain
// browser during `pnpm dev` / headless Chrome) a mock host answers every command.

interface Pending {
    resolve: (v: any) => void;
    reject: (e: any) => void;
}

interface WebViewLike {
    postMessage(msg: any): void;
    addEventListener(type: string, cb: (e: any) => void): void;
}

const g = typeof window !== "undefined" ? (window as any) : undefined;
const webview: WebViewLike | undefined = g?.chrome?.webview;
export const IS_MOCK = !webview;

let idCounter = 0;
const pending = new Map<number, Pending>();
const listeners = new Map<string, Set<(e: any) => void>>();
let mock: MockHost | null = null;

function getMock(): MockHost {
    if (!mock) mock = new MockHost();
    return mock;
}

function emit(event: string, payload: any): void {
    const set = listeners.get(event);
    if (set) for (const h of set) h(payload);
}

if (webview) {
    webview.addEventListener("message", (e: any) => {
        const msg = e?.data;
        if (!msg) return;
        if (typeof msg.id === "number") {
            const p = pending.get(msg.id);
            if (p) {
                pending.delete(msg.id);
                if (msg.ok) p.resolve(msg.result);
                else p.reject(new Error(msg.error || "The host reported an error."));
            }
        } else if (typeof msg.event === "string") {
            emit(msg.event, msg);
        }
    });
}

/** True when there is no C# host and the mock answers commands. */
export function isMock(): boolean {
    return IS_MOCK;
}

/** The mock host instance (only when there is no C# host); used by the self test. */
export function mockHost(): MockHost | null {
    return IS_MOCK ? getMock() : null;
}

export function call<T = any>(cmd: string, args: object = {}): Promise<T> {
    if (!webview) return getMock().call(cmd, args) as Promise<T>;
    const id = ++idCounter;
    return new Promise<T>((resolve, reject) => {
        pending.set(id, { resolve, reject });
        webview.postMessage({ id, cmd, ...args });
    });
}

export function on(event: string, handler: (e: any) => void): () => void {
    if (!webview) return getMock().on(event, handler);
    let set = listeners.get(event);
    if (!set) {
        set = new Set();
        listeners.set(event, set);
    }
    set.add(handler);
    return () => set!.delete(handler);
}
