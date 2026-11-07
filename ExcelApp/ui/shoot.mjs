// Headless-Chrome screenshot driver for the mock-host UI (lane A self check).
// Starts nothing itself: it expects `pnpm exec vite` and Chrome (with
// --remote-debugging-port=9222) already running, drives a local URL only, and
// writes PNGs. Uses Node's global WebSocket + fetch (no extra packages).
import { setTimeout as sleep } from "node:timers/promises";
import fs from "node:fs";

const PORT = 9222;
const url = process.argv[2] ?? "http://127.0.0.1:5173/";
const outDir = process.argv[3] ?? ".shots";

async function findTarget() {
    const hosts = ["[::1]", "127.0.0.1", "localhost"];
    for (let i = 0; i < 60; i++) {
        for (const h of hosts) {
            try {
                const r = await fetch(`http://${h}:${PORT}/json/list`);
                const list = await r.json();
                const page = list.find((t) => t.type === "page");
                if (page?.webSocketDebuggerUrl) {
                    // Rewrite ws host to match the reachable one.
                    return page.webSocketDebuggerUrl.replace("localhost", h).replace("127.0.0.1", h);
                }
            } catch {
                /* try next host */
            }
        }
        await sleep(250);
    }
    throw new Error("no Chrome page target on port " + PORT);
}

const ws = new WebSocket(await findTarget());
let id = 0;
const pending = new Map();
ws.addEventListener("message", (e) => {
    const m = JSON.parse(e.data);
    if (m.id && pending.has(m.id)) {
        pending.get(m.id)(m);
        pending.delete(m.id);
    }
});
const send = (method, params = {}) =>
    new Promise((resolve) => {
        const mid = ++id;
        pending.set(mid, resolve);
        ws.send(JSON.stringify({ id: mid, method, params }));
    });
await new Promise((r) => ws.addEventListener("open", r, { once: true }));

await send("Page.enable");
await send("Runtime.enable");
await send("Emulation.setDeviceMetricsOverride", { width: 1600, height: 900, deviceScaleFactor: 1, mobile: false });
await send("Page.navigate", { url });

// Wait until the app reports the ready state.
let state = "none";
for (let i = 0; i < 80; i++) {
    const r = await send("Runtime.evaluate", {
        expression: "(window.__app && window.__app.state) ? window.__app.state() : 'none'",
        returnByValue: true,
    });
    state = r.result?.result?.value ?? "none";
    if (state === "ready") break;
    await sleep(250);
}
await sleep(2000); // let the first blocks load and paint

async function shot(name) {
    const r = await send("Page.captureScreenshot", { format: "png" });
    fs.writeFileSync(`${outDir}/${name}`, Buffer.from(r.result.data, "base64"));
    console.log("wrote", name);
}

// Shot 1: the grid. Scroll a bit and select a range with the self-test hook.
await send("Runtime.evaluate", { expression: "window.__app.run('select', 3, 2, 8, 5)" });
await send("Runtime.evaluate", { expression: "window.__app.scrollTo(4)" });
await sleep(900);
await shot("mock-grid.png");

// Shot 2: the ribbon with the Fill colour palette open (Excel look).
await send("Runtime.evaluate", {
    expression: "document.querySelector('[data-testid=\"menu-fill\"]').click()",
});
await sleep(700);
await shot("mock-ribbon.png");

console.log("state was:", state);
ws.close();
process.exit(0);
