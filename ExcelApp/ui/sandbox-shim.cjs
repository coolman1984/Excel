// Test-only preload for running vitest inside the Codex sandbox, where Node may not start cmd.exe (spawn EPERM).
// Vite probes "net use" once at startup; a synchronous EPERM would abort it. Here the error goes to the callback,
// and Vite falls back to plain realpath. Use: NODE_OPTIONS=--require=./sandbox-shim.cjs (see package script test:sandbox).
const cp = require("node:child_process");
for (const name of ["exec", "execFile"]) {
  const original = cp[name];
  cp[name] = function (...args) {
    try { return original.apply(this, args); }
    catch (error) {
      const callback = args.findLast((a) => typeof a === "function");
      if (callback) process.nextTick(callback, error, "", "");
      return { on() { return this; }, kill() {}, stdout: null, stderr: null };
    }
  };
}
