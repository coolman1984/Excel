# Plan 4 restart (2026-11-08): read this, then PLAN-4-DATA-ENGINE.md, then continue your lane

The first run stopped on sandbox problems. The lead fixed them and proved each fix inside the Codex sandbox.
**Everything below works now. Do not stop for these again.** BLOCKED is only for a tool error you cannot work around
after two tries; then write the exact error in your answer and continue with the next step that does not need it.

## What changed
1. **Lane checks.** Run only your own suites:
   - Lane A: `powershell -NoProfile -File check.ps1 -Offline -Suites store,ingest`
   - Lane B: `powershell -NoProfile -File check.ps1 -Offline -Suites query,automation,formula`
   The older suites (host, style, save, the viewer tests) are the lead's and are not run by lanes; they are the ones
   that touch the approved cache folder. `check.ps1` now builds with one MSBuild node (`-m:1`): worker nodes cannot start
   in the sandbox, which was the "build failed with 0 errors" of the first run.
2. **Test folder.** `D:\ExcelAppTest` is writable for lanes A and B (`StoreRoot.ForTest` / `ConfigRoot.ForTest` under
   `D:\ExcelAppTest\syn*`). Never `%TEMP%`.
3. **Interface tools (lane C).** Inside `ExcelApp/ui`, with `$env:NODE_OPTIONS='--require=./sandbox-shim.cjs'` set:
   - `pnpm exec tsc --noEmit`
   - `pnpm exec vitest run --pool=threads --configLoader runner`
   - `pnpm exec vite build --configLoader runner`
   Do not run `pnpm install`: `node_modules` is already linked. Screenshots (C16) need Chrome, which cannot start in the
   sandbox: write the list of pages and URLs in your answer; **the lead takes and inspects the pictures** and sends defects back.
4. **Live tests** (Excel through COM: A13, the live part of B13). Excel cannot start in the sandbox. Write the live tests
   fully (they run when `offline == false`), run every offline test yourself, and say in the answer that the live run is
   the lead's. The lead runs `check.ps1` live and sends failures back.
5. **Answer and progress files** now live inside your worktree: `.lane/answer.md` and `.lane/progress.md`
   (create the folder). Same content as section 9 of the plan.

## Where each lane stands
- **Lane A** (`Excel-wt-a`, branch `plan4-a`): nothing written. Start at A1. Its first-run concerns are answered by points 1 and 2.
- **Lane B** (`Excel-wt-b`, branch `plan4-b`): B1 started (`Data/Query/QuerySpec.cs`, `Data/Query/Expr.cs`, `QueryTests.cs`).
  Review what is there, finish B1 (Pred interpretation and its tests), then B2 to B13.
- **Lane C** (`Excel-wt-c`, branch `plan4-c`): C1 and C2 done. `tsc` fails at `src/data/ResultGrid.tsx` line 47
  (`altKey` is not on `CellClickedEventArgs`: read the modifier from the event's native mouse event or track Alt with a
  key listener). Fix it, then C3 to C16.

## Rules (unchanged, section 0 of the plan)
Real data never leaves the store; Excel is the only reader of company files; never open `D:\ExcelViewerCache` or any `.xvc`;
made-up data only; no logging of cell values or names; no git, no other agents, no agent_bridge; own only your files.
Work through **every** step of your lane in order, checking after each one. Do not stop to ask.

## Contract changes decided by the lead (2026-11-08, after lane C's requests)
1. `DiffResult` now extends `ResultInfo`: `data.history.diff` and `data.sandbox.compare` return `resultId, rows, columns, ms`
   plus `added, removed, changed, changedColumns`. Lane A (history diff) and lane B (sandbox compare) return that shape.
2. `_status` is the one reserved column that is sent: listed in `ResultInfo.columns` of diff results and paged by
   `data.page` (lane B Format). `_v` and `_r` stay private.
3. `data.settings.get` / `.set` return `SettingsInfo { settings, storeFolder, configFolder, pdfApproved }` (lane A).
   `pdfApproved` is false until the owner confirms (plan 7.4); the interface shows AllowPdf disabled while it is false.
