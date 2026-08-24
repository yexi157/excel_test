# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

Browser-based online Excel system. Vue 3 + Vite + TS frontend with **Univer OSS** as the spreadsheet engine and an **MSW + IndexedDB mock backend** that fully implements the future REST contract. The real backend is intended to be Python/FastAPI but is out of scope at this stage — implement the 7 endpoints in `docs/superpowers/specs/2026-05-10-online-excel-design.md` §7 to swap it in.

Authoritative docs (read before non-trivial changes):
- `docs/superpowers/specs/2026-05-10-online-excel-design.md` — design spec (architecture, data flow, REST contract, risks)
- `docs/superpowers/plans/2026-05-10-online-excel.md` — implementation plan (32 tasks, with code snippets and known API quirks)

## Commands (run from `frontend/`)

```bash
npm run dev          # Vite dev server on :5173 (MSW auto-starts; IndexedDB seeds 3 fixtures on first load)
npm run build        # vue-tsc -b && vite build → dist/
npm run preview      # serve dist/ on :4173 (for offline-mode smoke test)
npm test             # vitest run (3 specs, ~19 tests)
npm test -- tests/xlsxConverter.spec.ts   # single spec
npm run test:watch   # vitest watch
npx vue-tsc --noEmit # standalone type-check (faster than full build)
```

Regenerate test fixtures (one-shot, does NOT add `exceljs` to project deps):
```bash
# from project root
npx --package=exceljs@^4 -- node frontend/scripts/generate-fixtures.mjs
```

Mock backend dev knobs (browser console / URL):
- `localStorage.setItem('mockFailRate', '0.5'); location.reload()` → 50% of API calls 500
- `?reset=1` URL param → wipes IndexedDB and re-seeds fixtures

## Architecture (the parts that span multiple files)

### Hard rule: backend never sees Excel structure

All `IWorkbookData` JSON lives only in the browser. The 7 REST endpoints exchange:
- **JSON** — only file-tree metadata (`FileTreeNode`, `FileMetadata`)
- **`.xlsx` binary** — for `getFile` / `saveFile` / `uploadFile` (Content-Type `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`)

Anywhere you see `xlsxConverter.toUniver` / `toXlsx` — that's the boundary. Don't push JSON snapshots through the API.

### Module dependency direction

```
components → fileStore (Pinia, the only cross-cut)
                ↓
       xlsxConverter   fileApi
                ↓         ↓
       LuckyExcel       axios → MSW → (future) FastAPI
```

`UniverHost.vue` is the **only** code that owns a `Univer` instance. It exposes a `recreate(workbookData)` callback to the store via `bindUniver(univer, api, recreate)`. `fileStore.openFile` and `removeNode` call `recreateUniver` — they never touch Univer directly.

### Why each file exists (the non-obvious ones)

- `utils/xlsxConverter.ts` — abstract `XlsxConverter` interface. The concrete impl (`xlsxConverter.luckyexcel.ts`) is swappable. R1 risk in spec §8.4 is "Luckyexcel breaks on real-world .xlsx"; the abstraction lets you replace just the impl with a SheetJS hand-written converter or backend Python without touching anything else.
- `mocks/{db,seed,handlers,browser}.ts` — full mock backend. `db.ts` is the IndexedDB schema (idb), `seed.ts` injects 3 sample fixtures on first load, `handlers.ts` is 7 MSW endpoints, `browser.ts` boots the worker. The same `handlers` array is re-used in node/test mode by `tests/fileApi.spec.ts` via `setupServer`.
- `utils/guardDirty.ts` — the "unsaved changes" confirmation dialog. Lives in utils (not store) because it depends on `ElMessageBox`. Imported lazily inside `fileStore.openFile` to keep the store testable in node env without DOM.
- `scripts/generate-fixtures.mjs` — generates the test `.xlsx` fixtures via npx-hosted ExcelJS (no project dep). Has a self-bootstrapping shim that locates the npx-cached exceljs through `PATH` (npx doesn't expose its cache to Node's import resolver). Generators are deterministic (Math.sin pseudo-random) but ExcelJS stamps `wb.created` so byte equality is not preserved across runs.

### Critical decisions baked into the code

These were learned the hard way during build-out — don't undo them without a stronger reason than "looks cleaner":

- **Each file open rebuilds the entire `Univer` instance** (not `disposeUnit + createUnit`). The dispose-and-recreate-unit path triggered `RefSelectionsRenderService → SheetsSelectionsService cannot resolve unit` runtime crashes (spec §8.4 R2). Cost is ~300-600ms per file open; the alternative is a hard crash.

- **Dirty detection uses `univerAPI.addEvent(api.Event.SheetValueChanged, ...)`**, NOT `commandService.onCommandExecuted` filtering by `MUTATION`. The mutation listener fires during Univer's own initialization stream (lasting > 200ms with no clean upper bound), so anything timing-based was unreliable. `SheetValueChanged` is user-facing — fires on actual edits, not on `createUnit` data hydration.

- **`fileApi.uploadFile` does NOT set `Content-Type` explicitly**. Setting it manually makes axios skip its FormData detection path and the `boundary=...` parameter is never appended → multipart unparseable. Let axios infer.

- **`fileApi.getFile` uses `responseType: 'arraybuffer'` and wraps in a `Blob` client-side**, not `responseType: 'blob'`. The latter is browser-only; in Node tests the http adapter silently downgrades to a string body. The wrap also gives us explicit MIME type.

- **`fileApi.spec.ts` runs under `// @vitest-environment node`** (not the project default jsdom). MSW's `request.formData()` aborts inside `multipartFormDataParser` under jsdom + Node's bundled undici due to `File`/`FormData` brand mismatches. node env also needs `localStorage` polyfill in `tests/setup.ts` (handlers reference it for the `mockFailRate` knob).

### Known API quirks (Univer 0.22 + @zwight/luckyexcel 1.1.6)

| What you'd expect | Actual |
|---|---|
| `univer.disposeUnit(id)` | Lives on `FUniver` (Facade). Use `univerAPI.disposeUnit(id)`. |
| `LuckyExcel.transformUniverToExcel({snapshot, fileName}, success, error)` | Single params object: `{snapshot, fileName, getBuffer, success, error}`. `success`/`error` are inside the params, not positional. |
| `LuckyExcel.transformExcelToUniver(file, success, error)` | Positional callbacks (asymmetric vs above). |
| `getBuffer: true` (export) | Returns `ArrayBuffer` to `success` callback instead of triggering browser download. |
| `@zwight/luckyexcel` dependency tree | It declares `@univerjs/core@^0.6.0`, but the converter only exchanges plain workbook snapshots across this boundary. Do not import or register its nested Univer runtime in the app. |
| Formula export | Version 1.1.6 reads formula text from `cell.si` instead of `cell.f`; `xlsxConverter` adapts formula cells with copy-on-write before export. |
| Export buffer type | `transformUniverToExcel` may return an `ArrayBuffer` or a Node-style `Buffer`/`Uint8Array`; normalize the view before constructing the Blob. |
| `FSheetHooks.onCellChange()` | Doesn't exist in 0.22. Use `api.Event.SheetValueChanged` (see Dirty detection above). |
| `IAuthzIoService` | Univer's permission DI token. Optional `AuthzIoLocalService` mock is bundled in `@univerjs/core` for single-user use; not registered by default. Register only if a UI permission dialog throws "Cannot resolve IAuthzIoService". |

### Locale dictionary requirement

`UniverHost.vue` MUST pass a `locales` map (merged from each plugin's `/locale/zh-CN` import), not just `locale: ZH_CN`. Without merged dictionaries, Univer throws `LocaleService: Locale not initialized` at runtime. 8 packs are merged: `design`, `ui`, `docs-ui`, `sheets`, `sheets-ui`, `sheets-formula`, `sheets-formula-ui`, `sheets-numfmt-ui`. Note `@univerjs/sheets-numfmt` and `@univerjs/docs` ship no locale dir — don't import from them.

### Plugin registration order (UniverHost.vue)

```
engines (render, formula) → UI → Vue3Adapter → docs (core, UI) → sheets (core, UI) → features (formula, numfmt + their UI)
```

`UniverDocsPlugin` + `UniverDocsUIPlugin` are **required** even though we don't edit docs — `sheets-ui`'s `EditorBridgeService` injects `IEditorService` from `docs-ui`. Without them, runtime crash.

## Hard constraint: no CDN dependencies

Spec §5.5 — production bundle must be self-contained for offline use. Don't add `vite-plugin-cdn-import`, don't `external: ['vue']` in rollup options, don't load fonts from Google Fonts. Verify after build:

```bash
cd frontend && npm run build
grep -rEoh 'https?://[^"'\'' )<>]+' dist/ --include='*.js' --include='*.css' --include='*.html' \
  | grep -v 'localhost' | grep -v 'sourceMappingURL' | sort -u
```

Output should contain only XML namespaces (`http://schemas.openxmlformats.org/...`), help-doc strings, and example placeholders — no runtime CDN refs. (The spec's original `grep -rE '...(?!localhost)...'` recipe uses a PCRE lookahead that BSD/macOS `grep -E` doesn't support and silently exits non-zero — false-positive clean.)

## Where to find what

- Test fixtures: `frontend/public/fixtures/*.xlsx` (also symlinked to `frontend/tests/fixtures/`). Generated, regenerable, but committed for stable test inputs.
- MSW service worker: `frontend/public/mockServiceWorker.js` (generated by `npx msw init`; do not edit).
- Skill references for Univer specifics: `.claude/skills/univer-integrate/references/` (Facade API guide, multi-unit management, permissions, framework integration patterns) — all symlinked from `.agents/skills/`.

## Workflow expectation

This project was built using `superpowers:subagent-driven-development` per the design spec and implementation plan. When making non-trivial changes: read the relevant spec section first, prefer extending the existing abstractions (`xlsxConverter`, `fileApi`, `fileStore`) over adding new cross-cuts, and update spec/plan if the change invalidates documented assumptions.
