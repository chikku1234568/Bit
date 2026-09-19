# Bit — repo structure

Where things live after **PR #2** (M3–M6 + M7 demo script) on `main`. Git root is this directory (`package.json` + `BIT-V0-DESIGN.md`). Nested clone on the laptop: `C:\Users\Srikar\Bit\Bit`.

---

## 1. Tree (what you actually have)

```
Bit/
├── BIT-V0-DESIGN.md          Product spec (locked decisions, merge table)
├── BIT-GUIDE.md              Short runbook + milestone map
├── README.md                 Install, run, API table
├── docs/
│   ├── REPO-STRUCTURE.md     This file
│   ├── FUNCTIONALITY.md      What works, how to use, comments
│   └── E2E-DEMO.md           Alex / Jordan click path
├── package.json              API + engines (bit@0.6.0)
├── vitest.config.ts          Tests: src/**/*.test.ts
├── tsconfig.json
├── fixtures/sample-budget.xlsx
├── scripts/roundtrip.ts      CLI: parse → write *.roundtrip.xlsx
├── src/
│   ├── xlsx/                 M1 — .xlsx ↔ snapshot
│   ├── domain/               Shared types
│   ├── store/                Local disk DB
│   ├── service/              Application logic
│   ├── api/                  Fastify HTTP
│   ├── diff/                 M4 — pure two-way diff
│   └── merge/                M5 — pure three-way merge
├── web/                      Vite 6 + React UI
└── data/                     Runtime store (gitignored)
```

Not in git: `node_modules/`, `web/node_modules/`, `data/`, `*.roundtrip.xlsx`, `web/dist/`.

---

## 2. Layers (read this if you are lost)

```
Excel (.xlsx)
    ▲  download / upload
    │
web/src/pages          React screens
web/src/api.ts         fetch wrappers → /api/*
    │  Vite proxy strips /api
src/api/app.ts         HTTP routes, multipart, errors
src/service/projects.ts  create / save / scenario / review / combine
    ├── src/xlsx       parseXlsx / writeXlsx
    ├── src/diff       diffSnapshots(base, compare)
    ├── src/merge      mergeSnapshots(base, ours, theirs)
    └── src/store      meta.json + blobs + xlsx cache
```

**Rule:** never reimplement parse/write in service or UI. Diff and merge take snapshots only — no ExcelJS, no Fastify, no `fs`.

---

## 3. Backend (`src/`)

| Path | Milestone | Role |
|---|---|---|
| `src/xlsx/types.ts` | M1 | `Cell` (`v`,`f`,`fmt`), `Sheet`, `WorkbookSnapshot` |
| `src/xlsx/parse.ts` | M1 | `.xlsx` Buffer or path → snapshot |
| `src/xlsx/write.ts` | M1 | snapshot → `.xlsx` Buffer |
| `src/xlsx/index.ts` | M1 | Public exports |
| `src/xlsx/roundtrip.test.ts` | M1 | FP&A fixture write→parse→write→parse |
| `src/domain/types.ts` | M2–M6 | `Project`, `Scenario`, `Version`, `Review` |
| `src/store/store.ts` | M2 | `data/meta.json`, `data/blobs/<sha>.json`, `data/xlsx/<versionId>.xlsx` |
| `src/service/projects.ts` | M2–M6 | All product operations |
| `src/service/projects.test.ts` | M2–M6 | Service + `app.inject` HTTP tests |
| `src/diff/diff.ts` | M4 | `diffSnapshots` → `{ changes[] }` |
| `src/diff/diff.test.ts` | M4 | Value / formula / format / sheet / formula-`v` ignore |
| `src/merge/merge.ts` | M5 | `mergeSnapshots`, `applyResolutions`, `cellsEqual` |
| `src/merge/merge.test.ts` | M5 | Design §9 table: no-op, disjoint, conflict, delete-vs-edit, sheet add |
| `src/api/app.ts` | M2–M6 | Fastify routes + error mapping |
| `src/api/server.ts` | M2 | Listen `:3001`, `BIT_DATA_DIR` |

### Domain objects

- **Project** — one workbook. Points at Main via `mainScenarioId`.
- **Scenario** — named line of versions. Main has `isMain: true`. Tip is `tipVersionId`.
- **Version** — a saved snapshot. `parentIds`: `[]` first version, one parent on Save, **two** after Combine.
- **Review** — scenario → Main. Freezes `baseVersionId` (Main tip at open) and `compareVersionId` (scenario tip at open). Status: `open` \| `changes-requested` \| `combined` \| `closed`.

### Store on disk

```
data/meta.json          projects[], scenarios[], versions[], reviews[]
data/blobs/<sha256>.json    WorkbookSnapshot (content-addressed)
data/xlsx/<versionId>.xlsx  export cache (rewritten from snapshot if missing)
```

`readMeta` fills `reviews: []` if an older M2 `meta.json` has no reviews key. Delete `data/` to reset.

---

## 4. HTTP API (`src/api/app.ts`)

Base URL: `http://127.0.0.1:3001`. UI calls the same paths under `/api` (Vite rewrite).

| Method | Path | Implemented in |
|---|---|---|
| GET | `/health` | `app.ts` |
| GET/POST | `/projects` | `listProjects` / `createProject` |
| GET | `/projects/:id` | `getProject` |
| GET/POST | `/projects/:id/scenarios` | `listScenarios` / `createScenario` |
| GET | `/scenarios/:id` | `getScenario` |
| GET/POST | `/projects/:id/versions` | history / save on **Main** |
| POST | `/scenarios/:id/versions` | save on a scenario |
| GET | `/versions/:id` | metadata |
| GET | `/versions/:id/xlsx` | download |
| GET | `/diff?base=&compare=` | `diffVersions` |
| GET | `/merge?base=&ours=&theirs=` | `previewMerge` (no write) |
| POST | `/scenarios/:id/reviews` | `createReview` |
| GET | `/projects/:id/reviews` | `listReviews` |
| GET | `/reviews/:id` | `getReview` |
| POST | `/reviews/:id/combine` | `combineReview` |
| POST | `/reviews/:id/request-changes` | `updateReviewStatus` |
| POST | `/reviews/:id/close` | `updateReviewStatus` |

Auth stub: header `X-Bit-Author`, or form/JSON `author`, else `demo-user`.

Errors: `NotFoundError` → 404, `ValidationError` → 400, `ConflictError` → 409 (unresolved cells on Combine).

---

## 5. Web UI (`web/`)

Vite 6, React 19, `react-router-dom`. Dev server `:5173`, `server.host: true`, proxy `/api` → `http://127.0.0.1:3001`.

| File | Screen | Route |
|---|---|---|
| `web/src/main.tsx` | Bootstrap | — |
| `web/src/App.tsx` | Shell + **You are** | all |
| `web/src/api.ts` | Typed fetch helpers | — |
| `web/src/index.css` | Layout / cards / tables | — |
| `web/src/pages/HomePage.tsx` | List + create project | `/` |
| `web/src/pages/ProjectPage.tsx` | Scenarios, save, history, reviews | `/projects/:id?scenario=` |
| `web/src/pages/WhatChangedPage.tsx` | Diff table + filters | `/projects/:id/what-changed` |
| `web/src/pages/ReviewPage.tsx` | Combine / Needs a decision | `/reviews/:reviewId` |

`web/package.json` pins **Vite 6** + `@vitejs/plugin-react` 4 + TypeScript 5.8. Vite 8 (rolldown) does not boot on Node 20.17.

---

## 6. Docs vs code

| File | Use |
|---|---|
| `BIT-V0-DESIGN.md` | Why / merge rules / non-goals. Do not treat as “how to run”. |
| `BIT-GUIDE.md` | Commands, copy table, toolchain pin. |
| `docs/E2E-DEMO.md` | Manual two-person walkthrough. |
| `docs/REPO-STRUCTURE.md` | This map. |
| `docs/FUNCTIONALITY.md` | Product + how-to + suggestions. |
| `GROK-BOT-PROMPT.md` | Offline agent brief (untracked unless you add it). |

---

## 7. Tests

```
npm test    # vitest run — 36 tests on current main
```

| File | Count (approx) | Covers |
|---|---|---|
| `src/xlsx/roundtrip.test.ts` | 7 | Fidelity bar |
| `src/service/projects.test.ts` | 11 | Create/save/download, scenarios, review, combine, HTTP |
| `src/diff/diff.test.ts` | 7 | Diff kinds |
| `src/merge/merge.test.ts` | 11 | Three-way table |

HTTP tests use Fastify `app.inject` and a temp `data` dir. They do not need the UI.

---

## 8. Git history (useful `main` commits)

```
9fb0340  M1 xlsx bridge
81f84fa  M2 versions (#1)
39cdaef  M2 follow-up: Vite 6 pin, host bind, BIT-GUIDE
310b0d4  M3 scenarios
f2c7d6b  M4 What changed
f6c7395  M5 merge engine
60c2010  M6 review + Combine
d74b481  M7 E2E demo script
e36d11d  Merge pull request #2 (offline-m3-m6)
```

Package version `0.6.0` means “M6 landed”, not a public semver.

---

## 9. Where to change what

| You want to… | Edit |
|---|---|
| Track a new format field | `src/xlsx/types.ts` + parse + write + round-trip test |
| Change merge rules | `src/merge/merge.ts` + `merge.test.ts` only |
| Add an API route | `src/api/app.ts` + `ProjectService` + HTTP test |
| Add a screen | `web/src/pages/*` + route in `App.tsx` + helper in `api.ts` |
| Change stored records | `src/domain/types.ts` + `store.ts` `MetaDb` |
| Demo copy / click path | `docs/E2E-DEMO.md` / `docs/FUNCTIONALITY.md` |
