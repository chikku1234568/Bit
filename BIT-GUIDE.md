# Bit — practical guide

Onboarding for the local demo. Design rules live in `BIT-V0-DESIGN.md`; this file is how to run and where things live.

## How to run

```bash
npm install
cd web && npm install && cd ..
```

| Process | Command | URL |
|--------|---------|-----|
| API | `npm run api` | `http://127.0.0.1:3001` |
| Web | `npm run dev:web` | `http://127.0.0.1:5173` (Vite binds `0.0.0.0`; proxies `/api` → API) |
| Tests | `npm test` | — |

Optional env: `PORT`, `HOST`, `BIT_DATA_DIR` (default `./data`), `VITE_API_BASE` (default `/api`).

Reset store: delete `./data/` (gitignored). Meta + blobs + cached `.xlsx` live there.

## Milestone map

| Milestone | Outcome | Status target |
|-----------|---------|---------------|
| **M1** | Parse/write `.xlsx`; snapshot schema; round-trip | Done |
| **M2** | Project, Main, Save version, history, download | Done |
| **M3** | Create scenario, switch, save on scenario | Done |
| **M4** | What changed (diff) | Done |
| **M5** | Three-way merge engine + preview API | Done |
| **M6** | Ask for review, Needs a decision, Combine into Main | Done |
| **M7** | Two-user E2E demo script (Alex/Jordan) | Next |

## Key paths

```
src/xlsx/           # parseXlsx / writeXlsx — import only; do not reimplement
src/domain/         # Project, Scenario, Version types
src/store/          # Local demo store: data/meta.json + blobs/ + xlsx/
src/service/        # ProjectService
src/api/            # Fastify HTTP API
src/diff/           # Pure snapshot diff (M4+)
src/merge/          # Pure three-way merge (M5+)
web/src/pages/      # Home, Project, What changed, Review
fixtures/           # Sample budget workbook
docs/               # E2E demo notes (M7)
```

## Auth stub (through M6)

No real login. Author comes from, in order:

1. `X-Bit-Author` request header  
2. Multipart form field `author`  
3. Default `demo-user`

UI: **You are** in the top bar → `localStorage` key `bit-author`.

## UI copy (never show Git words)

| Concept | Bit UI |
|---------|--------|
| Project | **Project** |
| Commit | **Save version** |
| Branch main | **Main** |
| Feature branch | **Scenario** |
| Diff | **What changed** |
| Pull request | **Ask for review** |
| Merge | **Combine into Main** |
| Conflict | **Needs a decision** |
| Clone/pull/push | **Sync** |
| Checkout | **Open this version** / **Switch scenario** |

## Web toolchain pin

Vite **6** + `@vitejs/plugin-react` **4** + TypeScript **5.8**. Do not upgrade to Vite 8 / TS 6 for V0.

## Constraints (V0)

- No Postgres, S3, Clerk, Auth.js, Redis, or cloud SDKs.
- Merge/diff engines are pure: snapshots in → result out (no Fastify / ExcelJS / fs).
- Reuse existing deps only (exceljs, fastify, cors, multipart, vitest, tsx, typescript, React 19, react-router-dom, Vite 6).
