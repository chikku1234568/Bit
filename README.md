# Bit — M6 Review + Combine

Version control for Excel. **M6** adds Ask for review, Needs a decision, and Combine into Main.

Practical onboarding: see **[BIT-GUIDE.md](BIT-GUIDE.md)**.

## Install

```bash
npm install
cd web && npm install && cd ..
```

Requires Node.js 18+.

## Run locally

Terminal 1 — API (default `http://127.0.0.1:3001`, data in `./data`):

```bash
npm run api
```

Terminal 2 — Web UI (Vite, proxies `/api` → API):

```bash
npm run dev:web
```

Open `http://127.0.0.1:5173`.

### Demo flow

1. Set **You are** in the top bar (stub author; no real auth in M2).
2. **Create project** — name + upload `.xlsx` → Main version 1.
3. Open the project — see Main tip and version history.
4. **Create scenario**, then **Switch scenario**.
5. Edit in Excel, **Save version** on the selected scenario (Main tip stays put).
6. On a scenario, **Ask for review** → open Review → **Combine into Main** (resolve **Needs a decision** if any).
7. **Download .xlsx** for the new Main tip.

Optional env:

- `PORT` / `HOST` — API listen address
- `BIT_DATA_DIR` — store root (default `./data`)
- `VITE_API_BASE` — UI API prefix (default `/api` via Vite proxy)

### Auth (V0 stub)

No login. Author is taken from:

1. `X-Bit-Author` request header, or
2. multipart form field `author`, or
3. default `demo-user`

The UI stores the name in `localStorage` (`bit-author`). Invite / two-user is M7 — out of scope here.

## Test

```bash
npm test
```

Includes M1–M6: xlsx, versions, scenarios, diff, merge, and review/combine tests.

## Round-trip CLI (M1)

```bash
npm run roundtrip -- path/to/file.xlsx
```

## Layout

```
src/xlsx/          # M1 bridge — parse/write snapshot (do not reimplement)
src/domain/        # Project, Version, Scenario (Main)
src/store/         # Local demo store: data/meta.json + blobs/ + xlsx/
src/service/       # ProjectService
src/diff/          # Pure snapshot diff (M4)
src/merge/         # Pure three-way merge (M5)
src/api/           # Fastify HTTP API
web/               # Vite + React UI (Home, Project)
data/              # Runtime store (gitignored)
fixtures/          # Sample budget workbook
```

## API (M6)

| Method | Path | Purpose |
|--------|------|---------|
| GET | `/projects` | List projects |
| POST | `/projects` | Create (multipart: `file`, `name`, `message`, `author`) |
| GET | `/projects/:id` | Project + Main tip + scenarios |
| GET | `/projects/:id/scenarios` | List scenarios (Main first) |
| POST | `/projects/:id/scenarios` | Create scenario from Main tip (`{ name }`) |
| GET | `/scenarios/:id` | Scenario metadata |
| GET | `/projects/:id/versions` | Version history |
| POST | `/projects/:id/versions` | Save version on Main |
| POST | `/scenarios/:id/versions` | Save version on a scenario |
| GET | `/versions/:id` | Version metadata |
| GET | `/versions/:id/xlsx` | Download export |
| GET | `/diff?base=&compare=` | Cell/sheet diff between versions |
| GET | `/merge?base=&ours=&theirs=` | Merge preview (conflicts, no write) |
| POST | `/scenarios/:id/reviews` | Ask for review (freeze base/compare) |
| GET | `/projects/:id/reviews` | List reviews |
| GET | `/reviews/:id` | Review detail |
| POST | `/reviews/:id/combine` | Combine into Main (resolutions) |
| POST | `/reviews/:id/request-changes` | Request changes |
| POST | `/reviews/:id/close` | Close review |

## Fidelity bar (tracked)

- Sheet names and order
- Cell values (`v`) and formulas (`f`)
- Basic formatting: bold, italic, fill colour, font colour, number format

## Out of scope (later milestones)

E2E demo script (M7).

## Design

See `BIT-V0-DESIGN.md`.
