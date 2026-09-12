# Bit — M2 Versions

Version control for Excel. **M2** adds Project, Main, Save version, history, and download on top of the M1 xlsx bridge.

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
4. Edit the workbook in Excel, then **Save version** (upload + note).
5. **Download .xlsx** for any version.

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

Includes M1 xlsx round-trips and M2 create → save version → history → download (service + HTTP).

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
src/api/           # Fastify HTTP API
web/               # Vite + React UI (Home, Project)
data/              # Runtime store (gitignored)
fixtures/          # Sample budget workbook
```

## API (M2)

| Method | Path | Purpose |
|--------|------|---------|
| GET | `/projects` | List projects |
| POST | `/projects` | Create (multipart: `file`, `name`, `message`, `author`) |
| GET | `/projects/:id` | Project + Main tip |
| GET | `/projects/:id/versions` | Version history |
| POST | `/projects/:id/versions` | Save version on Main |
| POST | `/scenarios/:id/versions` | Save version on a scenario |
| GET | `/versions/:id` | Version metadata |
| GET | `/versions/:id/xlsx` | Download export |

## Fidelity bar (tracked)

- Sheet names and order
- Cell values (`v`) and formulas (`f`)
- Basic formatting: bold, italic, fill colour, font colour, number format

## Out of scope (later milestones)

Scenarios UI (M3), What changed (M4), merge (M5), review (M6), invite / two-user E2E (M7).

## Design

See `BIT-V0-DESIGN.md`.
