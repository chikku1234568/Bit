# Bit — add-in runbook

The product is the **Excel add-in** (task pane + local agent). Full commands: [README.md](README.md). Limits: [docs/LIMITATIONS.md](docs/LIMITATIONS.md).

Use the **web lab** at http://127.0.0.1:5173 to try Save version, Fetch, Graph, What changed, and Make this Main **before** you sideload anything in Excel. Same agent, same project folder. The pane is how you snapshot the **open** workbook; the browser cannot read Excel.

## Start

```bat
npm install
cd web
npm install
cd ..
scripts\start-bit.cmd
```

Or: `npm run api` and `npm run dev:web`.

| | URL |
|--|-----|
| Agent | http://127.0.0.1:3001 |
| **Web lab (try first)** | http://127.0.0.1:5173 |
| Add-in (Excel loads this) | http://127.0.0.1:5173/addin.html |

Then sideload `addin\manifest.xml` — [README §4](README.md#4-sideload-the-excel-add-in-once-per-machine).

## Product (add-in)

Same **project folder**, two **workbooks**. Not one shared `.xlsx`.

| Button | Meaning |
|--------|---------|
| Choose folder | Point the agent at disk or a OneDrive/SharePoint **sync** path |
| Fetch | Reload the graph from that folder |
| Create project / Save version | Snapshot the **open** Excel file into the project |
| Open | New workbook from a version |
| What changed | Diff Main vs a version |
| Make this Main | Point Main at that snapshot |

**You are** is a display name (`localStorage` `bit-author`), not login.

## UI copy (never Git words)

| Concept | Bit |
|---------|-----|
| Repo | Project |
| Commit | Save version |
| `main` | Main |
| Branch | Scenario |
| Diff | What changed |
| Checkout | Open this version |

## Store

`project.json` + `versions/` + `blobs/` in the chosen folder (often OneDrive). Default if you never pick a folder: `./data` (gitignored). Last path: `%USERPROFILE%\.bit\config.json`.

## Fidelity

**Kept:** values, formulas, formatting/layout subset, hyperlinks, validation, names, comments, tables.

**Dropped on Open:** charts, pivots, images, macros / `.xlsm`.

## Layout

```
addin/manifest.xml
src/xlsx/      parse/write snapshot
src/store/     append-only project store
src/api/       agent :3001
web/           lab + add-in pane :5173
```
