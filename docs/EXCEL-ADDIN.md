# Bit Excel add-in (local / SharePoint-synced folder)

Windows Excel + a Bit **agent** on this PC. Storage is a **project folder** you pick (`project.json` + `versions/` + `blobs/`). That folder can be a normal disk path or a **SharePoint / OneDrive synced library** folder.

## Collaboration model (locked)

- The unit of collaboration is a **Bit project**, not one Excel file.
- Two people = **two workbooks, same project folder**. Never co-author one shared `.xlsx`.
- Remote = the project store on disk (or synced). Not the workbook.
- Collab V1: **Save version**, **Fetch**, **What changed**, **Make this Main**. No Git merge of workbooks in the ribbon.

## Run (every session)

From the repo root:

```bat
scripts\start-bit.cmd
```

Or two terminals:

```bat
npm run api
npm run dev:web
```

| Process | URL |
|--------|-----|
| Agent | http://127.0.0.1:3001 |
| Add-in task pane | http://127.0.0.1:5173/addin.html |
| Web lab (browser) | http://127.0.0.1:5173/ |

Keep both running. The pane talks to the agent through Vite’s `/api` proxy.

## Sideload in Excel (Microsoft 365 / Excel 2021+)

1. Start Bit as above.
2. Excel → **Insert** → **Add-ins** → **My Add-ins** → **Upload My Add-in**.
3. Choose `addin\manifest.xml` (this repo).
4. Home ribbon → **Bit** group → **Bit** opens the task pane.

If Upload is missing: File → Options → Trust Center → Trust Center Settings → **Trusted Add-in Catalogs** is the enterprise path; Upload My Add-in is enough for you.

First load of `office.js` uses Microsoft’s CDN. After that Excel caches it. Fully air-gapped Excel without that cache will not host the pane.

## Two-person SharePoint-synced project (no Graph OAuth)

1. Create a SharePoint document library (or OneDrive shared folder) for the project, e.g. `BitProjects/FY27`.
2. On each PC, sync that library with the **OneDrive** client so it appears as a normal folder under File Explorer.
3. Person A: in the Bit pane, **Choose folder** / **Open project** → that synced path. **Create project** from their workbook (snapshots into the folder).
4. Person B: same synced path → **Open project**. They see A’s versions after OneDrive sync. **Fetch** reloads tips/history from disk.
5. B: **Open** a version → Excel opens a **new** workbook (`Excel.createWorkbook`). B edits *their* copy, **Save version**, optionally **Make this Main**.
6. A: **Fetch** to see B’s versions / new Main tip. Use **What changed** (Main tip vs selected version).

If two people Save on the same tip at once, one gets **409** — “Project was updated — Fetch and try again.” Fetch, then retry. No Microsoft Graph app registration is required for this path.

Legacy `meta.json` (fat) is migrated on read to slim `project.json` + `versions/*.json`. Resetting `data/` is also fine for demos.

## Buttons

| Control | What it does |
|--------|----------------|
| **Choose folder** / **Open project** | Point the agent at the project store (disk or synced). |
| **Fetch** | Reload projects, scenarios, tips, and graph from the folder. |
| **You are** | Stub author (same as the web app). |
| **Create project** | Snapshot the **open workbook** → Main v1 in this folder. |
| **Save version** | Snapshot the open workbook onto the selected scenario (append-only). |
| **Scenarios / Add** | Optional branch. Switch chip, then Save version. |
| **History → Open** | Opens that version as a **new** workbook. Never writes into someone else’s open file. |
| **Make this Main** | Promote this version to Main tip (CAS). Confirm dialog; no conflict grid. |
| **Graph** | Version tree: nodes, parent edges. |
| **What changed** | Diff Main tip vs selected version (collab default). |

## Fetch the graph (API)

```
GET http://127.0.0.1:3001/projects/<id>/graph
POST http://127.0.0.1:3001/versions/<id>/promote
```

`GET /agent/status` → `{ ok, dataDir, projectCount }`.

## Requirements

- Windows desktop Excel (not Excel Online) for the add-in pane.
- Node 18+ (Node 20 recommended) as already used for Bit.
- `.xlsx` only.
- For shared folders: OneDrive sync client running.

## Charts / macros

Still dropped on Bit reconstruction. Save version reads the real open file via `getFileAsync`; download/Open uses the snapshot rebuild.
