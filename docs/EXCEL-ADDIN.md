# Bit Excel add-in (local folder)

Windows Excel + a Bit **agent** on this PC. Storage is a folder you pick (`meta.json` + `blobs/`). Not OneDrive multiplayer. Not AppSource.

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

## Buttons

| Control | What it does |
|--------|----------------|
| **Choose folder** | Native folder picker (agent). Default is `./data` under the repo, or `~/.bit/config.json`. |
| **You are** | Stub author (same as the web app). |
| **Create project** | Snapshot the **open workbook** → Main v1. |
| **Save version** | Snapshot the open workbook onto the selected scenario. |
| **Scenarios / Add** | Branch. Switch chip, then Save version. |
| **History → Open** | Opens that version as a **new** workbook (`Excel.createWorkbook`). |
| **Graph** | Version tree: nodes, parent edges, combine = two parents. |
| **Changed** | What changed vs Main tip (or previous version). |

Excel file lock: **Open** does not overwrite the file you are editing.

## Fetch the graph (API)

```
GET http://127.0.0.1:3001/projects/<id>/graph
```

```json
{
  "nodes": [{ "id", "message", "author", "timestamp", "scenarioName", "isTip", "parentIds" }],
  "edges": [{ "from", "to" }],
  "scenarios": [{ "id", "name", "isMain", "tipVersionId" }]
}
```

`GET /agent/status` → `{ ok, dataDir, projectCount }`.

## Requirements

- Windows desktop Excel (not Excel Online).
- Node 18+ as already used for Bit.
- `.xlsx` only.

## Charts / macros

Still dropped on Bit reconstruction. Save version reads the real open file via `getFileAsync`; download/Open uses the snapshot rebuild.
