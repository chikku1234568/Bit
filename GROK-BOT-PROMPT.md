# Grok bot — next ship: project collab (not Git merge, not one shared .xlsx)

Copy **PROMPT START** through **PROMPT END** into a new Grok session. Repo is `chikku1234568/Bit`, branch `main` after the Excel-add-in + tracking checkpoint.

## PROMPT START

You are continuing **Bit** (version control for Excel). The human already uses a **Windows Excel add-in** + **local Bit agent**. Do not rebuild M1–M6 or the add-in from scratch.

### Product decision (locked)

- **Git merge of workbooks is the wrong UX. Do not pitch or build Combine-as-git-merge as the collab story.**
- **The collaboration unit is a Bit project, not an Excel file.** Two users = two workbooks, **same project**. Never co-author one `.xlsx`.
- **Remote = the project store** (SharePoint library / synced folder), not the workbook.
- **Diffs are the collab feature.** Expand tracking so What changed is trustworthy.
- Collab V1 operation is **Promote to Main** (make this version Main), plus **What changed**. Optional later: apply selected cells. Keep the existing three-way merge *engine* in the repo but do not force it in the add-in ribbon.

Git mapping for *you* (never show Git words in the UI except we already banned them):

| Git | Bit |
|---|---|
| repo | Project |
| working copy | Each person’s own .xlsx |
| branch | Scenario (optional) |
| remotes/origin | Project folder on SharePoint / disk |
| merge | **Do not ship as the main path.** Use Promote + What changed |

### What is already on `main`

- ExcelJS snapshot: values, formulas, fonts, borders, alignment, widths/heights, merges, freeze, hidden, tab colour (`src/xlsx/`, `src/xlsx/format.ts`)
- Local store: `data/meta.json` + `blobs/` + rebuilt `xlsx/` (`src/store/store.ts`) — **one fat meta.json, not atomic**
- Fastify agent `:3001`, Vite add-in `:5173/addin.html`, `addin/manifest.xml`
- Scenarios, diff, merge engine, review/combine (web UI)
- Add-in: Save version from open workbook, scenarios, history, graph, what changed, folder picker
- Node **20.17** — stay on **Vite 6**. No new cloud SDKs unless specified below.
- Tests: `npm test` must stay green.

### What to ship (in order)

#### 1. Append-only project store (required before SharePoint)

Replace “rewrite all of `meta.json`” as the source of truth for versions:

- Keep a **manifest** `project.json` (or slim `meta.json`) for project id/name + **refs**: `main` tip, scenario tips.
- Each Save writes **new files only**: `versions/<versionId>.json` (metadata) and `blobs/<hash>.json` (snapshot). Never mutate an old version file.
- Updating a ref (Main tip) must be **compare-and-swap**: read expected previous tip, write new tip only if it still matches. If not, 409 “Project was updated — Fetch and try again.”
- Optional `project.lock` with expiry for Save/Promote (single writer). If lock exists and is fresh, 409.
- Local disk first (`BIT_DATA_DIR` / chosen folder). Same layout must work if that folder is a **SharePoint-synced OneDrive path**.

This is how two people share a library without Git and without clobbering history.

#### 2. Promote to Main (collab V1, not merge)

- `POST /versions/:id/promote` — set Main tip to this version (CAS on previous tip). New Main version record may be the same snapshot hash (no new blob required) or a pointer. `parentIds` = `[previousMainTip]` (and optionally the promoted id if you want two parents; prefer **one parent = previous Main** plus `promotedFromVersionId` on the version).
- Add-in + web: button **Make this Main** on History/Graph. Confirm. No conflict grid.
- What changed: Main tip vs selected version (already exists — wire it as the default collab view).

Do **not** require Combine/review for this path. Review/combine may stay in the API unused by the add-in.

#### 3. Fetch / two users on one project

- `GET /projects/:id/graph` already exists. Add **Fetch** in the add-in: reload graph, scenarios, tips from disk (the synced folder).
- **Create project** still snapshots *this* workbook. Second user: **Open project** from the same folder (pick folder that already has `meta`/`project.json`) then **Open** a version into a **new** workbook (`Excel.createWorkbook`). They edit their copy, **Save version** onto a scenario or onto Main via Promote.
- Never write into the other person’s open Excel file.

#### 4. SharePoint

**Do not** build Microsoft Graph OAuth / Azure app registration unless the local synced-folder path already works in tests.

Ship as: user picks a folder that is a SharePoint/OneDrive **synced library** (`C:\Users\…\SharePoint\BitProjects\FY27`). Document: OneDrive client must be running; two people must not Save at the same millisecond without CAS (you implemented CAS).

If Graph is trivial later: upload/download the same append-only files. Not a second data model.

#### 5. Tracking — expand for diffs only (as time allows)

Priority for What changed, not for merge:

- Data validation
- Named ranges
- Cell comments
- Tables / autofilter range
- Hyperlink URL

Still out: charts, pivots, VBA, images (optional original-xlsx sidecar per version if easy: save `getFileAsync` bytes next to the snapshot, **Download original** vs Bit copy).

Every new field: `types.ts` → parse → write → round-trip test → diff kind. Do not dump raw ExcelJS objects.

### Constraints

- No `npm` packages that need Node 20.19+ / Vite 8.
- No Postgres, Clerk, S3 required for this milestone.
- Windows PowerShell: `;` not `&&`.
- UI copy: Save version, Scenario, What changed, Make this Main, Fetch. Never commit/PR/merge/repo.
- `npm test` after each slice. Restart `npm run api` and `npm run dev:web` so the add-in picks up routes.
- Do not reimplement `parseXlsx` / `writeXlsx`.
- Update `docs/EXCEL-ADDIN.md` and `README.md` with two-person SharePoint-folder steps.

### Definition of done

1. Two local folders (or two machines sharing a synced folder) can Save version without wiping the other’s version files.
2. Simultaneous Save on Main returns 409 if tips raced; retry after Fetch works.
3. User B opens the project folder, sees User A’s versions on the graph, Opens a version to a new workbook, Saves their own version, **Make this Main**.
4. What changed between Main and that version lists cell/layout diffs.
5. Tests for CAS, append-only, promote.
6. Docs: “same project, two workbooks, SharePoint-synced folder; not one shared Excel.”

Start by reading `src/store/store.ts`, `src/api/app.ts`, `web/src/addin/AddinApp.tsx`, `docs/EXCEL-ADDIN.md`. Then implement store layout + CAS **before** UI.

## PROMPT END
