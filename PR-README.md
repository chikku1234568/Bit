# PR #3 — Collab: append-only project store + Make this Main

**Branch:** `collab-append-store` → `main`  
**PR:** https://github.com/chikku1234568/Bit/pull/3  
**Base:** `a773191` (Excel add-in + richer snapshot tracking)

This document is the reviewer / demo guide for everything in this PR. Product design lives in `BIT-V0-DESIGN.md`; day-to-day runbooks in `BIT-GUIDE.md` and `docs/EXCEL-ADDIN.md`.

---

## 1. Why this PR exists

Bit’s collab story is **not** “two people edit one shared `.xlsx`” and **not** “Git-merge two workbooks in the ribbon.”

| Locked decision | Meaning |
|-----------------|---------|
| Unit of collab | A **Bit project** |
| Two users | **Two workbooks**, same project folder |
| Remote | The **project store** on disk (or a SharePoint / OneDrive **synced** folder) |
| Collab V1 ops | **Save version**, **Fetch**, **What changed**, **Make this Main** |
| Merge engine | Still in the repo for the web lab; **not** the add-in collab path |

Never co-author one open Excel file. Never pitch Combine-as-git-merge as the primary UX.

---

## 2. What’s in this PR (by commit)

| Commit | Slice |
|--------|--------|
| `de7d777` | **Store:** append-only versions + CAS refs + lock; fat `meta.json` migrate-on-read |
| `d67aafa` | **Collab:** Make this Main (promote) API + UI; Fetch / Open project from folder |
| `2f21831` | **Docs:** two-person SharePoint-synced project folder |
| `05e32a9` | **Tracking:** cell hyperlink URL for What changed |
| `cecbc73` | **Tracking:** validation, named ranges, comments, tables, AutoFilter |

---

## 3. Store layout (append-only)

Per project directory (default under `BIT_DATA_DIR` / chosen folder):

```
<project-folder>/
  project.json          # slim manifest: id, name, refs (Main tip + scenario tips)
  project.lock          # optional short-lived lock for Save / Promote
  versions/<id>.json    # version metadata (immutable once written)
  blobs/<sha256>.json   # content-addressed workbook snapshots (immutable)
  xlsx/<versionId>.xlsx # optional rebuilt export cache
```

### Rules

1. **Save writes new files only** — never mutate an existing `versions/*.json` or blob.
2. **Tip updates are compare-and-swap (CAS)** — reader supplies expected previous tip; write succeeds only if the tip still matches. Otherwise **HTTP 409** with:  
   `Project was updated — Fetch and try again.`
3. **`project.lock`** — if a fresh lock exists, Save/Promote returns **409**.
4. **Legacy** — old fat `meta.json` is **migrated on read** to the slim layout. Deleting the data folder is fine for demos.

This is how two people share a library without Git and without clobbering history — including when the folder is a OneDrive/SharePoint sync root.

---

## 4. Promote to Main (collab V1)

### API

```http
POST /versions/:id/promote
```

- CAS on previous Main tip.
- New Main version record may **reuse the same snapshot hash** (no new blob required).
- `parentIds = [previousMainTip]`
- `promotedFromVersionId` points at the version that was promoted.
- **No conflict grid.** Confirm in UI, then promote.

### UI

- Add-in History / Graph: **Make this Main**
- Web Project page: same
- **What changed** default for collab: **Main tip vs selected version**

Review / Combine remain available in the web lab API/UI but are **not** required for this path and are not the add-in collab story.

---

## 5. Fetch + Open project (two users)

| Action | Behavior |
|--------|----------|
| **Choose folder / Open project** | Point the agent at a folder that already has `project.json` (or legacy `meta.json`). |
| **Fetch** | Reload projects, scenarios, tips, and graph from disk (after sync). |
| **Open** (history) | Opens that version as a **new** workbook (`Excel.createWorkbook`). Never writes into someone else’s open file. |
| **Create project** | Still snapshots *this* workbook into the folder as Main v1. |
| **Save version** | Snapshot the **open** workbook onto the selected scenario (append-only). |

### Two-person SharePoint-synced folder (no Graph OAuth)

1. Shared library synced with the **OneDrive** client on each PC.  
2. Person A: Choose folder → Create project from their workbook.  
3. Person B: same folder → Open project → Fetch after sync.  
4. B: Open a version → **new** workbook → edit → Save version → optionally **Make this Main**.  
5. A: Fetch → see new tip → What changed (Main vs that version).  
6. Tip race → **409** → Fetch → retry.

Details: `docs/EXCEL-ADDIN.md`.

---

## 6. What changed — tracking expansion

Diff engine: `src/diff/diff.ts`. Snapshots: `src/xlsx/`.

| Tracked for diffs | Diff kind (UI label) | Notes / fidelity limits |
|-------------------|----------------------|-------------------------|
| Values / formulas / sheets / layout / rich fmt | existing kinds | From earlier checkpoints |
| Cell hyperlink URL | `cell-hyperlink` (Hyperlink) | |
| Data validation | `validation` (Validation) | Multi-cell `sqref` expanded **per-cell** on read |
| Named ranges | `named-range` (Named range) | Skips `_xlnm.*`; sheet scope via `Sheet!Name` |
| Cell comments | `cell-comment` (Comment) | **Text only** — author not round-tripped (ExcelJS) |
| Tables | `table` (Table) | Metadata (`name`/`ref`/header/totals); styles not tracked |
| AutoFilter range | `auto-filter` (AutoFilter) | |

**Still out:** charts, pivots, VBA, images (optional original-xlsx sidecar later).

Every new field follows: `types.ts` → parse → write → round-trip test → diff kind. No raw ExcelJS objects in the snapshot.

---

## 7. How to run (Windows)

```bat
scripts\start-bit.cmd
```

Or:

```bat
npm run api
npm run addin
```

| Process | URL |
|---------|-----|
| Agent | http://127.0.0.1:3001 |
| Add-in pane | http://127.0.0.1:5173/addin.html |
| Web lab | http://127.0.0.1:5173/ |

Sideload `addin\manifest.xml` (Insert → Add-ins → My Add-ins → Upload My Add-in).

Env:

- `BIT_DATA_DIR` — store root (default `./data`)
- `PORT` / `HOST` — API listen
- `VITE_API_BASE` — UI API prefix (default `/api`)

Node **20.17** / **Vite 6** — do not upgrade to Vite 8 / rolldown on this machine.

---

## 8. Tests

```bat
npm test
```

**59 passed** at tip `cecbc73` (and later tip after this doc commit), covering:

- Append-only (old version file unchanged after later save)
- CAS race → 409
- Promote HTTP + tip advance
- Hyperlink / validation / named ranges / comments / tables / AutoFilter round-trip + diff
- Prior M1–M6 suites still green

---

## 9. Manual test plan (definition of done)

- [ ] Two local folders (or two machines on one synced folder) can **Save version** without wiping the other’s version files.
- [ ] Simultaneous Save on Main → **409**; **Fetch** then retry succeeds.
- [ ] User B opens the project folder, sees User A’s versions on the graph, **Opens** a version to a **new** workbook, Saves their own version, **Make this Main**.
- [ ] **What changed** between Main and that version lists cell/layout (and new) diffs.
- [ ] Sideload add-in: Fetch + Make this Main + What changed smoke.
- [ ] Docs match reality: “same project, two workbooks, SharePoint-synced folder; not one shared Excel.”

---

## 10. Key paths touched

```
src/store/          # append-only project.json + versions/ + CAS + lock
src/domain/         # promotedFromVersionId, etc.
src/service/        # promote, open-from-folder flows
src/api/app.ts      # POST /versions/:id/promote, 409 messages
src/xlsx/           # parse/write + new tracked fields
src/diff/           # new diff kinds + labels
web/src/addin/      # Make this Main, Fetch, Open project
web/src/pages/      # Promote + What changed defaults
docs/EXCEL-ADDIN.md
README.md
BIT-GUIDE.md
PR-README.md        # this file
```

Merge engine (`src/merge/`) and Ask for review / Combine (web) are **unchanged in intent** — keep for lab; add-in collab path does not depend on them.

---

## 11. Out of scope (intentionally)

- Microsoft Graph OAuth / Azure app registration
- Postgres / Clerk / S3 / free hosted “bithub.com” storage
- Co-authoring one `.xlsx`
- Insert-aware merge / forcing Combine in the add-in ribbon
- Charts, pivots, VBA, images as first-class snapshot fields
- Comment **author** round-trip (ExcelJS limitation)

**Later (product direction, not this PR):** pluggable BYO cloud bucket (AWS / GCP / Azure) with the **same** append-only file layout — local remains the default for demos.

---

## 12. Reviewer checklist

1. Read this file + skim `docs/EXCEL-ADDIN.md` § collaboration model.  
2. `npm test` green.  
3. Spot-check CAS / promote tests under `src/service/`.  
4. Optional: Windows sideload + two folders or a synced path.  
5. Confirm UI copy has no Git jargon (Save version / Scenario / What changed / Make this Main / Fetch).

---

## 13. Related docs

| Doc | Role |
|-----|------|
| `BIT-V0-DESIGN.md` | Product design |
| `BIT-GUIDE.md` | Practical onboarding |
| `docs/EXCEL-ADDIN.md` | Add-in + SharePoint-synced folder |
| `docs/E2E-DEMO.md` | Alex / Jordan click path |
| `docs/FUNCTIONALITY.md` | Feature inventory |
| `docs/REPO-STRUCTURE.md` | Tree map |

Questions / follow-ups: continue on this branch or open a successor PR after merge.
