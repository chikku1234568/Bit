# Bit V0 — Design Document

**Bit**: version control for Excel (Git-shaped, spreadsheet-native).

| | |
|---|---|
| Status | Draft for build |
| Version | V0 (minimum end-to-end) |
| Date | 2026-09-11 |
| Codename | Bit |

---

## 1. Purpose

Bit lets non-technical Excel users collaborate on a **single workbook** with real version history, scenarios (branches), review, and merge — without OneDrive-style last-write-wins, and without learning Git.

**V0 success (non-negotiable):** two users can collaborate on one `.xlsx` through Bit with tracking, version control, and merge rules — end to end.

---

## 2. Goals and non-goals

### 2.1 Goals (V0)

- End-to-end happy path for **exactly two collaborators** on **one workbook project**
- Save versions with notes and authorship
- Create a **scenario** (branch), edit, save versions
- **Ask for review** before combining into Main
- **Combine (merge)** with clear cell-level rules and conflict resolution in Bit’s UI
- Export a real `.xlsx` after merge
- Diff: what changed between two versions (cells, formulas, basic formatting, sheet structure)
- Language and UX fit **FP&A / finance** and **consultants** (not developers)

### 2.2 Non-goals (explicitly later)

- Excel add-in as the *shipped* V0 surface (north star; V0 demo is web-first)
- Mac / Excel for the web
- `.xlsm` / VBA / macros
- Spawning branches as separate Excel files (“branch into other excels”)
- Special handling for very large / exotic formulae
- Multi-sheet collaboration as a separate product concept (sheet-level permissions, per-sheet owners)
- Real-time co-editing cursors (Google Docs style)
- Teams larger than two, SSO, enterprise admin, compliance exports
- CLI as a user-facing product

---

## 3. Decisions (locked)

| Topic | V0 decision |
|---|---|
| Product name | **Bit** |
| Primary demo surface | **Web app** (fastest path); Excel add-in is the goal after V0 |
| Excel for demo | **Windows desktop Excel** (open/export `.xlsx` locally) |
| File type | **`.xlsx` only** |
| Collaboration | **Review flow in V0** (ask for review → approve/combine) |
| Project model | **One workbook = one Bit project** |
| Tracked content | Cell **values**, **formulas**, **sheet structure**, **basic formatting** |
| Conflict UX | Resolve in **Bit UI** (grid: ours / theirs / edit), then export `.xlsx` |
| Beachhead | **FP&A / finance** and **consultants** |
| Scope bar | **Minimum full flow that works E2E** — nothing extra |

---

## 4. Personas and jobs

### 4.1 Primary personas

1. **FP&A analyst** — owns the budget model; needs a safe Main and a clear audit of what changed.
2. **Consultant / second analyst** — proposes changes in a scenario without overwriting Main.

### 4.2 Jobs to be done

- “I need a copy of the truth that isn’t `Budget_FINAL_v7.xlsx`.”
- “I want a colleague to change assumptions without silently clobbering mine.”
- “I need to see exactly which cells changed before I accept their work.”
- “When we both edited the same cell, I need a simple way to choose.”

---

## 5. User language (hide Git)

| Internal / Git concept | Bit UI copy |
|---|---|
| Repository / project | **Project** (one workbook) |
| Commit | **Save version** |
| Branch `main` | **Main** |
| Feature branch | **Scenario** |
| Diff | **What changed** |
| Pull request | **Ask for review** |
| Merge | **Combine into Main** |
| Conflict | **Needs a decision** |
| Clone / pull / push | **Sync** |
| Checkout | **Open this version** / **Switch scenario** |

---

## 6. V0 end-to-end flow (the product)

Happy path for two users: **Alex** (owner) and **Jordan** (collaborator).

```
1. Alex creates a Bit project and uploads budget.xlsx
2. Bit parses → cell model → creates Main with version 1
3. Alex invites Jordan (email / link — simplest possible)
4. Jordan opens the project, creates scenario "Tax update"
5. Jordan downloads/opens working copy, edits in Excel (Windows), re-uploads / saves version
6. Jordan clicks Ask for review → review request targeting Main
7. Alex opens What changed (cell/formatting diff)
8. Alex Combine into Main
   - auto-merge clean cells
   - for conflicts: Bit grid → Ours / Theirs / Edit
9. Bit writes new Main version and offers Download .xlsx
10. Both Sync; history shows who changed what
```

If this path works reliably, V0 is done.

---

## 7. Information architecture

### 7.1 Project

- ID, name, created_at
- Exactly one workbook identity (canonical `.xlsx` artefact per version)
- Members: two seats in V0 (owner + collaborator)
- Scenarios: `Main` (default, protected) + zero or more named scenarios

### 7.2 Version (commit)

- Parent version(s) — one parent normally; two parents after combine
- Author, timestamp, message
- Snapshot pointer → immutable cell-model blob
- Optional stored `.xlsx` artefact (export cache)

### 7.3 Scenario (branch)

- Name (e.g. `Tax update`)
- Pointer to tip version
- `Main` cannot be force-overwritten; updates only via Combine or direct Save on Main (product rule: prefer changes via scenario + review)

### 7.4 Review request

- From scenario → Main
- Status: open / approved / combined / closed
- Base version (Main tip when review opened) + compare version (scenario tip)
- Discussion: V0 can be a single optional note field; threaded comments later

---

## 8. Domain model (cell snapshot)

Canonical snapshot stored per version (JSON or equivalent):

```json
{
  "sheets": {
    "Budget": {
      "dimensions": {"rows": 200, "cols": 20},
      "cells": {
        "A1": {
          "v": 100,
          "f": null,
          "fmt": {"numFmt": "0.00", "bold": true, "fill": "#FFFF00"}
        },
        "B1": {
          "v": null,
          "f": "=A1*1.1",
          "fmt": {"numFmt": "0.00"}
        }
      }
    }
  },
  "sheetOrder": ["Budget"]
}
```

### 8.1 Tracked in V0

- Sheet names and order (structure)
- Per-cell value (`v`) and formula (`f`)
- Basic formatting: bold/italic, fill colour, font colour, number format (extend only if cheap)

### 8.2 Not tracked in V0

- Charts, pivots, images, controls
- Data validation, conditional formatting rules (beyond what’s implied by stored direct formats)
- Named ranges, VBA, external links
- Calculated cache values as source of truth (store formula; value optional for display)

### 8.3 Round-trip

- **Import:** `.xlsx` → snapshot (parser)
- **Export:** snapshot → `.xlsx` (writer)
- V0 fidelity bar: values, formulas, sheet structure, basic formatting survive upload → save → download without corruption of tracked fields

---

## 9. Merge rules (core IP)

Three-way merge per cell address, per sheet, using:

- **Base** = common ancestor version  
- **Ours** = Main tip  
- **Theirs** = scenario tip  

### 9.1 Cell equality

Two cell states are equal if `v`, `f`, and tracked `fmt` match (normalised).

### 9.2 Per-cell outcomes

| Base | Ours | Theirs | Result |
|---|---|---|---|
| X | X | X | No change |
| X | Y | X | Take Ours |
| X | X | Y | Take Theirs |
| X | Y | Y | Take Y (both same) |
| X | Y | Z | **Conflict** (Y ≠ Z) |
| missing | Y | missing | Take Ours (added on Main) |
| missing | missing | Y | Take Theirs (added on scenario) |
| missing | Y | Z | **Conflict** if Y ≠ Z; else Y |
| X | missing | X | Deleted on Ours → delete |
| X | X | missing | Deleted on Theirs → delete |
| X | missing | Y | **Conflict** (delete vs edit) |
| X | Y | missing | **Conflict** (edit vs delete) |

Same rules apply to sheet add/remove/rename at sheet grain:

- Sheet only on one side → take that side  
- Sheet on both with different cell merges → merge cell-wise  
- Both renamed differently → conflict (V0: treat rename as delete+add if detection is weak; document limitation)

### 9.3 Row/column insertion

**V0 limitation (accepted):** merge is **address-stable** (A1 stays A1). Inserting a row that shifts content will look like mass cell edits and may over-conflict. Document this; do not build insert-aware alignment in V0.

### 9.4 Conflict resolution UX

For each conflicted cell, reviewer chooses:

- **Keep Main (Ours)**
- **Keep scenario (Theirs)**
- **Edit** (enter value and/or formula; formatting optional)

No combine completes until all conflicts are resolved. Resulting snapshot becomes a new Main version with two parents.

---

## 10. Review flow (V0)

1. Collaborator on a scenario clicks **Ask for review**
2. Owner sees review: summary counts (changed / conflicting) + **What changed** table
3. Owner can:
   - **Combine into Main** (runs merge; may open conflict grid)
   - **Request changes** (simple status + note; scenario remains open)
   - **Close** without combining
4. After successful combine: scenario can remain or be marked combined; Main tip updates; both users see new version on Sync

V0 does **not** require GitHub-style multiple reviewers or CI.

---

## 11. UX surface (V0 web app)

### 11.1 Screens (minimum)

1. **Sign in** — magic link or simple email/password (pick fastest)
2. **Home** — list of projects (few)
3. **Project** — Main tip, scenarios list, version history, members (2)
4. **Scenario** — tip version, Save version (upload), Ask for review
5. **What changed** — filterable cell diff (sheet, address, old → new, formatting deltas)
6. **Review** — diff + Combine CTA
7. **Needs a decision** — conflict grid
8. **Download** — export `.xlsx` for any version

### 11.2 Interaction pattern with Excel

Users edit in **Windows Excel**. Bit is the system of record for versions.

V0 loop:

1. Download working `.xlsx` from current scenario  
2. Edit in Excel  
3. Upload / **Save version** back to Bit  

(Add-in later collapses this to in-Excel buttons.)

### 11.3 Tone

Calm, plain language. No “commit”, “PR”, “HEAD”, “fast-forward”. Show progress and errors in human terms (“Couldn’t read this file — is it `.xlsx`?”).

---

## 12. System architecture

```
┌─────────────────────────────────────────┐
│  Bit Web App (V0 demo UI)               │
│  project · versions · review · merge UI │
└─────────────────┬───────────────────────┘
                  │ HTTPS / JSON
┌─────────────────▼───────────────────────┐
│  Bit API                                │
│  auth · projects · versions · reviews   │
└───────────┬─────────────────┬───────────┘
            │                 │
┌───────────▼──────┐  ┌───────▼────────────┐
│  Merge engine    │  │  Xlsx bridge       │
│  3-way cell merge│  │  parse / write     │
└───────────┬──────┘  └───────┬────────────┘
            │                 │
┌───────────▼─────────────────▼────────────┐
│  Store                                   │
│  metadata DB + content-addressed blobs   │
│  (snapshots, optional xlsx artefacts)    │
└──────────────────────────────────────────┘
```

### 12.1 Suggested stack (default; swappable)

| Layer | Suggestion | Rationale |
|---|---|---|
| Web UI | Next.js or simple React + Vite | Fast demo UI |
| API | Node (Nest/Fastify) or Python (FastAPI) | One language with xlsx libs |
| Xlsx | ExcelJS (Node) or openpyxl (Python) | Values, formulas, basic fmt |
| DB | Postgres | Projects, users, reviews |
| Blobs | S3-compatible or local disk for demo | Snapshots + xlsx |
| Auth | Magic link (Clerk/Auth.js) or basic email | Two-user demo speed |

**Engine purity:** merge logic must be unit-testable with no UI or Excel dependency (pure snapshot in → snapshot/conflicts out).

### 12.2 Git under the hood?

Optional. V0 can use an application object store (versions + parent links) without exposing Git. Using Git as a blob backend is fine if it speeds storage — users never see it.

---

## 13. API sketch (V0)

- `POST /projects` — create; upload initial `.xlsx`
- `GET /projects/:id`
- `POST /projects/:id/invite` — add second user
- `GET /projects/:id/scenarios`
- `POST /projects/:id/scenarios` — create scenario from Main tip
- `POST /scenarios/:id/versions` — upload `.xlsx` or snapshot; message
- `GET /versions/:id` / `GET /versions/:id/xlsx`
- `GET /diff?base=&compare=` — cell-level diff
- `POST /scenarios/:id/reviews` — ask for review → Main
- `POST /reviews/:id/combine` — body: conflict resolutions if needed
- `GET /reviews/:id/conflicts` — preview merge conflicts

---

## 14. Security and trust (V0-sized)

- Auth required; project visible only to its two members
- Signed URLs or auth’d downloads for `.xlsx`
- Audit fields on every version (who / when / message)
- Main protection: no silent overwrite from scenario without Combine
- Do not log full workbook contents in plain app logs

---

## 15. Testing strategy

### 15.1 Engine fixtures (must-have)

- No-op merge  
- Disjoint cell edits → clean combine  
- Same cell different values → conflict  
- Formula vs value change → conflict  
- Formatting-only change on one side → auto-take  
- Delete vs edit → conflict  
- Sheet added on one side → clean take  

### 15.2 Round-trip fixtures

- Sample FP&A-like sheet (assumptions block + simple formulas)
- Assert tracked fields survive parse → write → parse

### 15.3 E2E demo script

Script the Alex/Jordan flow on a staging deploy; treat it as the acceptance test for “V0 done”.

---

## 16. Milestones

| Milestone | Outcome |
|---|---|
| **M1 — Model + bridge** | Parse/write `.xlsx`; snapshot schema; round-trip demo file |
| **M2 — Versions** | Project, Main, Save version, history, download |
| **M3 — Scenarios** | Create scenario, switch, save on scenario |
| **M4 — Diff** | What changed UI (values/formulas/fmt/structure) |
| **M5 — Merge engine** | Three-way rules + conflict objects; unit tests green |
| **M6 — Review + Combine** | Ask for review, conflict grid, Combine into Main |
| **M7 — Two-user E2E** | Invite, sync, full Alex/Jordan path on Windows Excel |

**Exit criterion:** a recorded walkthrough of M7 on a realistic budget-style workbook with at least one deliberate conflict resolved in the Bit UI.

---

## 17. Risks and mitigations

| Risk | Mitigation |
|---|---|
| Formatting round-trip incomplete | Track a small fmt subset; show “formatting may be simplified” in UI |
| Row inserts explode conflicts | Document address-stable merge; guide users to avoid structural inserts in V0 scenarios |
| Upload/download friction vs OneDrive | Accept for web V0; prioritise Excel add-in immediately after E2E proof |
| Scope creep (macros, multi-workbook) | This doc’s non-goals are binding for V0 |
| “Review” feels heavy for two people | Keep one-click Ask for review + Combine; optional note only |

---

## 18. North star (post-V0)

- Excel **add-in** (ribbon: Save version, Scenarios, Ask for review, Sync)
- Mac support
- Richer formatting / validation / named ranges
- Insert-aware merge (anchor rows / keys)
- Larger teams, roles, audit export
- Optional: scenario → new workbook export (“branch into other excels”)

---

## 19. Open items (non-blocking for V0 start)

- Final auth vendor for demo
- Whether Main allows direct Save version or only via Combine (recommend: allow owner direct Save; collaborator uses scenarios)
- Exact formatting subset (start: bold, fill, number format)
- Hosting target for demo (local docker vs single cloud box)

---

## 20. Summary

**Bit V0** is the smallest product that proves: *two Excel people can collaborate on one workbook with versions, a scenario, a review, and a real merge — without being technical.*

Web demo first. Windows `.xlsx`. Values, formulas, structure, basic formatting. Conflicts resolved in Bit. Everything else waits.
