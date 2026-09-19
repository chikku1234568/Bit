# Bit — functionality, how to use, comments

What the product does **today** on `main` (M1–M6, plus an M7 demo script). For file locations see [`REPO-STRUCTURE.md`](./REPO-STRUCTURE.md). Product rules: `BIT-V0-DESIGN.md`.

UI: http://127.0.0.1:5173  
API: http://127.0.0.1:3001

```bash
npm test          # 36 passed (validated after PR #2)
npm run api
npm run dev:web
```

---

## 1. What Bit is

Version control for one Excel workbook. Two people (FP&A + a colleague) can:

1. Keep a safe **Main**
2. Propose work in a named **Scenario**
3. See **What changed**
4. **Ask for review**
5. **Combine into Main**, resolving **Needs a decision** cell by cell
6. **Download .xlsx**

Git words never appear in the UI.

| You say | Bit shows |
|---|---|
| commit | Save version |
| branch | Scenario |
| `main` | Main |
| diff | What changed |
| pull request | Ask for review |
| merge | Combine into Main |
| conflict | Needs a decision |

There is **no login**. Top bar **You are** is the author (`localStorage` `bit-author`). Switch the name to simulate Alex vs Jordan.

---

## 2. What is actually shipped

| Milestone | Status | You can do this |
|---|---|---|
| **M1** Bridge | Done | Parse/write `.xlsx` ↔ snapshot (values, formulas, basic format) |
| **M2** Versions | Done | Create project, Save version on Main, history, download |
| **M3** Scenarios | Done | Create scenario from Main tip, switch, save on that scenario |
| **M4** Diff | Done | What changed table between two versions |
| **M5** Merge engine | Done | Three-way cell merge + preview API (pure, unit-tested) |
| **M6** Review | Done | Ask for review, Needs a decision grid, Combine into Main (two parents) |
| **M7** Two-user E2E | Script only | [`E2E-DEMO.md`](./E2E-DEMO.md) — no email invite |

**Not built:** Excel add-in, real auth, invite emails, Postgres, insert-aware merge, macros / `.xlsm`, charts, named ranges.

**Tracked in a version:** sheet names/order, hidden sheets, tab colour, freeze panes; per-cell `v` / `f` / `fmt` (bold, italic, underline, strike, font name/size, fill, font colour, number format, borders, alignment); column widths, row heights, hidden rows/cols, merges. Theme colours are resolved to `#RRGGBB` when possible.

**Dropped on round-trip:** charts, pivots, images, data validation, conditional formatting, named ranges, tables, comments, VBA / macros.

---

## 3. How to run

From the git root (the folder with `package.json`):

```bash
npm install
cd web
npm install
cd ..
npm run api          # terminal 1
npm run dev:web      # terminal 2
```

Open http://127.0.0.1:5173.

Reset all demo data: delete the `data/` folder (gitignored).

Sample file: `fixtures/sample-budget.xlsx` (Budget + Assumptions, formulas, fills).

Node on this laptop is **20.17**. Stay on **Vite 6**. Vite 8 will not start.

---

## 4. How to use (happy path)

Use **Smoke FY27** (already in `data/` from validation) or create a new project.

### Create a project (Alex)

1. **You are** → `Alex`
2. Home → name `FY27 Budget` → upload `.xlsx` → **Create project**
3. You land on the project. Scenario is **Main**. History has version 1.

### Propose work (Jordan)

1. **You are** → `Jordan`
2. Open the same project
3. **Create scenario** named `Tax update` → it starts at Main’s current tip
4. **Switch scenario** to `Tax update` (URL gets `?scenario=…`)
5. **Download .xlsx**, edit in Excel (change a number), save locally
6. **Save version** — pick that file, note `Updated tax rate`
7. Main tip stays put. Scenario tip advances.

### What changed

On the project page, **What changed**. On a scenario it defaults to Main tip vs scenario tip. Filter by sheet / kind (value, formula, format, sheet add/remove).

### Ask for review → Combine (Alex)

1. Still on `Tax update` → **Ask for review** (disabled on Main)
2. Review page: counts + link to What changed
3. If both sides edited the **same cell** → **Needs a decision**: Keep Main / Keep scenario / Edit
4. **Combine into Main**
5. New Main version with **two parents**. Download `.xlsx`

**Request changes** and **Close** are on the review page (status only; they do not write a version).

### Sync

There is no push/pull. Refresh the page. Everyone hitting this API sees the same `data/` folder.

Step-by-step with a deliberate conflict: [`E2E-DEMO.md`](./E2E-DEMO.md).

---

## 5. What we validated after merging PR #2

| Check | Result |
|---|---|
| `npm test` | **36/36 passed** (xlsx, service/HTTP, diff, merge) |
| API `/health` | `{ ok: true }` |
| UI http://127.0.0.1:5173 | 200, Vite 6 |
| Create project + scenario + save on scenario | Main tip unchanged; scenario parent = Main tip |
| `GET /diff` | B2 value change detected |
| Ask for review + Combine (no conflicts) | Review `combined`; Main tip `parentIds` length 2 |
| Combined snapshot | Budget `B2` **100000 → 125000**; formulas kept |

Existing **FY27 Budget** from the earlier M2 demo is still in `data/` alongside **Smoke FY27**.

---

## 6. Behaviour details worth knowing

**Create scenario** copies the *pointer* to Main’s tip version (same version id, no duplicated blob). The first Save on the scenario creates a new version parented on that shared tip.

**Save version** always uploads a full `.xlsx`. Bit does not patch cells in the browser.

**Review freeze:** opening a review stores Main tip and scenario tip at that moment. Combine uses **current Main tip** as “ours” and the **frozen scenario tip** as “theirs”, three-way against the frozen base. If Alex saved on Main after the review opened, those edits still participate.

**Conflicts** are per cell address (`Budget!B2`). Row inserts look like hundreds of edits — accepted V0 limit (address-stable).

**Combine** writes a new Main version, caches `.xlsx`, sets review to `combined`. You cannot combine twice.

**Author** is whatever **You are** says. No permissions: Jordan can save on Main if they switch to it.

---

## 7. Suggestions and comments

Honest notes after reading and running PR #2. Not blockers for a V0 demo.

### Product / UX

1. **Default conflict choice is Keep Main.** The grid initialises every conflict to `keep-ours`. Easy to Combine without looking. Prefer empty selection and refuse Combine until each row is chosen, or highlight unresolved rows.
2. **Edit in the grid only sets a value**, not a formula (`f: null`). If the conflict is formula-vs-formula, Edit silently turns it into a constant. Add a formula field or say “value only”.
3. **Conflict cells render raw JSON** (`{"v":100000,"f":null,...}`). Fine for us, wrong for FP&A. Show “100,000 vs 125,000” and formula text.
4. **Ask for review has no note field** in the UI (API accepts `note`). One optional box would match the design.
5. **History on a scenario** includes the shared Main tip plus scenario saves — slightly confusing when the first row looks like Main’s message. Label “started from Main version …”.
6. **No “you are looking at Main” warning** when Jordan could Save version on Main. A one-liner (“Changes on Main skip review”) would help.
7. **Two-user is theatre.** Same browser, change the name. For a recorded demo that’s enough; don’t pretend it’s invite.

### Engine / API

8. **`POST /reviews/:id/combine` returns the full merge snapshot** in JSON. Large and unused by the UI (`version` + `review` are enough). Trim the payload.
9. **xlsx download + Node fetch** hit `ECONNRESET` once after Combine (server still logged 200; store was correct). Fastify `reply.send(Buffer)` + keep-alive is a bit sharp. Worth a follow-up if the UI download ever flakes; browser `<a href>` is fine.
10. **`hashSnapshot` uses `JSON.stringify` key order.** Same cells in different key order = different hashes. Harmless for V0; if you ever hash in two code paths, canonicalise keys.
11. **`meta.json` is not atomic.** Two overlapping Save versions can drop one. Single-user demo only.
12. **Merge preview `autoChangeCount`** counts cells that differ from base after auto-merge. Useful. Surface it more clearly on the project page, not only on Review.
13. **Cannot Ask for review from Main** — good. Cannot delete a scenario — fine for V0.
14. **Formula cells with cached `v` vs `null`** are *not* a value change in diff (correct). Merge `cellsEqual` *does* compare `v`, so a formula-only recalc cache mismatch could theoretically conflict. Unlikely with Bit’s own writer (`v` often null on formula cells). If you see phantom conflicts, ignore `v` when `f` matches, same as diff.

### Scope to resist

15. Don’t add Postgres / Clerk / S3 until the Excel add-in is the next bet. Disk store is the right V0 demo.
16. Don’t build insert-aware merge yet. Document “don’t insert rows in a scenario” on the Save version card.
17. Don’t reimplement `src/xlsx` inside merge or the UI.

### Nice next steps (post-V0 demo)

- Record the [`E2E-DEMO.md`](./E2E-DEMO.md) path with one deliberate same-cell conflict.
- Members list on the project (`Alex`, `Jordan`) still stub, no authz.
- “Formatting may be simplified” banner on upload.
- Excel add-in: Save version / Switch scenario / Ask for review — that’s the real distribution story.

---

## 8. Click checklist for you

1. Open http://127.0.0.1:5173
2. Open **Smoke FY27** — Main tip should read **Combined Tax update into Main** (already combined in validation)
3. Open **FY27 Budget** if you want a clean-ish older project, or create a new one
4. Create a scenario, save a version, What changed, Ask for review, Combine
5. To wipe: delete `data/` and restart `npm run api`

If the API or UI is down:

```bash
npm run api
npm run dev:web
```
