# Bit V0 — E2E demo script (Alex / Jordan)

Offline walkthrough of the happy path. Auth is the **You are** stub (no email invite).

Use two browser profiles or just switch the **You are** name between steps.

| Role | You are name |
|------|----------------|
| Owner (FP&A) | `Alex` |
| Collaborator | `Jordan` |

## Prep

```bash
npm install && cd web && npm install && cd ..
npm run api          # :3001
npm run dev:web      # :5173
```

Optional: `rm -rf data` for a clean store.

Sample workbook: `fixtures/sample-budget.xlsx`.

## Click path

### 1. Alex creates the project

1. Set **You are** → `Alex`.
2. Home → **Create project**: name `FY27 Budget`, upload `fixtures/sample-budget.xlsx`.
3. Open the project. Confirm **Main** tip and version history show the initial version.

### 2. Jordan creates a scenario

1. Set **You are** → `Jordan`.
2. Open the same project (refresh Home if needed).
3. **Create scenario** named `Tax update`.
4. **Switch scenario** to `Tax update` (URL `?scenario=<id>`).

### 3. Jordan edits and saves

1. **Download .xlsx** from the scenario tip.
2. Edit in Excel (e.g. change a revenue / tax assumption cell).
3. **Save version** on `Tax update` with note `Updated tax rate`.
4. Confirm Main tip is unchanged; scenario tip advanced.

### 4. What changed

1. Open **What changed** (defaults: Main tip vs scenario tip).
2. Confirm the edited cell appears (value / formula / format as applicable).

### 5. Ask for review

1. Still on `Tax update`, click **Ask for review**.
2. Review page opens: summary counts + link to **What changed**.
3. Base/compare are frozen at open.

### 6. Alex combines

1. Set **You are** → `Alex`.
2. Open the review from the project **Reviews** list (or the review URL).
3. If **Needs a decision** appears: choose **Keep Main** / **Keep scenario** / **Edit** per cell.
4. Click **Combine into Main**.
5. Confirm new Main version with two parents; **Download .xlsx**.

### 7. Sync / history

1. Both users refresh the project (**Sync** = reload).
2. Main tip shows the combined version; history lists authors and notes.

## Deliberate conflict (optional)

1. Alex on Main and Jordan on a scenario both change the **same cell** after a review is opened (or open review, then Alex saves on Main overlapping Jordan’s cells).
2. Combine without resolutions → error (**Needs a decision**).
3. Resolve in the grid → Combine succeeds.

## Out of scope here

- Real email invite / members seats (design M7 invite — stub only).
- Excel add-in, macros, insert-aware merge.

## Acceptance

This script completed on a realistic budget workbook with at least one deliberate conflict resolved in Bit UI = V0 E2E done.
