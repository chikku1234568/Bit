# Bit V1 — current limitations

This is a **desktop preview**, not an AppSource listing and not a hosted cloud service. Read this before putting real work in a project folder.

## Product

| Topic | Reality |
|--------|---------|
| Who it is for | Windows desktop Excel + a local Bit agent. Not Excel on the web. |
| Install | Clone this repo and run locally. Not in the Office Store. **MY ADD-INS** empty is expected. |
| Collab unit | A **Bit project folder**, not one shared `.xlsx`. Two people = two workbooks, same folder. |
| Remote | Disk path or OneDrive / SharePoint **synced** folder. No Microsoft Graph app. |
| Identity | **You are** is a display name. Anyone who can write the folder can Save version or Make this Main. |
| Merge | Add-in collab path is **What changed** + **Make this Main**. Cell-level Combine exists in the browser lab only. |

## What Save / Open keep

Values, formulas, sheet order, a large formatting/layout subset, hyperlinks, validation, named ranges, comments, tables / AutoFilter (for diffs). See README fidelity list.

## What they drop

Charts, pivots, images, shapes, macros / `.xlsm`, conditional formatting rules as first-class merge, print setup, Power Query, external links. **Open** rebuilds `.xlsx` from JSON; it is not your original file.

Inserting rows is not tracked as “insert”; addresses stay `A1` = `A1`.

## Store / collab

- One agent per PC (`localhost:3001`). A colleague cannot use your machine’s Bit.
- OneDrive sync delay: **Fetch** after the client shows a checkmark.
- Two Saves on the same Main tip at once → HTTP **409**; Fetch and retry.
- `project.json` tips use compare-and-swap. Do not hand-edit those files while Bit is running.
- Putting the folder in OneDrive is **not** live co-authoring and **not** a database. Do not also Excel-co-author the same workbook.

## Security

- No login, no encryption at rest beyond whatever OneDrive provides.
- Do not put secrets in the repo. The **project folder** (often outside the git clone) holds workbook snapshots in JSON.
- Treat a shared BitProjects library like a shared drive: anyone with edit access can change Main.

## Not in V1

AppSource, Excel Online, Google Sheets, hosted multi-tenant API, real SSO, insert-aware merge, chart round-trip.
