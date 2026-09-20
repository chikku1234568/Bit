# Bit

Version control for Excel. **V1 preview** — desktop Excel + a local agent + a **project folder** (disk or OneDrive/SharePoint sync).

Collaboration is **the same Bit project, two workbooks**. It is not two people editing one shared `.xlsx`. OneDrive shares the **project folder**. Each person still installs Bit on **their** PC.

| You want | Do this |
|---|---|
| Run Bit on this PC | [Install](#1-install-once) → [Start](#2-start-bit-every-session) → [Sideload](#3-sideload-the-excel-add-in-once-per-machine) |
| Join a teammate’s project | Same install, then [Collaborate](#4-collaborate-onedrivesharepoint) |
| How Fetch / Open works | [How collaboration works](#5-how-collaboration-works) |

Add-in details and sideload troubleshooting: **[docs/EXCEL-ADDIN.md](docs/EXCEL-ADDIN.md)**.

---

## What you need

- **Windows** desktop Excel (Microsoft 365). Not Excel in the browser.
- **Node.js 18+** (20.x is what we run). [nodejs.org](https://nodejs.org/)
- Git, if you clone this repo
- For two people: **OneDrive** (or SharePoint library **Sync**) on both PCs

---

## 1. Install (once)

```bat
git clone https://github.com/chikku1234568/Bit.git
cd Bit
npm install
cd web
npm install
cd ..
```

If the repo is already on disk, skip clone. From the folder that contains `package.json`:

```bat
npm install
cd web
npm install
cd ..
```

---

## 2. Start Bit (every session)

**Leave both processes running** while you use Excel. Closing those windows stops Bit.

**Option A — Windows double-click**

```bat
scripts\start-bit.cmd
```

**Option B — two terminals** (from the repo root):

Terminal 1 — agent:

```bat
npm run api
```

Terminal 2 — add-in page:

```bat
npm run dev:web
```

You should see:

| Process | URL |
|--------|-----|
| Agent | http://127.0.0.1:3001 |
| Add-in (Excel loads this) | http://127.0.0.1:5173/addin.html |
| Browser lab (optional) | http://127.0.0.1:5173 |

Check the agent: open http://127.0.0.1:3001/agent/status — you want `"ok": true`.

The Excel pane talks to the **agent on this PC**. A colleague cannot use your `localhost`. They start Bit on **their** machine.

---

## 3. Sideload the Excel add-in (once per machine)

Bit is **not** on the Office Store. **MY ADD-INS** will look empty. That is normal.

### If you have **Upload My Add-in**

1. Start Bit ([§2](#2-start-bit-every-session)).
2. Excel → **Insert** (or **Home**) → **Add-ins** → **My Add-ins**.
3. **Upload My Add-in** → choose `addin\manifest.xml` in this repo.
4. Home ribbon → **Bit** → task pane. Green **agent** pill = good.

### If you only see MY ADD-INS | STORE (typical on a personal Microsoft account)

Use a **Shared Folder** catalog:

1. Excel → **File** → **Options** → **Trust Center** → **Trust Center Settings** → **Trusted Add-in Catalogs**.
2. Catalog Url = a **network share** that contains `manifest.xml` (not a `C:\` path). On the machine that created the share this repo uses:

   `\\YOUR-PC-NAME\BitAddin`

   First time: right-click `scripts\share-addin-catalog.cmd` → **Run as administrator**, then use `\\%COMPUTERNAME%\BitAddin`.
3. **Add catalog** → tick **Show in Menu** → OK. **Quit Excel fully** and reopen.
4. Home → **Add-ins** → **SHARED FOLDER** (not Store) → **Refresh** → **Bit** → **Add**.

Icons and the pane load from `http://127.0.0.1:5173`, so Bit must already be started.

If the pane says **offline**, start the agent again and reopen **Bit** on the ribbon.

---

## 4. Collaborate (OneDrive/SharePoint)

OneDrive does **not** install Bit. It only syncs the **project folder**. Each person: install Bit → start Bit → sideload → **Choose folder** → **Fetch**.

### Person A (creates the project)

1. In Explorer, create a folder inside OneDrive (or a synced SharePoint library), e.g.

   `C:\Users\<you>\OneDrive\BitProjects`

2. Share that folder (or the library) with Person B (**Can edit**). Wait until OneDrive shows a checkmark, not “processing”.
3. Start Bit. Excel pane → **Choose folder** → that path (or **Open project** / paste path → **Use path**).
4. **You are** → your name.
5. Open *your* `.xlsx` → **Create project**.

You should see `project.json`, `versions\`, and `blobs\` appear in that folder, then sync.

### Person B (joins)

1. OneDrive → **Shared** → **Sync** / **Add shortcut to My files** so the library is a normal folder. Their path will **not** match Person A’s (different `C:\Users\...`).
2. Clone/install Bit, start Bit, sideload ([§1–3](#1-install-once)).
3. Excel pane → **Choose folder** → *their* synced `BitProjects` path.
4. **Fetch**.

They should see Person A’s project and version graph.

### Then they work

| Button | Meaning |
|--------|---------|
| **Fetch** | Reload the graph from disk (after OneDrive sync). Does not open Excel files by itself. |
| **Open** (History / Graph) | New workbook from that **version**. That file is *their* copy. Never writes into the other person’s open Excel. |
| **Save version** | Snapshot the workbook they have **open** into the shared project. |
| **What changed** | Main tip vs the selected version. |
| **Make this Main** | Point Main at that version (whole snapshot). Teammates see it after **Fetch**. |

If both Save on the same Main tip at once: one gets **409** — “Project was updated — Fetch and try again.” Fetch, then retry.

**Do not** both edit the same live `.xlsx` in Excel co-authoring. Two files, one project folder.

---

## 5. How collaboration works

```
OneDrive / SharePoint folder     ←  the Bit project (graph + snapshots)
        ▲
        │  Fetch / Save version / Make this Main
        │
  A's Excel file              B's Excel file
  (A's working copy)          (B's working copy)
```

1. **Fetch** = “show me what’s in the folder now.”
2. Pick a node on the graph → **Open** = “give me a new workbook from that snapshot.”
3. Edit only that workbook → **Save version** = “add my snapshot to the project.”
4. **Make this Main** when the team wants that snapshot as official.

Browser lab (no Excel pane): http://127.0.0.1:5173 after `npm run api` and `npm run dev:web`. Same agent, same folder. Useful if sideload fails.

---

## Tests

```bat
npm test
```

---

## More docs

| Doc | What |
|-----|------|
| [docs/EXCEL-ADDIN.md](docs/EXCEL-ADDIN.md) | Add-in buttons, sideload, SharePoint notes |
| [docs/FUNCTIONALITY.md](docs/FUNCTIONALITY.md) | What is saved vs not, how to use |
| [docs/REPO-STRUCTURE.md](docs/REPO-STRUCTURE.md) | Where code lives |
| [BIT-GUIDE.md](BIT-GUIDE.md) | Runbook |
| [BIT-V0-DESIGN.md](BIT-V0-DESIGN.md) | Product decisions |

### Fidelity (short)

**Kept:** values, formulas, a lot of formatting and layout, hyperlinks, validation, names, comments, tables (for What changed).

**Dropped on Open / rebuild:** charts, pivots, images, macros / `.xlsm`.

### Layout

```
addin/manifest.xml   Excel sideload manifest
src/xlsx/            parse/write snapshot
src/store/           append-only project.json + versions/ + blobs/
src/api/             agent (Fastify :3001)
web/                 Vite UI + add-in task pane (:5173)
scripts/start-bit.cmd
```

The agent remembers the last folder in `%USERPROFILE%\.bit\config.json`. Override with `BIT_DATA_DIR`.
