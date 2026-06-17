# THUNDRR GO — Creator Hub

All-in-one hub for the TikTok GO roster. Login → dashboard + managed roster → tap any creator for contact details and a full day-to-day / month-over-month stats report. Same workflow as the LIVE Creator Manager: VAs drag-and-drop exports into `/data`, GitHub auto-rebuilds on Netlify, the app reads one compiled `data.json`.

**Login:** `ben` / `ben123` (client-side gate, change in `index.html` → `CRED`).

---

## The three files

| File (drop into `/data`) | What it powers | Cadence |
|---|---|---|
| `CreatorAnalysis_ManagedCreators_<start>_<end>.xlsx` | GMV, orders, AOV, views, conversion (authoritative period totals) | Daily |
| `ContentAnalysis_VideoList_<start>_<end>.xlsx` | Per-post history → daily views chart + month-over-month trends + recent posts | Daily |
| `Creator_Roster_Main_<id>.xlsx` | Contact info (email, phone, Discord, manager) + who you manage | Biweekly / as needed |

The **username** is the join key across all three (`Creator ID` = `Creator name` = `What is your TikTok username?`).

### Date ranges & history
Filenames carry the date range — that's how snapshots get keyed. Keep one `CreatorAnalysis` file per month and **GMV month-over-month unlocks automatically** (drop in last month's export and the profile MoM card fills in). The `ContentAnalysis` file already carries post history back several weeks, so daily + MoM *content* trends work from a single upload.

---

## Daily workflow (VA)
1. Export the two analysis files from the TikTok GO backstage portal.
2. Drag both into the repo's `/data` folder on GitHub and commit.
3. Netlify rebuilds in ~1 min. Done — the hub shows fresh numbers.

Biweekly: do the same with the updated `Creator_Roster_Main` export. Use **Sync from roster file** in the app to pull any new creators into your managed roster.

---

## Deploy

**Netlify (recommended — runs the build):**
- New site from this repo. Build command `npm install && node build.js`, publish dir `.` (already set in `netlify.toml`).

**GitHub Pages (no build step):**
- Works too — `data.json` is committed as a fallback. To refresh on Pages, run `node build.js` locally and commit the regenerated `data.json`.

**Local preview:**
```bash
npm install
node build.js          # regenerate data.json from /data
npx serve .            # or any static server (must be http://, not file://)
```

---

## How the roster works
- First open seeds your managed roster with creators that have **both** performance data and a contact (≈your active managed creators). Stored in the browser (`localStorage`).
- **Roster tab:** type a username (or paste several) to add. Tap the ✕ on a card to remove. Filter box searches by name/handle.
- **Import all with data** pulls in every managed creator from the latest export. **Sync from roster file** adds anyone new from the contacts file.
- Creators you add who don't have stats yet show *awaiting data* and fill in on the next upload.

## Files
```
index.html      → the app (single file, light default + dark toggle, mobile-first)
build.js        → compiles /data/*.xlsx → data.json
data.json       → compiled output (committed; regenerated each build)
data/           → drop your exports here
netlify.toml    → build config
```
