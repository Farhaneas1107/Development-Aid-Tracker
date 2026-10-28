# Naue Geosynthetics — International Development Aid Opportunity Tracker

An interactive web portal that transforms `Naue_Project_Opportunity_Tracker.xlsx` into an easy-to-read, searchable, satellite-map-enabled intelligence and multi-reviewer collaboration platform covering **all 31 international development aid programs** (`Development Aid`, `Development Aid_2`, and `Development Aid_3`).

## Key Capabilities
- **All 31 Active Programs (`415` Sub-Tenders · `159` Actionable Civil/Spec Opportunities)** across Batches 1–3.
- **6 Interactive Views**:
  1. **Story Cards**: Scannable executive briefs with relevance badges, Naue product chips, review counts, and latest review previews.
  2. **Comparison Table**: Sortable engineering matrix with quick stage selectors.
  3. **Satellite & Site Map**: Opens directly in **Esri World Imagery + Place Labels** (Zero API key required) with exact GPS pins, Yemen irrigation canal alignments, and 1-click Google Earth links.
  4. **Review Pipeline**: 6-stage drag-and-drop Kanban workflow showing inline review notes, reviewer names, and automatic timestamps.
  5. **💬 All Reviews**: Dedicated hub to inspect every review across all projects — switch between **Grouped by Project** (showing multiple colleagues' reviews stacked under each project) and **Chronological Stream**, or filter by reviewer.
  6. **Naue Product Guide**: Engineering application guide & fit matrix across all 8 Naue geosynthetic product families.
- **Multi-Reviewer Safe & Automatic Timestamps**:
  - Every submitted review gets a unique ID and an **automatic system timestamp** (no manual date entry needed).
  - Multiple people reviewing the same project **never overwrite one another** — all reviews are merged cleanly across Local Server, Google Sheets (Option B), and Firebase (Option A).

## Running Locally
Double-click **`start_website.bat`** (or run `python server.py --open --port 8080`) and open **http://127.0.0.1:8080**.
- Saves reviews automatically to `data/team_reviews.json`.
- Exports a 3-sheet styled `.xlsx` workbook (`Reviewed Opportunity Tracker`, `All Individual Reviews`, and `Full Verbatim Details`).

## Hosting Live on GitHub Pages + Shared Team Reviews
1. Push this repository to GitHub and enable **Settings → Pages → Deploy from branch (`main` / root)**.
2. To share live multi-user reviews across everyone visiting your GitHub Pages link, click **`☁️ Cloud Sync`** in the top header (or edit `js/cloud-config.js` once before pushing):
   - **Option B — Google Sheets Live Backend**: Copy the script from `google_sheets_backend.gs` (or via the button in the `☁️ Cloud Sync` modal) into a Google Sheet's **Extensions → Apps Script**, deploy as a **Web App (`Anyone`)**, and paste the URL into `googleSheetsWebAppUrl` in `js/cloud-config.js`. It automatically maintains both a `Team_Reviews` summary tab and an `All_Reviews_Log` tab using `LockService` so concurrent reviewers never collide.
   - **Option A — Firebase Realtime Database**: Create a free Firebase Realtime Database and paste its URL into `firebaseDatabaseUrl` in `js/cloud-config.js`. Each reviewer's entry is stored atomically under `/naue_reviews/{projectId}/entries/{entryId}`.

## Updating Data from Excel
Upload a revised `.xlsx` file directly inside the website (**⇄ Import / Export → Upload / Sync Updated Excel**), or replace `Naue_Project_Opportunity_Tracker.xlsx` and run:
```bash
python scripts/build_data.py
```
