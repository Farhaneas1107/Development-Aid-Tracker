/**
 * ============================================================================
 * NAUE OPPORTUNITY TRACKER — SHARED CLOUD DATABASE CONFIGURATION (FOR GITHUB)
 * ============================================================================
 * When hosting on GitHub Pages, set EITHER Option A (Firebase) OR Option B
 * (Google Sheets Web App URL) below and commit this file to GitHub.
 * Every team member who opens your GitHub Pages link will automatically
 * read and save reviews to the shared cloud database in real time!
 */
window.NAUE_CLOUD_CONFIG = {
  // Choose: "auto" | "google_sheets" | "firebase" | "local"
  provider: "auto",

  // --------------------------------------------------------------------------
  // OPTION B: GOOGLE SHEETS LIVE BACKEND (Recommended if your team uses Google)
  // Paste your deployed Google Apps Script Web App URL (ends with /exec):
  // --------------------------------------------------------------------------
  googleSheetsWebAppUrl: "",

  // --------------------------------------------------------------------------
  // OPTION A: FIREBASE REALTIME DATABASE (Recommended for instant cloud JSON)
  // Paste your Firebase Realtime Database URL (e.g. "https://your-app-default-rtdb.firebaseio.com"):
  // --------------------------------------------------------------------------
  firebaseDatabaseUrl: "",

  // Optional auto-refresh interval in seconds to pull teammates' new notes (0 = off)
  autoPollSeconds: 30
};

