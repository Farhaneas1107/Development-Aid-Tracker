/**
 * NAUE GEOSYNTHETICS — DEVELOPMENT AID OPPORTUNITY INTELLIGENCE PORTAL
 * Main Application Controller (Dashboard, 5 Views, Map, Kanban, Dossier & Reviews)
 */
(function () {
  const STORAGE_KEY_REVIEWS = "naue_opportunity_reviews_v1";
  const STORAGE_KEY_THEME = "naue_portal_theme";
  const STORAGE_KEY_CLOUD = "naue_portal_cloud_settings_v1";
  const STORAGE_KEY_AUTHOR = "naue_reviewer_name_v1";
  const DEFAULT_GSHEET_URL =
    "https://script.google.com/macros/s/AKfycbxpnNWCpok61AJrLubZUlrnikFeQARHsancCsgA_ojbqQSk9XBHw0uP5lVssgLWrsmW/exec";

  const GSHEET_SCRIPT_TEMPLATE = `var SUMMARY_SHEET = "Team_Reviews";
var LOG_SHEET = "All_Reviews_Log";
function ensureSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s1 = ss.getSheetByName(SUMMARY_SHEET);
  if (!s1) {
    s1 = ss.insertSheet(SUMMARY_SHEET);
    s1.getRange(1,1,1,9).setValues([["Project ID","Current Stage","Starred","Assigned Lead","Review Count","Reviewers","Latest Review Summary","Last Updated (Auto)","Raw Merged JSON"]]).setFontWeight("bold").setBackground("#00549F").setFontColor("#FFFFFF");
    s1.setFrozenRows(1);
  }
  var s2 = ss.getSheetByName(LOG_SHEET);
  if (!s2) {
    s2 = ss.insertSheet(LOG_SHEET);
    s2.getRange(1,1,1,6).setValues([["Timestamp (Auto)","Project ID","Reviewer Name","Reviewer Verdict / Stage","Review Comment / Assessment","Unique Review ID"]]).setFontWeight("bold").setBackground("#059669").setFontColor("#FFFFFF");
    s2.setFrozenRows(1);
  }
  return { summary: s1, log: s2 };
}
function readExistingMap(s1) {
  var lastRow = s1.getLastRow(), map = {};
  if (lastRow >= 2) {
    var data = s1.getRange(2, 1, lastRow - 1, 9).getValues();
    for (var i = 0; i < data.length; i++) {
      var pid = String(data[i][0] || "").trim(), raw = String(data[i][8] || "").trim();
      if (pid && raw) { try { map[pid] = JSON.parse(raw); } catch (e) {} }
    }
  }
  return map;
}
function mergeProjectRecord(ex, inc) {
  if (!ex) return inc; if (!inc) return ex;
  var byId = {}, exL = ex.entries || ex.comments || [], inL = inc.entries || inc.comments || [];
  for (var i = 0; i < exL.length; i++) { var e = exL[i]; byId[e.id || (e.author+"_"+(e.isoTime||i))] = e; }
  for (var j = 0; j < inL.length; j++) { var n = inL[j]; byId[n.id || (n.author+"_"+(n.isoTime||j))] = n; }
  var merged = Object.keys(byId).map(function(k){return byId[k];}).sort(function(a,b){return String(b.isoTime||"").localeCompare(String(a.isoTime||""));});
  var newer = String(inc.updatedAt||"") >= String(ex.updatedAt||"") ? inc : ex;
  return { status: newer.status||ex.status||"unreviewed", starred: Boolean(newer.starred), assignee: newer.assignee||ex.assignee||"", checklist: Object.assign({}, ex.checklist||{}, inc.checklist||{}), updatedAt: newer.updatedAt||new Date().toISOString(), updatedBy: newer.updatedBy||ex.updatedBy||"", entries: merged };
}
function doGet() {
  var sh = ensureSheets();
  return ContentService.createTextOutput(JSON.stringify({ ok: true, provider: "google_sheets", reviews: readExistingMap(sh.summary) })).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  var lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    var sh = ensureSheets(), cur = readExistingMap(sh.summary), inc = (JSON.parse(e.postData.contents||"{}")).reviews || {};
    Object.keys(inc).forEach(function(pid){ cur[pid] = mergeProjectRecord(cur[pid], inc[pid]); });
    var keys = Object.keys(cur).sort(), sumRows = [], logRows = [];
    keys.forEach(function(pid){
      var rec = cur[pid]||{}, entries = rec.entries||[], authors = [];
      entries.forEach(function(en){
        var a = en.author||"Reviewer"; if (authors.indexOf(a)===-1) authors.push(a);
        logRows.push([en.displayTime||en.date||en.isoTime||"", pid, a, en.verdictLabel||en.verdict||rec.status||"", en.text||"", en.id||""]);
      });
      var latest = entries.length ? "["+(entries[0].author||"Reviewer")+"] "+(entries[0].text||entries[0].verdictLabel||"") : "";
      sumRows.push([pid, rec.status||"unreviewed", rec.starred?"Yes":"No", rec.assignee||"", entries.length, authors.join(", "), latest, rec.updatedAt||"", JSON.stringify(rec)]);
    });
    if (sh.summary.getLastRow()>=2) sh.summary.getRange(2,1,sh.summary.getLastRow()-1,9).clearContent();
    if (sumRows.length) sh.summary.getRange(2,1,sumRows.length,9).setValues(sumRows);
    logRows.sort(function(a,b){return String(b[0]).localeCompare(String(a[0]));});
    if (sh.log.getLastRow()>=2) sh.log.getRange(2,1,sh.log.getLastRow()-1,6).clearContent();
    if (logRows.length) sh.log.getRange(2,1,logRows.length,6).setValues(logRows);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, totalReviews: logRows.length })).setMimeType(ContentService.MimeType.JSON);
  } finally { lock.releaseLock(); }
}`;

  const REVIEW_STATUSES = [
    { id: "unreviewed", label: "Unreviewed", color: "#64748b" },
    { id: "pursuing", label: "Pursuing (Priority)", color: "#059669" },
    { id: "consultant_spec", label: "Specify w/ Consultant", color: "#2563eb" },
    { id: "contractor_bid", label: "Target Contractors", color: "#0284c7" },
    { id: "monitoring", label: "Monitoring / Hold", color: "#d97706" },
    { id: "not_applicable", label: "Not Applicable", color: "#94a3b8" }
  ];

  const DEFAULT_CHECKLIST = [
    { id: "verify_status", label: "Verify current tender / ground execution status" },
    { id: "id_consultant", label: "Identify design / supervision engineering consultant" },
    { id: "send_spec", label: "Share Naue technical specification & BoQ text" },
    { id: "id_contractors", label: "Map shortlisted civil works contractors / local partner" }
  ];

  const TIER_META = {
    "Very High": { cls: "rel-very-high", color: "#059669" },
    "High": { cls: "rel-high", color: "#2563eb" },
    "Medium-to-High": { cls: "rel-medhigh-fix", badgeCls: "rel-medium-to-high", color: "#0891b2" },
    "Moderate": { cls: "rel-moderate", color: "#d97706" },
    "Low-to-Medium": { cls: "rel-low-to-medium", color: "#6366f1" },
    "Low": { cls: "rel-low", color: "#64748b" },
    "None": { cls: "rel-none", color: "#94a3b8" }
  };

  const state = {
    payload: window.NAUE_INITIAL_DATA || { projects: [], productsCatalog: [] },
    reviews: {},
    serverConnected: false,
    view: "cards", // cards | table | map | kanban | reviews | products
    reviewsHubLayout: "grouped", // grouped | feed
    reviewsAuthorFilter: "all",
    chartDimension: "country", // country | sector | donor
    dashboardCollapsed: false,
    compareIds: [],
    activeDossierId: null,
    dossierFullscreen: false,
    filters: {
      search: "",
      tier: "all", // all | actionable | top_priority | specific tier
      country: "all",
      product: "all",
      sector: "all",
      channel: "all",
      folder: "all",
      reviewStatus: "all", // all | starred | specific status
      sort: "relevance_desc"
    },
    leafletMap: null,
    leafletLayerGroup: null
  };

  /* ========================================================================
     UTILITIES & FORMATTING
     ======================================================================== */
  function escHtml(str) {
    return String(str ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function highlightSearch(escapedText) {
    const q = state.filters.search.trim();
    if (!q || q.length < 2) return escapedText;
    const safeQ = q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    return escapedText.replace(new RegExp(`(${safeQ})`, "gi"), '<mark class="hl">$1</mark>');
  }

  function formatEngineeringText(raw) {
    let html = escHtml(raw);
    html = html.replace(
      /((?:US\$\s?|EUR\s?|€|\$)\d[\d,\.]*(?:\s?(?:Million|M|k))?(?:\s?[–\-]\s?\$?\d[\d,\.]*(?:M|k)?)?)/g,
      '<strong style="color:var(--tier-vhigh-text)">$1</strong>'
    );
    html = html.replace(
      /(Carbofol®?|Bentofix®?|Secutex®?\s*Soft\s*Rock|Secutex®?|Secugrid®?|Combigrid®?|Secudrain®?|Secumat®?)/gi,
      '<strong style="color:var(--brand-primary)">$1</strong>'
    );
    html = html.replace(/('FILTER CLOTH'|FILTER CLOTH)/g, '<strong style="background:var(--tier-vhigh-bg);color:var(--tier-vhigh-text);padding:1px 5px;border-radius:4px">$1</strong>');
    return highlightSearch(html);
  }

  function getTierBadgeClass(tier) {
    const m = TIER_META[tier];
    return (m && (m.badgeCls || m.cls)) || "rel-moderate";
  }

  function getTierColor(tier) {
    return (TIER_META[tier] && TIER_META[tier].color) || "#64748b";
  }

  function getStatusLabel(statusId) {
    return REVIEW_STATUSES.find(s => s.id === statusId)?.label || statusId || "Unreviewed";
  }

  function makeAutoTimestamp(dateObj = new Date()) {
    const isoTime = dateObj.toISOString();
    const displayTime = dateObj.toLocaleString(undefined, {
      year: "numeric",
      month: "short",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit"
    });
    return { isoTime, displayTime };
  }

  function getSavedReviewerName() {
    try {
      return localStorage.getItem(STORAGE_KEY_AUTHOR) || "";
    } catch (e) {
      return "";
    }
  }

  function setSavedReviewerName(name) {
    try {
      if (name && name.trim()) localStorage.setItem(STORAGE_KEY_AUTHOR, name.trim());
    } catch (e) {}
  }

  function getProductInfo(prodId) {
    return (
      state.payload.productsCatalog.find(p => p.id === prodId) || {
        id: prodId,
        name: prodId + "®",
        color: "#0284c7",
        category: "Geosynthetics",
        shortDesc: ""
      }
    );
  }

  function normalizeReviewRecord(rec) {
    const base = rec && typeof rec === "object" ? rec : {};
    let rawEntries = [];
    if (Array.isArray(base.entries)) {
      rawEntries = base.entries;
    } else if (base.entries && typeof base.entries === "object") {
      // Firebase map format { rev_123: {...}, rev_456: {...} }
      rawEntries = Object.values(base.entries);
    } else if (Array.isArray(base.comments)) {
      rawEntries = base.comments;
    }

    const byId = {};
    rawEntries.forEach((item, i) => {
      if (!item || typeof item !== "object") return;
      const id = item.id || `rev_legacy_${item.author || "user"}_${item.isoTime || item.date || i}`;
      byId[id] = {
        id,
        author: item.author || "Reviewer",
        verdict: item.verdict || base.status || "unreviewed",
        verdictLabel: item.verdictLabel || getStatusLabel(item.verdict || base.status || "unreviewed"),
        text: item.text || "",
        isoTime: item.isoTime || base.updatedAt || "",
        displayTime: item.displayTime || item.date || item.isoTime || "Auto-recorded"
      };
    });

    // Migrate legacy quickNote into an entry if no entries exist yet
    if (base.quickNote && Object.keys(byId).length === 0) {
      const legacyId = "rev_quicknote_legacy";
      byId[legacyId] = {
        id: legacyId,
        author: base.assignee || "Reviewer",
        verdict: base.status || "unreviewed",
        verdictLabel: getStatusLabel(base.status || "unreviewed"),
        text: base.quickNote,
        isoTime: base.updatedAt || "",
        displayTime: base.displayUpdatedAt || "Recorded"
      };
    }

    const entries = Object.values(byId).sort((a, b) => String(b.isoTime || "").localeCompare(String(a.isoTime || "")));

    return {
      status: base.status || "unreviewed",
      starred: Boolean(base.starred),
      assignee: base.assignee || "",
      checklist: base.checklist && typeof base.checklist === "object" ? base.checklist : {},
      updatedAt: base.updatedAt || (entries[0]?.isoTime || null),
      displayUpdatedAt: base.displayUpdatedAt || (entries[0]?.displayTime || ""),
      updatedBy: base.updatedBy || (entries[0]?.author || ""),
      entries
    };
  }

  function mergeReviewsMaps(existingMap, incomingMap) {
    const result = {};
    const keys = new Set([...Object.keys(existingMap || {}), ...Object.keys(incomingMap || {})]);
    keys.forEach(pid => {
      const ex = existingMap && existingMap[pid] ? normalizeReviewRecord(existingMap[pid]) : null;
      const inc = incomingMap && incomingMap[pid] ? normalizeReviewRecord(incomingMap[pid]) : null;
      if (!ex) {
        result[pid] = inc;
        return;
      }
      if (!inc) {
        result[pid] = ex;
        return;
      }
      const entryById = {};
      ex.entries.forEach(e => (entryById[e.id] = e));
      inc.entries.forEach(e => (entryById[e.id] = e));
      const mergedEntries = Object.values(entryById).sort((a, b) =>
        String(b.isoTime || "").localeCompare(String(a.isoTime || ""))
      );
      const newer = String(inc.updatedAt || "") >= String(ex.updatedAt || "") ? inc : ex;
      result[pid] = {
        status: newer.status || ex.status || "unreviewed",
        starred: Boolean(newer.starred),
        assignee: newer.assignee || ex.assignee || "",
        checklist: { ...ex.checklist, ...inc.checklist },
        updatedAt: newer.updatedAt || ex.updatedAt,
        displayUpdatedAt: newer.displayUpdatedAt || ex.displayUpdatedAt,
        updatedBy: newer.updatedBy || ex.updatedBy,
        entries: mergedEntries
      };
    });
    return result;
  }

  function getReview(projectId) {
    state.reviews[projectId] = normalizeReviewRecord(state.reviews[projectId]);
    return state.reviews[projectId];
  }

  function addReviewEntry(projectId, { author, verdict, text, autoStatusChangeOnly = false }) {
    const rev = getReview(projectId);
    const { isoTime, displayTime } = makeAutoTimestamp();
    const cleanAuthor = (author || getSavedReviewerName() || "Team Reviewer").trim();
    if (author && author.trim()) setSavedReviewerName(cleanAuthor);

    const chosenVerdict = verdict || rev.status || "unreviewed";
    rev.status = chosenVerdict;
    rev.updatedAt = isoTime;
    rev.displayUpdatedAt = displayTime;
    rev.updatedBy = cleanAuthor;

    let createdEntry = null;
    if (!autoStatusChangeOnly || (text && text.trim())) {
      createdEntry = {
        id: `rev_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
        author: cleanAuthor,
        verdict: chosenVerdict,
        verdictLabel: getStatusLabel(chosenVerdict),
        text: (text || "").trim() || `Marked project stage as "${getStatusLabel(chosenVerdict)}".`,
        isoTime,
        displayTime
      };
      rev.entries.unshift(createdEntry);
    }

    saveReviews(projectId, createdEntry);
    return createdEntry;
  }

  function getAllReviewsCount() {
    let total = 0;
    state.payload.projects.forEach(p => {
      total += getReview(p.id).entries.length;
    });
    return total;
  }

  function getCloudSettings() {
    const cfg = window.NAUE_CLOUD_CONFIG || {};
    let saved = {};
    try {
      saved = JSON.parse(localStorage.getItem(STORAGE_KEY_CLOUD) || "{}");
    } catch (e) {}
    const hasCustomFirebase = Boolean(saved.firebaseDatabaseUrl && !saved.googleSheetsWebAppUrl);
    const gsheetUrl = hasCustomFirebase
      ? ""
      : (saved.googleSheetsWebAppUrl || cfg.googleSheetsWebAppUrl || DEFAULT_GSHEET_URL).trim();
    const firebaseUrl = (saved.firebaseDatabaseUrl || cfg.firebaseDatabaseUrl || "").trim().replace(/\/+$/, "");
    let activeProvider = "local";
    if (gsheetUrl) activeProvider = "google_sheets";
    else if (firebaseUrl) activeProvider = "firebase";
    return { activeProvider, gsheetUrl, firebaseUrl };
  }

  function updateCloudStatusUI() {
    const { activeProvider, gsheetUrl, firebaseUrl } = getCloudSettings();
    const dot = document.getElementById("cloud-status-dot");
    const label = document.getElementById("cloud-status-label");
    const modalText = document.getElementById("cloud-modal-status-text");
    const gInp = document.getElementById("cloud-gsheet-url");
    const fInp = document.getElementById("cloud-firebase-url");

    if (gInp && gsheetUrl) gInp.value = gsheetUrl;
    if (fInp && firebaseUrl) fInp.value = firebaseUrl;

    if (activeProvider === "google_sheets") {
      if (dot) dot.style.background = "#10b981";
      if (label) label.textContent = "☁️ Google Sheets Live";
      if (modalText) modalText.innerHTML = `<strong style="color:#10b981">● Connected by default to shared Google Sheets Backend (Auto-Sync Active)</strong>`;
    } else if (activeProvider === "firebase") {
      if (dot) dot.style.background = "#10b981";
      if (label) label.textContent = "☁️ Firebase Live";
      if (modalText) modalText.innerHTML = `<strong style="color:#10b981">● Connected to Firebase Multi-Reviewer Database (Option A)</strong>`;
    } else {
      if (dot) dot.style.background = state.serverConnected ? "#38bdf8" : "#94a3b8";
      if (label) label.textContent = state.serverConnected ? "💻 Local Server" : "☁️ Cloud Sync";
      if (modalText) {
        modalText.textContent = state.serverConnected
          ? "Current storage: Local Disk (data/team_reviews.json) + Browser LocalStorage"
          : "Current storage: Browser LocalStorage (Connect Option A or B above for shared GitHub Pages sync)";
      }
    }
  }

  async function pullReviewsFromCloud(silent = true) {
    const { activeProvider, gsheetUrl, firebaseUrl } = getCloudSettings();
    try {
      if (activeProvider === "google_sheets" && gsheetUrl) {
        const res = await fetch(gsheetUrl, { cache: "no-store" });
        const json = await res.json();
        if (json && json.reviews && typeof json.reviews === "object") {
          state.reviews = mergeReviewsMaps(state.reviews, json.reviews);
          localStorage.setItem(STORAGE_KEY_REVIEWS, JSON.stringify(state.reviews));
          renderAll();
          if (!silent) showToast("Merged all team reviews from Google Sheets!");
        }
      } else if (activeProvider === "firebase" && firebaseUrl) {
        const res = await fetch(`${firebaseUrl}/naue_reviews.json`, { cache: "no-store" });
        const json = await res.json();
        if (json && typeof json === "object") {
          state.reviews = mergeReviewsMaps(state.reviews, json);
          localStorage.setItem(STORAGE_KEY_REVIEWS, JSON.stringify(state.reviews));
          renderAll();
          if (!silent) showToast("Merged all team reviews from Firebase!");
        }
      }
    } catch (err) {
      if (!silent) showToast("Cloud sync warning: " + err.message);
    }
  }

  async function pushReviewsToCloud(changedProjectId = null, newEntry = null) {
    const { activeProvider, gsheetUrl, firebaseUrl } = getCloudSettings();
    try {
      if (activeProvider === "google_sheets" && gsheetUrl) {
        let payloadReviews = {};
        if (changedProjectId) {
          payloadReviews[changedProjectId] = getReview(changedProjectId);
        } else {
          for (const [pid, rawRec] of Object.entries(state.reviews)) {
            const rec = normalizeReviewRecord(rawRec);
            const hasActivity =
              rec.status !== "unreviewed" ||
              rec.starred ||
              Boolean(rec.assignee) ||
              rec.entries.length > 0 ||
              Object.values(rec.checklist || {}).some(Boolean);
            if (hasActivity) payloadReviews[pid] = rec;
          }
        }
        if (Object.keys(payloadReviews).length === 0) return;
        // Google Apps Script uses LockService + mergeProjectRecord by unique entry ID
        await fetch(gsheetUrl, {
          method: "POST",
          mode: "no-cors",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: JSON.stringify({ reviews: payloadReviews })
        });
      } else if (activeProvider === "firebase" && firebaseUrl) {
        if (changedProjectId) {
          const rec = getReview(changedProjectId);
          // PATCH metadata without touching other reviewers' entries
          await fetch(`${firebaseUrl}/naue_reviews/${changedProjectId}.json`, {
            method: "PATCH",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              status: rec.status,
              starred: rec.starred,
              assignee: rec.assignee,
              checklist: rec.checklist,
              updatedAt: rec.updatedAt,
              displayUpdatedAt: rec.displayUpdatedAt,
              updatedBy: rec.updatedBy
            })
          });
          // Store individual review entry at its own unique key so concurrent reviewers NEVER collide!
          if (newEntry && newEntry.id) {
            await fetch(`${firebaseUrl}/naue_reviews/${changedProjectId}/entries/${newEntry.id}.json`, {
              method: "PUT",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify(newEntry)
            });
          }
        } else {
          // Initial sync: push each project's entries keyed by ID
          for (const [pid, rawRec] of Object.entries(state.reviews)) {
            const rec = normalizeReviewRecord(rawRec);
            const entriesMap = {};
            rec.entries.forEach(e => (entriesMap[e.id] = e));
            await fetch(`${firebaseUrl}/naue_reviews/${pid}.json`, {
              method: "PATCH",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                status: rec.status,
                starred: rec.starred,
                assignee: rec.assignee,
                checklist: rec.checklist,
                updatedAt: rec.updatedAt,
                displayUpdatedAt: rec.displayUpdatedAt,
                updatedBy: rec.updatedBy
              })
            });
            for (const [eid, eObj] of Object.entries(entriesMap)) {
              await fetch(`${firebaseUrl}/naue_reviews/${pid}/entries/${eid}.json`, {
                method: "PUT",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(eObj)
              });
            }
          }
        }
      }
    } catch (err) {
      console.warn("Cloud push error:", err);
    }
  }

  function saveReviews(changedProjectId = null, newEntry = null) {
    try {
      localStorage.setItem(STORAGE_KEY_REVIEWS, JSON.stringify(state.reviews));
    } catch (e) {
      console.warn("localStorage unavailable:", e);
    }
    if (state.serverConnected) {
      fetch("/api/reviews", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(state.reviews)
      })
        .then(r => r.json())
        .then(json => {
          if (json && json.reviews) {
            state.reviews = mergeReviewsMaps(state.reviews, json.reviews);
          }
        })
        .catch(() => {});
    }
    pushReviewsToCloud(changedProjectId, newEntry);
  }

  function showToast(msg) {
    const el = document.getElementById("app-toast");
    if (!el) return;
    el.textContent = msg;
    el.classList.add("show");
    clearTimeout(showToast._timer);
    showToast._timer = setTimeout(() => el.classList.remove("show"), 2600);
  }

  /* ========================================================================
     FILTERING & SORTING ENGINE
     ======================================================================== */
  function getFilteredProjects() {
    const f = state.filters;
    const q = f.search.trim().toLowerCase();

    const filtered = state.payload.projects.filter(p => {
      const rev = getReview(p.id);

      if (f.tier === "actionable" && p.relevanceRank === 0) return false;
      if (f.tier === "top_priority" && p.relevanceRank < 5) return false;
      if (
        f.tier !== "all" &&
        f.tier !== "actionable" &&
        f.tier !== "top_priority" &&
        p.relevanceTier !== f.tier
      ) {
        return false;
      }

      if (f.country !== "all" && !p.countries.includes(f.country)) return false;
      if (f.product !== "all" && !p.products.includes(f.product)) return false;
      if (f.sector !== "all" && !p.sectors.includes(f.sector)) return false;
      if (f.channel !== "all" && p.salesChannel !== f.channel) return false;
      if (f.folder !== "all" && p.folder !== f.folder) return false;

      if (f.reviewStatus === "starred" && !rev.starred) return false;
      if (f.reviewStatus === "reviewed_any" && rev.status === "unreviewed" && !rev.starred && rev.entries.length === 0) return false;
      if (
        f.reviewStatus !== "all" &&
        f.reviewStatus !== "starred" &&
        f.reviewStatus !== "reviewed_any" &&
        rev.status !== f.reviewStatus
      ) {
        return false;
      }

      if (q) {
        const hay = [
          p.id,
          p.shortName,
          p.title,
          p.countryRaw,
          p.locationRaw,
          p.tenderRefRaw,
          p.procuringOrgRaw,
          p.technicalScopeRaw,
          p.timelineRaw,
          p.relevanceRaw,
          p.products.join(" "),
          rev.assignee,
          rev.entries.map(e => `${e.author} ${e.text}`).join(" ")
        ]
          .join(" ")
          .toLowerCase();
        if (!hay.includes(q)) return false;
      }

      return true;
    });

    filtered.sort((a, b) => {
      switch (f.sort) {
        case "relevance_desc":
          return b.relevanceRank - a.relevanceRank || b.stats.actionable - a.stats.actionable || a.seq - b.seq;
        case "opps_desc":
          return b.stats.actionable - a.stats.actionable || b.relevanceRank - a.relevanceRank;
        case "subprojects_desc":
          return b.stats.total - a.stats.total || b.relevanceRank - a.relevanceRank;
        case "country_asc":
          return a.countryRaw.localeCompare(b.countryRaw) || b.relevanceRank - a.relevanceRank;
        case "seq_asc":
        default:
          return a.seq - b.seq;
      }
    });

    return filtered;
  }

  /* ========================================================================
     POPULATE FILTER SELECTS
     ======================================================================== */
  function populateFilterDropdowns() {
    const projects = state.payload.projects;

    // Countries
    const countryCounts = {};
    projects.forEach(p => {
      p.countries.forEach(c => {
        countryCounts[c] = (countryCounts[c] || 0) + 1;
      });
    });
    const countrySel = document.getElementById("filter-country");
    const sortedCountries = Object.entries(countryCounts).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
    countrySel.innerHTML =
      `<option value="all">All Countries / Regions</option>` +
      sortedCountries.map(([c, n]) => `<option value="${escHtml(c)}">${escHtml(c)} (${n})</option>`).join("");

    // Products
    const prodSel = document.getElementById("filter-product");
    prodSel.innerHTML =
      `<option value="all">All Naue Products</option>` +
      state.payload.productsCatalog
        .map(pr => {
          const cnt = projects.filter(p => p.products.includes(pr.id)).length;
          return `<option value="${escHtml(pr.id)}">${escHtml(pr.name)} (${cnt})</option>`;
        })
        .join("");

    // Sectors
    const sectorCounts = {};
    projects.forEach(p => p.sectors.forEach(s => (sectorCounts[s] = (sectorCounts[s] || 0) + 1)));
    const sectorSel = document.getElementById("filter-sector");
    sectorSel.innerHTML =
      `<option value="all">All Engineering Sectors</option>` +
      Object.entries(sectorCounts)
        .sort((a, b) => b[1] - a[1])
        .map(([s, n]) => `<option value="${escHtml(s)}">${escHtml(s)} (${n})</option>`)
        .join("");

    // Channels
    const chanCounts = {};
    projects.forEach(p => (chanCounts[p.salesChannel] = (chanCounts[p.salesChannel] || 0) + 1));
    const chanSel = document.getElementById("filter-channel");
    chanSel.innerHTML =
      `<option value="all">All Sales Channels</option>` +
      Object.entries(chanCounts)
        .map(([c, n]) => `<option value="${escHtml(c)}">${escHtml(c)} (${n})</option>`)
        .join("");

    // Batches
    const folderCounts = {};
    projects.forEach(p => (folderCounts[p.folder] = (folderCounts[p.folder] || 0) + 1));
    const folderSel = document.getElementById("filter-folder");
    folderSel.innerHTML =
      `<option value="all">All Document Batches (3)</option>` +
      Object.entries(folderCounts)
        .map(([f, n]) => `<option value="${escHtml(f)}">${escHtml(f)} (${n})</option>`)
        .join("");
  }

  function syncFilterControlsUI() {
    document.getElementById("filter-search").value = state.filters.search;
    document.getElementById("search-clear-btn").style.display = state.filters.search ? "block" : "none";
    document.getElementById("filter-tier").value = state.filters.tier;
    document.getElementById("filter-country").value = state.filters.country;
    document.getElementById("filter-product").value = state.filters.product;
    document.getElementById("filter-sector").value = state.filters.sector;
    document.getElementById("filter-channel").value = state.filters.channel;
    document.getElementById("filter-folder").value = state.filters.folder;
    document.getElementById("filter-review").value = state.filters.reviewStatus;
    document.getElementById("filter-sort").value = state.filters.sort;
  }

  /* ========================================================================
     RENDER DASHBOARD (KPIs + INTERACTIVE CHARTS)
     ======================================================================== */
  function renderDashboard(filtered) {
    const all = state.payload.projects;
    const totalPrograms = all.length;
    const topPriorityCount = all.filter(p => p.relevanceRank >= 5).length;
    const vHighCount = all.filter(p => p.relevanceTier === "Very High").length;
    const highCount = all.filter(p => p.relevanceTier === "High").length;

    const totalActionable = filtered.reduce((acc, p) => acc + p.stats.actionable, 0);
    const totalSub = filtered.reduce((acc, p) => acc + p.stats.total, 0);
    const hitRate = totalSub > 0 ? ((totalActionable / totalSub) * 100).toFixed(1) : "0.0";

    const specCount = all.filter(p => p.salesChannel.includes("Specification")).length;
    const directCount = all.filter(p => p.salesChannel.includes("Direct")).length;

    const reviewedCount = all.filter(p => {
      const r = getReview(p.id);
      return r.status !== "unreviewed" || r.entries.length > 0;
    }).length;
    const totalIndividualReviews = getAllReviewsCount();
    const starredCount = all.filter(p => getReview(p.id).starred).length;

    const kpiContainer = document.getElementById("kpi-grid");
    kpiContainer.innerHTML = `
      <div class="kpi-card ${state.filters.tier === 'all' && state.filters.reviewStatus === 'all' && state.view !== 'reviews' ? 'active-kpi' : ''}" data-kpi-action="reset">
        <div class="kpi-label"><span>Total Programs</span><span>Batches 1–3</span></div>
        <div class="kpi-value">${totalPrograms} <small>Active Pipeline</small></div>
        <div class="kpi-sub">All 31 programs pending ground works</div>
      </div>
      <div class="kpi-card ${state.filters.tier === 'top_priority' ? 'active-kpi' : ''}" data-kpi-action="top_priority">
        <div class="kpi-label"><span>Flagship Priority</span><span style="color:var(--tier-vhigh-solid)">● Click to Filter</span></div>
        <div class="kpi-value">${topPriorityCount} <small>${vHighCount} Very High · ${highCount} High</small></div>
        <div class="kpi-sub">Immediate core landfill, canal & road targets</div>
      </div>
      <div class="kpi-card ${state.filters.tier === 'actionable' ? 'active-kpi' : ''}" data-kpi-action="actionable">
        <div class="kpi-label"><span>Actionable Opportunities</span><span>In View</span></div>
        <div class="kpi-value">${totalActionable} <small>of ${totalSub} tenders (${hitRate}%)</small></div>
        <div class="kpi-sub">Civil & specification packages matching Naue</div>
      </div>
      <div class="kpi-card ${state.filters.channel === 'Consultant / BoQ Specification' ? 'active-kpi' : ''}" data-kpi-action="spec_channel">
        <div class="kpi-label"><span>Sales Engagement Mix</span><span>Channels</span></div>
        <div class="kpi-value">${directCount} <small>Direct · ${specCount} Spec</small></div>
        <div class="kpi-sub">Click to filter BoQ consultant specification</div>
      </div>
      <div class="kpi-card ${state.view === 'reviews' || state.filters.reviewStatus === 'reviewed_any' ? 'active-kpi' : ''}" data-kpi-action="reviewed">
        <div class="kpi-label"><span>Team Reviews</span><span>💬 ${totalIndividualReviews} · ★ ${starredCount}</span></div>
        <div class="kpi-value">${reviewedCount} / ${totalPrograms} <small>${Math.round((reviewedCount / Math.max(totalPrograms, 1)) * 100)}%</small></div>
        <div class="kpi-sub">Click to inspect all team reviews</div>
      </div>
    `;

    // Chart 1: Relevance Tiers
    const tierOrder = ["Very High", "High", "Medium-to-High", "Moderate", "Low-to-Medium", "Low", "None"];
    const maxTier = Math.max(...tierOrder.map(t => all.filter(p => p.relevanceTier === t).length), 1);
    document.getElementById("chart-tiers").innerHTML = tierOrder
      .map(tier => {
        const list = all.filter(p => p.relevanceTier === tier);
        const opps = list.reduce((s, p) => s + p.stats.actionable, 0);
        const pct = (list.length / maxTier) * 100;
        const selected = state.filters.tier === tier ? "selected" : "";
        return `
          <div class="bar-row ${selected}" data-filter-tier="${escHtml(tier)}" title="Click to filter by ${escHtml(tier)}">
            <span class="bar-label">${escHtml(tier)}</span>
            <div class="bar-track"><div class="bar-fill" style="width:${pct}%;background:${getTierColor(tier)}"></div></div>
            <span class="bar-count">${list.length} (${opps})</span>
          </div>
        `;
      })
      .join("");

    // Chart 2: Naue Product Line Demand
    const prodStats = state.payload.productsCatalog.map(pr => {
      const matching = all.filter(p => p.products.includes(pr.id));
      const opps = matching.reduce((s, p) => s + p.stats.actionable, 0);
      return { ...pr, count: matching.length, opps };
    }).sort((a, b) => b.count - a.count);
    const maxProd = Math.max(...prodStats.map(x => x.count), 1);

    document.getElementById("chart-products").innerHTML = prodStats
      .map(pr => {
        const pct = (pr.count / maxProd) * 100;
        const selected = state.filters.product === pr.id ? "selected" : "";
        return `
          <div class="bar-row ${selected}" data-filter-product="${escHtml(pr.id)}" title="${escHtml(pr.category)} — Click to filter">
            <span class="bar-label">${escHtml(pr.name)}</span>
            <div class="bar-track"><div class="bar-fill" style="width:${pct}%;background:${pr.color}"></div></div>
            <span class="bar-count">${pr.count} prg</span>
          </div>
        `;
      })
      .join("");

    // Chart 3: Country / Sector / Donor Breakdown
    renderDimensionChart(all);
  }

  function renderDimensionChart(all) {
    const dim = state.chartDimension;
    const counts = {};
    all.forEach(p => {
      const arr = dim === "country" ? p.countries : dim === "sector" ? p.sectors : p.donors;
      arr.forEach(item => {
        if (!counts[item]) counts[item] = { count: 0, opps: 0 };
        counts[item].count += 1;
        counts[item].opps += p.stats.actionable;
      });
    });

    const entries = Object.entries(counts)
      .sort((a, b) => b[1].count - a[1].count || b[1].opps - a[1].opps)
      .slice(0, 7);
    const maxVal = Math.max(...entries.map(e => e[1].count), 1);

    document.getElementById("chart-dimension").innerHTML = entries
      .map(([name, info]) => {
        const pct = (info.count / maxVal) * 100;
        const isSelected =
          (dim === "country" && state.filters.country === name) ||
          (dim === "sector" && state.filters.sector === name);
        return `
          <div class="bar-row ${isSelected ? 'selected' : ''}" data-dim-type="${dim}" data-dim-val="${escHtml(name)}">
            <span class="bar-label" title="${escHtml(name)}">${escHtml(name)}</span>
            <div class="bar-track"><div class="bar-fill" style="width:${pct}%;background:#0284c7"></div></div>
            <span class="bar-count">${info.count} (${info.opps})</span>
          </div>
        `;
      })
      .join("");
  }

  /* ========================================================================
     RENDER ACTIVE FILTER CHIPS & SUMMARY
     ======================================================================== */
  function renderActiveFilterChips(filtered) {
    const totalOpps = filtered.reduce((s, p) => s + p.stats.actionable, 0);
    document.getElementById("results-summary").innerHTML = `
      <span>Showing <strong>${filtered.length}</strong> of <strong>${state.payload.projects.length}</strong> Programs</span>
      <span>·</span>
      <span><strong>${totalOpps}</strong> Actionable Opportunities</span>
    `;

    const chips = [];
    const f = state.filters;
    if (f.search) chips.push({ key: "search", label: `Search: "${f.search}"` });
    if (f.tier !== "all") chips.push({ key: "tier", label: `Relevance: ${f.tier}` });
    if (f.country !== "all") chips.push({ key: "country", label: `Country: ${f.country}` });
    if (f.product !== "all") chips.push({ key: "product", label: `Product: ${f.product}®` });
    if (f.sector !== "all") chips.push({ key: "sector", label: `Sector: ${f.sector}` });
    if (f.channel !== "all") chips.push({ key: "channel", label: `Channel: ${f.channel}` });
    if (f.folder !== "all") chips.push({ key: "folder", label: `Batch: ${f.folder}` });
    if (f.reviewStatus !== "all") chips.push({ key: "reviewStatus", label: `Review: ${f.reviewStatus}` });

    const bar = document.getElementById("active-filters-bar");
    if (chips.length === 0) {
      bar.style.display = "none";
      return;
    }
    bar.style.display = "flex";
    document.getElementById("filter-chips-list").innerHTML = chips
      .map(
        c =>
          `<span class="filter-chip">${escHtml(c.label)} <button data-clear-filter="${c.key}" title="Remove">×</button></span>`
      )
      .join("");
  }

  /* ========================================================================
     VIEW 1: EXECUTIVE STORY CARDS
     ======================================================================== */
  function renderCardsView(projects) {
    const container = document.getElementById("view-container");
    if (projects.length === 0) {
      container.innerHTML = renderEmptyState();
      return;
    }

    container.innerHTML = `<div class="cards-grid">${projects
      .map(p => {
        const rev = getReview(p.id);
        const isCompared = state.compareIds.includes(p.id);
        const digestText = p.whyItMatters || p.scopeLead;

        return `
          <article class="project-card ${rev.starred ? 'card-starred' : ''}" data-project-id="${p.id}">
            <div>
              <div class="card-top-bar">
                <div class="card-meta-left">
                  <span class="id-tag">${escHtml(p.id)}</span>
                  <span class="rel-badge ${getTierBadgeClass(p.relevanceTier)}">${escHtml(p.relevanceTier)}</span>
                  <span class="batch-tag">${escHtml(p.folder)}</span>
                </div>
                <div class="card-quick-actions">
                  <button class="icon-btn ${rev.starred ? 'starred' : ''}" data-action="star" data-id="${p.id}" title="Bookmark / Priority Star">★</button>
                  <button class="icon-btn ${isCompared ? 'compared' : ''}" data-action="compare" data-id="${p.id}" title="Add to Side-by-Side Comparison">⇄</button>
                </div>
              </div>

              <div class="card-country-row" style="margin-top:10px">
                <span>📍 ${highlightSearch(escHtml(p.countryRaw))}</span>
                <span style="color:var(--text-muted);font-weight:500">${escHtml(p.donors.join(" / ") || p.agencyShort)}</span>
              </div>

              <h2 class="card-title" data-action="open-dossier" data-id="${p.id}">
                ${highlightSearch(escHtml(p.shortName))}
              </h2>

              <p class="card-digest" style="margin-top:7px">
                ${highlightSearch(escHtml(digestText))}
              </p>
            </div>

            <div style="display:flex;flex-direction:column;gap:11px">
              <div class="card-metrics-strip">
                <div class="metric-cell">
                  <span>Naue Opps</span>
                  <strong style="color:var(--tier-vhigh-text)">${p.stats.actionable} <span style="font-weight:500;color:var(--text-muted)">/ ${p.stats.total} pkgs</span></strong>
                </div>
                <div class="metric-cell">
                  <span>Channel</span>
                  <strong title="${escHtml(p.salesChannel)}">${escHtml(p.salesChannel.replace("Consultant / BoQ ", "BoQ ").replace(" Civil Works", ""))}</strong>
                </div>
                <div class="metric-cell">
                  <span>Scope / Budget</span>
                  <strong title="${escHtml(p.budget)}">${escHtml(p.budget)}</strong>
                </div>
              </div>

              <div class="card-products">
                ${
                  p.products.length
                    ? p.products
                        .map(pid => {
                          const info = getProductInfo(pid);
                          return `<span class="prod-chip" data-filter-product="${escHtml(pid)}" title="Filter by ${escHtml(info.name)}">
                            <span class="prod-dot" style="background:${info.color}"></span>${escHtml(info.name)}
                          </span>`;
                        })
                        .join("")
                    : `<span style="font-size:11.5px;color:var(--text-muted)">No physical geosynthetic scope</span>`
                }
              </div>

              <div class="card-footer">
                <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap">
                  <select class="status-select-mini" data-status="${rev.status}" data-action="status-change" data-id="${p.id}">
                    ${REVIEW_STATUSES.map(
                      s => `<option value="${s.id}" ${rev.status === s.id ? "selected" : ""}>${escHtml(s.label)}</option>`
                    ).join("")}
                  </select>
                  ${
                    rev.entries.length > 0
                      ? `<button type="button" class="prod-chip" style="border-color:#059669;color:#047857" data-action="open-dossier" data-id="${p.id}" title="Click to read ${rev.entries.length} team review(s)">
                          💬 ${rev.entries.length} Review${rev.entries.length > 1 ? "s" : ""}
                        </button>`
                      : ""
                  }
                </div>
                <button class="btn btn-outline" style="padding:5px 11px;font-size:12px" data-action="open-dossier" data-id="${p.id}">
                  Read Full Details →
                </button>
              </div>
              ${
                rev.entries.length > 0
                  ? `<div class="kanban-review-snippet" style="margin-top:0;cursor:pointer" data-action="open-dossier" data-id="${p.id}">
                      <div class="kanban-review-meta">
                        <strong>👤 ${escHtml(rev.entries[0].author)}</strong>
                        <span>🕒 ${escHtml(rev.entries[0].displayTime)}</span>
                      </div>
                      <div style="white-space:nowrap;overflow:hidden;text-overflow:ellipsis">
                        "${escHtml(rev.entries[0].text)}"
                      </div>
                    </div>`
                  : ""
              }
            </div>
          </article>
        `;
      })
      .join("")}</div>`;
  }

  /* ========================================================================
     VIEW 2: STRUCTURED COMPARISON TABLE
     ======================================================================== */
  function renderTableView(projects) {
    const container = document.getElementById("view-container");
    if (projects.length === 0) {
      container.innerHTML = renderEmptyState();
      return;
    }

    container.innerHTML = `
      <div class="table-wrapper">
        <table class="tracker-table">
          <thead>
            <tr>
              <th>ID</th>
              <th data-sort-col="relevance_desc">Relevance ▾</th>
              <th data-sort-col="country_asc">Country / Batch</th>
              <th>Program & Procuring Agency</th>
              <th data-sort-col="opps_desc">Naue Opps ▾</th>
              <th>Applicable Naue Products</th>
              <th>Sales Channel & Timeline</th>
              <th>Team Review</th>
              <th>Action</th>
            </tr>
          </thead>
          <tbody>
            ${projects
              .map(p => {
                const rev = getReview(p.id);
                return `
                  <tr>
                    <td>
                      <div style="display:flex;align-items:center;gap:5px">
                        <button class="icon-btn ${rev.starred ? 'starred' : ''}" style="width:24px;height:24px;font-size:12px" data-action="star" data-id="${p.id}">★</button>
                        <span class="id-tag">${escHtml(p.id)}</span>
                      </div>
                    </td>
                    <td><span class="rel-badge ${getTierBadgeClass(p.relevanceTier)}">${escHtml(p.relevanceTier)}</span></td>
                    <td>
                      <div style="font-weight:700;color:var(--text-primary)">${highlightSearch(escHtml(p.countryRaw))}</div>
                      <div style="font-size:11px;color:var(--text-muted)">${escHtml(p.folder)}</div>
                    </td>
                    <td style="max-width:340px">
                      <div style="font-weight:700;color:var(--brand-primary);cursor:pointer" data-action="open-dossier" data-id="${p.id}">
                        ${highlightSearch(escHtml(p.shortName))}
                      </div>
                      <div style="font-size:11.5px;color:var(--text-muted);margin-top:2px">${escHtml(p.procuringOrgRaw.slice(0, 95))}${p.procuringOrgRaw.length > 95 ? "…" : ""}</div>
                    </td>
                    <td>
                      <strong style="font-size:14px;color:var(--tier-vhigh-text)">${p.stats.actionable}</strong>
                      <span style="color:var(--text-muted);font-size:12px"> / ${p.stats.total}</span>
                      <div style="font-size:11px;color:var(--text-muted)">${escHtml(p.budget)}</div>
                    </td>
                    <td style="max-width:250px">
                      <div style="display:flex;flex-wrap:wrap;gap:4px">
                        ${p.products
                          .map(pid => `<span class="prod-chip" data-filter-product="${escHtml(pid)}">${escHtml(pid)}®</span>`)
                          .join("") || '<span style="color:var(--text-muted);font-size:11.5px">—</span>'}
                      </div>
                    </td>
                    <td>
                      <div style="font-weight:600;font-size:12px">${escHtml(p.salesChannel)}</div>
                      <div style="font-size:11.5px;color:var(--text-muted)">${escHtml(p.timelineTag)}</div>
                    </td>
                    <td>
                      <select class="status-select-mini" data-status="${rev.status}" data-action="status-change" data-id="${p.id}">
                        ${REVIEW_STATUSES.map(
                          s => `<option value="${s.id}" ${rev.status === s.id ? "selected" : ""}>${escHtml(s.label)}</option>`
                        ).join("")}
                      </select>
                    </td>
                    <td>
                      <button class="btn btn-outline" style="padding:5px 10px;font-size:12px" data-action="open-dossier" data-id="${p.id}">Open</button>
                    </td>
                  </tr>
                `;
              })
              .join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  /* ========================================================================
     VIEW 3: INTERACTIVE GEOGRAPHIC MAP (LEAFLET + CANAL POLYLINES + SATELLITE)
     ======================================================================== */
  function renderMapView(projects) {
    const container = document.getElementById("view-container");
    const mappedProjects = projects.filter(p => p.mapLocations && p.mapLocations.length > 0);
    const totalPins = mappedProjects.reduce((s, p) => s + p.mapLocations.length, 0);

    container.innerHTML = `
      <div class="map-layout">
        <div class="map-sidebar">
          <div class="map-sidebar-header">
            <span>Mapped Programs (${mappedProjects.length}) · ${totalPins} Sites</span>
            <button class="btn btn-outline" style="padding:3px 8px;font-size:11px" id="map-fit-all">Fit All</button>
          </div>
          <div class="map-site-list" id="map-site-list">
            ${mappedProjects
              .map(
                p => `
              <div class="map-site-item" data-map-focus="${p.id}">
                <div style="display:flex;justify-content:space-between;align-items:center;gap:6px;margin-bottom:4px">
                  <span class="id-tag">${escHtml(p.id)}</span>
                  <span class="rel-badge ${getTierBadgeClass(p.relevanceTier)}" style="font-size:10.5px;padding:1px 7px">${escHtml(p.relevanceTier)}</span>
                </div>
                <div style="font-weight:700;font-size:12.5px;color:var(--text-primary)">${escHtml(p.shortName)}</div>
                <div style="font-size:11.5px;color:var(--text-muted);margin-top:2px">📍 ${escHtml(p.countryRaw)} (${p.mapLocations.length} site${p.mapLocations.length > 1 ? "s" : ""})</div>
                <div style="display:flex;flex-wrap:wrap;gap:4px;margin-top:6px">
                  ${p.mapLocations
                    .map(
                      (loc, lIdx) => `
                    <button
                      type="button"
                      class="prod-chip"
                      style="font-size:10.5px;padding:1px 6px"
                      data-site-lat="${loc.lat}"
                      data-site-lng="${loc.lng}"
                      data-site-pid="${p.id}"
                      data-site-idx="${lIdx}"
                      title="Zoom to ${escHtml(loc.name)} on Satellite"
                    >🎯 ${escHtml(loc.name.length > 26 ? loc.name.slice(0, 24) + "…" : loc.name)}</button>
                  `
                    )
                    .join("")}
                </div>
              </div>
            `
              )
              .join("")}
          </div>
        </div>
        <div class="map-canvas-wrap">
          <div class="map-mode-toolbar">
            <button type="button" class="map-mode-btn active" data-map-layer="hybrid">🛰️ Satellite + Labels</button>
            <button type="button" class="map-mode-btn" data-map-layer="satellite">🌍 Pure Satellite</button>
            <button type="button" class="map-mode-btn" data-map-layer="street">🗺️ Street Map</button>
            <button type="button" class="map-mode-btn" data-map-layer="topo">⛰️ Topographic</button>
          </div>
          <div id="leaflet-map"></div>
          <div class="map-legend">
            <div style="font-weight:700;margin-bottom:5px">Geosynthetic Relevance</div>
            <div style="display:flex;align-items:center;gap:6px"><span class="prod-dot" style="background:#059669"></span> Very High</div>
            <div style="display:flex;align-items:center;gap:6px"><span class="prod-dot" style="background:#2563eb"></span> High</div>
            <div style="display:flex;align-items:center;gap:6px"><span class="prod-dot" style="background:#0891b2"></span> Medium-to-High</div>
            <div style="display:flex;align-items:center;gap:6px"><span class="prod-dot" style="background:#d97706"></span> Moderate</div>
            <div style="display:flex;align-items:center;gap:6px"><span class="prod-dot" style="background:#64748b"></span> Low / Precursor</div>
            <div style="margin-top:5px;padding-top:5px;border-top:1px solid var(--border-color);color:var(--text-muted)">
              Click any 🎯 site chip on left to zoom to <strong>High-Res Satellite</strong>
            </div>
          </div>
        </div>
      </div>
    `;

    initLeafletMap(mappedProjects);
  }

  function initLeafletMap(mappedProjects) {
    const mapEl = document.getElementById("leaflet-map");
    if (typeof L === "undefined") {
      mapEl.innerHTML = `
        <div style="padding:40px;text-align:center;color:#e2e8f0">
          <h3>Interactive Map Library Requires Internet Connection</h3>
          <p>All project coordinates and direct Google Maps links remain available inside each Project Dossier.</p>
        </div>`;
      return;
    }

    if (state.leafletMap) {
      try {
        state.leafletMap.remove();
      } catch (e) {}
      state.leafletMap = null;
    }

    // Free, zero-API-key Satellite, Hybrid, Topographic & Street tile providers
    const satelliteImagery = L.tileLayer(
      "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
      {
        attribution: "Satellite &copy; Esri, Maxar, Earthstar Geographics",
        maxZoom: 19
      }
    );
    const placeLabelsOverlay = L.tileLayer(
      "https://services.arcgisonline.com/ArcGIS/rest/services/Reference/World_Boundaries_and_Places/MapServer/tile/{z}/{y}/{x}",
      {
        attribution: "Labels &copy; Esri",
        maxZoom: 19,
        pane: "overlayPane"
      }
    );
    const hybridGroup = L.layerGroup([satelliteImagery, placeLabelsOverlay]);

    const pureSatellite = L.tileLayer(
      "https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
      {
        attribution: "Satellite &copy; Esri, Maxar",
        maxZoom: 19
      }
    );

    const streetTiles = L.tileLayer(
      "https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png",
      {
        attribution: "&copy; OpenStreetMap &copy; CARTO",
        maxZoom: 19
      }
    );

    const topoTiles = L.tileLayer(
      "https://server.arcgisonline.com/ArcGIS/rest/services/World_Topo_Map/MapServer/tile/{z}/{y}/{x}",
      {
        attribution: "Topo &copy; Esri",
        maxZoom: 19
      }
    );

    const map = L.map(mapEl, {
      center: [31.5, 40.0],
      zoom: 5,
      layers: [hybridGroup]
    });

    const layerMap = {
      hybrid: hybridGroup,
      satellite: pureSatellite,
      street: streetTiles,
      topo: topoTiles
    };
    let currentBaseLayer = hybridGroup;

    document.querySelectorAll("[data-map-layer]").forEach(btn => {
      btn.addEventListener("click", () => {
        const key = btn.getAttribute("data-map-layer");
        const nextLayer = layerMap[key];
        if (!nextLayer || nextLayer === currentBaseLayer) return;
        map.removeLayer(currentBaseLayer);
        map.addLayer(nextLayer);
        currentBaseLayer = nextLayer;
        document.querySelectorAll("[data-map-layer]").forEach(b => b.classList.remove("active"));
        btn.classList.add("active");
      });
    });

    const bounds = [];
    const projectBoundsMap = {};
    const siteMarkers = {};

    mappedProjects.forEach(p => {
      const color = getTierColor(p.relevanceTier);
      projectBoundsMap[p.id] = [];

      p.mapLocations.forEach((loc, lIdx) => {
        const pt = [loc.lat, loc.lng];
        bounds.push(pt);
        projectBoundsMap[p.id].push(pt);

        const marker = L.circleMarker(pt, {
          radius: loc.precision === "exact" ? 9 : 8,
          fillColor: color,
          color: "#ffffff",
          weight: 2.5,
          opacity: 1,
          fillOpacity: 0.95
        }).addTo(map);

        siteMarkers[`${p.id}_${lIdx}`] = marker;

        // Draw canal reach polyline if endLat/endLng exist (e.g. Yemen 8 Wadi Tuban canals)
        if (loc.endLat && loc.endLng) {
          const endPt = [loc.endLat, loc.endLng];
          bounds.push(endPt);
          projectBoundsMap[p.id].push(endPt);
          L.polyline([pt, endPt], { color: "#38bdf8", weight: 5, opacity: 0.9 }).addTo(map);
          L.polyline([pt, endPt], { color, weight: 3, opacity: 1 }).addTo(map);
        }

        const precisionBadge =
          loc.precision === "exact"
            ? "Exact Engineering GPS"
            : loc.precision === "approximate"
            ? "Approximate Site Coordinates"
            : "Regional Reference Location";

        const popupHtml = `
          <div style="min-width:240px;font-family:Inter,sans-serif">
            <div style="font-size:11px;font-weight:700;color:${color};text-transform:uppercase">${escHtml(p.id)} · ${escHtml(p.relevanceTier)}</div>
            <div style="font-size:13.5px;font-weight:700;margin:3px 0;color:#0f172a">${escHtml(loc.name)}</div>
            <div style="font-size:11px;color:#0284c7;font-weight:600;margin-bottom:4px">📍 ${loc.lat.toFixed(5)}°N, ${loc.lng.toFixed(5)}°E (${precisionBadge})</div>
            <div style="font-size:12px;color:#475569;margin-bottom:6px">${escHtml(p.shortName)}</div>
            <div style="font-size:11.5px;background:#f1f5f9;padding:5px 7px;border-radius:4px;margin-bottom:8px;color:#1e293b">
              ${escHtml(loc.note || p.salesChannel)}
            </div>
            <div style="display:flex;gap:6px;justify-content:space-between;align-items:center">
              <a href="https://www.google.com/maps?q=${loc.lat},${loc.lng}&t=k" target="_blank" rel="noopener" style="font-size:11.5px;color:#0284c7;font-weight:600">Google Earth ↗</a>
              <button onclick="window.NaueApp.openDossier('${p.id}')" style="background:#00549f;color:#fff;border:none;padding:4px 9px;border-radius:4px;font-size:11.5px;cursor:pointer;font-weight:600">Open Dossier →</button>
            </div>
          </div>
        `;
        marker.bindPopup(popupHtml);
      });
    });

    setTimeout(() => {
      map.invalidateSize();
      if (bounds.length > 0) {
        map.fitBounds(bounds, { padding: [40, 40], maxZoom: 12 });
      }
    }, 100);

    document.getElementById("map-fit-all")?.addEventListener("click", () => {
      map.invalidateSize();
      if (bounds.length > 0) map.fitBounds(bounds, { padding: [40, 40] });
    });

    // Click individual 🎯 site chip to fly straight to high-res satellite view (zoom 15-16) and open popup
    document.querySelectorAll("[data-site-lat]").forEach(btn => {
      btn.addEventListener("click", e => {
        e.stopPropagation();
        const lat = parseFloat(btn.getAttribute("data-site-lat"));
        const lng = parseFloat(btn.getAttribute("data-site-lng"));
        const pid = btn.getAttribute("data-site-pid");
        const idx = btn.getAttribute("data-site-idx");
        map.flyTo([lat, lng], 15, { duration: 0.85 });
        const m = siteMarkers[`${pid}_${idx}`];
        if (m) setTimeout(() => m.openPopup(), 450);
      });
    });

    document.querySelectorAll("[data-map-focus]").forEach(item => {
      item.addEventListener("click", () => {
        const pid = item.getAttribute("data-map-focus");
        document.querySelectorAll(".map-site-item").forEach(el => el.classList.remove("active"));
        item.classList.add("active");
        const pts = projectBoundsMap[pid];
        if (pts && pts.length > 0) {
          map.flyToBounds(pts, { padding: [60, 60], maxZoom: 14, duration: 0.75 });
        }
      });
      item.addEventListener("dblclick", () => {
        openDossier(item.getAttribute("data-map-focus"));
      });
    });

    state.leafletMap = map;
  }

  /* ========================================================================
     VIEW 4: REVIEW PIPELINE KANBAN BOARD (WITH INLINE REVIEW PREVIEW)
     ======================================================================== */
  function renderKanbanView(projects) {
    const container = document.getElementById("view-container");
    container.innerHTML = `
      <div class="kanban-board">
        ${REVIEW_STATUSES.map(col => {
          const items = projects.filter(p => getReview(p.id).status === col.id);
          return `
            <div class="kanban-col" data-kanban-col="${col.id}">
              <div class="kanban-col-header" style="border-bottom-color:${col.color}">
                <span>${escHtml(col.label)}</span>
                <span class="id-tag">${items.length}</span>
              </div>
              <div class="kanban-cards" data-dropzone="${col.id}">
                ${items
                  .map(p => {
                    const rev = getReview(p.id);
                    const uniqueAuthors = [...new Set(rev.entries.map(e => e.author))];
                    return `
                      <div class="kanban-card" draggable="true" data-drag-id="${p.id}">
                        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px">
                          <span class="id-tag">${escHtml(p.id)}</span>
                          <span class="rel-badge ${getTierBadgeClass(p.relevanceTier)}" style="font-size:10px;padding:1px 6px">${escHtml(p.relevanceTier)}</span>
                        </div>
                        <div style="font-weight:700;color:var(--text-primary);cursor:pointer;margin-bottom:4px" data-action="open-dossier" data-id="${p.id}">
                          ${escHtml(p.shortName)}
                        </div>
                        <div style="font-size:11.5px;color:var(--text-muted);margin-bottom:6px">
                          📍 ${escHtml(p.countryRaw)} · <strong>${p.stats.actionable}</strong> opps
                        </div>
                        ${
                          rev.assignee
                            ? `<div style="font-size:11px;color:var(--brand-primary);font-weight:600;margin-bottom:4px">🎯 Lead: ${escHtml(rev.assignee)}</div>`
                            : ""
                        }
                        ${
                          rev.entries.length > 0
                            ? `<div class="kanban-review-snippet" style="cursor:pointer" data-action="open-dossier" data-id="${p.id}">
                                <div class="kanban-review-meta">
                                  <strong>💬 ${rev.entries.length} · ${escHtml(uniqueAuthors.join(", "))}</strong>
                                </div>
                                <div style="display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;margin-bottom:3px">
                                  "${escHtml(rev.entries[0].text)}"
                                </div>
                                <div style="font-size:10px;color:var(--text-muted)">🕒 ${escHtml(rev.entries[0].displayTime)}</div>
                              </div>`
                            : rev.displayUpdatedAt
                            ? `<div style="font-size:10.5px;color:var(--text-muted);margin-top:4px">
                                🕒 ${escHtml(rev.displayUpdatedAt)} · <span style="color:var(--brand-primary);cursor:pointer;font-weight:600" data-action="open-dossier" data-id="${p.id}">+ Add Review Note</span>
                              </div>`
                            : `<div style="margin-top:5px"><span style="font-size:11px;color:var(--brand-primary);cursor:pointer;font-weight:600" data-action="open-dossier" data-id="${p.id}">✎ Write Review →</span></div>`
                        }
                      </div>
                    `;
                  })
                  .join("")}
              </div>
            </div>
          `;
        }).join("")}
      </div>
    `;

    // Wire HTML5 Drag & Drop with Automatic Timestamping
    let draggedId = null;
    container.querySelectorAll("[data-drag-id]").forEach(card => {
      card.addEventListener("dragstart", () => {
        draggedId = card.getAttribute("data-drag-id");
        card.style.opacity = "0.5";
      });
      card.addEventListener("dragend", () => {
        card.style.opacity = "1";
      });
    });

    container.querySelectorAll("[data-dropzone]").forEach(zone => {
      zone.addEventListener("dragover", e => {
        e.preventDefault();
        zone.style.background = "var(--brand-soft)";
      });
      zone.addEventListener("dragleave", () => {
        zone.style.background = "";
      });
      zone.addEventListener("drop", e => {
        e.preventDefault();
        zone.style.background = "";
        const targetStatus = zone.getAttribute("data-dropzone");
        if (draggedId && targetStatus) {
          const rev = getReview(draggedId);
          const { isoTime, displayTime } = makeAutoTimestamp();
          rev.status = targetStatus;
          rev.updatedAt = isoTime;
          rev.displayUpdatedAt = displayTime;
          saveReviews(draggedId, null);
          renderAll();
          showToast(`Moved ${draggedId} to "${getStatusLabel(targetStatus)}" (${displayTime})`);
        }
      });
    });
  }

  /* ========================================================================
     VIEW 4B: ALL TEAM REVIEWS HUB (MULTI-USER LOG & GROUPED ROLLUP)
     ======================================================================== */
  function renderReviewsHubView(projects) {
    const container = document.getElementById("view-container");

    // Collect all individual entries + status-reviewed projects across filtered projects
    const allEntriesFlat = [];
    const authorSet = new Set();

    state.payload.projects.forEach(p => {
      const rev = getReview(p.id);
      rev.entries.forEach(e => {
        authorSet.add(e.author || "Reviewer");
      });
    });

    projects.forEach(p => {
      const rev = getReview(p.id);
      rev.entries.forEach(e => {
        if (state.reviewsAuthorFilter === "all" || e.author === state.reviewsAuthorFilter) {
          allEntriesFlat.push({ project: p, reviewRec: rev, entry: e });
        }
      });
    });

    allEntriesFlat.sort((a, b) => String(b.entry.isoTime || "").localeCompare(String(a.entry.isoTime || "")));

    const reviewedProjects = projects.filter(p => {
      const rev = getReview(p.id);
      if (state.reviewsAuthorFilter !== "all") {
        return rev.entries.some(e => e.author === state.reviewsAuthorFilter);
      }
      return rev.entries.length > 0 || rev.status !== "unreviewed" || rev.starred;
    });

    const authorsList = [...authorSet].sort();

    let bodyHtml = "";
    if (reviewedProjects.length === 0 && allEntriesFlat.length === 0) {
      bodyHtml = `
        <div style="background:var(--bg-surface);border:1px solid var(--border-color);border-radius:12px;padding:44px 24px;text-align:center">
          <h3 style="margin:0 0 8px">No Team Reviews Submitted Yet</h3>
          <p style="color:var(--text-muted);margin:0 0 16px;font-size:13.5px">
            Click any project to open its Dossier, enter your name and assessment in <strong>Section 5</strong>, and click <strong>Submit Review</strong>.<br/>
            Multiple people can review the same project—every person's review is saved separately with an automatic timestamp!
          </p>
          <button class="btn btn-primary" onclick="window.NaueApp.openDossier('NAUE-01')">Review First Project (NAUE-01) →</button>
        </div>
      `;
    } else if (state.reviewsHubLayout === "feed") {
      bodyHtml = `
        <div style="display:flex;flex-direction:column;gap:12px">
          ${allEntriesFlat
            .map(({ project: p, entry: c }) => {
              const initials = (c.author || "R").slice(0, 2);
              return `
                <div class="review-entry-card">
                  <div class="review-entry-header">
                    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
                      <span class="reviewer-avatar">${escHtml(initials)}</span>
                      <strong>${escHtml(c.author)}</strong>
                      <span style="color:var(--text-muted);font-size:12px">reviewed</span>
                      <span class="id-tag">${escHtml(p.id)}</span>
                      <strong style="color:var(--brand-primary);cursor:pointer" data-action="open-dossier" data-id="${p.id}">${escHtml(p.shortName)}</strong>
                      <span class="rel-badge rel-high" style="font-size:11px">${escHtml(c.verdictLabel || getStatusLabel(c.verdict))}</span>
                    </div>
                    <span class="auto-time-badge">🕒 ${escHtml(c.displayTime)}</span>
                  </div>
                  <div style="padding-left:38px;font-size:13.5px;color:var(--text-primary);white-space:pre-line">
                    ${escHtml(c.text)}
                  </div>
                </div>
              `;
            })
            .join("") || '<p style="color:var(--text-muted);text-align:center;padding:24px">No written review notes match this filter yet.</p>'}
        </div>
      `;
    } else {
      // Grouped by Project (shows ALL multiple reviewers side-by-side under each project!)
      bodyHtml = reviewedProjects
        .map(p => {
          const rev = getReview(p.id);
          const matchingEntries =
            state.reviewsAuthorFilter === "all"
              ? rev.entries
              : rev.entries.filter(e => e.author === state.reviewsAuthorFilter);
          const checkedCount = Object.values(rev.checklist || {}).filter(Boolean).length;

          return `
            <div class="project-review-group">
              <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:12px;flex-wrap:wrap;margin-bottom:12px;padding-bottom:10px;border-bottom:1px solid var(--border-color)">
                <div>
                  <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:4px">
                    <span class="id-tag">${escHtml(p.id)}</span>
                    <span class="rel-badge ${getTierBadgeClass(p.relevanceTier)}">${escHtml(p.relevanceTier)}</span>
                    <span style="font-size:12px;font-weight:700;color:var(--brand-accent)">📍 ${escHtml(p.countryRaw)}</span>
                    <span class="id-tag" style="color:var(--tier-vhigh-text)">Current Stage: ${escHtml(getStatusLabel(rev.status))}</span>
                    ${rev.assignee ? `<span class="id-tag">🎯 Lead: ${escHtml(rev.assignee)}</span>` : ""}
                    ${checkedCount > 0 ? `<span class="id-tag">✓ ${checkedCount}/4 Checklist</span>` : ""}
                  </div>
                  <h3 style="margin:4px 0 0;font-size:16px;color:var(--text-primary);cursor:pointer" data-action="open-dossier" data-id="${p.id}">
                    ${escHtml(p.shortName)}
                  </h3>
                </div>
                <button class="btn btn-primary" style="padding:6px 12px;font-size:12px" data-action="open-dossier" data-id="${p.id}">
                  + Add / View Full Dossier (${rev.entries.length} Review${rev.entries.length === 1 ? "" : "s"})
                </button>
              </div>

              ${
                matchingEntries.length > 0
                  ? `<div style="display:flex;flex-direction:column;gap:9px">
                      ${matchingEntries
                        .map(c => {
                          const initials = (c.author || "R").slice(0, 2);
                          return `
                            <div class="review-entry-card" style="background:var(--bg-subtle)">
                              <div class="review-entry-header">
                                <div style="display:flex;align-items:center;gap:9px;flex-wrap:wrap">
                                  <span class="reviewer-avatar">${escHtml(initials)}</span>
                                  <strong style="font-size:13.5px">${escHtml(c.author)}</strong>
                                  <span class="rel-badge rel-medium-to-high" style="font-size:11px;padding:2px 8px">
                                    Verdict: ${escHtml(c.verdictLabel || getStatusLabel(c.verdict))}
                                  </span>
                                </div>
                                <span class="auto-time-badge">🕒 ${escHtml(c.displayTime)}</span>
                              </div>
                              <div style="padding-left:37px;font-size:13.5px;color:var(--text-primary);white-space:pre-line">
                                ${escHtml(c.text)}
                              </div>
                            </div>
                          `;
                        })
                        .join("")}
                    </div>`
                  : `<div style="font-size:12.5px;color:var(--text-muted)">
                      Stage changed to <strong>${escHtml(getStatusLabel(rev.status))}</strong> ${rev.displayUpdatedAt ? `(🕒 ${escHtml(rev.displayUpdatedAt)})` : ""} — no written comment added yet.
                    </div>`
              }
            </div>
          `;
        })
        .join("");
    }

    container.innerHTML = `
      <div class="reviews-hub-toolbar">
        <div>
          <strong style="font-size:15px">💬 Team Reviews Overview</strong>
          <span style="font-size:12.5px;color:var(--text-muted);margin-left:8px">
            <strong>${allEntriesFlat.length}</strong> individual review${allEntriesFlat.length === 1 ? "" : "s"} across <strong>${reviewedProjects.length}</strong> active/reviewed project${reviewedProjects.length === 1 ? "" : "s"} · Multi-reviewer entries never overwrite each other
          </span>
        </div>
        <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
          <select id="reviews-author-filter" class="filter-select" style="width:auto">
            <option value="all">All Reviewers (${authorsList.length})</option>
            ${authorsList
              .map(
                a =>
                  `<option value="${escHtml(a)}" ${state.reviewsAuthorFilter === a ? "selected" : ""}>👤 ${escHtml(a)}</option>`
              )
              .join("")}
          </select>
          <div class="mini-tabs">
            <button type="button" class="mini-tab ${state.reviewsHubLayout === 'grouped' ? 'active' : ''}" data-rev-layout="grouped">Grouped by Project</button>
            <button type="button" class="mini-tab ${state.reviewsHubLayout === 'feed' ? 'active' : ''}" data-rev-layout="feed">Chronological Stream</button>
          </div>
        </div>
      </div>
      ${bodyHtml}
    `;

    document.getElementById("reviews-author-filter")?.addEventListener("change", e => {
      state.reviewsAuthorFilter = e.target.value;
      renderReviewsHubView(getFilteredProjects());
    });

    container.querySelectorAll("[data-rev-layout]").forEach(btn => {
      btn.addEventListener("click", () => {
        state.reviewsHubLayout = btn.getAttribute("data-rev-layout");
        renderReviewsHubView(getFilteredProjects());
      });
    });
  }

  /* ========================================================================
     VIEW 5: NAUE PRODUCT ENGINEERING GUIDE & FIT MATRIX
     ======================================================================== */
  function renderProductsView(projects) {
    const container = document.getElementById("view-container");
    const catalog = state.payload.productsCatalog;

    container.innerHTML = `
      <div class="product-catalog-grid">
        ${catalog
          .map(pr => {
            const matching = state.payload.projects.filter(p => p.products.includes(pr.id));
            const totalOpps = matching.reduce((s, p) => s + p.stats.actionable, 0);
            return `
              <div class="product-guide-card" style="border-top:4px solid ${pr.color}">
                <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px">
                  <div>
                    <h3 style="margin:0;font-size:17px;color:var(--text-primary)">${escHtml(pr.name)}</h3>
                    <div style="font-size:12px;font-weight:600;color:${pr.color}">${escHtml(pr.category)}</div>
                  </div>
                  <span class="id-tag">${matching.length} Programs · ${totalOpps} Opps</span>
                </div>
                <p style="font-size:13px;color:var(--text-secondary);margin:10px 0">${escHtml(pr.shortDesc)}</p>
                <div style="font-size:12px;color:var(--text-muted);margin-bottom:12px">
                  <strong style="color:var(--text-primary)">Key Applications in Portfolio:</strong>
                  <ul style="margin:4px 0 0;padding-left:18px">
                    ${pr.typicalUses.slice(0, 3).map(u => `<li>${escHtml(u)}</li>`).join("")}
                  </ul>
                </div>
                <button class="btn btn-outline" style="width:100%;justify-content:center" data-jump-product="${escHtml(pr.id)}">
                  View All ${matching.length} Matching Programs →
                </button>
              </div>
            `;
          })
          .join("")}
      </div>

      <div class="table-wrapper">
        <table class="tracker-table matrix-table">
          <thead>
            <tr>
              <th>ID</th>
              <th>Program Short Name</th>
              <th>Country</th>
              <th>Relevance</th>
              ${catalog.map(pr => `<th style="text-align:center">${escHtml(pr.name)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${projects
              .map(
                p => `
              <tr>
                <td><span class="id-tag">${escHtml(p.id)}</span></td>
                <td>
                  <strong style="color:var(--brand-primary);cursor:pointer" data-action="open-dossier" data-id="${p.id}">
                    ${escHtml(p.shortName)}
                  </strong>
                </td>
                <td>${escHtml(p.countryRaw)}</td>
                <td><span class="rel-badge ${getTierBadgeClass(p.relevanceTier)}">${escHtml(p.relevanceTier)}</span></td>
                ${catalog
                  .map(pr => {
                    const has = p.products.includes(pr.id);
                    const isExplicit = p.explicitProductMap && p.explicitProductMap[pr.id];
                    if (!has) return `<td class="matrix-hit" style="color:var(--border-strong)">·</td>`;
                    return `<td class="matrix-hit" title="${isExplicit ? 'Explicitly named in document' : 'Applicable system match'}">
                      <span style="color:${pr.color};font-size:16px">${isExplicit ? "●" : "◐"}</span>
                    </td>`;
                  })
                  .join("")}
              </tr>
            `
              )
              .join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  function renderEmptyState() {
    return `
      <div style="background:var(--bg-surface);border:1px solid var(--border-color);border-radius:12px;padding:48px 24px;text-align:center">
        <h3 style="margin:0 0 8px">No Programs Match Your Current Filters</h3>
        <p style="color:var(--text-muted);margin:0 0 16px;font-size:13.5px">Try clearing one or more filters or broadening your search query.</p>
        <button class="btn btn-primary" id="empty-reset-btn">Reset All Filters</button>
      </div>
    `;
  }

  /* ========================================================================
     DEEP-DIVE PROJECT DOSSIER DRAWER
     ======================================================================== */
  function openDossier(projectId) {
    const project = state.payload.projects.find(p => p.id === projectId);
    if (!project) return;
    state.activeDossierId = projectId;

    const filtered = getFilteredProjects();
    const idx = filtered.findIndex(p => p.id === projectId);
    const prevId = idx > 0 ? filtered[idx - 1].id : null;
    const nextId = idx >= 0 && idx < filtered.length - 1 ? filtered[idx + 1].id : null;

    const rev = getReview(projectId);
    const tot = Math.max(project.stats.total, 1);
    const hPct = (project.stats.high / tot) * 100;
    const mPct = (project.stats.medium / tot) * 100;
    const lPct = (project.stats.low / tot) * 100;

    document.getElementById("dossier-nav-pos").textContent =
      idx >= 0 ? `Program ${idx + 1} of ${filtered.length}` : project.id;
    const prevBtn = document.getElementById("dossier-prev-btn");
    const nextBtn = document.getElementById("dossier-next-btn");
    prevBtn.disabled = !prevId;
    nextBtn.disabled = !nextId;
    prevBtn.onclick = () => prevId && openDossier(prevId);
    nextBtn.onclick = () => nextId && openDossier(nextId);

    const bodyEl = document.getElementById("dossier-body");
    bodyEl.innerHTML = `
      <div class="dossier-hero">
        <div style="display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap">
          <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap">
            <span class="id-tag" style="font-size:12.5px">${escHtml(project.id)}</span>
            <span class="rel-badge ${getTierBadgeClass(project.relevanceTier)}" style="font-size:12.5px;padding:4px 12px">
              ${escHtml(project.relevanceTier)} Relevance
            </span>
            <span class="batch-tag">${escHtml(project.folderLabel)} · Excel Row ${project.excelRow}</span>
          </div>
          <button class="btn btn-outline no-print" id="dossier-star-btn">
            ${rev.starred ? "★ Starred Priority" : "☆ Bookmark Priority"}
          </button>
        </div>

        <h2 class="dossier-title">${formatEngineeringText(project.title)}</h2>
        <div style="font-size:13.5px;color:var(--text-secondary)">
          <strong>📍 Country & Location:</strong> ${escHtml(project.countryRaw)} — ${formatEngineeringText(project.locationRaw)}
        </div>

        <div class="ref-chips">
          ${project.tenderCodes
            .map(
              code =>
                `<span class="ref-chip" data-copy-ref="${escHtml(code)}" title="Click to copy reference">📋 ${escHtml(code)}</span>`
            )
            .join("")}
        </div>
      </div>

      <div class="dossier-summary-grid">
        <div class="summary-box">
          <label>Procuring / Funding Agency</label>
          <div>${escHtml(project.donors.join(" / ") || project.agencyShort)}</div>
        </div>
        <div class="summary-box">
          <label>Recommended Sales Channel</label>
          <div>${escHtml(project.salesChannel)}</div>
        </div>
        <div class="summary-box">
          <label>Documented Scope / Budget</label>
          <div>${escHtml(project.budget)}</div>
        </div>
        <div class="summary-box">
          <label>Procurement / Timeline Stage</label>
          <div>${escHtml(project.timelineTag)}</div>
        </div>
      </div>

      <div class="opp-funnel-box">
        <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;font-size:13px">
          <strong>Opportunity Breakdown Across Sub-Projects / Packages</strong>
          <span><strong style="color:var(--tier-vhigh-text)">${project.stats.actionable} Actionable Opportunities</strong> out of ${project.stats.total} Total Documented Packages</span>
        </div>
        <div class="opp-stacked-bar">
          <div style="width:${hPct}%;background:#059669" title="High / Very High: ${project.stats.high}"></div>
          <div style="width:${mPct}%;background:#0284c7" title="Medium / Moderate: ${project.stats.medium}"></div>
          <div style="width:${lPct}%;background:#cbd5e1" title="Low / Non-Civil: ${project.stats.low}"></div>
        </div>
        <div style="display:flex;gap:18px;font-size:12px;color:var(--text-muted);flex-wrap:wrap">
          <span><strong style="color:#059669">● ${project.stats.high}</strong> Very High / High Relevance</span>
          <span><strong style="color:#0284c7">● ${project.stats.medium}</strong> Medium / Moderate Relevance</span>
          <span><strong style="color:#64748b">● ${project.stats.low}</strong> Low / Non-Civil Tenders</span>
        </div>
      </div>

      <!-- SECTION 1: WHAT IS HAPPENING -->
      <section class="dossier-section">
        <h3>1. What Is Happening — Technical Scope & Project Packages</h3>
        <p style="font-size:14px;color:var(--text-primary);margin:0 0 12px;font-weight:500">
          ${formatEngineeringText(project.scopeLead)}
        </p>
        ${
          project.scopeItems.length > 0
            ? `<div class="scope-item-list">
                ${project.scopeItems
                  .map(
                    it => `
                  <div class="scope-item-card">
                    <div class="scope-num">${escHtml(it.num)}</div>
                    <div class="scope-text">
                      ${it.title ? `<strong>${escHtml(it.title)}:</strong> ` : ""}${formatEngineeringText(it.body)}
                    </div>
                  </div>
                `
                  )
                  .join("")}
              </div>`
            : ""
        }
        ${
          project.scopeOppCallout
            ? `<div class="pitch-callout" style="margin-top:12px">
                <strong>Direct Summary Opportunity:</strong> ${formatEngineeringText(project.scopeOppCallout)}
              </div>`
            : ""
        }
      </section>

      <!-- SECTION 2: RELEVANCE TO NAUE & PRODUCT FIT -->
      <section class="dossier-section">
        <h3>2. Why It Matters to Naue — Geosynthetic Fit & Pitch Strategy</h3>
        ${
          project.whyItMatters
            ? `<div class="pitch-callout"><strong>Strategic Rationale:</strong> ${formatEngineeringText(project.whyItMatters)}</div>`
            : `<div class="pitch-callout"><strong>Headline Assessment:</strong> ${formatEngineeringText(project.relevanceHeadline)}</div>`
        }

        ${
          project.relevanceBreakdown.length > 0
            ? `<div style="margin-bottom:14px">
                ${project.relevanceBreakdown
                  .map(
                    b => `
                  <div class="opp-breakdown-item tier-${b.tier}">
                    ${formatEngineeringText(b.text)}
                  </div>
                `
                  )
                  .join("")}
              </div>`
            : ""
        }

        ${
          project.products.length > 0
            ? `
            <div style="font-size:12.5px;font-weight:700;color:var(--text-muted);margin-bottom:8px;text-transform:uppercase">
              Applicable Naue Product Lines (${project.products.length})
            </div>
            <div class="dossier-products-grid">
              ${project.products
                .map(pid => {
                  const info = getProductInfo(pid);
                  return `
                    <div class="dossier-prod-card" style="border-left:4px solid ${info.color}">
                      <div style="font-weight:700;font-size:13.5px;color:var(--text-primary)">${escHtml(info.name)}</div>
                      <div style="font-size:11.5px;font-weight:600;color:${info.color};margin-bottom:4px">${escHtml(info.category)}</div>
                      <div style="font-size:12px;color:var(--text-secondary)">${escHtml(info.shortDesc)}</div>
                    </div>
                  `;
                })
                .join("")}
            </div>
          `
            : ""
        }
      </section>

      <!-- SECTION 3: PROCUREMENT AGENCY, REFERENCES & TIMELINE -->
      <section class="dossier-section">
        <h3>3. Procuring Organisation & Implementation Timeline</h3>
        <div style="background:var(--bg-subtle);border:1px solid var(--border-color);border-radius:8px;padding:12px 15px;margin-bottom:12px;font-size:13.5px">
          <div style="margin-bottom:6px"><strong>Procuring / Financing Institution:</strong> ${formatEngineeringText(project.procuringOrgRaw)}</div>
          <div><strong>Full Tender / Project Reference:</strong> <span style="font-family:var(--font-mono);font-size:12px">${escHtml(project.tenderRefRaw)}</span></div>
        </div>

        <div style="background:var(--bg-surface);border:1px solid var(--border-color);border-radius:8px;padding:12px 15px;font-size:13.5px">
          <div style="font-weight:700;margin-bottom:6px">Documented Schedule: ${formatEngineeringText(project.timelineHeadline)}</div>
          ${
            project.timelineMilestones.length > 0
              ? `<ul style="margin:0;padding-left:18px;color:var(--text-secondary)">
                  ${project.timelineMilestones.map(m => `<li style="margin-bottom:4px">${formatEngineeringText(m)}</li>`).join("")}
                </ul>`
              : `<p style="margin:0;color:var(--text-secondary)">${formatEngineeringText(project.timelineRaw)}</p>`
          }
          <div style="margin-top:10px;padding-top:8px;border-top:1px dashed var(--border-color);font-size:12px;color:var(--tier-vhigh-text)">
            <strong>● Execution Note:</strong> Physical ground execution across earlier-dated development aid packages remains unstarted/pending — active candidate for engineering verification, consultant specification, and contractor tracking.
          </div>
        </div>
      </section>

      <!-- SECTION 4: SITE LOCATIONS & SATELLITE COORDINATES -->
      <section class="dossier-section">
        <h3>4. Site Coordinates & Satellite Inspection Links</h3>
        <div style="font-size:12.5px;color:var(--text-muted);margin-bottom:8px;white-space:pre-line">${escHtml(project.coordinatesRaw)}</div>
        ${
          project.mapLocations.length > 0
            ? `
            <table class="coord-table">
              <thead>
                <tr>
                  <th>Site / Intervention Name</th>
                  <th>Coordinates</th>
                  <th>Engineering Note</th>
                  <th class="no-print">Satellite</th>
                </tr>
              </thead>
              <tbody>
                ${project.mapLocations
                  .map(
                    loc => `
                  <tr>
                    <td><strong>${escHtml(loc.name)}</strong></td>
                    <td style="font-family:var(--font-mono);font-size:11.5px">
                      ${loc.lat.toFixed(5)}°N, ${loc.lng.toFixed(5)}°E
                      <span style="color:var(--text-muted)">(${escHtml(loc.precision)})</span>
                    </td>
                    <td>${escHtml(loc.note || "")}</td>
                    <td class="no-print">
                      <a href="https://www.google.com/maps?q=${loc.lat},${loc.lng}&t=k" target="_blank" rel="noopener" style="color:var(--brand-primary);font-weight:600;text-decoration:none">
                        Google Earth / Maps ↗
                      </a>
                    </td>
                  </tr>
                `
                  )
                  .join("")}
              </tbody>
            </table>
          `
            : ""
        }
      </section>

      <!-- SECTION 5: MULTI-REVIEWER TEAM EVALUATION & ACTION WORKFLOW -->
      <section class="dossier-section">
        <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:10px">
          <h3 style="margin:0;border:none;padding:0">5. Team Reviews, Action Checklist & Multi-User Log (${rev.entries.length} Review${rev.entries.length === 1 ? "" : "s"})</h3>
          ${
            rev.displayUpdatedAt
              ? `<span class="auto-time-badge">🕒 Last activity: ${escHtml(rev.displayUpdatedAt)}${rev.updatedBy ? ` · ${escHtml(rev.updatedBy)}` : ""}</span>`
              : `<span class="auto-time-badge">🕒 Timestamps recorded automatically</span>`
          }
        </div>

        <div class="review-panel">
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px">
            <div>
              <label style="font-size:11.5px;font-weight:700;display:block;margin-bottom:4px">Overall Project Pursuit Stage</label>
              <select id="dossier-status-sel" class="filter-select">
                ${REVIEW_STATUSES.map(
                  s => `<option value="${s.id}" ${rev.status === s.id ? "selected" : ""}>${escHtml(s.label)}</option>`
                ).join("")}
              </select>
            </div>
            <div>
              <label style="font-size:11.5px;font-weight:700;display:block;margin-bottom:4px">Assigned Regional / Technical Lead</label>
              <input type="text" id="dossier-assignee-inp" class="filter-select" placeholder="e.g., MENA Sales / Landfill Spec Team" value="${escHtml(rev.assignee)}" />
            </div>
          </div>

          <div style="font-size:11.5px;font-weight:700;margin-bottom:4px">Shared Action Checklist</div>
          <div class="checklist-grid">
            ${DEFAULT_CHECKLIST.map(
              item => `
              <label class="check-item">
                <input type="checkbox" data-check-id="${item.id}" ${rev.checklist && rev.checklist[item.id] ? "checked" : ""} />
                <span>${escHtml(item.label)}</span>
              </label>
            `
            ).join("")}
          </div>

          <!-- ADD INDIVIDUAL REVIEW BOX (MULTIPLE PEOPLE SUPPORTED, NEVER OVERWRITES) -->
          <div class="no-print" style="margin-top:14px;background:var(--bg-surface);border:1px solid var(--border-strong);border-radius:10px;padding:14px">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:10px">
              <strong style="font-size:13px;color:var(--text-primary)">✍️ Submit Your Individual Review</strong>
              <span style="font-size:11.5px;color:var(--tier-vhigh-text);font-weight:600">
                ✓ Multi-Reviewer Safe — Each person's review is saved separately
              </span>
            </div>

            <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:10px">
              <div>
                <label style="font-size:11px;font-weight:700;color:var(--text-muted);display:block;margin-bottom:3px">Reviewer Name</label>
                <input type="text" id="dossier-comment-author" class="filter-select" placeholder="Your Name (e.g., Salam)" value="${escHtml(getSavedReviewerName())}" />
              </div>
              <div>
                <label style="font-size:11px;font-weight:700;color:var(--text-muted);display:block;margin-bottom:3px">Your Verdict / Stage</label>
                <select id="dossier-comment-verdict" class="filter-select">
                  ${REVIEW_STATUSES.map(
                    s =>
                      `<option value="${s.id}" ${(rev.status !== "unreviewed" ? rev.status : "pursuing") === s.id ? "selected" : ""}>${escHtml(s.label)}</option>`
                  ).join("")}
                </select>
              </div>
            </div>

            <div style="margin-bottom:10px">
              <label style="font-size:11px;font-weight:700;color:var(--text-muted);display:block;margin-bottom:3px">Review Assessment, Partner Info & Next Steps</label>
              <textarea id="dossier-comment-text" rows="3" class="filter-select" placeholder="Write your review here (technical fit, consultant / contractor details, priority notes)..."></textarea>
            </div>

            <div style="display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap">
              <span class="auto-time-badge">🕒 Automatic Timestamp — Date & exact time stamped automatically on submit</span>
              <button class="btn btn-primary" id="dossier-add-comment">✓ Submit Review</button>
            </div>
          </div>

          <!-- ALL SUBMITTED REVIEWS FOR THIS PROJECT -->
          <div style="margin-top:16px">
            <div style="display:flex;justify-content:space-between;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:8px">
              <strong style="font-size:12.5px;color:var(--text-secondary);text-transform:uppercase;letter-spacing:0.03em">
                💬 Submitted Reviews for ${escHtml(project.id)} (${rev.entries.length})
              </strong>
              <button type="button" class="btn btn-outline no-print" id="dossier-open-all-reviews-btn" style="padding:4px 10px;font-size:11.5px">
                Open All Team Reviews Tab →
              </button>
            </div>

            <div class="comments-list" id="dossier-comments-list" style="margin-top:0">
              ${
                rev.entries.length > 0
                  ? rev.entries
                      .map(c => {
                        const initials = (c.author || "R").slice(0, 2);
                        return `
                          <div class="review-entry-card">
                            <div class="review-entry-header">
                              <div style="display:flex;align-items:center;gap:9px;flex-wrap:wrap">
                                <span class="reviewer-avatar">${escHtml(initials)}</span>
                                <strong style="font-size:13.5px;color:var(--text-primary)">${escHtml(c.author || "Reviewer")}</strong>
                                <span class="rel-badge rel-medium-to-high" style="font-size:11px;padding:2px 8px">
                                  ${escHtml(c.verdictLabel || getStatusLabel(c.verdict))}
                                </span>
                              </div>
                              <span class="auto-time-badge">🕒 ${escHtml(c.displayTime)}</span>
                            </div>
                            <div style="padding-left:37px;font-size:13.5px;color:var(--text-primary);white-space:pre-line">
                              ${escHtml(c.text)}
                            </div>
                          </div>
                        `;
                      })
                      .join("")
                  : `<div style="padding:14px;border:1px dashed var(--border-strong);border-radius:8px;font-size:12.5px;color:var(--text-muted);text-align:center">
                      No individual reviews submitted for this project yet. Submit your assessment above — multiple colleagues can review this same project without overwriting each other.
                    </div>`
              }
            </div>
          </div>
        </div>
      </section>

      <!-- SECTION 6: VERBATIM SOURCE RECORD -->
      <details class="no-print" style="margin-top:16px;border:1px solid var(--border-color);border-radius:8px;padding:10px 14px;background:var(--bg-subtle)">
        <summary style="cursor:pointer;font-weight:700;font-size:12.5px;color:var(--text-muted)">
          View Unedited Verbatim Spreadsheet Text (100% Original Document Fidelity)
        </summary>
        <div style="margin-top:12px;font-size:12px;white-space:pre-wrap;color:var(--text-secondary);line-height:1.5">
<strong>[Technical Scope]</strong>
${escHtml(project.technicalScopeRaw)}

<strong>[Procurement & Implementation Timeline]</strong>
${escHtml(project.timelineRaw)}

<strong>[Relevance to Naue]</strong>
${escHtml(project.relevanceRaw)}
        </div>
      </details>
    `;

    // Bind Dossier interactive review controls (with automatic system timestamps)
    document.getElementById("dossier-star-btn").onclick = () => {
      rev.starred = !rev.starred;
      const { isoTime, displayTime } = makeAutoTimestamp();
      rev.updatedAt = isoTime;
      rev.displayUpdatedAt = displayTime;
      saveReviews(projectId, null);
      openDossier(projectId);
      renderAll();
    };

    document.getElementById("dossier-status-sel").onchange = e => {
      const { isoTime, displayTime } = makeAutoTimestamp();
      rev.status = e.target.value;
      rev.updatedAt = isoTime;
      rev.displayUpdatedAt = displayTime;
      saveReviews(projectId, null);
      openDossier(projectId);
      renderAll();
      showToast(`Stage updated to "${getStatusLabel(rev.status)}" (${displayTime})`);
    };

    document.getElementById("dossier-assignee-inp").onchange = e => {
      const { isoTime, displayTime } = makeAutoTimestamp();
      rev.assignee = e.target.value.trim();
      rev.updatedAt = isoTime;
      rev.displayUpdatedAt = displayTime;
      saveReviews(projectId, null);
      renderAll();
      showToast("Assigned lead saved");
    };

    bodyEl.querySelectorAll("[data-check-id]").forEach(cb => {
      cb.onchange = () => {
        if (!rev.checklist) rev.checklist = {};
        rev.checklist[cb.getAttribute("data-check-id")] = cb.checked;
        const { isoTime, displayTime } = makeAutoTimestamp();
        rev.updatedAt = isoTime;
        rev.displayUpdatedAt = displayTime;
        saveReviews(projectId, null);
        renderAll();
      };
    });

    document.getElementById("dossier-add-comment").onclick = () => {
      const authorInp = document.getElementById("dossier-comment-author");
      const verdictSel = document.getElementById("dossier-comment-verdict");
      const textInp = document.getElementById("dossier-comment-text");
      const txt = textInp.value.trim();
      if (!txt) {
        showToast("Please write a short review note before clicking Submit Review");
        textInp.focus();
        return;
      }
      const entry = addReviewEntry(projectId, {
        author: authorInp.value,
        verdict: verdictSel.value,
        text: txt
      });
      openDossier(projectId);
      renderAll();
      showToast(`Saved review by ${entry.author} (${entry.displayTime})`);
    };

    document.getElementById("dossier-open-all-reviews-btn").onclick = () => {
      closeDossier();
      state.view = "reviews";
      renderAll();
    };

    bodyEl.querySelectorAll("[data-copy-ref]").forEach(chip => {
      chip.onclick = () => {
        const txt = chip.getAttribute("data-copy-ref");
        navigator.clipboard?.writeText(txt);
        showToast(`Copied "${txt}"`);
      };
    });

    document.getElementById("dossier-backdrop").classList.add("open");
    document.getElementById("dossier-drawer").classList.add("open");
  }

  function closeDossier() {
    state.activeDossierId = null;
    document.getElementById("dossier-backdrop").classList.remove("open");
    document.getElementById("dossier-drawer").classList.remove("open");
  }

  /* ========================================================================
     SIDE-BY-SIDE COMPARISON MODAL
     ======================================================================== */
  function updateCompareBar() {
    const bar = document.getElementById("compare-floating-bar");
    const n = state.compareIds.length;
    if (n === 0) {
      bar.classList.remove("visible");
      return;
    }
    bar.classList.add("visible");
    document.getElementById("compare-count-text").textContent = `${n} Program${n > 1 ? "s" : ""} Selected for Comparison`;
  }

  function openCompareModal() {
    const items = state.compareIds
      .map(id => state.payload.projects.find(p => p.id === id))
      .filter(Boolean);
    if (items.length === 0) return;

    const modal = document.getElementById("compare-modal");
    const content = document.getElementById("compare-modal-body");

    content.innerHTML = `
      <div class="table-wrapper">
        <table class="tracker-table">
          <thead>
            <tr>
              <th style="width:160px">Attribute</th>
              ${items.map(p => `<th>${escHtml(p.id)} · ${escHtml(p.shortName)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><strong>Naue Relevance</strong></td>
              ${items.map(p => `<td><span class="rel-badge ${getTierBadgeClass(p.relevanceTier)}">${escHtml(p.relevanceTier)}</span></td>`).join("")}
            </tr>
            <tr>
              <td><strong>Country & Location</strong></td>
              ${items.map(p => `<td><strong>${escHtml(p.countryRaw)}</strong><br/><span style="font-size:12px;color:var(--text-muted)">${escHtml(p.locationRaw.slice(0, 120))}</span></td>`).join("")}
            </tr>
            <tr>
              <td><strong>Actionable Opps</strong></td>
              ${items.map(p => `<td><strong style="color:var(--tier-vhigh-text);font-size:15px">${p.stats.actionable}</strong> / ${p.stats.total} sub-tenders</td>`).join("")}
            </tr>
            <tr>
              <td><strong>Sales Channel</strong></td>
              ${items.map(p => `<td>${escHtml(p.salesChannel)}</td>`).join("")}
            </tr>
            <tr>
              <td><strong>Budget / Scale</strong></td>
              ${items.map(p => `<td>${escHtml(p.budget)}</td>`).join("")}
            </tr>
            <tr>
              <td><strong>Naue Products</strong></td>
              ${items.map(p => `<td>${p.products.map(x => `<span class="prod-chip" style="margin:2px">${escHtml(x)}®</span>`).join("") || "None"}</td>`).join("")}
            </tr>
            <tr>
              <td><strong>Core Pitch / Scope</strong></td>
              ${items.map(p => `<td style="font-size:12px">${escHtml((p.whyItMatters || p.scopeLead).slice(0, 260))}</td>`).join("")}
            </tr>
            <tr>
              <td><strong>Action</strong></td>
              ${items.map(p => `<td><button class="btn btn-primary" onclick="document.getElementById('compare-modal').classList.remove('open');window.NaueApp.openDossier('${p.id}')">Open Full Dossier</button></td>`).join("")}
            </tr>
          </tbody>
        </table>
      </div>
    `;
    modal.classList.add("open");
  }

  /* ========================================================================
     MAIN RENDER ORCHESTRATOR
     ======================================================================== */
  function renderAll() {
    syncFilterControlsUI();
    const filtered = getFilteredProjects();
    renderDashboard(filtered);
    renderActiveFilterChips(filtered);

    const revCountBadge = document.getElementById("all-reviews-tab-count");
    if (revCountBadge) {
      revCountBadge.textContent = String(getAllReviewsCount());
    }

    document.querySelectorAll(".view-tab").forEach(btn => {
      btn.classList.toggle("active", btn.getAttribute("data-view") === state.view);
    });

    if (state.view === "cards") renderCardsView(filtered);
    else if (state.view === "table") renderTableView(filtered);
    else if (state.view === "map") renderMapView(filtered);
    else if (state.view === "kanban") renderKanbanView(filtered);
    else if (state.view === "reviews") renderReviewsHubView(filtered);
    else if (state.view === "products") renderProductsView(filtered);

    updateCompareBar();
  }

  function resetFilters() {
    state.filters = {
      search: "",
      tier: "all",
      country: "all",
      product: "all",
      sector: "all",
      channel: "all",
      folder: "all",
      reviewStatus: "all",
      sort: "relevance_desc"
    };
    renderAll();
  }

  /* ========================================================================
     EVENT DELEGATION & INITIALIZATION
     ======================================================================== */
  function bindEvents() {
    // Theme toggle
    document.getElementById("theme-toggle-btn").addEventListener("click", () => {
      const html = document.documentElement;
      const next = html.getAttribute("data-theme") === "dark" ? "light" : "dark";
      html.setAttribute("data-theme", next);
      localStorage.setItem(STORAGE_KEY_THEME, next);
    });

    // Dashboard collapse toggle
    document.getElementById("toggle-dashboard-btn").addEventListener("click", e => {
      state.dashboardCollapsed = !state.dashboardCollapsed;
      document.getElementById("dashboard-section").style.display = state.dashboardCollapsed ? "none" : "block";
      e.currentTarget.textContent = state.dashboardCollapsed ? "Show Analytics" : "Hide Analytics";
    });

    // KPI clicks
    document.getElementById("kpi-grid").addEventListener("click", e => {
      const card = e.target.closest("[data-kpi-action]");
      if (!card) return;
      const act = card.getAttribute("data-kpi-action");
      if (act === "reset") resetFilters();
      else if (act === "top_priority") {
        state.filters.tier = state.filters.tier === "top_priority" ? "all" : "top_priority";
        renderAll();
      } else if (act === "actionable") {
        state.filters.tier = state.filters.tier === "actionable" ? "all" : "actionable";
        renderAll();
      } else if (act === "spec_channel") {
        state.filters.channel =
          state.filters.channel === "Consultant / BoQ Specification" ? "all" : "Consultant / BoQ Specification";
        renderAll();
      } else if (act === "reviewed") {
        if (state.view === "reviews") {
          state.view = "cards";
          state.filters.reviewStatus = "all";
        } else {
          state.view = "reviews";
        }
        renderAll();
      }
    });

    // Analytics chart clicks
    document.getElementById("dashboard-section").addEventListener("click", e => {
      const tierRow = e.target.closest("[data-filter-tier]");
      if (tierRow) {
        const val = tierRow.getAttribute("data-filter-tier");
        state.filters.tier = state.filters.tier === val ? "all" : val;
        renderAll();
        return;
      }
      const prodRow = e.target.closest("[data-filter-product]");
      if (prodRow) {
        const val = prodRow.getAttribute("data-filter-product");
        state.filters.product = state.filters.product === val ? "all" : val;
        renderAll();
        return;
      }
      const dimRow = e.target.closest("[data-dim-type]");
      if (dimRow) {
        const type = dimRow.getAttribute("data-dim-type");
        const val = dimRow.getAttribute("data-dim-val");
        if (type === "country") state.filters.country = state.filters.country === val ? "all" : val;
        else if (type === "sector") state.filters.sector = state.filters.sector === val ? "all" : val;
        else if (type === "donor") state.filters.search = state.filters.search === val ? "" : val;
        renderAll();
        return;
      }
      const miniTab = e.target.closest("[data-chart-dim]");
      if (miniTab) {
        state.chartDimension = miniTab.getAttribute("data-chart-dim");
        document.querySelectorAll("[data-chart-dim]").forEach(b => b.classList.remove("active"));
        miniTab.classList.add("active");
        renderDimensionChart(state.payload.projects);
      }
    });

    // View tabs
    document.querySelectorAll(".view-tab").forEach(btn => {
      btn.addEventListener("click", () => {
        state.view = btn.getAttribute("data-view");
        renderAll();
      });
    });

    // Filter controls
    document.getElementById("filter-search").addEventListener("input", e => {
      state.filters.search = e.target.value;
      renderAll();
    });
    document.getElementById("search-clear-btn").addEventListener("click", () => {
      state.filters.search = "";
      renderAll();
    });

    const selectBindings = [
      ["filter-tier", "tier"],
      ["filter-country", "country"],
      ["filter-product", "product"],
      ["filter-sector", "sector"],
      ["filter-channel", "channel"],
      ["filter-folder", "folder"],
      ["filter-review", "reviewStatus"],
      ["filter-sort", "sort"]
    ];
    selectBindings.forEach(([domId, key]) => {
      document.getElementById(domId).addEventListener("change", e => {
        state.filters[key] = e.target.value;
        renderAll();
      });
    });

    document.getElementById("reset-filters-btn").addEventListener("click", resetFilters);

    document.getElementById("active-filters-bar").addEventListener("click", e => {
      const btn = e.target.closest("[data-clear-filter]");
      if (!btn) return;
      const k = btn.getAttribute("data-clear-filter");
      state.filters[k] = k === "search" ? "" : "all";
      renderAll();
    });

    // Main View Container Delegation
    document.getElementById("view-container").addEventListener("click", e => {
      if (e.target.id === "empty-reset-btn") {
        resetFilters();
        return;
      }
      const openBtn = e.target.closest("[data-action='open-dossier']");
      if (openBtn) {
        openDossier(openBtn.getAttribute("data-id"));
        return;
      }
      const starBtn = e.target.closest("[data-action='star']");
      if (starBtn) {
        const id = starBtn.getAttribute("data-id");
        const rev = getReview(id);
        rev.starred = !rev.starred;
        const { isoTime, displayTime } = makeAutoTimestamp();
        rev.updatedAt = isoTime;
        rev.displayUpdatedAt = displayTime;
        saveReviews(id, null);
        renderAll();
        showToast(rev.starred ? `Bookmarked ${id}` : `Removed bookmark from ${id}`);
        return;
      }
      const cmpBtn = e.target.closest("[data-action='compare']");
      if (cmpBtn) {
        const id = cmpBtn.getAttribute("data-id");
        if (state.compareIds.includes(id)) {
          state.compareIds = state.compareIds.filter(x => x !== id);
        } else {
          if (state.compareIds.length >= 4) {
            showToast("Compare up to 4 programs at a time");
            return;
          }
          state.compareIds.push(id);
        }
        renderAll();
        return;
      }
      const prodChip = e.target.closest("[data-filter-product]");
      if (prodChip) {
        const pid = prodChip.getAttribute("data-filter-product");
        state.filters.product = state.filters.product === pid ? "all" : pid;
        renderAll();
        return;
      }
      const jumpProd = e.target.closest("[data-jump-product]");
      if (jumpProd) {
        state.filters.product = jumpProd.getAttribute("data-jump-product");
        state.view = "cards";
        renderAll();
        window.scrollTo({ top: 320, behavior: "smooth" });
        return;
      }
      const sortTh = e.target.closest("[data-sort-col]");
      if (sortTh) {
        state.filters.sort = sortTh.getAttribute("data-sort-col");
        renderAll();
      }
    });

    document.getElementById("view-container").addEventListener("change", e => {
      const sel = e.target.closest("[data-action='status-change']");
      if (!sel) return;
      const id = sel.getAttribute("data-id");
      const rev = getReview(id);
      const { isoTime, displayTime } = makeAutoTimestamp();
      rev.status = sel.value;
      rev.updatedAt = isoTime;
      rev.displayUpdatedAt = displayTime;
      saveReviews(id, null);
      renderAll();
      showToast(`Updated ${id} to "${getStatusLabel(sel.value)}" (${displayTime})`);
    });

    // Dossier controls
    document.getElementById("dossier-close-btn").addEventListener("click", closeDossier);
    document.getElementById("dossier-backdrop").addEventListener("click", closeDossier);
    document.getElementById("dossier-fullscreen-btn").addEventListener("click", e => {
      state.dossierFullscreen = !state.dossierFullscreen;
      document.getElementById("dossier-drawer").classList.toggle("fullscreen", state.dossierFullscreen);
      e.currentTarget.textContent = state.dossierFullscreen ? "Exit Full Width" : "Full Width";
    });
    document.getElementById("dossier-print-btn").addEventListener("click", () => window.print());

    // Keyboard navigation
    document.addEventListener("keydown", e => {
      if (e.key === "Escape") {
        closeDossier();
        document.querySelectorAll(".modal-overlay").forEach(m => m.classList.remove("open"));
      } else if (state.activeDossierId && !["INPUT", "TEXTAREA", "SELECT"].includes(document.activeElement?.tagName)) {
        if (e.key === "ArrowRight") document.getElementById("dossier-next-btn")?.click();
        if (e.key === "ArrowLeft") document.getElementById("dossier-prev-btn")?.click();
      }
    });

    // Compare bar
    document.getElementById("open-compare-btn").addEventListener("click", openCompareModal);
    document.getElementById("clear-compare-btn").addEventListener("click", () => {
      state.compareIds = [];
      renderAll();
    });
    document.getElementById("close-compare-modal").addEventListener("click", () => {
      document.getElementById("compare-modal").classList.remove("open");
    });

    // Cloud Sync Modal (Option A: Firebase & Option B: Google Sheets)
    document.getElementById("cloud-sync-btn")?.addEventListener("click", () => {
      updateCloudStatusUI();
      document.getElementById("cloud-modal").classList.add("open");
    });
    document.getElementById("close-cloud-modal")?.addEventListener("click", () => {
      document.getElementById("cloud-modal").classList.remove("open");
    });
    document.getElementById("copy-gsheet-script-btn")?.addEventListener("click", () => {
      navigator.clipboard?.writeText(GSHEET_SCRIPT_TEMPLATE);
      showToast("Copied Google Apps Script! Paste it into Extensions → Apps Script in Google Sheets.");
    });
    document.getElementById("save-gsheet-btn")?.addEventListener("click", async () => {
      const url = document.getElementById("cloud-gsheet-url").value.trim();
      if (!url) {
        showToast("Please paste a valid Google Apps Script Web App URL");
        return;
      }
      localStorage.setItem(STORAGE_KEY_CLOUD, JSON.stringify({ googleSheetsWebAppUrl: url, firebaseDatabaseUrl: "" }));
      updateCloudStatusUI();
      await pushReviewsToCloud();
      await pullReviewsFromCloud(false);
    });
    document.getElementById("save-firebase-btn")?.addEventListener("click", async () => {
      const url = document.getElementById("cloud-firebase-url").value.trim().replace(/\/+$/, "");
      if (!url) {
        showToast("Please paste a valid Firebase Realtime Database URL");
        return;
      }
      localStorage.setItem(STORAGE_KEY_CLOUD, JSON.stringify({ googleSheetsWebAppUrl: "", firebaseDatabaseUrl: url }));
      updateCloudStatusUI();
      await pushReviewsToCloud();
      await pullReviewsFromCloud(false);
    });
    document.getElementById("cloud-sync-now-btn")?.addEventListener("click", async () => {
      await pushReviewsToCloud();
      await pullReviewsFromCloud(false);
    });
    document.getElementById("cloud-disconnect-btn")?.addEventListener("click", async () => {
      localStorage.removeItem(STORAGE_KEY_CLOUD);
      document.getElementById("cloud-gsheet-url").value = "";
      document.getElementById("cloud-firebase-url").value = "";
      updateCloudStatusUI();
      const { activeProvider } = getCloudSettings();
      if (activeProvider === "google_sheets") {
        await pullReviewsFromCloud(true);
        showToast("Restored default shared Google Sheets connection!");
      } else {
        showToast("Disconnected cloud sync (using LocalStorage)");
      }
    });

    // Export & Import Modal
    document.getElementById("export-btn").addEventListener("click", () => {
      document.getElementById("io-modal").classList.add("open");
    });
    document.getElementById("close-io-modal").addEventListener("click", () => {
      document.getElementById("io-modal").classList.remove("open");
    });

    document.getElementById("btn-export-xlsx").addEventListener("click", () => {
      if (state.serverConnected) {
        window.location.href = "/api/export-excel";
        showToast("Downloading styled Reviewed Excel workbook...");
      } else {
        window.NaueParser.exportCSV(getFilteredProjects(), state.reviews);
        showToast("Exported Reviewed CSV spreadsheet (run start_website.bat for styled .xlsx)");
      }
    });

    document.getElementById("btn-export-csv").addEventListener("click", () => {
      window.NaueParser.exportCSV(getFilteredProjects(), state.reviews);
      showToast("Exported Reviewed CSV");
    });

    document.getElementById("excel-file-input").addEventListener("change", async e => {
      const file = e.target.files?.[0];
      if (!file) return;
      try {
        const buf = await file.arrayBuffer();
        if (state.serverConnected) {
          const resp = await fetch("/api/rebuild-from-excel", { method: "POST", body: buf });
          const json = await resp.json();
          if (json.ok && json.data) {
            state.payload = json.data;
          }
        } else {
          state.payload = window.NaueParser.parseWorkbookBuffer(buf, state.payload.productsCatalog);
        }
        populateFilterDropdowns();
        renderAll();
        document.getElementById("io-modal").classList.remove("open");
        showToast(`Successfully loaded ${state.payload.projects.length} programs from ${file.name}`);
      } catch (err) {
        alert("Could not parse workbook: " + err.message);
      }
    });
  }

  async function init() {
    // Restore theme
    const savedTheme = localStorage.getItem(STORAGE_KEY_THEME);
    if (savedTheme === "dark") {
      document.documentElement.setAttribute("data-theme", "dark");
    }

    // Load localStorage reviews first
    try {
      const localRev = localStorage.getItem(STORAGE_KEY_REVIEWS);
      if (localRev) state.reviews = mergeReviewsMaps({}, JSON.parse(localRev));
    } catch (e) {}

    populateFilterDropdowns();
    bindEvents();
    updateCloudStatusUI();
    renderAll();

    // Check if Python server.py is active for persistent file sync
    try {
      const res = await fetch("/api/state", { cache: "no-store" });
      if (res.ok) {
        const json = await res.json();
        state.serverConnected = true;
        if (json.data && json.data.projects && json.data.projects.length) {
          state.payload = json.data;
        }
        if (json.reviews && Object.keys(json.reviews).length > 0) {
          state.reviews = mergeReviewsMaps(state.reviews, json.reviews);
          localStorage.setItem(STORAGE_KEY_REVIEWS, JSON.stringify(state.reviews));
        }
        populateFilterDropdowns();
        updateCloudStatusUI();
        renderAll();
      }
    } catch (e) {
      // On GitHub Pages without a cloud backend, fall back to committed data/team_reviews.json baseline
      if (getCloudSettings().activeProvider === "local") {
        try {
          const staticRev = await fetch("data/team_reviews.json", { cache: "no-store" });
          if (staticRev.ok) {
            const jsonRev = await staticRev.json();
            if (jsonRev && typeof jsonRev === "object") {
              state.reviews = mergeReviewsMaps(jsonRev, state.reviews);
              renderAll();
            }
          }
        } catch (err2) {}
      }
    }

    // Pull latest reviews from Option A (Firebase) or Option B (Google Sheets) if configured
    await pullReviewsFromCloud(true);
    const pollSec = Number(window.NAUE_CLOUD_CONFIG?.autoPollSeconds ?? 30);
    if (pollSec >= 10) {
      setInterval(() => pullReviewsFromCloud(true), pollSec * 1000);
    }
  }

  window.NaueApp = { openDossier };
  document.addEventListener("DOMContentLoaded", init);
})();

