/**
 * ============================================================================
 * NAUE OPPORTUNITY TRACKER — GOOGLE SHEETS MULTI-USER BACKEND (OPTION B)
 * ============================================================================
 * Supports MULTIPLE PEOPLE reviewing the same project simultaneously without
 * overwriting each other! Every review entry has a unique ID and automatic
 * server/client timestamp, stored across two auto-created tabs:
 *   1. "Team_Reviews" (Merged state per project, preserving all reviewers' entries)
 *   2. "All_Reviews_Log" (Readable row-by-row table of every person's review)
 *
 * SETUP (60 Seconds, 100% Free):
 * 1. Open a Google Sheet (https://sheets.new) -> Extensions -> Apps Script
 * 2. Paste this entire file into Code.gs
 * 3. Click Deploy -> New deployment -> Web app
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 * 4. Copy the Web app URL (.../exec) into `js/cloud-config.js` or ☁️ Cloud Sync.
 */

var SUMMARY_SHEET = "Team_Reviews";
var LOG_SHEET = "All_Reviews_Log";

function ensureSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var s1 = ss.getSheetByName(SUMMARY_SHEET);
  if (!s1) {
    s1 = ss.insertSheet(SUMMARY_SHEET);
    var h1 = [
      "Project ID",
      "Current Stage",
      "Starred",
      "Assigned Lead",
      "Review Count",
      "Reviewers",
      "Latest Review Summary",
      "Last Updated (Auto)",
      "Raw Merged JSON"
    ];
    s1.getRange(1, 1, 1, h1.length).setValues([h1]).setFontWeight("bold").setBackground("#00549F").setFontColor("#FFFFFF");
    s1.setFrozenRows(1);
  }

  var s2 = ss.getSheetByName(LOG_SHEET);
  if (!s2) {
    s2 = ss.insertSheet(LOG_SHEET);
    var h2 = [
      "Timestamp (Auto)",
      "Project ID",
      "Reviewer Name",
      "Reviewer Verdict / Stage",
      "Review Comment / Assessment",
      "Unique Review ID"
    ];
    s2.getRange(1, 1, 1, h2.length).setValues([h2]).setFontWeight("bold").setBackground("#059669").setFontColor("#FFFFFF");
    s2.setFrozenRows(1);
  }

  return { summary: s1, log: s2 };
}

function readExistingMap(summarySheet) {
  var lastRow = summarySheet.getLastRow();
  var map = {};
  if (lastRow >= 2) {
    var data = summarySheet.getRange(2, 1, lastRow - 1, 9).getValues();
    for (var i = 0; i < data.length; i++) {
      var pid = String(data[i][0] || "").trim();
      var raw = String(data[i][8] || "").trim();
      if (pid && raw) {
        try {
          map[pid] = JSON.parse(raw);
        } catch (e) {}
      }
    }
  }
  return map;
}

function mergeProjectRecord(existing, incoming) {
  if (!existing) return incoming;
  if (!incoming) return existing;

  // Merge all individual review entries by unique id so NO reviewer overwrites another!
  var byId = {};
  var exList = existing.entries || existing.comments || [];
  var inList = incoming.entries || incoming.comments || [];

  for (var i = 0; i < exList.length; i++) {
    var e = exList[i];
    var eid = e.id || (e.author + "_" + (e.isoTime || e.date || i));
    byId[eid] = e;
  }
  for (var j = 0; j < inList.length; j++) {
    var n = inList[j];
    var nid = n.id || (n.author + "_" + (n.isoTime || n.date || j));
    byId[nid] = n;
  }

  var mergedEntries = Object.keys(byId)
    .map(function (k) { return byId[k]; })
    .sort(function (a, b) {
      return String(b.isoTime || "").localeCompare(String(a.isoTime || ""));
    });

  // Keep whichever top-level status/metadata has the newer updatedAt timestamp
  var exTime = String(existing.updatedAt || "");
  var inTime = String(incoming.updatedAt || "");
  var newer = inTime >= exTime ? incoming : existing;

  // Merge checklist booleans (union of checked items if same timestamp, or newer)
  var mergedChecklist = Object.assign({}, existing.checklist || {}, incoming.checklist || {});

  return {
    status: newer.status || existing.status || "unreviewed",
    starred: Boolean(newer.starred),
    assignee: newer.assignee || existing.assignee || "",
    checklist: mergedChecklist,
    updatedAt: newer.updatedAt || new Date().toISOString(),
    updatedBy: newer.updatedBy || existing.updatedBy || "",
    entries: mergedEntries
  };
}

function doGet() {
  var sheets = ensureSheets();
  var reviews = readExistingMap(sheets.summary);
  return ContentService.createTextOutput(
    JSON.stringify({ ok: true, provider: "google_sheets", reviews: reviews })
  ).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var sheets = ensureSheets();
    var currentMap = readExistingMap(sheets.summary);
    var body = JSON.parse(e.postData.contents || "{}");
    var incomingMap = body.reviews || {};

    var inKeys = Object.keys(incomingMap);
    for (var i = 0; i < inKeys.length; i++) {
      var pid = inKeys[i];
      currentMap[pid] = mergeProjectRecord(currentMap[pid], incomingMap[pid]);
    }

    var allKeys = Object.keys(currentMap).sort();
    var summaryRows = [];
    var logRows = [];

    for (var k = 0; k < allKeys.length; k++) {
      var projId = allKeys[k];
      var rec = currentMap[projId] || {};
      var entries = rec.entries || [];
      var authors = [];
      for (var m = 0; m < entries.length; m++) {
        var auth = entries[m].author || "Reviewer";
        if (authors.indexOf(auth) === -1) authors.push(auth);
        logRows.push([
          entries[m].displayTime || entries[m].date || entries[m].isoTime || "",
          projId,
          auth,
          entries[m].verdictLabel || entries[m].verdict || rec.status || "",
          entries[m].text || "",
          entries[m].id || ""
        ]);
      }
      var latestSummary = entries.length > 0
        ? "[" + (entries[0].author || "Reviewer") + "] " + (entries[0].text || entries[0].verdictLabel || "")
        : "";

      summaryRows.push([
        projId,
        rec.status || "unreviewed",
        rec.starred ? "★ Yes" : "No",
        rec.assignee || "",
        entries.length,
        authors.join(", "),
        latestSummary,
        rec.updatedAt || "",
        JSON.stringify(rec)
      ]);
    }

    if (sheets.summary.getLastRow() >= 2) {
      sheets.summary.getRange(2, 1, sheets.summary.getLastRow() - 1, 9).clearContent();
    }
    if (summaryRows.length > 0) {
      sheets.summary.getRange(2, 1, summaryRows.length, 9).setValues(summaryRows);
    }

    // Sort logRows newest first
    logRows.sort(function (a, b) { return String(b[0]).localeCompare(String(a[0])); });
    if (sheets.log.getLastRow() >= 2) {
      sheets.log.getRange(2, 1, sheets.log.getLastRow() - 1, 6).clearContent();
    }
    if (logRows.length > 0) {
      sheets.log.getRange(2, 1, logRows.length, 6).setValues(logRows);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ ok: true, mergedProjects: summaryRows.length, totalReviews: logRows.length })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, error: String(err) })
    ).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
