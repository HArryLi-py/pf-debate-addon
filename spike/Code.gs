/**
 * FIREBASE RTDB SPIKE — Apps Script add-on server (bound Google Doc).
 *
 * Goal: verify whether a Google Docs add-on sidebar (HtmlService IFRAME
 * sandbox) can sustain a CLIENT-SIDE realtime connection to Firebase RTDB
 * with sub-second propagation between two sidebars on the same Doc.
 *
 * This is the ONE open technical risk before committing the Phase 2
 * architecture (Firebase-backed shared timer/phase/chat).
 *
 * Jobs here: onOpen (menu) / showSidebar (serve Sidebar.html) /
 *            getBootstrap (round-id + viewer-id) /
 *            Tier-3 fallback (google.script.run -> UrlFetchApp -> Firebase REST).
 */

// ---------- Edit this one block ----------
var FIREBASE_CONFIG = {
  dbHost:   "https://YOUR-TEST-PROJECT.firebaseio.com", // <-- your test project
  auth:     "",        // "" while rules are open; DB secret/legacy token to lock down
  testPath: "spike/rounds"
};
// ----------------------------------------

function onOpen() {
  return DocumentApp.getUi()
    .createMenu("Firebase Spike")
    .addItem("Open realtime sidebar", "showSidebar")
    .addToUi();
}

function showSidebar() {
  var ui = DocumentApp.getUi();
  var t  = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Firebase RTDB spike")
    .setXFrameOptions(HtmlService.XFrameOptionsMode.ALLOWALL);
  ui.showSidebar(t);
}

function getBootstrap() {
  var docId;
  try { docId = DocumentApp.getActiveDocument().getId(); } catch (e) { docId = "standalone"; }
  return {
    dbHost:   FIREBASE_CONFIG.dbHost,
    auth:     FIREBASE_CONFIG.auth,
    testPath: FIREBASE_CONFIG.testPath,
    // Shared key: BOTH sidebars derive the same roundId from the same Doc.
    roundId:  "doc-" + docId.slice(-6),
    viewerId: "u-" + Math.random().toString(36).slice(2, 8) + "-" + Date.now().toString(36),
    serverTime: new Date().toISOString()
  };
}

// TIER 3 FALLBACK: server-side UrlFetchApp proxy + polling.
// Apps Script server CANNOT hold WebSocket or use EventSource (browser APIs).
// ~250 ms RTT — NOT sub-second realtime. Last resort only.
function serverReadRound(roundId) {
  var url = FIREBASE_CONFIG.dbHost + "/" + FIREBASE_CONFIG.testPath + "/" +
            encodeURIComponent(roundId) + "/.json" +
            (FIREBASE_CONFIG.auth ? "?auth=" + encodeURIComponent(FIREBASE_CONFIG.auth) : "");
  try {
    var resp = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
    var code = resp.getResponseCode();
    var body = resp.getContentText();
    if (code === 200) return body;
    return "ERR " + code + ": " + body.slice(0, 200);
  } catch (e) { return "FETCH_ERR: " + e.message; }
}

function serverWriteRound(roundId, payload) {
  var url = FIREBASE_CONFIG.dbHost + "/" + FIREBASE_CONFIG.testPath + "/" +
            encodeURIComponent(roundId) + "/.json" +
            (FIREBASE_CONFIG.auth ? "?auth=" + encodeURIComponent(FIREBASE_CONFIG.auth) : "");
  try {
    var resp = UrlFetchApp.fetch(url, {
      method: "put", contentType: "application/json",
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    var code = resp.getResponseCode();
    var body = resp.getContentText();
    if (code === 200) return "OK " + body.slice(0, 120);
    return "ERR " + code + ": " + body.slice(0, 200);
  } catch (e) { return "FETCH_ERR: " + e.message; }
}
