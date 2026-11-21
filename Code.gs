/**
 * EZScript Christmas Party 2026 — RSVP & tracking backend
 * --------------------------------------------------------
 * Google Apps Script bound to a Google Sheet. Receives POSTs from index.html
 * (hosted on GitHub Pages) and writes them into three tabs:
 *
 *   RSVPs     – one row per guest (keyed by email; re-submits update the row)
 *   Activity  – every tracked event: page views, steps, answers, calendar adds…
 *   Summary   – live headcount / funnel formulas
 *
 * Download as Excel any time: File → Download → Microsoft Excel (.xlsx).
 * Setup steps: see README.md.
 */

// ---- Settings -------------------------------------------------------------
const NOTIFY_EMAIL = "";                 // e.g. "events@ezscript.com" — emails you on every new RSVP. Leave "" to disable.
const TIMEZONE     = "America/Los_Angeles";
const EVENT_ID     = "ezscript-christmas-2026";
const SPREADSHEET_ID = "";               // Only needed if this script was NOT created from the Sheet (Extensions → Apps Script).
                                         // Paste the ID from the Sheet URL: docs.google.com/spreadsheets/d/<THIS PART>/edit

// ---- Sheet layout ---------------------------------------------------------
const RSVP_SHEET = "RSVPs";
const ACT_SHEET  = "Activity";
const SUM_SHEET  = "Summary";

const RSVP_HEADERS = ["Email", "Name", "Phone", "Attending", "Guests (incl. self)", "Invite ID",
  "First Submitted", "Last Updated", "Times Submitted", "OS", "Device", "Browser", "Visitor ID", "Session ID"];
const ACT_HEADERS = ["Timestamp", "Event", "Detail", "Step", "Invite ID", "Email", "Visitor ID", "Session ID",
  "OS", "Device", "Browser", "Screen", "Language", "Timezone", "Page URL", "Referrer"];

const ALLOWED_EVENTS = ["page_view", "rsvp_started", "step_view", "validation_error", "answer",
  "calendar_add", "edit_response", "rsvp_submitted"];

// ---- Web app entry points -------------------------------------------------
/** Health check — open the /exec URL in a browser; the page's ?debug mode calls this too. */
function doGet() {
  try {
    const ss = getSpreadsheet_();
    ensureSheets_(ss);
    return json_({ ok: true, service: "EZScript RSVP", event: EVENT_ID, sheet: ss.getName(),
      rsvps: Math.max(0, ss.getSheetByName(RSVP_SHEET).getLastRow() - 1),
      activity: Math.max(0, ss.getSheetByName(ACT_SHEET).getLastRow() - 1) });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doPost(e) {
  let body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || "{}"); }
  catch (err) { return json_({ ok: false, error: "invalid_json" }); }

  if (body.hp) return json_({ ok: true });                       // honeypot filled → bot; silently drop
  if (body.eventId && body.eventId !== EVENT_ID) return json_({ ok: false, error: "wrong_event" });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const ss = getSpreadsheet_();
    ensureSheets_(ss);

    if (body.type === "rsvp") {
      const r = validateRSVP_(body);
      if (!r.ok) return json_({ ok: false, error: r.error });
      const isNew = upsertRSVP_(ss, r.rsvp, body);
      logActivity_(ss, Object.assign({}, body, { event: "rsvp_submitted", detail: r.rsvp.attending + (r.rsvp.guests ? " / " + r.rsvp.guests : "") }));
      if (NOTIFY_EMAIL && isNew) notify_(r.rsvp, body);
    } else if (body.type === "event") {
      if (ALLOWED_EVENTS.indexOf(body.event) === -1) return json_({ ok: false, error: "unknown_event" });
      logActivity_(ss, body);
    } else {
      return json_({ ok: false, error: "unknown_type" });
    }
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  return json_({ ok: true });
}

// ---- One-time setup (run from the Apps Script editor) ---------------------
function setup() {
  const ss = getSpreadsheet_();
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", ss.getId());
  ss.setSpreadsheetTimeZone(TIMEZONE);
  ensureSheets_(ss);
  const def = ss.getSheetByName("Sheet1");
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  ss.setActiveSheet(ss.getSheetByName(SUM_SHEET));
}

/**
 * Self-test — run this from the editor (select testWrite → Run).
 * It writes a sample RSVP exactly like the website does. If it shows up in the RSVPs tab,
 * the script side is fine and any problem is the deployment (access / URL / version).
 */
function testWrite() {
  const res = doPost({ postData: { contents: JSON.stringify({
    type: "rsvp", eventId: EVENT_ID, name: "Test Guest", email: "test.guest@example.com", phone: "9495550123",
    attending: "Yes", guests: 2, invite: "TEST", os: "editor", device: "desktop", browser: "Apps Script",
    visitorId: "test", sessionId: "test" }) } });
  Logger.log(res.getContent());
}

// ---- Internals ------------------------------------------------------------
function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();       // works when the script is bound to the Sheet
  if (active) return active;
  const id = SPREADSHEET_ID || PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id) throw new Error("Script is not attached to a Sheet. Create it from the Sheet (Extensions → Apps Script) or set SPREADSHEET_ID in Code.gs.");
  return SpreadsheetApp.openById(id);
}

function ensureSheets_(ss) {
  let r = ss.getSheetByName(RSVP_SHEET);
  if (!r) {
    r = ss.insertSheet(RSVP_SHEET);
    styleHeader_(r, RSVP_HEADERS);
    r.getRange("C:C").setNumberFormat("@");                       // phone as text (keeps leading zeros)
    r.getRange("G:H").setNumberFormat("yyyy-mm-dd hh:mm");
    r.setColumnWidth(1, 220); r.setColumnWidth(2, 180);
  }
  let a = ss.getSheetByName(ACT_SHEET);
  if (!a) {
    a = ss.insertSheet(ACT_SHEET);
    styleHeader_(a, ACT_HEADERS);
    a.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");
  }
  if (!ss.getSheetByName(SUM_SHEET)) buildSummary_(ss.insertSheet(SUM_SHEET, 0));
}

function styleHeader_(sh, headers) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#15532E");
  sh.setFrozenRows(1);
}

function buildSummary_(sh) {
  const rows = [
    ["EZScript Christmas Party — Dec 17, 2026 · Newport Beach, CA", ""],
    ["", ""],
    ["RSVPs", ""],
    ["Total responses",               "=COUNTA(RSVPs!A2:A)"],
    ["Attending — Yes",               '=COUNTIF(RSVPs!D2:D,"Yes")'],
    ["Attending — Maybe",             '=COUNTIF(RSVPs!D2:D,"Maybe")'],
    ["Attending — No",                '=COUNTIF(RSVPs!D2:D,"No")'],
    ["Confirmed headcount (Yes)",     '=SUMIF(RSVPs!D2:D,"Yes",RSVPs!E2:E)'],
    ["Possible headcount (Yes + Maybe)", '=SUMIF(RSVPs!D2:D,"Yes",RSVPs!E2:E)+SUMIF(RSVPs!D2:D,"Maybe",RSVPs!E2:E)'],
    ["", ""],
    ["Funnel (unique visitors)", ""],
    ["Opened invitation",             '=IFERROR(COUNTUNIQUE(FILTER(Activity!G2:G,Activity!B2:B="page_view")),0)'],
    ["Started RSVP",                  '=IFERROR(COUNTUNIQUE(FILTER(Activity!G2:G,Activity!B2:B="rsvp_started")),0)'],
    ["Submitted RSVP",                '=IFERROR(COUNTUNIQUE(FILTER(Activity!G2:G,Activity!B2:B="rsvp_submitted")),0)'],
    ["Added to calendar",             '=IFERROR(COUNTUNIQUE(FILTER(Activity!G2:G,Activity!B2:B="calendar_add")),0)'],
    ["Conversion (submitted ÷ opened)", "=IFERROR(B14/B12,0)"],
    ["", ""],
    ["Devices (RSVPs)", ""],
    ["Mobile",  '=COUNTIF(RSVPs!K2:K,"mobile")'],
    ["Tablet",  '=COUNTIF(RSVPs!K2:K,"tablet")'],
    ["Desktop", '=COUNTIF(RSVPs!K2:K,"desktop")']
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange("A1").setFontSize(14).setFontWeight("bold").setFontColor("#15532E");
  ["A3", "A11", "A18"].forEach(function (c) { sh.getRange(c).setFontWeight("bold").setFontColor("#15532E"); });
  sh.getRange("B16").setNumberFormat("0%");
  sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 120);
}

function validateRSVP_(b) {
  const name = clean_(b.name, 80);
  const email = clean_(b.email, 120).toLowerCase();
  const phone = String(b.phone || "").replace(/\D/g, "").slice(0, 15);
  const attending = ["Yes", "No", "Maybe"].indexOf(b.attending) > -1 ? b.attending : "";
  const g = parseInt(b.guests, 10);
  const guests = attending === "No" ? "" : (g >= 1 && g <= 4 ? g : "");
  if (name.length < 2) return { ok: false, error: "name" };
  if (!/^[^\s@]+@[^\s@]+\.[A-Za-z]{2,}$/.test(email)) return { ok: false, error: "email" };
  if (phone.length < 10) return { ok: false, error: "phone" };
  if (!attending) return { ok: false, error: "attending" };
  if (attending !== "No" && guests === "") return { ok: false, error: "guests" };
  return { ok: true, rsvp: { name: name, email: email, phone: phone, attending: attending, guests: guests } };
}

/** Insert or update by email. Returns true when it's a new guest. */
function upsertRSVP_(ss, r, b) {
  const sh = ss.getSheetByName(RSVP_SHEET);
  const now = new Date();
  const last = sh.getLastRow();
  let rowIdx = -1;
  if (last > 1) {
    const emails = sh.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < emails.length; i++) {
      if (String(emails[i][0]).toLowerCase() === r.email) { rowIdx = i + 2; break; }
    }
  }
  const invite = clean_(b.invite, 64);
  if (rowIdx === -1) {
    sh.appendRow([r.email, r.name, r.phone, r.attending, r.guests, invite, now, now, 1,
      clean_(b.os, 20), clean_(b.device, 20), clean_(b.browser, 30), clean_(b.visitorId, 64), clean_(b.sessionId, 64)]);
    return true;
  }
  const cur = sh.getRange(rowIdx, 1, 1, RSVP_HEADERS.length).getValues()[0];
  sh.getRange(rowIdx, 1, 1, RSVP_HEADERS.length).setValues([[
    r.email, r.name, r.phone, r.attending, r.guests, invite || cur[5], cur[6] || now, now, (Number(cur[8]) || 0) + 1,
    clean_(b.os, 20), clean_(b.device, 20), clean_(b.browser, 30), clean_(b.visitorId, 64), clean_(b.sessionId, 64)
  ]]);
  return false;
}

function logActivity_(ss, b) {
  ss.getSheetByName(ACT_SHEET).appendRow([
    new Date(), clean_(b.event, 40), clean_(b.detail, 200), clean_(b.step, 20), clean_(b.invite, 64),
    clean_(b.email, 120).toLowerCase(), clean_(b.visitorId, 64), clean_(b.sessionId, 64),
    clean_(b.os, 20), clean_(b.device, 20), clean_(b.browser, 30), clean_(b.screen, 20),
    clean_(b.language, 20), clean_(b.timezone, 50), clean_(b.url, 300), clean_(b.referrer, 300)
  ]);
}

function notify_(r, b) {
  try {
    MailApp.sendEmail(NOTIFY_EMAIL,
      "New RSVP: " + r.name + " — " + r.attending + (r.guests ? " (" + r.guests + ")" : ""),
      ["Name: " + r.name, "Email: " + r.email, "Phone: " + r.phone, "Attending: " + r.attending,
       "Guests: " + (r.guests || "—"), "Invite ID: " + (b.invite || "—"), "Device: " + b.device + " / " + b.os].join("\n"));
  } catch (err) { console.error(err); }
}

/** Trim, cap length and neutralise spreadsheet formula injection. */
function clean_(v, max) {
  let s = String(v == null ? "" : v).replace(/[\u0000-\u001F]/g, " ").trim().slice(0, max || 200);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
