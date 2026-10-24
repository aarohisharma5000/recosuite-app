/**
 * RecoSuite support tickets -> Google Sheet ("Support" tab) + optional email notification.
 * Paste this whole file into Extensions -> Apps Script of the Google Sheet that should hold the tickets.
 * Deploy as Web app: Execute as "Me", Who has access "Anyone". See SETUP.md.
 */

// ===================== SETTINGS (safe to change) =====================
var SEND_EMAIL_NOTIFICATION = true;                         // false = only write to the sheet, no email
var NOTIFY_EMAIL            = "aarohisharma5000@gmail.com";  // where new-ticket emails go
var SHEET_NAME              = "Support";
var MIN_TIME_MS             = 3000;                          // spam check: form filled faster than this = rejected
var MAX_TICKETS_PER_EMAIL   = 5;                             // per email address ...
var RATE_WINDOW_SECONDS     = 600;                           // ... per 10 minutes
// =====================================================================

var HEADERS = ["Timestamp IST", "Issue ID", "Email", "Issue Type", "Description", "Status", "User Agent"];
var ISSUE_TYPES = ["New Enquiry", "Access Issue", "Tool Error", "Wrong Results", "Download Issue", "Payment Issue",
                   "Upgrade Request", "Feature Request", "Other"];
var LIMITS = { email: 254, description: 5000, userAgent: 200 };

function doPost(e) {
  try {
    var p = (e && e.parameter) || {};

    // 1) spam checks. Honeypot filled -> pretend success, store nothing.
    if (clean_(p.website, 200)) return json_({ ok: true });
    var elapsed = parseInt(p.elapsedMs, 10);
    if (!(elapsed >= MIN_TIME_MS)) return json_({ ok: false, error: "Please wait a few seconds and submit again." });

    // 2) validation + length limits
    var issueId = clean_(p.issueId, 20).toUpperCase();
    if (!/^RS-\d{6}-[A-Z0-9]{4}$/.test(issueId)) return json_({ ok: false, error: "Invalid issue ID. Please reload the page." });
    var email = clean_(p.email, LIMITS.email).toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return json_({ ok: false, error: "Please enter a valid email address." });
    var issueType = clean_(p.issueType, 40);
    if (ISSUE_TYPES.indexOf(issueType) < 0) issueType = "Other";
    var description = clean_(p.description, LIMITS.description, true);
    if (description.length < 10) return json_({ ok: false, error: "Please describe your issue." });
    var userAgent = clean_(p.userAgent, LIMITS.userAgent);

    // 3) simple rate limit per email
    var cache = CacheService.getScriptCache();
    var rateKey = "rate_" + email;
    var count = parseInt(cache.get(rateKey) || "0", 10);
    if (count >= MAX_TICKETS_PER_EMAIL) return json_({ ok: false, error: "Too many tickets from this email. Please try again later or use WhatsApp." });

    // 4) write (locked so two tickets at the same moment don't collide)
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    var duplicate = false;
    try {
      var sheet = getSheet_();
      duplicate = sheet.getLastRow() > 1 &&
        sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).createTextFinder(issueId).matchEntireCell(true).findNext() !== null;
      if (!duplicate) {        // the same Issue ID sent twice (e.g. "Try again") is stored once
        var ts = Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");
        sheet.appendRow([ts, issueId, email, issueType, description, "Open", userAgent].map(safe_));
      }
    } finally {
      lock.releaseLock();
    }
    if (duplicate) return json_({ ok: true, issueId: issueId, duplicate: true });
    cache.put(rateKey, String(count + 1), RATE_WINDOW_SECONDS);

    // 5) optional email notification (a mail problem never loses the ticket)
    if (SEND_EMAIL_NOTIFICATION) {
      try {
        MailApp.sendEmail({
          to: NOTIFY_EMAIL,
          replyTo: email,
          subject: "[RecoSuite Support] " + issueType + " — " + issueId,
          body: "New support ticket\n\nIssue ID: " + issueId + "\nEmail: " + email + "\nIssue Type: " + issueType +
                "\nTime (IST): " + Utilities.formatDate(new Date(), "Asia/Kolkata", "yyyy-MM-dd HH:mm") +
                "\nUser Agent: " + userAgent + "\n\nDescription:\n" + description +
                "\n\n— Reply to this email to answer the customer directly."
        });
      } catch (mailErr) {
        console.warn("Notification email failed: " + mailErr);
      }
    }
    return json_({ ok: true, issueId: issueId });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: "Server error. Please use WhatsApp or email instead." });
  }
}

// Opening the Web App URL in a browser shows this. Handy to check that the deployment works.
function doGet() {
  return json_({ ok: true, service: "RecoSuite support", hint: "POST tickets here from support.html" });
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.getRange("A:G").setNumberFormat("@");            // everything is plain text
    sheet.setColumnWidth(5, 420);
  }
  return sheet;
}

// Trim, drop control characters (keep new lines in the description), cut to max length.
function clean_(v, max, multiline) {
  var s = String(v == null ? "" : v);
  s = multiline ? s.replace(/\r\n?/g, "\n").replace(/[\u0000-\u0009\u000B-\u001F\u007F]/g, "")
                : s.replace(/[\u0000-\u001F\u007F]/g, " ");
  return s.trim().slice(0, max);
}

// Formula-injection guard: a value starting with = + - @ (or tab / CR) is stored as text.
function safe_(v) {
  var s = String(v);
  return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
