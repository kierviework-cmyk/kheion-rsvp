/**
 * Kheion's Birthday & Christening RSVP — Apps Script backend.
 * Bound to the "Kheion RSVP" Google Sheet. Deployed as a Web App.
 */

var RSVP_SHEET = 'RSVPs';
var CONFIG_SHEET = 'Config';
var SUMMARY_SHEET = 'Summary';

var RSVP_HEADERS = ['Timestamp', 'Full name', 'Phone', 'Email', 'Address', 'Attending', 'Headcount', 'Godparent volunteer', 'Gift claimed', 'Notes', 'Companions'];
var CONFIG_HEADERS = ['key', 'value'];

var DEFAULT_CONFIG = [
  ['celebrant_name', 'Kheion'],
  ['event_title', "Kheion's Birthday & Christening"],
  ['event_date', '2026-09-16'],
  ['start_time', '10:00 AM'],
  ['end_time', '2:00 PM'],
  ['timezone', 'Asia/Manila'],
  ['venue_name', '[VENUE_TBD]'],
  ['venue_address', '[VENUE_ADDRESS_TBD]'],
  ['host_message', "You're invited to a double celebration!"],
  ['rsvp_deadline', '[RSVP_DEADLINE_TBD]']
];

var MONTH_NAMES_ = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

function formatDateWords_(dateStr) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateStr || '');
  if (!m) return dateStr || '';
  var monthName = MONTH_NAMES_[parseInt(m[2], 10) - 1];
  if (!monthName) return dateStr;
  return monthName + ' ' + parseInt(m[3], 10) + ', ' + m[1];
}

/**
 * Run this once from the Apps Script editor (select it in the function
 * dropdown and click Run) to create and seed the three tabs. Safe to
 * re-run — it never overwrites existing data, only fills in what's missing.
 */
function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var rsvpSheet = ss.getSheetByName(RSVP_SHEET) || ss.insertSheet(RSVP_SHEET);
  if (rsvpSheet.getLastRow() === 0) {
    rsvpSheet.appendRow(RSVP_HEADERS);
    rsvpSheet.setFrozenRows(1);
  } else {
    // Migration: add any headers (e.g. "Companions") missing from sheets
    // created before this column existed.
    var existingHeaders = rsvpSheet.getRange(1, 1, 1, rsvpSheet.getLastColumn()).getValues()[0];
    RSVP_HEADERS.forEach(function (header) {
      if (existingHeaders.indexOf(header) === -1) {
        rsvpSheet.getRange(1, rsvpSheet.getLastColumn() + 1).setValue(header);
      }
    });
  }

  var configSheet = ss.getSheetByName(CONFIG_SHEET) || ss.insertSheet(CONFIG_SHEET);
  if (configSheet.getLastRow() === 0) {
    configSheet.appendRow(CONFIG_HEADERS);
    configSheet.setFrozenRows(1);
    // Force the values column to Plain Text *before* writing, so Sheets
    // never silently reinterprets "10:00 AM" / "2026-09-16" as a Date cell.
    configSheet.getRange('B2:B1000').setNumberFormat('@');
    DEFAULT_CONFIG.forEach(function (row) {
      configSheet.appendRow(row);
    });
  } else {
    // Repair: earlier runs let Sheets auto-convert date/time-looking text
    // into Date cells before Plain Text formatting existed. Force the
    // format now and rewrite any values that already got corrupted.
    configSheet.getRange('B2:B1000').setNumberFormat('@');
    var defaultByKey = {};
    DEFAULT_CONFIG.forEach(function (row) { defaultByKey[row[0]] = row[1]; });
    var configValues = configSheet.getDataRange().getValues();
    for (var r = 1; r < configValues.length; r++) {
      var key = configValues[r][0];
      var val = configValues[r][1];
      if (val instanceof Date && defaultByKey.hasOwnProperty(key)) {
        configSheet.getRange(r + 1, 2).setValue(defaultByKey[key]);
      }
    }
  }

  var summarySheet = ss.getSheetByName(SUMMARY_SHEET) || ss.insertSheet(SUMMARY_SHEET);
  if (summarySheet.getLastRow() === 0) {
    seedSummarySheet_(summarySheet);
  }

  // Remove the default "Sheet1" if it's still empty and unused.
  var sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && sheet1.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }
}

/**
 * Run this once from the editor after adding the confirmation-email feature.
 * Sending mail needs a permission the original deployment authorization
 * didn't cover; calling any Mail service here triggers that consent prompt
 * ahead of time so the first real guest submission doesn't fail silently.
 */
function authorizeMailSending() {
  console.log('Remaining daily email quota: ' + MailApp.getRemainingDailyQuota());
}

/**
 * One-time cleanup: wipes all RSVP test rows (keeps the header) so the
 * sheet starts clean before real guests use the form. Run once from the
 * editor, then delete this function — it's destructive and has no
 * business staying in the deployed script long-term.
 */
function clearAllTestData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var rsvpSheet = ss.getSheetByName(RSVP_SHEET);
  if (rsvpSheet && rsvpSheet.getLastRow() > 1) {
    rsvpSheet.getRange(2, 1, rsvpSheet.getLastRow() - 1, rsvpSheet.getLastColumn()).clearContent();
  }

  console.log('Cleared all RSVP rows.');
}

// Live-formula dashboard — every cell recalculates automatically as rows
// are appended to RSVPs, no code changes or re-runs needed.
function seedSummarySheet_(sheet) {
  sheet.getRange('A1').setValue('RSVP Summary').setFontWeight('bold').setFontSize(14);

  var stats = [
    ['Total RSVPs', '=COUNTA(RSVPs!B2:B)'],
    ['Attending (Yes)', '=COUNTIF(RSVPs!F2:F,"Yes")'],
    ['Not attending (No)', '=COUNTIF(RSVPs!F2:F,"No")'],
    ['Total headcount', '=SUM(RSVPs!G2:G)']
  ];
  sheet.getRange(3, 1, stats.length, 2).setValues(stats);
  sheet.getRange(3, 1, stats.length, 1).setFontWeight('bold');

  sheet.setFrozenRows(2);
  sheet.autoResizeColumns(1, 8);
}

function getConfigMap_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  var values = sheet.getDataRange().getValues();
  var config = {};
  for (var i = 1; i < values.length; i++) {
    var key = values[i][0];
    var value = values[i][1];
    if (key) config[key] = normalizeConfigValue_(key, value);
  }
  return config;
}

// Google Sheets silently converts cell text that looks like a date/time
// (e.g. "2026-09-16", "10:00 AM") into a Date object. Convert those back
// to the plain strings the frontend expects, regardless of key name.
function normalizeConfigValue_(key, value) {
  if (!(value instanceof Date)) return value;

  // Sheets stores date/time cells as timezone-less serial numbers; Apps
  // Script exposes them as Date objects whose UTC fields hold the literal
  // digits that were typed (e.g. "10:00 AM" -> UTC hours = 10). Format in
  // Etc/GMT (UTC, no offset) to recover those literal digits.
  if (/_time$/.test(key)) {
    return Utilities.formatDate(value, 'Etc/GMT', 'h:mm a');
  }
  if (/_date$/.test(key) || key === 'rsvp_deadline') {
    return Utilities.formatDate(value, 'Etc/GMT', 'yyyy-MM-dd');
  }
  return Utilities.formatDate(value, 'Etc/GMT', 'yyyy-MM-dd');
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Neutralize leading formula-trigger characters (=, +, -, @, tab, CR) before
// writing untrusted user input to a Sheet — SpreadsheetApp parses cell
// strings the same way manual entry does, so an unprefixed value starting
// with one of these becomes a live, auto-evaluating formula on write.
function sanitizeForSheet_(value) {
  var str = (value === null || value === undefined) ? '' : value.toString();
  return /^[=+\-@\t\r]/.test(str) ? ("'" + str) : str;
}

var MAX_FIELD_LEN = { name: 200, phone: 20, email: 254, address: 300, notes: 1000, companion: 200 };

// Cheap global throttle shared by every caller (Apps Script doesn't expose a
// reliable per-caller IP), capped low enough to blunt a flood/spam script
// without affecting normal RSVP traffic.
function isRateLimited_() {
  var cache = CacheService.getScriptCache();
  var key = 'rsvp_submit_count_60s';
  var count = Number(cache.get(key) || 0);
  if (count >= 20) return true;
  cache.put(key, String(count + 1), 60);
  return false;
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;

  if (action === 'info') {
    var config = getConfigMap_();
    return jsonResponse_({ status: 'ok', config: config });
  }

  return jsonResponse_({ status: 'ok', message: 'Kheion RSVP API is running.' });
}

function doPost(e) {
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch (parseErr) {
    return jsonResponse_({ status: 'error', message: 'Invalid submission format.' });
  }

  // Honeypot: a hidden field real users never see or fill. Pretend success
  // without writing anything, so an automated filler doesn't retry harder.
  if ((payload.website || '').toString().trim() !== '') {
    return jsonResponse_({ status: 'ok' });
  }

  if (isRateLimited_()) {
    return jsonResponse_({ status: 'error', message: 'Server is busy, please try again in a moment.' });
  }

  var lock = LockService.getScriptLock();
  var gotLock = lock.tryLock(10000);

  if (!gotLock) {
    return jsonResponse_({ status: 'error', message: 'Server is busy, please try again in a moment.' });
  }

  try {
    var name = (payload.name || '').toString().trim();
    var phone = (payload.phone || '').toString().trim();
    var email = (payload.email || '').toString().trim();
    var address = (payload.address || '').toString().trim();
    var attending = payload.attending === 'Yes' ? 'Yes' : 'No';
    var notes = (payload.notes || '').toString().trim();
    var companionsInput = Array.isArray(payload.companions) ? payload.companions : [];
    var headcount = '';
    var companionsText = '';

    if (!name || !email) {
      return jsonResponse_({ status: 'error', message: 'Name and email are required.' });
    }

    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      return jsonResponse_({ status: 'error', message: 'Please enter a valid email address.' });
    }

    if (name.length > MAX_FIELD_LEN.name || phone.length > MAX_FIELD_LEN.phone ||
        email.length > MAX_FIELD_LEN.email || address.length > MAX_FIELD_LEN.address ||
        notes.length > MAX_FIELD_LEN.notes) {
      return jsonResponse_({ status: 'error', message: 'One of the fields is too long. Please shorten your entry.' });
    }

    if (attending === 'Yes') {
      var companionNames = [];
      for (var c = 0; c < companionsInput.length; c++) {
        var cName = (companionsInput[c].name || '').toString().trim();
        var cPhone = (companionsInput[c].phone || '').toString().trim();
        if (!cName) continue;
        if (cName.length > MAX_FIELD_LEN.companion || cPhone.length > MAX_FIELD_LEN.companion) {
          return jsonResponse_({ status: 'error', message: 'A companion name or phone is too long.' });
        }
        companionNames.push(cPhone ? (cName + ' (' + cPhone + ')') : cName);
      }
      companionsText = companionNames.join('; ');
      headcount = 1 + companionNames.length;

      if (headcount > 20) {
        return jsonResponse_({ status: 'error', message: 'Total party size must be 20 or fewer.' });
      }
    }

    var rsvpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RSVP_SHEET);
    rsvpSheet.appendRow([
      new Date(),
      sanitizeForSheet_(name),
      sanitizeForSheet_(phone),
      sanitizeForSheet_(email),
      sanitizeForSheet_(address),
      attending,
      headcount,
      '',
      '',
      sanitizeForSheet_(notes),
      sanitizeForSheet_(companionsText)
    ]);

    notifyTelegram_(name, attending, headcount, notes);
    sendConfirmationEmail_(name, email, attending);

    return jsonResponse_({ status: 'ok' });
  } finally {
    lock.releaseLock();
  }
}

// Best-effort host notification — never let a Telegram outage break the
// RSVP flow, so failures are swallowed after one log entry.
function notifyTelegram_(name, attending, headcount, notes) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TELEGRAM_BOT_TOKEN');
  var chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chatId) return;

  var lines = ['🎉 New RSVP: ' + name];
  lines.push(attending === 'Yes' ? '✅ Attending (' + headcount + ' pax)' : '❌ Not attending');
  if (notes) lines.push('📝 Notes: ' + notes);

  var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: lines.join('\n') }),
      muteHttpExceptions: true
    });
  } catch (err) {
    console.error('Telegram notification failed: ' + err);
  }
}

// Best-effort guest confirmation — never let a Mail quota/permission issue
// break the RSVP flow itself, so failures are swallowed after one log entry.
function sendConfirmationEmail_(name, email, attending) {
  try {
    var config = getConfigMap_();
    var eventTitle = config.event_title || 'the celebration';
    var dateWords = formatDateWords_(config.event_date);
    var venue = (!config.venue_name || config.venue_name.toUpperCase().indexOf('TBD') !== -1) ? 'To be announced' : config.venue_name;
    var timeRange = (config.start_time || '') + '–' + (config.end_time || '');

    var subject, body;
    if (attending === 'Yes') {
      subject = "You're confirmed for " + eventTitle + "! 🦈";
      body = 'Hi ' + name + ',\n\n' +
        'Salamat for confirming your RSVP! We\'re so excited to celebrate with you.\n\n' +
        'Date: ' + dateWords + '\n' +
        'Time: ' + timeRange + '\n' +
        'Venue: ' + venue + '\n\n' +
        'See you there!\n\nWarmly,\nThe Family';
    } else {
      subject = 'Thanks for letting us know — ' + eventTitle;
      body = 'Hi ' + name + ',\n\n' +
        'Thank you for taking the time to respond. We\'re sorry you can\'t make it, but we truly appreciate you letting us know.\n\n' +
        'We\'ll be thinking of you on ' + dateWords + '!\n\nWarmly,\nThe Family';
    }

    MailApp.sendEmail(email, subject, body);
  } catch (err) {
    console.error('Confirmation email failed: ' + err);
  }
}
