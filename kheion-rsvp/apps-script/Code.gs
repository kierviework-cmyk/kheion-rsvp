/**
 * Kheion's Birthday & Christening RSVP — Apps Script backend.
 * Bound to the "Kheion RSVP" Google Sheet. Deployed as a Web App.
 */

var RSVP_SHEET = 'RSVPs';
var GIFTS_SHEET = 'Gifts';
var CONFIG_SHEET = 'Config';
var SUMMARY_SHEET = 'Summary';

var RSVP_HEADERS = ['Timestamp', 'Full name', 'Phone', 'Email', 'Address', 'Attending', 'Headcount', 'Godparent volunteer', 'Gift claimed', 'Notes', 'Companions'];
var GIFTS_HEADERS = ['Gift item', 'Claimed by', 'Claimed at'];
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

var DEFAULT_GIFTS = [
  'Rice cooker',
  'Diapers (size 2)',
  'Baby clothes (12–18 months)',
  'Feeding bottles set',
  'Educational toy set',
  'Baby bath essentials',
  'Cake (sponsor)',
  'Drinks (sponsor)',
  'Balloons & decorations (sponsor)'
];

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

  var giftsSheet = ss.getSheetByName(GIFTS_SHEET) || ss.insertSheet(GIFTS_SHEET);
  if (giftsSheet.getLastRow() === 0) {
    giftsSheet.appendRow(GIFTS_HEADERS);
    giftsSheet.setFrozenRows(1);
    DEFAULT_GIFTS.forEach(function (item) {
      giftsSheet.appendRow([item, '', '']);
    });
  }

  var configSheet = ss.getSheetByName(CONFIG_SHEET) || ss.insertSheet(CONFIG_SHEET);
  if (configSheet.getLastRow() === 0) {
    configSheet.appendRow(CONFIG_HEADERS);
    configSheet.setFrozenRows(1);
    DEFAULT_CONFIG.forEach(function (row) {
      configSheet.appendRow(row);
    });
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

// Live-formula dashboard — every cell recalculates automatically as rows
// are appended to RSVPs/Gifts, no code changes or re-runs needed.
function seedSummarySheet_(sheet) {
  sheet.getRange('A1').setValue('RSVP Summary').setFontWeight('bold').setFontSize(14);

  var stats = [
    ['Total RSVPs', '=COUNTA(RSVPs!B2:B)'],
    ['Attending (Yes)', '=COUNTIF(RSVPs!F2:F,"Yes")'],
    ['Not attending (No)', '=COUNTIF(RSVPs!F2:F,"No")'],
    ['Total headcount', '=SUM(RSVPs!G2:G)'],
    ['Ninong volunteers', '=COUNTIF(RSVPs!H2:H,"Ninong")'],
    ['Ninang volunteers', '=COUNTIF(RSVPs!H2:H,"Ninang")'],
    ['Gifts claimed', '=COUNTA(Gifts!B2:B)'],
    ['Gifts remaining', '=COUNTA(Gifts!A2:A)-COUNTA(Gifts!B2:B)']
  ];
  sheet.getRange(3, 1, stats.length, 2).setValues(stats);
  sheet.getRange(3, 1, stats.length, 1).setFontWeight('bold');

  sheet.getRange('A12').setValue('Godparent volunteers').setFontWeight('bold');
  sheet.getRange('A13').setFormula('=IFERROR(QUERY(RSVPs!B2:H,"select B, H where H=\'Ninong\' or H=\'Ninang\' label B \'Name\', H \'Role\'",0),"None yet")');

  sheet.getRange('D12').setValue('Gifts remaining').setFontWeight('bold');
  sheet.getRange('D13').setFormula('=IFERROR(QUERY(Gifts!A2:A,"select A where A is not null",0),"None left")');

  sheet.getRange('F12').setValue('Gifts claimed').setFontWeight('bold');
  sheet.getRange('F13').setFormula('=IFERROR(QUERY(Gifts!A2:C,"select A, B where B is not null label A \'Item\', B \'Claimed by\'",0),"None yet")');

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

function getGiftsSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GIFTS_SHEET);
}

function getAvailableGifts_() {
  var sheet = getGiftsSheet_();
  var values = sheet.getDataRange().getValues();
  var gifts = [];
  for (var i = 1; i < values.length; i++) {
    var item = values[i][0];
    var claimedBy = values[i][1];
    if (item && !claimedBy) gifts.push(item);
  }
  return gifts;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var action = e && e.parameter && e.parameter.action;

  if (action === 'info') {
    var config = getConfigMap_();
    var availableGifts = getAvailableGifts_();
    return jsonResponse_({ status: 'ok', config: config, availableGifts: availableGifts });
  }

  return jsonResponse_({ status: 'ok', message: 'Kheion RSVP API is running.' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var gotLock = lock.tryLock(10000);

  if (!gotLock) {
    return jsonResponse_({ status: 'error', message: 'Server is busy, please try again in a moment.' });
  }

  try {
    var payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonResponse_({ status: 'error', message: 'Invalid submission format.' });
    }

    var name = (payload.name || '').toString().trim();
    var phone = (payload.phone || '').toString().trim();
    var email = (payload.email || '').toString().trim();
    var address = (payload.address || '').toString().trim();
    var attending = payload.attending === 'Yes' ? 'Yes' : 'No';
    var isGodparent = payload.isGodparent === 'Yes';
    var godparent = isGodparent ? (payload.godparent || '').toString().trim() : '';
    var gift = (payload.gift || '').toString().trim();
    var notes = (payload.notes || '').toString().trim();
    var companionsInput = Array.isArray(payload.companions) ? payload.companions : [];
    var headcount = '';
    var companionsText = '';

    if (!name || !phone) {
      return jsonResponse_({ status: 'error', message: 'Name and phone are required.' });
    }

    if (attending === 'Yes') {
      var companionNames = [];
      for (var c = 0; c < companionsInput.length; c++) {
        var cName = (companionsInput[c].name || '').toString().trim();
        var cPhone = (companionsInput[c].phone || '').toString().trim();
        if (!cName) continue;
        companionNames.push(cPhone ? (cName + ' (' + cPhone + ')') : cName);
      }
      companionsText = companionNames.join('; ');
      headcount = 1 + companionNames.length;

      if (headcount > 20) {
        return jsonResponse_({ status: 'error', message: 'Total party size must be 20 or fewer.' });
      }

      if (isGodparent && godparent !== 'Ninong' && godparent !== 'Ninang') {
        return jsonResponse_({ status: 'error', message: 'Please select Ninong or Ninang.' });
      }
    } else {
      godparent = '';
      gift = '';
    }

    // Re-check gift availability inside the lock (someone may have just claimed it).
    if (attending === 'Yes' && isGodparent && gift) {
      var giftsSheet = getGiftsSheet_();
      var giftValues = giftsSheet.getDataRange().getValues();
      var giftRow = -1;
      for (var i = 1; i < giftValues.length; i++) {
        if (giftValues[i][0] === gift) {
          giftRow = i;
          break;
        }
      }

      var stillAvailable = giftRow !== -1 && !giftValues[giftRow][1];

      if (!stillAvailable) {
        return jsonResponse_({
          status: 'gift_taken',
          message: 'Someone just claimed that gift — please pick another.',
          availableGifts: getAvailableGifts_()
        });
      }

      giftsSheet.getRange(giftRow + 1, 2, 1, 2).setValues([[name, new Date()]]);
    } else if (attending === 'Yes' && isGodparent && !gift) {
      // Gift is required for godparents, unless none are left to choose from.
      if (getAvailableGifts_().length > 0) {
        return jsonResponse_({ status: 'error', message: 'Please select a gift to bring.' });
      }
    }

    var finalGodparent = attending === 'Yes' && isGodparent ? godparent : '';
    var finalGift = attending === 'Yes' && isGodparent ? gift : '';

    var rsvpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RSVP_SHEET);
    rsvpSheet.appendRow([
      new Date(),
      name,
      phone,
      email,
      address,
      attending,
      headcount,
      finalGodparent,
      finalGift,
      notes,
      companionsText
    ]);

    notifyTelegram_(name, attending, headcount, finalGodparent, finalGift, notes);

    return jsonResponse_({ status: 'ok' });
  } finally {
    lock.releaseLock();
  }
}

// Best-effort host notification — never let a Telegram outage break the
// RSVP flow, so failures are swallowed after one log entry.
function notifyTelegram_(name, attending, headcount, godparent, gift, notes) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('TELEGRAM_BOT_TOKEN');
  var chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chatId) return;

  var lines = ['🎉 New RSVP: ' + name];
  lines.push(attending === 'Yes' ? '✅ Attending (' + headcount + ' pax)' : '❌ Not attending');
  if (godparent) lines.push('🙏 Godparent: ' + godparent);
  if (gift) lines.push('🎁 Gift: ' + gift);
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
