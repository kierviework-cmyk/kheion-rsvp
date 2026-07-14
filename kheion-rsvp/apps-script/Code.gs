/**
 * Kheion's Birthday & Christening RSVP — Apps Script backend.
 * Bound to the "Kheion RSVP" Google Sheet. Deployed as a Web App.
 */

var RSVP_SHEET = 'RSVPs';
var GIFTS_SHEET = 'Gifts';
var CONFIG_SHEET = 'Config';

var RSVP_HEADERS = ['Timestamp', 'Full name', 'Phone', 'Email', 'Address', 'Attending', 'Headcount', 'Godparent volunteer', 'Gift claimed', 'Notes'];
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

  // Remove the default "Sheet1" if it's still empty and unused.
  var sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && sheet1.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }
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
    var godparent = (payload.godparent || 'None').toString().trim();
    var gift = (payload.gift || '').toString().trim();
    var notes = (payload.notes || '').toString().trim();
    var headcount = '';

    if (!name || !phone) {
      return jsonResponse_({ status: 'error', message: 'Name and phone are required.' });
    }

    if (attending === 'Yes') {
      headcount = parseInt(payload.headcount, 10);
      if (!headcount || headcount < 1 || headcount > 20) {
        return jsonResponse_({ status: 'error', message: 'Headcount must be between 1 and 20.' });
      }
    } else {
      godparent = '';
      gift = '';
    }

    // Re-check gift availability inside the lock (someone may have just claimed it).
    if (attending === 'Yes' && gift) {
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
    }

    var rsvpSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(RSVP_SHEET);
    rsvpSheet.appendRow([
      new Date(),
      name,
      phone,
      email,
      address,
      attending,
      headcount,
      godparent,
      attending === 'Yes' ? gift : '',
      notes
    ]);

    return jsonResponse_({ status: 'ok' });
  } finally {
    lock.releaseLock();
  }
}
