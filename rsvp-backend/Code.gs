/**
 * Wonder Women Utsav: RSVP backend (Google Apps Script)
 * ------------------------------------------------------
 * What it does, each time someone submits the RSVP form on the website:
 *   1. Saves the visitor to the "RSVPs" tab of this Google Sheet
 *   2. Emails them a confirmation (if they ticked the email box)
 *
 * The same phone number registering twice is not added again.
 */

// ====== EVENT DETAILS (edit these for future editions) ======
var EVENT = {
  name:      'Wonder Women Utsav',
  dateLabel: 'Saturday, 31 October 2026',
  timeLabel: '10 AM to 8 PM',
  venue:     'Sunder Bhawan, Community Centre',
  area:      'Sunder Vihar, New Delhi',
  mapsUrl:   'https://www.google.com/maps/dir/?api=1&destination=Community+Centre+Sunder+Vihar+New+Delhi&destination_place_id=ChIJUUZ6i6YFDTkReUAE_jUnqY4',
  contact:   '+91 76782 73883',
  whatsapp:  'https://wa.me/917678273883',
  instagram: 'https://instagram.com/wonder_women_utsav'
};

// The guest-list Google Sheet ("Wonder Women Utsav – RSVPs"). Leave empty if this
// script was opened from the Sheet via Extensions > Apps Script.
var SHEET_ID = '1UsuiqrwG1os7mkBJReAaHIluVX5fymxOwPG2S6vkbvU';
var SHEET_NAME = 'RSVPs';
var HEADERS = ['Timestamp', 'Name', 'Phone', 'Email', 'Guests',
               'Email updates OK', 'Email status', 'Source'];
var COL = { name: 2, phone: 3, email: 4, guests: 5, consent: 6, emailStatus: 7 };

// ====== WEB APP ENTRY POINTS ======

function doGet() {
  return json_({ ok: true, service: 'Wonder Women Utsav RSVP', time: new Date().toISOString() });
}

function doPost(e) {
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // Honeypot: real visitors never see or fill this field.
    if (data.website) return json_({ ok: true });

    var name = clean_(data.name, 80);
    var phone = normalisePhone_(data.phone);
    var email = clean_(data.email, 120).toLowerCase();
    var guests = String(data.guests || '1').replace(/[^0-9+]/g, '').slice(0, 2) || '1';
    var consent = data.emailConsent !== false;

    if (name.length < 2) return json_({ ok: false, error: 'Please enter your name.' });
    if (!phone) return json_({ ok: false, error: 'Please enter a valid 10-digit Indian mobile number.' });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) {
      return json_({ ok: false, error: 'Please enter a valid email address.' });
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    var sheet, row, duplicate = false;
    try {
      sheet = getSheet_();
      row = findRowByPhone_(sheet, phone);
      if (row) {
        duplicate = true;
        // Fill in an email if they didn't give one the first time.
        if (!sheet.getRange(row, COL.email).getValue()) sheet.getRange(row, COL.email).setValue(email);
      } else {
        sheet.appendRow([new Date(), name, "'+91 " + phone, email, guests,
                         consent ? 'Yes' : 'No', '', clean_(data.source, 200)]);
        row = sheet.getLastRow();
      }
    } finally {
      lock.releaseLock();
    }

    var person = { name: name, first: name.split(' ')[0], phone: phone, email: email, guests: guests };

    sheet.getRange(row, COL.emailStatus).setValue(consent ? sendEmail_(person) : 'Not sent (box unticked)');

    return json_({ ok: true, duplicate: duplicate });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'Something went wrong on our side. Please try again in a minute.' });
  }
}

// ====== EMAIL ======

function sendEmail_(p) {
  try {
    var subject = 'You’re registered for ' + EVENT.name + ' ✅';
    var html =
      '<div style="background-color:#050505;background-image:linear-gradient(#050505,#050505);padding:32px 16px;font-family:Arial,Helvetica,sans-serif">' +
      '<div style="max-width:520px;margin:0 auto;background:#F7F3EA;border-radius:4px;overflow:hidden">' +
        // The header is one picture with its black background baked in, so dark-mode
        // apps (Gmail on iPhone flips colours) can't turn it white and hide the logo.
        '<div style="background-color:#050505;background-image:linear-gradient(#050505,#050505);border-bottom:3px solid #C7AE32;line-height:0;font-size:0">' +
          '<img src="cid:wwuLogo" width="520" alt="Wonder Women Utsav" style="display:block;border:0;width:100%;max-width:520px;height:auto;margin:0 auto">' +
        '</div>' +
        '<div style="padding:28px">' +
          '<p style="font-size:12px;letter-spacing:3px;text-transform:uppercase;color:#8a7a1f;margin:0 0 8px">&#10004; You&rsquo;re registered</p>' +
          '<h1 style="font-family:Georgia,serif;font-size:26px;color:#171717;margin:0 0 14px">See you there, ' + esc_(p.first) + '!</h1>' +
          '<p style="font-size:15px;color:#4a4740;line-height:1.6;margin:0 0 20px">Thank you for confirming your visit to ' + EVENT.name +
            '. Entry is free. Just walk in and enjoy the shopping, food and celebration.</p>' +
          '<table style="width:100%;font-size:14px;color:#171717;border-collapse:collapse">' +
            row_('Date', EVENT.dateLabel) + row_('Time', EVENT.timeLabel) +
            row_('Venue', EVENT.venue + '<br>' + EVENT.area) + row_('Guests', p.guests) +
          '</table>' +
          '<p style="margin:24px 0 0;text-align:center"><a href="' + EVENT.mapsUrl + '" style="display:inline-block;background:#C7AE32;color:#050505;text-decoration:none;font-weight:bold;font-size:13px;letter-spacing:2px;text-transform:uppercase;padding:14px 26px;border-radius:3px">Get directions</a></p>' +
          '<a href="' + EVENT.whatsapp + '" style="display:block;margin:26px 0 0;padding:14px 16px;border:1px solid #cfe9d6;background:#effaf2;border-radius:4px;text-decoration:none">' +
            '<table role="presentation" cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tr>' +
              '<td style="vertical-align:middle;padding-right:12px"><img src="cid:waIcon" width="36" height="36" alt="WhatsApp" style="display:block;border:0"></td>' +
              '<td style="vertical-align:middle;font-size:14px;line-height:1.4;color:#171717"><strong>Questions? Chat with us on WhatsApp</strong><br>' +
              '<span style="color:#1f7a45">' + EVENT.contact + '</span></td>' +
            '</tr></table></a>' +
          '<p style="font-size:13px;color:#6b6862;line-height:1.6;margin:18px 0 0">Follow <a href="' + EVENT.instagram + '" style="color:#8a7a1f">@wonder_women_utsav</a> for updates.</p>' +
        '</div>' +
      '</div></div>';
    var text = 'See you there, ' + p.first + '!\n\nYour visit to ' + EVENT.name + ' is confirmed.' +
      '\n\n' + EVENT.dateLabel + ', ' + EVENT.timeLabel + '\n' + EVENT.venue + ', ' + EVENT.area +
      '\nDirections: ' + EVENT.mapsUrl + '\n\nQuestions? WhatsApp us on ' + EVENT.contact + ': ' + EVENT.whatsapp;
    MailApp.sendEmail({ to: p.email, subject: subject, body: text, htmlBody: html, name: EVENT.name,
                        inlineImages: { wwuLogo: imageBlob_(LOGO_PNG, 'wwu-logo.jpg'), waIcon: imageBlob_(WHATSAPP_PNG, 'whatsapp.png') } });
    return 'Sent';
  } catch (err) {
    console.error('Email failed', err);
    return 'Failed: ' + String(err.message || err).slice(0, 80);
  }
}

function row_(k, v) {
  return '<tr><td style="padding:8px 0;border-top:1px solid #e3dccb;color:#8f8b82;width:80px;vertical-align:top">' + k +
         '</td><td style="padding:8px 0;border-top:1px solid #e3dccb;font-weight:bold">' + v + '</td></tr>';
}

// ====== HELPERS ======

function getSheet_() {
  var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#050505').setFontColor('#C7AE32');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, HEADERS.length, 150);
  }
  // Keep the Phone column as plain text so "+91 98..." is never read as a number or formula.
  sheet.getRange('C:C').setNumberFormat('@');
  return sheet;
}

function findRowByPhone_(sheet, phone) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  var phones = sheet.getRange(2, COL.phone, last - 1, 1).getValues();
  for (var i = 0; i < phones.length; i++) {
    if (String(phones[i][0]).replace(/\D/g, '').slice(-10) === phone) return i + 2;
  }
  return 0;
}

function normalisePhone_(raw) {
  var d = String(raw || '').replace(/\D/g, '');
  if (d.length === 12 && d.indexOf('91') === 0) d = d.slice(2);
  else if (d.length === 11 && d.charAt(0) === '0') d = d.slice(1);
  return /^[6-9]\d{9}$/.test(d) ? d : null;
}

function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);
}

function esc_(s) {
  return String(s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run this once from the editor (select it, click Run) to create the RSVPs tab,
 * grant permissions, and send a test email to yourself.
 */
function setup() {
  var sheet = getSheet_();
  // Refresh the heading row while the list is still empty.
  if (sheet.getLastRow() <= 1) {
    sheet.clear();
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#050505').setFontColor('#C7AE32');
    sheet.setFrozenRows(1);
    sheet.getRange('C:C').setNumberFormat('@');
  }
  var me = Session.getActiveUser().getEmail();
  if (me) {
    sendEmail_({ name: 'Test Visitor', first: 'Test', phone: '9999999999', email: me, guests: '1' });
    console.log('Setup done. Test confirmation email sent to ' + me);
  } else {
    console.log('Setup done.');
  }
}

/** Deletes the setup test rows (names starting with "TEST Entry"). Safe to run any time. */
function removeTestRows() {
  var sheet = getSheet_();
  for (var r = sheet.getLastRow(); r >= 2; r--) {
    if (String(sheet.getRange(r, COL.name).getValue()).indexOf('TEST Entry') === 0) sheet.deleteRow(r);
  }
}

// ====== EMAIL IMAGES (embedded so they show without "load images") ======
function imageBlob_(b64, name) {
  var type = /\.jpe?g$/.test(name) ? 'image/jpeg' : 'image/png';
  return Utilities.newBlob(Utilities.base64Decode(b64), type, name);
}
var LOGO_PNG = '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';
var WHATSAPP_PNG = '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';

/** One-time: line up rows saved before the switch from WhatsApp to email. Safe to run again. */
function fixOldRows() {
  var sh = getSheet_();
  if (String(sh.getRange(1, 9).getValue()) !== 'Source') return;
  for (var r = 2; r <= sh.getLastRow(); r++) {
    var src = sh.getRange(r, 9).getValue();
    sh.getRange(r, 8).setValue(src);
    sh.getRange(r, 9).clearContent();
  }
  sh.getRange(1, 6, 1, 3).setValues([['Email updates OK', 'Email status', 'Source']]);
  sh.getRange(1, 9).clear();
}
