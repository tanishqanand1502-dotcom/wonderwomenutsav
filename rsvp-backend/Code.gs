/**
 * Wonder Women Utsav: RSVP backend (Google Apps Script)
 * ------------------------------------------------------
 * What it does, each time someone submits the RSVP form on the website:
 *   1. Saves the visitor to the "RSVPs" tab of this Google Sheet
 *   2. Emails them a confirmation (if they gave an email)
 *   3. Sends a WhatsApp confirmation (once WhatsApp is switched on, see SETUP.md)
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
  mapsUrl:   'https://www.google.com/maps/search/?api=1&query=Sunder+Bhawan+Sunder+Vihar+New+Delhi',
  contact:   '+91 98102 36151',
  instagram: 'https://instagram.com/wonder_women_utsav'
};

// The guest-list Google Sheet ("Wonder Women Utsav – RSVPs"). Leave empty if this
// script was opened from the Sheet via Extensions > Apps Script.
var SHEET_ID = '1UsuiqrwG1os7mkBJReAaHIluVX5fymxOwPG2S6vkbvU';
var SHEET_NAME = 'RSVPs';
var HEADERS = ['Timestamp', 'Name', 'Phone', 'Email', 'Guests',
               'WhatsApp OK', 'Email status', 'WhatsApp status', 'Source'];
var COL = { name: 2, phone: 3, email: 4, guests: 5, consent: 6, emailStatus: 7, waStatus: 8 };

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
    var consent = data.whatsappConsent !== false;

    if (name.length < 2) return json_({ ok: false, error: 'Please enter your name.' });
    if (!phone) return json_({ ok: false, error: 'Please enter a valid 10-digit Indian mobile number.' });
    if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) {
      return json_({ ok: false, error: 'That email address doesn’t look right.' });
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
        if (email && !sheet.getRange(row, COL.email).getValue()) sheet.getRange(row, COL.email).setValue(email);
      } else {
        sheet.appendRow([new Date(), name, '+91 ' + phone, email, guests,
                         consent ? 'Yes' : 'No', '', '', clean_(data.source, 200)]);
        row = sheet.getLastRow();
      }
    } finally {
      lock.releaseLock();
    }

    var person = { name: name, first: name.split(' ')[0], phone: phone, email: email, guests: guests };

    if (email) sheet.getRange(row, COL.emailStatus).setValue(sendEmail_(person));
    if (consent) sheet.getRange(row, COL.waStatus).setValue(sendWhatsApp_(person));
    else sheet.getRange(row, COL.waStatus).setValue('No consent');

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
      '<div style="background:#050505;padding:32px 16px;font-family:Arial,Helvetica,sans-serif">' +
      '<div style="max-width:520px;margin:0 auto;background:#F7F3EA;border-radius:4px;overflow:hidden">' +
        '<div style="background:#050505;padding:22px 28px;border-bottom:3px solid #C7AE32">' +
          '<div style="font-family:Georgia,serif;font-size:22px;color:#F7F3EA">Wonder Women <span style="color:#C7AE32;font-style:italic">Utsav</span></div>' +
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
          '<p style="font-size:13px;color:#6b6862;line-height:1.6;margin:24px 0 0">Questions? Call or WhatsApp us on ' + EVENT.contact +
            ', or follow <a href="' + EVENT.instagram + '" style="color:#8a7a1f">@wonder_women_utsav</a> for updates.</p>' +
        '</div>' +
      '</div></div>';
    var text = 'See you there, ' + p.first + '!\n\nYour visit to ' + EVENT.name + ' is confirmed.' +
      '\n\n' + EVENT.dateLabel + ', ' + EVENT.timeLabel + '\n' + EVENT.venue + ', ' + EVENT.area +
      '\nDirections: ' + EVENT.mapsUrl + '\n\nQuestions? ' + EVENT.contact;
    MailApp.sendEmail({ to: p.email, subject: subject, body: text, htmlBody: html, name: EVENT.name });
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

// ====== WHATSAPP (Meta WhatsApp Cloud API) ======
// Switched off until these Script Properties are set (Project Settings > Script properties):
//   WA_TOKEN            permanent access token from Meta
//   WA_PHONE_NUMBER_ID  the Phone number ID shown in WhatsApp > API Setup
//   WA_TEMPLATE         approved template name (default: rsvp_confirmation)
//   WA_LANG             template language code (default: en)
// The template body must have exactly 1 variable: {{1}} = first name.

function sendWhatsApp_(p) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('WA_TOKEN');
  var phoneId = props.getProperty('WA_PHONE_NUMBER_ID');
  if (!token || !phoneId) return 'Pending (WhatsApp not set up yet)';
  try {
    var body = {
      messaging_product: 'whatsapp',
      to: '91' + p.phone,
      type: 'template',
      template: {
        name: props.getProperty('WA_TEMPLATE') || 'rsvp_confirmation',
        language: { code: props.getProperty('WA_LANG') || 'en' },
        components: [{ type: 'body', parameters: [
          { type: 'text', text: p.first }
        ]}]
      }
    };
    var res = UrlFetchApp.fetch('https://graph.facebook.com/v21.0/' + phoneId + '/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code >= 200 && code < 300) return 'Sent';
    var msg = '';
    try { msg = JSON.parse(res.getContentText()).error.message; } catch (x) { msg = res.getContentText(); }
    return 'Failed (' + code + '): ' + String(msg).slice(0, 80);
  } catch (err) {
    return 'Failed: ' + String(err.message || err).slice(0, 80);
  }
}

/**
 * Run this by hand once WhatsApp is switched on: it sends WhatsApp confirmations
 * to everyone who registered before that (status "Pending...") or whose send failed.
 */
function sendPendingWhatsApp() {
  var sheet = getSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return;
  var rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  var sent = 0;
  rows.forEach(function (r, i) {
    var status = String(r[COL.waStatus - 1]);
    if (r[COL.consent - 1] !== 'Yes' || status === 'Sent') return;
    if (status && status.indexOf('Pending') !== 0 && status.indexOf('Failed') !== 0) return;
    var name = String(r[COL.name - 1]);
    var p = { name: name, first: name.split(' ')[0], phone: String(r[COL.phone - 1]).replace(/\D/g, '').slice(-10) };
    var result = sendWhatsApp_(p);
    sheet.getRange(i + 2, COL.waStatus).setValue(result);
    if (result === 'Sent') sent++;
    Utilities.sleep(300);
  });
  console.log('WhatsApp confirmations sent: ' + sent);
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
  }
  var me = Session.getActiveUser().getEmail();
  if (me) {
    sendEmail_({ name: 'Test Visitor', first: 'Test', phone: '9999999999', email: me, guests: '1' });
    console.log('Setup done. Test confirmation email sent to ' + me);
  } else {
    console.log('Setup done.');
  }
}
