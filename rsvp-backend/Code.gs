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
      '<div style="background:#050505;padding:32px 16px;font-family:Arial,Helvetica,sans-serif">' +
      '<div style="max-width:520px;margin:0 auto;background:#F7F3EA;border-radius:4px;overflow:hidden">' +
        '<div style="background:#050505;padding:18px 28px;border-bottom:3px solid #C7AE32">' +
          '<img src="cid:wwuLogo" width="150" height="83" alt="Wonder Women Utsav" style="display:block;border:0;width:150px;height:auto">' +
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
                        inlineImages: { wwuLogo: imageBlob_(LOGO_PNG, 'wwu-logo.png'), waIcon: imageBlob_(WHATSAPP_PNG, 'whatsapp.png') } });
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
  return Utilities.newBlob(Utilities.base64Decode(b64), 'image/png', name);
}
var LOGO_PNG = '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';
var WHATSAPP_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAEgAAABICAYAAABV7bNHAAAVSUlEQVR42u2ce5idVX3vP7+13ndfZvZM5pJ7goHcSaYhkiImCkNEfUS8lNaAetTKRSm1nMdCLeDxNORptbRFatWqhBLxtM9jS6BaEeS02DAgSKCSEEMgJCEhmQSSTGYmM7Ov77vW7/yxL7mQhOw9QbGHNU/+yHr2fvda3/X93X/rFV6PoUj3w8ttzwV4ZIWvTs/T5Ynol3vOV2MmiOpHFBTVt5qEPV1LThGR8vdVJWHFl9wORNYJiIrcI97vDX9r8iObZEXp0G8tN90PY3ouWOEQ9FRvRU7583SZQVa76sTs569ZoKofwMhZeD0H1TNMMqj9uhYdGvlXr0RBQoMkLdVt+2IMItsx8hRenxGRH78w99sbDn1nmUVWezh1QMmpY0y37VnaEwNMffyP0+nO4mWI/7Co/I6kA/BaBqPoFFQR0cp3DXKcdSiK4CusEhCRpBVJWjCC5mNU9Ieo+bf8geS/9C752zxA95ruoOeCnlPCqNEDdPcyy6VlxszbuCwRheN+H+/+xDYnZqtT/HAJFY1REREVEDPK0/CqooiqqASmJYFYwWVLL2DsrWG0/3ubulaXjl7brwegNd0BS3vimS9cmxRX+rQY8z9NKpjnizE+HzsRAcHyeg7FqSomHViTDPCFeJN6/3W1ibu2zv5GsbrGXy1AutzAzYqIzl7/2XdqyLdtS6rL52N8MXICgoyWKfWuSb2CmmRoTTrADRc2SsQ1Lyxc+bOyeN4shxuMkx2mIZGSFR4R5my++i9J2UdNIuiKB/OxFiMvIvZXDg6AiBERq8XIx4P52CSCLlL20Tmbr/5LREBWeO5eZl9XBnWv6Q56lvbEs35x5ZnSlvwbkw4vdn05X10gb6Sh6gHs2Cbj89H9Olj8wpZFdz5X3cMpZ1D3muVlcNZd+TZpCnsksBfHffkIEfOGA6d6YCIm7stHEtiLpSnsmbXuyrf1LO2Ju9csD04pg7p1edAjK+IZGz97nU3avyFS4yPnxIjlN2CoV2dCawnFu6L7wraulbdV9zRqBnWvKT9o9oarrgtbk1/VkkNj539TwAEQI1Zj57XkCFuTX5294arremTFSTFJToo566+6LuxIf9UNFWO82lpI8Dq69Pr66CXFiLOtySDqz1+/beE/vCaTjr/RstvuZq6/4q22KfW0j5zD6fG93jrBEAxVmJ16tPJX/YRBsJhyOKKg+FMDmqJY8Sa01uUKZ29duGpdda8nL2K63CCr3fSnr55l0skHvfMe52W04AiCFYNHGfF5+uMRBuMsHk8oAWlJ0mSSJCVAUYZ8jv54mGGfx6FYMZjRno8gOC/eeW/SyQenP331LGS1K/t2J8OgSiS+e0q/lah4n80k3xMPFkalkAXBiFDwJfK+RHuQoSs1jUXNM5mfegtTE2PpDFpImwSCkPclBtwIvaU+ni/08ovcVp7J7aAvPkjShDSZJF4PZ1xjijtoS1k3UvwPDZMfnLK7wx0rIyDHNucr4lkbPnNLMLb5hmj/SCTGhI0uxIqh5GOyvsCs1GSWtb+T94/5bWYmJ9X1nJ2l/Tw49Avu7v8Zz+Z3kjIhKZPAqR8FSD4Kx2XCuC/7V1sW3HFjde/HB6gS3M169rPnSiJ4TIuxogSN6hmD4aDPMins4Oqx7+OyjvMZY5sAiNVhTkJkFMWpJ6gQOOeL/HDwCf5+3/1sK71Mm2lGK59rkN6xJAPRUvyOLfNXrj06wJXDRQuWy8wt/aEUC0+bVDjP5yPfiBMoSFmHuDy/276YL026lMlhJwCROsLKZiN1bCu+zPbiXl6O+hl2eRSl2aaYGLQzLTmeWcnJpE2iBqqVsp3oj4e5de8P+N6Bn5I2SQJMY2pc1Zt0aHwh2qTJ1NlbZ3VEsEKroiZHhxEzN3zmiqAjfWfcX4hF6mePQYjVEeP50sTLuGrcewEoaUxCyo/bmH+JHw4+wcPDG9lV2sewy+PwNRuglec0mxSTEx0syZzJh9vOZXHz3Fc9618Hf84Xd3+Poo9JSojHN4ARcdCRCuL+/JVbF9yx6vBwpJrilOXcLHf+fCiZzmR/KSk7XQtOkfqCWYMQ4xHga6d9hovHnEOsHiNlcXs2v5Nv7LuPh4bXM+wKpExIUkLscUjq1BNpTN6XSJqQdzTP49oJH6gBVWXjE9nNXP3SNxlyOZKSqB8kxUvKihbci/mR5t+6cnFrcUUlWyEAy/Ruu1oudTM3/MEVQVt4ZzxQv9WSysmXNOY7b/lD3jdm0REn/bV9/8a39j3AsMvTatNlc6+v7d1ULaCqMuzzBGL5dOeF3DRxGSmToKgRSQn5RW4rn9j+VSIfY8XWrZPUqwvaUzYejK7cuuA7q6qYSFn3KBOe+WRTi0lvMKlguhZjX2/mz4phMM7y11Mv55OdS2vgHHQ5Pr9rJfcffIo2m8GKadjyVEEdcCMsaZ7Lt6d9jklhe+23/n1oHVft+Dppk2xAaauXZGB8IX5x2OcX7D3rH3MgmG6WW0S0xTRdErSlpvtCHDcKzsc7u/lk51IidSQkYNBl+dT223jg4H8xNhiDAF5PlIQ+8ah63OOCVtZmX+CjL/41u0p9JCSgpDHvbX0r147/IAdd9rhie6KIzRfiOGhLTW8xTZcgot0st6Zn9SYFMKIfU6dab5glCAUfcXpyAl+adBm+koUvacw1L32LtdnNjA1aidUhmJpOcWjDjnmkjvYgw5bCHq7a8XUG3AiBWGJ1fH7Chzi7aQZZV8DUmQ8UEdSpGtGPAfSs3qSGS1e7ueuv6CIMLvLZokJ9lqvqIf/JhEtotxmceqwYbnnlHh4aXk9n0EKkDoMw4vOMC8cwJzWFlIQV0BoDKVZHe9DM+vyL/K/d/4ip2OVQAm6c9JGKja7b7Ac+W1TC4KK566/o4tLVzgDEgbnYNIeidap/g5D1Bc5umsGH2s4lVk8olsezz3PH/v9bA8eKYcjl+HjHBTww82Z+MnsFd0y7FiNmVOFCpI7OoIV7Bx7nBwM/J6yw6PxMF90tXQy7PKZOUVOPN82hxIG5uLxHVUHlXPW+7hSsiBB5x+Vj311z/px6bn3lX8sb13JUPuLyLMnM45apn6IzaMFgWJyZy3mZeYz4QgP64tDwqqRNgtv2/bACSHkLl499d4WddR+AqPegci6qYuZxcygi52jBIXXwXRCKPmJachzvbX0rihKI4eHhX7I2u5kWm8bjK/MBN036CAZDrA5fUbYf7Ti/sS0cDhBKk0mwpbCbHwz+HIPB4zkvM485qankfamuDIAIUsZCzpnHzaEpbnxlCQk7QUuuLtNupBx1v735TFptU8103z3wM7ScwMEgFDXi9OR4zkqfgQJWbIUxQndLF2emTiPviqNKY3hVQgm5d+DxsqVTSEjI+S3zKfiozvyeGC05T8JOKG58ZYkR1aliTIg2do6LM3NQIBDL/vggT2Y3kzbJmjfrVOmwGQKxlURZdVOehIR8tON8ihqPKklZZlGSjfmdbC721kR2SfNcAjH1b00VMSYU1akG4ffKB1qfyvfqSZsk81NvqW36mdx29sdDhCY4LDcIBY2PKUYChGI5FQlcI0LW53kyu6U2Nys1hTbbjMPVq1y1QvLfM6gU62W3IBULkmFC2Fabf67QS6SHFlM1u3tKB+iLh2ppCV/JDm4rvszX9v6olgAbfTYVni/sqv1/XDCGsUErkTqof5OgUjTgz9WCK3dZ1PFdh6fVNjHGNtdM9e6o74hlKErCBLwSDfCfQ8+Uv6flqL2oMX+083b2xwcJCUZl7itSQSCWPaX+2m83m2RNP0p9SBstOMCfa0zCnl7pz5F6FWNKEiTkkDgNu/yrlK2qkpSQlX0PkvNFjBi8ekIxnJGcgFNfM82j5Y9BGPb5I2abTaoCvtQlIhp5TMKebsqdXQ1XL1/l3R4951HSNsHG/E6+tf9+DILDYzCsmPw/OC0xlpwvlisYJ8hMysky+6hA2DTqYwloyalp1HwIQqzxEemKpEmUTfwxgswxtom/3/cAj2WfIyEBkTrGBa383WmfxSKUNHqVwyhIrQLiK0GxnGBFHq1lH6uj6EuNV9qkwZp62Z8Rcq5YCQrLyx4btB43WVUOHJXP71zJ7tIBQrFE6licmcvtp1+LFUvWFSrugFQOwJGUkMXNc7EI/fFITcEfi8xOPWODMYcx2jPs8qPy1M0xj/wkam+BWA64YQ7Ew7X5GcmJxxUGjydlEuyJ+rl21+0UfUQgZc/63S1n8f3pX2B6ciJ98RCxxhgRcr7I8skfY/WMG7lnxk18uvNdWLEMxCPHVOoezxnJCTX2DbhhBtwIFtuYCVBVIwnbUCefxTLkcmwv7a3NLUifQdqkjmuyy6LWzOMjz3HD7rtqhcRIHYuaZnLfzP/NTROXMTHsYG80yLXjP8Cy9nfi8HSlp/FXUy/nRzO+xKc73/0qnaVadikWpWfU5nqjA2W/rIEMIwqSsGJ8ye2Q0NRdNxGBSGOezm0tbx7PnNQUZqUmUdDjxz9xJZfzz/2PcOPu70HFWSxpTItt4o8nfJj7Zv0Z98y4iZsmXYpWrJOvlH9mp6Zwy9Tf5x2ZM8lWAl2p6LCpYSdnN8+oFbHXZ18sx2L1qllFJTT4ktthwKyVlKXWTXry7CMhIY+NPFeJ3CEhAe9rPZuCj05oPZx62m2Gu/oe4vLtf8feaLDmLpQ0psNmeFfLgkpoUv4zFbbFlUA370vl8pJSE8cLWxcyxjbX/J5HRjY2pn8ELykLmLUG0WQjIla1GBvy29mU31VOrKNc2nEe44MxlHx8QtfKUQbp34fX8cGtf86/9D+Kq6Rqq6Wdw5saPEqkjkAMe6NBNuR3kDKJWmGxxaT5ZOfSWgp4V6mPJ7Mv0GyS+EZy4AqIJg3Kveqp9CHXn0QfdgVWD/ysZnWmhJ18ovMChnzuNantKJv//fFBPt97B7+77cvcM/AYwy5PQoKa+FQZFIrlQDzMn/Z+l4MuR4CpJeMuaV/M3NTUSqgj3D3wKH3xMIE0oKBVpYwJ9wYq0qveR0j9zQlePRmT4geDT3DF2PfwlsQ4PMo1497Pjw8+RW/pQEV09ITilpCAlAlZl9vOU7nbmZmcxLtaFnBh61nMTk7BiOFgnKVnZCN3HXiIHcV9ZEwKgIIvMSXRwXUTfqfiAgj9boR/7n+kwp6G/B/U+0hFeoNk18THo2df3isJO1VL9ZV7FEiagJejfv5jaD1Xjn1PJUIPCDh5y1EWk3LsJCL0lvq4ff+D3HXgp7SaJqwYcr7IsMuTNCEZky434QsUXIm/mHwNEw8r/3xn3wPsKvXREWSI6xYv9ZIIjJbcnmTXpMfNJm6OVPUpSVm0gYjRaTkXMy99Ws3/2JDfwbbSKyQlrMs4Vi1VQkLagwwpSZDXEsOuHF+1Bc21Z4qU6/NfnHQZF41ZRFQBZ0N+B6v6HqLVNjUADqiiZSz0qU3cHBlEFNG1Ykzd/rhUyjtTws5KxrD89Sezmyn6Uk2HWDGV5idTB6PKCtoitc6OamDr8QzGI9ww8SP80fiLidURiCXrC3yhdxUljRuPwUDFGBBdi4gagCD29/tspFJnIcmIUNASC5umH5HTeTK7hVQlJoo05qDLMRCPMOxzNcDqEeOyL3RIIQvCraddyfUTLim30VQU+Z/2fpcNuR002wYtFyAG47ORBrG/vxxq3L3MPr9w1Uai+CemOSlAXB8llSWZuTWrtjceZG12cy1OGx+M4f1jFnHz5I/z4ba3k/UFhlyu9vlqW121v6/q+1T9nqASoGZ9gYMuy3mZ+dw74yY+0XEBkcYEYjFi+OLu/8M9A4/RHjSPpqkqNs1JIYp/8vzCVRu5e5kNupfNkx7Aq3w/sPJ+VT2pAF8Oi9Lf1jSrNr8pv5M5qSksbVnA2zNzmJ+aVmuaAljW/g5W9T3E2uxm+uMRrBhCsQTYmlvg1RPjiL0jVkfGpDm3eTaf6ryQD7W9reYnJSSgqDE39H6X7/c/QkfQ0pDeOfywxYo4le8DdC+bJw03LxjK3mtXehr3zfyzmtgUfKkmXoeb8sODXIBNhV08MryR/8pu4aXSPvrdSM07TlWU9NTEWBY1zeC8zHwWNk0/Qg8JwpbiHm7ovYvHR56jPciMqh3veM0LAYIu09V29cJ/yrZs+IMvm1RwZ5yP9bXUhEg5bfrbzbMqIUC5pS5pwsomDkVjR+ucWB3zUqcxL3UajLuIokYMuRzFSokmlIBW20RKwiNANhWRLPqIuw78lG/u+zEDbqTCnFFdC0M9alMBvqBf3rvwn7LL9EN2teAabqAyGIZ8ljumXcv7WhfV+niqI+eLvFDYw9O5rTw28hx98UE+1tHNB9rOIWPSR5j2qt452pJVO82qI+sL/GjwSe468BAbcjtoMkkSJhglczhhA1VQLXM8vKbb9i7tyc/c8Jmv2KbwzrjgnBynj7pc1YiYFHawtGVBuZNVY7YVXmF9bhtPZDfzTG47O0t9ZH2BoGLin8xu4Zv77ufC1gW8q2UBXenT6QxajutCCDDgRtiYf4k1w7/koaH1bCnsISEBbUEzXnX04JQtpbdNYRAX4q/0Lvnb/MNrugOWSnxkLb6OJk5ByGuJBenTuWbcRTw68ixP515kR3FvzQwnTUhSgpo/UlX+1V7pUALGh2OYlhjPlEQnnbaFZpNCpNwQ0R+PsDs6wEvFveyNDlLSiJRJkK6kdf2puqxwsk2c5bpxfW3AAgz7Al59BZBD/YbHa/Q+vKUuUkdJY6JKYfFo2xmKJZSAhASIyKibx4+ziRO2AR/JjktXu+41y4Mt81eu1Xx0a9DRFKj30YmcuFbbRHuQIS2JmjKtesEn8pJ9xaI1myTtQYaOIEP7Uf+aTZpQgloIcqrBUe+joKMp0Hx065b5K9d2r1keHH0J+FdyFeGNOE72KsIxygNozwX4rbO/UXRF+ZwvxPvKeetToA3fMOiol4QVX4j3uaJ8buvsbxR7LsAf6579r+U61K8XnFNxHQpAVrtuXR5sXbhqXZwrXW9bEhaDa6RM9AZijmJwtiVh41zp+q0LV63r1uXB8cA5MUBA9dritoX/cJsfzF9vW5MBVvQ3UtxUPVbUtiYDP1i5bbjmte+tvnmp9zUrqycxakzqWnmbjpQWI+y3bSmrSvTGJw6RbUtZhP06Ulq8rWvlSTGnLgbVmPT/4YsF6rdIhzxNmbP56q+omhvFCi4fxQK//pcMlN/h4W06DNQpIv6WzXNu/yKgjbwN5s2Xm7wuAFXHm6/HqUvk3nzB0glO8c1XdJ308958ydvJMuq/x2sC/x/d1Gt/UWSFDwAAAABJRU5ErkJggg==';

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
