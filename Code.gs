const SHEET_NAME = 'Guestbook';
const NOTIFY_EMAIL = 'abdohisham237@gmail.com';

function doPost(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || 'save');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Timestamp', 'Guest name', 'Message']);
  }

  if (action === 'save') {
    const name = String(p.name || '').trim().slice(0, 70);
    const message = String(p.message || '').trim().slice(0, 700);
    if (!name || !message) return HtmlService.createHtmlOutput('Missing name or message.');

    // Save first so the guestbook keeps the message even if email delivery fails.
    sh.appendRow([new Date(), "'" + name, "'" + message]);

    let emailSent = true;
    try {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'New wedding wish from ' + name,
        body: 'You received a new message in the Abdulrahman & Nadine wedding guestbook.\n\nGuest: ' + name + '\n\nMessage:\n' + message + '\n\nReceived: ' + new Date().toLocaleString()
      });
    } catch (err) {
      emailSent = false;
      console.error('Email notification failed: ' + err);
    }

    return HtmlService.createHtmlOutput(emailSent
      ? '<p>Thank you! Your message has been received.</p>'
      : '<p>Your message was saved. Email notification could not be sent right now.</p>');
  }

  if (action === 'admin') {
    const supplied = String(p.password || '');
    const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
    if (!expected || supplied !== expected) return HtmlService.createHtmlOutput('<h2>Incorrect password</h2>');
    const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 3).getDisplayValues() : [];
    const esc = s => String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
    const body = rows.reverse().map(r => `<article><small>${esc(r[0])}</small><h3>${esc(r[1].replace(/^'/,''))}</h3><p>${esc(r[2].replace(/^'/,''))}</p></article>`).join('') || '<p>No messages yet.</p>';
    return HtmlService.createHtmlOutput(`<!doctype html><html><meta name="viewport" content="width=device-width,initial-scale=1"><title>Wedding Guestbook</title><style>body{font:16px Arial;background:#fffaf4;color:#4d342d;padding:20px;max-width:760px;margin:auto}h1{color:#9d7138}article{background:white;border:1px solid #dfc58e;padding:15px;margin:12px 0}small{color:#987}</style><h1>Abdulrahman &amp; Nadine — Guestbook</h1><p>${rows.length} messages</p>${body}</html>`);
  }
  return HtmlService.createHtmlOutput('Unknown action.');
}
