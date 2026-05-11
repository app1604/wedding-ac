// Google Apps Script — RSVP handler
// Paste this into Extensions → Apps Script in your Google Sheet.
// Deploy as: Web app · Execute as: Me · Access: Anyone

const HOST_EMAIL = 'alharkeml@gmail.com';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('RSVP');
    if (!sheet) {
      sheet = ss.insertSheet('RSVP');
      sheet.appendRow([
        'Timestamp', 'Име', 'Присъствие', 'Придружител',
        'Име на придружителя', 'Ястие', 'Алкохол', 'Пожелание',
      ]);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date().toLocaleString('bg-BG', { timeZone: 'Europe/Sofia' }),
      data.name        || '',
      data.attendance  || '',
      data.plusOne     || '',
      data.plusOneName || '',
      data.food        || '',
      data.drink       || '',
      data.message     || '',
    ]);

    const subject = `RSVP: ${data.name} — ${data.attendance}`;
    const body = [
      'Ново RSVP потвърждение',
      '',
      `Име:             ${data.name}`,
      `Присъствие:      ${data.attendance}`,
      `Придружител:     ${data.plusOne}${data.plusOneName ? ' (' + data.plusOneName + ')' : ''}`,
      `Ястие:           ${data.food  || '—'}`,
      `Алкохол:         ${data.drink || '—'}`,
      `Пожелание:       ${data.message || '—'}`,
      '',
      `Изпратено на: ${data.submittedAt}`,
    ].join('\n');

    MailApp.sendEmail(HOST_EMAIL, subject, body);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
