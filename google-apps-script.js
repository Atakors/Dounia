/**
 * Google Apps Script for Aidh Dounia (عيظ دنيا) Election Campaign Form
 * Section: تواصل مباشر مع المترشحة (#citizen-voice)
 *
 * HOW TO DEPLOY:
 * 1. Create a new Google Sheet (e.g., "انشغالات المواطنين - الحملة الانتخابية عيظ دنيا").
 * 2. In the top menu, click Extensions (الإضافات) -> Apps Script.
 * 3. Delete any existing code and paste this entire file.
 * 4. Click Save (💾), then click "Deploy" (نشر) -> "New deployment" (نشر جديد).
 * 5. Select type: "Web app" (تطبيق ويب):
 *    - Execute as: "Me" (أنا)
 *    - Who has access: "Anyone" (أي شخص)
 * 6. Click "Deploy", authorize access, and copy the Web App URL (https://script.google.com/macros/s/.../exec).
 * 7. Paste that URL into `GOOGLE_SHEETS_WEBAPP_URL` inside `index.html` (or tell the agent your URL).
 */

const SHEET_NAME = 'انشغالات_المواطنين';

function handleRequest(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = doc.getSheetByName(SHEET_NAME);

    // Create sheet and Arabic RTL header row if it doesn't exist yet
    if (!sheet) {
      sheet = doc.insertSheet(SHEET_NAME);
      sheet.setRightToLeft(true);
      const headers = [
        'التاريخ والوقت',
        'الاسم الكامل',
        'رقم الهاتف / التواصل',
        'البلدية (ولاية البليدة)',
        'موضوع الانشغال أو الاقتراح',
        'المترشحة'
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#0B6E4F')
        .setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    let params = (e && e.parameter) ? Object.assign({}, e.parameter) : {};

    // Also parse raw postData contents if sent as JSON or text/plain
    if (e && e.postData && e.postData.contents) {
      try {
        const parsed = JSON.parse(e.postData.contents);
        params = Object.assign(params, parsed);
      } catch (_) {
        // URL-encoded fallback
      }
    }

    const rowData = [
      params.timestamp || new Date().toLocaleString('ar-DZ'),
      params.name || 'مواطن(ة) من البليدة',
      params.phone || 'غير محدد',
      params.municipality || 'غير محدد',
      params.message || '',
      params.candidate || 'عيظ دنيا (رقم 46 - القائمة 13)'
    ];

    sheet.appendRow(rowData);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: 'Row added successfully' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return handleRequest(e);
}
