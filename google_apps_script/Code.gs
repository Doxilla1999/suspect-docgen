// วิธีติดตั้ง: ดูขั้นตอนละเอียดใน google_apps_script/README.md
// ตั้งรหัสผ่านทีมตรงนี้ก่อน deploy
const PIN = 'CHANGE_ME';

const SHEET_NAME = 'บันทึกจับกุม';
const HEADER_ROW = ['วันที่บันทึก', 'ชื่อ', 'นามสกุล', 'เลขบัตรประชาชน', 'ตำบลที่จับ'];

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADER_ROW);
  }
  return sheet;
}

function doPost(e) {
  const respond = (obj) =>
    ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);

  try {
    const data = JSON.parse(e.postData.contents);
    if (data.pin !== PIN) {
      return respond({ ok: false, error: 'รหัสผ่านไม่ถูกต้อง' });
    }
    if (!data.firstName && !data.lastName) {
      return respond({ ok: false, error: 'ไม่มีชื่อผู้ถูกจับกุม' });
    }
    const sheet = getOrCreateSheet_();
    sheet.appendRow([
      new Date(),
      data.firstName || '',
      data.lastName || '',
      data.idNumber || '',
      data.tambon || ''
    ]);
    return respond({ ok: true });
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  }
}
