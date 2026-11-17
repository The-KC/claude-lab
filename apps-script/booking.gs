/**
 * 小火箭 / 炫風足球俱樂部｜預約體驗課 → Google 試算表
 *
 * 安裝方式：
 * 1. 開一個新的 Google 試算表
 * 2. 選單「擴充功能 → Apps Script」，把這整個檔案貼上（取代原本內容）並儲存
 * 3. 右上角「部署 → 新增部署作業」→ 類型選「網頁應用程式」
 *    - 執行身分：我
 *    - 誰可以存取：所有人
 * 4. 按「部署」並完成授權，複製「網頁應用程式網址」（https://script.google.com/macros/s/.../exec）
 * 5. 把網址填到 index.html 的 CONFIG.sheet
 *
 * 之後修改這個檔案，要用「部署 → 管理部署作業 → 編輯 → 版本：新版本」才會生效（網址不變）。
 */

const SHEET_NAME = '預約名單';
const NOTIFY_EMAIL = true; // 有新預約時寄 Email 通知試算表擁有者；不需要就改成 false
const HEADERS = ['送出時間', '家長', '電話', '小朋友', '年齡', '隊伍', '建議班別', '場地', '時段', '處理狀態', '備註'];
const COL_WIDTHS = [150, 100, 120, 110, 70, 90, 100, 130, 140, 100, 240];

// 處理狀態選項：[名稱, 底色, 字色]
const STATUS = [
  ['待聯絡', '#FDE2E4', '#B00C18'],
  ['已聯絡', '#FFF4CC', '#7A5B00'],
  ['已排體驗', '#DCEBFF', '#1A4E9B'],
  ['已體驗', '#ECE6FF', '#4B32A8'],
  ['已報名', '#D9F2E3', '#1B6B3A'],
  ['不考慮', '#EEEEEE', '#777777']
];

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);

    // 防垃圾：隱藏欄位被填寫就當作機器人，假裝成功但不寫入
    if (d.website) return json({ ok: true });

    const required = ['parent', 'phone', 'child', 'age', 'team', 'day'];
    if (required.some(k => !String(d[k] || '').trim())) return json({ ok: false, error: 'missing' });

    const row = [
      new Date(),
      clean(d.parent),
      "'" + clean(d.phone).replace(/^'/, ''), // 保留電話開頭的 0
      clean(d.child),
      clean(d.age),
      clean(d.team),
      clean(d.group),
      clean(d.venue),
      clean(d.slot),
      '待聯絡'
    ];

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      getSheet().appendRow(row);
    } finally {
      lock.releaseLock();
    }

    if (NOTIFY_EMAIL) notify(row);
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

// 用瀏覽器打開網址時顯示，方便確認部署成功
function doGet() {
  return json({ ok: true, message: '預約表單接收服務運作中' });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    formatSheet(sheet);
  }
  return sheet;
}

/**
 * 美化試算表：在 Apps Script 編輯器上方選擇「setupSheet」後按「執行」。
 * 可以重複執行，不會影響已經存在的預約資料。
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet();
  formatSheet(sheet);

  // 刪掉沒用到的預設空白分頁（工作表1 / Sheet1）
  ss.getSheets().forEach(s => {
    if (s.getName() !== SHEET_NAME && s.getLastRow() === 0 && /^(工作表|Sheet)\d+$/.test(s.getName())) ss.deleteSheet(s);
  });
  ss.setActiveSheet(sheet);
}

function formatSheet(sheet) {
  const cols = HEADERS.length;
  if (sheet.getMaxColumns() < cols) sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
  const rows = sheet.getMaxRows();
  const all = sheet.getRange(1, 1, rows, cols);
  const body = sheet.getRange(2, 1, rows - 1, cols);
  const col = c => sheet.getRange(2, c, rows - 1, 1);

  // 標題列
  sheet.getRange(1, 1, 1, cols).setValues([HEADERS])
    .setFontWeight('bold').setFontColor('#FFFFFF').setFontSize(11)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 38);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2); // 往右捲時「送出時間、家長」保持可見

  // 欄寬、列高、對齊
  COL_WIDTHS.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setRowHeights(2, rows - 1, 30);
  body.setVerticalAlignment('middle').setFontSize(11).setWrap(false);
  sheet.getRange(2, 5, rows - 1, 6).setHorizontalAlignment('center'); // 年齡～處理狀態置中
  col(11).setWrap(true);                                               // 備註可換行
  col(1).setNumberFormat('yyyy/mm/dd hh:mm');

  // 斑馬紋（深色標題 + 白 / 淺灰交錯）
  sheet.getBandings().forEach(b => b.remove());
  all.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor('#0B0B0D').setFirstRowColor('#FFFFFF').setSecondRowColor('#F7F7F8');

  // 處理狀態：下拉選單 + 依狀態上色
  col(10).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(STATUS.map(s => s[0]), true).setAllowInvalid(false).build())
    .setFontWeight('bold');
  sheet.setConditionalFormatRules(STATUS.map(([name, bg, fg]) =>
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(name)
      .setBackground(bg).setFontColor(fg).setRanges([col(10)]).build()));

  // 篩選按鈕（可依隊伍、狀態篩選或排序）
  if (sheet.getFilter()) sheet.getFilter().remove();
  all.createFilter();
}

// 限制長度，並避免內容被試算表當成公式執行
function clean(v) {
  const s = String(v == null ? '' : v).trim().slice(0, 100);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function notify(row) {
  const to = Session.getEffectiveUser().getEmail();
  if (!to) return;
  const body = HEADERS.slice(1, 9).map((h, i) => h + '：' + String(row[i + 1]).replace(/^'/, '')).join('\n');
  MailApp.sendEmail(to, '【新預約】' + row[5] + '・' + row[3], body + '\n\n試算表：' + SpreadsheetApp.getActiveSpreadsheet().getUrl());
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
