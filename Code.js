/**
 * 職員朝礼伝達システム - Backend Logic (v6.0 Final)
 */

const SHEETS = {
  STAFF: '職員情報',
  DAILY: '行事',
  TRIP: '出張等',
  LEAVE: '休暇等',
  MEETING: '会議',
  ANNOUNCE: '伝達事項',
  ROOM: '特別教室予約',
  FIXED_MEETING: '定例会議',
  FIXED_CLASS: '特別教室固定'
};

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate().setTitle('職員朝礼伝達システム').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * データ取得API
 */
function getData(dateStr) {
  const result = {
    staff: getStaffData(),
    daily: [], // 複数件対応のため配列
    trips: [], leaves: [], meetings: [],
    announcements_staff: [], announcements_student: [], reservations: [],
    counts: { trip: 0, leave: 0, meeting: 0 }
  };

  const targetDate = new Date(dateStr);
  const dayMap = ['日', '月', '火', '水', '木', '金', '土'];
  const dayOfWeek = dayMap[targetDate.getDay()];

  // 1. 行事 (複数登録対応)
  result.daily = getRows(SHEETS.DAILY).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], time: r[2], content: r[3], note: r[4] }));

  // 2. 出張
  result.trips = getRows(SHEETS.TRIP).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], staff_name: r[2], purpose: r[3], location: r[4], time: r[5], note: r[6] }));
  result.counts.trip = result.trips.length;

  // 3. 休暇
  result.leaves = getRows(SHEETS.LEAVE).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], staff_name: r[2], type: r[3], time: r[4], note: r[5] }));
  result.counts.leave = result.leaves.length;

  // 4. 会議 (定例 + 個別)
  const nm = getRows(SHEETS.MEETING).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], name: r[2], time: r[3], place: r[4], is_fixed: false }));
  const fm = getRows(SHEETS.FIXED_MEETING).filter(r => r[0] === dayOfWeek)
    .map(r => ({ id: 'fixed', date: dateStr, name: r[1], time: r[2], place: r[3], is_fixed: true }));
  result.meetings = [...fm, ...nm];
  result.counts.meeting = result.meetings.length;

  // 5. 伝達事項
  const aa = getRows(SHEETS.ANNOUNCE).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], type: r[2], priority: r[3], target: r[4], content: r[5], reporter: r[6] }));
  result.announcements_staff = aa.filter(a => a.type === '職員');
  result.announcements_student = aa.filter(a => a.type === '生徒');

  // 6. 教室予約
  const nr = getRowsFixedCols(SHEETS.ROOM, 6).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ 
      id: r[0], date: r[1], room: r[2], period: r[3], content: r[4], reserver: r[5], is_fixed: false 
    }));
  const fr = getRowsFixedCols(SHEETS.FIXED_CLASS, 5).filter(r => r[0] === dayOfWeek)
    .map(r => ({ 
      id: 'fixed', date: dateStr, period: r[1], room: r[2], content: r[3], reserver: r[4], is_fixed: true 
    }));
  result.reservations = [...fr, ...nr];

  return JSON.stringify(result);
}

/**
 * データ保存API
 */
function saveData(category, data) {
  const id = Utilities.getUuid();
  let sheetName = "";
  let rowData = [];

  switch (category) {
    case 'daily':
      sheetName = SHEETS.DAILY;
      rowData = [id, data.date, data.time, data.content, data.note];
      break;
    case 'trip':
      sheetName = SHEETS.TRIP;
      rowData = [id, data.date, data.staff, data.purpose, data.location, data.time, data.note];
      break;
    case 'leave':
      sheetName = SHEETS.LEAVE;
      rowData = [id, data.date, data.staff, data.type, data.time, data.note];
      break;
    case 'meeting':
      sheetName = SHEETS.MEETING;
      rowData = [id, data.date, data.name, data.time, data.place];
      break;
    case 'announce':
      sheetName = SHEETS.ANNOUNCE;
      rowData = [id, data.date, data.type, data.priority, data.target, data.content, data.reporter];
      break;
    case 'room':
      sheetName = SHEETS.ROOM;
      rowData = [id, data.date, data.room, data.period, data.content, data.reserver];
      break;
  }

  if (sheetName) {
    getSS().getSheetByName(sheetName).appendRow(rowData);
  }
}

/**
 * データ削除API
 */
function deleteEvent(id, category) {
  let sheetName = "";
  if (category === 'room') sheetName = SHEETS.ROOM;
  if (category === 'daily') sheetName = SHEETS.DAILY;
  if (category === 'announce') sheetName = SHEETS.ANNOUNCE;
  
  if (!sheetName) return;

  const sheet = getSS().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

// --- Helper Functions ---
function getRows(name) {
  const sheet = getSS().getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
}

function getRowsFixedCols(sheetName, colCount) {
  const s = getSS().getSheetByName(sheetName);
  if (!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, colCount).getValues();
}

function getStaffData() {
  const sheet = getSS().getSheetByName(SHEETS.STAFF);
  if (!sheet) return [];
  const rows = getRows(SHEETS.STAFF);
  return rows.map(r => ({ id: r[0], name: r[1], role: r[2], order: r[3] }))
             .sort((a, b) => a.order - b.order);
}

function formatDate(dateObj) {
  if (!dateObj) return "";
  return Utilities.formatDate(new Date(dateObj), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}