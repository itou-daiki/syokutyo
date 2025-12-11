/**
 * 職員朝礼伝達システム - Backend Logic (v9.5)
 *
 * @description 教員朝礼で使用する伝達事項・予定管理システム
 * @version 9.5
 */

// =============================================================================
// 定数定義
// =============================================================================

const SHEETS = {
  STAFF: '職員情報',
  DAILY: '行事',
  TRIP: '出張等',
  LEAVE: '休暇等',
  MEETING: '会議',
  ANNOUNCE: '伝達事項',
  ROOM: '特別教室予約',
  FIXED_MEETING: '定例会議',
  FIXED_CLASS: '特別教室固定',
  TASK: 'タスク' // 新規追加
};

const ERROR_MESSAGES = {
  INVALID_DATE: '日付が無効です',
  MISSING_DATA: '必須項目が入力されていません',
  SHEET_NOT_FOUND: 'シートが見つかりません',
  SAVE_FAILED: 'データの保存に失敗しました',
  DELETE_FAILED: 'データの削除に失敗しました',
  UNKNOWN_ERROR: '予期しないエラーが発生しました'
};

// =============================================================================
// ユーティリティ関数 (変更なし)
// =============================================================================
function getSS() {
  try { return SpreadsheetApp.openById(SPREADSHEET_ID); } 
  catch (e) { logError('getSS', e); throw new Error('スプレッドシートを開けませんでした'); }
}
function logError(fn, e) { console.error(`[${fn}] Error:`, e.message, e.stack); }
function logInfo(fn, m) { console.info(`[${fn}] ${m}`); }

// =============================================================================
// Web App エントリーポイント (変更なし)
// =============================================================================
function doGet() {
  try {
    return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('職員朝礼伝達システム').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (e) {
    return HtmlService.createHtmlOutput('システムエラーが発生しました。');
  }
}
function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

// =============================================================================
// バリデーション関数 (変更なし)
// =============================================================================
function isValidDate(d) { if (!d) return false; const date = new Date(d); return date instanceof Date && !isNaN(date.getTime()); }
function validateRequiredFields(data, fields) {
  const missing = fields.filter(f => !data[f] || data[f].toString().trim() === '');
  return { valid: missing.length === 0, missingFields: missing };
}
function isValidCategory(cat) { return ['daily', 'trip', 'leave', 'meeting', 'announce', 'room', 'task'].includes(cat); }

// =============================================================================
// データ取得API
// =============================================================================
function getData(dateStr) {
  try {
    if (!isValidDate(dateStr)) throw new Error(ERROR_MESSAGES.INVALID_DATE);

    const todayData = getDataForDate(dateStr);
    
    // 明日
    const tomorrow = new Date(dateStr);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const tomorrowData = getDataForDate(tomorrowStr);

    // タスクデータ（全件取得し、フロントでフィルタリング）
    const tasks = getTasks();

    return JSON.stringify({
      staff: getStaffData(),
      today: todayData,
      tomorrow: tomorrowData,
      tasks: tasks
    });

  } catch (e) {
    logError('getData', e);
    throw new Error(`データ取得エラー: ${e.message}`);
  }
}

function getDataForDate(dateStr) {
  const result = {
    date: dateStr,
    daily: [], trips: [], leaves: [], meetings: [],
    announcements_staff: [], announcements_student: [], reservations: [],
    counts: { trip: 0, leave: 0, meeting: 0 }
  };

  const targetDate = new Date(dateStr);
  const dayMap = ['日', '月', '火', '水', '木', '金', '土'];
  const dayOfWeek = dayMap[targetDate.getDay()];

  // 1. 行事
  result.daily = getRows(SHEETS.DAILY).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], time: r[2]||'', content: r[3]||'', note: r[4]||'' }));

  // 2. 出張
  result.trips = getRows(SHEETS.TRIP).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], staff_name: r[2]||'', purpose: r[3]||'', location: r[4]||'', time: r[5]||'', note: r[6]||'' }));
  result.counts.trip = result.trips.length;

  // 3. 休暇
  result.leaves = getRows(SHEETS.LEAVE).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], staff_name: r[2]||'', type: r[3]||'', time: r[4]||'', note: r[5]||'' }));
  result.counts.leave = result.leaves.length;

  // 4. 会議
  const nm = getRows(SHEETS.MEETING).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], name: r[2]||'', time: r[3]||'', place: r[4]||'', is_fixed: false }));
  const fm = getRows(SHEETS.FIXED_MEETING).filter(r => r[0] === dayOfWeek)
    .map(r => ({ id: 'fixed', date: dateStr, name: r[1]||'', time: r[2]||'', place: r[3]||'', is_fixed: true }));
  result.meetings = [...fm, ...nm];
  result.counts.meeting = result.meetings.length;

  // 5. 伝達事項
  const aa = getRows(SHEETS.ANNOUNCE).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], type: r[2]||'', priority: r[3]||'', target: r[4]||'', content: r[5]||'', reporter: r[6]||'' }));
  result.announcements_staff = aa.filter(a => a.type === '職員');
  result.announcements_student = aa.filter(a => a.type === '生徒');

  // 6. 教室予約
  const nr = getRowsFixedCols(SHEETS.ROOM, 6).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], room: r[2]||'', period: r[3]||'', content: r[4]||'', reserver: r[5]||'', is_fixed: false }));
  const fr = getRowsFixedCols(SHEETS.FIXED_CLASS, 5).filter(r => r[0] === dayOfWeek)
    .map(r => ({ id: 'fixed', date: dateStr, period: r[1]||'', room: r[2]||'', content: r[3]||'', reserver: r[4]||'', is_fixed: true }));
  result.reservations = [...fr, ...nr];

  return result;
}

function getTasks() {
  // タスク: id, name, roll, content, due_date, check
  return getRows(SHEETS.TASK).map(r => ({
    id: r[0],
    name: r[1] || '',
    roll: r[2] || '',
    content: r[3] || '',
    due_date: formatDate(r[4]),
    check: r[5] === true || r[5] === 'TRUE'
  }));
}

// =============================================================================
// データ保存・更新API (大幅改修)
// =============================================================================
function saveData(category, data) {
  try {
    if (!isValidCategory(category)) throw new Error(`無効なカテゴリ: ${category}`);
    
    // タスク以外は日付チェック
    if (category !== 'task' && !isValidDate(data.date)) throw new Error(ERROR_MESSAGES.INVALID_DATE);

    const isUpdate = !!data.id; // IDがあれば更新モード
    const id = isUpdate ? data.id : Utilities.getUuid();
    
    let sheetName = "";
    let rowData = []; // IDを除くデータ配列（更新用）
    
    // カテゴリごとの設定
    switch (category) {
      case 'daily':
        sheetName = SHEETS.DAILY;
        rowData = [data.date, data.time||'', data.content, data.note||''];
        break;
      case 'trip':
        sheetName = SHEETS.TRIP;
        rowData = [data.date, data.staff, data.purpose, data.location, data.time||'1日', data.note||''];
        break;
      case 'leave':
        sheetName = SHEETS.LEAVE;
        rowData = [data.date, data.staff, data.type, data.time||'1日', data.note||''];
        break;
      case 'meeting':
        sheetName = SHEETS.MEETING;
        rowData = [data.date, data.name, data.time||'放課後', data.place||'会議室'];
        break;
      case 'announce':
        sheetName = SHEETS.ANNOUNCE;
        rowData = [data.date, data.type, data.priority||'・', data.target||'全職員', data.content, data.reporter];
        break;
      case 'room':
        sheetName = SHEETS.ROOM;
        rowData = [data.date, data.room, data.period, data.content, data.reserver||''];
        break;
      case 'task':
        sheetName = SHEETS.TASK;
        // id, name, roll, content, due_date, check
        rowData = [data.name, data.roll||'', data.content, data.due_date||'', data.check||false];
        break;
    }

    const sheet = getSS().getSheetByName(sheetName);
    if (!sheet) throw new Error(`${ERROR_MESSAGES.SHEET_NOT_FOUND}: ${sheetName}`);

    if (isUpdate) {
      // 更新処理: IDを検索して行を特定
      const dataRange = sheet.getDataRange();
      const values = dataRange.getValues();
      let rowIndex = -1;
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] == id) {
          rowIndex = i + 1;
          break;
        }
      }
      
      if (rowIndex > 0) {
        // 行の値を更新（ID列[0]は変更せず、その右側を更新）
        sheet.getRange(rowIndex, 2, 1, rowData.length).setValues([rowData]);
        return { success: true, message: 'データを更新しました', id: id };
      } else {
        // IDが見つからない場合は新規追加扱いにするかエラーにするか。ここでは新規追加に倒す
        const newRow = [id, ...rowData];
        sheet.appendRow(newRow);
        return { success: true, message: 'データが見つからなかったため新規追加しました', id: id };
      }

    } else {
      // 新規追加処理
      const newRow = [id, ...rowData];
      sheet.appendRow(newRow);
      return { success: true, message: 'データを保存しました', id: id };
    }

  } catch (e) {
    logError('saveData', e);
    throw new Error(`${ERROR_MESSAGES.SAVE_FAILED}: ${e.message}`);
  }
}

// タスクのチェック状態のみを切り替える軽量API
function toggleTaskCheck(id, isChecked) {
  try {
    const sheet = getSS().getSheetByName(SHEETS.TASK);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        // checkカラムは6列目(index 5)
        sheet.getRange(i + 1, 6).setValue(isChecked);
        return { success: true };
      }
    }
    throw new Error('タスクが見つかりません');
  } catch (e) {
    throw new Error(e.message);
  }
}

// =============================================================================
// データ削除API (変更なし)
// =============================================================================
function deleteEvent(id, category) {
  try {
    const categoryToSheet = {
      'room': SHEETS.ROOM, 'daily': SHEETS.DAILY, 'announce': SHEETS.ANNOUNCE,
      'trip': SHEETS.TRIP, 'leave': SHEETS.LEAVE, 'meeting': SHEETS.MEETING, 'task': SHEETS.TASK
    };
    const sheetName = categoryToSheet[category];
    const sheet = getSS().getSheetByName(sheetName);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'データを削除しました' };
      }
    }
    throw new Error('データが見つかりません');
  } catch (e) {
    logError('deleteEvent', e);
    throw new Error(`${ERROR_MESSAGES.DELETE_FAILED}: ${e.message}`);
  }
}

// =============================================================================
// ヘルパー関数 & 天気API (変更なし)
// =============================================================================
function getRows(name) {
  try {
    const sheet = getSS().getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  } catch (e) { return []; }
}
function getRowsFixedCols(sheetName, colCount) {
  try {
    const sheet = getSS().getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, colCount).getValues();
  } catch (e) { return []; }
}
function getStaffData() {
  try {
    const rows = getRows(SHEETS.STAFF);
    return rows.map(r => ({ id: r[0]||'', name: r[1]||'', role: r[2]||'', order: r[3]||999 }))
      .filter(s => s.name).sort((a, b) => a.order - b.order);
  } catch (e) { return []; }
}
function formatDate(dateObj) {
  try {
    if (!dateObj) return "";
    const date = new Date(dateObj);
    if (isNaN(date.getTime())) return "";
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) { return ""; }
}
function getWeatherData() {
  // ... (天気APIコードは前回と同じため省略。もし必要なら前回のをそのまま使う)
  // 今回の要件には変更がないため、既存のままでOK
  try {
    const HITA_LAT = 33.3219; const HITA_LON = 130.9414;
    const url = `https://api.open-meteo.com/v1/forecast?latitude=${HITA_LAT}&longitude=${HITA_LON}&current=temperature_2m,weather_code&daily=temperature_2m_max,temperature_2m_min,weather_code&timezone=Asia/Tokyo&forecast_days=2`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return response.getContentText();
  } catch (e) { throw new Error(e.message); }
}