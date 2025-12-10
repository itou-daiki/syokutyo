/**
 * 職員朝礼伝達システム - Backend Logic (v7.0)
 *
 * @description 教員朝礼で使用する伝達事項・予定管理システム
 * @version 7.0
 * @author System
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
  FIXED_CLASS: '特別教室固定'
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
// ユーティリティ関数
// =============================================================================

/**
 * スプレッドシートインスタンスを取得
 * @returns {Spreadsheet} スプレッドシートオブジェクト
 * @throws {Error} スプレッドシートが開けない場合
 */
function getSS() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (e) {
    logError('getSS', e);
    throw new Error('スプレッドシートを開けませんでした');
  }
}

/**
 * エラーログを記録
 * @param {string} functionName - 関数名
 * @param {Error} error - エラーオブジェクト
 */
function logError(functionName, error) {
  console.error(`[${functionName}] Error:`, error.message);
  console.error(`Stack trace:`, error.stack);
}

/**
 * 情報ログを記録
 * @param {string} functionName - 関数名
 * @param {string} message - メッセージ
 */
function logInfo(functionName, message) {
  console.info(`[${functionName}] ${message}`);
}

// =============================================================================
// Web App エントリーポイント
// =============================================================================

/**
 * Webアプリケーションのエントリーポイント
 * @returns {HtmlOutput} HTMLページ
 */
function doGet() {
  try {
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('職員朝礼伝達システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (e) {
    logError('doGet', e);
    return HtmlService.createHtmlOutput('システムエラーが発生しました。管理者に連絡してください。');
  }
}

/**
 * HTMLファイルを読み込んでコンテンツを返す
 * @param {string} filename - ファイル名
 * @returns {string} ファイルのコンテンツ
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =============================================================================
// バリデーション関数
// =============================================================================

/**
 * 日付文字列が有効かチェック
 * @param {string} dateStr - 日付文字列（yyyy-MM-dd形式）
 * @returns {boolean} 有効な場合true
 */
function isValidDate(dateStr) {
  if (!dateStr) return false;
  const date = new Date(dateStr);
  return date instanceof Date && !isNaN(date.getTime());
}

/**
 * 必須フィールドをチェック
 * @param {Object} data - チェックするデータ
 * @param {Array<string>} requiredFields - 必須フィールド名の配列
 * @returns {Object} {valid: boolean, missingFields: Array<string>}
 */
function validateRequiredFields(data, requiredFields) {
  const missingFields = requiredFields.filter(field => !data[field] || data[field].trim() === '');
  return {
    valid: missingFields.length === 0,
    missingFields: missingFields
  };
}

/**
 * データカテゴリが有効かチェック
 * @param {string} category - カテゴリ名
 * @returns {boolean} 有効な場合true
 */
function isValidCategory(category) {
  const validCategories = ['daily', 'trip', 'leave', 'meeting', 'announce', 'room'];
  return validCategories.includes(category);
}

// =============================================================================
// データ取得API
// =============================================================================

/**
 * 指定日のデータを取得
 * @param {string} dateStr - 日付文字列（yyyy-MM-dd形式）
 * @returns {string} JSON形式のデータ
 * @throws {Error} 無効な日付の場合
 */
function getData(dateStr) {
  try {
    logInfo('getData', `Fetching data for date: ${dateStr}`);

    // 日付のバリデーション
    if (!isValidDate(dateStr)) {
      throw new Error(ERROR_MESSAGES.INVALID_DATE);
    }

    // 本日のデータを取得
    const todayData = getDataForDate(dateStr);

    // 明日の日付を計算
    const tomorrow = new Date(dateStr);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = formatDate(tomorrow);

    // 明日のデータを取得
    const tomorrowData = getDataForDate(tomorrowStr);

    const result = {
      staff: getStaffData(),
      today: todayData,
      tomorrow: tomorrowData
    };

    logInfo('getData', `Data fetched successfully. Today events: ${todayData.daily.length}, Tomorrow events: ${tomorrowData.daily.length}`);
    return JSON.stringify(result);

  } catch (e) {
    logError('getData', e);
    throw new Error(`データ取得エラー: ${e.message}`);
  }
}

/**
 * 指定日のデータを取得（内部関数）
 * @param {string} dateStr - 日付文字列（yyyy-MM-dd形式）
 * @returns {Object} その日のデータ
 */
function getDataForDate(dateStr) {
  const result = {
    date: dateStr,
    daily: [],
    trips: [],
    leaves: [],
    meetings: [],
    announcements_staff: [],
    announcements_student: [],
    reservations: [],
    counts: { trip: 0, leave: 0, meeting: 0 }
  };

  const targetDate = new Date(dateStr);
  const dayMap = ['日', '月', '火', '水', '木', '金', '土'];
  const dayOfWeek = dayMap[targetDate.getDay()];

    // 1. 行事 (複数登録対応)
    result.daily = getRows(SHEETS.DAILY)
      .filter(r => formatDate(r[1]) === dateStr)
      .map(r => ({
        id: r[0],
        date: r[1],
        time: r[2] || '',
        content: r[3] || '',
        note: r[4] || ''
      }));

    // 2. 出張
    result.trips = getRows(SHEETS.TRIP)
      .filter(r => formatDate(r[1]) === dateStr)
      .map(r => ({
        id: r[0],
        date: r[1],
        staff_name: r[2] || '',
        purpose: r[3] || '',
        location: r[4] || '',
        time: r[5] || '',
        note: r[6] || ''
      }));
    result.counts.trip = result.trips.length;

    // 3. 休暇
    result.leaves = getRows(SHEETS.LEAVE)
      .filter(r => formatDate(r[1]) === dateStr)
      .map(r => ({
        id: r[0],
        date: r[1],
        staff_name: r[2] || '',
        type: r[3] || '',
        time: r[4] || '',
        note: r[5] || ''
      }));
    result.counts.leave = result.leaves.length;

    // 4. 会議 (定例 + 個別)
    const nm = getRows(SHEETS.MEETING)
      .filter(r => formatDate(r[1]) === dateStr)
      .map(r => ({
        id: r[0],
        date: r[1],
        name: r[2] || '',
        time: r[3] || '',
        place: r[4] || '',
        is_fixed: false
      }));
    const fm = getRows(SHEETS.FIXED_MEETING)
      .filter(r => r[0] === dayOfWeek)
      .map(r => ({
        id: 'fixed',
        date: dateStr,
        name: r[1] || '',
        time: r[2] || '',
        place: r[3] || '',
        is_fixed: true
      }));
    result.meetings = [...fm, ...nm];
    result.counts.meeting = result.meetings.length;

    // 5. 伝達事項
    const aa = getRows(SHEETS.ANNOUNCE)
      .filter(r => formatDate(r[1]) === dateStr)
      .map(r => ({
        id: r[0],
        date: r[1],
        type: r[2] || '',
        priority: r[3] || '',
        target: r[4] || '',
        content: r[5] || '',
        reporter: r[6] || ''
      }));
    result.announcements_staff = aa.filter(a => a.type === '職員');
    result.announcements_student = aa.filter(a => a.type === '生徒');

    // 6. 教室予約
    const nr = getRowsFixedCols(SHEETS.ROOM, 6)
      .filter(r => formatDate(r[1]) === dateStr)
      .map(r => ({
        id: r[0],
        date: r[1],
        room: r[2] || '',
        period: r[3] || '',
        content: r[4] || '',
        reserver: r[5] || '',
        is_fixed: false
      }));
    const fr = getRowsFixedCols(SHEETS.FIXED_CLASS, 5)
      .filter(r => r[0] === dayOfWeek)
      .map(r => ({
        id: 'fixed',
        date: dateStr,
        period: r[1] || '',
        room: r[2] || '',
        content: r[3] || '',
        reserver: r[4] || '',
        is_fixed: true
      }));
    result.reservations = [...fr, ...nr];

    return result;
}

// =============================================================================
// データ保存API
// =============================================================================

/**
 * データを保存
 * @param {string} category - カテゴリ名
 * @param {Object} data - 保存するデータ
 * @returns {Object} {success: boolean, message: string, id: string}
 * @throws {Error} 保存失敗時
 */
function saveData(category, data) {
  try {
    logInfo('saveData', `Saving data for category: ${category}`);

    // カテゴリのバリデーション
    if (!isValidCategory(category)) {
      throw new Error(`無効なカテゴリです: ${category}`);
    }

    // 日付のバリデーション
    if (!isValidDate(data.date)) {
      throw new Error(ERROR_MESSAGES.INVALID_DATE);
    }

    const id = Utilities.getUuid();
    let sheetName = "";
    let rowData = [];
    let requiredFields = [];

    switch (category) {
      case 'daily':
        sheetName = SHEETS.DAILY;
        requiredFields = ['date', 'content'];
        rowData = [id, data.date, data.time || '', data.content, data.note || ''];
        break;

      case 'trip':
        sheetName = SHEETS.TRIP;
        requiredFields = ['date', 'staff', 'purpose', 'location'];
        rowData = [id, data.date, data.staff, data.purpose, data.location, data.time || '1日', data.note || ''];
        break;

      case 'leave':
        sheetName = SHEETS.LEAVE;
        requiredFields = ['date', 'staff', 'type'];
        rowData = [id, data.date, data.staff, data.type, data.time || '1日', data.note || ''];
        break;

      case 'meeting':
        sheetName = SHEETS.MEETING;
        requiredFields = ['date', 'name'];
        rowData = [id, data.date, data.name, data.time || '放課後', data.place || '会議室'];
        break;

      case 'announce':
        sheetName = SHEETS.ANNOUNCE;
        requiredFields = ['date', 'type', 'content', 'reporter'];
        rowData = [id, data.date, data.type, data.priority || '・', data.target || '全職員', data.content, data.reporter];
        break;

      case 'room':
        sheetName = SHEETS.ROOM;
        requiredFields = ['date', 'room', 'period', 'content'];
        rowData = [id, data.date, data.room, data.period, data.content, data.reserver || ''];
        break;

      default:
        throw new Error(`未対応のカテゴリです: ${category}`);
    }

    // 必須フィールドのバリデーション
    const validation = validateRequiredFields(data, requiredFields);
    if (!validation.valid) {
      throw new Error(`${ERROR_MESSAGES.MISSING_DATA}: ${validation.missingFields.join(', ')}`);
    }

    // シートへの書き込み
    const sheet = getSS().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`${ERROR_MESSAGES.SHEET_NOT_FOUND}: ${sheetName}`);
    }

    sheet.appendRow(rowData);
    logInfo('saveData', `Data saved successfully with ID: ${id}`);

    return {
      success: true,
      message: 'データを保存しました',
      id: id
    };

  } catch (e) {
    logError('saveData', e);
    throw new Error(`${ERROR_MESSAGES.SAVE_FAILED}: ${e.message}`);
  }
}

// =============================================================================
// データ削除API
// =============================================================================

/**
 * データを削除
 * @param {string} id - データID
 * @param {string} category - カテゴリ名
 * @returns {Object} {success: boolean, message: string}
 * @throws {Error} 削除失敗時
 */
function deleteEvent(id, category) {
  try {
    logInfo('deleteEvent', `Deleting data with ID: ${id}, category: ${category}`);

    // IDとカテゴリのバリデーション
    if (!id || !category) {
      throw new Error('IDまたはカテゴリが指定されていません');
    }

    // カテゴリからシート名を取得
    const categoryToSheet = {
      'room': SHEETS.ROOM,
      'daily': SHEETS.DAILY,
      'announce': SHEETS.ANNOUNCE,
      'trip': SHEETS.TRIP,
      'leave': SHEETS.LEAVE,
      'meeting': SHEETS.MEETING
    };

    const sheetName = categoryToSheet[category];
    if (!sheetName) {
      throw new Error(`未対応のカテゴリです: ${category}`);
    }

    const sheet = getSS().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error(`${ERROR_MESSAGES.SHEET_NOT_FOUND}: ${sheetName}`);
    }

    const data = sheet.getDataRange().getValues();
    let rowDeleted = false;

    // IDが一致する行を探して削除
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        rowDeleted = true;
        logInfo('deleteEvent', `Data deleted successfully at row: ${i + 1}`);
        break;
      }
    }

    if (!rowDeleted) {
      throw new Error(`指定されたIDのデータが見つかりません: ${id}`);
    }

    return {
      success: true,
      message: 'データを削除しました'
    };

  } catch (e) {
    logError('deleteEvent', e);
    throw new Error(`${ERROR_MESSAGES.DELETE_FAILED}: ${e.message}`);
  }
}

// =============================================================================
// ヘルパー関数
// =============================================================================

/**
 * シートから全行を取得
 * @param {string} name - シート名
 * @returns {Array<Array>} データ行の配列（ヘッダー行を除く）
 */
function getRows(name) {
  try {
    const sheet = getSS().getSheetByName(name);
    if (!sheet) {
      logError('getRows', new Error(`シートが見つかりません: ${name}`));
      return [];
    }
    if (sheet.getLastRow() < 2) return [];

    return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  } catch (e) {
    logError('getRows', e);
    return [];
  }
}

/**
 * シートから指定列数の全行を取得
 * @param {string} sheetName - シート名
 * @param {number} colCount - 取得する列数
 * @returns {Array<Array>} データ行の配列（ヘッダー行を除く）
 */
function getRowsFixedCols(sheetName, colCount) {
  try {
    const sheet = getSS().getSheetByName(sheetName);
    if (!sheet) {
      logError('getRowsFixedCols', new Error(`シートが見つかりません: ${sheetName}`));
      return [];
    }
    if (sheet.getLastRow() < 2) return [];

    return sheet.getRange(2, 1, sheet.getLastRow() - 1, colCount).getValues();
  } catch (e) {
    logError('getRowsFixedCols', e);
    return [];
  }
}

/**
 * 職員データを取得
 * @returns {Array<Object>} 職員データの配列（並び順でソート済み）
 */
function getStaffData() {
  try {
    const sheet = getSS().getSheetByName(SHEETS.STAFF);
    if (!sheet) {
      logError('getStaffData', new Error(`職員情報シートが見つかりません`));
      return [];
    }

    const rows = getRows(SHEETS.STAFF);
    return rows
      .map(r => ({
        id: r[0] || '',
        name: r[1] || '',
        role: r[2] || '',
        order: r[3] || 999
      }))
      .filter(staff => staff.name) // 名前が空でないものだけ
      .sort((a, b) => a.order - b.order);
  } catch (e) {
    logError('getStaffData', e);
    return [];
  }
}

/**
 * 日付オブジェクトを文字列に変換
 * @param {Date|string} dateObj - 日付オブジェクトまたは文字列
 * @returns {string} yyyy-MM-dd形式の日付文字列
 */
function formatDate(dateObj) {
  try {
    if (!dateObj) return "";
    const date = new Date(dateObj);
    if (isNaN(date.getTime())) return "";
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (e) {
    logError('formatDate', e);
    return "";
  }
}