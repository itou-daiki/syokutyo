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
  TASK: 'タスク',
  EVENT: 'イベント' // New
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
function clearDataCache() {
  try {
    CacheService.getScriptCache().removeAll([]);
  } catch (e) {
    console.warn('Cache clear failed', e);
  }
}

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
// =============================================================================
// データ取得API
// =============================================================================
function getData(dateStr) {
  try {
    if (!isValidDate(dateStr)) throw new Error(ERROR_MESSAGES.INVALID_DATE);

    // キャッシュキーを作成（日付 + 時刻で5分単位）
    const cacheKey = `data_${dateStr}_${Math.floor(Date.now() / (5 * 60 * 1000))}`;
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);

    if (cached) {
      return cached;
    }

    // 前後1週間の日付範囲を計算
    const targetDate = new Date(dateStr);
    const startDate = new Date(targetDate);
    startDate.setDate(startDate.getDate() - 7);
    const endDate = new Date(targetDate);
    endDate.setDate(endDate.getDate() + 7);

    // Batch Load All Data (日付範囲でフィルタリング)
    const allData = getAllSheetData(startDate, endDate);

    const todayData = getDataForDate(dateStr, allData);

    // 明日
    const tomorrow = new Date(dateStr);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const tomorrowData = getDataForDate(tomorrowStr, allData);

    // タスクデータ（全件取得し、フロントでフィルタリング）
    let tasks = getTasks(allData.task);
    const staffList = getStaffData(allData.staff);

    // ユーザー識別 & フィルタリング
    const userEmail = Session.getActiveUser().getEmail();
    const currentUser = staffList.find(s => s.email === userEmail);

    if (currentUser) {
      tasks = tasks.filter(t => t.name === currentUser.name);
    }

    const result = JSON.stringify({
      staff: staffList,
      today: todayData,
      tomorrow: tomorrowData,
      tasks: tasks,
      currentUser: currentUser ? currentUser.name : null
    });

    // キャッシュに保存（5分間）
    cache.put(cacheKey, result, 300);

    return result;

  } catch (e) {
    logError('getData', e);
    throw new Error(`データ取得エラー: ${e.message}`);
  }
}

function getAllSheetData(startDate, endDate) {
  try { createMissingSheets(); } catch (e) { console.error('Auto-creation failed', e); }
  const ss = getSS();
  const data = {};

  // 日付フィルタリングが必要なシート
  const dateFilteredSheets = [SHEETS.DAILY, SHEETS.TRIP, SHEETS.LEAVE, SHEETS.MEETING,
                               SHEETS.ANNOUNCE, SHEETS.ROOM, SHEETS.EVENT];

  // 日付フィルタリング不要なシート（固定データ）
  const noFilterSheets = [SHEETS.FIXED_MEETING, SHEETS.FIXED_CLASS, SHEETS.STAFF, SHEETS.TASK];

  // 日付範囲を正規化（時刻を0にして比較）
  const startTime = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()).getTime();
  const endTime = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate()).getTime();

  // 日付フィルタリング対象シートの処理
  dateFilteredSheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) {
      const allRows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
      // 列B（インデックス1）に日付があると仮定してフィルタリング
      data[name] = allRows.filter(row => {
        const dateVal = row[1]; // 列B
        if (!dateVal) return false;
        try {
          const rowDate = new Date(dateVal);
          const rowTime = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate()).getTime();
          return rowTime >= startTime && rowTime <= endTime;
        } catch (e) {
          return false;
        }
      });
    } else {
      data[name] = [];
    }
  });

  // フィルタリング不要なシートの処理
  noFilterSheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) {
      data[name] = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    } else {
      data[name] = [];
    }
  });

  return data;
}

function getDataForDate(dateStr, allData) {
  const result = {
    date: dateStr,
    daily: [], trips: [], leaves: [], meetings: [],
    announcements_staff: [], announcements_student: [], reservations: [],
    counts: { trip: 0, leave: 0, meeting: 0 }
  };

  const targetDate = new Date(dateStr.replace(/-/g, '/'));
  function normalize(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime(); }
  const targetTime = normalize(targetDate);

  const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][targetDate.getDay()];

  // Helper to use memory data
  const getMem = (name) => allData[name] || [];

  // 行事: Start <= Target <= End
  // A:id, B:date, C:time, D:content, E:note, F:end_date, G:display_order, H:schedule_type
  let scheduleType = '';
  result.daily = getMem(SHEETS.DAILY).filter(r => {
    const start = r[1] ? new Date(formatDate(r[1]).replace(/-/g, '/')) : null;
    const end = r[5] ? new Date(formatDate(r[5]).replace(/-/g, '/')) : (start ? new Date(start) : null); // Default end to start

    if (!start) return false;
    const isMatch = normalize(start) <= targetTime && targetTime <= normalize(end);
    return isMatch;
  }).map(r => ({
    id: r[0],
    date: formatDate(r[1]),
    time: r[2] || '',
    content: r[3] || '',
    note: r[4] || '',
    end_date: formatDate(r[5]),
    order: r[6] ? parseInt(r[6]) : 999
    // schedule_type not needed in list item, but we returned it globally
  })).sort((a, b) => (a.order - b.order) || (a.time || '').localeCompare(b.time || ''));

  // Main Event & Schedule Info (EVENT Sheet)
  // A:ID, B:Date, C:Content, D:ScheduleType, E:CleaningStatus
  let mainEventObj = null;
  const events = getMem(SHEETS.EVENT);
  // Find matching date
  const matchedEvent = events.find(r => {
    const d = r[1] ? new Date(formatDate(r[1]).replace(/-/g, '/')) : null;
    return d && normalize(d) === targetTime;
  });

  if (matchedEvent) {
    mainEventObj = {
      id: matchedEvent[0],
      date: formatDate(matchedEvent[1]),
      content: matchedEvent[2] || '',
      scheduleType: matchedEvent[3] || '通常校時',
      cleaningStatus: matchedEvent[4] || '通常清掃'
    };
  } else {
    // Default
    mainEventObj = {
      id: null,
      date: dateStr,
      content: '',
      scheduleType: '通常校時',
      cleaningStatus: '通常清掃'
    };
  }
  result.main_event = mainEventObj;
  // Legacy support removed: result.scheduleType will be undefined or we can polyfill if needed, 
  // but frontend should check result.main_event.scheduleType.
  // For backward compatibility briefly:
  result.scheduleType = mainEventObj.scheduleType;

  // 出張: date == target
  result.trips = getMem(SHEETS.TRIP).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], staff_name: r[2] || '', purpose: r[3] || '', location: r[4] || '', time: r[5] || '', note: r[6] || '' }));
  result.counts.trip = result.trips.length;

  // 休暇: date == target
  result.leaves = getMem(SHEETS.LEAVE).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], staff_name: r[2] || '', type: r[3] || '', time: r[4] || '', note: r[5] || '' }));
  result.counts.leave = result.leaves.length;

  // 会議: date == target OR Fixed Meeting
  const normalMeetings = getMem(SHEETS.MEETING).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], name: r[2] || '', time: r[3] || '', place: r[4] || '', is_fixed: false }));

  const fixedMeetings = getMem(SHEETS.FIXED_MEETING).filter(r => r[0] === dayOfWeek)
    .map(r => ({ id: 'fixed', date: dateStr, name: r[1] || '', time: r[2] || '', place: r[3] || '', is_fixed: true }));

  result.meetings = [...fixedMeetings, ...normalMeetings];
  result.counts.meeting = result.meetings.length;

  // 伝達事項: date == target
  // A:id, B:date, C:type, D:priority, E:target, F:content, G:reporter, H:display_order
  const allAnnounce = getMem(SHEETS.ANNOUNCE).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({
      id: r[0], date: r[1], type: r[2], priority: r[3],
      target: r[4] || '', content: r[5] || '', reporter: r[6] || '',
      order: r[7] ? parseInt(r[7]) : 999
    })).sort((a, b) => (a.order - b.order) || (a.priority === '◎' ? -1 : 1));

  result.announcements_staff = allAnnounce.filter(a => a.type === '職員');
  result.announcements_student = allAnnounce.filter(a => a.type === '生徒');

  // 特別教室
  // Room logic needs care for column count if FixedCols used.
  // getMem returns all cols (from getDataRange equivalent above).
  // ROOM: A:id, B:date, C:period, D:room, E:content, F:reserver
  // Indices: 0,1,2,3,4,5
  const nr = getMem(SHEETS.ROOM).filter(r => formatDate(r[1]) === dateStr)
    .map(r => ({ id: r[0], date: r[1], room: r[3] || '', period: r[2] || '', content: r[4] || '', reserver: r[5] || '', is_fixed: false }));

  // FIXED: A:day, B:period, C:room, D:content, E:reserver
  // Indices: 0,1,2,3,4
  const fr = getMem(SHEETS.FIXED_CLASS).filter(r => r[0] === dayOfWeek)
    .map(r => ({ id: 'fixed', date: dateStr, period: r[1] || '', room: r[2] || '', content: r[3] || '', reserver: r[4] || '', is_fixed: true }));
  result.reservations = [...fr, ...nr];

  return result;
}

function getTasks(taskRows) {
  // タスク: id, name, roll, content, due_date, check, detail
  // accept optional argument
  const rows = taskRows || getRows(SHEETS.TASK);
  return rows.map(r => ({
    id: r[0],
    name: r[1] || '',
    roll: r[2] || '',
    content: r[3] || '',
    due_date: formatDate(r[4]),
    check: r[5] === true || r[5] === 'TRUE',
    detail: r[6] || ''
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
        // id, date, time, content, note, end_date, display_order, schedule_type
        // We preserve schedule_type from existing rows if possible, BUT here we are adding/updating a specific ITEM.
        // It shouldn't overwrite schedule_type column if it's not being set here.
        // However, rowData structure writes a whole row.
        // Strategy: Only update A-G (indices 0-6), leave H (7) alone if we can.
        // OR: backend logic for update handles this?
        // Current logic in `saveData`: rowData = [date, time, content, note, end_date, 999] (len 6)
        // If we write this, it writes to Col B~G. Col H (schedule_type) is untouched. Correct!
        rowData = [data.date, data.time || '', data.content, data.note || '', data.end_date || data.date, 999];
        break;
      case 'trip':
        sheetName = SHEETS.TRIP;
        rowData = [data.date, data.staff, data.purpose, data.location, data.time || '1日', data.note || ''];
        break;
      case 'leave':
        sheetName = SHEETS.LEAVE;
        rowData = [data.date, data.staff, data.type, data.time || '1日', data.note || ''];
        break;
      case 'meeting':
        sheetName = SHEETS.MEETING;
        rowData = [data.date, data.name, data.time || '放課後', data.place || '会議室'];
        break;
      case 'announce':
        sheetName = SHEETS.ANNOUNCE;
        // id, date, type, priority, target, content, reporter, display_order
        rowData = [data.date, data.type, data.priority || '・', data.target || '全職員', data.content, data.reporter, 999];
        break;
      case 'room':
        sheetName = SHEETS.ROOM;
        rowData = [data.date, data.period, data.room, data.content, data.reserver || ''];
        break;
      case 'task':
        sheetName = SHEETS.TASK;
        // id, name, roll, content, due_date, check, detail
        rowData = [data.name, data.roll || '', data.content, data.due_date || '', data.check || false, data.detail || ''];
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
        // 更新処理
        // 行事と伝達事項は列が増えているので注意
        // id列(0)は変更せず、その右側を更新。ただし、updateOrderで順番が変わっている可能性を考慮し、
        // display_orderは既存の値を維持したいが、saveDataには渡されないことが多い。
        // ここでは、データ配列の長さ分だけ更新する。もし既存のorder列よりデータ配列が短ければorderは維持される。
        // しかし、DailyのrowDataは [date, time, content, note, end_date, 999] と定義している。
        // 更新時に999で上書きすると順番がリセットされてしまう。
        // 対策: isUpdate時は既存のorderを読み取るか、order以外の列だけ更新する。

        // 簡易実装: 指定されたrowDataの長さだけ書き込む。ただしDaily/Announceの最終列はorderなので、
        // 渡されたdataにorderが含まれていない場合、backendで既存値を取得するか、clientからorderを送る必要がある。
        // 今回はシンプルに: 「更新時はorder列を書き換えない（rowDataから除外またはSpreadsheetApp操作で工夫）」

        let writeData = [...rowData];
        if ((category === 'daily' || category === 'announce') && isUpdate) {
          // 既存の行データを取得してorderを維持する
          const existingRow = values[rowIndex - 1]; // 0-based
          const orderIdx = category === 'daily' ? 6 : 7; // Daily:G(6), Announce:H(7)
          if (existingRow.length > orderIdx) {
            writeData[orderIdx - 1] = existingRow[orderIdx]; // overwrite the '999' or similar in rowData with existing
          }
        }

        sheet.getRange(rowIndex, 2, 1, writeData.length).setValues([writeData]);
        clearDataCache();
        return { success: true, message: 'データを更新しました', id: id };
      } else {
        // IDが見つからない場合は新規追加扱いにするかエラーにするか。ここでは新規追加に倒す
        const newRow = [id, ...rowData];
        sheet.appendRow(newRow);
        clearDataCache();
        return { success: true, message: 'データが見つからなかったため新規追加しました', id: id };
      }

    } else {
      // 新規追加処理

      // taskの場合、assignees（複数人）への一括登録をハンドル
      if (category === 'task' && data.assignees && Array.isArray(data.assignees)) {
        // assigneesリストの数だけ行を追加
        const rowsToAdd = data.assignees.map(name => {
          const newId = Utilities.getUuid();
          // id, name, roll, content, due_date, check, detail
          return [newId, name, data.roll || '', data.content, data.due_date || '', 0, data.detail || ''];
        });

        // 複数行一括追加 (appendRowは1行ずつなのでループするか、getRange().setValuesを使う)
        // ここではappendRowループで実装（大量ではない前提）
        rowsToAdd.forEach(row => sheet.appendRow(row));

        clearDataCache();
        return { success: true, message: `${rowsToAdd.length}件のタスクを一括登録しました` };
      }

      // 通常の単一行追加
      const newRow = [id, ...rowData];
      sheet.appendRow(newRow);
      clearDataCache();
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
        // 0=未着手, 1=着手中, 2=完了 (数値として保存)
        const newVal = parseInt(isChecked);
        sheet.getRange(i + 1, 6).setValue(isNaN(newVal) ? 0 : newVal);
        clearDataCache();
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
  // ... (No change needed, generic row deletion)
  const sheetName = category === 'daily' ? SHEETS.DAILY :
    category === 'meeting' ? SHEETS.MEETING :
      category === 'trip' ? SHEETS.TRIP :
        category === 'leave' ? SHEETS.LEAVE :
          category === 'announce' ? SHEETS.ANNOUNCE :
            category === 'task' ? SHEETS.TASK : null; // Room deletion logic is separate or covered here? (Room uses saveData usually for cancel)

  if (!sheetName) return { success: false, message: 'Invalid category' };

  const sheet = getSS().getSheetByName(sheetName);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      clearDataCache();
      return { success: true, message: '削除しました' };
    }
  }
  return { success: false, message: 'データが見つかりませんでした' };
}

function moveItem(category, id, direction, dateStr) {
  // direction: -1 (up), 1 (down)
  // dateStr is needed to filter items for the same day (to swap within the day view)
  const sheetName = category === 'daily' ? SHEETS.DAILY : (category === 'announce' ? SHEETS.ANNOUNCE : null);
  if (!sheetName) return;

  const sheet = getSS().getSheetByName(sheetName);
  const rows = sheet.getDataRange().getValues(); // Header row 0
  const headers = rows[0];

  // Identify columns
  const dateColIdx = 1; // B
  const orderColIdx = category === 'daily' ? 6 : 7; // G or H

  // Find all items for this date
  // For Daily, date range affects visibility, but for sorting we usually sort by start_date then order.
  // Actually the UI list contains items that *overlap* the date. Swapping order of items with different start dates is tricky.
  // Assumption: Users want to reorder items displayed *on a specific day*.
  // If we change order(G), it affects that item globally.
  // Let's implement simpler logic: Get all items that exactly match the Start Date (for Daily) or Date (Announce).
  // AND items on that list.

  // To keep it robust: We will get all items displayed on `dateStr` (from getData logic), find the target `id`, 
  // identify the adjacent item *in the current sort order*, and swap their `display_order` values.

  // 1. Get current data for date to determine the sequence
  const currentData = getDataForDate(dateStr);
  let list = [];
  if (category === 'daily') list = currentData.daily;
  else if (category === 'announce') list = [...currentData.announcements_staff, ...currentData.announcements_student]; // Mixed types... sorting might be separate.
  // Wait, announce is split by Staff/Student. Reordering usually happens within the same sub-list.
  // Let's assume we reorder within the same list (e.g. Staff Announce list).
  // Need to know which list the item belongs to. Code doesn't know.
  // But we can just search the sorted list for the ID.

  if (category === 'announce') {
    // Re-merge or find which one contains ID
    const s1 = currentData.announcements_staff.find(i => i.id === id);
    list = s1 ? currentData.announcements_staff : currentData.announcements_student;
  }

  const idx = list.findIndex(i => i.id === id);
  if (idx === -1) return; // Not found

  const swapIdx = idx + direction;
  if (swapIdx < 0 || swapIdx >= list.length) return; // Cannot move

  const targetItem = list[idx];
  const swapItem = list[swapIdx];

  // Update DB
  // We need to find the specific rows for these IDs and swap/update their order values.
  // If they share the same order value, we need to enforce distinction.

  // Simple approach: Assign explicit indices to the whole list to ensure stability, then swap.
  // Or just swap the 'order' values. If order is same/default(999), we prioritize index.

  let targetRowIdx = -1, swapRowIdx = -1;
  let targetOrderVal = 999, swapOrderVal = 999;

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] == targetItem.id) { targetRowIdx = i; targetOrderVal = rows[i][orderColIdx] || 999; }
    if (rows[i][0] == swapItem.id) { swapRowIdx = i; swapOrderVal = rows[i][orderColIdx] || 999; }
  }

  if (targetRowIdx === -1 || swapRowIdx === -1) return;

  // If both are 999 (default), assign distinct values based on current display index
  if (targetOrderVal === swapOrderVal) {
    targetOrderVal = idx;
    swapOrderVal = swapIdx;
  }

  // Swap
  const temp = targetOrderVal;
  targetOrderVal = swapOrderVal;
  swapOrderVal = temp;

  sheet.getRange(targetRowIdx + 1, orderColIdx + 1).setValue(targetOrderVal);
  sheet.getRange(swapRowIdx + 1, orderColIdx + 1).setValue(swapOrderVal);
  clearDataCache();
}

function saveScheduleInfo(dateStr, scheduleType, mainEventContent, cleaningStatus) {
  const sheet = getSS().getSheetByName(SHEETS.EVENT);
  // dateStr is expected to be yyyy-MM-dd
  // sheet date is likely yyyy/MM/dd or Date object

  const rows = sheet.getDataRange().getValues();
  let found = false;

  // Normalize comparisons
  const targetDate = new Date(dateStr.replace(/-/g, '/'));

  for (let i = 1; i < rows.length; i++) {
    const rDate = rows[i][1];
    if (!rDate) continue;
    const d = new Date(rDate);

    // Compare year/month/day
    if (d.getFullYear() === targetDate.getFullYear() &&
      d.getMonth() === targetDate.getMonth() &&
      d.getDate() === targetDate.getDate()) {

      // Update
      // C:Content, D:Type, E:Cleaning
      sheet.getRange(i + 1, 3).setValue(mainEventContent || '');
      sheet.getRange(i + 1, 4).setValue(scheduleType || '通常校時');
      sheet.getRange(i + 1, 5).setValue(cleaningStatus || '通常清掃');
      found = true;
      break;
    }
  }

  if (!found) {
    sheet.appendRow([
      Utilities.getUuid(),
      dateStr,
      mainEventContent || '',
      scheduleType || '通常校時',
      cleaningStatus || '通常清掃'
    ]);
  }
  clearDataCache();
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
function getStaffData(staffRows) {
  try {
    const rows = staffRows || getRows(SHEETS.STAFF);
    // 0:id, 1:name, 2:role, 3:order, 4:email, 5:grade, 6:dept1, 7:dept2, 8:dept3, 9:subject, 10:role_type, 11:chief_type
    return rows.map(r => ({
      id: r[0] || '', name: r[1] || '', role: r[2] || '', order: r[3] || 999, email: r[4] || '',
      grade: r[5] || '', subject: r[9] || '',
      depts: [r[6], r[7], r[8]].filter(d => d && d.toString().trim() !== ''), // Combine Depts
      role_type: r[10] || '',  // 担任, 副担任, 学年所属, 管理職
      chief_type: r[11] || ''  // 学年主任, 分掌主任 etc
    })).filter(s => s.name).sort((a, b) => a.order - b.order);
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