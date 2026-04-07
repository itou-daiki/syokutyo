/**
 * 職員朝礼伝達システム - Backend Logic (v13.0)
 *
 * @description 教員朝礼で使用する伝達事項・予定管理システム
 * @version 13.0 - Performance & Feature Improvements
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
  EVENT: 'イベント',
  REPORT: '報告'
};

const CATEGORY_SHEET_MAP = {
  daily: SHEETS.DAILY,
  trip: SHEETS.TRIP,
  leave: SHEETS.LEAVE,
  meeting: SHEETS.MEETING,
  announce: SHEETS.ANNOUNCE,
  room: SHEETS.ROOM,
  task: SHEETS.TASK,
  report: SHEETS.REPORT
};

const ERROR_MESSAGES = {
  INVALID_DATE: '日付が無効です',
  MISSING_DATA: '必須項目が入力されていません',
  SHEET_NOT_FOUND: 'シートが見つかりません',
  SAVE_FAILED: 'データの保存に失敗しました',
  DELETE_FAILED: 'データの削除に失敗しました',
  UNKNOWN_ERROR: '予期しないエラーが発生しました'
};

// 日付フィルタリング対象シート
const DATE_FILTERED_SHEETS = [
  SHEETS.DAILY, SHEETS.TRIP, SHEETS.LEAVE, SHEETS.MEETING,
  SHEETS.ANNOUNCE, SHEETS.ROOM, SHEETS.EVENT, SHEETS.REPORT
];

// 固定データシート
const STATIC_SHEETS = [SHEETS.FIXED_MEETING, SHEETS.FIXED_CLASS, SHEETS.STAFF, SHEETS.TASK];

// 静的データのキャッシュTTL（30分 - 職員情報・定例会議・固定教室は変更頻度が低い）
const STATIC_CACHE_TTL = 1800;
// 動的データのキャッシュTTL（5分）
const DYNAMIC_CACHE_TTL = 300;

// =============================================================================
// ユーティリティ関数
// =============================================================================

// リクエスト内でSpreadsheet参照をキャッシュ（同一実行内で複数回openByIdを防止）
var _ssCache = null;
function getSS() {
  if (_ssCache) return _ssCache;
  try {
    _ssCache = SpreadsheetApp.openById(SPREADSHEET_ID);
    return _ssCache;
  } catch (e) {
    logError('getSS', e);
    throw new Error('スプレッドシートを開けませんでした');
  }
}

// 排他制御: 書き込み操作をスクリプトロックで保護
function withLock(fn) {
  var lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(15000)) {
      throw new Error('他のユーザーが操作中です。しばらく待ってから再試行してください。');
    }
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function logError(fn, e) {
  console.error('[' + fn + '] Error:', e.message, e.stack);
}



// ターゲット指定キャッシュ無効化（全キャッシュクリアではなく影響範囲のみ）
function clearDataCache(affectedDate) {
  try {
    var cache = CacheService.getScriptCache();
    var tz = Session.getScriptTimeZone();
    var keys = [];

    if (affectedDate) {
      // 影響日とその前後1日のみクリア（複数日イベント対応）
      var af = new Date(affectedDate.replace(/-/g, '/'));
      if (!isNaN(af.getTime())) {
        for (var j = -1; j <= 1; j++) {
          var ad = new Date(af);
          ad.setDate(ad.getDate() + j);
          keys.push('data_' + Utilities.formatDate(ad, tz, 'yyyy-MM-dd'));
        }
      }
    } else {
      // affectedDate未指定時は今日前後3日のみ（旧: ±7日全削除）
      var today = new Date();
      for (var i = -3; i <= 3; i++) {
        var d = new Date(today);
        d.setDate(d.getDate() + i);
        keys.push('data_' + Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
      }
    }
    if (keys.length > 0) cache.removeAll(keys);
    bumpDataVersion(); // 他クライアントに変更を通知
  } catch (e) {
    console.warn('Cache clear failed', e);
  }
}

// 静的データキャッシュをクリア（タスク変更時等）
function clearStaticCache() {
  try {
    CacheService.getScriptCache().remove('static_sheets_v2');
  } catch (e) { /* ignore */ }
}

// 日付フォーマット（最適化版：不要なtry-catch除去、早期リターン）
function formatDate(dateObj) {
  if (!dateObj) return '';
  if (typeof dateObj === 'string') {
    // 既にyyyy-MM-dd形式ならそのまま返す
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateObj)) return dateObj;
    // yyyy/MM/dd形式をyyyy-MM-ddに変換
    if (/^\d{4}\/\d{2}\/\d{2}$/.test(dateObj)) return dateObj.replace(/\//g, '-');
  }
  var date = new Date(dateObj);
  if (isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// 日付を正規化してタイムスタンプ取得（時刻を0にする）
function normalizeDate(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

// =============================================================================
// Web App エントリーポイント
// =============================================================================

function doGet() {
  try {
    return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('職員朝礼伝達システム');
  } catch (e) {
    return HtmlService.createHtmlOutput('システムエラーが発生しました。');
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =============================================================================
// バリデーション関数
// =============================================================================

function isValidDate(d) {
  if (!d) return false;
  var date = new Date(d);
  return date instanceof Date && !isNaN(date.getTime());
}

function validateRequiredFields(data, fields) {
  var missing = fields.filter(function(f) {
    return !data[f] || data[f].toString().trim() === '';
  });
  return { valid: missing.length === 0, missingFields: missing };
}

function isValidCategory(cat) {
  return cat in CATEGORY_SHEET_MAP;
}

// =============================================================================
// データ取得API
// =============================================================================

function getData(dateStr) {
  try {
    if (!isValidDate(dateStr)) throw new Error(ERROR_MESSAGES.INVALID_DATE);

    // シンプルな日付ベースのキャッシュキー
    var cacheKey = 'data_' + dateStr;
    var cache = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);

    if (cached) {
      return cached;
    }

    // 前後7日間の日付範囲を計算
    var targetDate = new Date(dateStr);
    var startDate = new Date(targetDate);
    startDate.setDate(startDate.getDate() - 7);
    var endDate = new Date(targetDate);
    endDate.setDate(endDate.getDate() + 7);

    // 全シートデータをバッチ取得
    var allData = getAllSheetData(startDate, endDate);

    var todayData = buildDateData(dateStr, allData);

    // 明日のデータ
    var tomorrow = new Date(targetDate);
    tomorrow.setDate(tomorrow.getDate() + 1);
    var tomorrowStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var tomorrowData = buildDateData(tomorrowStr, allData);

    // タスクデータ
    var tasks = buildTasks(allData[SHEETS.TASK]);
    var staffList = buildStaffList(allData[SHEETS.STAFF]);

    // ユーザー識別 & フィルタリング
    var userEmail = '';
    try { userEmail = Session.getActiveUser().getEmail(); } catch (e) { /* 権限不足時 */ }
    var currentUser = userEmail ? staffList.find(function(s) { return s.email === userEmail; }) : null;

    // ユーザー特定できた場合は自分のタスクのみ、できない場合はタスク非表示（プライバシー保護）
    if (currentUser) {
      tasks = tasks.filter(function(t) { return t.name === currentUser.name; });
    } else {
      tasks = [];
    }

    var result = JSON.stringify({
      staff: staffList,
      today: todayData,
      tomorrow: tomorrowData,
      tasks: tasks,
      currentUser: currentUser ? currentUser.name : null,
      cacheRange: {
        start: Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        end: Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      }
    });

    // キャッシュに保存（5分間、100KB上限チェック）
    if (result.length < 100000) {
      cache.put(cacheKey, result, 300);
    }

    return result;

  } catch (e) {
    logError('getData', e);
    throw new Error('データ取得エラー: ' + e.message);
  }
}

// 静的データ（職員情報・定例会議・固定教室）を長TTLで分離キャッシュ
// 40人が同時アクセスしてもSpreadsheet APIコールを大幅削減
function getStaticSheetData(ss) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'static_sheets_v2';
  var cached = cache.get(cacheKey);

  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through */ }
  }

  var staticData = {};
  STATIC_SHEETS.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) {
      staticData[name] = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    } else {
      staticData[name] = [];
    }
  });

  // 静的データは30分キャッシュ（100KB上限チェック）
  var json = JSON.stringify(staticData);
  if (json.length < 100000) {
    cache.put(cacheKey, json, STATIC_CACHE_TTL);
  }
  return staticData;
}

function getAllSheetData(startDate, endDate) {
  var ss = getSS();
  var data = {};

  // 日付範囲を正規化
  var startTime = normalizeDate(startDate);
  var endTime = normalizeDate(endDate);

  // 終了日カラムのインデックス（複数日イベント対応）
  var endDateCol = {};
  endDateCol[SHEETS.DAILY] = 5;  // 列F: end_date
  endDateCol[SHEETS.TRIP] = 7;   // 列H: end_date

  // 日付フィルタリング対象シート
  DATE_FILTERED_SHEETS.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (sh && sh.getLastRow() > 1) {
      var allRows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
      var endCol = endDateCol[name];
      data[name] = allRows.filter(function(row) {
        var dateVal = row[1];
        if (!dateVal) return false;
        var rowDate = new Date(dateVal);
        if (isNaN(rowDate.getTime())) return false;
        var rowStart = normalizeDate(rowDate);

        // 複数日イベント: 終了日が範囲内ならinclude
        if (endCol !== undefined && row[endCol]) {
          var endDate = new Date(row[endCol]);
          if (!isNaN(endDate.getTime())) {
            var rowEnd = normalizeDate(endDate);
            return rowStart <= endTime && rowEnd >= startTime;
          }
        }

        return rowStart >= startTime && rowStart <= endTime;
      });
    } else {
      data[name] = [];
    }
  });

  // 静的データは分離キャッシュから取得（30分TTL）
  var staticData = getStaticSheetData(ss);
  STATIC_SHEETS.forEach(function(name) {
    data[name] = staticData[name] || [];
  });

  return data;
}

// =============================================================================
// 日別データ構築（旧getDataForDate）
// =============================================================================

function buildDateData(dateStr, allData) {
  var targetDate = new Date(dateStr.replace(/-/g, '/'));
  var targetTime = normalizeDate(targetDate);
  var dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][targetDate.getDay()];

  var getMem = function(name) { return allData[name] || []; };

  // 行事: Start <= Target <= End
  var daily = getMem(SHEETS.DAILY).filter(function(r) {
    var start = r[1] ? new Date(formatDate(r[1]).replace(/-/g, '/')) : null;
    var end = r[5] ? new Date(formatDate(r[5]).replace(/-/g, '/')) : (start ? new Date(start) : null);
    if (!start) return false;
    return normalizeDate(start) <= targetTime && targetTime <= normalizeDate(end);
  }).map(function(r) {
    return {
      id: r[0], date: formatDate(r[1]), time: r[2] || '', content: r[3] || '',
      note: r[4] || '', end_date: formatDate(r[5]), order: r[6] ? parseInt(r[6]) : 999
    };
  }).sort(function(a, b) {
    return (a.order - b.order) || (a.time || '').localeCompare(b.time || '');
  });

  // イベント情報
  var mainEventObj = buildMainEvent(getMem(SHEETS.EVENT), targetTime, dateStr);

  // 出張: Start <= Target <= End
  var trips = getMem(SHEETS.TRIP).filter(function(r) {
    var start = r[1] ? new Date(formatDate(r[1]).replace(/-/g, '/')) : null;
    var end = r[7] ? new Date(formatDate(r[7]).replace(/-/g, '/')) : (start ? new Date(start) : null);
    if (!start) return false;
    return normalizeDate(start) <= targetTime && targetTime <= normalizeDate(end);
  }).map(function(r) {
    return {
      id: r[0], date: formatDate(r[1]), staff_name: r[2] || '', purpose: r[3] || '',
      location: r[4] || '', time: r[5] || '', note: r[6] || '', end_date: formatDate(r[7])
    };
  });

  // 休暇: date == target
  var leaves = getMem(SHEETS.LEAVE).filter(function(r) {
    return formatDate(r[1]) === dateStr;
  }).map(function(r) {
    return {
      id: r[0], date: formatDate(r[1]), staff_name: r[2] || '', type: r[3] || '',
      time: r[4] || '', note: r[5] || ''
    };
  });

  // 会議: date == target OR Fixed Meeting
  var normalMeetings = getMem(SHEETS.MEETING).filter(function(r) {
    return formatDate(r[1]) === dateStr;
  }).map(function(r) {
    return { id: r[0], date: formatDate(r[1]), name: r[2] || '', time: r[3] || '', place: r[4] || '', is_fixed: false };
  });

  var fixedMeetings = getMem(SHEETS.FIXED_MEETING).filter(function(r) {
    return r[0] === dayOfWeek;
  }).map(function(r) {
    return { id: 'fixed', date: dateStr, name: r[1] || '', time: r[2] || '', place: r[3] || '', is_fixed: true };
  });

  // 伝達事項
  var allAnnounce = getMem(SHEETS.ANNOUNCE).filter(function(r) {
    return formatDate(r[1]) === dateStr;
  }).map(function(r) {
    return {
      id: r[0], date: formatDate(r[1]), type: r[2], priority: r[3],
      target: r[4] || '', content: r[5] || '', reporter: r[6] || '',
      order: r[7] ? parseInt(r[7]) : 999
    };
  }).sort(function(a, b) {
    return (a.order - b.order) || (a.priority === '◎' ? -1 : 1);
  });

  // 報告
  var reports = getMem(SHEETS.REPORT).filter(function(r) {
    return formatDate(r[1]) === dateStr;
  }).map(function(r) {
    return {
      id: r[0], date: formatDate(r[1]), category: r[2] || '', reporter: r[3] || '', content: r[4] || ''
    };
  });

  // 特別教室予約
  var normalRooms = getMem(SHEETS.ROOM).filter(function(r) {
    return formatDate(r[1]) === dateStr;
  }).map(function(r) {
    return {
      id: r[0], date: formatDate(r[1]), room: r[3] || '', period: r[2] || '',
      content: r[4] || '', reserver: r[5] || '', is_fixed: false
    };
  });

  var fixedRooms = getMem(SHEETS.FIXED_CLASS).filter(function(r) {
    return r[0] === dayOfWeek;
  }).map(function(r) {
    return {
      id: 'fixed', date: dateStr, period: r[1] || '', room: r[2] || '',
      content: r[3] || '', reserver: r[4] || '', is_fixed: true
    };
  });

  return {
    date: dateStr,
    daily: daily,
    main_event: mainEventObj,
    scheduleType: mainEventObj.scheduleType,
    trips: trips,
    leaves: leaves,
    meetings: fixedMeetings.concat(normalMeetings),
    announcements_staff: allAnnounce.filter(function(a) { return a.type === '職員'; }),
    announcements_student: allAnnounce.filter(function(a) { return a.type === '生徒'; }),
    reports: reports,
    reservations: fixedRooms.concat(normalRooms),
    counts: {
      trip: trips.length,
      leave: leaves.length,
      meeting: fixedMeetings.length + normalMeetings.length
    }
  };
}

function buildMainEvent(events, targetTime, dateStr) {
  var matched = events.find(function(r) {
    var d = r[1] ? new Date(formatDate(r[1]).replace(/-/g, '/')) : null;
    return d && normalizeDate(d) === targetTime;
  });

  if (matched) {
    return {
      id: matched[0],
      date: formatDate(matched[1]),
      content: matched[2] || '',
      scheduleType: matched[3] || '通常校時',
      cleaningStatus: matched[4] || '通常清掃'
    };
  }

  return {
    id: null, date: dateStr, content: '',
    scheduleType: '通常校時', cleaningStatus: '通常清掃'
  };
}

// =============================================================================
// タスク・職員データ構築
// =============================================================================

function buildTasks(taskRows) {
  var rows = taskRows || getRows(SHEETS.TASK);
  return rows.map(function(r) {
    // check: 0=未着手, 1=着手中, 2=完了 を数値で保持
    var checkVal = typeof r[5] === 'number' ? r[5] : (r[5] === true || r[5] === 'TRUE' ? 2 : 0);
    return {
      id: r[0], name: r[1] || '', roll: r[2] || '', content: r[3] || '',
      due_date: formatDate(r[4]), check: checkVal,
      detail: r[6] || ''
    };
  });
}

function buildStaffList(staffRows) {
  var rows = staffRows || getRows(SHEETS.STAFF);
  return rows.map(function(r) {
    return {
      id: r[0] || '', name: r[1] || '', role: r[2] || '', order: r[3] || 999,
      email: r[4] || '', grade: r[5] || '', subject: r[9] || '',
      depts: [r[6], r[7], r[8]].filter(function(d) { return d && d.toString().trim() !== ''; }),
      role_type: r[10] || '', chief_type: r[11] || ''
    };
  }).filter(function(s) { return s.name; })
    .sort(function(a, b) { return a.order - b.order; });
}

// =============================================================================
// データ保存・更新API
// =============================================================================

// カテゴリ別のrowData構築マッピング
// NOTE: 各builderの戻り値配列は列B以降に対応（列AのIDは含まない）
const ROW_BUILDERS = {
  daily: function(data) {
    return [data.date, data.time || '', data.content, data.note || '', data.end_date || data.date, 999];
  },
  trip: function(data) {
    return [data.date, data.staff, data.purpose, data.location, data.time || '1日', data.note || '', data.end_date || data.date];
  },
  leave: function(data) {
    return [data.date, data.staff, data.type, data.time || '1日', data.note || ''];
  },
  meeting: function(data) {
    return [data.date, data.name, data.time || '放課後', data.place || '会議室'];
  },
  announce: function(data) {
    return [data.date, data.type, data.priority || '・', data.target || '全職員', data.content, data.reporter, 999];
  },
  room: function(data) {
    return [data.date, data.period, data.room, data.content, data.reserver || ''];
  },
  task: function(data) {
    return [data.name, data.roll || '', data.content, data.due_date || '', data.check != null ? data.check : 0, data.detail || ''];
  },
  report: function(data) {
    return [data.date, data.category || '部活動', data.reporter || '', data.content];
  }
};

// display_orderカラムのインデックス（列A=0起算のシート列位置）
// updateRow()では writeData[orderIdx-1] で列B起算の配列インデックスに変換
const ORDER_COL_INDEX = { daily: 6, announce: 7 };

function saveData(category, data) {
  try {
    // --- バリデーション・準備（ロック外で実行 → ロック保持時間を短縮）---
    if (!isValidCategory(category)) throw new Error('無効なカテゴリ: ' + category);
    if (category !== 'task' && !isValidDate(data.date)) throw new Error(ERROR_MESSAGES.INVALID_DATE);

    var isUpdate = !!data.id;
    var id = isUpdate ? data.id : Utilities.getUuid();
    var sheetName = CATEGORY_SHEET_MAP[category];
    var rowData = ROW_BUILDERS[category](data);
    var ss = getSS();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(ERROR_MESSAGES.SHEET_NOT_FOUND + ': ' + sheetName);

    // --- ここからロック内（書込みのみ）---
    return withLock(function() {
      if (isUpdate) {
        return updateRow(sheet, id, rowData, category);
      }

      if (category === 'task' && data.assignees && Array.isArray(data.assignees)) {
        return bulkAddTasks(sheet, data);
      }

      // 教室予約: 同じ日付+教室+時限の既存エントリがあれば更新（競合検知付き）
      if (category === 'room') {
        var existing = sheet.getDataRange().getValues();
        for (var ri = 1; ri < existing.length; ri++) {
          if (formatDate(existing[ri][1]) === data.date &&
              existing[ri][2] === data.period &&
              existing[ri][3] === data.room) {
            // 他ユーザーが既に予約変更していた場合、上書き通知
            var prevContent = existing[ri][4] || '';
            var prevReserver = existing[ri][5] || '';
            sheet.getRange(ri + 1, 2, 1, rowData.length).setValues([rowData]);
            clearDataCache(data.date);
            var msg = (prevContent && prevReserver !== (data.reserver || ''))
              ? '教室予約を更新しました（' + prevReserver + 'の「' + prevContent + '」を上書き）'
              : '教室予約を更新しました';
            return { success: true, message: msg, id: existing[ri][0] };
          }
        }
      }

      sheet.appendRow([id].concat(rowData));
      clearDataCache(data.date);
      return { success: true, message: 'データを保存しました', id: id };
    });

  } catch (e) {
    logError('saveData', e);
    throw new Error(ERROR_MESSAGES.SAVE_FAILED + ': ' + e.message);
  }
}

function updateRow(sheet, id, rowData, category) {
  // TextFinderで高速ID検索
  var finder = sheet.getRange(1, 1, sheet.getLastRow(), 1).createTextFinder(String(id)).matchEntireCell(true);
  var found = finder.findNext();

  if (found) {
    var rowNum = found.getRow();
    var writeData = rowData.slice();

    // display_orderを維持
    var orderIdx = ORDER_COL_INDEX[category];
    if (orderIdx !== undefined) {
      var existingOrder = sheet.getRange(rowNum, orderIdx + 1).getValue();
      if (existingOrder !== '') writeData[orderIdx - 1] = existingOrder;
    }

    sheet.getRange(rowNum, 2, 1, writeData.length).setValues([writeData]);
    clearDataCache();
    return { success: true, message: 'データを更新しました', id: id };
  }

  // IDが見つからない場合は新規追加
  sheet.appendRow([id].concat(rowData));
  clearDataCache();
  return { success: true, message: 'データが見つからなかったため新規追加しました', id: id };
}

function bulkAddTasks(sheet, data) {
  var rows = data.assignees.map(function(name) {
    return [Utilities.getUuid(), name, data.roll || '', data.content, data.due_date || '', 0, data.detail || ''];
  });

  if (rows.length > 0) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  clearDataCache();
  clearStaticCache();
  return { success: true, message: rows.length + '件のタスクを一括登録しました' };
}

// =============================================================================
// タスクステータス切替
// =============================================================================

// タスクステータス切替: ロック不要（単一セル書込み・ユーザー固有操作）
// 40人同時利用時のロック競合を回避
function toggleTaskCheck(id, statusValue) {
  try {
    var sheet = getSS().getSheetByName(SHEETS.TASK);
    if (!sheet) throw new Error(ERROR_MESSAGES.SHEET_NOT_FOUND);

    var finder = sheet.getRange(1, 1, sheet.getLastRow(), 1).createTextFinder(String(id)).matchEntireCell(true);
    var found = finder.findNext();
    if (!found) throw new Error('タスクが見つかりません');

    var newVal = parseInt(statusValue);
    if (isNaN(newVal) || newVal < 0 || newVal > 2) newVal = 0;
    sheet.getRange(found.getRow(), 6).setValue(newVal);
    clearDataCache();
    clearStaticCache(); // タスクは静的キャッシュに含まれる
    return { success: true };
  } catch (e) {
    throw new Error(e.message);
  }
}

// =============================================================================
// データ削除API
// =============================================================================

function deleteEvent(id, category) {
  var sheetName = CATEGORY_SHEET_MAP[category];
  if (!sheetName) return { success: false, message: 'Invalid category' };

  return withLock(function() {
    var sheet = getSS().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: ERROR_MESSAGES.SHEET_NOT_FOUND };

    // TextFinderで高速ID検索
    var finder = sheet.getRange(1, 1, sheet.getLastRow(), 1).createTextFinder(String(id)).matchEntireCell(true);
    var found = finder.findNext();
    if (!found) return { success: false, message: 'データが見つかりませんでした' };

    sheet.deleteRow(found.getRow());
    clearDataCache();
    if (category === 'task') clearStaticCache();
    return { success: true, message: '削除しました' };
  });
}

// =============================================================================
// 順序変更API（最適化版：シート直接操作のみ）
// =============================================================================

function moveItem(category, id, direction, dateStr) {
  var sheetName = CATEGORY_SHEET_MAP[category];
  if (!sheetName || !(category in ORDER_COL_INDEX)) return;

  withLock(function() {
  var orderColIdx = ORDER_COL_INDEX[category];
  var sheet = getSS().getSheetByName(sheetName);
  if (!sheet) return;
  var rows = sheet.getDataRange().getValues();

  // announce: 対象アイテムのtype(職員/生徒)を特定し、同type内でのみ並べ替え
  var announceType = null;
  if (category === 'announce') {
    for (var j = 1; j < rows.length; j++) {
      if (String(rows[j][0]) === String(id)) { announceType = rows[j][2]; break; }
    }
  }

  // 対象日付に一致する行を収集し、order順にソート
  var dateMatchRows = [];
  var targetTime = normalizeDate(new Date(dateStr.replace(/-/g, '/')));

  for (var i = 1; i < rows.length; i++) {
    var isMatch = false;
    if (category === 'daily') {
      var start = rows[i][1] ? new Date(formatDate(rows[i][1]).replace(/-/g, '/')) : null;
      var end = rows[i][5] ? new Date(formatDate(rows[i][5]).replace(/-/g, '/')) : start;
      if (start) {
        isMatch = normalizeDate(start) <= targetTime && targetTime <= normalizeDate(end);
      }
    } else {
      isMatch = (formatDate(rows[i][1]) === dateStr);
    }

    // announce: 同typeのみ
    if (isMatch && announceType !== null && rows[i][2] !== announceType) {
      isMatch = false;
    }

    if (isMatch) {
      dateMatchRows.push({
        sheetRow: i + 1,
        id: rows[i][0],
        order: rows[i][orderColIdx] || 999
      });
    }
  }

  // order順にソート
  dateMatchRows.sort(function(a, b) { return a.order - b.order; });

  // 対象アイテムを見つける
  var idx = dateMatchRows.findIndex(function(r) { return String(r.id) === String(id); });
  if (idx === -1) return;

  var swapIdx = idx + direction;
  if (swapIdx < 0 || swapIdx >= dateMatchRows.length) return;

  var target = dateMatchRows[idx];
  var swap = dateMatchRows[swapIdx];

  // 同じorder値なら明示的に異なる値を付与
  if (target.order === swap.order) {
    target.order = idx;
    swap.order = swapIdx;
  }

  // orderを交換してシートに書き込み
  sheet.getRange(target.sheetRow, orderColIdx + 1).setValue(swap.order);
  sheet.getRange(swap.sheetRow, orderColIdx + 1).setValue(target.order);
  clearDataCache(dateStr);
  }); // withLock
}

// =============================================================================
// スケジュール情報保存（最適化版）
// =============================================================================

function saveScheduleInfo(dateStr, scheduleType, mainEventContent, cleaningStatus) {
  withLock(function() {
  var sheet = getSS().getSheetByName(SHEETS.EVENT);
  if (!sheet) throw new Error(ERROR_MESSAGES.SHEET_NOT_FOUND + ': ' + SHEETS.EVENT);
  var rows = sheet.getDataRange().getValues();

  var targetDateStr = formatDate(new Date(dateStr.replace(/-/g, '/')));

  for (var i = 1; i < rows.length; i++) {
    if (formatDate(rows[i][1]) === targetDateStr) {
      // 一括更新（3セル分をまとめて書き込み）
      sheet.getRange(i + 1, 3, 1, 3).setValues([[
        mainEventContent || '',
        scheduleType || '通常校時',
        cleaningStatus || '通常清掃'
      ]]);
      clearDataCache(dateStr);
      return;
    }
  }

  // 新規追加
  sheet.appendRow([
    Utilities.getUuid(), dateStr, mainEventContent || '',
    scheduleType || '通常校時', cleaningStatus || '通常清掃'
  ]);
  clearDataCache(dateStr);
  }); // withLock
}

// =============================================================================
// データバージョン管理（同時編集検知）
// =============================================================================

// 全書込み操作で呼ばれ、バージョンをインクリメント
// クライアントはポーリングでバージョン変化を検知→自動リロード
function bumpDataVersion() {
  try {
    var cache = CacheService.getScriptCache();
    var current = parseInt(cache.get('data_version') || '0');
    cache.put('data_version', String(current + 1), 21600); // 6時間TTL
  } catch (e) { /* ignore */ }
}

// クライアントが定期的に呼ぶ軽量API（シートアクセスなし）
function getDataVersion() {
  try {
    return parseInt(CacheService.getScriptCache().get('data_version') || '0');
  } catch (e) { return 0; }
}

// =============================================================================
// ヘルパー関数
// =============================================================================

function getRows(name) {
  try {
    var sheet = getSS().getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return [];
    return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  } catch (e) { return []; }
}

// =============================================================================
// 天気API
// =============================================================================

function getWeatherData() {
  try {
    var HITA_LAT = 33.3219;
    var HITA_LON = 130.9414;
    var url = 'https://api.open-meteo.com/v1/forecast?latitude=' + HITA_LAT +
      '&longitude=' + HITA_LON +
      '&current=temperature_2m,weather_code&daily=temperature_2m_max,temperature_2m_min,weather_code&timezone=Asia/Tokyo&forecast_days=2';
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    return response.getContentText();
  } catch (e) {
    throw new Error(e.message);
  }
}
