/**
 * データベース雛形作成 & シート初期化
 * - スプレッドシートのカスタムメニューから実行可能
 * - 全シートのヘッダー・列幅・書式・サンプルデータを自動設定
 */

// スプレッドシートを開いたときにメニューを追加
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('朝礼システム')
    .addItem('データベース雛形を作成', 'setupDatabase')
    .addItem('不足シートのみ追加', 'createMissingSheets')
    .addToUi();
}

// 全シート定義
var SHEET_DEFS = [
  {
    name: '職員情報',
    headers: ['ID', '氏名', '職名', '表示順', 'メール', '学年', '分掌1', '分掌2', '分掌3', '教科', '役割', '主任種別'],
    widths:  [120,   90,    70,     60,      200,     50,    80,     80,     80,    60,   80,    80],
    sample:  ['（自動）', '山田太郎', '教諭', '1', 'yamada@school.ed.jp', '1年', '教務部', '', '', '数学', '担任', '']
  },
  {
    name: '行事',
    headers: ['ID', '日付', '時間', '内容', '備考', '終了日', '表示順'],
    widths:  [120,  100,    70,    200,    150,    100,     60],
    dateCol: [2, 6],
    sample:  ['（自動）', '2026-04-06', '放課後', '職員会議', '16:00開始', '2026-04-06', '']
  },
  {
    name: 'イベント',
    headers: ['ID', '日付', 'メインイベント', '校時', '清掃'],
    widths:  [120,  100,    200,            100,   100],
    dateCol: [2],
    sample:  ['（自動）', '2026-04-06', '', '通常校時', '通常清掃']
  },
  {
    name: '出張等',
    headers: ['ID', '日付', '氏名', '用務', '行先', '時間', '備考', '終了日'],
    widths:  [120,  100,    90,    200,    150,    80,    150,    100],
    dateCol: [2, 8],
    sample:  ['（自動）', '2026-04-06', '山田太郎', '教育委員会研修', '県庁', '終日', '', '2026-04-06']
  },
  {
    name: '休暇等',
    headers: ['ID', '日付', '氏名', '種別', '時間', '備考'],
    widths:  [120,  100,    90,    80,    80,    200],
    dateCol: [2],
    sample:  ['（自動）', '2026-04-06', '鈴木花子', '年休', '終日', '']
  },
  {
    name: '会議',
    headers: ['ID', '日付', '会議名', '時間', '場所'],
    widths:  [120,  100,    200,     80,    100],
    dateCol: [2],
    sample:  ['（自動）', '2026-04-06', '学年会', '放課後', '会議室']
  },
  {
    name: '伝達事項',
    headers: ['ID', '日付', '種別', '重要度', '対象', '内容', '報告者', '表示順'],
    widths:  [120,  100,    60,    60,      80,    300,    90,     60],
    dateCol: [2],
    sample:  ['（自動）', '2026-04-06', '職員', '・', '全職員', '明日の時程について', '山田太郎', '']
  },
  {
    name: '特別教室予約',
    headers: ['ID', '日付', '時限', '教室名', '内容', '予約者'],
    widths:  [120,  100,    60,    100,     200,    90],
    dateCol: [2],
    sample:  ['（自動）', '2026-04-06', '5限', 'PC室', '情報の授業', '田中先生']
  },
  {
    name: '定例会議',
    headers: ['曜日', '会議名', '時間', '場所'],
    widths:  [60,    200,     80,    100],
    sample:  ['月', '学年会', '放課後', '会議室']
  },
  {
    name: '特別教室固定',
    headers: ['曜日', '時限', '教室名', '内容', '使用教員'],
    widths:  [60,    60,    100,     200,    90],
    sample:  ['月', '3限', 'PC室', '情報', '田中先生']
  },
  {
    name: 'タスク',
    headers: ['ID', '担当者', 'ロール', '内容', '期限', 'ステータス', '詳細'],
    widths:  [120,  90,      80,     250,    100,   80,       300],
    dateCol: [5],
    sample:  ['（自動）', '山田太郎', '', '年度末書類提出', '2026-04-10', '0', '教務部に提出']
  },
  {
    name: '報告',
    headers: ['ID', '日付', '種別', '報告者', '内容'],
    widths:  [120,  100,    80,    90,      300],
    dateCol: [2],
    sample:  ['（自動）', '2026-04-06', '部活動', '山田太郎', 'サッカー部 練習試合結果報告']
  }
];

/**
 * データベース雛形を作成（全シートを初期化）
 */
function setupDatabase() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    'データベース雛形作成',
    '全シートのヘッダー・書式・サンプルデータを設定します。\n既存データは上書きされません（新規シートのみ作成）。\n\n実行しますか？',
    ui.ButtonSet.OK_CANCEL
  );

  if (result !== ui.Button.OK) return;

  var ss = getSS();
  var created = 0;
  var skipped = 0;

  SHEET_DEFS.forEach(function(def) {
    var sheet = ss.getSheetByName(def.name);
    var isNew = !sheet;

    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      created++;
    } else {
      skipped++;
    }

    // ヘッダーが未設定の場合のみ設定
    var firstCell = sheet.getRange(1, 1).getValue();
    if (!firstCell) {
      // ヘッダー設定
      sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
      sheet.setFrozenRows(1);

      // ヘッダー書式
      var headerRange = sheet.getRange(1, 1, 1, def.headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e8eaf6');
      headerRange.setHorizontalAlignment('center');

      // 列幅設定
      for (var i = 0; i < def.widths.length; i++) {
        sheet.setColumnWidth(i + 1, def.widths[i]);
      }

      // 日付列の書式設定
      if (def.dateCol) {
        def.dateCol.forEach(function(col) {
          sheet.getRange(2, col, sheet.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
        });
      }

      // サンプルデータ（新規シートのみ）
      if (isNew && def.sample) {
        sheet.getRange(2, 1, 1, def.sample.length).setValues([def.sample]);
        // サンプル行を薄いグレーで表示（削除用の目印）
        sheet.getRange(2, 1, 1, def.sample.length).setFontColor('#999999');
      }
    }
  });

  ui.alert(
    '完了',
    '新規作成: ' + created + 'シート\nスキップ（既存）: ' + skipped + 'シート\n\nサンプルデータ（灰色の行）は不要になったら削除してください。',
    ui.ButtonSet.OK
  );
}

/**
 * 不足シートのみ追加（ヘッダー付き）
 */
function createMissingSheets() {
  var ss = getSS();
  var created = 0;

  SHEET_DEFS.forEach(function(def) {
    var sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      sheet = ss.insertSheet(def.name);
      sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
      sheet.setFrozenRows(1);

      var headerRange = sheet.getRange(1, 1, 1, def.headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#e8eaf6');
      headerRange.setHorizontalAlignment('center');

      for (var i = 0; i < def.widths.length; i++) {
        sheet.setColumnWidth(i + 1, def.widths[i]);
      }

      created++;
      console.info('Created sheet: ' + def.name);
    }
  });

  if (created > 0) {
    console.info(created + ' sheets created.');
  } else {
    console.info('All sheets already exist.');
  }
}
