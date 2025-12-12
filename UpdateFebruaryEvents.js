function updateFebruaryEventsOneTime() {
    const dailyEvents = [
        { id: Utilities.getUuid(), date: '2026/02/03', time: '', content: '推薦入試検査日', note: '', end_date: '2026/02/03', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/04', time: '4限', content: '運委（入試選考会議, 4限）、職員会議（入試号否判定会議, 15:45-16:35, 会議室）', note: '短縮・清掃カット', end_date: '2026/02/04', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2026/02/07', time: '', content: '大学入学共通テスト模試（2年）', note: '', end_date: '2026/02/07', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/10', time: '', content: '第5回入試準備委員会、第7回教務委員会', note: '', end_date: '2026/02/10', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/11', time: '', content: '（建国記念の日）', note: '', end_date: '2026/02/11', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/12', time: '', content: 'SSH成果発表会・STEAM講演会', note: '', end_date: '2026/02/12', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/13', time: '', content: 'HRA（）学校評議員会（13:30-15:00）、高校入試(一次)願書受付（〜2/19）', note: '', end_date: '2026/02/13', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/17', time: '', content: '学年末考査①', note: '', end_date: '2026/02/17', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/18', time: '', content: '学年末考査②', note: '', end_date: '2026/02/18', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/19', time: '', content: '学年末考査③、高校入試(一次)願書受付（締切日）', note: '', end_date: '2026/02/19', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/20', time: '', content: '学年末考査④、サイエンスダイアログ（14:00-15:00）、久大地区高人解研研究大会（13:30〜、会議室）', note: '', end_date: '2026/02/20', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/23', time: '', content: '天皇誕生日', note: '', end_date: '2026/02/23', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/24', time: '', content: '高校入試(一次)志願変更（〜2/27）', note: '', end_date: '2026/02/24', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/25', time: '4限', content: 'クラスマッチ（2年）、運委（4限）', note: '', end_date: '2026/02/25', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/26', time: '', content: 'クラスマッチ（1年）', note: '', end_date: '2026/02/26', type: '' },
        { id: Utilities.getUuid(), date: '2026/02/27', time: '正午', content: '（卒業式設営）、素点入力完了（正午）、職員会議（15:15-16:35, 大会議室）、高校入試(一次)志願変更（締切日）', note: '短縮', end_date: '2026/02/27', type: '短縮校時' }
    ];

    const mainEvents = [
        { date: '2026/02/03', content: '推薦入試' },
        { date: '2026/02/11', content: '建国記念の日' },
        { date: '2026/02/17', content: '学年末考査' },
        { date: '2026/02/18', content: '学年末考査' },
        { date: '2026/02/19', content: '学年末考査' },
        { date: '2026/02/20', content: '学年末考査' },
        { date: '2026/02/23', content: '天皇誕生日' },
        { date: '2026/02/25', content: '2年クラスマッチ' },
        { date: '2026/02/26', content: '1年クラスマッチ' }
    ];

    const ss = getSS();

    // Update DAILY
    const dailySheet = ss.getSheetByName(SHEETS.DAILY);
    const dailyRows = dailyEvents.map(d => [d.id, d.date, d.time, d.content, d.note, d.end_date, 999, d.type]);
    if (dailyRows.length > 0) {
        dailySheet.getRange(dailySheet.getLastRow() + 1, 1, dailyRows.length, 8).setValues(dailyRows);
    }

    // Update EVENT
    const eventSheet = ss.getSheetByName(SHEETS.EVENT);
    const eventRows = mainEvents.map(m => [Utilities.getUuid(), m.date, m.content]);
    if (eventRows.length > 0) {
        eventSheet.getRange(eventSheet.getLastRow() + 1, 1, eventRows.length, 3).setValues(eventRows);
    }

    console.log('February 2026 Events Imported.');
}
