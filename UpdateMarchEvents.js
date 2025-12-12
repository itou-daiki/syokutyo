function updateMarchEventsOneTime() {
    const dailyEvents = [
        { id: Utilities.getUuid(), date: '2026/03/01', time: '', content: '卒業式', note: '', end_date: '2026/03/01', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/02', time: '', content: '振替休業（3/1分）', note: '', end_date: '2026/03/02', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/04', time: '4限', content: '運委（4限）', note: '', end_date: '2026/03/04', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/05', time: '正午', content: '評点入力完了（正午）', note: '', end_date: '2026/03/05', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/06', time: '', content: 'HRA（前期生徒会選挙）、拡大学年会議（15:45-16:35）', note: '', end_date: '2026/03/06', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/09', time: '', content: '入試会場設営', note: '午後カット', end_date: '2026/03/09', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2026/03/10', time: '', content: '一次入試（生徒は自宅学習）、検査日', note: '', end_date: '2026/03/10', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/11', time: '', content: '一次入試（生徒は自宅学習）、採点', note: '', end_date: '2026/03/11', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/12', time: '', content: '一次入試（生徒は自宅学習）、選考会議', note: '', end_date: '2026/03/12', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/13', time: '', content: 'HRA（）、運委（？限）、合格通知発送', note: '', end_date: '2026/03/13', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/16', time: '', content: '【SSC】英語発表会（2年SS/5〜6限）、合格者説明会、（第二志願校出願（〜3/19））', note: '', end_date: '2026/03/16', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/17', time: '', content: '1年県内最先端科学研修（次年度SS予定者）、（第二志願校出願（締切日））、（特例選抜B検査日）', note: '', end_date: '2026/03/17', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/18', time: '', content: '1年県内最先端科学研修（次年度SS予定者）、（入試選考委員会（1限？）・職員会議（午後）、第2志願・特例選抜選考）', note: '午後カット？', end_date: '2026/03/18', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2026/03/20', time: '', content: '春分の日', note: '', end_date: '2026/03/20', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/24', time: '', content: '修了式、大掃除', note: '', end_date: '2026/03/24', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/25', time: '', content: '諸表簿提出（〆切16時）、（転入考査）', note: '', end_date: '2026/03/25', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/26', time: '', content: '運委（10:00-11:00）', note: '', end_date: '2026/03/26', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/27', time: '', content: '職員会議（10:00-11:00, 大会議室）', note: '', end_date: '2026/03/27', type: '' },
        { id: Utilities.getUuid(), date: '2026/03/30', time: '', content: '離任式', note: '', end_date: '2026/03/30', type: '' }
    ];

    const mainEvents = [
        { date: '2026/03/01', content: '卒業式' },
        { date: '2026/03/10', content: '一次入試' },
        { date: '2026/03/11', content: '一次入試' },
        { date: '2026/03/12', content: '一次入試' },
        { date: '2026/03/20', content: '春分の日' },
        { date: '2026/03/24', content: '修了式' },
        { date: '2026/03/30', content: '離任式' }
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

    console.log('March 2026 Events Imported.');
}
