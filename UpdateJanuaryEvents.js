function updateJanuaryEventsOneTime() {
    const dailyEvents = [
        { id: Utilities.getUuid(), date: '2026/01/01', time: '', content: '元旦', note: '', end_date: '2026/01/01', type: '' },
        { id: Utilities.getUuid(), date: '2026/01/02', time: '', content: '年末・年始の休業日', note: '', end_date: '2026/01/02', type: '' },
        { id: Utilities.getUuid(), date: '2026/01/07', time: '', content: '第3回入試準備委員会（9:00-9:50）、運営委員会（10:30〜11:20）', note: '', end_date: '2026/01/07', type: '' },
        { id: Utilities.getUuid(), date: '2026/01/08', time: '', content: '始業式、大掃除、職員会議（13:30〜15:00）', note: '', end_date: '2026/01/08', type: '' },
        { id: Utilities.getUuid(), date: '2026/01/09', time: '', content: '課題・実力考査、共通テスト結団式', note: '', end_date: '2026/01/09', type: '' },
        { id: Utilities.getUuid(), date: '2026/01/12', time: '', content: '成人の日', note: '', end_date: '2026/01/12', type: '' },
        { id: Utilities.getUuid(), date: '2026/01/13', time: '', content: '課題・実力考査、第4回入試準備委員会（7限, 大会議室）', note: '', end_date: '2026/01/13', type: '' },
        { id: Utilities.getUuid(), date: '2026/01/14', time: '', content: '面接旬間（〜1/23）', note: '短縮', end_date: '2026/01/14', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2026/01/15', time: '', content: '', note: '短縮', end_date: '2026/01/15', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2026/01/16', time: '', content: '人権学習HRA（１年）', note: '短縮', end_date: '2026/01/16', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2026/01/17', time: '', content: '共通テスト、ベネッセ総合学力テスト（1,2年）', note: '', end_date: '2026/01/17', type: '' },
        { id: Utilities.getUuid(), date: '2026/01/18', time: '', content: '共通テスト、ベネッセ総合学力テスト（2年）', note: '', end_date: '2026/01/18', type: '' },
        { id: Utilities.getUuid(), date: '2026/01/19', time: '', content: '共通テスト自己採点、第8回推薦委員会', note: '短縮', end_date: '2026/01/19', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2026/01/20', time: '', content: '', note: '短縮', end_date: '2026/01/20', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2026/01/21', time: '4限', content: '運委', note: '短縮', end_date: '2026/01/21', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2026/01/22', time: '', content: 'PTA役員会（15:00-17:00）', note: '短縮', end_date: '2026/01/22', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2026/01/23', time: '', content: '主権者教育HRA（1,2年）、志望校検討会議第3回', note: '短縮', end_date: '2026/01/23', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2026/01/27', time: '', content: '職員会議（15:45-16:35, 大会議室）', note: '短縮・清掃カット', end_date: '2026/01/27', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2026/01/28', time: '', content: '【WS】学年発表会（1年全/5〜6限）', note: '', end_date: '2026/01/28', type: '' },
        { id: Utilities.getUuid(), date: '2026/01/30', time: '', content: '公開人権HRA（2年）、人権職員研修（対話会, 全職員）', note: '', end_date: '2026/01/30', type: '' }
    ];

    const mainEvents = [
        { date: '2026/01/01', content: '元旦' },
        { date: '2026/01/02', content: '年末・年始の休業日' },
        { date: '2026/01/08', content: '始業式' },
        { date: '2026/01/12', content: '成人の日' },
        { date: '2026/01/17', content: '共通テスト' },
        { date: '2026/01/18', content: '共通テスト' }
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

    console.log('January 2026 Events Imported.');
}
