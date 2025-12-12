function updateDecemberEventsOneTime() {
    const dailyEvents = [
        { id: Utilities.getUuid(), date: '2025/12/01', time: '', content: '高教研書道部会公開授業研究（午後）', note: '', end_date: '2025/12/01', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/02', time: '', content: '評点入力完了（２年のみ, 16:40）、第２回入試準備委員会（3限、大会議室）、職員研修（15:45-16:40）', note: '短縮・清掃カット', end_date: '2025/12/02', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2025/12/03', time: '4限', content: '運委', note: '', end_date: '2025/12/03', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/04', time: '', content: '評点入力完了（1,3年, 16:40）、拡大学年会議（２年）', note: '短縮・2年清掃カット', end_date: '2025/12/04', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2025/12/05', time: '', content: 'HRA２年（修学旅行結団式）、第2回学校生活アンケート、いじめ防止・いじめ対策委員会（応接室、15:20-16:30）', note: '', end_date: '2025/12/05', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/06', time: '', content: '修学旅行①（2年）', note: '', end_date: '2025/12/06', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/07', time: '', content: '修学旅行②（2年）', note: '', end_date: '2025/12/07', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/08', time: '', content: '修学旅行③（2年）', note: '', end_date: '2025/12/08', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/09', time: '', content: '修学旅行④（2年）', note: '', end_date: '2025/12/09', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/10', time: '', content: '修学旅行⑤（2年）', note: '', end_date: '2025/12/10', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/11', time: '', content: '1年クラスマッチ、（2年振替休業）、拡大学年会議（1,3年）（15:45-16:35）', note: '短縮・清掃カット', end_date: '2025/12/11', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2025/12/12', time: '', content: 'HRA１年（僕らの未来会議参加、パトリア14:00-15:30）、（2年振替休業）', note: '1年のみ清掃カット', end_date: '2025/12/12', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/13', time: '', content: 'GTEC（1年、自宅受験）', note: '', end_date: '2025/12/13', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/14', time: '', content: 'GTEC（1年、自宅受験）', note: '', end_date: '2025/12/14', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/15', time: '', content: '【SSC】APU留学生との交流会②（2年SS/5〜6限）、運営委員会（７限）', note: '', end_date: '2025/12/15', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/16', time: '', content: '海外研修（台湾）①', note: '', end_date: '2025/12/16', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/17', time: '', content: '海外研修（台湾）②、遠隔配信視察（他県より、5,6限）', note: '', end_date: '2025/12/17', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/18', time: '', content: '海外研修（台湾）③', note: '', end_date: '2025/12/18', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/19', time: '', content: 'HRA（２学期の振り返り）、海外研修（台湾）④', note: '', end_date: '2025/12/19', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/21', time: '', content: '爽風館スクーリング（第２棟立入禁止）', note: '', end_date: '2025/12/21', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/22', time: '', content: '職員会議（15:40-16:35, 大会議室）、編入試験（午前中）', note: '短縮・清掃カット', end_date: '2025/12/22', type: '短縮校時' },
        { id: Utilities.getUuid(), date: '2025/12/23', time: '', content: '２年クラスマッチ', note: '', end_date: '2025/12/23', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/24', time: '', content: '終業式、大掃除、第３回トイレサミット', note: '', end_date: '2025/12/24', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/25', time: '', content: '3年共通テストプレⅡ（代ゼミ）①、（転入考査）', note: '', end_date: '2025/12/25', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/26', time: '', content: '3年共通テストプレⅡ（代ゼミ）②', note: '', end_date: '2025/12/26', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/29', time: '', content: '年末・年始の休業日', note: '', end_date: '2025/12/29', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/30', time: '', content: '年末・年始の休業日', note: '', end_date: '2025/12/30', type: '' },
        { id: Utilities.getUuid(), date: '2025/12/31', time: '', content: '年末・年始の休業日', note: '', end_date: '2025/12/31', type: '' }
    ];

    const mainEvents = [
        { date: '2025/12/06', content: '修学旅行' },
        { date: '2025/12/07', content: '修学旅行' },
        { date: '2025/12/08', content: '修学旅行' },
        { date: '2025/12/09', content: '修学旅行' },
        { date: '2025/12/10', content: '修学旅行' },
        { date: '2025/12/11', content: '1年クラスマッチ' },
        { date: '2025/12/16', content: '海外研修' },
        { date: '2025/12/17', content: '海外研修' },
        { date: '2025/12/18', content: '海外研修' },
        { date: '2025/12/19', content: '海外研修' },
        { date: '2025/12/23', content: '2年クラスマッチ' },
        { date: '2025/12/24', content: '終業式' },
        { date: '2025/12/29', content: '年末・年始の休業日' },
        { date: '2025/12/30', content: '年末・年始の休業日' },
        { date: '2025/12/31', content: '年末・年始の休業日' }
    ];

    const ss = getSS();

    // Update DAILY
    // A:ID, B:Date, C:Time, D:Content, E:Note, F:EndDate, G:Order, H:Type
    const dailySheet = ss.getSheetByName(SHEETS.DAILY);
    const dailyRows = dailyEvents.map(d => [d.id, d.date, d.time, d.content, d.note, d.end_date, 999, d.type]);
    if (dailyRows.length > 0) {
        dailySheet.getRange(dailySheet.getLastRow() + 1, 1, dailyRows.length, 8).setValues(dailyRows);
    }

    // Update EVENT
    const eventSheet = ss.getSheetByName(SHEETS.EVENT);
    // A:ID, B:Date, C:Content
    const eventRows = mainEvents.map(m => [Utilities.getUuid(), m.date, m.content]);
    if (eventRows.length > 0) {
        eventSheet.getRange(eventSheet.getLastRow() + 1, 1, eventRows.length, 3).setValues(eventRows);
    }

    console.log('December Events Imported.');
}
