function createMissingSheets() {
    const ss = getSS();

    // Define required sheets and their headers
    const REQUIRED_SHEETS = [
        { name: '特別教室予約', headers: ['ID', 'Date', 'Period', 'Room', 'Content', 'Reserver'] }
    ];

    REQUIRED_SHEETS.forEach(info => {
        let sheet = ss.getSheetByName(info.name);
        if (!sheet) {
            console.log(`Creating missing sheet: ${info.name}`);
            sheet = ss.insertSheet(info.name);
            if (info.headers) {
                sheet.getRange(1, 1, 1, info.headers.length).setValues([info.headers]);
                sheet.setFrozenRows(1);
            }
        } else {
            console.log(`Sheet already exists: ${info.name}`);
        }
    });
}
