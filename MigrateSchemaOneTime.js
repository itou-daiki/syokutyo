function migrateSchemaOneTime() {
    const ss = getSS();
    const dailySheet = ss.getSheetByName(SHEETS.DAILY);
    const eventSheet = ss.getSheetByName(SHEETS.EVENT);

    if (!dailySheet || !eventSheet) throw new Error('Sheets not found');

    // 1. Setup EVENT Headers
    // A:ID, B:Date, C:Content, D:ScheduleType, E:CleaningStatus
    // Check header
    const headerRange = eventSheet.getRange(1, 1, 1, 5);
    const headers = headerRange.getValues()[0];
    if (headers[3] !== 'ScheduleType') eventSheet.getRange('D1').setValue('ScheduleType');
    if (headers[4] !== 'CleaningStatus') eventSheet.getRange('E1').setValue('CleaningStatus');

    // 2. Read Map from DAILY
    const dailyData = dailySheet.getDataRange().getValues(); // Row 1 is header
    const dateScheduleMap = {};

    // Start from row 2
    for (let i = 1; i < dailyData.length; i++) {
        const row = dailyData[i];
        const date = row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy/MM/dd') : row[1];
        const type = row[7]; // H

        if (date && type) {
            dateScheduleMap[date] = type;
            // Clear from DAILY (optional, but requested)
            // dailySheet.getRange(i+1, 8).clearContent(); 
            // Actually writing back row by row is slow. Let's just create a batch update event later if needed.
            // For safety, let's just leave legacy data for now or clear it in batch? 
            // Plan said: "Remove schedule_type usage". So clearing is good.
        }
    }

    // Clear DAILY Col H (Batch)
    if (dailyData.length > 1) {
        dailySheet.getRange(2, 8, dailyData.length - 1, 1).clearContent();
    }

    // 3. Update EVENT Sheet
    const eventLastRow = eventSheet.getLastRow();
    let eventData = [];
    if (eventLastRow > 1) {
        eventData = eventSheet.getRange(2, 1, eventLastRow - 1, 5).getValues();
    }

    const existingDates = new Set();

    // Update existing
    const updatedEventData = eventData.map(row => {
        const date = row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy/MM/dd') : row[1];
        existingDates.add(date);

        // Set ScheduleType defaults
        if (!row[3]) {
            row[3] = dateScheduleMap[date] || '通常校時';
        }
        // Set CleaningStatus defaults
        if (!row[4]) {
            row[4] = '通常清掃';
        }
        return row;
    });

    // Write back existing updates
    if (updatedEventData.length > 0) {
        eventSheet.getRange(2, 1, updatedEventData.length, 5).setValues(updatedEventData);
    }

    // 4. Create new rows for missing dates
    const newRows = [];
    Object.keys(dateScheduleMap).forEach(date => {
        if (!existingDates.has(date)) {
            newRows.push([
                Utilities.getUuid(),
                date,
                '', // Content
                dateScheduleMap[date],
                '通常清掃'
            ]);
        }
    });

    if (newRows.length > 0) {
        eventSheet.getRange(eventSheet.getLastRow() + 1, 1, newRows.length, 5).setValues(newRows);
    }

    console.log('Schema Migration Complete.');
}
