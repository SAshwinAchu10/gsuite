function scheduleShifts() {
    var sheet = SpreadsheetApp.getActiveSheet();
    var calId = sheet.getRange("B1").getValue();

    var headerRows = 3;
    var range = sheet.getDataRange();
    var data = range.getValues();
    var changes = 0;
    for (i in data) {
        if (i < headerRows) continue;
        var row = data[i];
        var interviewer = row[0];
        var candidate = row[1];
        var title = row[2];
        var description = row[3];
        if (!interviewer) continue;
        var startTime = new Date(row[4]);
        var endTime = new Date(row[5]);
        var eventIdInSheet = row[6];
        if (interviewer != null && startTime != null && endTime != null && description != null && candidate != null && title != null) {
            var event = null;
            try {
                if (eventIdInSheet) {
                    event = CalendarApp.getCalendarById(calId).getEventById(eventIdInSheet);
                }
            } catch (e) {
            }
            var insert = 0;
            if (event != null && event.getGuestByEmail(interviewer) != null && event.getGuestByEmail(candidate) != null &&
                String(event.getStartTime()) == String(startTime) &&
                String(event.getEndTime()) == String(endTime)) {
            } else if (event != null) {
                var existingEvent = CalendarApp.getCalendarById(calId).getEventById(eventIdInSheet)
                try {
                    existingEvent.deleteEvent();
                } catch (e) {
                }
                insert = 1;
            } else {
                insert = 1;
            }
            if (insert) {
                var newEvent = CalendarApp.getCalendarById(calId).createEvent(title != null ? title : 'Interview', startTime, endTime,
                    { description: description, sendInvites: true, guests: `${interviewer},${candidate}` });
                row[6] = newEvent.getId();
                changes++;
            }
        }

    }
    if (changes) {
        range.setValues(data);
    }
    return changes;
}
