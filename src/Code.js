/* 
 *
 * Syncs events in "Year at a Glance" spreadsheet with a specified Google Calendar.
 * 
 * NOTE: CREATE A NEW GOOGLE CALENDAR FOR THIS - YOUR EXISTING EVENTS NOT IN SPREADSHEET WILL BE DELETED!
 */
function syncEventsWithGoogleCalendar() {
  let CALENDAR_ID = "aa933b9e7ca36cdf14e1f86ae12de5ca8e7c6d707947e5d3e6d5f51e989a9240@group.calendar.google.com";
  let calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    if (sheet.getName() === "README") {
      continue;
    }
    let sheetStartDate = sheet.getRange("A2").getValue();
    let startDate = new Date(sheetStartDate.getFullYear(), 0, 1)
    let endDate = new Date(sheetStartDate.getFullYear(), 11, 31);
    deleteEvents(calendar, startDate, endDate);
    createEventsFromSheet(calendar, spreadsheet, sheet);
  }
}

/* 
 * Delete all events between two specified dates (inclusive).
 * 
 * @param calendarId The ID of the calendar containing the event to be deleted.
 * @param startDate The start date from which events should be deleted.
 * @param endDate The end date from which events should be deleted.
 */
function deleteEvents(calendar, startDate, endDate) {
  let events = calendar.getEvents(startDate, endDate);
  for (let i = 0; i < events.length; i++) {
    events[i].deleteEvent();
  }
}

/*
 * Create events from sheet.
 * 
 * @param calendarId The ID of the calendar containing the event to be deleted.
 * @param sheet The sheet to create events from.
 */
function createEventsFromSheet(calendar, spreadsheet, sheet) {
  let eventDataColumns = [
    [1, 2],
    [3, 4],
    [5, 6],
    [7, 8],
    [9, 10],
    [11, 12],
    [13, 14],
    [15, 16],
    [17, 18],
    [19, 20],
    [21, 22],
    [23, 24],
  ]; // List of date/event name pairs in the form of [Date, Event Title]

  for (let i=0; i<eventDataColumns.length; i++) {
    let dateColumnValues = sheet.getRange(2, eventDataColumns[i][0], sheet.getLastRow() - 1).getValues();
    let eventNameColumn = sheet.getRange(2, eventDataColumns[i][1], sheet.getLastRow() - 1);
    let eventNameColumnRichTextValues = eventNameColumn.getRichTextValues();

    for (let j=0; j<dateColumnValues.length; j++) {
      let date = dateColumnValues[j][0];
      let eventCellRichTextValue = eventNameColumnRichTextValues[j][0];
      let eventCellText = eventCellRichTextValue.getText();
      if (date && eventCellText) {
        let eventTitleList = eventCellText.split("|");
        for (let k=0; k<eventTitleList.length; k++) {
          let eventTitle = eventTitleList[k];
          let start = eventCellText.indexOf(eventTitle);
          let end = start + eventTitle.length - 1;
          let eventUrl = eventCellRichTextValue.getLinkUrl(start, end);
          calendar.createAllDayEvent(eventTitle,
            date, {description: `${eventUrl}<br><br><b><a href="${spreadsheet.getUrl()}#gid=${sheet.getSheetId()}&range=${eventNameColumn.getCell(j+1, 1).getA1Notation()}">Click here to see this event in the Year Calendar spreadsheet</a></b>`});
        }
      } else {
        console.log (`Entry "${eventCellText}" on ${date} did not have both a date and an event name. Ignoring.`)
      }
    }
  }
}
