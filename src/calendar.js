/* If you identify any issues with this code or would like to submit a feature request for review, please submit an issue in the associated GitHub repository: https://github.com/quinn-p-mchugh/year-calendar */

const CALENDAR_ID = SpreadsheetApp.getActiveSpreadsheet()
  .getSheetByName("⚙️CONFIG")
  .getRange("B3")
  .getValue(); // The ID of the calendar to use for syncing.

/**
 * Sync events in "Year Calendar" spreadsheet with a specified Google Calendar.
 *
 * ⚠️ WARNING: All events the specified calendar will be DELETED prior to syncing.
 */
function syncEventsWithGoogleCalendar() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    if (!sheet.getRange("A1").getValue().includes("🗓️")) {
      continue;
    }
    let sheetStartDate = sheet.getRange("A3").getValue();
    let startDate = new Date(sheetStartDate.getFullYear(), 0, 1);
    let endDate = new Date(sheetStartDate.getFullYear(), 11, 31);
    deleteEvents(calendar, startDate, endDate);
    createEventsFromSheet(calendar, sheet);
  }
}

/*
 * Delete all events between two specified dates.
 *
 * @param {Calendar} calendar - The calendar to be searched.
 * @param {Date} dateStart - The start date from which events should be deleted (inclusive).
 * @param {Date} dateEnd - The end date from which events should be deleted (inclusive).
 */
function deleteEvents(calendar, dateStart, dateEnd) {
  let events = calendar.getEvents(dateStart, dateEnd);
  for (let i = 0; i < events.length; i++) {
    events[i].deleteEvent();
  }
}

/*
 * Create Google Calendar events from entries in the spreadsheet. Entries separated by a "|" symbol will be treated as separate events.
 *
 * @param {Calendar} calendar - The calendar to add events to.
 * @param {Sheet} sheet - The sheet from which events will be created.
 */
function createEventsFromSheet(calendar, sheet) {
  const eventDataColumns = [
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
  ]; // Column numbers for event data. Each item in the list indicates the column index of the date and the event title for a given month.

  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const sheetId = sheet.getSheetId();
  for (let i = 0; i < eventDataColumns.length; i++) {
    const dateColumnValues = sheet
      .getRange(3, eventDataColumns[i][0], sheet.getLastRow() - 1)
      .getValues();
    const eventNameColumn = sheet.getRange(
      3,
      eventDataColumns[i][1],
      sheet.getLastRow() - 1
    );
    const eventNameColumnRichTextValues = eventNameColumn.getRichTextValues();

    for (let j = 0; j < dateColumnValues.length; j++) {
      const date = dateColumnValues[j][0];
      const eventCellRichTextValue = eventNameColumnRichTextValues[j][0];
      const eventCellText = eventCellRichTextValue.getText();
      if (!date || !eventCellText) {
        //console.log(`Entry "${eventCellText}" on ${date} did not have both a date and an event name. Ignoring.`);
        continue;
      }

      const eventTitleEntryList = eventCellText.split("|");
      for (let k = 0; k < eventTitleEntryList.length; k++) {
        const eventTitleEntry = eventTitleEntryList[k];
        const eventTitleCleaned = eventTitleEntry.trim();
        const start = eventCellText.indexOf(eventTitleEntry);
        const end = start + eventTitleEntry.length - 1;
        const eventUrl = eventCellRichTextValue.getLinkUrl(start, end);
        let eventUrlText = `${eventUrl}<br><br></br>`;
        if (!eventUrl) {
          eventUrlText = "";
        }
        calendar.createAllDayEvent(eventTitleCleaned, date, {
          description: `${eventUrlText}<b><a href="${spreadsheetUrl}#gid=${sheetId}&range=${eventNameColumn
            .getCell(j + 1, 1)
            .getA1Notation()}">Click here to see this event in the Year Calendar spreadsheet</a></b>`,
        });
        console.log(
          `Google Calendar event created for "${eventTitleCleaned}" on ${date}`
        );
      }
    }
  }
}
