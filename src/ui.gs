function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('YearSpread Menu')
      .addItem('Toggle Wrap - Wrap or clip all cells in active sheet', 'toggleWrapStrategy')
      .addToUi();
}

/*
 * Toggle the wrap strategy of all cells in the active sheet between WRAP and CLIP.
 */
function toggleWrapStrategy() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let range = sheet.getDataRange();
  
  let currentStrategy = range.getCell(1, 1).getWrapStrategy();
  
  if (currentStrategy === SpreadsheetApp.WrapStrategy.WRAP) {
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  } else {
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
}