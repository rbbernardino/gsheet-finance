function onOpen() {
  createMenu();
}

function createMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('RodCash')
    .addItem('Abrir Menu', 'openSideBar')
    .addToUi();
}


function openSideBar() {
  var html = HtmlService.createTemplateFromFile('sidebar')
    .evaluate()
    .setTitle('RodCash')
    .setWidth(270);
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

// used by createTemplateFromFile
function include(filename: string) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function ordenarPorValor() {
  const comparadorSheet = SpreadsheetApp.getActive().getSheetByName("Comparador");
  if (comparadorSheet) {
    comparadorSheet.getRange("A3:C")
      .sort([{ column: 3, ascending: true }]);

    comparadorSheet.getRange("E3:G")
      .sort([{ column: 5, ascending: true }]);
  }
}

function alinharValores() {
  const comparadorSheet = SpreadsheetApp.getActive().getSheetByName("Comparador");
  if (comparadorSheet) {
    const lastFilledBank = getLastFilledLine(comparadorSheet, "E", 3);
    for (let curLn = 3; curLn < lastFilledBank; curLn++) {
      const curMobillsValue: number = comparadorSheet.getRange(`C${curLn}`).getValue();
      const curBankValue: number = comparadorSheet.getRange(`E${curLn}`).getValue();
      if (Math.abs(curMobillsValue - curBankValue) > 0.02 && curMobillsValue > curBankValue) {
        comparadorSheet.getRange(curLn, 1, 1, 3) // columns "A:C"
          .insertCells(SpreadsheetApp.Dimension.ROWS);
      }
    }
  }
}

function getLastFilledLine(sheet: GoogleAppsScript.Spreadsheet.Sheet, colLetter: string, startFrom = 0) {
  var colData = sheet.getRange(colLetter + ":" + colLetter).getValues();
  for (let i = startFrom; i < colData.length; i++) { // length --> height --> rows
    const cellData = colData[i][0];
    if (cellData == "")
      return i;
  }
  return colData.length;
}
