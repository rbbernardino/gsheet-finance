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

}
