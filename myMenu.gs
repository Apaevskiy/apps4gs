function onOpen() {
  initMenu();
}

function initMenu() {
  let menu = SpreadsheetApp.getUi().createMenu("👉CONSOLIDATE👈")
  menu.addItem("👍Run addon", "getConsolidateForm");
  menu.addToUi();
}

function getConsolidateForm() {
  const htmlOutput = HtmlService.createTemplateFromFile("mainHTML").evaluate();
  htmlOutput.setTitle('Consolidate sheets');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}
function getPickerForm() {
  const htmlOutput = HtmlService.createTemplateFromFile("googlePickerHTML").evaluate();
  htmlOutput.setHeight(600);
  htmlOutput.setWidth(800);
  htmlOutput.setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Select Folder');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
    .getContent();
}
