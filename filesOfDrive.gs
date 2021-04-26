function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}

function getMainSpreadsheet(){
  return getSpreadsheet(SpreadsheetApp.getActiveSpreadsheet());
}
function getListSpreadsheet(messageWithListId) {
  let listId = messageWithListId.toString().split(',');
  let listOfSpreadSheets = [];
  for(let i=0;i<listId.length;i++){
    let spreadsheet = SpreadsheetApp.openById(listId[i]);
    listOfSpreadSheets.push(getSpreadsheet(spreadsheet));
  }
  return listOfSpreadSheets;
}
function getSpreadsheet(spreadsheet) {
    let itemOfSpreadsheet = {
      nameSpreadsheet: "",
      id: "",
      listOfNameSheet: []
    };
    itemOfSpreadsheet.nameSpreadsheet = spreadsheet.getName();
    itemOfSpreadsheet.id = spreadsheet.getId();
    sheetsOfFile = spreadsheet.getSheets();
    for (i = 0; i < sheetsOfFile.length; i++) {
      itemOfSpreadsheet.listOfNameSheet.push(sheetsOfFile[i].getName())
    }
    return itemOfSpreadsheet;
}
