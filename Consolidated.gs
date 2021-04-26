function consolidation(listSheets){
  var activeSpreadsheet = SpreadsheetApp.create("Consolidated sheets");
  //var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var yourNewSheet = activeSpreadsheet.getSheetByName("SheetConsolidation");

    if (yourNewSheet != null) {
        activeSpreadsheet.deleteSheet(yourNewSheet);
    }
  yourNewSheet = activeSpreadsheet.insertSheet("SheetConsolidation");

  let listTagInSheets = getTagInSheets(listSheets);
  let list = listTagInSheets.tags;
  let listOfColumns = listCol(list);
  let listOfRows = listRow(list);
  yourNewSheet.getRange(1,1,1,listOfColumns.length).setValues([listOfColumns]);
  yourNewSheet.getRange(2,1,listOfRows.length, 1).setValues(transpose(listOfRows));
  
  for(let i = 0; i < list.length; i++){
    let row = getIndexOfTagRow(yourNewSheet, list[i].key.row)
    let col = getIndexOfTagCol(yourNewSheet, list[i].key.col)
    yourNewSheet.getRange(row+1,col+1,1,1).setValue(list[i].value);
  }
  return {numberOfSheets: listTagInSheets.numberOfSheets, url: activeSpreadsheet.getUrl()};
}

function getTagInSheets(listSheets) {
  let list = [];
  for (let table of listSheets) {
    
    let spreadsheet = SpreadsheetApp.openById(table.idFile);
    let sheet = spreadsheet.getSheetByName(table.nameSheet);
    var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
    var values = range.getValues();
    
    let result = {
      key: {
        row: 0,
        col: 0
      },
      value: 0
    };
    for (let row = 1; row < values.length; row++) {
      for (let col = 1; col < values[row].length; col++) {
        keyOfValues = {
          row: values[row][0],
          col: values[0][col]
        };
        
        if(typeof values[row][col] === 'number'){
        let item = listHas(list, keyOfValues);
        if (item != null) {
          item.value += values[row][col];
        } else {
          list.push({
            key: keyOfValues,
            value: values[row][col]
          });
        }
        }
      }
    }
  }
  list.sort(function(a,b) {
	  var x = a.key.row;
	  var y = b.key.row;
	  return x < y ? -1 : x > y ? 1 : 0;
  });
  return {tags: list, numberOfSheets: listSheets.length};
}

function listCol(resultConsolidation) {
  let bufferList = new Set(['Product']);
  for (let item of resultConsolidation) {
      bufferList.add(item.key.col);
  }
  return Array.from(bufferList);
}
function listRow(resultConsolidation) {
  let bufferList = new Set();
  for (let item of resultConsolidation) {
      bufferList.add(item.key.row);
  }
  return Array.from(bufferList);
}

function listHas(list, newKey) {
  for (let item of list) {
    if (item.key.row == newKey.row && item.key.col == newKey.col) {
      return item;
    }
  }
  return null;
}

function getIndexOfTagRow(sheet,name){
  var values = sheet.getDataRange().getValues();
  let column = [];
  for(let i=0;i<values.length;i++){
    column.push(values[i][0]);
  }
  let headerIdx = column.indexOf(name);
  return headerIdx;
}
function getIndexOfTagCol(sheet, name){
  var values = sheet.getDataRange().getValues().splice(0,1);
  let headerIdx = values[0].indexOf(name);
  return headerIdx;
}

function transpose(arr) {
  let newArr = [];
  for(let i = 0; i < arr.length; i++){
    newArr.push([]);
    newArr[i][0] = [arr[i]];
  }
  return newArr;
}
