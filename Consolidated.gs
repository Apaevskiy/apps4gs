function consolidation(listSheets) {
  CacheService.getScriptCache().put('NumberOfSheets', JSON.stringify({number: 0, lengthConsolidation: listSheets.length}), 3);
  let files = DriveApp.getFilesByName('Consolidated sheets'); // Ищем файл по имени
  let yourNewSpreadsheet;
  if(files.hasNext()){  // Если есть совпадения, то открываем
    yourNewSpreadsheet = SpreadsheetApp.openById(files.next().getId());
  } else {  // Если нет, то создаём
    yourNewSpreadsheet = SpreadsheetApp.create("Consolidated sheets");
  }

  let yourNewSheet = yourNewSpreadsheet.getSheetByName("SheetConsolidation");  // Тот же принцип, но с листом
  if (yourNewSheet != null) {
    yourNewSpreadsheet.deleteSheet(yourNewSheet);
  }
  yourNewSheet = yourNewSpreadsheet.insertSheet("SheetConsolidation");
  
  let setTagRow = new Set();
  let setTagCol = new Set();
  let numberOfSheets = 0; //  Подсчёт количесва листов
  for (let itemSheet of listSheets) {
    CacheService.getScriptCache().put('ProccesConsolidationOfSheet', JSON.stringify({percent: 0}),30);
    let spreadsheet = SpreadsheetApp.openById(itemSheet.idFile);
    let sheet = spreadsheet.getSheetByName(itemSheet.nameSheet);


    let values = sheet.getDataRange().getValues();  // Данные с выбранных листов

    let array_row = values.map(function(value) {
      return value[0];
    }); // Получение заголовков строк
    let array_column = values[0]; // Получение заголовков столбцов

    mapTagRow = new Map();  // Map для хранения индексов заголовков {key: "QTY1", value: 1}
    mapTagCol = new Map();  // По-идее при больших объёмах данных это должно ускорить работу

    for (let value of array_row) setTagRow.add(value);
    for (let value of array_column) setTagCol.add(value);

    let newSheetValues = yourNewSheet.getRange(2, 2, setTagRow.size-1, setTagCol.size-1).getValues();
    // Так как лист пустой, то DataRange возвращает только одну ячейку, потому беру область под вводимые данные

    let index = 0;  // Запись индексов к заголовкам
    for (let value of setTagRow) mapTagRow.set(value, index++);
    index = 0;
    for (let value of setTagCol) mapTagCol.set(value, index++);

    for (let i = 1; i < values.length; i++) {
      for (let j = 1; j < values[0].length; j++) {
        let col = mapTagCol.get(values[0][j])-1;  // Берём из Map'а индекс строки и столбца
        let row = mapTagRow.get(values[i][0])-1;
        if (typeof values[i][j] == 'number' && values[i][j] != null) {
          newSheetValues[row][col] = (newSheetValues[row][col]==null) ? Number(values[i][j]) : Number(newSheetValues[row][col])+values[i][j];
          
        } // Если данных по этой переменной ещё нет, то записывается соответствующее значение
      }
      CacheService.getScriptCache().put('ProccesConsolidationOfSheet', JSON.stringify({percent: Math.floor(i / values.length*100)}),3);
    }
    numberOfSheets++;
    CacheService.getScriptCache().put('NumberOfSheets', 
    JSON.stringify({number: numberOfSheets, lengthConsolidation: listSheets.length}), 3);
    yourNewSheet.getRange(2, 2, setTagRow.size-1, setTagCol.size-1).setValues(newSheetValues);
    // Записывается результат в таблицу
  }
  let array_row = Array.from(setTagRow);
  let array_column = Array.from(setTagCol); // Добавление заголовков в созданный лист

  yourNewSheet.getRange(1, 1, 1, array_column.length).setValues([array_column]);
  yourNewSheet.getRange(1, 1, array_row.length, 1).setValues(transpose(array_row));
  return {numberOfSheets: numberOfSheets, url: yourNewSpreadsheet.getUrl()};
}
function getDataForProgressBar(){
  let bufferNumberSheet = JSON.parse(CacheService.getScriptCache().get('NumberOfSheets'));
  let bufferPercent = JSON.parse(CacheService.getScriptCache().get('ProccesConsolidationOfSheet'));
  return {sheet: bufferNumberSheet, percentWork: bufferPercent};
}
function transpose(arr) { // Транспонирование заголовков, пришлось добавить, т.к. столбец заполнить однострочным массивом не даёт
  let newArr = [];
  for (let i = 0; i < arr.length; i++) {
    newArr.push([]);
    newArr[i][0] = [arr[i]];
  }
  return newArr;
}
