/**
 * blocks.js - Обработка блоков данных для библиотеки CalcMProLib
 * Предоставляет функцию XBLOCK для автоматической обработки блоков данных
 */

/**
 * Обрабатывает блок с указанным номером, считывает данные,
 * применяет формулы и записывает результат
 */
function XBLOCK(index) {
  // Проверяем корректность номера блока
  if (!index || index < 1 || index > 14) {
    return [["Ошибка: номер блока должен быть от 1 до 14"]];
  }
    try {
    // Убедимся, что библиотека инициализирована
    if (typeof FormulasLib === 'undefined' || !FormulasLib.SUPPLIER_ARTICLE) {
      init();
    }
    
    // Получаем активную таблицу
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Определяем имена листов для этого блока
    const blockSheetName = `Блок${index}`;
    const dataSheetName = `Лист${index}.0`;
    
    // Получаем листы
    const blockSheet = ss.getSheetByName(blockSheetName);
    const dataSheet = ss.getSheetByName(dataSheetName);
    const modelSheet = ss.getSheetByName("МодельСМЕТЫ");
    
    // Проверяем наличие листов
    if (!blockSheet) throw new Error(`Лист "${blockSheetName}" не найден`);
    if (!dataSheet) throw new Error(`Лист "${dataSheetName}" не найден`);
    if (!modelSheet) throw new Error(`Лист "МодельСМЕТЫ" не найден`);
    
    // Собираем входные данные для формул
    const a2 = modelSheet.getRange("A2").getValue();
    const modelSmetaCell = `B${26 + index}`;
    const bValue = modelSheet.getRange(modelSmetaCell).getValue();
    const aRange = blockSheet.getRange("A3:A38").getValues();
    const dRange = blockSheet.getRange("D3:D38").getValues();
    const eRange = blockSheet.getRange("E3:E38").getValues();
    const gRange = blockSheet.getRange("G3:G38").getValues();
    const iRange = blockSheet.getRange("I3:I38").getValues();
    const i2 = blockSheet.getRange("I2").getValue();
    const flagC28 = blockSheet.getRange("C28").getValue() === true;
    const valueE28 = blockSheet.getRange("E28").getValue();
    
    // Собираем результаты всех формул
    const results = {};
    
    // Вычисляем каждую формулу для блока
    results.B = FormulasLib.MY_CUSTOM_FUNC(a2, bValue, aRange, index);
    results.A = FormulasLib.FLAG_CHECK(dRange, flagC28, valueE28, index);
    results.C = FormulasLib.SUPPLIER_ARTICLE(dRange, index);
    results.F = FormulasLib.UNIT_MEASURE(dRange, index);
    results.H = FormulasLib.SUM_PRODUCT_ARR(a2, eRange, gRange, index);
    results.L = FormulasLib.PHOTO_ARRAY(dRange, index);
    results.M = FormulasLib.SPECIFICATION_WITH_LINK(dRange, index);
    results.N = FormulasLib.CODE_MS(dRange, index);
    
    // Вычисляем ненулевое значение в ячейке O1
    const checkQuantity = FormulasLib.CHECK_QUANTITY(dRange, eRange, index);
    blockSheet.getRange("O1").setValue(checkQuantity);
    
    // Вычисляем цену с наценками
    results.G = FormulasLib.CALCULATE_PRICE_DYNAMIC(dRange, iRange, index, i2);
    
    // Если функция вызвана как пользовательская функция (@customfunction),
    // возвращаем результаты последней формулы
    if (isCustomFunctionCall()) {
      return [["XBLOCK успешно обработан для блока " + index]];
    }
    
    // Если функция вызвана как обычный скрипт, записываем результаты в лист
    writeResultsToSheet(blockSheet, results);
    
    // Возвращаем сводку для отладки
    return [["XBLOCK успешно обработан для блока " + index]];
    
  } catch (error) {
    return [["Ошибка при обработке блока: " + error.message]];
  }
}

/**
 * Записывает результаты формул в соответствующие столбцы листа
 * @param {Sheet} sheet - Лист, куда нужно записать результаты
 * @param {Object} results - Объект с результатами формул по столбцам
 */
function writeResultsToSheet(sheet, results) {
  // Для каждого результата записываем его в соответствующий столбец
  for (const column in results) {
    if (results.hasOwnProperty(column)) {
      // Получаем результаты формулы
      const columnData = results[column];
      
      // Число строк для результата
      const numRows = columnData.length;
      
      // Определяем диапазон для записи (A1:A{numRows})
      const range = sheet.getRange(`${column}1:${column}${numRows}`);
      
      // Записываем результаты
      range.setValues(columnData);
    }
  }
}

/**
 * Определяет, вызвана ли функция как пользовательская функция в ячейке
 * @return {boolean} true, если вызвана как @customfunction
 */
function isCustomFunctionCall() {
  // В Google Apps Script трудно точно определить, вызвана ли
  // функция как customfunction или как обычный скрипт.
  // Используем эвристику на основе контекста вызова.
  return false; // Пока всегда пишем в лист
}

/**
 * Обрабатывает все блоки сразу
 * Удобно для обновления всех блоков одним нажатием
 */
function processAllBlocks() {
  for (let i = 1; i <= 14; i++) {
    XBLOCK(i);
  }
  SpreadsheetApp.getUi().alert("Все блоки успешно обработаны");
}

/**
 * Обрабатывает блок с указанным номером и возвращает результат
 * Удобно для вызова из других функций или скриптов
 */
function processXBLOCK(index) {
  // Вызов оригинальной функции XBLOCK
  return XBLOCK(index);
}

/**
 * Инициализирует модуль blocks.js и экспортирует его функции
 * @param {Object} exports - Объект для экспорта функций
 */
function blocksInit(exports) {
  exports.XBLOCK = XBLOCK;
  exports.processAllBlocks = processAllBlocks;
  exports.processXBLOCK = processXBLOCK;
  
  return exports;
}