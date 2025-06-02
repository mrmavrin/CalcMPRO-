/**
 * CalcMProLib.js - Библиотека формул для работы с данными в Google Sheets
 * Предоставляет набор пользовательских функций для использования в Google Sheets
 * через библиотечную связь.
 * 
 * Версия: 1.0.0 (Библиотека)
 */

// ======= НАСТРОЙКИ ПО УМОЛЧАНИЮ (могут быть переопределены) =======

let DEFAULT_LOOKUP_SHEET_FORMAT = "Лист{blockNum}.0";
let DEFAULT_PRICE_SHEET = "Лист2.0";
let DEFAULT_MARKUP_SHEET = "НАЦЕНКИ";
let DEFAULT_ESTIMATE_SHEET = "МодельСМЕТЫ";
let DEFAULT_BLOCK_COUNT = 14;
let DEFAULT_MODEL_SMETA_FIRST_CELL = "B27"; // Начиная с этой ячейки для Блок1

// Массив соответствия блоков с их ячейками и листами
let BLOCK_MAPPINGS = [];

/**
 * Инициализирует соответствия для блоков
 */
function initBlockMappings() {
  BLOCK_MAPPINGS = [];
  for (let i = 1; i <= DEFAULT_BLOCK_COUNT; i++) {
    const cellRow = 26 + i; // B27, B28, ..., B40
    BLOCK_MAPPINGS.push({
      blockNum: i,
      lookupSheet: DEFAULT_LOOKUP_SHEET_FORMAT.replace("{blockNum}", i),
      modelSmetaCell: "B" + cellRow
    });
  }
}

// Инициализируем соответствия по умолчанию
initBlockMappings();

/**
 * Функция инициализации библиотеки с пользовательскими настройками
 * @param {Object} options - Объект с настройками
 * @param {string} [options.lookupSheetFormat] - Шаблон имени справочного листа (напр. "Лист{blockNum}.0")
 * @param {string} [options.priceSheet] - Имя листа с ценами (по умолчанию "Лист2.0")
 * @param {string} [options.markupSheet] - Имя листа с наценками (по умолчанию "НАЦЕНКИ") 
 * @param {string} [options.estimateSheet] - Имя листа сметы (по умолчанию "МодельСМЕТЫ")
 * @param {number} [options.blockCount] - Количество блоков (по умолчанию 14)
 * @param {Array<Object>} [options.customMappings] - Полностью настраиваемые соответствия блоков
 */
function init(options = {}) {
  if (options.lookupSheetFormat) DEFAULT_LOOKUP_SHEET_FORMAT = options.lookupSheetFormat;
  if (options.priceSheet) DEFAULT_PRICE_SHEET = options.priceSheet;
  if (options.markupSheet) DEFAULT_MARKUP_SHEET = options.markupSheet;
  if (options.estimateSheet) DEFAULT_ESTIMATE_SHEET = options.estimateSheet;
  
  let mappingsChanged = false;
  
  if (options.blockCount !== undefined && options.blockCount !== DEFAULT_BLOCK_COUNT) {
    DEFAULT_BLOCK_COUNT = options.blockCount;
    mappingsChanged = true;
  }
  
  if (mappingsChanged) {
    initBlockMappings();
  }
  
  // Если предоставлены пользовательские соответствия, используем их
  if (options.customMappings) {
    BLOCK_MAPPINGS = options.customMappings;
  }
  
  return {
    lookupSheetFormat: DEFAULT_LOOKUP_SHEET_FORMAT,
    priceSheet: DEFAULT_PRICE_SHEET,
    markupSheet: DEFAULT_MARKUP_SHEET,
    estimateSheet: DEFAULT_ESTIMATE_SHEET,
    blockCount: DEFAULT_BLOCK_COUNT,
    blockMappings: BLOCK_MAPPINGS
  };
}

/**
 * Получает информацию о блоке по его номеру
 * @param {number} blockNum - Номер блока (1-14)
 * @return {Object} Информация о блоке
 */
function getBlockInfo(blockNum) {
  if (!blockNum || blockNum < 1 || blockNum > DEFAULT_BLOCK_COUNT) {
    throw new Error(`Неверный номер блока: ${blockNum}. Номер блока должен быть от 1 до ${DEFAULT_BLOCK_COUNT}`);
  }
  
  return BLOCK_MAPPINGS[blockNum - 1];
}

/**
 * Проверка флага и установка соответствующего значения в зависимости от блока
 * @param {Array} dRange - Диапазон данных для проверки
 * @param {boolean} flagC28 - Флаг проверки (не используется, сохранен для обратной совместимости)
 * @param {number} valueE28 - Значение при активном флаге (не используется, сохранен для обратной совместимости)
 * @param {number} [blockNum] - Номер блока (1-14), обязательный
 * @return {Array} Двумерный массив результатов
 */
function FLAG_CHECK(dRange, flagC28, valueE28, blockNum) {
  if (!Array.isArray(dRange)) dRange = [[dRange]];
  
  // Проверка наличия blockNum
  if (!blockNum) {
    throw new Error("Номер блока не указан. Пожалуйста, укажите номер блока (1-14).");
  }
  
  // Получаем активную таблицу
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var modelSheet = ss.getSheetByName("МодельСМЕТЫ");
  
  if (!modelSheet) {
    throw new Error("Лист 'МодельСМЕТЫ' не найден.");
  }
  
  // Вычисляем номера строк для C и E столбцов в зависимости от номера блока
  var cRowForFlag = 26 + blockNum;      // C27 для блока 1, C28 для блока 2, и т.д.
  var eRowForValue = 26 + blockNum;     // E27 для блока 1, E28 для блока 2, и т.д.
  var cRowForLastValue = 26 + blockNum; // C27 для блока 1, C28 для блока 2, и т.д.
  
  // Получаем значения из нужных ячеек
  var flagCValue = modelSheet.getRange("C" + cRowForFlag).getValue() === true;
  var valueEValue = modelSheet.getRange("E" + eRowForValue).getValue();
  var flagCLastValue = modelSheet.getRange("C" + cRowForLastValue).getValue() === true;
  
  // Формируем результат в соответствии с исходной формулой
  var result = [];
  result.push(["ФЛАГ ПРОВЕРКИ"]);
  
  // Вторая строка: ЕСЛИ('МодельСМЕТЫ'!C30=ИСТИНА;'МодельСМЕТЫ'!E27; 0)
  var secondRowValue = flagCValue ? valueEValue : 0;
  result.push([secondRowValue]);
  
  // Перебираем dRange и применяем BYROW(D3:D38; LAMBDA(d; ЕСЛИ(d=""; 0; 1)))
  for (var i = 0; i < dRange.length; i++) {
    var val = dRange[i][0];
    result.push([(val === "" || val === null || val === undefined) ? 0 : 1]);
  }
  
  // Добавляем две пустые строки
  result.push([""]);
  result.push([""]);
  
  // Последняя строка: ЕСЛИ(ЕСЛИ('МодельСМЕТЫ'!C27=ИСТИНА; 'МодельСМЕТЫ'!E27; 0)=0; 0; 1)
  var lastRowValue = flagCLastValue ? valueEValue : 0;
  result.push([lastRowValue === 0 ? 0 : 1]);
  
  return result;
}

/**
 * Пользовательская функция для формирования номеров строк
 * @param {number} a2 - Значение из $A$2 (число)
 * @param {*} b27 - Значение из ячейки модели сметы для соответствующего блока
 * @param {Array} aRange - Массив из A3:A38 (2D массив)
 * @param {number} [blockNum] - Номер блока (1-14), необязательный
 * @return {Array} Двумерный массив результатов
 */
function MY_CUSTOM_FUNC(a2, b27, aRange, blockNum) {
  // Здесь blockNum используется только в комментарии
  const blockComment = blockNum ? `Блок${blockNum}` : "БЛОКУ";
  
  if (!Array.isArray(aRange)) aRange = [[aRange]];
  var flatRange = aRange.map(row => row[0]);
  
  var result = [];
  
  result.push(["№"]);
  result.push([b27]);
  
  if (a2 === 0 || a2 === "0") {
    for (var i = 0; i < flatRange.length; i++) {
      result.push([0]);
    }
  } else {
    var cumulativeSum = 0;
    for (var i = 0; i < flatRange.length; i++) {
      var val = flatRange[i];
      if (val !== 1 && val !== "1") {
        result.push([""]);
      } else {
        cumulativeSum += Number(val) || 0;
        result.push([String(a2) + "." + cumulativeSum]);
      }
    }
  }
  
  result.push([""]);
  result.push([`Если необходимы уточнения, отметьте их здесь (B41) в строку. Данные перенесутся как комментарий к ${blockComment}.`]);
  
  return result;
}

/**
 * Находит артикул поставщика по значению из указанного листа
 * @param {Array} dRange - Диапазон с артикулами для поиска
 * @param {string|number} sheetNameOrBlockNum - Имя листа с данными поставщиков или номер блока (1-14)
 * @return {Array} Двумерный массив артикулов поставщиков
 */
function SUPPLIER_ARTICLE(dRange, sheetNameOrBlockNum) {
  let sheetName;
  
  // Определяем имя листа на основе номера блока или используем переданное имя
  if (typeof sheetNameOrBlockNum === "number" || (typeof sheetNameOrBlockNum === "string" && !isNaN(sheetNameOrBlockNum))) {
    // Если передан номер блока
    const blockNum = parseInt(sheetNameOrBlockNum);
    const blockInfo = getBlockInfo(blockNum);
    sheetName = blockInfo.lookupSheet;
  } else {
    // Если передано имя листа
    sheetName = sheetNameOrBlockNum || DEFAULT_LOOKUP_SHEET_FORMAT.replace("{blockNum}", 1);
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error('Лист с именем "' + sheetName + '" не найден');
  }

  // Получаем столбцы E и F динамически с выбранного листа
  var colE = sheet.getRange('E:E').getValues().flat();
  var colF = sheet.getRange('F:F').getValues().flat();

  // Функция очистки значения (аналог TO_PURE_NUMBER + TRIM + remove spaces)
  function cleanValue(val) {
    if (val === null || val === undefined) return "";
    return String(val).replace(/\s+/g, '').replace(/[^0-9.]/g, '');
  }

  if (!Array.isArray(dRange)) {
    dRange = [[dRange]];
  }

  var result = [];
  result.push(["Артикул постащика"]); // Заголовок
  result.push([""]);                  // Пустая строка

  for (var i = 0; i < dRange.length; i++) {
    var d = dRange[i][0];
    if (d === "" || d === null || d === undefined) {
      result.push([""]);
      continue;
    }

    var cleanedD = cleanValue(d);
    var foundIndex = -1;

    // Ищем совпадения по colF
    for (var j = 0; j < colF.length; j++) {
      var cleanedF = cleanValue(colF[j]);
      if (cleanedF === cleanedD) {
        foundIndex = j;
        break;
      }
    }

    if (foundIndex === -1) {
      result.push(["-"]);
    } else {
      var valE = colE[foundIndex];
      result.push([valE === undefined ? "-" : valE]);
    }
  }

  return result;
}

/**
 * Находит единицы измерения по артикулу
 * @param {Array} dRange - Диапазон с артикулами для поиска
 * @param {string|number} sheetNameOrBlockNum - Имя листа с данными или номер блока (1-14)
 * @return {Array} Двумерный массив единиц измерения
 */
function UNIT_MEASURE(dRange, sheetNameOrBlockNum) {
  let sheetName;
  
  // Определяем имя листа на основе номера блока или используем переданное имя
  if (typeof sheetNameOrBlockNum === "number" || (typeof sheetNameOrBlockNum === "string" && !isNaN(sheetNameOrBlockNum))) {
    // Если передан номер блока
    const blockNum = parseInt(sheetNameOrBlockNum);
    const blockInfo = getBlockInfo(blockNum);
    sheetName = blockInfo.lookupSheet;
  } else {
    // Если передано имя листа
    sheetName = sheetNameOrBlockNum || DEFAULT_LOOKUP_SHEET_FORMAT.replace("{blockNum}", 1);
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) throw new Error('Лист "' + sheetName + '" не найден');

  // Получаем диапазон F1:V (22 столбца, 1-й из них F)
  var dataRange = sheet.getRange('F1:V').getValues();

  // Приводим dRange к массиву
  if (!Array.isArray(dRange)) {
    dRange = [[dRange]];
  }

  var result = [];
  result.push(["ЕД.ИЗМ."]); // заголовок
  result.push([""]);        // пустая строка

  // Индекс для быстрого поиска: ключ из столбца F → вся строка
  var lookupMap = {};
  for (var i = 0; i < dataRange.length; i++) {
    var key = dataRange[i][0]; // столбец F — первый в диапазоне
    lookupMap[key] = dataRange[i];
  }

  for (var j = 0; j < dRange.length; j++) {
    var val = dRange[j][0];
    if (val === null || val === undefined || val === "") {
      result.push([""]);
      continue;
    }
    var row = lookupMap[val];
    if (row === undefined) {
      result.push([""]);
    } else {
      // 17-й столбец диапазона F:V → индекс 16 (нумерация с 0)
      result.push([row[16] === undefined ? "" : row[16]]);
    }
  }

  return result;
}

/**
 * Расчет суммы произведений двух массивов с учетом флага
 * @param {number} a2 - Флаг (0 или другое значение)
 * @param {Array} eRange - Первый массив значений
 * @param {Array} gRange - Второй массив значений
 * @param {number} [blockNum] - Номер блока (1-14), необязательный
 * @return {Array} Двумерный массив с результатами
 */
function SUM_PRODUCT_ARR(a2, eRange, gRange, blockNum) {
  // Здесь blockNum не используется, но включен для единообразия
  if (!Array.isArray(eRange)) eRange = [[eRange]];
  if (!Array.isArray(gRange)) gRange = [[gRange]];
  
  var arr = [];
  
  for (var i = 0; i < gRange.length; i++) {
    var gVal = gRange[i][0];
    var eVal = (i < eRange.length) ? eRange[i][0] : 0;
    if (gVal === "" || gVal === null || gVal === undefined) {
      arr.push("");
    } else {
      arr.push(eVal * gVal);
    }
  }
  
  // cleanArr — arr без пустых значений
  var cleanArr = arr.filter(function(x) { return x !== "" && x !== null && x !== undefined; });
  
  var sumValue = (a2 === 0 || a2 === "0") ? 0 : cleanArr.reduce(function(acc, val) { return acc + val; }, 0);
  
  var result = [];
  result.push(["СУММА"]);
  result.push([sumValue]);
  
  // Добавляем весь arr построчно как двумерный массив
  for (var j = 0; j < arr.length; j++) {
    result.push([arr[j]]);
  }
  
  return result;
}

/**
 * Находит ссылки на фото по артикулу
 * @param {Array} dRange - Диапазон с артикулами для поиска
 * @param {number} [blockNum] - Номер блока (1-14)
 * @return {Array} Двумерный массив ссылок на фото
 */
function PHOTO_ARRAY(dRange, blockNum) {
  let sheetName;
  
  if (blockNum) {
    const blockInfo = getBlockInfo(blockNum);
    sheetName = blockInfo.lookupSheet;
  } else {
    sheetName = DEFAULT_LOOKUP_SHEET_FORMAT.replace("{blockNum}", 1);
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  // Получаем диапазон F1:V (столбец F — ключ, 3-й столбец — индекс 2 в массиве)
  var dataRange = sheet.getRange('F1:V').getValues();

  if (!Array.isArray(dRange)) {
    dRange = [[dRange]];
  }

  var result = [];
  result.push(["ФОТО"]); // заголовок
  result.push([""]);      // пустая строка

  // Создаем индекс: ключ из столбца F → строка данных
  var lookupMap = {};
  for (var i = 0; i < dataRange.length; i++) {
    var key = dataRange[i][0]; // столбец F — 1-й в диапазоне
    lookupMap[key] = dataRange[i];
  }

  for (var j = 0; j < dRange.length; j++) {
    var val = dRange[j][0];
    if (val === null || val === undefined || val === "") {
      result.push([""]);
      continue;
    }
    var row = lookupMap[val];
    if (row === undefined) {
      result.push([""]);
    } else {
      var link = row[2]; // 3-й столбец от F — индекс 2
      if (link === undefined || link === "") {
        result.push([""]);
      } else {
        result.push([link]);
      }
    }
  }

  return result;
}

/**
 * Находит спецификацию по артикулу и создает ссылки
 * @param {Array} dRange - Диапазон с артикулами для поиска
 * @param {string|number} sheetNameOrBlockNum - Имя листа с данными или номер блока (1-14)
 * @return {Array} Двумерный массив спецификаций с гиперссылками
 */
function SPECIFICATION_WITH_LINK(dRange, sheetNameOrBlockNum) {
  let sheetName;
  
  // Определяем имя листа на основе номера блока или используем переданное имя
  if (typeof sheetNameOrBlockNum === "number" || (typeof sheetNameOrBlockNum === "string" && !isNaN(sheetNameOrBlockNum))) {
    // Если передан номер блока
    const blockNum = parseInt(sheetNameOrBlockNum);
    const blockInfo = getBlockInfo(blockNum);
    sheetName = blockInfo.lookupSheet;
  } else {
    // Если передано имя листа
    sheetName = sheetNameOrBlockNum || DEFAULT_LOOKUP_SHEET_FORMAT.replace("{blockNum}", 1);
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) throw new Error('Лист "' + sheetName + '" не найден');

  // Получаем значения и rich text диапазона F1:V
  var dataValues = sheet.getRange('F1:V').getValues();
  var richTextValues = sheet.getRange('F1:V').getRichTextValues();

  if (!Array.isArray(dRange)) {
    dRange = [[dRange]];
  }

  var result = [];
  result.push(["СПЕЦИФИКАЦИЯ"]); // заголовок
  result.push([""]);             // пустая строка

  // Индексируем по ключу (столбец F)
  var lookupMap = {};
  for (var i = 0; i < dataValues.length; i++) {
    var key = dataValues[i][0];
    lookupMap[key] = {
      text: dataValues[i][1],
      richText: richTextValues[i][1]
    };
  }

  for (var j = 0; j < dRange.length; j++) {
    var val = dRange[j][0];
    if (val === null || val === undefined || val === "") {
      result.push([""]);
      continue;
    }
    var item = lookupMap[val];
    if (!item) {
      result.push([""]);
    } else {
      var richText = item.richText;
      if (richText && richText.getLinkUrl()) {
        var url = richText.getLinkUrl();
        var displayText = richText.getText();
        result.push(['=HYPERLINK("' + url + '")']);
      } else {
        result.push([item.text || ""]);
      }
    }
  }

  return result;
}

/**
 * Находит код МС по артикулу
 * @param {Array} dRange - Диапазон с артикулами для поиска
 * @param {string|number} sheetNameOrBlockNum - Имя листа с данными или номер блока (1-14)
 * @return {Array} Двумерный массив кодов МС
 */
function CODE_MS(dRange, sheetNameOrBlockNum) {
  let sheetName;
  
  // Определяем имя листа на основе номера блока или используем переданное имя
  if (typeof sheetNameOrBlockNum === "number" || (typeof sheetNameOrBlockNum === "string" && !isNaN(sheetNameOrBlockNum))) {
    // Если передан номер блока
    const blockNum = parseInt(sheetNameOrBlockNum);
    const blockInfo = getBlockInfo(blockNum);
    sheetName = blockInfo.lookupSheet;
  } else {
    // Если передано имя листа
    sheetName = sheetNameOrBlockNum || DEFAULT_LOOKUP_SHEET_FORMAT.replace("{blockNum}", 1);
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) throw new Error('Лист "' + sheetName + '" не найден');

  var colA = sheet.getRange('A:A').getValues().flat();
  var colF = sheet.getRange('F:F').getValues().flat();

  if (!Array.isArray(dRange)) {
    dRange = [[dRange]];
  }

  var result = [];
  result.push(["КОД МС"]); // заголовок
  result.push([""]);       // пустая строка

  for (var i = 0; i < dRange.length; i++) {
    var val = dRange[i][0];
    if (val === null || val === undefined || val === "") {
      result.push([""]);
      continue;
    }

    var index = colF.indexOf(val);
    if (index === -1) {
      result.push([""]);
    } else {
      var code = colA[index];
      result.push([code === undefined ? "" : code]);
    }
  }

  return result;
}

/**
 * Проверяет наличие количества для всех артикулов
 * @param {Array} dRange - Диапазон артикулов
 * @param {Array} eRange - Диапазон количеств
 * @param {number} [blockNum] - Номер блока (1-14), необязательный
 * @return {string} Строка с предупреждением или пустая строка
 */
function CHECK_QUANTITY(dRange, eRange, blockNum) {
  // Здесь blockNum не используется, но включен для единообразия
  if (!Array.isArray(dRange)) dRange = [[dRange]];
  if (!Array.isArray(eRange)) eRange = [[eRange]];
  
  var count = 0;
  
  for (var i = 0; i < dRange.length; i++) {
    var dVal = dRange[i][0];
    var eVal = (i < eRange.length) ? eRange[i][0] : null;
    
    var dNotEmpty = dVal !== "" && dVal !== null && dVal !== undefined;
    var eEmpty = eVal === "" || eVal === null || eVal === undefined;
    
    if (dNotEmpty && eEmpty) {
      count++;
    }
  }
  
  return count > 0 ? "НЕТ КОЛ-ВА" : "";
}

/**
 * Рассчитывает цену на основе наценок
 * @param {Array} dRange - Диапазон артикулов
 * @param {Array} iRange - Диапазон дополнительных значений
 * @param {string|number} sheetNameOrBlockNum - Имя листа с прайсом или номер блока (1-14)
 * @param {string} markupNameValue - Название наценки
 * @return {Array} Двумерный массив цен
 */
function CALCULATE_PRICE_DYNAMIC(dRange, iRange, sheetNameOrBlockNum, markupNameValue) {
  let sheetNameList;
  
  // Определяем имя листа на основе номера блока или используем переданное имя
  if (typeof sheetNameOrBlockNum === "number" || (typeof sheetNameOrBlockNum === "string" && !isNaN(sheetNameOrBlockNum))) {
    // Если передан номер блока
    const blockNum = parseInt(sheetNameOrBlockNum);
    const blockInfo = getBlockInfo(blockNum);
    sheetNameList = blockInfo.lookupSheet;
  } else {
    // Если передано имя листа
    sheetNameList = sheetNameOrBlockNum || DEFAULT_PRICE_SHEET;
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetList = ss.getSheetByName(sheetNameList);
  var sheetSmeta = ss.getSheetByName(DEFAULT_ESTIMATE_SHEET);
  var sheetMarkup = ss.getSheetByName(DEFAULT_MARKUP_SHEET);

  if (!sheetList || !sheetSmeta || !sheetMarkup) {
    throw new Error('Один или несколько листов не найдены. Проверьте имена: ' + 
                   sheetNameList + ', ' + DEFAULT_ESTIMATE_SHEET + ', ' + DEFAULT_MARKUP_SHEET);
  }

  var lookupRange = sheetList.getRange('F1:Y').getValues();
  var mode = (sheetSmeta.getRange('F1').getValue() || "").toString().trim();
  var markupName = (markupNameValue || "").toString().trim().toLowerCase();

  var markupTable = sheetMarkup.getRange('B6:D20').getValues();

  function findMarkup(name) {
    name = (name || "").toString().trim().toLowerCase();
    for (var i = 0; i < markupTable.length; i++) {
      var key = (markupTable[i][0] || "").toString().trim().toLowerCase();
      if (key === name) {
        return { value: markupTable[i][1], type: markupTable[i][2] };
      }
    }
    return { value: 0, type: "" };
  }

  var globalMarkup = findMarkup("Общая наценка");
  var localMarkup = findMarkup(markupName);

  if (!Array.isArray(dRange)) dRange = [[dRange]];
  if (!Array.isArray(iRange)) iRange = [[iRange]];

  var lookupMap = {};
  for (var i = 0; i < lookupRange.length; i++) {
    var row = lookupRange[i];
    var key = row[0];
    lookupMap[key] = row[19];
  }

  var result = [];
  result.push(["ЦЕНА"]);
  result.push(["Итого:"]);

  for (var idx = 0; idx < dRange.length; idx++) {
    var d = dRange[idx][0];
    var iVal = (idx < iRange.length) ? iRange[idx][0] : "";

    var rawBase = lookupMap.hasOwnProperty(d) ? lookupMap[d] : "";
    if (rawBase === "" || rawBase === null || rawBase === undefined) rawBase = "";

    var priceWithMarkup;

    if (mode === "Запретить наценки") {
      priceWithMarkup = rawBase;
    } else {
      if (markupName === "без наценок") {
        priceWithMarkup = rawBase;
      } else if (!localMarkup.type) {
        priceWithMarkup = "Покажите эту надпись руководителю";
      } else if (localMarkup.type === "%") {
        priceWithMarkup = rawBase * (1 + localMarkup.value / 100);
      } else if (localMarkup.type === "коэф") {
        priceWithMarkup = rawBase * localMarkup.value;
      } else {
        priceWithMarkup = "Покажите эту надпись руководителю";
      }
    }

    var finalVal;
    if (d === "" || d === null || d === undefined) {
      finalVal = "";
    } else {
      if (iVal === "" || iVal === null || iVal === undefined) {
        finalVal = priceWithMarkup;
      } else {
        if (typeof priceWithMarkup === "number" && typeof iVal === "number") {
          finalVal = priceWithMarkup + iVal;
        } else {
          finalVal = priceWithMarkup;
        }
      }
    }

    result.push([finalVal]);
  }

  return result;
}

/**
 * Выполняет полный набор формул для блока с указанным номером
 * Вспомогательная функция для использования в автоматизированных скриптах
 * 
 * @param {number} blockNum - Номер блока (1-14)
 * @param {Object} data - Объект с данными для формул
 * @param {Array} data.dRange - Диапазон артикулов
 * @param {Array} data.eRange - Диапазон количеств
 * @param {Array} data.gRange - Диапазон стоимости за единицу
 * @param {Array} data.iRange - Диапазон дополнительных значений
 * @param {number} data.a2 - Значение из A2
 * @param {string} data.markupName - Название наценки
 * @param {boolean} data.flagC28 - Флаг проверки
 * @param {number} data.valueE28 - Значение при активном флаге
 * @return {Object} - Объект с результатами всех формул
 */
function processBlockFormulas(blockNum, data) {
  // Получаем информацию о блоке
  const blockInfo = getBlockInfo(blockNum);
  
  // Получаем значение из ячейки в МодельСМЕТЫ для этого блока
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetEstimate = ss.getSheetByName(DEFAULT_ESTIMATE_SHEET);
  var modelSmetaValue = sheetEstimate.getRange(blockInfo.modelSmetaCell).getValue();
  
  // Выполняем все формулы
  return {
    blockNum: blockNum,
    blockInfo: blockInfo, 
    customFunc: MY_CUSTOM_FUNC(data.a2, modelSmetaValue, data.dRange, blockNum),
    flag: FLAG_CHECK(data.dRange, data.flagC28, data.valueE28, blockNum),
    supplierArticle: SUPPLIER_ARTICLE(data.dRange, blockNum),
    unitMeasure: UNIT_MEASURE(data.dRange, blockNum),
    sumProduct: SUM_PRODUCT_ARR(data.a2, data.eRange, data.gRange, blockNum),
    photo: PHOTO_ARRAY(data.dRange, blockNum),
    spec: SPECIFICATION_WITH_LINK(data.dRange, blockNum),
    code: CODE_MS(data.dRange, blockNum),
    checkQuantity: CHECK_QUANTITY(data.dRange, data.eRange, blockNum),
    price: CALCULATE_PRICE_DYNAMIC(data.dRange, data.iRange, blockNum, data.markupName)
  };
}

/**
 * Создает меню для работы с формулами
 * @param {Object} [options] - Настройки меню
 */
function createMenu(options = {}) {
  const menuName = options.menuName || "Формулы МС";
  
  SpreadsheetApp.getUi()
    .createMenu(menuName)
    .addItem("Настроить библиотеку", "CalcMProLib.showConfigDialog")
    .addItem("О библиотеке", "CalcMProLib.showInfoDialog")
    .addToUi();
}

/**
 * Показывает диалоговое окно с информацией о библиотеке
 */
function showInfoDialog() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    "О библиотеке CalcMProLib",
    "CalcMProLib v1.0.0\n\n" +
    "Библиотека предоставляет набор функций для работы с блоками данных.\n" +
    "Каждая функция может принимать номер блока (1-14) в качестве параметра.\n\n" + 
    "Пример использования:\n" +
    "=CalcMProLib.SUPPLIER_ARTICLE(D3:D38; 1) - для Блок1\n" +
    "=CalcMProLib.SUPPLIER_ARTICLE(D3:D38; 2) - для Блок2",
    ui.ButtonSet.OK
  );
}

/**
 * Инициализирует объект FormulasLib
 * @param {Object} exports - Объект для экспорта функций
 */
function FormulasLibInit(exports) {
  exports.FLAG_CHECK = FLAG_CHECK;
  exports.MY_CUSTOM_FUNC = MY_CUSTOM_FUNC;
  exports.SUPPLIER_ARTICLE = SUPPLIER_ARTICLE;
  exports.UNIT_MEASURE = UNIT_MEASURE;
  exports.SUM_PRODUCT_ARR = SUM_PRODUCT_ARR;
  exports.PHOTO_ARRAY = PHOTO_ARRAY;
  exports.SPECIFICATION_WITH_LINK = SPECIFICATION_WITH_LINK;  exports.CODE_MS = CODE_MS;
  exports.CHECK_QUANTITY = CHECK_QUANTITY;
  exports.CALCULATE_PRICE_DYNAMIC = CALCULATE_PRICE_DYNAMIC;
  // Удаляем экспорт XBLOCK из FormulasLib, так как теперь он экспортируется из blocks.js
  exports.init = init;
  
  return exports;
}

// ... оставшийся код FormulasLib.js без изменений ...