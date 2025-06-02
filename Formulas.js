// А1

function FLAG_CHECK(dRange, flagC28, valueE28) {
  if (!Array.isArray(dRange)) dRange = [[dRange]];
  
  var result = [];
  result.push(["ФЛАГ ПРОВЕРКИ"]);

  var secondRowValue = (flagC28 === true) ? valueE28 : 0;
  result.push([secondRowValue]);

  for (var i = 0; i < dRange.length; i++) {
    var val = dRange[i][0];
    result.push([(val === "" || val === null || val === undefined) ? 0 : 1]);
  }

  result.push([""]);
  result.push([""]);

  result.push([secondRowValue === 0 ? 0 : 1]);

  return result;
}

// B1

function MY_CUSTOM_FUNC(a2, b27, aRange) {
  // a2 — значение из $A$2 (число)
  // b27 — значение из 'МодельСМЕТЫ'!B27 (одна ячейка)
  // aRange — массив из A3:A38 (2D массив)
  
  // Преобразуем aRange в плоский массив чисел
  if (!Array.isArray(aRange)) aRange = [[aRange]];
  var flatRange = aRange.map(row => row[0]);
  
  var result = [];
  
  // 1) Заголовок
  result.push(["№"]);
  
  // 2) Значение из B27
  result.push([b27]);
  
  // 3) Блок с условием и накопительной суммой
  if (a2 === 0 || a2 === "0") {
    // Если a2=0 — выводим 0 столько строк, сколько элементов в диапазоне
    for (var i = 0; i < flatRange.length; i++) {
      result.push([0]);
    }
  } else {
    // a2 != 0 — строим массив согласно условию
    var cumulativeSum = 0;
    for (var i = 0; i < flatRange.length; i++) {
      var val = flatRange[i];
      if (val !== 1 && val !== "1") {
        // Если значение не равно 1 — пустая строка
        result.push([""]);
      } else {
        // Если равно 1 — формируем строку a2 + "." + накопленная сумма с добавлением текущего значения
        cumulativeSum += Number(val) || 0;
        result.push([String(a2) + "." + cumulativeSum]);
      }
    }
  }
  
  // 4) Пустая строка
  result.push([""]);
  
  // 5) Текст-комментарий
  result.push(["Если необходимы уточнения, отметьте их здесь (B41) в строку. Данные перенесутся как комментарий к БЛОКУ."]);
  
  return result;
}

// C1
//
// =SUPPLIER_ARTICLE(D3:D38; "Лист3.0")


function SUPPLIER_ARTICLE(dRange, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Если имя листа не передано, ставим 'Лист1.0' по умолчанию
  sheetName = sheetName || 'Лист1.0';

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
  result.push([""]); // Пустая строка

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



// F- единицы измерения
// формула
// =UNIT_MEASURE(D3:D38; "Лист3.0")


function UNIT_MEASURE(dRange, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Если не указан sheetName, используем 'Лист1.0' по умолчанию
  sheetName = sheetName || 'Лист1.0';
  
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


// H - сумма
//
// =SUM_PRODUCT_ARR(A2; E3:E38; G3:G38)

function SUM_PRODUCT_ARR(a2, eRange, gRange) {
  // a2 — значение из A2 (число)
  // eRange, gRange — массивы значений из E3:E38 и G3:G38
  
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

// L фото - неизвестно пока

function PHOTO_ARRAY(dRange) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Лист1.0');

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

// M - спецификаия - не кликабельна пока
//
// =SPECIFICATION_WITH_LINK(D3:D38; "Лист3.0")


function SPECIFICATION_WITH_LINK(dRange, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetName = sheetName || 'Лист1.0';
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
        result.push(['=HYPERLINK("' + url + '"; "' + displayText + '")']);
      } else {
        result.push([item.text || ""]);
      }
    }
  }

  return result;
}



// N - код МС
//
// =CODE_MS(D3:D38; "Лист5.0")


function CODE_MS(dRange, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetName = sheetName || 'Лист1.0';
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


// формула для напоминания внести количество

function CHECK_QUANTITY(dRange, eRange) {
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

// Стоимоть согласно наценкам

function CALCULATE_PRICE_DYNAMIC(dRange, iRange, sheetNameList, markupNameValue) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheetList = ss.getSheetByName(sheetNameList || 'Лист2.0');
  var sheetSmeta = ss.getSheetByName('СМЕТА');
  var sheetMarkup = ss.getSheetByName('НАЦЕНКИ');

  if (!sheetList || !sheetSmeta || !sheetMarkup) {
    throw new Error('Один или несколько листов не найдены. Проверьте имена.');
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
  for (var row of lookupRange) {
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
