/**
 * ReestrOffers.js - Библиотека для сохранения и восстановления данных в Google Sheets
 * Функциональность: создаёт "отпечаток" состояния определённых ячеек таблицы и позволяет
 * восстанавливать эти состояния. По сути - система версионного контроля для данных.
 * 
 * Версия: 1.0.0 (Библиотека)
 */

// ======= НАСТРОЙКИ ПО УМОЛЧАНИЮ (могут быть переопределены) =======

// Используем let вместо const, чтобы можно было переопределять значения
let STORAGE_SHEET_NAME = "Docs";
let REGISTER_SHEET_NAME = "Реестр";
let MODEL_SHEET_NAME = "МодельСМЕТЫ"; // Добавляем настройку имени листа модели

// Диапазоны для сохранения (могут быть переопределены)
let MODEL_SMETA_RANGES = [
  "B13:B18", // 6 ячеек - основные реквизиты
  "B20",     // отдельная ячейка - номер документа
  "B22",     // отдельная ячейка - дата
  "B24",     // отдельная ячейка - основание
  "G27:G40"  // 14 ячеек - итоговые суммы
];

// Блоки для сохранения (могут быть переопределены)
let BLOCKS = [];

// Количество блоков (может быть переопределено)
let NUM_BLOCKS = 14;

// Имя блока (может быть переопределено)
let BLOCK_NAME_TEMPLATE = "Блок{i}";

// Диапазоны блоков по умолчанию
let DEFAULT_BLOCK_RANGES = ["D3:D38", "E3:E38", "I2:I38"];
let DEFAULT_BLOCK_CELLS = ["B41"];

// Функция инициализации блоков
function initializeBlocks() {
  BLOCKS = [];
  for(let i=1; i<=NUM_BLOCKS; i++) {
    BLOCKS.push({
      sheetName: BLOCK_NAME_TEMPLATE.replace("{i}", i),
      ranges: DEFAULT_BLOCK_RANGES.slice(), // Копируем массив
      singleCells: DEFAULT_BLOCK_CELLS.slice() // Копируем массив
    });
  }
}

// Инициализируем блоки по умолчанию
initializeBlocks();

/**
 * Функция инициализации библиотеки с пользовательскими настройками
 * @param {Object} options - Объект с настройками
 * @param {string} [options.storageSheetName] - Имя листа для хранения данных
 * @param {string} [options.registerSheetName] - Имя листа реестра
 * @param {string} [options.modelSheetName] - Имя листа модели сметы
 * @param {Array<string>} [options.modelRanges] - Массив диапазонов модели
 * @param {number} [options.numBlocks] - Количество блоков
 * @param {string} [options.blockNameTemplate] - Шаблон имени блока
 * @param {Array<string>} [options.blockRanges] - Диапазоны для всех блоков
 * @param {Array<string>} [options.blockCells] - Отдельные ячейки для всех блоков
 * @param {Array<Object>} [options.customBlocks] - Полностью настраиваемые блоки
 */
function init(options = {}) {
  // Настраиваем имена листов
  if (options.storageSheetName) STORAGE_SHEET_NAME = options.storageSheetName;
  if (options.registerSheetName) REGISTER_SHEET_NAME = options.registerSheetName;
  if (options.modelSheetName) MODEL_SHEET_NAME = options.modelSheetName;
  
  // Настраиваем диапазоны модели
  if (options.modelRanges) MODEL_SMETA_RANGES = options.modelRanges;
  
  // Настраиваем блоки
  let blocksChanged = false;
  
  if (options.numBlocks !== undefined) {
    NUM_BLOCKS = options.numBlocks;
    blocksChanged = true;
  }
  
  if (options.blockNameTemplate) {
    BLOCK_NAME_TEMPLATE = options.blockNameTemplate;
    blocksChanged = true;
  }
  
  if (options.blockRanges) {
    DEFAULT_BLOCK_RANGES = options.blockRanges;
    blocksChanged = true;
  }
  
  if (options.blockCells) {
    DEFAULT_BLOCK_CELLS = options.blockCells;
    blocksChanged = true;
  }
  
  // Если что-то из настроек блоков изменилось, переинициализируем их
  if (blocksChanged) {
    initializeBlocks();
  }
  
  // Если предоставлены собственные блоки, используем их
  if (options.customBlocks) {
    BLOCKS = options.customBlocks;
  }
  
  // Возвращаем текущую конфигурацию
  return {
    storageSheetName: STORAGE_SHEET_NAME,
    registerSheetName: REGISTER_SHEET_NAME,
    modelSheetName: MODEL_SHEET_NAME,
    modelRanges: MODEL_SMETA_RANGES,
    blocks: BLOCKS
  };
}

/**
 * Создаёт пользовательское меню со специальными командами
 * @param {Object} [options] - Настройки меню
 * @param {string} [options.menuName] - Название меню
 * @param {string} [options.saveLabel] - Текст пункта для сохранения
 * @param {string} [options.restoreLabel] - Текст пункта для восстановления
 */
function createMenu(options = {}) {
  const menuName = options.menuName || "Отпечаток КП";
  const saveLabel = options.saveLabel || "Сохранить данные";
  const restoreLabel = options.restoreLabel || "Открыть данные";
  
  SpreadsheetApp.getUi()
    .createMenu(menuName)
    .addItem(saveLabel, "CalcMProLib.saveDataWithComment")
    .addItem(restoreLabel, "CalcMProLib.showRestoreDialog")
    .addToUi();
}

/**
 * Запрашивает у пользователя комментарий для сохраняемой записи
 * и запускает процесс сохранения с этим комментарием
 */
function saveDataWithComment() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.prompt(
    "Сохранение данных",
    "Введите комментарий для этой записи:",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Сохранение отменено.");
    return;
  }
  
  const comment = result.getResponseText();

  try {
    const newId = saveSnapshot(comment);
    ui.alert(`Данные успешно сохранены. Номер записи: ${newId}`);
  } catch (e) {
    ui.alert("Ошибка при сохранении: " + e.message);
  }
}

/**
 * Основная функция сохранения данных в хранилище
 * @param {string} comment - Комментарий к сохраняемой записи
 * @return {number} ID новой записи
 */
function saveSnapshot(comment) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let storageSheet = ss.getSheetByName(STORAGE_SHEET_NAME);
  if (!storageSheet) storageSheet = ss.insertSheet(STORAGE_SHEET_NAME);

  let registerSheet = ss.getSheetByName(REGISTER_SHEET_NAME);
  if (!registerSheet) {
    registerSheet = ss.insertSheet(REGISTER_SHEET_NAME);
    registerSheet.appendRow(["ID", "Дата", "Комментарий"]);
  }

  const newId = storageSheet.getLastRow() + 1;

  // Используем переменную MODEL_SHEET_NAME вместо строки
  const modelSheet = ss.getSheetByName(MODEL_SHEET_NAME);
  if (!modelSheet) throw new Error(`Лист '${MODEL_SHEET_NAME}' не найден`);

  let valuesFlat = [];

  MODEL_SMETA_RANGES.forEach(rangeA1 => {
    const range = modelSheet.getRange(rangeA1);
    const vals = range.getValues();
    vals.forEach(row => row.forEach(val => valuesFlat.push(val)));
  });

  BLOCKS.forEach(block => {
    const sheet = ss.getSheetByName(block.sheetName);
    if (!sheet) throw new Error(`Лист '${block.sheetName}' не найден`);
    
    block.ranges.forEach(r => {
      const vals = sheet.getRange(r).getValues();
      vals.forEach(row => row.forEach(val => valuesFlat.push(val)));
    });
    
    block.singleCells.forEach(cellA1 => {
      const val = sheet.getRange(cellA1).getValue();
      valuesFlat.push(val);
    });
  });

  const rowToWrite = [newId, ...valuesFlat];
  storageSheet.appendRow(rowToWrite);

  const nowStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd HH:mm:ss");
  registerSheet.appendRow([newId, nowStr, comment]);

  return newId;
}

/**
 * Отображает диалоговое окно со списком сохранённых записей для восстановления
 * @param {Object} [options] - Настройки диалога
 * @param {string} [options.dialogTitle] - Заголовок диалогового окна
 * @param {string} [options.headerText] - Текст заголовка в диалоге
 * @param {string} [options.cancelButtonText] - Текст кнопки отмены
 * @param {string} [options.restoreButtonText] - Текст кнопки восстановления
 * @param {string} [options.commentLabel] - Текст метки комментария
 */
function showRestoreDialog(options = {}) {
  const dialogTitle = options.dialogTitle || "Открытие данных";
  const headerText = options.headerText || "Выберите запись";
  const cancelButtonText = options.cancelButtonText || "Отмена";
  const restoreButtonText = options.restoreButtonText || "Открыть";
  const commentLabel = options.commentLabel || "Комментарий:";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const registerSheet = ss.getSheetByName(REGISTER_SHEET_NAME);
  if (!registerSheet) {
    SpreadsheetApp.getUi().alert(`Лист '${REGISTER_SHEET_NAME}' не найден. Нет сохранённых записей.`);
    return;
  }
  
  const dataRange = registerSheet.getDataRange();
  const data = dataRange.getValues();

  // Переименовываем переменную с options на recordOptions чтобы избежать конфликта
  let recordOptions = [];
  for (let i = 1; i < data.length; i++) {
    const [id, dateStr, comment] = data[i];
    
    let formattedDate = "";
    try {
      const dateObj = new Date(dateStr);
      formattedDate = Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone(), "dd.MM.yyyy HH:mm");
    } catch (e) {
      formattedDate = dateStr;
    }
    
    recordOptions.push({ 
      id,
      dateStr: formattedDate,
      comment
    });
  }
  
  if (recordOptions.length === 0) {
    SpreadsheetApp.getUi().alert("Нет сохранённых записей");
    return;
  }

  const html = HtmlService.createHtmlOutput(`
    <style>
      /* Базовые стили для всего документа */
      body { 
        font-family: 'Google Sans', 'Roboto', Arial, sans-serif; 
        padding: 16px; 
        margin: 0;
        color: #202124;
        background-color: #ffffff;
      }
      
      /* Стили для заголовка */
      h3 {
        margin-top: 0;
        font-weight: 500;
        margin-bottom: 16px;
        color: #202124;
        font-size: 16px;
        text-align: center;
      }
      
      /* Стили для текста подсказки */
      .header-text {
        color: #4285f4;
        font-size: 13px;
        text-align: center;
        margin-bottom: 16px;
      }
      
      /* Стили для выпадающего списка */
      select {
        width: 100%;
        font-family: 'Google Sans', 'Roboto', Arial, sans-serif;
        font-size: 14px;
        padding: 8px;
        border: 1px solid #dadce0;
        border-radius: 4px;
        margin-bottom: 16px;
        background-color: #ffffff;
        appearance: none;
        /* Добавляем стрелку выпадающего списка через SVG */
        background-image: url('data:image/svg+xml;utf8,<svg fill="%23333" height="24" viewBox="0 0 24 24" width="24" xmlns="http://www.w3.org/2000/svg"><path d="M7 10l5 5 5-5z"/></svg>');
        background-repeat: no-repeat;
        background-position-x: 98%;
        background-position-y: 50%;
      }
      
      /* Контейнер для кнопок */
      .button-container {
        display: flex;
        justify-content: space-between;
        width: 100%;
        margin-top: 16px;
      }
      
      /* Базовые стили для кнопок */
      button {
        font-family: 'Google Sans', 'Roboto', Arial, sans-serif;
        background-color: transparent;
        color:hsl(0, 7.20%, 13.50%);
        border: 1px solid #dadce0;
        border-radius: 4px;
        padding: 8px 16px;
        font-size: 14px;
        cursor: pointer;
        flex: 1;
        max-width: 48%;
        transition: all 0.2s ease;
        font-weight: 500;
      }
      
      /* Эффект при наведении на кнопки */
      button:hover {
        background-color: rgba(66, 133, 244, 0.04);
      }
      
      /* Стили для основной кнопки */
      button.primary {
        color: #4285f4;
        background-color: white;
        border: 1px solid #4285f4;
      }
      
      /* Эффект при наведении на основную кнопку */
      button.primary:hover {
        background-color: #4285f4;
        color: white;
      }
      
      /* Стили для второстепенной кнопки */
      button.secondary {
        color: #202124;
        background-color: white;
        border: 1px solid #dadce0;
      }
      
      /* Эффект при наведении на второстепенную кнопку */
      button.secondary:hover {
        background-color: #5f6368;  /* Светло-серый фон при наведении */
        color: white;
      }
      
      /* Стили для блока информации о записи */
      .record-info {
        display: flex;
        flex-direction: column;
        margin-top: 0;
        background-color: #f8f9fa;
        padding: 16px;
        border-radius: 4px;
        font-size: 14px;
        border: 1px solid #dadce0;
      }
      
      /* Стили для подписи к полю комментария */
      .label {
        font-weight: 400;
        margin-bottom: 8px;
        color: #5f6368;
        font-size: 13px;
      }
      
      /* Стили для текста комментария */
      .comment {
        white-space: pre-wrap;
        word-break: break-word;
        margin-bottom: 0;
        line-height: 1.4;
        min-height: 40px;
        color: #202124;
        background-color: white;
        padding: 8px;
        border-radius: 4px;
        border: 1px solid #dadce0;
      }
    </style>
  
    <!-- Заголовок диалога -->
    <div class="header-text">${headerText}</div>
    
    <!-- Выпадающий список с сохраненными записями -->
    <select id="recordSelect" onchange="showDetails()">
      ${recordOptions.map(opt => `<option value="${opt.id}" 
          data-date="${opt.dateStr}" 
          data-comment="${typeof opt.comment === 'string' ? opt.comment.replace(/"/g, '&quot;') : (opt.comment || '')}"
        >№${opt.id}, ${opt.dateStr}</option>`).join("")}
    </select>
    
    <!-- Блок с информацией о выбранной записи -->
    <div id="recordDetails" class="record-info">
      <div class="label">${commentLabel}</div>
      <div id="commentText" class="comment"></div>
    </div>
    
    <!-- Кнопки действий -->
    <div class="button-container">
      <button class="secondary" onclick="google.script.host.close()">${cancelButtonText}</button>
      <button class="primary" onclick="confirmRestore()">${restoreButtonText}</button>
    </div>
    
    <script>
      function confirmRestore() {
        const sel = document.getElementById('recordSelect').value;
        google.script.run.withSuccessHandler(function() {
          google.script.host.close();
        }).restoreByIdWithConfirm(parseInt(sel));
      }
      
      function showDetails() {
        const select = document.getElementById('recordSelect');
        const option = select.options[select.selectedIndex];
        document.getElementById('commentText').textContent = option.dataset.comment || '';
      }
      
      // Показываем информацию о первой записи при загрузке
      document.addEventListener('DOMContentLoaded', showDetails);
    </script>
  `).setWidth(360).setHeight(350);
  
  SpreadsheetApp.getUi().showModalDialog(html, dialogTitle);
}

/**
 * Функция восстановления данных по ID с предварительным подтверждением
 * @param {number} id - ID записи для восстановления
 * @param {Object} [options] - Настройки
 * @param {string} [options.confirmTitle] - Заголовок подтверждения
 * @param {string} [options.confirmMessage] - Сообщение подтверждения
 * @param {string} [options.successMessage] - Сообщение об успехе
 */
function restoreByIdWithConfirm(id, options = {}) {
  const confirmTitle = options.confirmTitle || 'Открытие данных';
  const confirmMessage = options.confirmMessage || 'Текущие данные будут заменены. Продолжить?';
  const successMessage = options.successMessage || 'Данные успешно открыты.';
  
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    confirmTitle,
    confirmMessage,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Открытие отменено.');
    return;
  }

  try {
    restoreSnapshotById(id);
    ui.alert(successMessage);
  } catch (e) {
    ui.alert('Ошибка при открытии: ' + e.message);
  }
}

/**
 * Основная функция восстановления данных из сохраненной записи
 * @param {number} id - ID записи для восстановления
 */
function restoreSnapshotById(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const storageSheet = ss.getSheetByName(STORAGE_SHEET_NAME);
  if (!storageSheet) throw new Error(`Лист '${STORAGE_SHEET_NAME}' не найден`);

  const data = storageSheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === id) {
      rowIndex = i;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Запись с ID=" + id + " не найдена");

  const row = data[rowIndex];
  const values = row.slice(1);
  const ssTZ = ss.getSpreadsheetTimeZone();

  // Используем переменную MODEL_SHEET_NAME вместо строки
  let pos = 0;
  const modelSheet = ss.getSheetByName(MODEL_SHEET_NAME);
  if (!modelSheet) throw new Error(`Лист '${MODEL_SHEET_NAME}' не найден`);

  MODEL_SMETA_RANGES.forEach(rangeA1 => {
    const range = modelSheet.getRange(rangeA1);
    const numCells = range.getNumRows() * range.getNumColumns();
    const valsSlice = values.slice(pos, pos + numCells);

    let values2D = [];
    for (let r = 0; r < range.getNumRows(); r++) {
      values2D[r] = [];
      for (let c = 0; c < range.getNumColumns(); c++) {
        values2D[r][c] = valsSlice[r * range.getNumColumns() + c];
      }
    }
    
    range.setValues(values2D);
    pos += numCells;
  });

  BLOCKS.forEach(block => {
    const sheet = ss.getSheetByName(block.sheetName);
    if (!sheet) throw new Error(`Лист '${block.sheetName}' не найден`);

    block.ranges.forEach(r => {
      sheet.getRange(r).clearContent();
    });
    
    block.singleCells.forEach(cellA1 => {
      sheet.getRange(cellA1).clearContent();
    });

    block.ranges.forEach(r => {
      const range = sheet.getRange(r);
      const numCells = range.getNumRows() * range.getNumColumns();
      const valsSlice = values.slice(pos, pos + numCells);

      let values2D = [];
      for (let rr = 0; rr < range.getNumRows(); rr++) {
        values2D[rr] = [];
        for (let cc = 0; cc < range.getNumColumns(); cc++) {
          values2D[rr][cc] = valsSlice[rr * range.getNumColumns() + cc];
        }
      }
      
      range.setValues(values2D);
      pos += numCells;
    });

    block.singleCells.forEach(cellA1 => {
      sheet.getRange(cellA1).setValue(values[pos]);
      pos++;
    });
  });
  
  // После восстановления данных запускаем функцию hideShow
  hideShow();
}

/**
 * Функция скрытия/показа строк на основе значений в колонке A
 * Скрывает строки со значением меньше 1 и показывает строки со значением больше 0.5
 */
function hideShow() {
  var startRow = 16; // С 16 строки работает скрипт
  var colToCheck = 1; // Column A (индекс 1)
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('СМЕТА');
  
  if (!sh) {
    console.log('Лист СМЕТА не найден');
    return;
  }
  
  var lastRow = sh.getLastRow();
  if (lastRow < startRow) {
    return; // Нет данных для обработки
  }
  
  var rg = sh.getRange(startRow, colToCheck, lastRow - startRow + 1, 1);
  var vA = rg.getValues();
  
  for (var i = 0; i < vA.length; i++) {
    if (vA[i][0] < 1) {
      sh.hideRows(i + startRow);
    }
    if (vA[i][0] > 0.5) {
      sh.showRows(i + startRow);
    }
  }
}

/**
 * В Google Apps Script все функции верхнего уровня автоматически экспортируются
 * и доступны при подключении скрипта как библиотеки.
 * 
 * Доступные внешне функции библиотеки:
 * - init(options): настройка библиотеки
 * - createMenu(options): создание меню
 * - saveDataWithComment(): сохранение данных с комментарием
 * - showRestoreDialog(options): диалог для восстановления
 * - restoreByIdWithConfirm(id, options): восстановление с подтверждением
 * - hideShow(): скрытие/отображение строк на основе значений
 */

/**
 * Инициализирует объект ReestrOffersLib
 * @param {Object} exports - Объект для экспорта функций
 */
function ReestrOffersLibInit(exports) {
  exports.saveDataWithComment = saveDataWithComment;
  exports.showRestoreDialog = showRestoreDialog;
  exports.restoreByIdWithConfirm = restoreByIdWithConfirm;
  exports.init = init;
  
  return exports;
}
