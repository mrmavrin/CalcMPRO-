// UI.js - Обработка пользовательского интерфейса для библиотеки CalcMProLib

/**
 * Создает меню для работы с формулами
 * @param {Object} [options] - Настройки меню
 */
function createMenu(options = {}) {
  const menuName = options.menuName || "Формулы МС";
  
  SpreadsheetApp.getUi()
    .createMenu(menuName)
    .addItem("Сохранить данные", "CalcMProLib.saveDataWithComment")
    .addItem("Открыть данные", "CalcMProLib.showRestoreDialog")  // Функция вызывается БЕЗ параметров
    .addSeparator()
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
 * Показывает диалоговое окно для настройки библиотеки
 */
function showConfigDialog() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutputFromFile('Config')
      .setWidth(400)
      .setHeight(300);
  ui.showModalDialog(html, 'Настройка библиотеки');
}

/**
 * Инициализирует объект UI
 * @param {Object} exports - Объект для экспорта функций
 */
function UIInit(exports) {
  exports.createMenu = createMenu;
  exports.showInfoDialog = showInfoDialog;
  exports.showConfigDialog = showConfigDialog;
  
  return exports;
}