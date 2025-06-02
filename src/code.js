/**
 * code.js - Точка входа в библиотеку CalcMProLib
 * 
 * Экспортирует все необходимые функции библиотеки
 * для использования в Google Sheets
 */

/**
 * Глобальные переменные для обмена данными между файлами
 */
var FormulasLib = {};
var ReestrOffersLib = {};
var UI = {};
var Blocks = {}; // Добавляем новый объект для blocks.js

// Инициализация библиотеки при запуске
function init() {
  // Передаем объекты между файлами
  FormulasLibInit(FormulasLib);
  ReestrOffersLibInit(ReestrOffersLib);
  UIInit(UI);
  blocksInit(Blocks); // Инициализируем blocks.js
  
  return {
    FormulasLib: FormulasLib,
    ReestrOffersLib: ReestrOffersLib,
    UI: UI,
    Blocks: Blocks
  };
}

/**
 * Функция выполняется при открытии таблицы
 */
function onOpen() {
  init(); // Инициализируем библиотеку
  UI.createMenu(); // Создаем меню с помощью UI
}

// Экспорт функций библиотеки для использования в Google Sheets
this.FLAG_CHECK = FormulasLib.FLAG_CHECK;
this.MY_CUSTOM_FUNC = FormulasLib.MY_CUSTOM_FUNC;
this.SUPPLIER_ARTICLE = FormulasLib.SUPPLIER_ARTICLE;
this.UNIT_MEASURE = FormulasLib.UNIT_MEASURE;
this.SUM_PRODUCT_ARR = FormulasLib.SUM_PRODUCT_ARR;
this.PHOTO_ARRAY = FormulasLib.PHOTO_ARRAY;
this.SPECIFICATION_WITH_LINK = FormulasLib.SPECIFICATION_WITH_LINK;
this.CODE_MS = FormulasLib.CODE_MS;
this.CHECK_QUANTITY = FormulasLib.CHECK_QUANTITY;
this.CALCULATE_PRICE_DYNAMIC = FormulasLib.CALCULATE_PRICE_DYNAMIC;
this.XBLOCK = Blocks.XBLOCK; // Используем XBLOCK из модуля Blocks

/**
 * Создает единое меню для доступа ко всем функциям библиотеки
 */
function createMenu() {
  const ui = SpreadsheetApp.getUi();
  
  // Создаем меню "Формулы МС"
  ui.createMenu('Формулы МС')  // Заменить 'CalcMProLib' на 'Формулы МС'
    .addItem('Сохранить данные', 'saveDataWithComment')
    .addItem('Открыть данные', 'showRestoreDialog')
    .addSeparator()
    .addItem('Настройки библиотеки', 'showConfigDialog')
    .addItem('О библиотеке', 'showInfoDialog')
    .addToUi();
}

/**
 * Инициализирует библиотеку с настройками по умолчанию
 * или с пользовательскими настройками
 * @param {Object} options - Пользовательские настройки
 */
function initLibrary(options = {}) {
  // Инициализируем компоненты библиотеки
  const formulaConfig = FormulasLib.init(options);
  const reestrConfig = ReestrOffersLib.init(options);
  
  return {
    formula: formulaConfig,
    reestr: reestrConfig,
    blocks: Blocks // Добавляем модуль blocks в возвращаемый объект
  };
}

// Экспортируем общие функции
this.onOpen = onOpen;
this.initLibrary = initLibrary;
this.createMenu = createMenu;

// Экспортируем функции для работы с блоками
this.processAllBlocks = Blocks.processAllBlocks;
this.processXBLOCK = Blocks.processXBLOCK;

// Функции, которые будут доступны через меню
function saveDataWithComment() {
  return ReestrOffersLib.saveDataWithComment();
}

function showRestoreDialog() {
  return ReestrOffersLib.showRestoreDialog();
}

function showConfigDialog() {
  return UI.showConfigDialog();
}

function showInfoDialog() {
  return UI.showInfoDialog();
}

// Экспортируются эти функции
this.saveDataWithComment = saveDataWithComment;
this.showRestoreDialog = showRestoreDialog;
this.showConfigDialog = showConfigDialog;
this.showInfoDialog = showInfoDialog;

// Экспорт функции для восстановления данных
this.restoreByIdWithConfirm = function(id) { return ReestrOffersLib.restoreByIdWithConfirm(id); };