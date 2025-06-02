/**
 * UserScript.gs - Скрипт для работы с библиотекой CalcMProLib
 * 
 * Этот файл нужно добавить в проект скрипта пользовательской таблицы
 * (Extensions > Apps Script в меню Google Sheets)
 */

/**
 * Функция выполняется при открытии таблицы
 */
function onOpen() {
  // Создаем меню
  CalcMProLib.createMenu();
}

/**
 * Функции-обертки для использования в таблице без префикса библиотеки
 */

function FLAG_CHECK(dRange, flagC28, valueE28, blockNum) {
  return CalcMProLib.FLAG_CHECK(dRange, flagC28, valueE28, blockNum);
}

function MY_CUSTOM_FUNC(a2, b27, aRange, blockNum) {
  return CalcMProLib.MY_CUSTOM_FUNC(a2, b27, aRange, blockNum);
}

function SUPPLIER_ARTICLE(dRange, blockNum) {
  return CalcMProLib.SUPPLIER_ARTICLE(dRange, blockNum);
}

function UNIT_MEASURE(dRange, blockNum) {
  return CalcMProLib.UNIT_MEASURE(dRange, blockNum);
}

function SUM_PRODUCT_ARR(a2, eRange, gRange, blockNum) {
  return CalcMProLib.SUM_PRODUCT_ARR(a2, eRange, gRange, blockNum);
}

function PHOTO_ARRAY(dRange, blockNum) {
  return CalcMProLib.PHOTO_ARRAY(dRange, blockNum);
}

function SPECIFICATION_WITH_LINK(dRange, blockNum) {
  return CalcMProLib.SPECIFICATION_WITH_LINK(dRange, blockNum);
}

function CODE_MS(dRange, blockNum) {
  return CalcMProLib.CODE_MS(dRange, blockNum);
}

function CHECK_QUANTITY(dRange, eRange, blockNum) {
  return CalcMProLib.CHECK_QUANTITY(dRange, eRange, blockNum);
}

function CALCULATE_PRICE_DYNAMIC(dRange, iRange, blockNum, markupName) {
  return CalcMProLib.CALCULATE_PRICE_DYNAMIC(dRange, iRange, blockNum, markupName);
}

function XBLOCK(index) {
  return CalcMProLib.XBLOCK(index);
}

/**
 * Функции для работы с реестром
 */
function saveDataWithComment() {
  return CalcMProLib.saveDataWithComment();
}

function showRestoreDialog() {
  return CalcMProLib.showRestoreDialog();
}

function restoreByIdWithConfirm(id) {
  return CalcMProLib.restoreByIdWithConfirm(id);
}