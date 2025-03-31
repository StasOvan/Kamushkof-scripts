// var userKey = Session.getTemporaryActiveUserKey();
// var sheetName = 'Расчет (User_' + userKey + ')';
// Logger.log(userKey);

const MAIN_SHEET = "Основная";
const TEMPLATE_SHEET = "Таблица просчета";
const SS = SpreadsheetApp.getActiveSpreadsheet();

function onOpenTable() {
  var userEmail, userName;
  var mainSheet = SS.getSheetByName(MAIN_SHEET);

  // Пытаемся получить email пользователя с обработкой ошибок
  try {  
    userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      Browser.msgBox("Не найден Email. Выполните вход в аккаунт Google.");
      mainSheet.showSheet();
      SS.setActiveSheet(mainSheet);
      mainSheet.activate();  
      hideAllSheetsByExcept(MAIN_SHEET);
      return;
    }
    if (!userEmail) throw new Error("Email не доступен");
    userName = userEmail.split('@')[0];
  } catch (e) {
    //Browser.msgBox("Не удалось получить email: " + e.message);
    //return;
  }
  
  var sheetName = 'Расчет (' + userName.replace(/[\\\/\?\*\[\]]/g, '_') + ')';

  // Проверяем существование листа
  var targetSheet = SS.getSheetByName(sheetName);
  if (!targetSheet) {
    var templateSheet = SS.getSheetByName(TEMPLATE_SHEET);
    targetSheet = templateSheet.copyTo(SS);
    targetSheet.setName(sheetName);
    //  } else {
    //   // Если шаблон не найден, создаем пустой лист (резервный вариант)
    //   targetSheet = SS.insertSheet(sheetName);
    //   Browser.msgBox("Шаблон 'Таблица просчета' не найден. Создан пустой лист.");
    // }
  } 
  
  // Скрываем остальные листы
  hideAllSheetsByExcept(sheetName);
  
  targetSheet.activate();
  
}

function hideAllSheetsByExcept(sheetName) {

  var sheets = SS.getSheets();
  var targetSheet = SS.getSheetByName(sheetName);
  var mainSheet = SS.getSheetByName(MAIN_SHEET);
 
  // Если целевой лист не найден, используем MAIN_SHEET
  if (!targetSheet) {
    targetSheet = mainSheet;
  };
  
  sheets.forEach(sheet => {
    Logger.log(sheet.getName() + " : " + targetSheet.getName());
    Logger.log(sheet.getName() != targetSheet.getName());
    if (sheet.getName() != targetSheet.getName() && sheet.getName() != MAIN_SHEET) 
      sheet.hideSheet();
    else
      sheet.showSheet();
  });
  
  SS.setActiveSheet(targetSheet);

}
