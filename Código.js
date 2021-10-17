/////////////////////////////////////////////////////////////////
// MENU
function addMenu()
{
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('CALENDARIO')
      .addItem('Actualizar', 'createTriggersAndUpdateCalendar')
      .addItem('Borrar histórico de dividendos', 'deleteDividendHistory')
      .addSeparator()
//      .addItem('¿Un café 😄?', 'donate')
      .addToUi();
}

function onOpen()
{
  addMenu();
}

/////////////////////////////////////////////////////////////////
// TRIGGERS
function createTriggers() {
  deleteAllTriggers();
  createSpreadsheetCronTrigger();
}

function setUpdateTime(){
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Configuración").getRange(32, 2).setValue(new Date());
}

function deleteAllTriggers() {  
 // Deletes all triggers in the current project.
 var triggers = ScriptApp.getProjectTriggers();
  
 for (var i = 0; i < triggers.length; i++) {
   ScriptApp.deleteTrigger(triggers[i]);
 }
}

function createSpreadsheetCronTrigger() {
  // Trigger every day.
//  ScriptApp.newTrigger('updateStocksInfoAndCreateEventsAndSendMails')
//      .timeBased()
//      .everyHours(8)
//      //.everyMinutes(5)
//      .create();
  
  // Trigger every day.
  ScriptApp.newTrigger('updateStocksInfoAndCreateEventsAndSendMails')
      .timeBased()
      .atHour(4)
      .everyDays(1)
      .create();
  
  // Trigger every day.
  ScriptApp.newTrigger('updateStocksInfoAndCreateEventsAndSendMails')
      .timeBased()
      .atHour(16)
      .everyDays(1)
      .create();
  
  // Trigger every day.
  ScriptApp.newTrigger('getHolidays')
      .timeBased()
      .atHour(5)
      .everyDays(1)
      .create();
  
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('checkDonations')
      .forSpreadsheet(ss)
      .onOpen()
      .create();
}

function createTriggersAndUpdateCalendar(){
  createTriggers();
  updateStocksInfoAndCreateEventsAndSendMails();
  getHolidays();
}

/////////////////////////////////////////////////////////////////

function getConfigurationParameter(row) {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var calendarSheet = spreadSheet.getSheetByName("Configuración");
  var parameter = calendarSheet.getRange(row, 2).getValue();
  
  return parameter;
}

function getColumnSummary() {
  return getConfigurationParameter(2);
}

function getColumnDescription() {
  return getConfigurationParameter(3);
}

function getColumnResultsDate() {
  return getConfigurationParameter(4);
}

function getColumnPaidDate() {
  return getConfigurationParameter(5);
}

function getColumnExDate() {
  return getConfigurationParameter(6);
}

function getColumnAmount() {
  return getConfigurationParameter(7);
}

function getColumnMorningstarId() {
  return getConfigurationParameter(8);
}

function getColumnInvestingId() {
  return getConfigurationParameter(9);
}

function getColumnCurrency() {
  return getConfigurationParameter(10);
}

function getColumnCurrencyRate() {
  return getConfigurationParameter(11);
}

function getColumnCompanyName() {
  return getConfigurationParameter(12);
}

function getColumnShares() {
  return getConfigurationParameter(13);
}

function getColumnCountry() {
  return getConfigurationParameter(14);
}

function getCalendarExDividends() {
  return getConfigurationParameter(16).trim();
}

function getCalendarIdDividends() {
  return getConfigurationParameter(17).trim();
}

function getCalendarIdResults() {
  return getConfigurationParameter(18).trim();
}

function getCalendarIdHolidays() {
  return getConfigurationParameter(19).trim();
}

function getEmail() {
  return getConfigurationParameter(20).trim();;
}

function getIsActiveNotificationChange() {
  return getConfigurationParameter(21);
}

function getIsActiveNotificationToday() {
  return getConfigurationParameter(22);
}


function getIsActivePaidDate() {
  return getConfigurationParameter(24);
}

function getIsActiveExDividendDate() {
  return getConfigurationParameter(25);
}

function getIsActiveResultsDate() {
  return getConfigurationParameter(26);
}

function getIsActiveHolidaysDate() {
  return getConfigurationParameter(27);
}


function getTodayDate(){
 return new Date(new Date().getFullYear(),new Date().getMonth(), new Date().getDate());
}

function getYesterdayDate(){
 return new Date(new Date().getFullYear(),new Date().getMonth(), new Date().getDate() - 1);
}

function getTomorrowDate(){
 return new Date(new Date().getFullYear(),new Date().getMonth(), new Date().getDate() + 1);
}