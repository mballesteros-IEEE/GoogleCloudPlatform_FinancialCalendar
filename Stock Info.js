function updateStocksInfoAndCreateEventsAndSendMails(){
  updateStocksInfo();
  createEventsAndSendMails();
  
  PropertiesService.getUserProperties().setProperty('LastNotificationTime', new Date());
  setUpdateTime();
}

function updateStocksInfo(){
  ///////////////////////////////////////////////////////////////////////////////////
  // CONFIGURATION
  var columnCompanyName = getColumnCompanyName();
  var columnIndexResults = getColumnResultsDate();
  var columnIndexDate = getColumnPaidDate();
  var columnIndexExDate = getColumnExDate();
  var columnIndexAmount = getColumnAmount();
  var columnIndexCurrency = getColumnCurrency();
  var columnIndexCurrencyRate = getColumnCurrencyRate();
  
  var columnIndexMorningstarId = getColumnMorningstarId();
  var columnIndexInvestingId = getColumnInvestingId();
  
  ///////////////////////////////////////////////////////////////////////////////////
    
  var calendarSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendario");
  var lastRow = calendarSheet.getLastRow();
  var range = calendarSheet.getRange("A1:Z" + lastRow + "");
  var values = range.getValues();
  
  var today = getTodayDate();
  var oneYearAgo = new Date(today);
  oneYearAgo.setYear(today.getYear() - 1);
  
  var documentProperties = PropertiesService.getUserProperties();
  
  for (x=1; x < values.length; x++) {
    try {
      var row = values[x];
      
      var morningstarId = row[columnIndexMorningstarId - 1].trim();
      var investingId = row[columnIndexInvestingId - 1].trim();
      
      if(morningstarId == "" && investingId == "")
        continue;
      
      var companyName = row[columnCompanyName - 1].trim();

      var message = "Cargando...";
      range.getCell(x + 1, columnIndexDate).setValue(message);
      range.getCell(x + 1, columnIndexExDate).setValue(message);
      range.getCell(x + 1, columnIndexAmount).setValue(message);
      range.getCell(x + 1, columnIndexResults).setValue(message);
      
      SpreadsheetApp.flush();
      
      ///////////////////////////////////////////////////////////////////////////////////
      // GET DIVIDEND INFO
      var dividendInfo = {latestAmount: 0, paidDate: new Date(), exDivDate: new Date()};      
      var isDividendInfoOk = setDividendData(morningstarId, dividendInfo);
            
      if(isDividendInfoOk && dividendInfo && dividendInfo.paidDate > oneYearAgo) {
        range.getCell(x + 1, columnIndexDate).setValue(dividendInfo.paidDate);
        range.getCell(x + 1, columnIndexExDate).setValue(dividendInfo.exDivDate);
        range.getCell(x + 1, columnIndexAmount).setValue(dividendInfo.latestAmount);
        
        var currency = row[columnIndexCurrency - 1];
        var rate = getRate(dividendInfo.paidDate, currency);
        range.getCell(x + 1, columnIndexCurrencyRate).setValue(rate);
        
        if (dividendInfo.paidDate <= today){
          var shares = row[getColumnShares() - 1];
          var country = row[getColumnCountry() - 1];
          documentProperties.setProperty("history#" + morningstarId + "#" + dividendInfo.paidDate, companyName + "#" + dividendInfo.latestAmount + "#" + currency + "#" + shares + "#" + country);  
        }
      }
      else {
        range.getCell(x + 1, columnIndexDate).setValue("");
        range.getCell(x + 1, columnIndexExDate).setValue("");
        range.getCell(x + 1, columnIndexAmount).setValue("");
        range.getCell(x + 1, columnIndexCurrencyRate).setValue("");
      }
      ///////////////////////////////////////////////////////////////////////////////////
      
      ///////////////////////////////////////////////////////////////////////////////////
      // GET RESULTS INFO
      var newResultsDate = getResultsData(investingId);
      if (newResultsDate && newResultsDate > oneYearAgo) {
        range.getCell(x + 1, columnIndexResults).setValue(newResultsDate);
      }
      else {
        range.getCell(x + 1, columnIndexResults).setValue("");
      }
      
      SpreadsheetApp.flush();      
    } catch (e) {
      console.error('updateStocksInfo() yielded an error: ' + e);
      
      try {
        SpreadsheetApp.getUi().alert('Se ha producido un error inesperado: ' + e);
      }
      catch(e){
      }
    }
  }
  
  setDividendHistory();
}

function createEventsAndSendMails(){
  ///////////////////////////////////////////////////////////////////////////////////
  // CONFIGURATION
  var columnIndexResults = getColumnResultsDate();
  var columnIndexPaidDate = getColumnPaidDate();
  var columnIndexExDate = getColumnExDate();
  
  var columnIndexMorningstarId = getColumnMorningstarId();
  var columnIndexInvestingId = getColumnInvestingId();
  
  var createPaidDividendEvent = getIsActivePaidDate();
  var createExDividendEvent = getIsActiveExDividendDate();
  
  var sendEmailWithChanges = getIsActiveNotificationChange();
  var sendEmailWithTodayEvents = getIsActiveNotificationToday();
  var sendTodayNotifications = getSendTodayNotifications();
  
  ///////////////////////////////////////////////////////////////////////////////////
  // CALENDARS
  var exDividendsEventCal = CalendarApp.getCalendarById(getCalendarExDividends());
  if(createExDividendEvent && exDividendsEventCal == null){
    SpreadsheetApp.getUi().alert('El Id de calendario para fechas ex dividendo no es válido. Deberás configurarlo en la hoja "Configuración".');
    return;
  }
  
  var dividendsEventCal = CalendarApp.getCalendarById(getCalendarIdDividends());
  if(createPaidDividendEvent && dividendsEventCal == null){
    SpreadsheetApp.getUi().alert('El Id de calendario para pagos de dividendo no es válido. Deberás configurarlo en la hoja "Configuración".');
    return;
  }
  
  var createResultsEvent = getIsActiveResultsDate();
  var resultsEventCal = CalendarApp.getCalendarById(getCalendarIdResults());
  if(createResultsEvent && resultsEventCal == null){
    SpreadsheetApp.getUi().alert('El Id de calendario para resultados no es válido. Deberás configurarlo en la hoja "Configuración".');
    return;
  }
  
  deleteAllEvents(exDividendsEventCal);
  deleteAllEvents(dividendsEventCal);
  deleteAllEvents(resultsEventCal);
  ///////////////////////////////////////////////////////////////////////////////////
    
  var documentProperties = PropertiesService.getUserProperties();
  //documentProperties.deleteAllProperties();
  
  var mail = getEmail();
      
  SpreadsheetApp.flush();
  
  var calendarSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Calendario");
  var lastRow = calendarSheet.getLastRow();
  var range = calendarSheet.getRange("A1:Z" + lastRow + "");
      
  SpreadsheetApp.flush();
  
  var values = range.getValues();
  
  var today = getTodayDate();
  var yesterday = getYesterdayDate();
  var tomorrow = getTomorrowDate();
  
  for (x=1; x < values.length; x++) {
    try {
      var row = values[x];
      
      var morningstarId = row[columnIndexMorningstarId - 1].trim();
      var investingId = row[columnIndexInvestingId - 1].trim();
      
      if (morningstarId == "" && investingId == "")
        continue;
      
      var paidDate = row[columnIndexPaidDate - 1];
      var exDivDate = row[columnIndexExDate - 1];
      var resultsDate = row[columnIndexResults - 1];
      
      ///////////////////////////////////////////////////////////////////////////////////
      // CREATE EVENTS
      
      var summary = calendarSheet.getRange(x + 1, getColumnSummary()).getValue();
      var description = calendarSheet.getRange(x + 1, getColumnDescription()).getValue() + getFooter();
      
      var options = {
        'description': description.replace(/\n/g, '\n'),
      }
      
      if(paidDate != "" && createPaidDividendEvent && paidDate && paidDate >= yesterday) {
        dividendsEventCal.createAllDayEvent(summary + ' Dividendo', paidDate, options);
      }

      if(exDivDate != "" && createExDividendEvent && exDivDate && exDivDate >= yesterday) {
        exDividendsEventCal.createAllDayEvent(summary + ' Ex-dividendo', exDivDate, options);
      }
      
      if(resultsDate != "" && createResultsEvent && resultsDate && resultsDate >= yesterday) {
        resultsEventCal.createAllDayEvent(summary + ' Resultados', resultsDate, options);
      }
      ///////////////////////////////////////////////////////////////////////////////////
      
      ///////////////////////////////////////////////////////////////////////////////////
      // SEND MAILS
      var oldDataDividend = documentProperties.getProperty(morningstarId);
      oldDataDividend = new Date(oldDataDividend);
      if (sendEmailWithChanges && paidDate && oldDataDividend < paidDate && paidDate >= today) {
        var dateFormatted = Utilities.formatDate(paidDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        MailApp.sendEmail(mail, '🆕 ' + summary + ' Dividendo ' + dateFormatted, description);
        
        documentProperties.setProperty(morningstarId, paidDate);
      }

      var oldDateResults = documentProperties.getProperty(investingId);
      oldDateResults = new Date(oldDateResults);
      if (sendEmailWithChanges && oldDateResults < resultsDate && resultsDate >= today) {
        var dateFormatted = Utilities.formatDate(resultsDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        MailApp.sendEmail(mail, '🆕 ' + summary + ' Resultados ' + dateFormatted, description);
        
        documentProperties.setProperty(investingId, resultsDate);
      }
      
      if(sendEmailWithTodayEvents && sendTodayNotifications){
        if(exDivDate 
           && exDivDate
           && exDivDate.getDate() == tomorrow.getDate()
           && exDivDate.getMonth() == tomorrow.getMonth() 
           && exDivDate.getFullYear() == tomorrow.getFullYear()) {
          MailApp.sendEmail(mail, '📝 ' + summary + ' Mañana: Fecha ex dividendo', description);
        }
        
        if(paidDate
           && paidDate
           && paidDate.getDate() == today.getDate() 
           && paidDate.getMonth() == today.getMonth() 
           && paidDate.getFullYear() == today.getFullYear()) {
          MailApp.sendEmail(mail, '💸 ' + summary + ' Hoy: Pago de dividendo', description);
        }
        
        if(resultsDate
           && resultsDate.getDate() == today.getDate()
           && resultsDate.getMonth() == today.getMonth() 
           && resultsDate.getFullYear() == today.getFullYear()) {
          MailApp.sendEmail(mail, '📊 ' + summary + ' Hoy: Publicación de resultados', description);
        }
      }
      ///////////////////////////////////////////////////////////////////////////////////
      
    } catch (e) {
      console.error('createEventsAndSendMails() yielded an error: ' + e);
      
      try {
        SpreadsheetApp.getUi().alert('Se ha producido un error inesperado: ' + e);
      }
      catch(e){
      }
    }
  }
}

function getDividendEvents(){
  ///////////////////////////////////////////////////////////////////////////////////
  // CALENDARS  
  var dividendsEventCal = CalendarApp.getCalendarById(getCalendarIdDividends());
  if(createPaidDividendEvent && dividendsEventCal == null){
    SpreadsheetApp.getUi().alert('El Id de calendario para pagos de dividendo no es válido. Deberás configurarlo en la hoja "Configuración".');
    return;
  }
  ///////////////////////////////////////////////////////////////////////////////////
  
  var dividendsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Histórico Dividendos");
  dividendsSheet.clearContents();
  
  if (dividendsEventCal == null)
    return;
  
  var events = dividendsEventCal.getEvents(new Date(1900, 1, 1), getTodayDate());
  
  for (x=0; x<events.length; x++) {
    try {
      var event = events[x];
      
      if(event.getDescription().indexOf("herr")>-1) {
        event.deleteEvent();
      }
    } catch (e) {
      console.error('deleteAllEvents() yielded an error: ' + e);
    } 
  }
  
}

function deleteAllEvents(eventCal)
{
  if (eventCal == null)
    return;
  
  var events = eventCal.getEvents(getYesterdayDate(), new Date(2500, 1, 1));
  
  for (x=0; x<events.length; x++) {
    try {
      var event = events[x];
      
      if(event.getDescription().indexOf("herr")>-1) {
        event.deleteEvent();
      }
    } catch (e) {
      console.error('deleteAllEvents() yielded an error: ' + e);
    } 
  }
}

function setDividendData(idMorningstar, dividendInfo) {
  try{
    if (idMorningstar && idMorningstar != "") {
      var fecha, dato;
      var url = Utilities.formatString('http://tools.morningstar.es/es/stockreport/default.aspx?SecurityToken=%s%5D3%5D0%5DE0WWE%24%24ALL', idMorningstar);
      var html = UrlFetchApp.fetch(url, {muteHttpExceptions: true}).getContentText();
      
      dividendInfo.paidDate = getDateBySeparator(html, '<td class="date colLatest" headers="MsStockReportOdD2">', 3, "/");
      
      if(dividendInfo.paidDate && dividendInfo.paidDate != undefined){
        dividendInfo.latestAmount = getFloat(html, '<td class="number colLatest" headers="MsStockReportOdD2">', 1);
        dividendInfo.exDivDate = getDateBySeparator(html, '<td class="date colLatest" headers="MsStockReportOdD2">', 2,"/");    
      }
      else{
        dividendInfo.paidDate = getDateBySeparator(html, '<td class="date colLatest" headers="MsStockReportOdhD2">',3, "/");
        
        if(dividendInfo.paidDate == undefined)
          return false;
        
        dividendInfo.latestAmount = getFloat(html, '<td class="number colLatest" headers="MsStockReportOdhD2">', 1) / 100;
        dividendInfo.exDivDate = getDateBySeparator(html, '<td class="date colLatest" headers="MsStockReportOdhD2">', 2, "/");
      }
      
      return true;
    }
    else {
      return false;
    }
  } catch (e) {
    console.error('setDividendData() yielded an error: ' + e);    
    
    return false;
  }
}

function getResultsData(investingId) {
  try{
    if(investingId && investingId != ""){
      var url = Utilities.formatString('https://es.investing.com/equities/%s', investingId);
      var html = UrlFetchApp.fetch(url, {muteHttpExceptions: true}).getContentText();
      
      var href = Utilities.formatString("<a href='/equities/%s-earnings'>", investingId);
      var date = getDateBySeparator(html, href, 1, ".");
      
      return date; 
    }
    else {
      return false;
    }
  } catch (e) {
    console.error('setDividendData() yielded an error: ' + e);
    
    return false;
  }
}

function getRate(date, quote){
  if(!quote || quote == "" || quote == "EUR")
    return 1.0;
  
  var dateFormatted = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var rateString = '';
  var today = getTodayDate();
  
  if (date < today) {
    var url = 'https://api.exchangeratesapi.io/history?start_at=' + dateFormatted + '&end_at=' + dateFormatted + '&symbols=' + quote;
    var plain = UrlFetchApp.fetch(url);
    var json = JSON.parse(plain);
    
    if (json.rates[dateFormatted] == undefined){
      date.setDate(date.getDate() - 1);
      return getRate(date, quote);
    }
    else {
      rateString = json.rates[dateFormatted][quote];      
    }    
  }
  else {
    var url = 'https://api.exchangeratesapi.io/latest?symbols=' + quote;
    var plain = UrlFetchApp.fetch(url);
    var json = JSON.parse(plain);
    rateString = json.rates[quote];
  }
  
  var rate = parseFloat(rateString);
  
  return rate;
}

function testDividend(){  
  var dividends = {latestAmount: 0, paidDate: new Date(), exDivDate: new Date()};
  
  getDividendData("", dividends);
  
  var latestAmount = dividends.latestAmount;
  var x = "";
}

function testResultsDate(){
  var documentProperties = PropertiesService.getUserProperties();
  documentProperties.deleteProperty('LastNotificationTime');
}

function getSendTodayNotifications() {
  var documentProperties = PropertiesService.getUserProperties();
  var keys = documentProperties.getKeys();  
  var lastNotificationProperty = undefined;
  if(keys.includes('LastNotificationTime'))
    lastNotificationProperty = documentProperties.getProperty('LastNotificationTime');
   
  var sendNotification = false;
  if(lastNotificationProperty) {
    var lastNotificationDateTime = new Date(lastNotificationProperty);
    var today = new getTodayDate();
    
    if(lastNotificationDateTime < today){
      sendNotification = true;
    }
  }
  else {
    sendNotification = true;
  }
  
  return sendNotification;
}

function getFooter(){
 return "\n\n" + "www.manuelballesteros.eu"; 
}

function setDividendHistory(){
  var documentProperties = PropertiesService.getUserProperties();
  var properties = documentProperties.getProperties();
  
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var historySheet = spreadSheet.getSheetByName("Histórico dividendos");
  
  var row = 1;
  var values = [];
  for (var key in properties) {
    if(key.startsWith("history")){
      var keySplitted = key.split("#");
      var paidDate = new Date(keySplitted[2]);
      
      var valueSplitted = properties[key].split("#");
      var companyName = valueSplitted[0];
      var amount = valueSplitted[1].replace(".",",");
      var currency = valueSplitted[2];
      var shares = valueSplitted[3];
      var country = valueSplitted[4];
      
      values.push([paidDate, companyName, amount, currency, shares, country]);
    }
  }
  
  values.sort(function(a, b) {
    a = a[0];
    b = b[0];
    return a>b ? -1 : a<b ? 1 : 0;
  });

  var rows = historySheet.getMaxRows();
  var rowStart = 25;
  historySheet.getRange(rowStart, 1, rows - rowStart + 1, values[0].length).clearContent();
  historySheet.getRange(rowStart, 1, values.length, values[0].length).setValues(values);
}

function deleteDividendHistory(){
  var documentProperties = PropertiesService.getUserProperties();
  var properties = documentProperties.getProperties();
  
  var row = 1;
  var values = [];
  for (var key in properties) {
    if(key.startsWith("history")){
      documentProperties.deleteProperty(key);
    }
  }
}