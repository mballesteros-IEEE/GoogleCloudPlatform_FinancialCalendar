function getHolidays() {  
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var calendarSheet = spreadSheet.getSheetByName("Calendario Cierres");
  var configSheet = spreadSheet.getSheetByName("Configuración");
  
  // Get user countries
  var countriesArray = configSheet.getRange(2, 5, 20, 1).getValues();  
  var countries = [];
  for (x=0; x < countriesArray.length; x++) {
    countries[x] = countriesArray[x][0];
  }  
  if(countries.length == 0)
    return;
  
  var mail = getEmail();
  var today = getTodayDate();
  
  // Delete events
  var calendar = CalendarApp.getCalendarById(getCalendarIdHolidays());
  if (calendar != null){
    var events = calendar.getEvents(getYesterdayDate(), new Date(2500, 1, 1));
    
    for (x=0; x<events.length; x++) {
      try {
        var event = events[x];
        
        if(event.getDescription().indexOf("herr")>-1) {
          event.deleteEvent();
        }
      } catch (e) {
        console.error('deleteEvent() yielded an error: ' + e);
      } 
    } 
  }
  
  // Load holidays
  var range = calendarSheet.getRange(1, 2);
  range.clear();
  SpreadsheetApp.flush();
  range.setFormula('=IMPORTHTML("https://es.investing.com/holiday-calendar/";"TABLE";1)');
  SpreadsheetApp.flush();
  waitForLoading(range,calendarSheet,10);
  
  var calendarSheet = spreadSheet.getSheetByName("Calendario Cierres");
  var data = calendarSheet.getRange(3, 1, 300, 5).getValues();
  
  var sendEmail = getSendTodayHolidays();
  
  for (x=0; x < data.length; x++) {
    try {
      var row = data[x];
      
      var date = row[0];
      var country = row[2].trim();
      var market = row[3].trim();
      var holiday = row[4].trim();
      
      if(countries.includes(country)) {
        var dateFormatted = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
        var message = market + ' (' + country  + ') ' + 'no abrirá el día ' + dateFormatted + '\n\n' + holiday + "\n\nwww.herramientasdeinversor.com";
        
        // Create calendar
        if(calendar != null && getIsActiveHolidaysDate()){
          var options = {
            'description': message,
          }
          calendar.createAllDayEvent(market + ' (' + holiday + ')', date, options);          
        }
        
        // Send email
        if(date.valueOf() === today.valueOf() && sendEmail){
          var subject = '🔒 ' + country + ' - Cierre de mercado por festivo hoy';
          MailApp.sendEmail(mail, subject, message);
        }
      }
    }
    catch(e){
      console.error('getHolidays() yielded an error: ' + e);      
    }
  }
  
  PropertiesService.getUserProperties().setProperty('LastHolidayTime', new Date());
}

function waitForLoading(dataRange, sheet, maxWaitTimeInSec)
{
  for(i = 0; i< maxWaitTimeInSec ; i++)
  {
    var value = dataRange.getCell(1,1).getValue();
    if(value.search("Loading") !== -1) {
      Utilities.sleep(1000);
      dataRange = sheet.getDataRange();
    } else {
      return true;
    }
    
  }
  return false;
}

function getSendTodayHolidays() {
  var documentProperties = PropertiesService.getUserProperties();
  var keys = documentProperties.getKeys();  
  var lastNotificationProperty = undefined;
  if(keys.includes('LastHolidayTime'))
    lastNotificationProperty = documentProperties.getProperty('LastHolidayTime');
   
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
