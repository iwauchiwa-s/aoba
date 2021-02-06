// Google Apps Script 'GCalender_access'
// version 1.0 (2021/02/06)

function GetCalenderList(){
    var calendars = CalendarApp.getAllCalendars();
    var sheet = SpreadsheetApp.getActive().getSheetByName('Gカレンダー');
    sheet.getRange('A10:A110').clear();
    for(var i=0; i<calendars.length; i++)
    {
      var calendar = calendars[i];      
      sheet.getRange('A'+(i+10)).
      setValue(calendar.getName());
    }
}


function GetCalendarEvent(){
  var sheet = SpreadsheetApp.getActive().getSheetByName('Gカレンダー');
  var objCalendar = CalendarApp.getOwnedCalendarsByName(sheet.getRange('A5').getValue());
  sheet.getRange('C10:G900').clear();
  var startDate = new Date(sheet.getRange('A6').getValue()); 
  var endDate = new Date(sheet.getRange('A7').getValue());
  var events = objCalendar[0].getEvents(startDate,endDate); 
  for(var i=0; i<events.length; i++){
    sheet.getRange('D'+(i+10)).setValue(events[i].getTitle());
    sheet.getRange('E'+(i+10)).setValue(events[i].getStartTime());
    sheet.getRange('F'+(i+10)).setValue(events[i].getEndTime());
    sheet.getRange('G'+(i+10)).setValue(events[i].getDescription());
  }
}

function DeleteEvents() {

  // Reload events for check
  var sheet = SpreadsheetApp.getActive().getSheetByName('Gカレンダー');
  var objCalendar = CalendarApp.getOwnedCalendarsByName(sheet.getRange('A5').getValue());
  var start = new Date(sheet.getRange('A6').getValue()); 
  var end = new Date(sheet.getRange('A7').getValue());
  var events = objCalendar[0].getEvents(start,end); 

  //Double-check variables
  var chkTitle ;
  var chkTime ;
  var objTitle ;
  var objTime ;
  var objFlag ;

  for(var i=0; i<events.length; i++){

    chkTitle = events[i].getTitle();
    chkTime = (events[i].getStartTime()).getTime();

    objTitle = sheet.getRange('D'+(i+10)).getValue();
    objTime = (sheet.getRange('E'+(i+10)).getValue()).getTime();
    objFlag = sheet.getRange('C'+(i+10)).getValue();
    
    if ( objFlag == 'd'){
      if ( chkTitle == objTitle ){
        if ( chkTime == objTime){
          events[i].deleteEvent() ;
        }
      }
    }
  }
}

function CreateEvents() {

  var sheet = SpreadsheetApp.getActive().getSheetByName('Gカレンダー');
  var objCalendar = CalendarApp.getOwnedCalendarsByName(sheet.getRange('A5').getValue());
  var id = objCalendar[0].getId();
  var CalendarbyID = CalendarApp.getCalendarById(id);

  var objTitle ;
  var objStartTime ;
  var objEndTime ;
  var objDescription ;
  
  var objValD = sheet.getRange('D10:D110').getValues();
  var objNum = objValD.filter(String).length;

  for(var i=0; i<objNum+1; i++){

    objTitle = sheet.getRange('D'+(i+10)).getValue();
    objStartTime = new Date(sheet.getRange('E'+(i+10)).getValue());
    objEndTime = new Date (sheet.getRange('F'+(i+10)).getValue());
    objDescription = sheet.getRange('G'+(i+10)).getValue();

    CalendarbyID.createEvent(objTitle, objStartTime, objEndTime, 
               {description: objDescription}); 
  }
}
