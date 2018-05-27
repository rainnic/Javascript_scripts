function export_gcal_to_gsheet(){

//
// Export Google Calendar Events to a Google Spreadsheet
// Original script written by Justin Gale
// found in Export Google Calendar Entries to a Google Spreadsheet (https://www.cloudbakers.com/blog/export-google-calendar-entries-to-a-google-spreadsheet)
//
// This code retrieves events between 2 dates for the specified calendar.
// It logs the results in the current spreadsheet starting at cell A2 listing the events,
// dates/times, etc and even calculates event duration (via creating formulas in the spreadsheet) and formats the values.
//
// I do re-write the spreadsheet header in Row 1 with every run, as I found it faster to delete then entire sheet content,
// change my parameters, and re-run my exports versus trying to save the header row manually...so be sure if you change
// any code, you keep the header in agreement for readability!
//
// 1. Please modify the value for mycal to be YOUR calendar email address or one visible on your MY Calendars section of your Google Calendar
// 2. Please modify the values for events to be the date/time range you want and any search parameters to find or omit calendar entires
// Note: Events can be easily filtered out/deleted once exported from the calendar
// 
// Reference Websites:
// https://developers.google.com/apps-script/reference/calendar/calendar
// https://developers.google.com/apps-script/reference/calendar/calendar-event
//

// SETTINGS
var calendarID = "PUT_HERE_YOUR_CALENDAR_ID"; //for example j4k34jl65hl5jh3ljj4l3@group.calendar.google.com
var sheetTitle = "Working hours of Nicola Rainiero";
var startingDate = "2018/05/01";
var endDate = "2018/05/31";
var holydays = "\
DATE(YEAR(A1); 1; 1);\
DATE(YEAR(A1); 1; 6);\
DATE(YEAR(A1); 4; 1);\
DATE(YEAR(A1); 4; 2);\
DATE(YEAR(A1); 4; 25);\
DATE(YEAR(A1); 5; 1);\
DATE(YEAR(A1); 6; 2);\
DATE(YEAR(A1); 6; 13);\
DATE(YEAR(A1); 8; 15);\
DATE(YEAR(A1); 11; 1);\
DATE(YEAR(A1); 12; 8);\
DATE(YEAR(A1); 12; 25);\
DATE(YEAR(A1); 12; 26)\
";
  
var mycal = calendarID;
var cal = CalendarApp.getCalendarById(mycal);
  
// Optional variations on getEvents
// var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"));
// var events = cal.getEvents(new Date("January 3, 2014 00:00:00 CST"), new Date("January 14, 2014 23:59:59 CST"), {search: 'word1'});
// 
// Explanation of how the search section works (as it is NOT quite like most things Google) as part of the getEvents function:
//    {search: 'word1'}              Search for events with word1
//    {search: '-word1'}             Search for events without word1
//    {search: 'word1 word2'}        Search for events with word2 ONLY
//    {search: 'word1-word2'}        Search for events with ????
//    {search: 'word1 -word2'}       Search for events without word2
//    {search: 'word1+word2'}        Search for events with word1 AND word2
//    {search: 'word1+-word2'}       Search for events with word1 AND without word2
//
// var events = cal.getEvents(new Date("January 12, 2014 00:00:00 CST"), new Date("January 18, 2014 23:59:59 CST"), {search: '-project123'});
var events = cal.getEvents(new Date(startingDate+" 00:00:00 UTC +1"), new Date(endDate+" 23:59:59 UTC +1"));


var sheet = SpreadsheetApp.getActiveSheet();
// Uncomment this next line if you want to always clear the spreadsheet content before running - Note people could have added extra columns on the data though that would be lost
sheet.clearContents();
sheet.clearFormats();

// Header of the sheet
sheet.getRange(1,1).setValue(events[0].getStartTime()).setNumberFormat("YYYY/MMMM").setHorizontalAlignment("left");
sheet.getRange(1,2).setValue(sheetTitle).setNumberFormat('0').setHorizontalAlignment("left");
  
sheet.getRange(1,3).setValue(startingDate).setNumberFormat("Fro\\m DD").setHorizontalAlignment("left");
sheet.getRange(1,4).setValue(endDate).setNumberFormat("To DD").setHorizontalAlignment("left");
  
// Create a header record on the current spreadsheet in cells A1:N1 - Match the number of entries in the "header=" to the last parameter
// of the getRange entry below
var header = [["Day", "Title", "Start", "End", "Duration", "Description", "Location"]]
var range = sheet.getRange(2,1,1,7);
range.setValues(header);

  
// Loop through all calendar events found and write them out starting on calulated ROW 2 (i+2)
for (var i=0;i<events.length;i++) {
var row=i+3;
var myformula_placeholder = '';
// Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
// NOTE: I've had problems with the getVisibility for some older events not having a value, so I've had do add in some NULL text to make sure it does not error
var details=[[events[i].getStartTime(), events[i].getTitle(), events[i].getStartTime(), events[i].getEndTime(), myformula_placeholder, events[i].getDescription(), events[i].getLocation()]];
var range=sheet.getRange(row,1,1,7);
range.setValues(details);

// Writing formulas from scripts requires that you write the formulas separate from non-formulas
// Write the formula out for this specific row in column 7 to match the position of the field myformula_placeholder from above: foumula over columns E-D for time calc
var cell=sheet.getRange(row,5);
cell.setFormula('=((DAY(D' +row+ ')*24+HOUR(D' +row+ ')+(MINUTE(D' +row+ ')/60))-(DAY(C' +row+ ')*24+HOUR(C' +row+ ')+(MINUTE(C' +row+ ')/60)))');
cell.setNumberFormat('.00');

}
  
var totalRows = sheet.getLastRow();
var firstRowDate = 3;

for (var i=firstRowDate; i <= totalRows; i+=1){
    sheet.getRange(i,1).setNumberFormat("-DD-").setHorizontalAlignment("center");
    sheet.getRange(i,3,totalRows,2).setNumberFormat("HH:mm");
}

sheet.getRange(totalRows+2,4).setValue('SUM').setNumberFormat('0').setHorizontalAlignment("right");
sheet.getRange(totalRows+2,5).setFormula('=SUM(E2:E' +totalRows+ ')').setNumberFormat('0.00').setHorizontalAlignment("right"); // sum duration
  
}
function onOpen() {
  Browser.msgBox('App Instructions - Please Read This Message', '1) Click Tools then Script Editor\\n2) Read/update the code with your desired values.\\n3) Then when ready click Run export_gcal_to_gsheet from the script editor.', Browser.Buttons.OK);

}

