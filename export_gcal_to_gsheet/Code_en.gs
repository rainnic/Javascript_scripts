function export_gcal_to_gsheet(){

// -----------------------------------------------------
// Export Google Calendar Events to a Google Spreadsheet
// -----------------------------------------------------
// The script was modified by Nicola Rainiero (https://rainnic.altervista.org/tag/google-apps-script)
// It was written by Justin Gale and found in Export Google Calendar Entries to a Google Spreadsheet (https://www.cloudbakers.com/blog/export-google-calendar-entries-to-a-google-spreadsheet)
//
// This code retrieves events between 2 dates for the specified calendar.
// It logs the results in the current spreadsheet starting at cell A2 listing the events,
// dates/times, etc and even calculates event duration (via creating formulas in the spreadsheet) and formats the values.
//
// I do re-write the spreadsheet header in Row 1 with every run, as I found it faster to delete then entire sheet content,
// change my parameters, and re-run my exports versus trying to save the header row manually...so be sure if you change
// any code, you keep the header in agreement for readability!
//
// In the SETTINGS section you have to edit:
// 1. the value for calendarID to be YOUR ID or one visible on your MY Calendars section of your Google Calendar;
// 2. the values for events to be the date/time range you want and any search parameters to find or omit calendar entires.
// Note: Events can be easily filtered out/deleted once exported from the calendar
// 
// Reference Websites:
// https://developers.google.com/apps-script/reference/calendar/calendar
// https://developers.google.com/apps-script/reference/calendar/calendar-event

// SETTINGS
var calendarID = "PUT_HERE_YOUR_CALENDAR_ID"; // must be like this example j4k34jl65hl5jh3ljj4l3@group.calendar.google.com
var sheetTitle = "Working hours of YOUR NAME"; // title of the Spreadsheet
var startingDate = "2019/06/01"; // starting date formatted in YEAR/MONTH/DAY
var endDate = "2019/06/30"; // end date formatted in YEAR/MONTH/DAY
var night_timing = [22, 6]; // to set the range of night hours --> night_timing[0]
var night = 1; // for adding the night column on the final sheet on(1)/off(0)  
var feast = 1; // for adding the holiday column on the final sheet on(1)/off(0)
var feast_days = "\
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
"; // list of national holidays formatted in YEAR/MONTH/DAY
// Styles for the table
var headerColor = "#EDD400"; // yellow for the header
var firstColor = "#FFFFFF"; // white for the first alternate colours
var secondColor = "#E0E0E0"; // gray for the second alternate colours
  
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
var events = cal.getEvents(new Date(startingDate+" 00:00:00 UTC +2"), new Date(endDate+" 23:59:59 UTC +2"));


var sheet = SpreadsheetApp.getActiveSheet();
// Uncomment this next line if you want to always clear the spreadsheet content before running - Note people could have added extra columns on the data though that would be lost
sheet.clearContents();
sheet.clearFormats();

// Header of the sheet
sheet.getRange(1,1).setValue(events[0].getStartTime()).setNumberFormat("YYYY/MMMM").setHorizontalAlignment("left");
sheet.getRange(1,2).setValue(sheetTitle).setNumberFormat('0').setHorizontalAlignment("left");
  
sheet.getRange(1,3).setValue(startingDate).setNumberFormat("Fro\\m MM/DD").setHorizontalAlignment("left");
sheet.getRange(1,4).setValue(endDate).setNumberFormat("To MM/DD").setHorizontalAlignment("left");
  
// Create a header record on the current spreadsheet in cells A1:N1 - Match the number of entries in the "header=" to the last parameter
// of the getRange entry below
var header = [["Day", "Title", "Start", "End", "Duration (hours)", "Description", "Location"]]
var range = sheet.getRange(2,1,1,7);
range.setValues(header).setBackground(headerColor);

  
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
  
// Variables to fix the dimension of the sheet loaded from Google Calendar
var totalRows = sheet.getLastRow();
var totalColumns = sheet.getLastColumn();
var firstRowDate = 3;
  
// To add new columns after the initial others
// Night hours
if (night) {sheet.getRange(2,totalColumns+1).setValue("Night hours").setHorizontalAlignment("center").setBackground(headerColor);}
var total_night_hours = 0;
// Holiday shift hours
if (feast) {sheet.getRange(2,totalColumns+2).setValue("Holiday hours").setHorizontalAlignment("center").setBackground(headerColor);}
var total_feast_hours = 0;
// Working days
var total_working_days = 0;
  
// To set the new number of columns to paint
if ((night && feast) || (!night && feast)) {
   var totalColoredColumns = totalColumns+2;
 } else if (night && !feast) {
   var totalColoredColumns = totalColumns+1;
 } else {
   var totalColoredColumns = totalColumns;
 }
  
// Variables used to alternate colours
var columnColorCalc = 28;
var color = firstColor;
var FirstWorkingDay = sheet.getRange(firstRowDate,columnColorCalc).setFormula('=(DATE(YEAR(A' +firstRowDate+ ');MONTH(A' +firstRowDate+ ');DAY(A' +firstRowDate+ '))-DATE(YEAR(A' +firstRowDate+ ');1;0))').getValue();

// Code to improve the stilish of the table
for (var i=firstRowDate; i <= totalRows; i+=1){
    sheet.getRange(i,1).setNumberFormat("-DD/MM-").setHorizontalAlignment("center");
    sheet.getRange(i,3,totalRows,2).setNumberFormat("HH:mm");

    // Code to alternate colours
    var workingDay = sheet.getRange(i,columnColorCalc).setFormula('=(DATE(YEAR(A' +i+ ');MONTH(A' +i+ ');DAY(A' +i+ '))-DATE(YEAR(A' +i+ ');1;0))').getValue();
    if( FirstWorkingDay == workingDay ){
        sheet.getRange(i, 1, 1, totalColoredColumns).setBackground(color);
    } else if (color == firstColor) { var FirstWorkingDay = sheet.getRange(i,columnColorCalc).setFormula('=(DATE(YEAR(A' +i+ ');MONTH(A' +i+ ');DAY(A' +i+ '))-DATE(YEAR(A' +i+ ');1;0))').getValue(); var color = secondColor; sheet.getRange(i, 1, 1, totalColoredColumns).setBackground(color);
    } else if (color == secondColor) { var FirstWorkingDay = sheet.getRange(i,columnColorCalc).setFormula('=(DATE(YEAR(A' +i+ ');MONTH(A' +i+ ');DAY(A' +i+ '))-DATE(YEAR(A' +i+ ');1;0))').getValue(); var color = firstColor; sheet.getRange(i, 1, 1, totalColoredColumns).setBackground(color);
    }
     // Code to alternate colours

    // Code to count working night hours
    sheet.getRange(i,totalColumns+1).setFormula('=HOUR(VALUE((MOD((TIME(HOUR(D' +i+ ');MINUTE(D' +i+ ');0))-(TIME(HOUR(C' +i+ ');MINUTE(C' +i+ ');0));1)*24'
                                                +'-((TIME(HOUR(D' +i+ ');MINUTE(D' +i+ ');0))<(TIME(HOUR(C' +i+ ');MINUTE(C' +i+ ');0)))*('+night_timing[0]+'-'+night_timing[1]+')'
                                                +'+MEDIAN('+night_timing[1]+';'+night_timing[0]+';(TIME(HOUR(C' +i+ ');'
                                                +'MINUTE(C' +i+ ');0))*24)-MEDIAN((TIME(HOUR(D' +i+ ');MINUTE(D' +i+ ');0))*24;'+night_timing[1]+';'+night_timing[0]+'))/24))+MINUTE(VALUE((MOD((TIME(HOUR(D' +i+ ');MINUTE(D' +i+ ');0))'
                                                +'-(TIME(HOUR(D' +i+ ');MINUTE(C' +i+ ');0));1)*24-((TIME(HOUR(D' +i+ ');MINUTE(D' +i+ ');0))<(TIME(HOUR(C' +i+ ');MINUTE(C' +i+ ');0)))*('+night_timing[0]+'-'+night_timing[1]+')'
                                                +'+MEDIAN('+night_timing[1]+';'+night_timing[0]+';(TIME(HOUR(C' +i+ ');'
                                                +'MINUTE(C' +i+ ');0))*24)-MEDIAN((TIME(HOUR(D' +i+ ');MINUTE(D' +i+ ');0))*24;'+night_timing[1]+';'+night_timing[0]+'))/24))/60').setNumberFormat('0.00').setHorizontalAlignment("center");
    total_night_hours = total_night_hours + sheet.getRange(i,totalColumns+1).getValue();
    // sheet.getRange(i,totalColumns+2).setValue(total_night_hours);
    // Code to count night hours
  
    // Code to count working holiday hours
    sheet.getRange(i,totalColumns+2).setFormula('=IF(NETWORKDAYS.INTL(A' +i+ '; A' +i+ '; "0000000";{'+ feast_days +'})=0;E' +i+ ';0)').setHorizontalAlignment("center");
    total_feast_hours = total_feast_hours + sheet.getRange(i,totalColumns+2).getValue();
    // sheet.getRange(i,totalColumns+3).setValue(total_night_hours);
    // Code to count working holiday hours

}
  
// Code to count working days
  sheet.getRange(totalRows+1,columnColorCalc).setFormula('=COUNTUNIQUE(AB3:AB' +totalRows+ ')').setNumberFormat('0');
  total_working_days = sheet.getRange(totalRows+1,columnColorCalc).getValue();
  sheet.getRange(totalRows+1,columnColorCalc).clear(); // To clear the cell after storing the variable

// Clear columns added for calculations
for (var i=firstRowDate; i <= totalRows; i+=1){
    sheet.getRange(i,columnColorCalc).clear(); // The column used to change the colors
    if (!night) {sheet.getRange(i,totalColumns+1).clear()}; // The column used to count night hours
    if (!feast) {sheet.getRange(i,totalColumns+2).clear()}; // The column used to count holiday hours
}

sheet.getRange(totalRows+2,4).setValue('Σ=').setNumberFormat('0').setHorizontalAlignment("right");
sheet.getRange(totalRows+2,5).setFormula('=SUM(E2:E' +totalRows+ ')').setNumberFormat('0.00 \\h\\o\\u\\r\\s').setHorizontalAlignment("left"); // shows total duration
  
sheet.getRange(totalRows+3,4).setValue('Σnight('+night_timing[0]+'-'+night_timing[1]+')=').setNumberFormat('0').setHorizontalAlignment("right");
sheet.getRange(totalRows+3,5).setValue(total_night_hours).setNumberFormat('0.00 \\h\\o\\u\\r\\s').setHorizontalAlignment("right"); // shows total night hours

sheet.getRange(totalRows+4,4).setValue('Σholiday=').setNumberFormat('0').setHorizontalAlignment("right");
sheet.getRange(totalRows+4,5).setValue(total_feast_hours).setNumberFormat('0.00 \\h\\o\\u\\r\\s').setHorizontalAlignment("right"); // shows total holiday hours
  
sheet.getRange(totalRows+6,4).setValue('Σworking=').setNumberFormat('0').setHorizontalAlignment("right");
  if (total_working_days == 1) {
    sheet.getRange(totalRows+6,5).setValue(total_working_days).setNumberFormat('0 \\d\\a\\y').setHorizontalAlignment("right"); // shows total working day
} else {
    sheet.getRange(totalRows+6,5).setValue(total_working_days).setNumberFormat('0 \\d\\a\\y\\s').setHorizontalAlignment("right"); // shows total working days
}
  
}
function onOpen() {
  Browser.msgBox('App Instructions - Please Read This Message', '1) Click Tools then Script Editor\\n2) Read/update the code with your desired values.\\n3) Then when ready click Run export_gcal_to_gsheet from the script editor.', Browser.Buttons.OK);

}
