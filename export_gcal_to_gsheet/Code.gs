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

var mycal = "justingale2@gmail.com";
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
var events = cal.getEvents(new Date("January 12, 2014 00:00:00 CST"), new Date("January 18, 2014 23:59:59 CST"), {search: '-project123'});


var sheet = SpreadsheetApp.getActiveSheet();
// Uncomment this next line if you want to always clear the spreadsheet content before running - Note people could have added extra columns on the data though that would be lost
// sheet.clearContents();  

// Create a header record on the current spreadsheet in cells A1:N1 - Match the number of entries in the "header=" to the last parameter
// of the getRange entry below
var header = [["Calendar Address", "Event Title", "Event Description", "Event Location", "Event Start", "Event End", "Calculated Duration", "Visibility", "Date Created", "Last Updated", "MyStatus", "Created By", "All Day Event", "Recurring Event"]]
var range = sheet.getRange(1,1,1,14);
range.setValues(header);

  
// Loop through all calendar events found and write them out starting on calulated ROW 2 (i+2)
for (var i=0;i<events.length;i++) {
var row=i+2;
var myformula_placeholder = '';
// Matching the "header=" entry above, this is the detailed row entry "details=", and must match the number of entries of the GetRange entry below
// NOTE: I've had problems with the getVisibility for some older events not having a value, so I've had do add in some NULL text to make sure it does not error
var details=[[mycal,events[i].getTitle(), events[i].getDescription(), events[i].getLocation(), events[i].getStartTime(), events[i].getEndTime(), myformula_placeholder, ('' + events[i].getVisibility()), events[i].getDateCreated(), events[i].getLastUpdated(), events[i].getMyStatus(), events[i].getCreators(), events[i].isAllDayEvent(), events[i].isRecurringEvent()]];
var range=sheet.getRange(row,1,1,14);
range.setValues(details);

// Writing formulas from scripts requires that you write the formulas separate from non-formulas
// Write the formula out for this specific row in column 7 to match the position of the field myformula_placeholder from above: foumula over columns F-E for time calc
var cell=sheet.getRange(row,7);
cell.setFormula('=(HOUR(F' +row+ ')+(MINUTE(F' +row+ ')/60))-(HOUR(E' +row+ ')+(MINUTE(E' +row+ ')/60))');
cell.setNumberFormat('.00');

}
}
function onOpen() {
  Browser.msgBox('App Instructions - Please Read This Message', '1) Click Tools then Script Editor\\n2) Read/update the code with your desired values.\\n3) Then when ready click Run export_gcal_to_gsheet from the script editor.', Browser.Buttons.OK);

}


