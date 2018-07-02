function export_gcal_to_gsheet(){

// -----------------------------------------------------------
// Generare un foglio Google dalle voci del proprio calendario
// -----------------------------------------------------------
// Script modificato da Nicola Rainiero (https://rainnic.altervista.org/tag/google-apps-script)
// L'originale è stato scritto da Justin Gale e pubblicato su Export Google Calendar Entries to a Google Spreadsheet (https://www.cloudbakers.com/blog/export-google-calendar-entries-to-a-google-spreadsheet).
//
// Il codice importa gli eventi compresi tra 2 date per un calendario specifico.
// Registra i risultati nel foglio di calcolo corrente a partire dalla cella A3 che elenca gli eventi,
// data/ora, ecc., calcola anche la durata dell'evento (tramite la creazione di formule nel foglio di calcolo) e formatta i valori.
//
// Nella sezione IMPOSTAZIONI:
// 1. modifica il valore di calendarID con l'ID del tuo calendario visibile nella sezione I miei calendari di Google Calendar;
// 2. mdifica i valori per gli eventi in base all'intervallo di data/ora desiderato e qualsiasi parametro di ricerca per trovare o omettere le voci del calendario.
// Nota: gli eventi possono essere facilmente filtrati/eliminati una volta esportati dal calendario
// 
// Fonti:
// https://developers.google.com/apps-script/reference/calendar/calendar
// https://developers.google.com/apps-script/reference/calendar/calendar-event

// IMPOSTAZIONI
var calendarID = "METTI_QUI_IL_TUO_ID_DEL_CALENDARIO"; // del tipo j4k34jl65hl5jh3ljj4l3@group.calendar.google.com
var sheetTitle = "Ore lavorate di TUO NOME"; // titolo della tabella
var startingDate = "2018/06/01"; // data iniziale scritta in ANNO/MESE/GIORNO
var endDate = "2018/06/30"; // data finale scritta in ANNO/MESE/GIORNO
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
"; // festività nazionali e locali scritte in ANNO/MESE/GIORNO
// Stili per la tabella
var headerColor = "#EDD400"; // giallo per l'intestazione della tabella
var firstColor = "#FFFFFF"; // bianco per la prima riga alternata
var secondColor = "#E0E0E0"; // grigio per la seconda riga alternata

  
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
  
sheet.getRange(1,3).setValue(startingDate).setNumberFormat("\\Da DD/MM").setHorizontalAlignment("left");
sheet.getRange(1,4).setValue(endDate).setNumberFormat("A DD/MM").setHorizontalAlignment("left");
  
// Create a header record on the current spreadsheet in cells A1:N1 - Match the number of entries in the "header=" to the last parameter
// of the getRange entry below
var header = [["Giorno", "Titolo", "Inizio", "Fine", "Durata (ore)", "Descrizione", "Luogo"]]
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
  
var totalRows = sheet.getLastRow();
var totalColumns = sheet.getLastColumn();
var firstRowDate = 3;
// Variabili usate per i colori alternati
var columnColorCalc = 28;
var color = firstColor;
var FirstWorkingDay = sheet.getRange(firstRowDate,columnColorCalc).setFormula('=(DATE(YEAR(A' +firstRowDate+ ');MONTH(A' +firstRowDate+ ');DAY(A' +firstRowDate+ '))-DATE(YEAR(A' +firstRowDate+ ');1;0))').getValue();

// Qui miglioro la formattazione della tabella
for (var i=firstRowDate; i <= totalRows; i+=1){
    sheet.getRange(i,1).setNumberFormat("-DD/MM-").setHorizontalAlignment("center");
    sheet.getRange(i,3,totalRows,2).setNumberFormat("HH:mm");

    // Codice per i colori alternati
    var workingDay = sheet.getRange(i,columnColorCalc).setFormula('=(DATE(YEAR(A' +i+ ');MONTH(A' +i+ ');DAY(A' +i+ '))-DATE(YEAR(A' +i+ ');1;0))').getValue();
    if( FirstWorkingDay == workingDay ){
        sheet.getRange(i, 1, 1, totalColumns).setBackground(color);
    } else if (color == firstColor) { var FirstWorkingDay = sheet.getRange(i,columnColorCalc).setFormula('=(DATE(YEAR(A' +i+ ');MONTH(A' +i+ ');DAY(A' +i+ '))-DATE(YEAR(A' +i+ ');1;0))').getValue(); var color = secondColor; sheet.getRange(i, 1, 1, totalColumns).setBackground(color);
    } else if (color == secondColor) { var FirstWorkingDay = sheet.getRange(i,columnColorCalc).setFormula('=(DATE(YEAR(A' +i+ ');MONTH(A' +i+ ');DAY(A' +i+ '))-DATE(YEAR(A' +i+ ');1;0))').getValue(); var color = firstColor; sheet.getRange(i, 1, 1, totalColumns).setBackground(color);
    }
    // Codice per i colori alternati

}

// Puliza delle colonne usate per i calcoli
for (var i=firstRowDate; i <= totalRows; i+=1){
    // La colonna usata per cambiare i colori
    sheet.getRange(i,columnColorCalc).clear();
}
  
sheet.getRange(totalRows+2,4).setValue('Σ=').setNumberFormat('0').setHorizontalAlignment("right");
sheet.getRange(totalRows+2,5).setFormula('=SUM(E2:E' +totalRows+ ')').setNumberFormat('0.00 \\o\\r\\e').setHorizontalAlignment("left"); // sum duration
  
}
function onOpen() {
  Browser.msgBox('Istruzioni - Leggi questo messaggio prima', '1) Clicca Strumenti e poi Editor di script\\n2) Leggi/aggiorna il codice con i tuoi valori.\\n3) Quando pronto clicca su Esegui -> Esegui funzione -> export_gcal_to_gsheet dall`Editor di script.', Browser.Buttons.OK);

}
