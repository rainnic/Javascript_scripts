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
var startingDate = "2021/10/01"; // data iniziale scritta in ANNO/MESE/GIORNO
var endDate = "2021/10/31"; // data finale scritta in ANNO/MESE/GIORNO
var night_timing = [22, 6]; // intervallo di orario notturno --> night_timing[0]
var night = 1; // per aggiungere la colonna delle ore notturne sul foglio si(1)/no(0)
var feast = 1; // per aggiungere la colonna delle ore festive sul foglio si(1)/no(0)
var feast_days = "\
DATE(YEAR(A1); 1; 1);\
DATE(YEAR(A1); 1; 6);\
DATE(YEAR(A1); 4; 21);\
DATE(YEAR(A1); 4; 22);\
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

// Check if calendar is empty
if (events.length == 0) {
  startTime = startingDate;
  sheet.getRange(3,2).setValue('Il calendario è vuoto!').setFontStyle('bold').setHorizontalAlignment("center");
  } else {startTime = events[0].getStartTime()}

// Header of the sheet
sheet.getRange(1,1).setValue(startTime).setNumberFormat("YYYY/MMMM").setHorizontalAlignment("left");
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

// Variabili per determinare la lunghezza e larghezza della tabella
var totalRows = sheet.getLastRow();
var totalColumns = sheet.getLastColumn();
var firstRowDate = 3;
  
// Per aggiungere altre voci alla intestazione della tabella
// Ore notturne
if (night) {sheet.getRange(2,totalColumns+1).setValue("Ore notturne").setHorizontalAlignment("center").setBackground(headerColor);}
var total_night_hours = 0;
// Ore festive
if (feast) {sheet.getRange(2,totalColumns+2).setValue("Ore festive").setHorizontalAlignment("center").setBackground(headerColor);}
var total_feast_hours = 0;
// Giorni lavorati
var total_worked_days = 0;
  
// Variabili per determinare la nuova larghezza della tabella da colorare
if ((night && feast) || (!night && feast)) {
   var totalColoredColumns = totalColumns+2;
 } else if (night && !feast) {
   var totalColoredColumns = totalColumns+1;
 } else {
   var totalColoredColumns = totalColumns;
 }
  
// Variabili usate per i colori alternati
var columnColorCalc = 28;
var color = firstColor;
var FirstWorkedDay = sheet.getRange(firstRowDate,columnColorCalc).setFormula('=(DATE(YEAR(A' +firstRowDate+ ');MONTH(A' +firstRowDate+ ');DAY(A' +firstRowDate+ '))-DATE(YEAR(A' +firstRowDate+ ');1;0))').getValue();

// Qui miglioro la formattazione della tabella
for (var i=firstRowDate; i <= totalRows; i+=1){
    sheet.getRange(i,1).setNumberFormat("-DD/MM-").setHorizontalAlignment("center");
    sheet.getRange(i,3,totalRows,2).setNumberFormat("HH:mm");

    // Codice per i colori alternati
    var workedDay = sheet.getRange(i,columnColorCalc).setFormula('=(DATE(YEAR(A' +i+ ');MONTH(A' +i+ ');DAY(A' +i+ '))-DATE(YEAR(A' +i+ ');1;0))').getValue();
    if( FirstWorkedDay == workedDay ){
        sheet.getRange(i, 1, 1, totalColoredColumns).setBackground(color);
    } else if (color == firstColor) { var FirstWorkedDay = sheet.getRange(i,columnColorCalc).setFormula('=(DATE(YEAR(A' +i+ ');MONTH(A' +i+ ');DAY(A' +i+ '))-DATE(YEAR(A' +i+ ');1;0))').getValue(); var color = secondColor; sheet.getRange(i, 1, 1, totalColoredColumns).setBackground(color);
    } else if (color == secondColor) { var FirstWorkedDay = sheet.getRange(i,columnColorCalc).setFormula('=(DATE(YEAR(A' +i+ ');MONTH(A' +i+ ');DAY(A' +i+ '))-DATE(YEAR(A' +i+ ');1;0))').getValue(); var color = firstColor; sheet.getRange(i, 1, 1, totalColoredColumns).setBackground(color);
    }
    // Codice per i colori alternati
  
    // Codice per l'orario notturno
    sheet.getRange(i,totalColumns+1).setFormula('=HOUR(VALUE((MOD((TIME(HOUR(D' +i+ ');MINUTE(D' +i+ ');0))-(TIME(HOUR(C' +i+ ');MINUTE(C' +i+ ');0));1)*24'
                                                +'-((TIME(HOUR(D' +i+ ');MINUTE(D' +i+ ');0))<(TIME(HOUR(C' +i+ ');MINUTE(C' +i+ ');0)))*('+night_timing[0]+'-'+night_timing[1]+')'
                                                +'+MEDIAN('+night_timing[1]+';'+night_timing[0]+';(TIME(HOUR(C' +i+ ');'
                                                +'MINUTE(C' +i+ ');0))*24)-MEDIAN((TIME(HOUR(D' +i+ ');MINUTE(D' +i+ ');0))*24;'+night_timing[1]+';'+night_timing[0]+'))/24))+MINUTE(VALUE((MOD((TIME(HOUR(D' +i+ ');MINUTE(D' +i+ ');0))'
                                                +'-(TIME(HOUR(D' +i+ ');MINUTE(C' +i+ ');0));1)*24-((TIME(HOUR(D' +i+ ');MINUTE(D' +i+ ');0))<(TIME(HOUR(C' +i+ ');MINUTE(C' +i+ ');0)))*('+night_timing[0]+'-'+night_timing[1]+')'
                                                +'+MEDIAN('+night_timing[1]+';'+night_timing[0]+';(TIME(HOUR(C' +i+ ');'
                                                +'MINUTE(C' +i+ ');0))*24)-MEDIAN((TIME(HOUR(D' +i+ ');MINUTE(D' +i+ ');0))*24;'+night_timing[1]+';'+night_timing[0]+'))/24))/60').setNumberFormat('0.00').setHorizontalAlignment("center");
    total_night_hours = total_night_hours + sheet.getRange(i,totalColumns+1).getValue();
    // sheet.getRange(i,totalColumns+2).setValue(total_night_hours);
    // Codice per l'orario notturno
  
    // Codice per l'orario festivo
    sheet.getRange(i,totalColumns+2).setFormula('=IF(NETWORKDAYS.INTL(A' +i+ '; A' +i+ '; "0000000";{'+ feast_days +'})=0;E' +i+ ';0)').setHorizontalAlignment("center");
    total_feast_hours = total_feast_hours + sheet.getRange(i,totalColumns+2).getValue();
    // sheet.getRange(i,totalColumns+3).setValue(total_night_hours);
    // Codice per l'orario festivo
}

// Calcolo giorni lavorati
  sheet.getRange(totalRows+1,columnColorCalc).setFormula('=COUNTUNIQUE(AB3:AB' +totalRows+ ')').setNumberFormat('0');
  total_worked_days = sheet.getRange(totalRows+1,columnColorCalc).getValue();
  sheet.getRange(totalRows+1,columnColorCalc).clear(); // Cancella la cella dopo averne salvato il contenuto
  
// Puliza delle colonne usate per i calcoli
for (var i=firstRowDate; i <= totalRows; i+=1){
    sheet.getRange(i,columnColorCalc).clear();              // La colonna usata per cambiare i colori
    if (!night) {sheet.getRange(i,totalColumns+1).clear()}; // La colonna usata per le ore notturne
    if (!feast) {sheet.getRange(i,totalColumns+2).clear()}; // La colonna usata per le ore festive
}

// Sommatorie  
sheet.getRange(totalRows+2,4).setValue('Σ=').setNumberFormat('0').setHorizontalAlignment("right");
sheet.getRange(totalRows+2,5).setFormula('=SUM(E2:E' +totalRows+ ')').setNumberFormat('0.00 \\o\\r\\e').setHorizontalAlignment("right"); // somma ore totali
  
sheet.getRange(totalRows+3,4).setValue('Σnotturne('+night_timing[0]+'-'+night_timing[1]+')=').setNumberFormat('0').setHorizontalAlignment("right");
sheet.getRange(totalRows+3,5).setValue(total_night_hours).setNumberFormat('0.00 \\o\\r\\e').setHorizontalAlignment("right"); // somma ore notturne

sheet.getRange(totalRows+4,4).setValue('Σfestive=').setNumberFormat('0').setHorizontalAlignment("right");
sheet.getRange(totalRows+4,5).setValue(total_feast_hours).setNumberFormat('0.00 \\o\\r\\e').setHorizontalAlignment("right"); // somma ore festive
  
sheet.getRange(totalRows+6,4).setValue('Σlavorati=').setNumberFormat('0').setHorizontalAlignment("right");
  if (total_worked_days == 1) {
    sheet.getRange(totalRows+6,5).setValue(total_worked_days).setNumberFormat('0 \\g\\i\\o\\r\\n\\o').setHorizontalAlignment("right"); // se 1 mostra il giorno lavorato
} else {
    sheet.getRange(totalRows+6,5).setValue(total_worked_days).setNumberFormat('0 \\g\\i\\o\\r\\n\\i').setHorizontalAlignment("right"); // somma giorni lavorati totali
}
  
}
  
function onOpen() {
  Browser.msgBox('Istruzioni - Leggi questo messaggio prima', '1) Clicca Strumenti e poi Editor di script\\n2) Leggi/aggiorna il codice con i tuoi valori.\\n3) Quando pronto clicca su Esegui -> Esegui funzione -> export_gcal_to_gsheet dall`Editor di script.', Browser.Buttons.OK);

}
