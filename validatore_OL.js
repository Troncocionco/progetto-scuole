// Use these first two lines for all four examples
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];
 
// Single cell
var cell = sheet.getRange("B2");
 
// Plain text
cell.setNumberFormat("@");

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

// Single column, select every row from 2..n in column B
// To select a limited range put the ending row number (“B2:B10”)
var column = sheet.getRange("B2:B");
 
// Simple date format
column.setNumberFormat("M/d/yy");

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

// Select all rows and columns from a spreadsheet
var range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
 
// Money format
range.setNumberFormat("$#,##0.00;$(#,##0.00)");

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

// Number format with color
var row = sheet.getRange("A1:D1");
row.setNumberFormat('$#,##0.00[green];$(#,##0.00)[red];"zero";@[blue]');
 
// Sample values
var values = [
   [ "100.0", "-100.0", "0", "ABC" ],
];
row.setValues(values);

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
INIZIO

//console.log(sheet);

VALIDA_OL("Note");
VALIDA_OL("cognomeDSReggente");
VALIDA_OL("nomeDSReggente");
VALIDA_OL("Comune");
VALIDA_OL("Provincia");
VALIDA_OL("Indirizzo");


function VALIDA_OL(rangeName) {

  //Definire preventivamente i Range su cui applicare la validazione
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var sheet = ss.getSheets()[0];
  var sheet = ss.getSheetByName("Sheet1");


  var drng = sheet.getDataRange();
  var rng = sheet.getRange(rangeName);
  
  //Get Column index of the Range selected
  var col_index = rng.getColumn()
  console.log(col_index);

  //notes.setNumberFormat("@[red]");
  var rngA = rng.getValues();
  //console.log(values[0][0].toLowerCase());
  var notes_length = rng.getValues().length;
  console.log(notes_length);
  console.log(rngA[0]);

  let i = 0;

  var format = /[\*'\.,\\\/\(\)"°:;]/gm;
  //var special = false;


  while (i <= notes_length - 1){
    
    var value = rngA[i].toString();
      
    console.log(value);

    if(value.match(format) ){
      console.log("Trovato un carattere speciale");
      //sheet.getRange("G".concat((i+2).toString())).setNumberFormat("@[red]");
      sheet.getRange(i+2, col_index).setNumberFormat("@[red]");
    }
    else{
      sheet.getRange("G".concat((i+2).toString())).setNumberFormat("@[black]");
    }
 
    i++;

  }
  
  //console.log(notes);
  
}

function isEmpty(str) {
    return (!str || str.length === 0 );
}




