
function onOpen() {
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Custom Menu')
      .addItem('2.VALIDA_INPUT', 'EX_VALIDATION')
      .addItem('1.CHECK_EXCEPTION', 'EX_CHECKER')
      .addItem('0.UPDATE_DATA_RANGE','updateDataRangeIndex')
      .addToUi();
}

function EX_VALIDATION() {

  let ui = SpreadsheetApp.getUi(); // Same variations.
  spreadS = SpreadsheetApp.getActive()
  spreadS.toast("Inizio Validazione OL…");

  insertFormula();

  VALIDA_OL("Note");
  VALIDA_OL("Provincia");
  VALIDA_OL("Comune");
  VALIDA_OL("miurDenominazioneScuola");
  VALIDA_OL("Indirizzo");
  VALIDA_OL("nomeDSReggente");
  VALIDA_OL("cognomeDSReggente");
  
}

function VALIDA_OL(rangeName) {

  //Define previously Data Range in G Sheet
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi();
  //let sheet = ss.getSheets()[0];
  let sheet = ss.getSheetByName("Input");

  let drng = sheet.getDataRange();
  let rng = sheet.getRange(rangeName);
  
  //Get Column index of the Range selected
  let col_index = rng.getColumn()
  console.log("Indice della colonna: ", col_index);


  let rngA = rng.getValues();

  let notes_length = rng.getValues().length;
  console.log(notes_length);
  console.log(rngA[0]);

  let i = 0;
  let counter = 0;
  
  let format = /[\*'\.,\\\/\(\)"°:;^\?\!]/gm;


  while (i <= notes_length - 1){
    
    let value = rngA[i].toString(); //cell's value
    let keepGoing = true; //controller interrupt script

    //Look for special char
    if(value.match(format) ){

      sheet.getRange(i+2, col_index).setNumberFormat("@[blu]");

      //Do not change "Note" column
      if(rangeName != "Note"){
              
              //Edit cell with special char
              let r = EDIT_FIELD(value);
              spreadS.toast(r);
              value = r[0];
              keepGoing = r[1];
              

              if (keepGoing){
                sheet.getRange(i+2, col_index).setValue(value);
              } else if (!keepGoing){
                break; //User stopped the script
              }        
      }
      counter ++;
    }
    else{
      //sheet.getRange("A".concat((i+2).toString())).setNumberFormat("@[black]");
      sheet.getRange(i+2, col_index).setNumberFormat("@[black]");
    }
 
    i++;

  }

  //console.log("Totale eccezioni individuate: ", counter);
  
}


function EDIT_FIELD(field_value){
   
  let ui = SpreadsheetApp.getUi(); // Same variations.
  let fixed_value = FIX_EXCEPTION(field_value);  
  let result = ui.prompt(
          'Trovato un carattere speciale!  Yes: accetta correzione No: Inserisci nuovo valore  Cancel: Non applicare nessuna correzione   X: Interrompi Script',
          field_value + ' --> '+ fixed_value,
          ui.ButtonSet.YES_NO_CANCEL);
    
  // Process the user's response.
  let button = result.getSelectedButton();
  let text = result.getResponseText();
        
  if (button == ui.Button.YES) {
    // User clicked "yes".
    spreadS.toast('Nuovo valore: ' + fixed_value);
    return [fixed_value, true];
  } else if (button == ui.Button.NO ) {
    spreadS.toast('Nuovo valore: ' + text);
    return [text, true];
  } else if (button == ui.Button.CANCEL){
    spreadS.toast('Nessuna correzione.');
    return [field_value, true];
  }  else if (button == ui.Button.CLOSE){
    spreadS.toast("SCRIPT INTERROTTO");
    return [field_value, false];
  }

}

function FIX_EXCEPTION(field_value) {
  let output = field_value.replaceAll("\" "," ")
  .replaceAll(" \""," ")
  .replaceAll('"',"")
  .replaceAll('('," ")
  .replaceAll(":"," ")
  .replaceAll(')'," ")
  .replaceAll("\*", "")
  .replaceAll("'"," ")
  .replaceAll("/"," ")
  .replaceAll(". "," ")
  .replaceAll("."," ")
  .replaceAll("?"," ");

  return output;
}
/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Inserisci Formula "Descrizione Lavori"
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/

function insertFormula() {

  let spreadS = SpreadsheetApp.getActive()
  spreadS.toast("Inserimento Formula 'Descrizione Lavori'…");

  let ui = SpreadsheetApp.getUi();
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  let inputSheet = ss.getSheetByName("Input");

  let length = ss.getLastRow();

  let i = 2;

  let result = ui.alert("Lotto con backup LTE?",ui.ButtonSet.YES_NO);


  if(result == ui.Button.YES){
    
    //Formula LTE
    while (i <= length ) {
      console.log("LTE: Inizio inserimento riga #"+i);
      inputSheet.getRange(i,7).setValue(`=CONCATENATE(D${i};" - ";MID(BJ${i};1;1); MID(BJ${i};7;7);" - VECCHIO COD ";K${i};" - NUOVO COD ";J${i}; " - TIPO ";R${i};" ";S${i};" - PIAN ";AW${i}; " - REF PS D GENTILE 3357825750 - SCUOLA ";Y${i};" - ";V${i};" REF ";AN${i};" ";AO${i};" - T ";AP${i}; " - BCK LTE")`);

      console.log("LTE: Inserita riga #"+i);
      i++;
    }

  }
  
  else {
    //Formula NO LTE
    while (i <= length ) {
      console.log("NO LTE: Inizio inserimento riga #"+i);
      inputSheet.getRange(i, 7).setValue(`=CONCATENATE(D${i};" - ";MID(BJ${i};1;1); MID(BJ${i};7;7);" - VECCHIO COD ";K${i};" - NUOVO COD ";J${i}; " - TIPO ";R${i};" ";S${i};" - PIAN ";AW${i};" - REF PS DANIELE GENTILE 3357825750 - SCUOLA ";Y${i};" - ";V${i};" REF ";AN${i};" ";AO${i};" - T ";AP${i}2)`);


      console.log("NO LTE: Inserita riga #"+i);
      i++;
    }

  }
  
}

/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Check Eccezioni nel file OL
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/

function EX_CHECKER() {

  let ui = SpreadsheetApp.getUi(); // Same variations.
  let spreadS = SpreadsheetApp.getActive()
  spreadS.toast("Inizio Controllo OL …");

  CHECK_OL("Note");
  CHECK_OL("Provincia");
  CHECK_OL("Comune");
  CHECK_OL("miurDenominazioneScuola");
  CHECK_OL("Indirizzo");
  CHECK_OL("nomeDSReggente");
  CHECK_OL("cognomeDSReggente");
  

}

function CHECK_OL(rangeName) {

  //Definire preventivamente i Range su cui applicare la validazione
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi();
  let sheet = ss.getSheetByName("Input");

  let drng = sheet.getDataRange();
  var rng = sheet.getRange(rangeName);
  
  //Get Column index of the Range selected
  var col_index = rng.getColumn()

  //Get Column values from selected range
  var rngA = rng.getValues();

  var notes_length = rng.getValues().length;


  let i = 0;
  let counter = 0;
  
  var format = /[\*'\.,\\\/\(\)"°:;^\?\!]/gm;

  while (i <= notes_length - 1){
    
    var value = rngA[i].toString();
      

    if(value.match(format) ){
      
      sheet.getRange(i+2, col_index).setNumberFormat("@[red]");
      
      counter ++;
    }
    else{
      sheet.getRange(i+2, col_index).setNumberFormat("@[black]");
    }
 
    i++;

  }
  ss.toast(`Trovati ${counter} campi con caratteri speciali in: ${rangeName}`);
  //console.log("Totale eccezioni individuate: ", counter);
  
}

/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Updates dei DataRange
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/
function updateDataRangeIndex() {
  let ui = SpreadsheetApp.getUi(); // Same variations.
  let ss = SpreadsheetApp.getActiveSpreadsheet();

  let inputSheet = ss.getSheetByName("Input");

  let length = ss.getLastRow();
  let width = ss.getLastColumn();

  const campi = {
    "Note": "G",
    "Provincia": "U",
    "Comune":"V",
    "miurDenominazioneScuola":"Y",
    "Indirizzo":"AC",
    "cognomeDSReggente":"AN",
    "nomeDSReggente":"AO"
  }
  
  
  //console.log(campi);
  //console.log(length, width);
  
  //ss.setNamedRange("Pippo",ss.getRange("B2:B50"));
  //console.log(`${campi["Note"]}2: ${campi["Note"]}${length}`);

  for (let x in campi) {
    //console.log(x);
    //console.log(campi[x] + "2: "+ campi[x]+length);
    ss.setNamedRange(x, ss.getRange(campi[x] + "2:"+ campi[x]+length));
  }  

}