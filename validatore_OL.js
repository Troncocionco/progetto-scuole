/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Crea oggetto Sheet e riferimenti per interfacce
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/

const SS = {
  ui: SpreadsheetApp.getUi(),
  active: SpreadsheetApp.getActiveSpreadsheet(),
  input: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input")
}


function onOpen() {

  // Or DocumentApp or SlidesApp or FormApp.
  SS.ui.createMenu('Custom Menu')
      .addItem('2.VALIDA_INPUT', 'EX_VALIDATION')
      .addItem('1.CHECK_EXCEPTION', 'EX_CHECKER')
      .addItem('0.UPDATE_DATA_RANGE','updateDataRangeIndex')
      .addToUi();
}


function EX_VALIDATION() {


  SS.active.toast("Inizio Validazione OL…");

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

  let drng = SS.input.getDataRange();
  let rng = SS.input.getRange(rangeName);
  
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

      SS.input.getRange(i+2, col_index).setNumberFormat("@[blu]");

      //Do not change "Note" column
      if(rangeName != "Note"){
              
              //Edit cell with special char
              let r = EDIT_FIELD(value);
              SS.active.toast(r);
              value = r[0];
              keepGoing = r[1];
              

              if (keepGoing){
                SS.input.getRange(i+2, col_index).setValue(value);
              } else if (!keepGoing){
                break; //User stopped the script
              }        
      }
      counter ++;
    }
    else{
      //sheet.getRange("A".concat((i+2).toString())).setNumberFormat("@[black]");
      SS.input.getRange(i+2, col_index).setNumberFormat("@[black]");
    }
 
    i++;

  }

  //console.log("Totale eccezioni individuate: ", counter);
  
}


function EDIT_FIELD(field_value){
   

  let fixed_value = FIX_EXCEPTION(field_value);  
  let result = SS.ui.prompt(
          'Trovato un carattere speciale!  Yes: accetta correzione No: Inserisci nuovo valore  Cancel: Non applicare nessuna correzione   X: Interrompi Script',
          field_value + ' --> '+ fixed_value,
          SS.ui.ButtonSet.YES_NO_CANCEL);
    
  // Process the user's response.
  let button = result.getSelectedButton();
  let text = result.getResponseText();
        
  if (button == SS.ui.Button.YES) {
    // User clicked "yes".
    SS.active.toast('Nuovo valore: ' + fixed_value);
    return [fixed_value, true];
  } else if (button == SS.ui.Button.NO ) {
    SS.active.toast('Nuovo valore: ' + text);
    return [text, true];
  } else if (button == SS.ui.Button.CANCEL){
    SS.active.toast('Nessuna correzione.');
    return [field_value, true];
  }  else if (button == SS.ui.Button.CLOSE){
    SS.active.toast("SCRIPT INTERROTTO");
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

function insertFormula(ss) {

  
  SS.active.toast("Inserimento Formula 'Descrizione Lavori'…");


  let length = SS.active.getLastRow();

  let i = 2;

  let result = SS.ui.alert("Lotto con backup LTE?",SS.ui.ButtonSet.YES_NO);


  if(result == SS.ui.Button.YES){
    
    //Formula LTE
    while (i <= length ) {
      console.log("LTE: Inizio inserimento riga #"+i);
      SS.input.getRange(i,7).setValue(`=CONCATENATE(D${i};" - ";MID(BJ${i};1;1); MID(BJ${i};7;7);" - VECCHIO COD ";K${i};" - NUOVO COD ";J${i}; " - TIPO ";R${i};" ";S${i};" - PIAN ";AW${i}; " - REF PS D GENTILE 3357825750 - SCUOLA ";Y${i};" - ";V${i};" REF ";AN${i};" ";AO${i};" - T ";AP${i}; " - BCK LTE")`);

      console.log("LTE: Inserita riga #"+i);
      i++;
    }

  }
  
  else {
    //Formula NO LTE
    while (i <= length ) {
      console.log("NO LTE: Inizio inserimento riga #"+i);
      SS.input.getRange(i, 7).setValue(`=CONCATENATE(D${i};" - ";MID(BJ${i};1;1); MID(BJ${i};7;7);" - VECCHIO COD ";K${i};" - NUOVO COD ";J${i}; " - TIPO ";R${i};" ";S${i};" - PIAN ";AW${i};" - REF PS DANIELE GENTILE 3357825750 - SCUOLA ";Y${i};" - ";V${i};" REF ";AN${i};" ";AO${i};" - T ";AP${i}2)`);


      console.log("NO LTE: Inserita riga #"+i);
      i++;
    }

  }
  
}

/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Check Eccezioni nel file OL
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/

function EX_CHECKER() {

  SS.active.toast("Inizio Controllo OL …");

  CHECK_OL("Note");
  CHECK_OL("Provincia");
  CHECK_OL("Comune");
  CHECK_OL("miurDenominazioneScuola");
  CHECK_OL("Indirizzo");
  CHECK_OL("nomeDSReggente");
  CHECK_OL("cognomeDSReggente");
  

}

function CHECK_OL(rangeName) {
  
  let drng = SS.input.getDataRange();
  var rng = SS.input.getRange(rangeName);
  
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
      
      SS.input.getRange(i+2, col_index).setNumberFormat("@[red]");
      
      counter ++;
    }
    else{
      SS.input.getRange(i+2, col_index).setNumberFormat("@[black]");
    }
 
    i++;

  }
  SS.active.toast(`Trovati ${counter} campi con caratteri speciali in: ${rangeName}`);
  //console.log("Totale eccezioni individuate: ", counter);
  
}

/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Updates dei DataRange
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/
function updateDataRangeIndex() {

  let length = SS.input.getLastRow();
  let width = SS.input.getLastColumn();

  const campi = {
    "Note": "G",
    "Provincia": "U",
    "Comune":"V",
    "miurDenominazioneScuola":"Y",
    "Indirizzo":"AC",
    "cognomeDSReggente":"AN",
    "nomeDSReggente":"AO"
  }

  for (let x in campi) {
    //console.log(x);
    //console.log(campi[x] + "2: "+ campi[x]+length);
    SS.input.setNamedRange(x, SS.input.getRange(campi[x] + "2:"+ campi[x]+length));
  }  

}