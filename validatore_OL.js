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
      .addItem('3.TEST','test')
      .addItem('2.VALIDA_INPUT', 'EX_VALIDATION')
      .addItem('1.CHECK_EXCEPTION', 'EX_CHECKER')
      .addItem('0.UPDATE_DATA_RANGE','updateDataRangeIndex')
      .addToUi();
}


function EX_VALIDATION() {

  SS.active.toast("Inizio Validazione OL…");

  //Aggiorna campo "Note" (Descrizione Lavori) con formule per OL LTE o OL NON LTE
  insertFormula();

  VALIDA_OL("Note");
  VALIDA_OL("Provincia");
  VALIDA_OL("Comune");
  VALIDA_OL("miurDenominazioneScuola");
  VALIDA_OL("Indirizzo");
  VALIDA_OL("nomeDSReggente");
  VALIDA_OL("cognomeDSReggente");

  let PARAMETRI = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parametri");

  //Aggiorna campo "DataConcordata" per emissione OL: 2 mesi dopo emissione
  updataDataConcordata();

//*END FUNCTION EX_VALIDATION*//  
}

function updataDataConcordata() {
    let today = new Date()
    console.log(today.getDate())
    let dd = String(today.getDate()).padStart(2, '0');
    let mm = String(today.getMonth() + 3).padStart(2, '0'); //January is 0!
    let yyyy = today.getFullYear();

    today = dd + '/' + mm + '/' + yyyy;
    
    PARAMETRI.getRange(2,1).setValue(today);

//*END FUNCTION updateDataConcordata*//  
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

  let counter = 0;
  
  let format = /[\*'\.,\\\/\(\)"°:;^\?\!]/gm;

  for (let i = 0; i <= notes_length - 1; i++) { 
    
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

  }

  //console.log("Totale eccezioni individuate: ", counter);

//*END FUNCTION  VALIDA_OL*//    
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
        
  if (button === SS.ui.Button.YES) {
    // User clicked "yes".
    SS.active.toast('Nuovo valore: ' + fixed_value);
    return [fixed_value, true];
  } else if (button === SS.ui.Button.NO ) {
    SS.active.toast('Nuovo valore: ' + text);
    return [text, true];
  } else if (button === SS.ui.Button.CANCEL){
    SS.active.toast('Nessuna correzione.');
    return [field_value, true];
  }  else if (button === SS.ui.Button.CLOSE){
    SS.active.toast("SCRIPT INTERROTTO");
    return [field_value, false];
  }

//*END FUNCTION EDIT_FIELD*//  
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
  //*END FUNCTION FIX_EXCEPTION*//
}

/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Inserisci Formula "Descrizione Lavori"
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/

function insertFormula(ss) {

  SS.active.toast("Inserimento Formula 'Descrizione Lavori'…");

  let length = SS.active.getLastRow();

  let result_LTE = SS.ui.alert("Lotto con backup LTE?",SS.ui.ButtonSet.YES_NO);

  if(result_LTE === SS.ui.Button.YES){
    
    //Formula LTE
    for(let i = 2; i <= length; i++) {
      console.log("LTE: Inizio inserimento riga #"+i);
      SS.input.getRange(i,7).setValue(`=CONCATENATE(D${i};" - ";MID(BJ${i};1;1); MID(BJ${i};7;7);" - VECCHIO COD ";K${i};" - NUOVO COD ";J${i}; " - TIPO ";R${i};" ";S${i};" - PIAN ";AW${i}; " - REF PS D GENTILE 3357825750 - SCUOLA ";Y${i};" - ";V${i};" REF ";AN${i};" ";AO${i};" - T ";AP${i}; " - BCK LTE")`);

      //console.log("LTE: Inserita riga #"+i);
    }
  }
  else {
    //Formula NO LTE

    for(let i = 2; i <= length; i++) {
      console.log("NO LTE: Inizio inserimento riga #"+i);
      SS.input.getRange(i, 7).setValue(`=CONCATENATE(D${i};" - ";MID(BJ${i};1;1); MID(BJ${i};7;7);" - VECCHIO COD ";K${i};" - NUOVO COD ";J${i}; " - TIPO ";R${i};" ";S${i};" - PIAN ";AW${i};" - REF PS DANIELE GENTILE 3357825750 - SCUOLA ";Y${i};" - ";V${i};" REF ";AN${i};" ";AO${i};" - T ";AP${i})`);


      //console.log("NO LTE: Inserita riga #"+i);
    }
  }




//*END FUNCTION insertFormula*//
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
  NMUCHECK("nmuSFP");
  CONTACTCHECK("telefonoScuola");
  CONTACTCHECK("prefisso");
  CONTACTCHECK("tel");
  CONTACTCHECK("mail");
}

function CHECK_OL(rangeName) {
  
  //let drng = SS.input.getDataRange();
  let rng = SS.input.getRange(rangeName);
  
  //Get Column index of the Range selected
  let col_index = rng.getColumn()

  //Get Column values from selected range
  let rngA = rng.getValues();

  var col_length = rng.getValues().length;


  
  let counter = 0;
  
  let format = /[\*'\.,\\\/\(\)"°:;^\?\!]/gm;

  for (let i = 0; i <= col_length - 1; i++){
    
    let value = rngA[i].toString();
      

    if(value.match(format) ){
      
      SS.input.getRange(i+2, col_index).setNumberFormat("@[red]");
      
      counter ++;
    }
    else{
      SS.input.getRange(i+2, col_index).setNumberFormat("@[black]");
    }

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
    "nomeDSReggente":"AO",
    "nmuSFP":"BM",
    "telefonoScuola": "AP",
    "prefisso": "AQ",
    "tel": "AR",
    "mail": "AS"
  }

  for (let x in campi) {
    //console.log(x);
    //console.log(campi[x] + "2: "+ campi[x]+length);
    SS.active.setNamedRange(x, SS.input.getRange(campi[x] + "2:"+ campi[x]+length));
  }  

}

function updateDataConcordata() {
  ss

}

/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Check NMU SFP
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/
function NMUCHECK(rangeName) {

  let rng = SS.input.getRange(rangeName);
  
  //Get Column index of the Range selected
  let col_index = rng.getColumn()

  //Get Column values from selected range
  let rngA = rng.getValues();

  var col_length = rng.getValues().length;


  
  let counter = 0;
  
  //let format = /[\*'\.,\\\/\(\)"°:;^\?\!]/gm;

  for (let i = 0; i <= col_length - 1; i++){
    
    let value = rngA[i].toString();
    console.log(value);
    console.log(i+2);
      
    if(value==="780110" || value==="780111"){
      
      SS.input.getRange(i+2, col_index).setNumberFormat("@[black]");
      
    }
    else{
      SS.input.getRange(i+2, col_index).setNumberFormat("@[red]");
      counter ++;
    }

  }
  SS.active.toast(`Trovati ${counter} campi con caratteri speciali in: ${rangeName}`);
  //console.log("Totale eccezioni individuate: ", counter);

}


/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Check Phone Check
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/

function CONTACTCHECK(rangeName) {

  let rng = SS.input.getRange(rangeName);
  
  //Get Column index of the Range selected
  let col_index = rng.getColumn()

  //Get Column values from selected range
  let rngA = rng.getValues();

  var col_length = rng.getValues().length;


  
  let counter = 0;
  let format

  switch(rangeName) {
    case "mail":
      format = /([a-zA-Z0-9]*)@.*/gm;
      break;
    default:
      format = /([0-9]{2,10})/gm;
  }
   

  for (let i = 0; i <= col_length - 1; i++){
    
    let value = rngA[i].toString();
      
    if(value.match(format) ){
      
      SS.input.getRange(i+2, col_index).setNumberFormat("@[black]");
      
    }
    else{
      SS.input.getRange(i+2, col_index).setNumberFormat("@[red]");
      counter ++;
    }

  }
  SS.active.toast(`Trovati ${counter} campi con caratteri speciali in: ${rangeName}`);
  //console.log("Totale eccezioni individuate: ", counter);
  

}


function test() {

  let result_BANDA = SS.ui.alert("Lotto con upgrade di banda?",SS.ui.ButtonSet.YES_NO);

  if(result_BANDA === SS.ui.Button.YES){
    let new_banda = SS.ui.prompt("Inserisci Valore di banda");
    SS.ui.alert("Nuova banda inserita: " + new_banda.getResponseText());

    PARAMETRI.getRange(2,2).setValue(new_banda.getResponseText());

  }
}
