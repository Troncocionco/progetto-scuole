/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Crea oggetto Sheet e riferimenti per interfacce
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/
let PARAMETRI = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Parametri");

const SS = {
  ui: SpreadsheetApp.getUi(),
  active: SpreadsheetApp.getActiveSpreadsheet(),
  input: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Input"),
  summary: SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary"),
  skynetMode: PARAMETRI.getRange(2,2).getValue(),
  format: /[\*'\.,\\\/\(\)"°:;^\?\!%&]/gm
}

function onOpen() {

  // Or DocumentApp or SlidesApp or FormApp.
  SS.ui.createMenu('Custom Menu')
      .addItem('3.TEST','test')
      .addItem('2.VALIDA_INPUT', 'EX_VALIDATION')
      .addItem('1.CHECK_EXCEPTION', 'EX_CHECKER')
      .addItem('0.INIT','init')
      .addToUi();  
}
/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Setup (Update data range, wipe text-sheet)
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/
function init() {

  let textSheets = ["FTTO_Text","Giga_Text"];

  updateDataRangeIndex();

  textSheets.forEach(x => sheetWiper(x));

  let skynetMode = SS.ui.alert("Vuoi usare la SkynetMode?", SS.ui.ButtonSet.YES_NO);

  let lteMode = SS.ui.alert("Validazione di un Lotto con Backup LTE?", SS.ui.ButtonSet.YES_NO);

  if (skynetMode === SS.ui.Button.YES) {
    //Modifica valore
    PARAMETRI.getRange(2,2).setValue("TRUE")
  }
  else if (skynetMode === SS.ui.Button.NO) {
    //Modifica valore
    PARAMETRI.getRange(2,2).setValue("FALSE")
  }

  if (lteMode === SS.ui.Button.YES) {
    //Modifica valore
    PARAMETRI.getRange(6,2).setValue("TRUE")
  }
  else if (lteMode === SS.ui.Button.NO) {
    PARAMETRI.getRange(6,2).setValue("FALSE")
  }

}

/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
  ESECUTORI
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/

function EX_VALIDATION() {

  SS.active.toast("Inizio Validazione OL…");

  //Aggiorna campo "Note" (Descrizione Lavori) con formule per OL LTE o OL NON LTE
  insertFormula();

  let campi = ["Note","Provincia", "Comune", "miurDenominazioneScuola", "Indirizzo", "nomeDSReggente", "cognomeDSReggente"];

  campi.forEach(campo => VALIDA_OL(campo));

  //Aggiorna campo "DataConcordata" per emissione OL: 2 mesi dopo emissione
  updataDataConcordata();
  

//*END FUNCTION EX_VALIDATION*//
}

function EX_CHECKER() {

  SS.active.toast("Inizio Controllo OL …");

  let campi = ["Note","Provincia", "Comune", "miurDenominazioneScuola", "Indirizzo", "nomeDSReggente", "cognomeDSReggente"];

  campi.forEach(campo => CHECK_OL(campo));
  
  let campiContatto = ["telefonoScuola", "prefisso", "tel", "mail"];

  campiContatto.forEach(campo => CONTACTCHECK(campo));

  NMUCHECK("nmuSFP");

}
/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Valida OL
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/
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

  let countCorrection = 0;
  
  for (let i = 0; i <= notes_length - 1; i++) { 
    
    let value = rngA[i].toString(); //cell's value
    let keepGoing = true; //controller interrupt script

    //Look for special char
    if(value.match(SS.format) ){

      SS.input.getRange(i+2, col_index).setNumberFormat("@[blue]");

      //Do not change "Note" column
      if(rangeName != "Note"){
              
              //Edit cell with special char
              let r = EDIT_FIELD(value);
              SS.active.toast(r);
              value = r[0];       //Fixed text
              keepGoing = r[1];   //User wants to implement correction?
              
              if (keepGoing){
                SS.input.getRange(i+2, col_index).setValue(value); //Write fixed value
              } else if (!keepGoing){
                break; //Don't edit this field
              }        
      }
      countCorrection ++;
    }
    else{
      
      SS.input.getRange(i+2, col_index).setNumberFormat("@[black]");
    }

  }
  SS.summary.getRange("sum_" + rangeName).setValue(countCorrection);
  //console.log("Totale eccezioni individuate: ", countCorrection);

//*END FUNCTION  VALIDA_OL*//    
}
/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Check Eccezioni nel file OL
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/

function CHECK_OL(rangeName) {
  
  //let drng = SS.input.getDataRange();
  let rng = SS.input.getRange(rangeName);
  
  //Get Column index of the Range selected
  let col_index = rng.getColumn()

  //Get Column values from selected range
  let rngA = rng.getValues();

  var col_length = rng.getValues().length;

  let counter = 0;
  
  for (let i = 0; i <= col_length - 1; i++){
    
    let value = rngA[i].toString();
      

    if(value.match(SS.format) ){
      
      SS.input.getRange(i+2, col_index).setNumberFormat("@[red]");
      SS.input.getRange(i+2, col_index).setBackground("yellow");
      
      counter ++;
    }
    else{
      SS.input.getRange(i+2, col_index).setNumberFormat("@[black]");
      SS.input.getRange(i+2, col_index).setBackground("white");
    }

  }
  SS.active.toast(`Trovati ${counter} campi con caratteri speciali in: ${rangeName}`);
  SS.summary.getRange("sum_" + rangeName).setValue(counter);
  //console.log("Totale eccezioni individuate: ", counter);
  
}
/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Editor field
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/
function EDIT_FIELD(field_value){
   
  //let fixed_value = FIX_EXCEPTION(field_value);

  let fixed_value = field_value.replaceAll("\" "," ")
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
  .replaceAll(","," ")
  .replaceAll("&"," ")
  .replaceAll("%"," ")
  .replaceAll("?"," ");
  
  if ( SS.skynetMode ) {
    return [fixed_value, true];
  }
  else {
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
  }  

//*END FUNCTION EDIT_FIELD*//  
}



//**------------------------------------------------------------------------------------------ */




/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Inserisci Formula "Descrizione Lavori"
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/
function insertFormula(ss) {

  SS.active.toast("Inserimento Formula 'Descrizione Lavori'…");

  let length = SS.active.getLastRow();

  //let result_LTE = SS.ui.alert("Lotto con backup LTE?",SS.ui.ButtonSet.YES_NO);

  //if(result_LTE === SS.ui.Button.YES){
  if (PARAMETRI.getRange(6,2).getValue()) {  
    
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
Sheet Wiper
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/
function sheetWiper(sheetName, dataRange = sheetName) {

  let TT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  let rng = TT.getRange(dataRange);
  
  rng.setValue("");

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
/*%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Updates dei DataConcordata
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%*/
function updataDataConcordata() {
    let today = new Date()
    console.log(today.getDate())
    let dd = String(today.getDate()).padStart(2, '0');
    let mm = String(today.getMonth() + 3).padStart(2, '0'); //January is 0!
    let yyyy = today.getFullYear();

    today = dd + '/' + mm + '/' + yyyy;
    
    PARAMETRI.getRange(1,2).setValue(today);

//*END FUNCTION updateDataConcordata*//  
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

  for (let i = 0; i <= col_length - 1; i++){
    
    let value = rngA[i].toString();
    console.log(value);
    console.log(i+2);
      
    if(value==="780110" || value==="780111"){
      
      SS.input.getRange(i+2, col_index).setNumberFormat("@[black]");
      
    }
    else{
      SS.input.getRange(i+2, col_index).setNumberFormat("@[red]");
      SS.input.getRange(i+2, col_index).setBackground("yellow");
      counter ++;
    }

  }
  SS.active.toast(`Trovati ${counter} campi con caratteri speciali in: ${rangeName}`);
  SS.summary.getRange("sum_" + rangeName).setValue(counter);
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
      SS.input.getRange(i+2, col_index).setBackground("yellow");
      counter ++;
    }

  }
  SS.active.toast(`Trovati ${counter} campi con caratteri speciali in: ${rangeName}`);
  SS.summary.getRange("sum_" + rangeName).setValue(counter);
  //console.log("Totale eccezioni individuate: ", counter);
  
}
