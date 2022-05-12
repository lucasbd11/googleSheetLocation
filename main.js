var date1Column = 8;
var date2Column = 9;
var verifColumn = 10;
var matosColumn = 7;
var prixColumn = 11;
var cautionColumn = 12;
var locaDATA = SpreadsheetApp.getActive().getSheetByName('Prix/Caution');
var SPLITCHAR = "\n";

function creationLst(){
  var row = 2;

  var lstMatos = [];
  var lstPrix = [];
  var lstCaution = [];

  while (locaDATA.getRange(row,1).getValue() != ""){
    lstMatos.push(locaDATA.getRange(row,1).getValue());
    lstPrix.push(locaDATA.getRange(row,2).getValue());
    lstCaution.push(locaDATA.getRange(row,3).getValue());
    row++;
  

  }

  return [lstMatos,lstPrix,lstCaution];



}

function onEdit(e) {
  var ss=SpreadsheetApp.getActiveSheet()

  var activeCell = ss.getActiveCell();
  
  
  var lstInfo = creationLst();
  console.log(lstInfo);

  var col = activeCell.getColumn();
  var row = activeCell.getRow();


  if ((ss.getRange(1,col).getValue()) == "Séléction matériel" && (row != 1)){
    if ((ss.getRange(row,matosColumn).getValue()).split(SPLITCHAR).includes(activeCell.getValue())){
      
      var lstMatosCell = ss.getRange(row,matosColumn).getValue().split(SPLITCHAR);

      var index = lstMatosCell.indexOf(activeCell.getValue());
      if (index !== -1) {
        lstMatosCell.splice(index, 1);
      } 
      
      if (lstMatosCell[0] != ""){
        ss.getRange(row,matosColumn).setValue(lstMatosCell.join(SPLITCHAR));
      } else {
        ss.getRange(row,matosColumn).setValue(lstMatosCell[1]);
      }

    }else {

      var lstMatosCell = ss.getRange(row,matosColumn).getValue().split(SPLITCHAR);
      lstMatosCell.push(activeCell.getValue());
      console.log(lstMatosCell);
      console.log("log ok");
      if (lstMatosCell[0] != ""){
        ss.getRange(row,matosColumn).setValue(lstMatosCell.join(SPLITCHAR));
      } else {
        ss.getRange(row,matosColumn).setValue(lstMatosCell[1]);
      }


    }

    col = matosColumn;
    activeCell.setValue("");

    

  }
  


  //if ((listMatos.includes(ss.getRange(1,col).getValue())) && (row != 1)){
  if ((ss.getRange(1,col).getValue()) == "Matériel" && (row != 1)){  
    ss.getRange(row,verifColumn).setValue("Vérification...");
    const dateTEST1 = new Date(ss.getRange(row,date1Column).getValue());
    const dateTEST2 = new Date(ss.getRange(row,date2Column).getValue());

    var CHECK = 1;
    var lstMatos = ss.getRange(row,col).getValue().split(SPLITCHAR);
    console.log(lstMatos);
    for (let j = 0; j < lstMatos.length; j++){
      var matos = lstMatos[j];
      console.log(matos);
      for (let i = 2; i < row; i++) {
        
        //if (ss.getRange(i,col).getValue() == true){
        if ((ss.getRange(i,col).getValue().split(SPLITCHAR)).includes(matos)){

          var dateStr1 = ss.getRange(i,date1Column).getValue();
          var dateStr2 = ss.getRange(i,date2Column).getValue();
          const date1 = new Date(dateStr1);
          const date2 = new Date(dateStr2);
          
          console.log(date1.getTime() < dateTEST2.getTime());
          console.log(date1.getTime());
          console.log(dateTEST2.getTime());
          if (((date1.getTime() < dateTEST1.getTime()) && (dateTEST1.getTime() < date2.getTime()))){
            CHECK = 0;
          }
          if ((date1.getTime() < dateTEST2.getTime()) && (dateTEST2.getTime() < date2.getTime())){
            CHECK = 0;
          }
        }

      }
    }
    
    if (CHECK == 1){
      ss.getRange(row,verifColumn).setValue("OK");
      ss.getRange(row,verifColumn).setBackgroundRGB(255,255,255);
    } else {
      ss.getRange(row,verifColumn).setValue("PAS OK");
      ss.getRange(row,verifColumn).setBackgroundRGB(255,0,0);
    }
    
    
    var cell = SpreadsheetApp.getActive().getRange('A1');
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No']).build();
    cell.setDataValidation(rule);








  }

    console.log("OK");

}



function getCellName(col,row){
  var alpha = " ABCDEFGHIJKLMNOPQRSTUVWXYZ";

  var cellID;

  var reste = col%26;

  if (col>26){
    cellID = "A"+alpha[reste];
  } else {
    cellID = alpha[col];
  }

  cellID = cellID + row.toString();

  return cellID;
}
  
