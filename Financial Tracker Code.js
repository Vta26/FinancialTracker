//---Create New Trigger---
function createOnEditTrigger(TriggerName) {
  var triggers = ScriptApp.getProjectTriggers();
  var shouldCreateTrigger = true;
  triggers.forEach(function (trigger) {
    if(trigger.getEventType() === ScriptApp.EventType.ON_EDIT && trigger.getHandlerFunction() === TriggerName) {
      shouldCreateTrigger = false; 
    }
  });
  
  if(shouldCreateTrigger) {
    ScriptApp.newTrigger(TriggerName)
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
  }
}

//---Authorize New Users---
function Authorize() {
  var ui = SpreadsheetApp.getUi();

  var result = ui.alert('Thanks for authorizing :)')

  createOnEditTrigger("EditExpenses")

  createOnEditTrigger("EditGiftCard")
}

//---Expense Functions---

function EditExpenses(e){
  var range = e.range;
  var value = range.getValue();
  if (value == "" || value == "FALSE"){
    return;
  }
  var col = range.getColumn();
  var row = range.getRow();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[1];
  if (ss.getActiveSheet().getName() != sheet.getName()){
    return;
  }
  if (col == 1.0 && row == 8){//Add new Expense
    var lastcell = (sheet.getRange(4,5).getValue());
    if(lastcell == 0){
      sheet.getRange(4,3).setValue(value);
      sheet.getRange(4,4).setValue(sheet.getRange(7,1).getValue());
      range.clearContent();
      return;
    }
    r1 = sheet.getRange(lastcell+4,3);
    r1.insertCells(SpreadsheetApp.Dimension.ROWS);
    r1.setValue(value);
    r2 = sheet.getRange(lastcell+4,4);
    r2.insertCells(SpreadsheetApp.Dimension.ROWS);
    r2.setValue(sheet.getRange(7,1).getValue());
    range.clearContent();
    return;
  }
  if (col == 2 && row == 16){//Remove Last Inputted Expense
    var lastcell = sheet.getRange(5,5).getValue();
    if (lastcell<2){
      sheet.getRange(4,3,2).clearContent();
      sheet.getRange(4,4,2).clearContent();
      return;
    }
    sheet.getRange(3+lastcell,3).deleteCells(SpreadsheetApp.Dimension.ROWS);
    sheet.getRange(3+lastcell,4).deleteCells(SpreadsheetApp.Dimension.ROWS);
  }
  if (col == 2 && row == 17){//Remove Last Non-Numerical Value
    var lastcell = sheet.getRange(4,5).getValue();
    if (lastcell<2){
      sheet.getRange(4,3,2).clearContent();
      sheet.getRange(4,4,2).clearContent();
      return;
    }
    var check = sheet.getRange(5,5).getValue();
    if(lastcell != check){
      sheet.getRange(4+lastcell,3).deleteCells(SpreadsheetApp.Dimension.ROWS);
      sheet.getRange(4+lastcell,4).deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
  }
  if (col == 2 && row == 18){//Reset
    var history = ss.getSheets()[2];
    var lastcell = sheet.getRange(4,5).getValue();
    var currentdate = new Date();
    var range3 = sheet.getRange(14,1);
    var range3val = range3.getValue();
    history.getRange(1,2,35).insertCells(SpreadsheetApp.Dimension.COLUMNS).insertCells(SpreadsheetApp.Dimension.COLUMNS);
    history.getRange(1,2).setValue(history.getRange(2,4).getValue());
    history.getRange(2,2).setValue(currentdate);
    history.getRange(31,2).setValue("=SUM(B3:B29)");
    history.getRange(33,2).setValue(sheet.getRange(4,1).getValue());
    history.getRange(35,2).setValue("=B33-B31");
    sheet.getRange(4,3,lastcell).copyTo(history.getRange(3,2));
    sheet.getRange(4,4,lastcell).copyTo(history.getRange(3,3));
    for(i=0;i<7;i++){//update category calc
      var amtspent = history.getRange(38,2+i);
      var categorytype = history.getRange(37,2+i).getA1Notation();
      amtspent.setValue(amtspent.getFormula() + "+SUMIF($C$3:$C$29,"+categorytype+",$B$3:$B$29)");
    }
    if(range3val<0){//Overspent
      sheet.getRange(4,7).setValue(Math.abs(range3val));
    }
    else{
      sheet.getRange(4,7).setValue(0);
    }
    if(lastcell>=2){
      var range = sheet.getRange(6,3,lastcell-1);
      var range3 = sheet.getRange(6,4,lastcell-1);
      range.deleteCells(SpreadsheetApp.Dimension.ROWS);
      range3.deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
    var range2 = sheet.getRange(4,3,2);
    range2.clearContent();
    sheet.getRange(4,4,2).clearContent();
    return;
  }
}

//---GiftCard Functions---

function SuccessClear(CurrentSheet){
  CurrentSheet.getRange(1,3,4,1).clearContent();
}

function DeleteGiftCard(CurrentSheet,OtherCheck, LastRow){
  //Cannot Delete if Other is Selected
  if (OtherCheck){
    CurrentSheet.getRange(1,3).setValue("Other is Selected, Cannot Delete");
    return;
  }
  SuccessClear(CurrentSheet);
  var colnum = CurrentSheet.getRange(1,1).getValue();
  CurrentSheet.getRange(1,colnum-1,LastRow,3).deleteCells(SpreadsheetApp.Dimension.COLUMNS);
  var rownum = CurrentSheet.getRange(9,1).getValue();
  CurrentSheet.getRange(rownum,1,1,2).deleteCells(SpreadsheetApp.Dimension.ROWS);
}

function SetUp(CurrentSheet){
  CurrentSheet.getRange('G2:I8').copyTo(CurrentSheet.getRange(2,4));
  CurrentSheet.getRange(2,5).setValue(CurrentSheet.getRange(2,2).getValue());
  CurrentSheet.getRange(4,5).setValue('='+ CurrentSheet.getRange(4,2).getValue() + ' - SUM(D7:D8)');//In case there's more than 2 entries when copied over
  CurrentSheet.getRange('D7:E8').clearContent();
}

function EditGiftCard(e) {
  //Grabbing Edited Cell
  const range = e.range;
  const value = range.getValue();
  if (value == "" || value == "FALSE"){
    return;
  }
  const col = range.getColumn();
  const row = range.getRow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  //Grabbing Current Sheet and Checking If Gift Card Sheet
  const sheet = ss.getSheets()[3];
  if (ss.getActiveSheet().getName() != sheet.getName()){
    return;
  }

  //Check if Other is Selected
  var CurrentCompany = sheet.getRange(1,2).getValue();
  var IsOther = false;
  if (CurrentCompany == "Other"){
    IsOther = true;
  }

  //Add Value
  if (col == 2 && row == 6){
    if (IsOther){//Add new Entry if new company name is given
      if (sheet.getRange(3,1).getValue()){
        if (sheet.getRange(2,2).getValue()==""){
          sheet.getRange(2,3).setValue("Cannot Add, No Name Given");
        }
        else{
          sheet.getRange(4,3).setValue("Cannot Add, No Starter Value Given");
        }
        return;
      }
      sheet.getRange('D1:F').insertCells(SpreadsheetApp.Dimension.COLUMNS);
      SetUp(sheet);
      sheet.getRange('A12:B12').insertCells(SpreadsheetApp.Dimension.ROWS);
      sheet.getRange(12,1).setValue(sheet.getRange(2,2).getValue());
      sheet.getRange(12,2).setValue('=IFNA(XLOOKUP(A12,C2:M2,C4:M4),"")');
      sheet.getRange(1,2).setValue(sheet.getRange(2,2).getValue());
      sheet.getRange(2,2).clearContent();
    }
    else{//else just add transaction
      var colnum = sheet.getRange(1,1).getValue();
      var rownum = sheet.getRange(5,colnum - 1).getValue();
      var transactiondate = new Date();
      if(rownum == 7){
        sheet.getRange(7,colnum - 1).setValue(sheet.getRange(4,2).getValue());
        sheet.getRange(7,colnum).setValue(transactiondate);
        SuccessClear(sheet);
        return;
      }
      sheet.getRange(rownum,colnum).insertCells(SpreadsheetApp.Dimension.ROWS).setValue(transactiondate);
      sheet.getRange(rownum,colnum - 1).insertCells(SpreadsheetApp.Dimension.ROWS).setValue(sheet.getRange(4,2).getValue());
    }
    SuccessClear(sheet);
  }

  //Delete Gift Card
  if (col == 2 && row == 8){
    DeleteGiftCard(sheet,IsOther,sheet.getLastRow());
    sheet.getRange(1,2).setValue("Other");
  }
}