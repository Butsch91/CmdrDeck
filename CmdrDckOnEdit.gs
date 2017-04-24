// CmdrDckOnEdit

// **********************************************
// function fcnOnEdit()
//
// Executes a few functions if some cells are edited
// 
// 
// **********************************************

function fcnOnEdit() {
  
  // Gets Sheet Data
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();								
  var actShtName = actSht.getSheetName();
  var MaxCol = actSht.getMaxColumns();
  
  // Gets Cell Data
  var CellRng = actSht.getActiveCell();
  var CellRow = CellRng.getRow();
  var CellCol = CellRng.getColumn();
  var Value = CellRng.getValue();
  
  // Opens the Configuration Sheet and Gets the Deck and Sideboard Ranges
  var ConfigSht = ss.getSheetByName('Config');
  var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
  
  // Parameters
  var LinkCol = 14;
  var LinkRng = actSht.getRange(CellRow, LinkCol);
  var ClrEnable = actSht.getRange(1,13).getValue();
  var StatusRow = 2;
  var StatusCol = 3;
  var CardTypeCol = 7;
  var CmdRow = 1;
  var CmdCol = 9;
  var GenDeckRow = 2;
  var GenDeckCol = 9;
    
  // CHECKS IF THE SELECTED SHEET IS VALID
  if(actShtName != 'Template' &&  actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef'){
    
    // IF THE MODIFIED CELL IS THE SORT COMMAND CELL, EXECUTES SORT CARDS FUNCTION
    if (CellRow == CmdRow && CellCol == CmdCol && Value != ''){
      fcnSortCardsCmd(Value);
    }
    
    // IF THE MODIFIED CELL IS THE GENERATE NEW DECK VERSION COMMAND CELL, EXECUTES THE GENERATE NEW DECK VERSION FUNCTION
    if (CellRow == GenDeckRow && CellCol == GenDeckCol && Value == 'Generate New Deck'){
      fcnGenerateNewVersion();
    }
    
    // IF A CARD IS ADDED, INSERT GATHERER LINK IN APPROPRIATE COLUMN
    if (CellRow >= DeckFirstRow - 1 && CellCol == 3 && Value != '' ) { 
      var CardName = Value;
      LinkRng.setValue('=HYPERLINK("http://gatherer.wizards.com/Pages/Card/Details.aspx?name='+CardName+'","'+CardName+'")');
    }
    
    // IF THE CARD ADDED WAS A LAND, LAND CATEGORY AND LAND COLOR IS AUTOMATICALLY ADDED. IF ARTIFACT, COLORLESS IS AUTOMATICALLY ADDED
    if (CellRow >= DeckFirstRow && CellCol == CardTypeCol && Value != ''){
      var CardType = Value;
      var CategoryCol = 9;
      var CardColorCol = 11;
      
      if (CardType == 'Land' || CardType == 'Basic Land'){
        actSht.getRange(CellRow,CategoryCol).setValue('Land');
        actSht.getRange(CellRow,CardColorCol).setValue('L');
      }
      
      if (CardType == "Artifact"){
        actSht.getRange(CellRow,CardColorCol).setValue('C');
      }
    }
    
    // IF CARD NAME WAS CLEARED, DELETES ALL INFO IF CLEAR IS ENABLED  
    if (CellRow >= DeckFirstRow && CellCol == 3 && Value == '' && ClrEnable == 'Enable Clear') { 
      actSht.getRange(CellRow,4).setValue("");
      actSht.getRange(CellRow,7).setValue("");
      actSht.getRange(CellRow,9).setValue("");
      actSht.getRange(CellRow,11).setValue("");
      actSht.getRange(CellRow,13).setValue("");
      actSht.getRange(CellRow,14).setValue("");
      
      // Executes the Sort Command
      var SortCmd = actSht.getRange(CmdRow, CmdCol).getValue();
      fcnSortCardsCmd(SortCmd);
    }
    
    // IF STATUS IS UPDATED, SETS STATUS AND TAB COLOR ACCORDING TO STATUS 
    if (CellRow == StatusRow && CellCol == StatusCol){
      fcnUpdateDeckStatus(Value);
    }
  }
}




