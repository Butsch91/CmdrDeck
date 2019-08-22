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
  
  Logger.log("Row:%s / Col: %s / Value: %s",CellRow, CellCol, Value);  
    
  // Opens the Configuration Sheet and Gets the Deck and Sideboard Ranges
  var ConfigSht = ss.getSheetByName('Config');
  var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
  var SdBdFirstRow = ConfigSht.getRange(6,8).getValue();
  var cfgRowCol = ConfigSht.getRange(18, 8, 17, 1).getValues();
  
  // Parameters
  var LinkCol =       cfgRowCol[ 0][0];
  var DeckStatusRow = cfgRowCol[ 1][0];
  var DeckStatusCol = cfgRowCol[ 2][0];
  var StplStatusRow = cfgRowCol[ 3][0];
  var StplStatusCol = cfgRowCol[ 4][0];
  var CardNameCol =   cfgRowCol[ 5][0];
  var CardTypeCol =   cfgRowCol[ 6][0];
  var CategoryCol =   cfgRowCol[ 7][0];
  var CardColorCol =  cfgRowCol[ 8][0];
  var NoteCol =       cfgRowCol[ 9][0];
  var CmdRow =        cfgRowCol[10][0];
  var CmdCol =        cfgRowCol[11][0];
  var GenDeckRow =    cfgRowCol[12][0];
  var GenDeckCol =    cfgRowCol[13][0];
  var SetCol =        cfgRowCol[14][0];
  var LinkRng = actSht.getRange(CellRow, LinkCol);
  var ClrEnable = actSht.getRange(2,13).getValue();
  
  Logger.log(ClrEnable);  
  //Logger.log(cfgRowCol);
    
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
      LinkRng.setFontLine('none');
      // If Card Quantity is null, put 1
      var rngCardQty = actSht.getRange(CellRow, 2);
      var CardQty = rngCardQty.getValue();
      if(CardQty == '') rngCardQty.setValue(1); 
    }
    
    // IF THE CARD TYPE MODIFIED IS A LAND, LAND CATEGORY AND LAND COLOR IS AUTOMATICALLY ADDED. IF ARTIFACT, COLORLESS IS AUTOMATICALLY ADDED
    if (CellRow >= DeckFirstRow && CellCol == CardTypeCol && (Value == 'Land' || Value == 'Basic Land')){
      var CardType = Value;
      
      if (CardType == 'Land' || CardType == 'Basic Land'){
        actSht.getRange(CellRow,CategoryCol).setValue('Land');
        actSht.getRange(CellRow,CardColorCol).setValue('L');
      }
      
      if (CardType == "Artifact"){
        actSht.getRange(CellRow,CardColorCol).setValue('C');
      }
    }
    
    // IF THE CARD ADDED WAS A BASIC LAND, LAND CATEGORY AND LAND COLOR IS AUTOMATICALLY ADDED.
    if (CellRow >= DeckFirstRow && CellCol == CardNameCol && (Value == 'Plains' || Value == 'Island' || Value == 'Swamp' || Value == 'Mountain' || Value == 'Forest' || Value == 'Wastes')){

      actSht.getRange(CellRow,CardTypeCol).setValue('Basic Land');
      actSht.getRange(CellRow,CategoryCol).setValue('Land');
      actSht.getRange(CellRow,CardColorCol).setValue('L');
    }
    
    // IF A NOTE CELL IS EDITED AND IT CONTAINS CERTAIN WORDS, HIGHLIGHT THE LINE 
    if (CellRow >= DeckFirstRow && CellCol == NoteCol){
      
      var NoteKeywordPresent = 0;
      
      // IF NOTE CELL IS NOT EMPTY
      if (Value != '') {
        
        var colorBckgnd = null;
        
        // LOOK FOR THE WORDS "Replace", "replace" AND ITS OTHER FORMS
        if (NoteKeywordPresent == 0 && ((Value.indexOf('Replace') > -1) || (Value.indexOf('Replaces') > -1) || (Value.indexOf('Replaced') > -1) || (Value.indexOf('replace') > -1) || (Value.indexOf('replaces') > -1) || (Value.indexOf('replaced') > -1)) ) {
          NoteKeywordPresent = 1;
          colorBckgnd = '#ffe599';
        }
        // LOOK FOR THE WORDS "Remove", "remove" AND ITS OTHER FORMS        
        if (NoteKeywordPresent == 0 && ((Value.indexOf('Remove') > -1) || (Value.indexOf('Removes') > -1) || (Value.indexOf('Removed') > -1) || (Value.indexOf('remove') > -1) || (Value.indexOf('removes') > -1) || (Value.indexOf('removed') > -1)) ) {
          NoteKeywordPresent = 1;
          colorBckgnd = '#ffe599';
        }
        
        // LOOK FOR THE WORDS "Add", "add", "Get", "get"        
        if (NoteKeywordPresent == 0 && ((Value.indexOf('Add') > -1) || (Value.indexOf('add') > -1) || (Value.indexOf('Get') > -1) || (Value.indexOf('get') > -1)) ) {
          NoteKeywordPresent = 1;
          colorBckgnd = '#ffe599';
        }
        
        // LOOK FOR THE WORDS "Missing", "missing"        
        if (NoteKeywordPresent == 0 && ((Value.indexOf('Missing') > -1) || (Value.indexOf('missing') > -1)) ) {
          NoteKeywordPresent = 1;
          colorBckgnd = '#ffe599';
        }

        // LOOK FOR THE WORDS "In", "in"     
        if (NoteKeywordPresent == 0 && ((Value.indexOf('In') > -1) || (Value.indexOf('in') > -1)) ) {
          NoteKeywordPresent = 1;
          colorBckgnd = '#d9ead3';
        }
        
        // LOOK FOR THE WORDS "Out", "out"        
        if (NoteKeywordPresent == 0 && ((Value.indexOf('Out') > -1) || (Value.indexOf('out') > -1)) ) {
          NoteKeywordPresent = 1;
          colorBckgnd = '#f4cccc';
          actSht.getRange(CellRow,2).setValue("");
        }
        
        if(NoteKeywordPresent == 1){
        // Present
          actSht.getRange(CellRow, 2, 1, MaxCol-1).setBackground(colorBckgnd);
        }
        // Not Present
        else
          actSht.getRange(CellRow, 1, 1, MaxCol).setBackground(null);
      }
      
      // IF NOTE CELL IS CLEARED
      if (Value == '') {
        actSht.getRange(CellRow, 1, 1, MaxCol).setBackground(null);
      }
    }
    
    // IF CARD NAME WAS CLEARED, DELETES ALL INFO IF CLEAR IS ENABLED  
    if (CellRow >= DeckFirstRow && CellCol == 3 && Value == '' && ClrEnable == 'Enable Clear') { 
      if(CellRow < SdBdFirstRow) actSht.getRange(CellRow,2).setValue("");
      actSht.getRange(CellRow,StplStatusCol).setValue("");
      actSht.getRange(CellRow,SetCol).setValue("");
      actSht.getRange(CellRow,CardTypeCol).setValue("");
      actSht.getRange(CellRow,CategoryCol).setValue("");
      actSht.getRange(CellRow,CardColorCol).setValue("");
      actSht.getRange(CellRow,NoteCol).setValue("");
      actSht.getRange(CellRow,LinkCol).setValue("");
      actSht.getRange(CellRow, 1, 1, MaxCol).setBackground(null);
      
      // Executes the Sort Command
      var SortCmd = actSht.getRange(CmdRow, CmdCol).getValue();
      fcnSortCardsCmd(SortCmd);
    }
    
    // IF DECK STATUS IS UPDATED, SETS STATUS AND TAB COLOR ACCORDING TO STATUS 
    if (CellRow == DeckStatusRow && CellCol == DeckStatusCol){
      fcnUpdateDeckStatus(Value);
    }
    
    // IF STAPLE STATUS IS UPDATED, SETS STATUS AND TAB COLOR ACCORDING TO STATUS 
    if (CellRow == StplStatusRow && CellCol == StplStatusCol){
      fcnUpdateStapleStatus(Value);
    }
  }
}




