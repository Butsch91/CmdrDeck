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
  var cfgRowCol = ConfigSht.getRange(17, 8, 26, 1).getValues();
  
  // Parameters
  var colNb 	   = cfgRowCol[ 1][0];
  var colName      = cfgRowCol[ 2][0];
  var colSet 	   = cfgRowCol[ 3][0];
  var colStaple    = cfgRowCol[ 4][0];
  var colTypeSort  = cfgRowCol[ 5][0];
  var colType 	   = cfgRowCol[ 6][0];
  var colCatSort   = cfgRowCol[ 7][0];
  var colCat       = cfgRowCol[ 8][0];
  var colColorSort = cfgRowCol[ 9][0];
  var colColor     = cfgRowCol[10][0];
  var colNote      = cfgRowCol[11][0];
  var colLink      = cfgRowCol[12][0];
  
  var rowDeckStatus = cfgRowCol[18][0];
  var colDeckStatus = cfgRowCol[19][0];
  var rowStplStatus = cfgRowCol[20][0];
  var colStplStatus = cfgRowCol[21][0];
  var rowCmd 		= cfgRowCol[22][0];
  var colCmd 		= cfgRowCol[23][0];
  var rowGenDeck 	= cfgRowCol[24][0];
  var colGenDeck 	= cfgRowCol[25][0];
  
  var LinkRng = actSht.getRange(CellRow, colLink);
  var ClrEnable = actSht.getRange(2,13).getValue();
  
  Logger.log(ClrEnable);  
  //Logger.log(cfgRowCol);
    
  // CHECKS IF THE SELECTED SHEET IS VALID
  if(actShtName != 'Template' &&  actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef'){
    
    // IF THE MODIFIED CELL IS THE SORT COMMAND CELL, EXECUTES SORT CARDS FUNCTION
    if (CellRow == rowCmd && CellCol == colCmd && Value != ''){
      fcnSortCardsCmd(Value);
    }
    
    // IF THE MODIFIED CELL IS THE GENERATE NEW DECK VERSION COMMAND CELL, EXECUTES THE GENERATE NEW DECK VERSION FUNCTION
    if (CellRow == rowGenDeck && CellCol == colGenDeck && Value == 'Generate New Deck'){
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
    if (CellRow >= DeckFirstRow && CellCol == colType && (Value == 'Land' || Value == 'Basic Land')){
      var CardType = Value;
      
      if (CardType == 'Land' || CardType == 'Basic Land'){
        actSht.getRange(CellRow,colCat).setValue('Land');
        actSht.getRange(CellRow,colColor).setValue('L');
      }
      
      if (CardType == "Artifact"){
        actSht.getRange(CellRow,colColor).setValue('C');
      }
    }
    
    // IF THE CARD ADDED WAS A BASIC LAND, LAND CATEGORY AND LAND COLOR IS AUTOMATICALLY ADDED.
    if (CellRow >= DeckFirstRow && CellCol == colName && (Value == 'Plains' || Value == 'Island' || Value == 'Swamp' || Value == 'Mountain' || Value == 'Forest' || Value == 'Wastes')){

      actSht.getRange(CellRow,colType).setValue('Basic Land');
      actSht.getRange(CellRow,colCat).setValue('Land');
      actSht.getRange(CellRow,colColor).setValue('L');
    }
    
    // IF A NOTE CELL IS EDITED AND IT CONTAINS CERTAIN WORDS, HIGHLIGHT THE LINE 
    if (CellRow >= DeckFirstRow && CellCol == colNote){
      
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
      actSht.getRange(CellRow,colStplStatus).setValue("");
      actSht.getRange(CellRow,colSet).setValue("");
      actSht.getRange(CellRow,colType).setValue("");
      actSht.getRange(CellRow,colCat).setValue("");
      actSht.getRange(CellRow,colColor).setValue("");
      actSht.getRange(CellRow,colNote).setValue("");
      actSht.getRange(CellRow,colLink).setValue("");
      actSht.getRange(CellRow, 1, 1, MaxCol).setBackground(null);
      
      // Executes the Sort Command
      var SortCmd = actSht.getRange(rowCmd, colCmd).getValue();
      fcnSortCardsCmd(SortCmd);
    }
    
    // IF DECK STATUS IS UPDATED, SETS STATUS AND TAB COLOR ACCORDING TO STATUS 
    if (CellRow == rowDeckStatus && CellCol == colDeckStatus){
      fcnUpdateDeckStatus(Value);
    }
    
    // IF STAPLE STATUS IS UPDATED, SETS STATUS AND TAB COLOR ACCORDING TO STATUS 
    if (CellRow == rowStplStatus && CellCol == colStplStatus){
      fcnUpdateStapleStatus(Value);
    }
  }
}




