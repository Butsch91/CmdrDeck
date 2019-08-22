// CmdrDckSort

// **********************************************
// function fcnSortCardsCmd()
//
// When the sort command cell is modified,
// the function sorts cards according to the choice 
// 
// **********************************************

function fcnSortCardsCmd(SortCmd) {
  
  switch (SortCmd) {
    case 'Type / Card' : fcnSortDeckTypeName(); break;
    case 'Type / Color' : fcnSortDeckTypeColor(); break;
    case 'Color' : fcnSortDeckColor(); break;
    case 'Category' : fcnSortDeckCategory(); break;
    case 'Card Name' : fcnSortDeckCardName(); break;
    case 'Staple' : fcnSortDeckStaple(); break;
    case 'Notes' : fcnSortDeckNotes(); break;
  }
}

// **********************************************
// function fcnSortDeckCardName()
//
// Sorts by Card(3), Type(6), Color(10), Staple(4), Category(8)
//
// **********************************************

function fcnSortDeckCardName(){
                     
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();

  if(actShtName != 'Template'  && actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef'){

    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Gets the Deck and Sideboard Ranges
    var ConfigSht = ss.getSheetByName('Config');
    var cfgRowCol = ConfigSht.getRange(18, 8, 17, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Parameters
    var CmdRow =        cfgRowCol[10][0];
    var CmdCol =        cfgRowCol[11][0];
    var StplStatusCol = cfgRowCol[ 4][0];
    var CardNameCol =   cfgRowCol[ 5][0];
    var CardTypeCol =   cfgRowCol[ 6][0];
    var CategoryCol =   cfgRowCol[ 7][0];
    var CardColorCol =  cfgRowCol[ 8][0];
    var NoteCol =       cfgRowCol[ 9][0];
    var SetCol =        cfgRowCol[14][0];
    
    
    // Looks for a Partner Commander *CMDR* in tag column (4)
    var CmdrFlag = 0;
    for(var TagRow = DeckFirstRow; TagRow <= (DeckFirstRow + DeckNumRows); TagRow++){
      var TagVal = actSht.getRange(TagRow, 4).getValue();
      if (TagVal == "*CMDR*"){
        CmdrFlag = 1;
        actSht.getRange(DeckFirstRow,2,TagRow-DeckFirstRow+1,NumCols).sort({column: 4, ascending: false});
      }
    }
    
    // Selects the Deck and Sideboad Ranges
    var DeckRange = actSht.getRange(DeckFirstRow+CmdrFlag, 2, DeckNumRows-CmdrFlag, NumCols);
    var SideRange = actSht.getRange(SideFirstRow, 2, SideNumRows, NumCols);    
     
    // Sets the Sorting Deck By: (Row 1 Column I)    
    actSht.getRange(CmdRow,CmdCol).setValue("Card Name");
    DeckRange.sort([CardNameCol,CardTypeCol-1,CardColorCol-1,4,8]);
    SideRange.sort([CardNameCol,CardTypeCol-1,CardColorCol-1,4,8]);
  }
}

// **********************************************
// function fcnSortDeckTypeColor()
//
// Sorts by Type(6), Color(10), Staple(4), Card(3), Category(8) 
//
// **********************************************

function fcnSortDeckTypeColor(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();
  var TestSht = ss.getSheetByName("Test");
    
  if(actShtName != 'Template' && actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef'){

    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Gets the Deck and Sideboard Ranges
    var ConfigSht = ss.getSheetByName('Config');
    var cfgRowCol = ConfigSht.getRange(18, 8, 17, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Parameters
    var CmdRow =        cfgRowCol[10][0];
    var CmdCol =        cfgRowCol[11][0];
    var StplStatusCol = cfgRowCol[ 4][0];
    var CardNameCol =   cfgRowCol[ 5][0];
    var CardTypeCol =   cfgRowCol[ 6][0];
    var CategoryCol =   cfgRowCol[ 7][0];
    var CardColorCol =  cfgRowCol[ 8][0];
    var NoteCol =       cfgRowCol[ 9][0];
    var SetCol =        cfgRowCol[14][0];
    
    // Looks for a Partner Commander *CMDR* in tag column (4)
    var CmdrFlag = 0;
    for(var TagRow = DeckFirstRow; TagRow <= (DeckFirstRow + DeckNumRows); TagRow++){
      var TagVal = actSht.getRange(TagRow, 4).getValue();
      if (TagVal == "*CMDR*"){
        CmdrFlag = 1;
        actSht.getRange(DeckFirstRow,2,TagRow-DeckFirstRow+1,NumCols).sort({column: 4, ascending: false});
      }
    }
    
    // Selects the Deck and Sideboad Ranges
    var DeckRange = actSht.getRange(DeckFirstRow+CmdrFlag, 2, DeckNumRows-CmdrFlag, NumCols);
    var SideRange = actSht.getRange(SideFirstRow, 2, SideNumRows, NumCols);    
     
    // Sets the Sorting Deck By: (Row 1 Column I) 
    actSht.getRange(CmdRow,CmdCol).setValue("Type / Color");
    DeckRange.sort([CardTypeCol-1,CardColorCol-1,CardNameCol,8]);
    SideRange.sort([CardTypeCol-1,CardColorCol-1,CardNameCol,8]);
  }
}


// **********************************************
// function fcnSortDeckTypeName()
//
// Sorts by Type(6), Card(3), Staple(4), Color(10), Category(8) 
//
// **********************************************

function fcnSortDeckTypeName(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();
  var TestSht = ss.getSheetByName("Test");
  
  if(actShtName != 'Template' && actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef'){

    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Gets the Deck and Sideboard Ranges
    var ConfigSht = ss.getSheetByName('Config');
    var cfgRowCol = ConfigSht.getRange(18, 8, 17, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Parameters
    var CmdRow =        cfgRowCol[10][0];
    var CmdCol =        cfgRowCol[11][0];
    var StplStatusCol = cfgRowCol[ 4][0];
    var CardNameCol =   cfgRowCol[ 5][0];
    var CardTypeCol =   cfgRowCol[ 6][0];
    var CategoryCol =   cfgRowCol[ 7][0];
    var CardColorCol =  cfgRowCol[ 8][0];
    var NoteCol =       cfgRowCol[ 9][0];
    var SetCol =        cfgRowCol[14][0];
    
    
    // Looks for a Partner Commander *CMDR* in tag column (4)
    var CmdrFlag = 0;
    for(var TagRow = DeckFirstRow; TagRow <= (DeckFirstRow + DeckNumRows); TagRow++){
      var TagVal = actSht.getRange(TagRow, 4).getValue();
      if (TagVal == "*CMDR*"){
        CmdrFlag = 1;
        actSht.getRange(DeckFirstRow,2,TagRow-DeckFirstRow+1,NumCols).sort({column: 4, ascending: false});
      }
    }
    
    // Selects the Deck and Sideboad Ranges
    var DeckRange = actSht.getRange(DeckFirstRow+CmdrFlag, 2, DeckNumRows-CmdrFlag, NumCols);
    var SideRange = actSht.getRange(SideFirstRow, 2, SideNumRows, NumCols);    
     
    // Sets the Sorting Deck By: (Row 1 Column I) 
    actSht.getRange(CmdRow,CmdCol).setValue("Type / Card");
    DeckRange.sort([CardTypeCol-1,CardNameCol,CardColorCol-1,CategoryCol-1]);
    SideRange.sort([CardTypeCol-1,CardNameCol,CardColorCol-1,CategoryCol-1]);
  }
}


// **********************************************
// function fcnSortDeckColor()
//
// Sorts by Color(10), Type(6), Staple(4), Card(3), Category(8)
//
// **********************************************

function fcnSortDeckColor(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();
    
  if(actShtName != 'Template' && actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef'){

    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Gets the Deck and Sideboard Ranges
    var ConfigSht = ss.getSheetByName('Config');
    var cfgRowCol = ConfigSht.getRange(18, 8, 17, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Parameters
    var CmdRow =        cfgRowCol[10][0];
    var CmdCol =        cfgRowCol[11][0];
    var StplStatusCol = cfgRowCol[ 4][0];
    var CardNameCol =   cfgRowCol[ 5][0];
    var CardTypeCol =   cfgRowCol[ 6][0];
    var CategoryCol =   cfgRowCol[ 7][0];
    var CardColorCol =  cfgRowCol[ 8][0];
    var NoteCol =       cfgRowCol[ 9][0];
    var SetCol =        cfgRowCol[14][0];
    
    
    // Looks for a Partner Commander *CMDR* in tag column (4)
    var CmdrFlag = 0;
    for(var TagRow = DeckFirstRow; TagRow <= (DeckFirstRow + DeckNumRows); TagRow++){
      var TagVal = actSht.getRange(TagRow, 4).getValue();
      if (TagVal == "*CMDR*"){
        CmdrFlag = 1;
        actSht.getRange(DeckFirstRow,2,TagRow-DeckFirstRow+1,NumCols).sort({column: 4, ascending: false});
      }
    }
    
    // Selects the Deck and Sideboad Ranges
    var DeckRange = actSht.getRange(DeckFirstRow+CmdrFlag, 2, DeckNumRows-CmdrFlag, NumCols);
    var SideRange = actSht.getRange(SideFirstRow, 2, SideNumRows, NumCols);    
     
    // Sets the Sorting Deck By: (Row 1 Column I) 
    actSht.getRange(CmdRow,CmdCol).setValue("Color");
    DeckRange.sort([CardColorCol-1,CardTypeCol-1,4,CardNameCol,CategoryCol-1]);
    SideRange.sort([CardColorCol-1,CardTypeCol-1,4,CardNameCol,CategoryCol-1]);
  }
}


// ***************************************************
// function fcnSortDeckCategory()
//
// Sorts by Category(8), Type(6), Card(3)
//
// ***************************************************

function fcnSortDeckCategory(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();

  if(actShtName != 'Template'  && actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef'){

    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Gets the Deck and Sideboard Ranges
    var ConfigSht = ss.getSheetByName('Config');
    var cfgRowCol = ConfigSht.getRange(18, 8, 17, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Parameters
    var CmdRow =        cfgRowCol[10][0];
    var CmdCol =        cfgRowCol[11][0];
    var StplStatusCol = cfgRowCol[ 4][0];
    var CardNameCol =   cfgRowCol[ 5][0];
    var CardTypeCol =   cfgRowCol[ 6][0];
    var CategoryCol =   cfgRowCol[ 7][0];
    var CardColorCol =  cfgRowCol[ 8][0];
    var NoteCol =       cfgRowCol[ 9][0];
    var SetCol =        cfgRowCol[14][0];
    
    
    // Looks for a Partner Commander *CMDR* in tag column (4)
    var CmdrFlag = 0;
    for(var TagRow = DeckFirstRow; TagRow <= (DeckFirstRow + DeckNumRows); TagRow++){
      var TagVal = actSht.getRange(TagRow, 4).getValue();
      if (TagVal == "*CMDR*"){
        CmdrFlag = 1;
        actSht.getRange(DeckFirstRow,2,TagRow-DeckFirstRow+1,NumCols).sort({column: 4, ascending: false});
      }
    }
    
    // Selects the Deck and Sideboad Ranges
    var DeckRange = actSht.getRange(DeckFirstRow+CmdrFlag, 2, DeckNumRows-CmdrFlag, NumCols);
    var SideRange = actSht.getRange(SideFirstRow, 2, SideNumRows, NumCols);    
     
    // Sets the Sorting Deck By: (Row 1 Column I) 
    actSht.getRange(CmdRow,CmdCol).setValue("Category");   
    DeckRange.sort([CategoryCol-1,CardTypeCol-1,CardNameCol]);
    SideRange.sort([CategoryCol-1,CardTypeCol-1,CardNameCol]);
  }
}



// **********************************************
// function fcnSortDeckStaple()
//
// Sorts by Staple(4), Color(10), Type(6), Card(3), Category(8) 
//
// **********************************************

function fcnSortDeckStaple(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();

  if(actShtName != 'Template'  && actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef'){

    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Gets the Deck and Sideboard Ranges
    var ConfigSht = ss.getSheetByName('Config');
    var cfgRowCol = ConfigSht.getRange(18, 8, 17, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Parameters
    var CmdRow =        cfgRowCol[10][0];
    var CmdCol =        cfgRowCol[11][0];
    var StplStatusCol = cfgRowCol[ 4][0];
    var CardNameCol =   cfgRowCol[ 5][0];
    var CardTypeCol =   cfgRowCol[ 6][0];
    var CategoryCol =   cfgRowCol[ 7][0];
    var CardColorCol =  cfgRowCol[ 8][0];
    var NoteCol =       cfgRowCol[ 9][0];
    var SetCol =        cfgRowCol[14][0];
    
    
    // Looks for a Partner Commander *CMDR* in tag column (4)
    var CmdrFlag = 0;
    for(var TagRow = DeckFirstRow; TagRow <= (DeckFirstRow + DeckNumRows); TagRow++){
      var TagVal = actSht.getRange(TagRow, 4).getValue();
      if (TagVal == "*CMDR*"){
        CmdrFlag = 1;
        actSht.getRange(DeckFirstRow,2,TagRow-DeckFirstRow+1,NumCols).sort({column: 4, ascending: false});
      }
    }
    
    // Selects the Deck and Sideboad Ranges
    var DeckRange = actSht.getRange(DeckFirstRow+CmdrFlag, 2, DeckNumRows-CmdrFlag, NumCols);
    var SideRange = actSht.getRange(SideFirstRow, 2, SideNumRows, NumCols);    
     
    // Sets the Sorting Deck By: (Row 1 Column I) 
    actSht.getRange(CmdRow,CmdCol).setValue("Staple");
    DeckRange.sort([StplStatusCol,CardColorCol-1,CardNameCol,CardTypeCol-1,CategoryCol-1]);
    SideRange.sort([StplStatusCol,CardColorCol-1,CardNameCol,CardTypeCol-1,CategoryCol-1]);
    }
}

// **********************************************
// function fcnSortDeckNotes()
//
// Sorts by Type(6), Notes(13), Card(3), Staple(4), Color(10), Category(8)
//
// **********************************************

function fcnSortDeckNotes(){
                     
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();

  if(actShtName != 'Template'  && actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef'){

    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Gets the Deck and Sideboard Ranges
    var ConfigSht = ss.getSheetByName('Config');
    var cfgRowCol = ConfigSht.getRange(18, 8, 17, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Parameters
    var CmdRow =        cfgRowCol[10][0];
    var CmdCol =        cfgRowCol[11][0];
    var CardNameCol =   cfgRowCol[ 5][0];
    var CardTypeCol =   cfgRowCol[ 6][0];
    var CategoryCol =   cfgRowCol[ 7][0];
    var CardColorCol =  cfgRowCol[ 8][0];
    var NoteCol =       cfgRowCol[ 9][0];
    var SetCol =        cfgRowCol[14][0];
    
    
    // Looks for a Partner Commander *CMDR* in tag column (4)
    var CmdrFlag = 0;
    for(var TagRow = DeckFirstRow; TagRow <= (DeckFirstRow + DeckNumRows); TagRow++){
      var TagVal = actSht.getRange(TagRow, 4).getValue();
      if (TagVal == "*CMDR*"){
        CmdrFlag = 1;
        actSht.getRange(DeckFirstRow,2,TagRow-DeckFirstRow+1,NumCols).sort({column: 4, ascending: false});
      }
    }
    
    // Selects the Deck and Sideboad Ranges
    var DeckRange = actSht.getRange(DeckFirstRow+CmdrFlag, 2, DeckNumRows-CmdrFlag, NumCols);
    var SideRange = actSht.getRange(SideFirstRow, 2, SideNumRows, NumCols);    
     
    // Sets the Sorting Deck By: (Row 1 Column I)    
    actSht.getRange(CmdRow,CmdCol).setValue("Notes");
    DeckRange.sort([NoteCol,CardTypeCol-1,CardNameCol,CardColorCol-1,CategoryCol-1]);
    SideRange.sort([NoteCol,CardTypeCol-1,CardNameCol,CardColorCol-1,CategoryCol-1]);
  }
}