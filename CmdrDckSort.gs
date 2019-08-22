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
    case 'Staple' : fcnSortDeckStaple(); break;
    case 'Type / Card' : fcnSortDeckTypeName(); break;
    case 'Type / Color' : fcnSortDeckTypeColor(); break;
    case 'Color' : fcnSortDeckColor(); break;
    case 'Card Name' : fcnSortDeckCardName(); break;
    case 'Category' : fcnSortDeckCategory(); break;
    case 'Notes' : fcnSortDeckNotes(); break;
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
    var cfgRowCol = ConfigSht.getRange(17, 8, 26, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Parameters
    var colName      = cfgRowCol[ 2][0];
    var colSet 	     = cfgRowCol[ 3][0];
    var colStaple    = cfgRowCol[ 4][0];
    var colTypeSort  = cfgRowCol[ 5][0];
    var colCatSort   = cfgRowCol[ 7][0];
    var colColorSort = cfgRowCol[ 9][0];
    var colNote      = cfgRowCol[11][0];
    
    var rowCmd 		= cfgRowCol[22][0];
    var colCmd 		= cfgRowCol[23][0];
     
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
    actSht.getRange(rowCmd,colCmd).setValue("Staple");
    DeckRange.sort([colStaple,colColorSort,colName,colTypeSort,colCatSort]);
    SideRange.sort([colStaple,colColorSort,colName,colTypeSort,colCatSort]);
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
  
  Logger.log(actShtName);
  
  if(actShtName != 'Template' && actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef'){

    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Gets the Deck and Sideboard Ranges
    var ConfigSht = ss.getSheetByName('Config');
    var cfgRowCol = ConfigSht.getRange(17, 8, 26, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Parameters
    var colName      = cfgRowCol[ 2][0];
    var colSet 	     = cfgRowCol[ 3][0];
    var colStaple    = cfgRowCol[ 4][0];
    var colTypeSort  = cfgRowCol[ 5][0];
    var colCatSort   = cfgRowCol[ 7][0];
    var colColorSort = cfgRowCol[ 9][0];
    var colNote      = cfgRowCol[11][0];
    
    var rowCmd 		= cfgRowCol[22][0];
    var colCmd 		= cfgRowCol[23][0];
        
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
    actSht.getRange(rowCmd,colCmd).setValue("Type / Card");
    DeckRange.sort([colTypeSort,colName,colColorSort,colCatSort]);
    SideRange.sort([colTypeSort,colName,colColorSort,colCatSort]);
  }
}

  
// **********************************************
// function fcnSortDeckTypeColor()
//
// Sorts by Type(6), Color(10), Staple(4), Card(3), Category(colCatSort) 
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
    var cfgRowCol = ConfigSht.getRange(17, 8, 26, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Parameters
    var colName      = cfgRowCol[ 2][0];
    var colSet 	     = cfgRowCol[ 3][0];
    var colStaple    = cfgRowCol[ 4][0];
    var colTypeSort  = cfgRowCol[ 5][0];
    var colCatSort   = cfgRowCol[ 7][0];
    var colColorSort = cfgRowCol[ 9][0];
    var colNote      = cfgRowCol[11][0];
    
    var rowCmd 		= cfgRowCol[22][0];
    var colCmd 		= cfgRowCol[23][0];
        
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
    actSht.getRange(rowCmd,colCmd).setValue("Type / Color");
    DeckRange.sort([colTypeSort,colColorSort,colName,8]);
    SideRange.sort([colTypeSort,colColorSort,colName,8]);
  }
}


// **********************************************
// function fcnSortDeckCardName()
//
// Sorts by Card Name(colName), Type(colTypeSort), Color(colColorSort), Staple(colStaple), Category(colCatSort)
//
// **********************************************

function fcnSortDeckCardName(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();
  var actShtName = actSht.getName();
  var TestSht = ss.getSheetByName("Test");
  
  
  // IF SHEET IS ALLOWED TO BE SORTED, EXECUTE SWITCH
  if(actShtName != 'Template'  && actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef'){
    
    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Gets the Deck and Sideboard Ranges
    var ConfigSht = ss.getSheetByName('Config');
    var cfgRowCol = ConfigSht.getRange(17, 8, 26, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Parameters
    var colName      = cfgRowCol[ 2][0];
    var colSet 	     = cfgRowCol[ 3][0];
    var colStaple    = cfgRowCol[ 4][0];
    var colTypeSort  = cfgRowCol[ 5][0];
    var colCatSort   = cfgRowCol[ 7][0];
    var colColorSort = cfgRowCol[ 9][0];
    var colNote      = cfgRowCol[11][0];
    
    var rowCmd 		= cfgRowCol[22][0];
    var colCmd 		= cfgRowCol[23][0];
    
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
    actSht.getRange(rowCmd,colCmd).setValue("Card Name");
    DeckRange.sort([colName,colTypeSort,colColorSort,colStaple,colCatSort]);
    SideRange.sort([colName,colTypeSort,colColorSort,colStaple,colCatSort]);
    
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
    var cfgRowCol = ConfigSht.getRange(17, 8, 26, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Parameters
    var colName      = cfgRowCol[ 2][0];
    var colSet 	     = cfgRowCol[ 3][0];
    var colStaple    = cfgRowCol[ 4][0];
    var colTypeSort  = cfgRowCol[ 5][0];
    var colCatSort   = cfgRowCol[ 7][0];
    var colColorSort = cfgRowCol[ 9][0];
    var colNote      = cfgRowCol[11][0];
    
    var rowCmd 		= cfgRowCol[22][0];
    var colCmd 		= cfgRowCol[23][0];
        
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
    actSht.getRange(rowCmd,colCmd).setValue("Color");
    DeckRange.sort([colColorSort,colTypeSort,4,colName,colCatSort]);
    SideRange.sort([colColorSort,colTypeSort,4,colName,colCatSort]);
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
    var cfgRowCol = ConfigSht.getRange(17, 8, 26, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Parameters
    var colName      = cfgRowCol[ 2][0];
    var colSet 	     = cfgRowCol[ 3][0];
    var colStaple    = cfgRowCol[ 4][0];
    var colTypeSort  = cfgRowCol[ 5][0];
    var colCatSort   = cfgRowCol[ 7][0];
    var colColorSort = cfgRowCol[ 9][0];
    var colNote      = cfgRowCol[11][0];
    
    var rowCmd 		= cfgRowCol[22][0];
    var colCmd 		= cfgRowCol[23][0];
    
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
    actSht.getRange(rowCmd,colCmd).setValue("Category");   
    DeckRange.sort([colCatSort,colTypeSort,colName]);
    SideRange.sort([colCatSort,colTypeSort,colName]);
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
    var cfgRowCol = ConfigSht.getRange(17, 8, 26, 1).getValues();
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
    // Defines number of columns to sort
    var MaxCols = actSht.getMaxColumns();
    var NumCols = MaxCols - 1;
    
    // Parameters
    var colName      = cfgRowCol[ 2][0];
    var colSet 	     = cfgRowCol[ 3][0];
    var colStaple    = cfgRowCol[ 4][0];
    var colTypeSort  = cfgRowCol[ 5][0];
    var colCatSort   = cfgRowCol[ 7][0];
    var colColorSort = cfgRowCol[ 9][0];
    var colNote      = cfgRowCol[11][0];
    
    var rowCmd 		= cfgRowCol[22][0];
    var colCmd 		= cfgRowCol[23][0];
     
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
    actSht.getRange(rowCmd,colCmd).setValue("Notes");
    DeckRange.sort([colNote,colTypeSort,colName,colColorSort,colCatSort]);
    SideRange.sort([colNote,colTypeSort,colName,colColorSort,colCatSort]);
  }
}
