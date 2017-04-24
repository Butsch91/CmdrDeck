// CmdrDckSort

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
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
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
    actSht.getRange(1,9).setValue("Type / Color");
    DeckRange.sort([6,10,3,8]);
    SideRange.sort([6,10,3,8]);
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
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
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
    actSht.getRange(1,9).setValue("Type / Card");
    DeckRange.sort([6,3,10,8]);
    SideRange.sort([6,3,10,8]);
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
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
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
    actSht.getRange(1,9).setValue("Color");
    DeckRange.sort([10,6,4,3,8]);
    SideRange.sort([10,6,4,3,8]);
  }
}


// ***************************************************
// function fcnSortDeckCategory()
//
// Sorts by Category(8), Type(6), Staple(4), Card(3)
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
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
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
    actSht.getRange(1,9).setValue("Category");   
    DeckRange.sort([8,6,4,3]);
    SideRange.sort([8,6,4,3]);
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
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
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
    actSht.getRange(1,9).setValue("Card Name");
    DeckRange.sort([3,6,10,4,8]);
    SideRange.sort([3,6,10,4,8]);
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
    var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
    var DeckNumRows = ConfigSht.getRange(5,8).getValue();
    var SideFirstRow = ConfigSht.getRange(6,8).getValue();
    var SideNumRows = ConfigSht.getRange(7,8).getValue();
    
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
    actSht.getRange(1,9).setValue("Staple");
    DeckRange.sort([4,10,3,6,8]);
    SideRange.sort([4,10,3,6,8]);
    }
}