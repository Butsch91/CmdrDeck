// CmdrDckCode

// **********************************************
// function fcnGathererLink()
//
// Creates a link to the Gatherer Card Page for
// each card in the list
// 
// **********************************************

function fcnGathererLink() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();								
  var actShtName = actSht.getSheetName();
  var MaxRow = actSht.getMaxRows();
  var Row;
  var CardCol = 3;
  var CardName;
  var CardRng;
  var LinkCol = 14;
  var LinkRng;
  
  // Gets the Deck and Sideboard Ranges
  var ConfigSht = ss.getSheetByName('Config');
  var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
 
  if(actShtName != 'Template'  && actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef') { 
    for (Row = DeckFirstRow - 1; Row <= MaxRow; Row++){
      CardRng = actSht.getRange(Row, CardCol);
      LinkRng = actSht.getRange(Row, LinkCol);
      CardName = CardRng.getValue();
      if(CardName != '') LinkRng.setValue('=HYPERLINK("http://gatherer.wizards.com/Pages/Card/Details.aspx?name='+CardName+'","'+CardName+'")');
      LinkRng.setFontLine('none');
    }
  }
}


// **********************************************
// function fcnClearGathererLink()
//
// Clears all Gatherer Links for
// each card in the list
// 
// **********************************************

function fcnClearGathererLink() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();								
  var actShtName = actSht.getSheetName();
  var MaxRow = actSht.getMaxRows();
  var Row;
  var CardCol = 3;
  var CardName;
  var CardRng;
  var LinkCol = 14;
  var LinkRng;

  // Gets the Deck and Sideboard Ranges
  var ConfigSht = ss.getSheetByName('Config');
  var DeckFirstRow = ConfigSht.getRange(4,8).getValue();
  
  if(actShtName != 'Template' && actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef') { 
    for (Row = DeckFirstRow - 1; Row <= MaxRow; Row++){
      CardRng = actSht.getRange(Row, CardCol);
      LinkRng = actSht.getRange(Row, LinkCol);
      CardName = CardRng.getValue();
      if(CardName != '') LinkRng.clear();
    }
  }
}

// **********************************************
// function fcnHideColumns()
//
// Hide all columns with the "hide" title
// 
// 
// **********************************************

function fcnHideColumns() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();								
  var actShtName = actSht.getSheetName();
  var MaxRow = actSht.getMaxRows();
  var MaxCol = actSht.getMaxColumns();
  var ColTitle;
  
  if(actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef') { 
    for (var Col = 1; Col <= MaxCol; Col++){
      ColTitle = actSht.getRange(1, Col).getValue();
      if(ColTitle == 'hide') {
        var ColRng = actSht.getRange(1, Col, MaxRow, 1);
        
        actSht.hideColumn(ColRng);
      }
    }
  }
}

// **********************************************
// function fcnShowColumns()
//
// Hide all columns with the "hide" title
// 
// 
// **********************************************

function fcnShowColumns() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();								
  var actShtName = actSht.getSheetName();
  var MaxRow = actSht.getMaxRows();
  var MaxCol = actSht.getMaxColumns();
  var ColTitle;
  
  if(actShtName != 'Config' && actShtName != 'Conversion' && actShtName != 'Test' && actShtName != 'Staple CrossRef') { 
    for (var Col = 1; Col <= MaxCol; Col++){
      ColTitle = actSht.getRange(1, Col).getValue();
      if(ColTitle == 'hide') {
        var ColRng = actSht.getRange(1, Col, MaxRow, 1);
        actSht.unhideColumn(ColRng);
      }
    }
  }
}

// **********************************************
// function fcnUpdateDeckStatus()
//
// 
// 
// **********************************************
function fcnUpdateDeckStatus(Status) {
  
  // Gets Sheet Data
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();								
  var actShtName = actSht.getSheetName();
  var MaxCol = actSht.getMaxColumns();
  var BackColor;
  var FontColor;
    
  // Opens the Configuration Sheet and Gets the Deck and Sideboard Ranges
  var ConfigSht = ss.getSheetByName('Config');
  
  // Selects the Tab Color
  if (Status == 'Ready to Play'){
    BackColor = ConfigSht.getRange(7, 6).getBackground();
    FontColor = ConfigSht.getRange(7, 6).getFontColor();
  } 
  
  if (Status == 'Recheck Deck and List'){
    BackColor = ConfigSht.getRange(8, 6).getBackground();
    FontColor = ConfigSht.getRange(8, 6).getFontColor();
  } 
  
  if (Status == 'Update List'){
    BackColor = ConfigSht.getRange(9, 6).getBackground();
    FontColor = ConfigSht.getRange(9, 6).getFontColor();
  } 
  
  if (Status == 'Update Deck'){
    BackColor = ConfigSht.getRange(10, 6).getBackground();
    FontColor = ConfigSht.getRange(10, 6).getFontColor();
  } 
  
  if (Status == 'Update Everything'){
    BackColor = ConfigSht.getRange(11, 6).getBackground();
    FontColor = ConfigSht.getRange(11, 6).getFontColor();
  }
  
  if (Status == 'In Construction'){
    BackColor = ConfigSht.getRange(14, 6).getBackground();
    FontColor = ConfigSht.getRange(14, 6).getFontColor();
  }   
  
  if (Status == 'Not Listed'){
    BackColor = ConfigSht.getRange(15, 6).getBackground();
    FontColor = ConfigSht.getRange(15, 6).getFontColor();
  } 
  
  // Sets the Status and Tab Color according to the Status Value
  if (Status != ''){
    actSht.getRange(2,1,1,3).setBackground(BackColor);
    actSht.getRange(2,1,1,3).setFontColor(FontColor);
    actSht.setTabColor(BackColor);
  }
    
  // Sets the Status and Tab Color according to the Status Value
  if (Status == ''){
    BackColor = null;
    actSht.getRange(2,1,1,3).setBackground('#cfe2f3');
    actSht.getRange(2,1,1,3).setFontColor('black');
    actSht.setTabColor(BackColor);
  }  
  
  // Gets the Deck Name
  var DeckName = actSht.getRange(1, 1).getValue();
  
  // Opens Spreadsheet _Cmdr Staple Binder to update the Deck Tab Color
  if (DeckName != 'Silvos' && DeckName != 'Ob Nixilis'){ 
    var ssStaple = SpreadsheetApp.openById('1l44UmxpachzK7SHETkOX1qkOk56uNUoyvOjzb9BrhgQ');
    var StapleDeckSht = ssStaple.getSheetByName(DeckName);
    StapleDeckSht.setTabColor(BackColor);
  }
}


// **********************************************
// function fcnSortCardsCmd()
//
// When the sort command cell is modified,
// the function sorts cards according to the choice 
// 
// **********************************************

function fcnSortCardsCmd(SortCmd) {
  
  if (SortCmd == 'Type / Card'){
    fcnSortDeckTypeName();
  }
  
  if (SortCmd == 'Type / Color'){
    fcnSortDeckTypeColor();
  }
  
  if (SortCmd == 'Color'){
    fcnSortDeckColor();
  }
  
  if (SortCmd == 'Category'){
    fcnSortDeckCategory();
  }
  
  if (SortCmd == 'Card Name'){
    fcnSortDeckCardName();
  }
  
  if (SortCmd == 'Staple'){
    fcnSortDeckStaple();
  }
}



// **********************************************
// function GetStapleCrossRef()
//
// Gets the Cross Reference sheet from
// the Staple Binder Spreadsheet
//
// **********************************************

function GetStapleCrossRef() {
  
  // Opens the current Equipment Equivalences sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var StapleRefSht = ss.getSheetByName('Staple CrossRef');

  // Opens Spreadsheet _Cmdr Staple Binder
  var ssStaple = SpreadsheetApp.openById('1l44UmxpachzK7SHETkOX1qkOk56uNUoyvOjzb9BrhgQ');

  // Deletes the current Cross Ref Sheet
  if(StapleRefSht != null) ss.deleteSheet(StapleRefSht);
 
  // Opens and copies the Source Sheet from Inbound Spreadsheet to this Spreadsheet
  var srcRefSht = ssStaple.getSheetByName('Staple CrossRef');
  srcRefSht.copyTo(ss);
  
  // Opens and renames the copy
  StapleRefSht = ss.getSheetByName('Copy of Staple CrossRef');
  StapleRefSht.setName('Staple CrossRef');

  
//  // Sets the First Sheet as Active Sheet
//  var FirstSht =  ss.getSheets()[0];
//  var NbSheets = ss.getNumSheets();
//  
  ss.setActiveSheet(StapleRefSht);
  
}