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
  var LinkRng;
  
  // Gets the Deck and Sideboard Ranges
  var ConfigSht = ss.getSheetByName('Config');
  var DeckFirstRow = ConfigSht.getRange(4,8).getValue();  
  var cfgRowCol = ConfigSht.getRange(18, 8, 17, 1).getValues();
  
  // Parameters
  var LinkCol =       cfgRowCol[ 0][0];
 
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
  var rngStatus = actSht.getRange(2,1,1,3);
  var BackColor;
  var FontColor;
  var IgnoreDeck = 0;
    
  // Opens the Configuration Sheet and Gets the Deck and Sideboard Ranges
  var shtConfig = ss.getSheetByName('Config');
  var StatusVal = shtConfig.getRange('F6:F15').getValues();
  var StatusBck = shtConfig.getRange('F6:F15').getBackgrounds();
  var StatusFnt = shtConfig.getRange('F6:F15').getFontColors();
  
  // Compares the Deck Status to the Status List 
  switch (Status){
  
    case StatusVal[1][0] : {
      BackColor = StatusBck[1][0];
      FontColor = StatusFnt[1][0];
      break;
    }
    case StatusVal[2][0] : {
      BackColor = StatusBck[2][0];
      FontColor = StatusFnt[2][0];
      break;
    }
    case StatusVal[3][0] : {
      BackColor = StatusBck[3][0];
      FontColor = StatusFnt[3][0];
      break;
    }
    case StatusVal[4][0] : {
      BackColor = StatusBck[4][0];
      FontColor = StatusFnt[4][0];
      break;
    }
    case StatusVal[5][0] : {
      BackColor = StatusBck[5][0];
      FontColor = StatusFnt[5][0];
      break;
    }
    case StatusVal[6][0] : {
      BackColor = StatusBck[6][0];
      FontColor = StatusFnt[6][0];
      break;
    }
    case StatusVal[7][0] : {
      BackColor = StatusBck[7][0];
      FontColor = StatusFnt[7][0];
      break;
    }
    case StatusVal[8][0] : {
      BackColor = StatusBck[8][0];
      FontColor = StatusFnt[8][0];
      break;
    }
    case StatusVal[9][0] : {
      BackColor = StatusBck[9][0];
      FontColor = StatusFnt[9][0];
      break;
    }
  }
  
  // Sets the Status and Tab Color according to the Status Value
  if (Status != ''){
    rngStatus.setBackground(BackColor);
    rngStatus.setFontColor(FontColor);
    actSht.setTabColor(BackColor);
  }
    
  // Sets the Status and Tab Color according to the Status Value
  if (Status == ''){
    BackColor = null;
    rngStatus.setBackground('#cfe2f3');
    rngStatus.setFontColor('black');
    actSht.setTabColor(BackColor);
  }  
}

// **********************************************
// function fcnUpdateStapleStatus()
//
// 
// 
// **********************************************
function fcnUpdateStapleStatus(Status) {
  
  // Gets Sheet Data
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var actSht = ss.getActiveSheet();								
  var actShtName = actSht.getSheetName();
  var MaxCol = actSht.getMaxColumns();
  var rngStatus = actSht.getRange(2,4);
  var BackColor;
  var FontColor;
  var IgnoreDeck = 0;
    
  // Opens the Configuration Sheet and Gets the Deck and Sideboard Ranges
  var shtConfig = ss.getSheetByName('Config');
  var StatusVal = shtConfig.getRange('F16:F25').getValues();
  var StatusBck = shtConfig.getRange('F16:F25').getBackgrounds();
  var StatusFnt = shtConfig.getRange('F16:F25').getFontColors();
  
  // Compares the Deck Status to the Status List 
  switch (Status){
  
    case StatusVal[1][0] : {
      BackColor = StatusBck[1][0];
      FontColor = StatusFnt[1][0];
      break;
    }
    case StatusVal[2][0] : {
      BackColor = StatusBck[2][0];
      FontColor = StatusFnt[2][0];
      break;
    }
    case StatusVal[3][0] : {
      BackColor = StatusBck[3][0];
      FontColor = StatusFnt[3][0];
      break;
    }
    case StatusVal[4][0] : {
      BackColor = StatusBck[4][0];
      FontColor = StatusFnt[4][0];
      break;
    }
    case StatusVal[5][0] : {
      BackColor = StatusBck[5][0];
      FontColor = StatusFnt[5][0];
      break;
    }
    case StatusVal[6][0] : {
      BackColor = StatusBck[6][0];
      FontColor = StatusFnt[6][0];
      break;
    }
    case StatusVal[7][0] : {
      BackColor = StatusBck[7][0];
      FontColor = StatusFnt[7][0];
      break;
    }
    case StatusVal[8][0] : {
      BackColor = StatusBck[8][0];
      FontColor = StatusFnt[8][0];
      break;
    }
    case StatusVal[9][0] : {
      BackColor = StatusBck[9][0];
      FontColor = StatusFnt[9][0];
      break;
    }
  }
  
  // Sets the Status and Tab Color according to the Status Value
  if (Status != ''){
    rngStatus.setBackground(BackColor);
    rngStatus.setFontColor(FontColor);
//    actSht.setTabColor(BackColor);
  }
    
  // Sets the Status and Tab Color according to the Status Value
  if (Status == ''){
    BackColor = null;
    rngStatus.setBackground('#cfe2f3');
    rngStatus.setFontColor('black');
    actSht.setTabColor(BackColor);
  }  
  
  // Gets the Deck Name
  var DeckName = actSht.getRange(1, 1).getValue();
  // Get List of Decks to Ignore for Commander Staple Binder
  var IgnoreDeckList = shtConfig.getRange('E23:E32').getValues();
  // Look for Deck to Ignore
  for(var i = 0; i< IgnoreDeckList.length; i++){
    if(DeckName == IgnoreDeckList[i][0]) IgnoreDeck = 1;
  }

  // Opens Spreadsheet _Cmdr Staple Binder to update the Deck Tab Color
  if (IgnoreDeck == 0){ 
    var ssStaple = SpreadsheetApp.openById('1l44UmxpachzK7SHETkOX1qkOk56uNUoyvOjzb9BrhgQ');
    var StapleDeckSht = ssStaple.getSheetByName(DeckName);
    StapleDeckSht.setTabColor(BackColor);
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