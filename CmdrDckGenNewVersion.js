// CmdrDckGenNewVersion

// **********************************************
// function fcnGenerateNewVersion()
//
// Generate New Version of Deck
// 
// 
// **********************************************

function fcnGenerateNewVersion() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var NbSheets = ss.getNumSheets();
  var actSht = ss.getActiveSheet();
  
  var OldDeck = actSht;
  var OldDeckNum = OldDeck.getIndex();
  var OldVersionNum = OldDeck.getRange(1, 4).getValue();
  var OldDeckName = OldDeck.getSheetName();
  
  var DeckName = OldDeck.getRange(1, 1).getValue();
  
  var CopyName;
  var ArchiveDeck;
  
  var NewDeck;
  var NewDeckNum;
  var NewVersionNum = OldVersionNum+1 ;
  var NewDeckName = DeckName + ' v' + NewVersionNum;
  
  // INSERTS TAB BEFORE "DECKLISTS" TAB
  ss.insertSheet(NewDeckName, OldDeckNum, {template: OldDeck});
  NewDeck = ss.getSheets()[OldDeckNum];
  
  // Open Archive Spreadsheet
  var ssArchive = SpreadsheetApp.openById("1m31w7ZDJTDimAnbhWGPo9ulrJy07cpal2lcqje6hU34");
  
  // Copy Old Deck to Archive Spreadsheet
  OldDeck.copyTo(ssArchive);
  
  // Renames the copy
  CopyName = "Copy of " + OldDeckName;
  ArchiveDeck = ssArchive.getSheetByName(CopyName);
  // Removes the "Copy Of"
  ArchiveDeck.setName(OldDeckName);
  ss.deleteSheet(OldDeck);
    
  // Opens the new deck sheet and modify appropriate data (version number, status etc)
  NewDeck = ss.getSheetByName(NewDeckName);
  NewDeck.getRange(1,4).setValue(NewVersionNum);  
  ss.setActiveSheet(NewDeck);
  var Status = NewDeck.getRange(2, 3).getValue();
  fcnUpdateDeckStatus(Status);
  NewDeck.getRange(2, 9).setValue('');
  
}