// CmdrDckOnOpen

// **********************************************
// function OnOpen()
//
// Function executed everytime the Spreadsheet is
// opened or refreshed
//
// **********************************************

function OnOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var FirstSht =  ss.getSheets()[0];
  
  GetStapleCrossRef();
  
  ss.setActiveSheet(FirstSht);
  
  var FuncMenuButtons = [{name: "Generate New Deck Version", functionName: "fcnGenerateNewVersion"}, 
                         {name: "Import Staple Cross Reference",functionName: "GetStapleCrossRef"}, 
                         {name: "Generate Gatherer Links", functionName: "fcnGathererLink"}, 
                         {name: "Clear Gatherer Links", functionName: "fcnClearGathererLink"}, 
                         {name: "Hide Columns", functionName: "fcnHideColumns"}, 
                         {name: "Show Columns", functionName: "fcnShowColumns"}];
  
  var SortMenuButtons = [{name: "Sort Deck by Staple", functionName: "fcnSortDeckStaple"},
                         {name: "Sort Deck by Type/Card Name", functionName: "fcnSortDeckTypeName"},
                         {name: "Sort Deck by Type/Color", functionName: "fcnSortDeckTypeColor"}, 
                         {name: "Sort Deck by Color", functionName: "fcnSortDeckColor"}, 
                         {name: "Sort Deck by Card Name", functionName: "fcnSortDeckCardName"}, 
                         {name: "Sort Deck by Category", functionName: "fcnSortDeckCategory"}, 
                         {name: "Sort Deck by Notes", functionName: "fcnSortDeckNotes"}];
  
  ss.addMenu("General Fctn", FuncMenuButtons);
  ss.addMenu("Sort Fctn", SortMenuButtons);
}