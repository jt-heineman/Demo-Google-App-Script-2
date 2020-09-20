/** @OnlyCurrentDoc */

// Add a custom menu to the active spreadsheet, including a separator and a sub-menu.
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Tracker Tools')
 .addSubMenu(SpreadsheetApp.getUi().createMenu('New Projects')
      .addItem('Create New Project Tracker', 'alertCreateNewTab')
      )

.addSubMenu(SpreadsheetApp.getUi().createMenu('Data')
      .addItem('Consolidate Data', 'alertCallGetDataFromTab')
      .addSeparator()
      .addItem('Consolidation Execution 1/4', 'CallGetDataFromTab_Trigger_1')
      .addItem('Consolidation Execution 2/4', 'CallGetDataFromTab_Trigger_2')
      .addItem('Consolidation Execution 3/4', 'CallGetDataFromTab_Trigger_3')
      .addItem('Consolidation Execution 4/4', 'CallCopyPaste_Trigger')
      )

  .addSubMenu(SpreadsheetApp.getUi().createMenu('Utils')
      .addItem('Protect Current Sheet', 'ProtectActiveSheet')
      .addItem('Fix Broken Formulas on Current Sheet', 'FixFormulasActiveSheet')
      .addItem('Create Backup Copy', 'makeCopy')       
      )
      .addToUi();
};

function alertCallGetDataFromTab(){
//function to start consolitation process
  var sh=SpreadsheetApp.getUi();
  var response=sh.alert("Did you select which tabs to consolidate?",sh.ButtonSet.YES_NO);

  if(response==sh.Button.YES) {  
    CallGetDataFromTab()
  }else{
  
  var spreadsheet = SpreadsheetApp.getActive(); 
    //tab with all sheet names, sheet names generate by custom cell function
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('ADMIN'), true);
    sh.alert("Please select which tabs you would like to consolidate");
  }
}

function CallGetDataFromTab() {
  //cleaning previous range
  var PasteSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  PasteSheet.getRange('A2:AE10000').activate();
  PasteSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  var spreadsheet = SpreadsheetApp.getActive();
  //tab with all sheet names and parameters for managing what tabs to pass in GetDataFromTab
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('ADMIN'), true);
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  //getting tab names to reference in next function 
  for (var i = 0; i < data.length; i++) {
    Logger.log('Tab name: ' + data[i][0]);
    var StrSheetName = data[i][0];  
    //check if tab is needed
    if(data[i][1] === 'Yes' && data[i][0] != ''){
      GetDataFromTab(StrSheetName);
    }
  }
  // write data to sheet
  PasteSheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  Logger.log(k);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Data'), true);
}

function GetDataFromTab(ShName) {
//return updates on global array
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(ShName), true);
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var strType;
  var strTab;
  var strInitiative;
  var strPRC;
  var strProject;
  var strCategory;
  var strExpenseType;
  var strName;
  var dtStart;
  var dtEnd;
  var dtProjStart;
  var dtProjEnd;
  var strDays;
  var strVendor;
  var strInvoiceNum;
  var strBudget;
  var strEstimate;
  var strActuals;
  var strUnit;
  var strRate;
  var strQuantity;
  var strNotes;
  var strCode;
  var StrExpenseMonth = new Date();
  var strChildrenKey;
  var strAssetCount;
  var strGLCode;
  var strCostCenter;
  var strBillTo;
  var strForecast;
  var strFUP;
  var strInitiative;
  var strDateNum;
  var strSavingsCheck;

//loop to get data from tab 
  for (var i = 9; i < data.length; i++) {
          
      if(data[i][1]==='Project'){
        strProject=data[i][3];
      }
      
      if(data[i][1]==='Line Item' || data[i][1]==='Project Budget'){
        strType = data[i][1];
        strCategory = data[i][2];
        strExpenseType = data[i][3];
        strName = data[i][5];
        strVendor = data[i][6];
        strInvoiceNum = data[i][7];
        strBudget = data[i][8];
        strUnit = data[i][9];
        strRate = data[i][10];
        strQuantity = data[i][11];
        strEstimate = data[i][12];
        strActuals = data[i][13];
        strAssetCount = data[i][14];
        strNotes = data[i][15];
        strGLCode = data[i][16];
        strCostCenter = data[i][19];
        strCode = data[i][20];
        StrExpenseMonth = data[i][21];
        strChildrenKey = '' //data[i][3];
        strBillTo = data[i][22];
        strForecast = data[i][23];
        strFUP = data[i][24];
        strSavingsCheck = data[i][25];
        strInitiative = data[3][2];
        
        if (StrExpenseMonth>0){
          strDateNum = StrExpenseMonth.getDate();
        }else{
          strDateNum = ""
        };
             
        // push a row of data as 2d array
        values.push([strType,ShName,strInitiative,'PRC',strProject,strCategory,strExpenseType,strName,strVendor,strInvoiceNum,strBudget,strUnit,strRate,strQuantity,strEstimate,strActuals,strAssetCount,strNotes,strGLCode,strCostCenter,strCode,StrExpenseMonth,strBillTo,strForecast,strFUP,strSavingsCheck,strDateNum,i]);
    }    
  }  
}

//public variable for data consolidation
var k = 2;
var ii = 1;
var values = [];

function CallGetDataFromTab_Trigger_1() {
  GetDataFromTab_Trigger(1);
}

function CallGetDataFromTab_Trigger_2() {
  GetDataFromTab_Trigger(2);
}

function CallGetDataFromTab_Trigger_3() {
  GetDataFromTab_Trigger(3); 
}

function CallCopyPaste_Trigger() {
  //cleaning previous range
  var DestinationSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  DestinationSheet.getRange('A1:AE10000').activate();
  DestinationSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  var spreadsheet = SpreadsheetApp.getActive();
   //tab with all temp data
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Temp Data'), true);
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  // write data to sheet
  DestinationSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Data'), true);
}

function GetDataFromTab_Trigger(runNumber) {
  Logger.log('Call Run Number 2:'+ runNumber);  
  var CallnumRun = runNumber 
  var PasteSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Temp Data');
  var rowPaste;
  var lastRow = goToFirstRowAfterLastRowWithData(PasteSheet,1)

  if(CallnumRun === 1){
    //cleaning previous range
    rowPaste = 2;
    PasteSheet.getRange('A2:AE10000').activate();
    PasteSheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
      
  }else{   
    rowPaste = lastRow;
  }
  
  Logger.log('Call Run Number 3:'+ CallnumRun);
  var spreadsheet = SpreadsheetApp.getActive();
  //tab with all sheet names, sheet names generate by custom cell function
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('ADMIN'), true);  
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  //getting tab names to reference in next iteration 
  for (var i = 0; i < data.length; i++) {
    Logger.log('Tab name: ' + data[i][0] + '-' + data[i][1] + '-' + data[i][2]);
    var StrSheetName = data[i][0];  
    //check if tab is needed
    if(data[i][2] === runNumber && data[i][0] != ''){
    Logger.log('Found:' + CallnumRun + '-' + data[i][2]);
    GetDataFromTab(StrSheetName);
    }    
  }
// write data to sheet
PasteSheet.getRange(rowPaste, 1, values.length, values[0].length).setValues(values); 
//activate sheet
spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Temp Data'), true);
}

function alertCreateNewTab(){
  var sh=SpreadsheetApp.getUi();
  var response=sh.alert("Would you like to proceed creating new Project Trackers? ",sh.ButtonSet.YES_NO);

  if(response==sh.Button.YES) {
  
    CreateNewTab()
      
  }else{
  
  var spreadsheet = SpreadsheetApp.getActive();
  
    //tab with all sheet names, sheet names generate by custom cell function
    spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Intake'), true);
       
    sh.alert("Before proceeding update all projects to add on column S - Tab Name");
     
  }
}

function CreateNewTab() {
  
//intake tab 
var app = SpreadsheetApp;
var ssintake = app.getActiveSpreadsheet().getSheetByName('Intake');
var rangeIntake = ssintake.getDataRange().getValues();
var sh=SpreadsheetApp.getUi();
//loop intake range to get info needed
for (var i = 1; i < rangeIntake.length; i++) {  
    //getting variables    
    var ProjectCode = rangeIntake[i][2];
    var NewSheetName = rangeIntake[i][18];
    var TemplateType = rangeIntake[i][19];
    var CreatedTime = rangeIntake[i][20];
    var Title = rangeIntake[i][1];
    var CPname = rangeIntake[i][8];
    var scope = rangeIntake[i][10];
    var Budget = rangeIntake[i][12];
    var GLCode = rangeIntake[i][13];
    var CostCenter = rangeIntake[i][14];
    var IOCode = rangeIntake[i][15]; 
    var ExpenseTo = rangeIntake[i][16];
        
    //check if tab is needed to add, if sheet name is filled then run 
    if(NewSheetName != '' && CreatedTime === ''){
      
    //run code to add tabs 
      var source = SpreadsheetApp.getActiveSpreadsheet();
      
      //define what template to use
      if(TemplateType === "Default - Simple Project" ){  //|| TemplateType === ''
      
        var TemplateSheet = source.getSheetByName('Template 1');
        var fillRange1 = "S10:S17";
        var fillRange2 = "T10:T17";
        var fillRange3 = "U10:U17";
        var fillRange4 = "V10:V17";
        var fillRange5 = "A10:A16";
        
        Logger.log('1' + TemplateType);

      }else if(TemplateType === "Full Project"){
      
        var TemplateSheet = source.getSheetByName('Template 2');
        var fillRange1 = "S10:S29";
        var fillRange2 = "T10:T29";
        var fillRange3 = "U10:U29";
        var fillRange4 = "V10:V29";
        var fillRange5 = "A10:A28";
        Logger.log('2' + TemplateType);
        
      }else{
      
        var TemplateSheet = source.getSheetByName('Template 1');
        var fillRange1 = "S10:S17";
        var fillRange2 = "T10:T17";
        var fillRange3 = "U10:U17";
        var fillRange4 = "V10:V17";
        var fillRange5 = "A10:A16";
        Logger.log('3' + TemplateType);
            
      };      
      
      Logger.log(source);
      Logger.log(NewSheetName);
      //add new tab based on template
      TemplateSheet.copyTo(source).setName(NewSheetName);
      // write data to sheet
      var PasteSheet = SpreadsheetApp.getActive();
      PasteSheet.setActiveSheet(PasteSheet.getSheetByName(NewSheetName), true);    
      //update header
      PasteSheet.getRange('C5').setValue(ProjectCode);
      PasteSheet.getRange('C6').setValue(CPname);
      PasteSheet.getRange('D4').setValue(IOCode);
      PasteSheet.getRange('D5').setValue(IOCode);
      PasteSheet.getRange('F4').setValue(Budget);
      PasteSheet.getRange('F5').setValue(Budget);
      PasteSheet.getRange('C3').setValue(Title);
      PasteSheet.getRange('D10').setValue(Title);
      PasteSheet.getRange('R6').setValue(scope);
      PasteSheet.getRange('C7').setValue(scope);
      PasteSheet.getRange('E5').setValue(ExpenseTo);
      //update columns
      PasteSheet.getRange('S10').setValue(GLCode);
      var fillDownRange = PasteSheet.getRangeByName(fillRange1); 
      PasteSheet.getRange("S10").copyTo(fillDownRange); 
      PasteSheet.getRange('T10').setValue(CostCenter);
      var fillDownRange = PasteSheet.getRangeByName(fillRange2);
      PasteSheet.getRange("T10").copyTo(fillDownRange);
      PasteSheet.getRange('U10').setValue(IOCode);
      var fillDownRange = PasteSheet.getRangeByName(fillRange3);
      PasteSheet.getRange("U10").copyTo(fillDownRange);
      PasteSheet.getRange('V10').setValue(ExpenseTo);
      var fillDownRange = PasteSheet.getRangeByName(fillRange4); 
      PasteSheet.getRange("V10").copyTo(fillDownRange);
      PasteSheet.getRange('A10').setValue(ProjectCode);
      var fillDownRange = PasteSheet.getRangeByName(fillRange5); 
      PasteSheet.getRange("A10").copyTo(fillDownRange);
      //stamp tab creation
      ssintake.getRange(i+1,21).setValue('Created');
      
      //send email and update check
      var Project = ssintake.getRange(i+1, 2).getValue();
      var Title = "Project Created: " + Project ;
      var Comment = "For your acknowledgement, a project was created, please verify if all is accordingly on the Budget Tracker"
      MailApp.sendEmail("user@company.com",Title , Comment);  
      ssintake.getRange(i+1, 22).setValue("TRUE");

      } 
   }  
}

function CallLockTabs() {

  var spreadsheet = SpreadsheetApp.getActive();
  //tab with all sheet names, sheet names generate by custom cell function
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('TABS'), true);
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  //getting tab names to reference in next function 
  for (var i = 1; i < data.length; i++) {
    
    var StrSheetName = data[i][0].toString();  
    var LockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(StrSheetName);
    Logger.log('Tab name: ' + StrSheetName);
    //calling  each tab
    var dataTab = LockSheet.getDataRange().getValues();
    
    if(dataTab[2][22] === 'Complete'){
    
      var protection = LockSheet.protect().setDescription('Project Complete');
      Logger.log('Lock Tab: ' + StrSheetName);
    }
  }
};

function ProtectActiveSheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var protection = spreadsheet.getActiveSheet().protect();
  protection.setDescription('Project Complete')
};

function CallFixFormulas() {
  var spreadsheet = SpreadsheetApp.getActive();
  //tab with all sheet names, sheet names generate by custom cell function
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('TABS'), true);
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  //getting tab names to be updated 
  for (var i = 0; i < data.length; i++) { 
    Logger.log('Tab name: ' + data[i][0]);
    var StrSheetName = data[i][0];
    //verify if needs update, 0 or 1
    if(data[i][1] === 'Yes'){
    FixFormulas(StrSheetName);
    }
  } 
};

function FixFormulas(StrSheet) {
  var ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(ss.getSheetByName(StrSheet), true);
  ss.getRange("V10").setFormula('=iferror(if(and(MONTH(U10)=$V$3,year(U10)=$V$4),R10&" (Current)",if(U10>$V$5,"2003999 (Prepay)",if(and(U10<=$V$5,M10>0,X10<>true,U10>0),"Paid",if(and(U10<=$V$5,or(M10=0,M10=""),X10=true),"Accrual",if(and(U10<=$V$5,M10>0,X10=true),"Accrual",""))))),"")');
  var lr = ss.getLastRow();
  var fillDownRange = ss.getRangeByName("V10:V200"); //(10,22,lr -12,1);
  ss.getRange("V10").copyTo(fillDownRange);
  ss.getRange("W10").setFormula('=if(and(B10="Line Item",T10<>""), IF(or(M10<>"",M10>0),M10,L10),"")');
  var fillDownRange = ss.getRangeByName("W10:W200"); //(10,22,lr -12,1);
  ss.getRange("W10").copyTo(fillDownRange);

};

function FixFormulasActiveSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //ss.setActiveSheet(ss.getSheetByName(StrSheet), true);
  ss.getRange("V10").setFormula('=iferror(if(and(MONTH(U10)=$V$3,year(U10)=$V$4),R10&" (Current)",if(U10>$V$5,"2003999 (Prepay)",if(and(U10<=$V$5,M10>0,X10<>true,U10>0),"Paid",if(and(U10<=$V$5,or(M10=0,M10=""),X10=true),"Accrual",if(and(U10<=$V$5,M10>0,X10=true),"Accrual",""))))),"")');
  var lr = ss.getLastRow();
  var fillDownRange = ss.getRangeByName("V10:V200"); //(10,22,lr -12,1);
  ss.getRange("V10").copyTo(fillDownRange);
  ss.getRange("W10").setFormula('=if(and(B10="Line Item",T10<>""), IF(or(M10<>"",M10>0),M10,L10),"")');
  var fillDownRange = ss.getRangeByName("W10:W200"); //(10,22,lr -12,1);
  ss.getRange("W10").copyTo(fillDownRange);
 
};

function makeCopy() {
  //function to create backups of entire application in Google Drive 
  // generates the timestamp and stores in variable formattedDate as year-month-date hour-minute-second
  var formattedDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd' 'HH:mm:ss");
  // gets the name of the original file and appends the word "copy" followed by the timestamp stored in formattedDate
  var name = SpreadsheetApp.getActiveSpreadsheet().getName() + " Copy " + formattedDate;
  // gets the destination folder by their ID. REPLACE xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx with your folder's ID that you can get by opening the folder in Google Drive and checking the URL in the browser's address bar
  var destination = DriveApp.getFolderById("1RI0xPtGou3Ju7uUCV_Tbar7rI6-pWvWe");
  // gets the current Google Sheet file
  var file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId())
  // makes copy of "file" with "name" at the "destination"
  file.makeCopy(name, destination);
//folder path
//https://drive.google.com/open?id=1RI0xPtGou3Ju7uUCV_Tbar7rI6-pWvWe
}

/**
 * Gets the Sheet Name of a selected Sheet.
 * @param {number} option 0 - Current Sheet, 1  All Sheets, 2 Spreadsheet filename
 * @return The input multiplied by 2.
 * @customfunction
 */

function SHEETNAME(option) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var thisSheet = sheet.getName();
  
  //Current option Sheet Name
  if(option === 0){
    return thisSheet;
  
  //All Sheet Names in Spreadsheet
  }else if(option === 1){
    var sheetList = [];
    ss.getSheets().forEach(function(val){
       sheetList.push(val.getName())
    });
    return sheetList;
  
  //The Spreadsheet File Name
  }else if(option === 2){
    return ss.getName();
  
  //Error  
  }else{
    return "#N/A";
  };
};

function goToFirstRowAfterLastRowWithData(sheet, colNum) {
  var v = sheet.getRange(1, colNum, sheet.getLastRow()).getValues(),
      l = v.length,
      r;
  while (l > 0) {
    if (v[l] && v[l][0].toString().length > 0) {
      r = (l + 2);
      break;
    } else {
      l--;
    }
  }
  return r || 1;
}
