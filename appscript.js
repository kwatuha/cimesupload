var HEADERS = {
  "Authorization": "Basic " + Utilities.base64Encode("root" + ":" + "reset123")
};

var SERVER = "http://134.122.4.64/cloud/ws/api/smartcounty/";
var attachmentUrl=SERVER+"attachment/types?IsVoided=0&_=1714240431638"; 
var usersUrl=SERVER+"system/users?IsVoided=0&_=1714384904093"; 
var wardsUrl=SERVER+"wards?_=1714384904095"; 
var departmentsUrl=SERVER+"county/department?_=1714384904096"; 
var subcountiesUrl=SERVER+"subcounties?_=1714392967300";
var projectTypesUrl=SERVER+"project/types?IsVoided=0&_=1714466326993";
var projectProgressUrl=SERVER+"project/progress/status?IsVoided=0&_=1714466326992";
var projectMilestoneUrl=SERVER+"project/status?IsVoided=0&_=1714466326994";
var cidpRegistryUrl=SERVER+"cidp/registry?IsVoided=0&_=1714488272905";
var financialYearUrl=SERVER+"financial/years?IsVoided=0&_=1714488272906";
var contractorsUrl=SERVER+"contractors?IsVoided=0&_=1714488272910";
var deptSectionsUrl=SERVER+"department/sections?IsVoided=0&_=1714488272913";
var sectorStrategiesUrl=SERVER+"sector/strategies?IsVoided=0&_=1714712317274";
var subProgramsUrl=SERVER+"strategy/subprogramme?IsVoided=0&_=1714712317276";
var projectsSearchUrl=SERVER+"projects?IsVoided=0&_=1714820290993";
var projectsPostUrl=SERVER+"projects";
//var DEPT_SECTIONS= fetchUrlData('dept_directorates');



//myLoginScren();
//addProjectsDataRangeFirst();
addProjectsDataRangeNewFirst()
 
addCIMESMenu() ;
importDeptDirectorates();

function getMetaData() {
      importWards();
      importProjectTypes();
      importProjectStatus();
      importProjectMilestone();
      importDepartments();      
      importCidp(); 
      importFinancialYears(); 
      importContractors(); 
      importDeptDirectorates();
      importSectorPrograms();
      importSectorSubPrograms(); 
}
function fetchUrlData(type){
  var urlType='';
 switch(type){
   case 'wards': urlType=wardsUrl;
   break;
   case 'users': urlType=usersUrl;
   break;
   case 'departments': urlType=departmentsUrl;
   break;
   case 'attachment_types': urlType=attachmentUrl;
   break; 
   case 'subcounties': urlType=subcountiesUrl;
   break;

   case 'project_types': urlType=projectTypesUrl;
   break;
   case 'progress_status': urlType=projectProgressUrl;
   break; 
   case 'BoQ_phase_milestones': urlType=projectMilestoneUrl;
   break;
   case 'cidp': urlType=cidpRegistryUrl;
   break;
   case 'financial_year': urlType=financialYearUrl;
   break;
   case 'contractors': urlType=contractorsUrl;
   break;
   case 'dept_directorates': urlType=deptSectionsUrl;
   break;
   case 'sector_programs': urlType=sectorStrategiesUrl;
   break;
   case 'sect_sub_programs': urlType=subProgramsUrl;
   break;
   case 'existing_projects': urlType=projectsSearchUrl;
   break;
   
 }
 if(!isEmpty(type)){
   return getData(urlType);
 }
}


function getAttachmentTypes(){
var url=SERVER+"attachment/types?IsVoided=0&_=1714240431638"; 
var options = {
  headers: HEADERS,
  muteHttpExceptions: true
};
var response = UrlFetchApp.fetch(url, options);
var content = response.getContentText();
var results = JSON.parse(content).result
//Logger.log(results)
return results;
}

function getData(url){
var options = {
  headers: HEADERS,
  muteHttpExceptions: true
};
var response = UrlFetchApp.fetch(url, options);
var content = response.getContentText();
var results = JSON.parse(content).result
//Logger.log(results)
return results;
}

function importDepartments() {
var results = fetchUrlData('departments');

// Clear existing data in columns A to C of the "main" sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("departments");
sheet.getRange("A:B").clear(); // Assuming you want to clear columns A to C, adjust as needed

// Write headers
sheet.getRange("A1:B1").setValues([["Name","DepartmentID"]]);

// Write data to columns
var rowData = [];
results.forEach(function(item, index) {
  rowData.push([item.Name,item.DepartmentID]);
});
sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

}

function importUsers() {
var results = fetchUrlData('users');

// Clear existing data in columns A to C of the "main" sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("users");
sheet.getRange("A:B").clear(); // Assuming you want to clear columns A to C, adjust as needed

// Write headers
sheet.getRange("A1:B1").setValues([["UserName","UserID"]]);

// Write data to columns
var rowData = [];
results.forEach(function(item, index) {
  rowData.push([item.UserName,item.UserID]);
});
sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

}
function importWards() {
var results = fetchUrlData('wards');
// Clear existing data in columns A to C of the "main" sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("wards");
sheet.getRange("A:D").clear(); // Assuming you want to clear columns A to C, adjust as needed

// Write headers
sheet.getRange("A1:D1").setValues([["Name","SubCountyID", "WardID","SubCounty"]]);

// Write data to columns
var rowData = [];
results.forEach(function(item, index) {
  rowData.push([item.Name,item.SubCountyID, item.WardID,item.SubCountyID]);
});
sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

}


function importProjectTypes() {
var results = fetchUrlData('project_types');

// Clear existing data in columns A to C of the "main" sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("project_types");
sheet.getRange("A:C").clear(); // Assuming you want to clear columns A to C, adjust as needed

// Write headers
sheet.getRange("A1:C1").setValues([["TypeName","TypeID",  "Description"]]);

// Write data to columns
var rowData = [];
results.forEach(function(item, index) {
  rowData.push([item.TypeName,item.TypeID,  item.Description]);
});
sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

}

function importProjectStatus() {
var results = fetchUrlData('progress_status');

// Clear existing data in columns A to C of the "main" sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("progress_status");
sheet.getRange("A:C").clear(); // Assuming you want to clear columns A to C, adjust as needed

// Write headers
sheet.getRange("A1:C1").setValues([["Status","StatusID",  "MandEAction"]]);

// Write data to columns
var rowData = [];
results.forEach(function(item, index) {
  rowData.push([item.Status,item.StatusID,  item.MandEAction]);
});
sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

}

function importCidp() {
    var results = fetchUrlData('cidp');
   
    // Clear existing data in columns A to C of the "main" sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("cidp");
    sheet.getRange("A:E").clear(); // Assuming you want to clear columns A to C, adjust as needed
    
    // Write headers
    sheet.getRange("A1:E1").setValues([["CIDPName", "CIDPID", "StartDate", "EndDate","Theme"]]);
    
    // Write data to columns
    var rowData = [];
    results.forEach(function(item, index) {
      rowData.push([item.CIDPName,item.CIDPID,  item.StartDate, item.EndDate,item.Theme]);
    });
    sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);
    
    }
    
    
    function importFinancialYears() {
    var results = fetchUrlData('financial_year');
    
    // Clear existing data in columns A to C of the "main" sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("financial_year");
    sheet.getRange("A:D").clear(); // Assuming you want to clear columns A to C, adjust as needed
    
    // Write headers
    sheet.getRange("A1:D1").setValues([["FinYearName","FinYearID",  "StartDate", "EndDate"]]);
    
    // Write data to columns
    var rowData = [];
    results.forEach(function(item, index) {
      rowData.push([item.FinYearName,item.FinYearID,  item.StartDate, item.EndDate]);
    });
    sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);
    
    }
    
    
    function importContractors() {
    var results = fetchUrlData('contractors');
    
    // Clear existing data in columns A to C of the "main" sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("contractors");
    sheet.getRange("A:D").clear(); // Assuming you want to clear columns A to C, adjust as needed
    
    // Write headers
    sheet.getRange("A1:D1").setValues([["CompanyName","ContractorID",  "ContactName", "Phone"]]);
    
    // Write data to columns
    var rowData = [];
    results.forEach(function(item, index) {
      rowData.push([item.CompanyName,item.ContractorID,  item.ContactName, item.Phone]);
    });
    sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);
    
    }
   
    function importSectorPrograms() {
    var results = fetchUrlData('sector_programs');
    
    // Clear existing data in columns A to C of the "main" sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sector_programs");
    sheet.getRange("A:D").clear(); // Assuming you want to clear columns A to C, adjust as needed
    
    // Write headers
    sheet.getRange("A1:D1").setValues([["Programme","StrategyID", "CIDPID", "DepartmentID"]]);
    
    // Write data to columns
    var rowData = [];
    results.forEach(function(item, index) {
      rowData.push([item.Programme,item.StrategyID, item.CIDPID, item.DepartmentID]);
    });
    sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);
    
    }
    
  
    function importSectorSubPrograms() {
    var results = fetchUrlData('sect_sub_programs');
    
    // Clear existing data in columns A to C of the "main" sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sect_sub_programs");
    sheet.getRange("A:D").clear(); // Assuming you want to clear columns A to C, adjust as needed
    
    // Write headers
    sheet.getRange("A1:D1").setValues([["SubProgramme","SubProgrammeID", "StrategyID",  "KeyOutcome"]]);
    
    // Write data to columns
    var rowData = [];
    results.forEach(function(item, index) {
      rowData.push([item.SubProgramme,item.SubProgrammeID, item.StrategyID,  item.KeyOutcome]);
    });
    sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);
    
    }
    

    function importDeptDirectorates() {
    var results = fetchUrlData('dept_directorates');
    var deptResults = fetchUrlData('departments');
    
    // Clear existing data in columns A to C of the "main" sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dept_directorates");
    sheet.getRange("A:D").clear(); // Assuming you want to clear columns A to C, adjust as needed
    var foundDpt = 
    // Write headers
    sheet.getRange("A1:D1").setValues([[ "Name","SectionID", "Department Name","DepartmentID" ]]);
    
    // Write data to columns
    var rowData = [];
    results.sort((a, b) => a.DepartmentID - b.DepartmentID);
    results.forEach(function(item, index) {
        var deptid=item.DepartmentID;
        var deptName='Unknown';
        if(deptid<12)
        deptName=deptResults.filter(function (entry) { return entry.DepartmentID === deptid; })[0].Name;
      rowData.push([item.Name,item.SectionID, deptName, item.DepartmentID ]);
    });
    sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);
    
    }


function importProjectMilestone() {
var results = fetchUrlData('BoQ_phase_milestones');

// Clear existing data in columns A to C of the "main" sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BoQ_phase_milestones");
sheet.getRange("A:E").clear(); // Assuming you want to clear columns A to C, adjust as needed

// Write headers
sheet.getRange("A1:E1").setValues([["Status","StatusID", "TypeID",  "Percentage", "Description"]]);

// Write data to columns
var rowData = [];

results.sort(function (a, b) {
  let af = a.TypeID;
  let bf = b.TypeID;
  let as = a.Percentage;
  let bs = b.Percentage;

  // If first value is same
  if (af == bf) {
      return (as < bs) ? -1 : (as > bs) ? 1 : 0;
  } else {
      return (af < bf) ? -1 : 1;
  }
});


results.forEach(function(item, index) {
  rowData.push([item.Status,item.StatusID, item.TypeID,  item.Percentage, item.Description]);
});
sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

}




function customImportJSONWithAuth() {

var results = getAttachmentTypes()

// Clear existing data in columns A to C of the "main" sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("attachment_types");
sheet.getRange("A:C").clear(); // Assuming you want to clear columns A to C, adjust as needed

// Write headers
sheet.getRange("A1:C1").setValues([["CreatedOn", "AttachmentName", "Description"]]);

// Write data to columns
var rowData = [];
results.forEach(function(item, index) {
  rowData.push([item.CreatedOn, item.AttachmentName, item.Description]);
});
sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

}

// TODO make this function to accept array
function isEmpty(value) {
return (value == null || (typeof value === "string" && value.trim().length === 0));
}

function uddateAttachmentTypes() {

// get current attachment types
var existingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("attachment_types");
var missingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("missing_attachment_types");
// Get the data range containing all rows
var allMissingatAtchmentTypes = missingSheet.getDataRange().getValues();
var allExistingatAtchmentTypes = existingSheet.getDataRange().getValues();

// TODO  (remove duplicates) get rows in allMissingatAtchmentTypes that does not exist in allExistingatAtchmentTypes an save as in var called filtered
 var filteredAttachmentTypes = allMissingatAtchmentTypes

 filteredAttachmentTypes.forEach(function(item, index) {
  // TODO: remove column headers
  if(!isEmpty(item[0]) && !isEmpty(item[1])){
      console.log("is not empty" )
      var formattedData = {
          CreatedOn: "2018-09-03T18:02:09.000Z", // Assuming the first column contains the CreatedOn data
          AttachmentName: item[0], // Assuming the second column contains the AttachmentName data
          Description: item[1] // Assuming the third column contains the Description data
        };

        // Example: Post data to server
        var url = SERVER+"attachment/types";
        var options = {
          headers: HEADERS,
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify(formattedData),
          muteHttpExceptions: true
        };

        var response = UrlFetchApp.fetch(url, options);
       //Logger.log(response.getContentText()); // Log the response from the server
  }
  
});



}

function myLoginScren() {
  // Create a user interface for the sidebar
  var htmlOutput = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_top">
        <H6> CIMES LOGIN</H6>
      </head>
      <body>
        <div>Username:<input type="text" id="username" /></div>
        <div>Password:<input type="password" id="password" /></div>  
        <button onclick="handleButtonClick()">Login</button>
        <script>
          function handleButtonClick() {
            var username=document.getElementById("username").value;
            var password=document.getElementById("password").value;
            google.script.run.doSomething(username,password); // Call server-side function when button is clicked
          }
        </script>
      </body>
    </html>`);
  
  // Show the sidebar
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// Function to be called when the button is clicked
function doSomething(username,password) {
  // Add your desired functionality here
  var details=getUserDetails(username,password)
  SpreadsheetApp.getActiveSpreadsheet().toast('Button Clicked!='+details.length+'==='+username+'=pwd='+password, 'Status', 3);
  //SpreadsheetApp.getActiveSpreadsheet().toast(details.length, 'Status', 3);
}

function hideSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = "Sheet1"; // Replace "Sheet1" with the name of the sheet you want to hide

  // Get the sheet by name
  var sheet = spreadsheet.getSheetByName(sheetName);

  if (sheet) {
    // Hide the sheet
    sheet.hideSheet();
    //Logger.log('Sheet hidden: ' + sheetName);
    SpreadsheetApp.getActiveSpreadsheet().toast('Sheet hidden: ' + sheetName, 'Status', 3);
  } else {
   // Logger.log('Sheet not found: ' + sheetName);
    SpreadsheetApp.getActiveSpreadsheet().toast('Sheet not found: ' + sheetName, 'Error', 3);
  }
}

function hideColumn() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Sheet1"); // Replace "Sheet1" with your sheet name
  var columnIndex = 2; // Column index to hide (e.g., column B is index 2)
  var numColumns = 1; // Number of columns to hide

  // Hide the specified columns
  sheet.hideColumns(columnIndex, numColumns);

 // Logger.log('Columns hidden: ' + columnIndex + ' to ' + (columnIndex + numColumns - 1));
  SpreadsheetApp.getActiveSpreadsheet().toast('Columns hidden: ' + columnIndex + ' to ' + (columnIndex + numColumns - 1), 'Status', 3);
}



function setDropdownValidation() {
  //var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("testme");;
  var range = sheet.getRange("wards!C1:C55"); // Specify the range containing dropdown options

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(range)
    .setAllowInvalid(false)
    .build();
    var sheetName="testme";
    var targetColumn=2;//B
    var numberofvalues=getColumnValues(sheetName,targetColumn);
    numberofvalues=numberofvalues+5;
    var targetplace="B2:B"+numberofvalues;

  var targetRange = sheet.getRange(targetplace); // Specify the target range for data validation
  targetRange.setDataValidation(rule);
}


function getColumnValues(sheetName,targetColumn) {
  var sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var column = 1; // Specify the column index (A=1, B=2, C=3, etc.)
  
  // Get the range for the entire column
  var range = sheet.getRange(1, targetColumn, sheet.getLastRow(), 1); // Start from row 1 to the last row
  
  // Get values from the column range
  var values = range.getValues();
  return values.length;
}

  
  function getUserDetails(username,password) {
    var url=SERVER+"system/users?IsVoided=0&_=1714384904093";
    var custheader = {
        "Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)
      };

    var options = {
        headers: custheader,
        muteHttpExceptions: true
      };
      var response = UrlFetchApp.fetch(url, options);
      var content = response.getContentText();
      var results = JSON.parse(content).result
      Logger.log("UUUUUUUUUUUUUUUUUUU===");
      Logger.log(results);
    
return results
    }

function onEdit(e) {
       //filterDropDownListEdit(e);
        //updateMetaData(e);
       // setDropDownDataRange(e);
}

function updateMetaData(e){
        var editedRange = e.range;
        var editedColumn = editedRange.getColumn();
        var sheet =editedRange.getSheet().getSheetName();
        switch(sheet){
          case 'wards': importWards();
          break;
          case 'users': importUsers();
          break;
          case 'departments': importDepartments();
          break; 
          case 'progress_status': importProjectStatus();
          break; 
          case 'BoQ_phase_milestones': importProjectMilestone();
          break;
          case 'cidp':importCidp();
          break;
          case 'financial_year': importFinancialYears();
          break;
          case 'contractors': importContractors();
          break;
          case 'dept_directorates': importDeptDirectorates();
          break;
    
        }
}

function setDropDownDataRange(e) {
    var editedRange = e.range;
    var sheet =editedRange.getSheet().getSheetName();
      switch(sheet){
          case 'projects': addProjectsDataRange(e);
          break;
          case 'monitoring': addMonitorDataRange(e);
          break;
    
        }
}

function addProjectsDataRangeFa(e){
 /* var editedRange = e.range;
  var editedsheet =editedRange.getSheet().getSheetName();
  var editedColumn = editedRange.getColumn();
  var editedRow = editedRange.getRow();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("projects");;
  var dataRangeSource = null;
  // sheet.getRange("wards!C1:C55");//source of dropdownlist
  editedRow=getMaxDataRow('projects');
  var targetplace=getTargetPlace(editedColumn,editedRow);
  //      sheet.getRange("A:W").clear();
  if(editedsheet==='projects'){

    sheet.getRange("A1:W1").setValues([[
        "CIDPID",
        "FinYearID",
        "DepartmentID",
        "SectionID",
        "StrategyID",
        "SubProgrammeID",
        "TypeID",
        "ProjectName",
        "Description",
        "IsFlagship",
        "Status",
        "ProjStatus",
        "StartDate",
        "EndDate",
        "ContractorID",
        "TenderNum",
        "ContractSum",
        "Budget",
        "Funding",
        "SubCountyID",
        "WardID",
        "Location",
        "CreatedOn"
      ]]);
 
}*/
}

function addProjectsDataRange(e){
        var editedRange = e.range;
        var editedsheet =editedRange.getSheet().getSheetName();
        var editedColumn = editedRange.getColumn();
        var editedRow = editedRange.getRow();
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("projects");;
        var dataRangeSource = null;
        // sheet.getRange("wards!C1:C55");//source of dropdownlist
        editedRow=getMaxDataRow('projects');
        var targetplace=getTargetPlace(editedColumn,editedRow);
        //      sheet.getRange("A:W").clear();
        if(editedsheet==='projects'){
    
          sheet.getRange("A1:X1").setValues([[
              "CIDPID",
              "FinYearID",
              "DepartmentID",
              "SectionID",
              "StrategyID",
              "SubProgrammeID",
              "TypeID",
              "ProjectName",
              "Description",
              "IsFlagship",
              "Status",
              "ProjStatus",
              "StartDate",
              "EndDate",
              "ContractorID",
              "TenderNum",
              "ContractSum",
              "Budget",
              "Funding",
              "SubCountyID",
              "WardID",
              "Location",
              "CreatedOn",
              "StausUpdate"
            ]]);

            if(editedColumn==1){
              dataRangeSource=sheet.getRange("cidp!A2:A55");
              targetplace = getTargetPlace(1,editedRow);
              addColumnRule(dataRangeSource,targetplace,sheet);
            }
          
            if(editedColumn==2){
            dataRangeSource=sheet.getRange("financial_year!A2:A55");
            targetplace = getTargetPlace(2,editedRow);
            addColumnRule(dataRangeSource,targetplace,sheet);
            }
            if(editedColumn==3){
            dataRangeSource=sheet.getRange("dept_directorates!D2:D55"); 
            targetplace = getTargetPlace(3,editedRow);
            addColumnRule(dataRangeSource,targetplace,sheet);
          }
          if(editedColumn==4){
            dataRangeSource=sheet.getRange("dept_directorates!A2:A55"); 
            targetplace = getTargetPlace(4,editedRow);
            addColumnRule(dataRangeSource,targetplace,sheet);
          }
          if(editedColumn==5){
            dataRangeSource=sheet.getRange("sector_programs!D2:D55");
            targetplace = getTargetPlace(5,editedRow);
            addColumnRule(dataRangeSource,targetplace,sheet);
          }
          if(editedColumn==6){
            dataRangeSource=sheet.getRange("sect_sub_programs!C2:C55");
            targetplace = getTargetPlace(6,editedRow);
            addColumnRule(dataRangeSource,targetplace,sheet);
          }
          if(editedColumn==7){
            dataRangeSource=sheet.getRange("project_types!A2:A55");
            targetplace = getTargetPlace(7,editedRow);
            addColumnRule(dataRangeSource,targetplace,sheet); 
          }
          if(editedColumn==11){
            dataRangeSource=sheet.getRange("progress_status!A2:A55");
            targetplace = getTargetPlace(11,editedRow);
            addColumnRule(dataRangeSource,targetplace,sheet);
          }
          if(editedColumn==15){
            dataRangeSource=sheet.getRange("BoQ_phase_milestones!C2:C55");
            targetplace = getTargetPlace(15,editedRow);
            addColumnRule(dataRangeSource,targetplace,sheet); 
          }
          if(editedColumn==15){
            dataRangeSource=sheet.getRange("contractors!A2:A55");
            targetplace = getTargetPlace(15,editedRow);
            addColumnRule(dataRangeSource,targetplace,sheet);
          
          }
          if(editedColumn==20){   dataRangeSource=sheet.getRange("wards!A2:A55");
            targetplace = getTargetPlace(20,editedRow);
            addColumnRule(dataRangeSource,targetplace,sheet);
          }
          
          if(editedColumn==21){
            dataRangeSource=sheet.getRange("wards!C2:C55");
            targetplace = getTargetPlace(21,editedRow);
            addColumnRule(dataRangeSource,targetplace,sheet);           
           }
          
         }
      
       
}


function addProjectsDataRangeFirst(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("projects");
  sheet.getRange("A1:X1").setValues([[
        "CIDPID",
        "FinYearID",
        "DepartmentID",
        "SectionID",
        "StrategyID",
        "SubProgrammeID",
        "TypeID",
        "ProjectName",
        "Description",
        "IsFlagship",
        "Status",
        "ProjStatus",
        "StartDate",
        "EndDate",
        "ContractorID",
        "TenderNum",
        "ContractSum",
        "Budget",
        "Funding",
        "SubCountyID",
        "WardID",
        "Location",
        "CreatedOn",
        "StatusUpdate"
      ]]);

      dataRangeSource=sheet.getRange("cidp!A2:A55");
      targetplace = getTargetPlace(1,2);
      addColumnRule(dataRangeSource,targetplace,sheet);
      
      dataRangeSource=sheet.getRange("financial_year!A2:A55");
      targetplace = getTargetPlace(2,2);
      addColumnRule(dataRangeSource,targetplace,sheet);
      
      dataRangeSource=sheet.getRange("dept_directorates!C2:C55"); 
      targetplace = getTargetPlace(3,2);
      addColumnRule(dataRangeSource,targetplace,sheet);
      
      dataRangeSource=sheet.getRange("dept_directorates!A2:A55"); 
      targetplace = getTargetPlace(4,2);
      addColumnRule(dataRangeSource,targetplace,sheet);
      
      dataRangeSource=sheet.getRange("sector_programs!D2:D55");
      targetplace = getTargetPlace(5,2);
      addColumnRule(dataRangeSource,targetplace,sheet);
      
      dataRangeSource=sheet.getRange("sect_sub_programs!C2:C55");
      targetplace = getTargetPlace(6,2);
      addColumnRule(dataRangeSource,targetplace,sheet);
      
      dataRangeSource=sheet.getRange("project_types!A2:A55");
      targetplace = getTargetPlace(7,2);
      addColumnRule(dataRangeSource,targetplace,sheet); 
      
      dataRangeSource=sheet.getRange("progress_status!A2:A55");
      targetplace = getTargetPlace(11,2);
      addColumnRule(dataRangeSource,targetplace,sheet);
      
      dataRangeSource=sheet.getRange("BoQ_phase_milestones!C2:C55");
      targetplace = getTargetPlace(15,2);
      addColumnRule(dataRangeSource,targetplace,sheet); 
      
      dataRangeSource=sheet.getRange("contractors!A2:A55");
      targetplace = getTargetPlace(15,2);
      addColumnRule(dataRangeSource,targetplace,sheet);
      
      dataRangeSource=sheet.getRange("wards!A2:A55");
      targetplace = getTargetPlace(20,2);
      addColumnRule(dataRangeSource,targetplace,sheet);
      
      dataRangeSource=sheet.getRange("wards!C2:C55");
      targetplace = getTargetPlace(21,2);
      addColumnRule(dataRangeSource,targetplace,sheet);           
      
 
}

function addColumnRule(dataRangeSource,targetplace,sheet){
  var rule = SpreadsheetApp.newDataValidation()
  .requireValueInRange(dataRangeSource)
  .setAllowInvalid(false)
  .build();
var targetRange = sheet.getRange(targetplace); // Specify the target range for data validation

targetRange.setDataValidation(rule);
}

function getMaxDataRow(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  // Get the data range of the sheet
  var range = sheet.getDataRange();
  
  // Get the values in the range
  var values = range.getValues();
  
  // Initialize maxRow to 0
  var maxRow = 0;
  
  // Iterate through each row in the values array
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    
    // Check if the entire row is empty
    var rowEmpty = true;
    for (var j = 0; j < row.length; j++) {
      if (row[j] !== "") {
        rowEmpty = false;
        break;
      }
    }
    
    // If the row is not empty, update maxRow
    if (!rowEmpty) {
      maxRow = i + 1; // Use i + 1 to get the correct row index
    }
  }
  
  return maxRow;
}




function getTargetPlace(editedColumn,editedRow){
var cellRef='';
  switch(editedColumn){
    case 	1: cellRef="A";
    break;
   case 	2: cellRef="B";  
    break;
   case 	3: cellRef="C";
    break;
   case 	4: cellRef="D";
    break;
   case 	5: cellRef="E";
    break;
   case 	6: cellRef="F";
    break;
   case 	7: cellRef="G";
    break;
   case 	8: cellRef="H";
    break;
   case 	9: cellRef="I";
    break;
   case 	10: cellRef="J";
    break;
   case 	11: cellRef="K";
    break;
   case 	12: cellRef="L";
    break;
   case 	13: cellRef="M";
    break;
   case 	14: cellRef="N";
    break;
   case 	15: cellRef="O";
    break;
   case 	16: cellRef="P";
    break;
   case 	17: cellRef="Q";
    break;
   case 	18: cellRef="R";
    break;
   case 	19: cellRef="S";
    break;
   case 	20: cellRef="T";
    break;
   case 	21: cellRef="U";
    break;
   case 	22: cellRef="w";
    break;
   case 	23: cellRef="X";
    break;

  }
  
  var nextRows=countRowsInColumn('projects', cellRef);
  nextRows=nextRows+3;
  nextRows=1000;
 return  cellRef+'2:'+cellRef+nextRows;

}


function countRowsInColumn(sheetName, columnName) {
  // Access the active spreadsheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Access the specified sheet by name
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log('Sheet not found');
    return;
  }
  
  // Get the range of data in the specified column
  var columnRange = sheet.getRange(columnName + '1:' + columnName + sheet.getLastRow());
  
  // Get all values in the column
  var columnValues = columnRange.getValues();
  
  // Count non-empty cells in the column
  var rowCount = 0;
  for (var i = 0; i < columnValues.length; i++) {
    if (columnValues[i][0] !== "") {
      rowCount++;
    }
  }
  
  return rowCount;
}



function addMonitorDataRange(e){
}

function addNewProject(){
  var payload={
  ProjectID:0,
  TypeID:1,
  ProjectName:	"Tests",
  TenderNum:	"",
  Description:	"wwewe",
  CIDPID:	"2",
  StrategyID:	"1",
  SubProgrammeID	:"1",
  FinYearID	:"1",
  DepartmentID	:"10",
  OutputKPI	:":",
  ADPOutputMetric	:"",
  ADPBaseline	:"6",
  ADPTarget	:"4",
  QTR2Targets	:"1",
  QTR1Targets	:"1",
  QTR3Targets	:"1",
  QTR4Targets	:"1",
  TenderedTarget	:"0",
  FYActualOutput	:"0",
  QTR1ActualOutput	:"0",
  QTR2ActualOutput	:"0",
  QTR3ActualOutput	:"0",
  QTR4ActualOutput	:"0",
  ProjStatus	:"",
  Percentage	:"",
  Status	:"New",
  Staging	:"Project",
  Funding	:"CG",
  IsFlagship	:"true",
  KPI	:"1",
  KeyOutcome	:"2",
  ContractorID	:"1",
  SubCountyID	:"2",
  WardID	:"9",
  Location	:"west+kisumu",
  UserID	:"4",
  Budget	:"5000000",
  PercBudgetRecvd	:"100",
  ContractSum	:"40000",
  TotAmountPaid	:"0",
  StartDate	:"Fri+May+03+2024+07:58:37)",
  EndDate	:"Fri+May+03+2024+07:58:37",
  CreatedBy	:"1",
  UpdatedBy	:"1",
  VoidedBy	:"1",
  LastUpdate	:"Fri+May+03+2024+07:58:37",
  CreatedOn	:"Fri+May+03+2024+07:58:37",
  SectionID :"d",
  }
  
  }

  function postProjectData(data) {
    
     data.forEach(function(item, index) {
      // TODO: remove column headers
      item["PercBudgetRecvd"]=100;
      item["Staging"]="Project";
      Logger.log('ooooooooooooo=spRowNumber='+item);
      Logger.log(item);
      Logger.log('=========='+item);
      Logger.log(JSON.stringify(item));
      Logger.log('END=spRowNumber='+item);
       var rowNum=2;
       var cItem=item;
       if(cItem.spRowNumber) rowNum=cItem.spRowNumber+1;
      var url = SERVER+"projects";
        var options = {
              headers: HEADERS,
              method: "post",
              contentType: "application/json",
              payload: JSON.stringify(item),
              muteHttpExceptions: true
            };
    
            var response = UrlFetchApp.fetch(url, options);
            var res=JSON.parse(response.getContentText());
            
             Logger.log(res); 
             Logger.log(res.result);
             Logger.log('ooooooooooooo=rowNum='+rowNum);
                if(res){
                //if(res.result.hasOwnProperty('insertId')) {
                 updateStatusAfterUpload(rowNum, res.result.insertId);
                  // updateStatusAfterUpload(spRowNumber,insertId)
                //}
                }

                
           Logger.log(response.getContentText()); // Log the response from the server
     
      
    });
  }  
  function updateStatusAfterUpload(spRowNumber,insertId){
    //{"startIndex":0,"result":{"fieldCount":0,"affectedRows":1,"insertId":36,"serverStatus":2,"warningCount":0,"message":"","protocol41":true,"changedRows":0}}
    var value='uploaded:'+insertId;
    setValueToCell('projects',24,spRowNumber,value);

  }

  function setValueToCell(sheetName,column,row,value) {
    // Access the active spreadsheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName); // Change "Sheet1" to your sheet name
    var cell = sheet.getRange(row, column);
    cell.setValue(value); // Replace "Hello, World!" with the value you want to set
  }

  
  function addCIMESMenu() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('E-CIMES')
      .addItem('Upload Project', 'showMessage')
      .addItem('Upload Contractors', 'insertTimestamp')
      .addItem('Upload Payments', 'insertTimestamp')
      .addItem('Update Meta Data', 'updateMetaData')
      .addToUi();
  }
 
function updateMetaData(){
  getMetaData();
  SpreadsheetApp.getActiveSpreadsheet().toast('System Meta Data Has Been Update', 'Message', 3000);
}


  function showMessage() {
    var existingProjects = fetchUrlData('existing_projects');
  
    //Logger.log('Existing found');
    //Logger.log(existingProjects);
    //Logger.log('Entered Data');

   var newData= fetchDataIntoArrayObjects('projects') ;
 

   postProjectData(newData);
    SpreadsheetApp.getActiveSpreadsheet().toast('Project Data Uploaded', 'Message', 3000);
  }
  
  
  function importSystemProjects() {
    /*var results = fetchUrlData('existing_projects');
    
    // Clear existing data in columns A to C of the "main" sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("system_data");
    sheet.getRange("A:C").clear(); // Assuming you want to clear columns A to C, adjust as needed
    
    // Write headers
    sheet.getRange("A1:C1").setValues([["ProjectName","ProjectID",  "Description"]]);
    
    // Write data to columns ProjectName Description ProjectID
    var rowData = [];
    results.forEach(function(item, index) {
      rowData.push([item.ProjectName,item.ProjectID,  item.Description]);
    });
    sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);
    */
  }
  
  

  function insertTimestamp() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var cell = sheet.getActiveCell();
    cell.setValue(new Date());
  }


  function fetchDataIntoArrayObjects(sheetName) {
  
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var range = sheet.getDataRange();
    //var range = sheet.getRange("A1:X1");
    var values = range.getValues();
   
    // Get column headers (field names) from the first row
    var headers = values[0];


   
  
    // Initialize an array to store objects
    var dataArray = [];
  
    // Iterate over each row starting from the second row (index 1)
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var dataObject = {};
  
      // Iterate over each column
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        var value = row[j];
  
        // Use column header as key and cell value as value in the object
        dataObject[header] = removeDecimalPlaces (header, value) ;

        //Logger.log('Entered Insidet Data Each Object');
       // Logger.log(dataObject);
        //Logger.log('Entered Outside Data Endo of Object');

      }
      dataObject["spRowNumber"] =  i ;
      // Push the object to the dataArray
    
      if(!isEmpty(dataObject.ProjectName) && !isEmpty(dataObject.StatusUpdate) && isNotUploaded(dataObject.StatusUpdate))
      dataArray.push(dataObject);
      

    }
  
   
    // Return the array of objects
    return dataArray;
  }

  function isNotUploaded(strInput){
    var  substring = 'uploaded:';
    var foundAt=strInput.indexOf(substring);
    if(foundAt==-1){
      return true;     
      }
  
    return false;
  }
  function excelSpreadSheetData(sheetName) {
  
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var range = sheet.getDataRange();
    //var range = sheet.getRange("A1:X1");
    var values = range.getValues();
   
    // Get column headers (field names) from the first row
    var headers = values[0];


   
  
    // Initialize an array to store objects
    var dataArray = [];
  
    // Iterate over each row starting from the second row (index 1)
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var dataObject = {};
      dataObject["spRowNumber"] =  i ;

        Logger.log('Entered Insidet Data Each Object ^^^^^^^^^^^^=='+i);
        Logger.log(dataObject);
        Logger.log('Entered Outside Data Endo of Object *********');

      // Iterate over each column
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        var value = row[j];
  
        // Use column header as key and cell value as value in the object
        dataObject[header] =  value ;
        

        //Logger.log('Entered Insidet Data Each Object');
       // Logger.log(dataObject);
        //Logger.log('Entered Outside Data Endo of Object');

      }
  
      // Push the object to the dataArray
      
      dataArray.push(dataObject);
      

    }
  
   
    // Return the array of objects
    return dataArray;
  }
  function addProjectsDataRangeNewFirst(){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("projects");
    sheet.getRange("A1:X1").setValues([[
          "CIDP",
          "FinYear",
          "Department",
          "Section",
          "Strategy",
          "SubProgramme",
          "ProjectType",
          "ProjectName",
          "Description",
          "IsFlagship",
          "Status",
          "MilestoneStatus",
          "StartDate",
          "EndDate",
          "Contractor",
          "TenderNum",
          "ContractSum",
          "Budget",
          "Funding",
          "SubCounty",
          "Ward",
          "Location",
          "CreatedOn",
          "StatusUpdate"
        ]]);
  
        dataRangeSource=sheet.getRange("cidp!A2:A55");
        targetplace = getTargetPlace(1,2);
        addColumnRule(dataRangeSource,targetplace,sheet);
        
        dataRangeSource=sheet.getRange("financial_year!A2:A55");
        targetplace = getTargetPlace(2,2);
        addColumnRule(dataRangeSource,targetplace,sheet);
        
        dataRangeSource=sheet.getRange("dept_directorates!C2:C55"); 
        targetplace = getTargetPlace(3,2);
        addColumnRule(dataRangeSource,targetplace,sheet);

        dataRangeSource=sheet.getRange("dept_directorates!A2:A55"); 
        targetplace = getTargetPlace(4,2);
        addColumnRule(dataRangeSource,targetplace,sheet);
        
        dataRangeSource=sheet.getRange("sector_programs!A2:A55");
        targetplace = getTargetPlace(5,2);
        addColumnRule(dataRangeSource,targetplace,sheet);
        
        dataRangeSource=sheet.getRange("sect_sub_programs!A2:A55");
        targetplace = getTargetPlace(6,2);
        addColumnRule(dataRangeSource,targetplace,sheet);
        
        dataRangeSource=sheet.getRange("project_types!A2:A55");
        targetplace = getTargetPlace(7,2);
        addColumnRule(dataRangeSource,targetplace,sheet); 
        
        dataRangeSource=sheet.getRange("progress_status!A2:A55");
        targetplace = getTargetPlace(11,2);
        addColumnRule(dataRangeSource,targetplace,sheet);
        
        dataRangeSource=sheet.getRange("BoQ_phase_milestones!A2:A55");
        targetplace = getTargetPlace(12,2);
        addColumnRule(dataRangeSource,targetplace,sheet); 
        
        dataRangeSource=sheet.getRange("contractors!A2:A55");
        targetplace = getTargetPlace(15,2);
        addColumnRule(dataRangeSource,targetplace,sheet);
        
        dataRangeSource=sheet.getRange("wards!A2:A55");
        targetplace = getTargetPlace(20,2);
        addColumnRule(dataRangeSource,targetplace,sheet);
        
        dataRangeSource=sheet.getRange("wards!A2:A55");
        targetplace = getTargetPlace(21,2);
        addColumnRule(dataRangeSource,targetplace,sheet);           
    
       
      
  }


 

  function removeDecimalPlaces (str, dataValue) {
    var strArray=["CIDPID",
                  "FinYearID",
                  "DepartmentID",
                  "SectionID",
                  "StrategyID",
                  "SubProgrammeID",
                  "TypeID",
                  "Status",
                  "ProjStatus",
                  "ContractorID",
                  "SubCountyID",
                  "WardID"];

               var isFound=   strArray.indexOf(str);
   

        if (isFound >-1){        
       // Logger.log('Entered Insidet Data Each Object dataValue ');
        //Logger.log(str+'====='+dataValue);
        //Logger.log('Entered Outside Data Endo of Object EDDDDDDDD');

          //if(dataValue)   return dataValue.replace(".0",'');
        
    }
    return dataValue;
}

function filterDropDownListEdit(e) {
  const sheet = e.source.getActiveSheet();
  const editedCell = e.range;
  
  // Check if the edited cell is in the range of Dropdown1
  Logger.log(' SOURCED ====='+sheet.getName()+'CLMN='+ editedCell.getRow());
  if (sheet.getName() == "projects" && editedCell.getColumn() == 3 && editedCell.getRow() >= 2 && editedCell.getRow() <= 1000) {
    const selectedValue = editedCell.getValue();
   // const dropdown2Cell = sheet.getRange(editedCell.getRow(), 2); // Assuming Dropdown2 is in the adjacent column
   const dropdown2Cell = sheet.getRange(editedCell.getRow(), 4); // Assuming Dropdown2 is in the adjacent column
    //const dropdown2Cell = sheet.getRange("E2:E1000"); // Assuming Dropdown2 is in the adjacent column

    // Clear existing data validation in Dropdown2
    dropdown2Cell.clearDataValidations();

    // Set new data validation in Dropdown2 based on the selected value in Dropdown1
    let newOptions;
    let directorates=excelSpreadSheetData('dept_directorates');
    var cellSelectedData= [];
    

    Logger.log('directorates BBBBBBBBBBBBBBBBBBBBBBB');
    Logger.log(directorates);
    newOptions=directorates.filter(function (entry) { return entry.DepartmentName === selectedValue; });
    Logger.log(newOptions);
       newOptions.forEach(function(item, index) {
      cellSelectedData.push([item.Name]);
      });
         Logger.log('Entered Insidet Data Each Object dataValue ');
        Logger.log('-------------------------------');
        Logger.log(cellSelectedData);
        Logger.log('Entered Outside Data Endo of Object EDDDDDDDD');


    // Define options for Dropdown2 based on the selected value in Dropdown1
    if (selectedValue == "Option1") {
      newOptions = ["OptionA", "OptionB", "OptionC"];
    } else if (selectedValue == "Option2") {
      newOptions = ["OptionX", "OptionY", "OptionZ"];
    }

    

    // Add more conditions as needed

    const rule = SpreadsheetApp.newDataValidation().requireValueInList(cellSelectedData).build();
    dropdown2Cell.setDataValidation(rule);
  }
}