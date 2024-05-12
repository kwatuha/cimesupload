var user=retrieveUserData();

var HEADERS = {
  "Authorization": "Basic " + Utilities.base64Encode(user.username+ ":" +user.password)
};


var SERVER = "http://134.122.4.64/cloud/ws/api/smartcounty/";
var attachmentUrl=SERVER+"attachment/types?IsVoided=0&_=1714240431638"; 
var usersUrl=SERVER+"system/users?IsVoided=0&_=1714384904093"; 
var wardsUrl=SERVER+"wards?IsVoided=0&_=1714384904095"; 
var departmentsUrl=SERVER+"county/department?IsVoided=0&_=1714384904096"; 
var subcountiesUrl=SERVER+"subcounties?IsVoided=0&_=1714392967300";
var projectTypesUrl=SERVER+"project/types?IsVoided=0&_=1714466326993";
var projectProgressUrl=SERVER+"project/progress/status?IsVoided=0&_=1714466326992";
var projectMilestoneUrl=SERVER+"project/status?IsVoided=0&_=1714466326994";
var cidpRegistryUrl=SERVER+"cidp/registry?IsVoided=0&_=1714488272905";
var financialYearUrl=SERVER+"financial/years?IsVoided=0&_=1714488272906";
var contractorsUrl=SERVER+"contractors?IsVoided=0&_=1714488272910";
var deptSectionsUrl=SERVER+"department/sections?IsVoided=0&_=1714488272913";
var sectorStrategiesUrl=SERVER+"sector/strategies?IsVoided=0&_=1714712317274";
var subProgramsUrl=SERVER+"strategy/subprogramme?IsVoided=0&_=1714712317276";
var projectsSearchUrl=SERVER+"projects/details?IsVoided=0";
var projectsPostUrl=SERVER+"projects";
 
function getMetaData() {
  
 var foundCidp=importCidp(); 

  if(foundCidp){
    
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
    addProjectObs(); 
     addContractors();
     addPrograms();
     addSubPrograms();
     importSystemProjects();  
     addCIMESMenu() ;
     addProjectsDataRangeNewFirst();
     addProjectPayments() ;
  }

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
var subcountyResults = fetchUrlData('subcounties');
// Clear existing data in columns A to C of the "main" sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("wards");
sheet.getRange("A:E").clear(); // Assuming you want to clear columns A to C, adjust as needed

// Write headers
sheet.getRange("A1:E1").setValues([["Name","SubCountyID", "WardID","SubCountyName","SubCounty"]]);

// Write data to columns
var rowData = [];
results.forEach(function(item, index) {

  var subcountyid=item.SubCountyID;
var subcountyName='Unknown';
 
var srsubcountys=subcountyResults.filter(function (entry) { return entry.SubCountyID === subcountyid; })
if(srsubcountys.length>0) subcountyName=srsubcountys[0].Name;

  rowData.push([item.Name,item.SubCountyID, item.WardID,subcountyName,item.SubCountyID]);
});
sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);

}


function importProjectTypes() {
var results = fetchUrlData('project_types');

// Clear existing data in columns A to C of the "main" sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("project_types");
sheet.getRange("A:E").clear(); // Assuming you want to clear columns A to C, adjust as needed

// Write headers
sheet.getRange("A1:E1").setValues([["TypeName","TypeID",  "Description","TypeID","TypeName"]]);

// Write data to columns
var rowData = [];
results.forEach(function(item, index) {
  rowData.push([item.TypeName,item.TypeID,  item.Description,item.TypeID,item.TypeName ]);
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
  var results=[];
    results = fetchUrlData('cidp');
    if(results){
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
    return results.length;
    }
    return 0;
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
    sheet.getRange("A:E").clear(); // Assuming you want to clear columns A to C, adjust as needed
    
    // Write headers
    sheet.getRange("A1:E1").setValues([["CompanyName","ContractorID",  "ContactName", "Phone","Company"]]);
    
    // Write data to columns
    var rowData = [];
    results.forEach(function(item, index) {
      rowData.push([item.CompanyName,item.ContractorID,  item.ContactName, item.Phone,item.CompanyName]);
    });
    sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);
    
    }
   
    function importSectorPrograms() {
    var results = fetchUrlData('sector_programs');
    var deptResults = fetchUrlData('departments');
    // Clear existing data in columns A to C of the "main" sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sector_programs");
    sheet.getRange("A:E").clear(); // Assuming you want to clear columns A to C, adjust as needed
    
    // Write headers
    sheet.getRange("A1:E1").setValues([["Programme","StrategyID", "CIDPID", "DepartmentID","Department"]]);
    
    // Write data to columns
    var rowData = [];
    results.forEach(function(item, index) {
      var deptid=item.DepartmentID;
      var deptName='Unknown';

      var srDepts=deptResults.filter(function (entry) { return entry.DepartmentID === deptid; })
      if(srDepts.length>0) deptName=srDepts[0].Name;
      rowData.push([item.Programme,item.StrategyID, item.CIDPID, item.DepartmentID,deptName]);
    });
    sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);
    
    }
    
  
    function importSectorSubPrograms() {
    var results = fetchUrlData('sect_sub_programs');
    var strategyResults = fetchUrlData('sector_programs');
    // Clear existing data in columns A to C of the "main" sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sect_sub_programs");
    sheet.getRange("A:E").clear(); // Assuming you want to clear columns A to C, adjust as needed
    
    // Write headers
    sheet.getRange("A1:E1").setValues([["SubProgramme","SubProgrammeID","StrategyID","KeyOutcome","Strategy"]]);
    
    // Write data to columns
    var rowData = [];
    results.forEach(function(item, index) {

      var strategyid=item.StrategyID;
        var strategyName='Unknown';
        var srstrategys=strategyResults.filter(function (entry) { return entry.StrategyID === strategyid; })
        if(srstrategys.length>0) strategyName=srstrategys[0].Programme;

      rowData.push([item.SubProgramme,item.SubProgrammeID, item.StrategyID,  item.KeyOutcome,strategyName]);
    });
    if(results.length>0)
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
        var srDepts=deptResults.filter(function (entry) { return entry.DepartmentID === deptid; })
        if(srDepts.length>0) deptName=srDepts[0].Name;
        
      rowData.push([item.Name,item.SectionID, deptName, item.DepartmentID ]);
    });
    sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);
    
    }


function importProjectMilestone() {
var results = fetchUrlData('BoQ_phase_milestones');
var projTypeResults = fetchUrlData('project_types');



// Clear existing data in columns A to C of the "main" sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BoQ_phase_milestones");
sheet.getRange("A:F").clear(); // Assuming you want to clear columns A to C, adjust as needed

// Write headers
sheet.getRange("A1:F1").setValues([["Status","StatusID", "TypeID",  "Percentage", "Description","ProjectType"]]);

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

  var pTypeId=item.TypeID;
var projTypeName='Unknown';
var srProjects=projTypeResults.filter(function (entry) { return entry.TypeID === pTypeId; });

if(srProjects.length>0) projTypeName=srProjects[0].TypeName;

  rowData.push([item.Status,item.StatusID, item.TypeID,  item.Percentage, item.Description,projTypeName]);
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

  


function onEdit(e) {
       //filterDropDownListEdit(e);
       //updateMetaData(e);
       // setDropDownDataRange(e);
       //filterMilestones(e)
       filterDropDownBy(e) 
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

function addContractors(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("contractor_reg");;
    sheet.getRange("A1:J1").setValues([[
        "CompanyName",
        "ContactName",
        "Phone",
        "Email",
        "Address",
        "PostalCode",
        "City",
        "County",    
        "Region",
        "StatusUpdate"
      ]]);
 
}

function addProjectObs(){
    //http://134.122.4.64/cloud/ws/api/smartcounty/project/observations
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("me_observations");
    //ObservationID 
    //sheet.getRange("A:P").clear();
      sheet.getRange("A1:P1").setValues([[
          "ProjectName",
          "ProjectStatus",
          "ProgressStatus",
          "OutputAchieved",
          "OutputPercentage",
          "Findings",
          "Challenges",
          "DelayReason",
          "Remarks",
          "Recommendations",
          "ObservationDate",
          "PostalCode",
          "City",
          "County",    
          "Region",
          "StatusUpdate"
        ]]);

        dataRangeSource=sheet.getRange("syspro!A2:A1000");
        targetplace = getTargetPlace(1,2);
        addColumnRule(dataRangeSource,targetplace,sheet); 

        dataRangeSource=sheet.getRange("progress_status!A2:A55");
        targetplace = getTargetPlace(3,2);
        addColumnRule(dataRangeSource,targetplace,sheet);
        
        /*
        
        dataRangeSource=sheet.getRange("BoQ_phase_milestones!A2:A55");
        targetplace = getTargetPlace(3,2);
        addColumnRule(dataRangeSource,targetplace,sheet); 

        dataRangeSource=sheet.getRange("BoQ_phase_milestones!A2:A55");
        targetplace = getTargetPlace(5,2);
        addColumnRule(dataRangeSource,targetplace,sheet); */
        
   
  }

  function filterMilestones(e) {
    var sheet = e.source.getActiveSheet();
    var editedCell = e.range;
    var editedRow = editedCell.getRow();
    var editedsheet =editedCell.getSheet().getSheetName(); 

    // Check if the edited cell is in the category column (e.g., column A)
    if (editedCell.getColumn() == 1 && editedsheet=='me_observations') {
      var category = editedCell.getValue();
     // var dataRange = sheet.getRange("A2:B" + sheet.getLastRow()); // Assuming data starts from row 2
      var dataRange = sheet.getRange("BoQ_phase_milestones!A2:E55");// Assuming data starts from row 2
      var data = dataRange.getValues();
      
      var projectYped="Q"+ editedRow;
      var searchkey=sheet.getRange(projectYped);
      var selProjectType=searchkey.getValue();
   
      var filteredItems = data.filter(function(row) {
        return row[2] == selProjectType;
      }).map(function(row) {
        return row[0];
      });
     
      
      // Set data validation for the items column with the filtered items
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(filteredItems)
        .build();
      
      sheet.getRange(2, 2, filteredItems.length, 1).setDataValidation(rule);
    }
  }

  
 
  
  /*UserType=Employee
  ObservationID=0
ProjectID=71
StatusID=13
TypeID
OutputAchieved=12
OutputMetric=N%2FA
Targets=FY%3A+N%2FA+%3D+(BL%3A+N%2FA%2C+QTR1%3A+N%2FA%2C+QTR2%3A+N%2FA%2C+QTR3%3A+N%2FA%2C+QTR4%3A+N%2FA)
CIDPKPI=KPI%3A+++%7C+Outcome%3A+Empowered+high+school+youth
CIDPTargets=Baseline%3A++1%2C+Y1%3A+10%2C+Y2%3A+20%2C+Y3%3A+60%2C+Y4%3A+10%2C+Y5%3A+10
OutputPercentage=33
ObservationDate=Mon+May+06+2024+21%3A14%3A19+GMT%2B0300+(East+Africa+Time)
Remarks=Test
Findings=Worked
Challenges=No+challenge
Recommendations=Frequent+review
ProgressStatus=Accepted
DelayReason=
TypeID=7
UserType=Employee
CreatedBy=1
UpdatedBy=1
VoidedBy=1*/

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
  
  
  nextRows=1000;
 return  cellRef+'2:'+cellRef+nextRows;

}

function getTargetPlaceWithStartRow(editedColumn,startAtRow){
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
   
    nextRows=1000;
   return  cellRef+startAtRow+':'+cellRef+nextRows;
  
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
          
                if(res){
                 updateStatusAfterUpload(rowNum, res.result.insertId);
                }   
           Logger.log(response.getContentText()); // Log the response from the server
     
      
    });
  }  
  function updateStatusAfterUpload(spRowNumber,insertId){
    //{"startIndex":0,"result":{"fieldCount":0,"affectedRows":1,"insertId":36,"serverStatus":2,"warningCount":0,"message":"","protocol41":true,"changedRows":0}}
    var user=retrieveUserData();
    var value='uploaded:'+user.username;
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
      .addItem('Refresh', 'updateMetaData')
      .addItem('Upload Project', 'showMessage')
      .addItem('Upload Contractors', 'saveContrators')
      .addItem('Upload M&E ', 'saveMEObservations')
      .addItem('Upload Payments', 'insertTimestamp')     
      .addItem('Upload Departmental Strategies', 'saveSectorStrategies')
      .addItem('Upload Sector Programs', 'saveSectorPrograms')
       .addItem('Upload Payements', 'saveProjectPayments')
      .addToUi();
  }
 
  // Function to be called when the button is clicked
function doSomething(username,password) {
  // Add your desired functionality here
  //var details=getUserDetails(username,password)
  
  storeUserData( username,password) 

  var foundCidp=importCidp(); 
  if(foundCidp){
    SpreadsheetApp.getActiveSpreadsheet().toast('Login Successful', 'Status', 3);
  
  }   else {
  
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Login Failed', 'Status', 3);
  
  }
 


    //setCookie(cookieName, cookieValue, expirationDays);
    //return setCookie(cookieName, cookieValue, expirationDays);
  //====
  
  //SpreadsheetApp.getActiveSpreadsheet().toast(details.length, 'Status', 3);
}



  function setLogin(){
    //handleButtonClick();
    //dodGet()
    retrieveUserData();
  }

  function setCookie(name, value, expirationDays) {
    var expirationDate = new Date();
    expirationDate.setDate(expirationDate.getDate() + expirationDays);
  
    var cookie = name + "=" + encodeURIComponent(value) + ";expires=" + expirationDate.toUTCString() + ";path=/";
    return ContentService.createTextOutput().appendHeader('Set-Cookie', cookie).setMimeType(ContentService.MimeType.TEXT);
  }

  function setLocalStorageVariable(key, value) {
    PropertiesService.getUserProperties().setProperty(key, value);
  }
  

  function storeUserData( username,password) {
    var username = username;
    var userEmail = password;
  
    setLocalStorageVariable("username", username);
    setLocalStorageVariable("password", password);
  
    Logger.log("User data stored in local storage variables.");
  }
  
  function getLocalStorageVariable(key) {
    return PropertiesService.getUserProperties().getProperty(key);
  }

  function dodGet() {
    // Store user data when the web app is accessed
   // storeUserData();
  
    // Retrieve and display user data
   
  
    return ContentService.createTextOutput("User data stored and retrieved successfully.");
  }
  function retrieveUserData() {
    var storedUsername = getLocalStorageVariable("username");
    var storedpassword = getLocalStorageVariable("password");
  
    if (storedUsername && storedpassword) {
      //Logger.log("Username: " + storedUsername);
     // Logger.log("storedpassword: " + storedpassword);
    } else {
     // Logger.log("User data not found.");
    }
    return {username:storedUsername, password:storedpassword}
  }
  
function updateMetaData(){
  getMetaData();
  SpreadsheetApp.getActiveSpreadsheet().toast('System Refreshed', 'Message', 3000);
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
  
  function saveContrators() {
    var newData= fetchContractors();
    postContractorData(newData);
    getMetaData();
     SpreadsheetApp.getActiveSpreadsheet().toast('Contractor List Uploaded', 'Message', 3000);
   }
 
  
 function fetchContractors() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('contractor_reg');
    var range = sheet.getDataRange();
    //var range = sheet.getRange("A1:X1");
    var values = range.getValues();
    var headers = values[0];
    var dataArray = [];
    // Iterate over each row starting from the second row (index 1)
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var dataObject = {};
      // Iterate over each column
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        var value = row[j];
        dataObject[header] =  value ;
      }
      dataObject["spRowNumber"] =  i ;
      if(!isEmpty(dataObject.CompanyName) && !isEmpty(dataObject.StatusUpdate) && isNotUploaded(dataObject.StatusUpdate))
      dataArray.push(dataObject);
    }
    return dataArray;
  }


  function postContractorData(data) {
    data.forEach(function(item, index) {
     item["ProductService"]=68120; 
      var rowNum=2;
      var cItem=item;
      if(cItem.spRowNumber) rowNum=cItem.spRowNumber+1;
     var url = SERVER+"contractors";
       var options = {
             headers: HEADERS,
             method: "post",
             contentType: "application/json",
             payload: JSON.stringify(item),
             muteHttpExceptions: true
           };
           var response = UrlFetchApp.fetch(url, options);
           var res=JSON.parse(response.getContentText());
         
               if(res){
                var user=retrieveUserData();
                var value='uploaded:'+user.username;
                setValueToCell('contractor_reg',10,rowNum,value);
               }   
          Logger.log(response.getContentText()); // Log the response from the server  
   });
 }  

  function importSystemProjects() {
    var results = fetchUrlData('existing_projects');
    // syspro ProjectID TypeID TenderNum ProjectName DepartmentID FinYearID CIDPID Status ProjStatus WardID ContractorID SubCountyID WardID
    // Clear existing data in columns A to C of the "main" sheet
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("syspro");
    sheet.getRange("A:L").clear(); // Assuming you want to clear columns A to C, adjust as needed
    
    // Write headersProjStatus
    sheet.getRange("A1:L1").setValues([[
"ProjectName","ProjectID", "TypeID", "DepartmentID","FinYearID", "ContractorID","Status","WardID", "SubCountyID","ProjStatus","TotAmountPaid","RecPercPaid"

    ]]);
    
    // Write data to columns ProjectName Description ProjectID
    var rowData = [];
    results.forEach(function(item, index) {
      rowData.push([
item.ProjectName,item.ProjectID,item.TypeID,item.DepartmentID,item.FinYearID,item.ContractorID,item.Status,item.WardID,item.SubCountyID, item.ProjStatus,item.TotAmountPaid, item.RecPercPaid

      ]);
    });
     if(results.length>0)
    sheet.getRange(2, 1, rowData.length, rowData[0].length).setValues(rowData);
    
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
    sheet.getRange("A:T").clear();
    sheet.getRange("A:T").clearDataValidations();
    sheet.getRange("A5:T5").setValues([[
          "CIDP",
          "FinYear",
          "DepartmentalStrategies",
          "Programme",
          "ProjectName",
          "Description",
          "IsFlagship",
          "ProjectType",
          "MilestoneStatus",
          "Status",         
          "TenderNum",
          "ContractSum",
          "Budget",
          "Funding",
          "Contractor",
          "SubCounty",
          "Ward",
          "Location",
          "CreatedOn",
          "StatusUpdate"
        ]]);
      
        
      targetplace = 'B2';
      dataRangeSource=sheet.getRange("dept_directorates!C2:C1000"); 
      addColumnRule(dataRangeSource,targetplace,sheet);

      targetplace = 'B3';
      dataRangeSource=sheet.getRange("dept_directorates!A2:A1000"); 
      addColumnRule(dataRangeSource,targetplace,sheet);

      var startAtRow=6;

      dataRangeSource=sheet.getRange("cidp!A2:A55");
      targetplace = getTargetPlaceWithStartRow(1,startAtRow);
      addColumnRule(dataRangeSource,targetplace,sheet);
      
     /* dataRangeSource=sheet.getRange("financial_year!A2:A55");
      targetplace = getTargetPlaceWithStartRow(2,startAtRow);
      addColumnRule(dataRangeSource,targetplace,sheet);

     dataRangeSource=sheet.getRange("sector_programs!A2:A1000");
      targetplace = getTargetPlaceWithStartRow(3,startAtRow);
      addColumnRule(dataRangeSource,targetplace,sheet);

     dataRangeSource=sheet.getRange("sect_sub_programs!A2:A1000");
      targetplace = getTargetPlaceWithStartRow(4,startAtRow);
      addColumnRule(dataRangeSource,targetplace,sheet);*/
/**
 *         "CIDP",1
          "FinYear",2
          "DepartmentalStrategies",3
          "Programme",4
          "ProjectName",5
           "Description",6
          "IsFlagship",7
          */
      dataRangeSource=sheet.getRange("project_types!A2:A55");
      targetplace = getTargetPlaceWithStartRow(8,startAtRow);
      addColumnRule(dataRangeSource,targetplace,sheet); 

      /*dataRangeSource=sheet.getRange("BoQ_phase_milestones!A2:A55");
      targetplace = getTargetPlaceWithStartRow(9,startAtRow);
      addColumnRule(dataRangeSource,targetplace,sheet); */
      
      dataRangeSource=sheet.getRange("progress_status!A2:A55");
      targetplace = getTargetPlaceWithStartRow(10,startAtRow);
      addColumnRule(dataRangeSource,targetplace,sheet);
       /*
          "ProjectType",8
          "MilestoneStatus",9
          "Status", 10       
          "TenderNum",11
          "ContractSum",12
          "Budget",13
          */    
      dataRangeSource=sheet.getRange("contractors!A2:A1000");
      targetplace = getTargetPlaceWithStartRow(15,startAtRow);
      addColumnRule(dataRangeSource,targetplace,sheet);
      
      dataRangeSource=sheet.getRange("wards!D2:D55");
      targetplace = getTargetPlaceWithStartRow(16,startAtRow);
      addColumnRule(dataRangeSource,targetplace,sheet);
      
      /*dataRangeSource=sheet.getRange("wards!A2:A55");
      targetplace = getTargetPlaceWithStartRow(17,startAtRow);
      addColumnRule(dataRangeSource,targetplace,sheet);           
    
       
          "Funding",14
          "Contractor",15
          "SubCounty",16
          "Ward",17
          "Location",18
          "CreatedOn",19
          "StatusUpdate"20
          */
      
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
    


    newOptions=directorates.filter(function (entry) { return entry.DepartmentName === selectedValue; });
    Logger.log(newOptions);
       newOptions.forEach(function(item, index) {
      cellSelectedData.push([item.Name]);
      });
     


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


function saveMEObservations() {
  var newData= fetchMEObservations();
  postMEObservationsData(newData);
  getMetaData();
  
   SpreadsheetApp.getActiveSpreadsheet().toast('M&E Observation List Uploaded', 'Message', 3000);
 }


function fetchMEObservations() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('me_observations');
  var range = sheet.getDataRange();
  //var range = sheet.getRange("A1:X1");
  var values = range.getValues();
  var headers = values[0];
  var dataArray = [];
  // Iterate over each row starting from the second row (index 1)
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var dataObject = {};
    // Iterate over each column
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      var value = row[j];
      dataObject[header] =  value ;
    }
    dataObject["spRowNumber"] =  i ;
    if(!isEmpty(dataObject.ProjectID) && !isEmpty(dataObject.StatusUpdate) && isNotUploaded(dataObject.StatusUpdate))
    dataArray.push(dataObject);
  }
  return dataArray;
}


function postMEObservationsData(data) {
  data.forEach(function(item, index) {
   item["ProductService"]=68120; 
    var rowNum=2;
    var cItem=item;
    if(cItem.spRowNumber) rowNum=cItem.spRowNumber+1;
   var url = SERVER+"project/observations";
     var options = {
           headers: HEADERS,
           method: "post",
           contentType: "application/json",
           payload: JSON.stringify(item),
           muteHttpExceptions: true
         };
         var response = UrlFetchApp.fetch(url, options);
         var res=JSON.parse(response.getContentText());
         Logger.log(response.getContentText());
         Logger.log(res);
         Logger.log("ldlldlldlldld");
       
             if(res){
              var user=retrieveUserData();
              var value='uploaded:'+user.username;
              setValueToCell('me_observations',16,rowNum,value);
             }   
        Logger.log(response.getContentText()); // Log the response from the server  
 });
}  


function addPrograms(){
  //http://134.122.4.64/cloud/ws/api/smartcounty/project/observations
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Departmental_strategies");

    sheet.getRange("A7:G7").setValues([[
        "Programme",
        "NeedsPriorities",
        "Objectives",
        "Outcomes",
        "Strategies",
        "Remarks",
        "StatusUpdate"
      ]]);


      var targetplace='';
      dataRangeSource=sheet.getRange("cidp!A2:A55");
      targetplace = 'B2';
      addColumnRule(dataRangeSource,targetplace,sheet);
    
      targetplace = 'B3';
      dataRangeSource=sheet.getRange("dept_directorates!C2:C55"); 
      addColumnRule(dataRangeSource,targetplace,sheet);

      targetplace = 'B4';
      dataRangeSource=sheet.getRange("dept_directorates!A2:A55"); 
      addColumnRule(dataRangeSource,targetplace,sheet);
    
      
 
}

function addSubPrograms(){
  //http://134.122.4.64/cloud/ws/api/smartcounty/project/observations
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("sector_sub_programs");

    sheet.getRange("A5:L5").setValues([[
        "SubProgramme",
        "KeyOutcome",
        "KPI",
        "Baseline",
        "Yr1Targets",
        "Yr2Targets",
        "Yr3Targets",
        "Yr4Targets",
        "Yr5Budget",
        "TotalBudget",
        "Remarks",
        "StatusUpdate"
      ]]);

      var targetplace='B3';
      dataRangeSource=sheet.getRange("sector_programs!A2:A55");
      addColumnRule(dataRangeSource,targetplace,sheet);
    
      
 
}

function saveSectorStrategies() {
  var newData= fetchSectorStrategies();
  postSectorStrategiesData(newData);
  
  getMetaData();
 
   SpreadsheetApp.getActiveSpreadsheet().toast('Sector Strategies List Uploaded', 'Message', 3000);
 }


function fetchSectorStrategies() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Departmental_strategies');
  var range = sheet.getDataRange();
  //var range = sheet.getRange("A1:X1");
  var CIDPID=sheet.getRange('I2').getValue();
  var DepartmentID=sheet.getRange('I3').getValue();
  var SectionID=sheet.getRange('I4').getValue();

  var values = range.getValues();
  var headers = values[6];
  var dataArray = [];

  // Iterate over each row starting from the second row (index 1)
  for (var i = 7; i < values.length; i++) {
    var row = values[i];
    var dataObject = {};
    // Iterate over each column
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      var value = row[j];
      dataObject[header] =  value ;
    }
    dataObject["spRowNumber"] =  i ;
    dataObject["CIDPID"] =  CIDPID ;
    dataObject["DepartmentID"] =  DepartmentID ;
    dataObject["SectionID"] =  SectionID ;
    if(!isEmpty(dataObject.Programme) && !isEmpty(dataObject.StatusUpdate) && isNotUploaded(dataObject.StatusUpdate))
    dataArray.push(dataObject);
  }
  return dataArray;
}


function postSectorStrategiesData(data) {
  data.forEach(function(item, index) {
   item["ProductService"]=68120; 
    var rowNum=2;
    var cItem=item;
    if(cItem.spRowNumber) rowNum=cItem.spRowNumber+1;
   var url = SERVER+"sector/strategies";
     var options = {
           headers: HEADERS,
           method: "post",
           contentType: "application/json",
           payload: JSON.stringify(item),
           muteHttpExceptions: true
         };
         var response = UrlFetchApp.fetch(url, options);
         var res=JSON.parse(response.getContentText());
         Logger.log(response.getContentText());
         Logger.log(res);
         Logger.log("ldlldlldlldld");
       
             if(res){
              var user=retrieveUserData();
              var value='uploaded:'+user.username;
              setValueToCell('Departmental_strategies',7,rowNum,value);
             }   
        Logger.log(response.getContentText()); // Log the response from the server  
 });
} 

function saveSectorPrograms() {
  var newData= fetchSectorPrograms();
  postSectorProgramsData(newData);
  getMetaData();
   SpreadsheetApp.getActiveSpreadsheet().toast('Sector Programs List Uploaded', 'Message', 3000);
 }


function fetchSectorPrograms() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sector_sub_programs');
  var range = sheet.getDataRange();
  //var range = sheet.getRange("A1:X1");
  var StrategyID=sheet.getRange('N3').getValue();


  var values = range.getValues();
  var headers = values[4];
  var dataArray = [];

  // Iterate over each row starting from the second row (index 1)
  for (var i = 5; i < values.length; i++) {
    var row = values[i];
    var dataObject = {};
    // Iterate over each column
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      var value = row[j];
      dataObject[header] =  value ;
    }
    dataObject["spRowNumber"] =  i ;
    dataObject["StrategyID"] =  StrategyID ;
 
    if(!isEmpty(dataObject.SubProgramme) && !isEmpty(dataObject.StatusUpdate) && isNotUploaded(dataObject.StatusUpdate))
    dataArray.push(dataObject);
  }
  return dataArray;
}


function postSectorProgramsData(data) {
  data.forEach(function(item, index) {
   item["ProductService"]=68120; 
    var rowNum=2;
    var cItem=item;
    if(cItem.spRowNumber) rowNum=cItem.spRowNumber+1;
   var url = SERVER+"strategy/subprogramme";
     var options = {
           headers: HEADERS,
           method: "post",
           contentType: "application/json",
           payload: JSON.stringify(item),
           muteHttpExceptions: true
         };
         var response = UrlFetchApp.fetch(url, options);
         var res=JSON.parse(response.getContentText());
         Logger.log(response.getContentText());
         Logger.log(res);
         Logger.log("ldlldlldlldld");
       
             if(res){
              var user=retrieveUserData();
              var value='uploaded:'+user.username;
              setValueToCell('sector_sub_programs',12,rowNum,value);
             }   
        Logger.log(response.getContentText()); // Log the response from the server  
 });
}  

function saveProjectPayments() {
  var newData= fetchPayments();
  postPaymentData(newData);
  getMetaData();
   SpreadsheetApp.getActiveSpreadsheet().toast('Payment List Uploaded', 'Message', 3000);
 }


function fetchPayments() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('payments');
  var range = sheet.getDataRange();
  //var range = sheet.getRange("A1:X1");
  var values = range.getValues();
  var headers = values[0];
  var dataArray = [];
  // Iterate over each row starting from the second row (index 1)
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var dataObject = {};
    // Iterate over each column
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      var value = row[j];
      dataObject[header] =  value ;
    }
    dataObject["spRowNumber"] =  i ;
    
    
    if(!isEmpty(dataObject.ProjectID) && !isEmpty(dataObject.StatusUpdate) && isNotUploaded(dataObject.StatusUpdate))
    dataArray.push(dataObject);
  }
  return dataArray;
}


function postPaymentData(data) {
  data.forEach(function(item, index) {
   item["ProductService"]=68120; 
   item["Staging"] =  "Payment";
   item["CertificateID"] =  "17";
    var rowNum=2;
    var cItem=item;
    if(cItem.spRowNumber) rowNum=cItem.spRowNumber+1;
   var url = SERVER+"project/payments";
     var options = {
           headers: HEADERS,
           method: "post",
           contentType: "application/json",
           payload: JSON.stringify(item),
           muteHttpExceptions: true
         };
         var response = UrlFetchApp.fetch(url, options);
         var res=JSON.parse(response.getContentText());
       
             if(res){
             // var value='uploaded:'+ res.result.insertId;

              var user=retrieveUserData();
              var value='uploaded:'+user.username;
              setValueToCell('payments',10,rowNum,value);
             }   
        Logger.log(response.getContentText()); // Log the response from the server  
 });
} 

function addProjectPayments(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("payments");
    sheet.getRange("A1:N1").setValues([[
        "Project",
        "ProjectStatus",
        "ProgressStatus",
        "Contractor",
        "PaymentType",
        "AmountPaid",
        "PercPaid",
        "IFMISID",
        "PaymentNumber",
        "TransactionID",        
        "TransactionDetails",
        "PaymentRemarks",
        "PaymentDate",
        "StatusUpdate"
      ]]);

}



function getEditedValue(editedColumn,editedColumn){
  editedColumn
}

function getCellRefByColumn(editedColumn){

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
   case 	22: cellRef="V";
    break;
   case 	23: cellRef="W";
    break;
    case 	24: cellRef="X";
    break;
    case 	25: cellRef="Y";
    break;
    case 	26: cellRef="Z";
    break;

  }
  return cellRef;
}
function getValueByColRow(sheet,editedColumn,editedRow){
 
  var cellRef= getCellRefByColumn(editedColumn);
  var ref= cellRef+editedRow;
  console.log(editedColumn+'====---------refrefref===='+ref)
  var editedCell= sheet.getRange(ref);
  return  editedCell.getValue();
  
  }


  function filterTargetColumn(sheet,strKey, dataSourceRange,filterCln,rtnValue,dropColumn,dropRow) {
       // Check if the edited cell is in the category column (e.g., column A)
     var dataRange = sheet.getRange(dataSourceRange);// Assuming data starts from row 2
      var data = dataRange.getValues();
 
      var cellRef= getCellRefByColumn(dropColumn);
      var ref= cellRef+dropRow
      var dropTargetCell= sheet.getRange(ref);
    
      var filteredItems = data.filter(function(row) {
        console.log(row[filterCln]+'UUUUUUUUUUUUUU^^^^^^^^^^^^^^^^^^^^^^^^^^^===='+strKey+'filterCln='+filterCln)
        console.log(row)
        return row[filterCln] == strKey;
      }).map(function(row) {
        return row[rtnValue];
      });

     Logger.log('XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'+strKey)
     Logger.log(filteredItems)

      // Set data validation for the items column with the filtered items
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(filteredItems)
        .build();

        dropTargetCell.clearDataValidations();
        dropTargetCell.setDataValidation(rule);
        
    
  }

  
  function filterTargetCIDPAdp(sheet,strKey, dataSourceRange,dropColumn,dropRow) {
    // Check if the edited cell is in the category column (e.g., column A)
  var dataRange = sheet.getRange(dataSourceRange);// Assuming data starts from row 2
   var data = dataRange.getValues();

   var cellRef= getCellRefByColumn(dropColumn);
   var ref= cellRef+dropRow
   var dropTargetCell= sheet.getRange(ref);

    
   var filteredItems = compareFYCIDP(strKey,data)
   Logger.log(data)
   // Set data validation for the items column with the filtered items
   var rule = SpreadsheetApp.newDataValidation()
     .requireValueInList(filteredItems)
     .build();

     dropTargetCell.clearDataValidations();

     dropTargetCell.setDataValidation(rule);
     
 
}

function compareFYCIDP(cidp,yrs) {
  console.log("CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC=="+cidp)
  console.log(yrs)
  console.log("CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC")
  //cidp=cidp.replace('CIDP:','');
  var  cidpStart=  cidp.substring(5,9);
  var  cidpEnd=    cidp.substring(10,15);
  console.log(cidpStart+'==xxxxxxCIDPxxxxxxx'+cidp+'=='+cidpEnd);
  var c=cidp.split('-');
  var foundYrs=[];
  yrs.forEach( (row) =>{

    if(row){
      console.log('xxxxxxxxxxxxx'+row)
      var  yearStart= row[0].replace(' ','');
      //FY2023-2024
      yearStart=yearStart.substring(2, 6);
      var  yearEnd= row[0].replace(' ','');
      yearEnd=yearEnd.substring(7,11);
      console.log(yearStart+'==xxxxxxxxxxxxx'+row+'=='+yearEnd);
      var isValid=comparePeriod(cidpStart,cidpEnd,yearStart,yearEnd);
         if(isValid==true) foundYrs.push(row);

    }
   
  });
  console.log('xxxxxxxxxxxxx----------')
  console.log(foundYrs)
  console.log('xxxxxxxxxxxxx----------')
  return [... new Set(foundYrs.sort())];
}
//2013-2017  2017-2018
function comparePeriod(cidpStart,cidpEnd,yearStart,YearEnd){
  if(yearStart*1>=cidpStart*1 && YearEnd*1<=cidpEnd*1){
          console.log(cidpStart,cidpEnd,yearStart,YearEnd)
 return true;
}
return false;
}
  
  function filterDropDownBy(e) {
    var sheet = e.source.getActiveSheet();
    var editedCell = e.range;
    var editedRow = editedCell.getRow();
    var editedColumn = editedCell.getColumn() ;
    var editedsheet =editedCell.getSheet().getSheetName(); 
  
    var editedValue = getValueByColRow(sheet,editedColumn,editedRow);
  
    // CIDP

   // addDataValidationWithFixedValues(sheet,dropColumn,rowNum)

    if(editedsheet=="projects" && editedColumn==1){
    Logger.log("CIDP:   "+editedValue)
    var DrcellRef= getCellRefByColumn(2);
    var Drref= DrcellRef+editedRow;
    var dropTargetCell= sheet.getRange(Drref);
    setCellState(dropTargetCell, 'Pending')

   var findProjectTypeKey=getValueByColRow(sheet,1,editedRow);
   // 9 dropdwon for milestone
   // filterTargetColumn(sheet,findProjectTypeKey, 'sect_sub_programs!A2:E1000',4,0,4,editedRow);
    filterTargetCIDPAdp(sheet,editedValue, 'financial_year!A2:A55',2,editedRow)
    Utilities.sleep(30);
    setCellState(dropTargetCell, 'Done')  
    }     
   
     // FinYear

     if(editedsheet=="projects" && editedColumn==2 && editedRow==2){
      var findProjectTypeKey=getValueByColRow(sheet,2,editedRow);
      // filterTargetByFixedCell(sheet,strKey, dataSourceRange,filterCln,rtnValue,dropColumn,startRow) 


       
    
      var DrcellRef= getCellRefByColumn(2);
     var Drref= DrcellRef+3;
     var dropTargetCell= sheet.getRange(Drref);
     setCellState(dropTargetCell, 'Pending')


    // 9 dropdwon for milestone
     filterTargetByFixedCell(sheet,findProjectTypeKey, 'sector_programs!A2:E1000',4,0,3,6,1000)
     filterTargetByFixedCell(sheet,findProjectTypeKey, 'dept_directorates!A2:C63',2,0,2,3,3)
     //filterTargetColumn(sheet,strKey, dataSourceRange,filterCln,rtnValue,dropColumn,dropRow)
     Utilities.sleep(30);
     setCellState(dropTargetCell, 'Done')

     }

    

     if(editedsheet=="projects" && editedColumn==3){
     Logger.log("Strategies:   "+editedValue)
  
     var DrcellRef= getCellRefByColumn(4);
     var Drref= DrcellRef+editedRow;
     var dropTargetCell= sheet.getRange(Drref);
     setCellState(dropTargetCell, 'Pending')

    var findProjectTypeKey=getValueByColRow(sheet,3,editedRow);
    // 9 dropdwon for milestone
     filterTargetColumn(sheet,findProjectTypeKey, 'sect_sub_programs!A2:E1000',4,0,4,editedRow);
     //filterTargetColumn(sheet,strKey, dataSourceRange,filterCln,rtnValue,dropColumn,dropRow)
     Utilities.sleep(30);
     setCellState(dropTargetCell, 'Done')


     }



     
       /*
  
  CIDPID 24 	FinYearID 25
  StrategyID 26	SubProgrammeID 27	 4
  TypeID 28	    ProjStatus 29 9
  SubCountyID 31	WardID 32  17
  StrategyID 26 AA2
  filterTargetColumn(sheet,strKey, dataSourceRange,filterCln,rtnValue,dropColumn,dropRow) 
  */
        // ProjectType
        
    if(editedsheet=="projects" && editedColumn==5){
          addDataValidationWithFixedValues(sheet,7,editedRow);
    }
        
    if(editedsheet=="projects" && editedColumn==8){
      var DrcellRef= getCellRefByColumn(9);
      var Drref= DrcellRef+editedRow;
      var dropTargetCell= sheet.getRange(Drref);
      setCellState(dropTargetCell, 'Pending')

     var findProjectTypeKey=getValueByColRow(sheet,8,editedRow);
     // 9 dropdwon for milestone
      filterTargetColumn(sheet,findProjectTypeKey, 'BoQ_phase_milestones!A2:F1000',5,0,9,editedRow);
      //filterTargetColumn(sheet,strKey, dataSourceRange,filterCln,rtnValue,dropColumn,dropRow)
      Utilities.sleep(30);
      setCellState(dropTargetCell, 'Done')
    }
  
  
    if(editedsheet=="projects" && editedColumn==16){
      Logger.log('Subcounties'+   editedValue)
       var DrcellRef= getCellRefByColumn(17);
      var Drref= DrcellRef+editedRow;
      var dropTargetCell= sheet.getRange(Drref);
      setCellState(dropTargetCell, 'Pending'); 
     var findProjectTypeKey=getValueByColRow(sheet,16,editedRow);
      filterTargetColumn(sheet,findProjectTypeKey, 'wards!A2:E1000',3,0,17,editedRow);
      Utilities.sleep(30);
      setCellState(dropTargetCell, 'Done')
     }




    
  
  
  }
  

  function filterTargetByFixedCell(sheet,strKey, dataSourceRange,filterCln,rtnValue,dropColumn,startRow,rowsToAddRule) {
    // Check if the edited cell is in the category column (e.g., column A)
  var dataRange = sheet.getRange(dataSourceRange);// Assuming data starts from row 2
   var data = dataRange.getValues();
   
   Logger.log('dataSourceRange0000000000000000000000000000000000000===')
   Logger.log(data)
   Logger.log(dataSourceRange)
   var filteredItems = data.filter(function(row) {

   
    Logger.log(rtnValue+'=====PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP=='+ row[filterCln])
Logger.log(row)
Logger.log('PPPPPPP^^^^^^^^^^^^^^^^^^^^^^^^^^^^PPPPPPPPPP======'+strKey)
     return row[filterCln] == strKey;
   }).map(function(row) {
     return row[rtnValue];
   });

   


   // Set data validation for the items column with the filtered items
   var rule = SpreadsheetApp.newDataValidation()
     .requireValueInList(filteredItems)
     .build();

     var cellRef= getCellRefByColumn(dropColumn);
     var ref= cellRef+startRow+':'+cellRef+rowsToAddRule;
     var dropTargetCell= sheet.getRange(ref);

     
     dropTargetCell.clearDataValidations();

     dropTargetCell.setDataValidation(rule);
     
}

function setCellState(targetCell, type){
  Logger.log(" targetCell.setValue('');"+targetCell.getValue())
  if(type=='Pending'){
  targetCell.setValue('');
  targetCell.setBackground('#d8d8d8');
   } else if(type=='Done'){
    targetCell.setBackground('#fff')
   }
}


function addDataValidationWithFixedValues(sheet,dropColumn,rowNum) {
  var cellRef= getCellRefByColumn(dropColumn);
  cellRef=cellRef+rowNum;
  // Define the fixed set of values for data validation
  var allowedValues = ["True", "False"];
  
  // Create a data validation rule based on the fixed set of values
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(allowedValues)
    .setAllowInvalid(false) // Prevent invalid entries
    .build();
  
  // Apply the data validation rule to cell A1
  var cellWithValidation = sheet.getRange(cellRef);
  cellWithValidation.setDataValidation(rule);
}
