/**
 * Google Apps Script for Workshop Management
 * 1. Create Folders in Drive
 * 2. Log Results to Spreadsheet
 */

const SPREADSHEET_ID = "1f1OE-Y0KaTmvFZ38P2ZJOp6wlJRAtiy4mnrYZPJKpP0";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    if (action === "initFolders") {
      return ContentService.createTextOutput(JSON.stringify(handleInitFolders(data.groups)))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === "submitResult") {
      return ContentService.createTextOutput(JSON.stringify(handleSubmitResult(data)))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === "getResults") {
      return ContentService.createTextOutput(JSON.stringify(handleGetResults()))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === "getResults") {
      const result = handleGetResults();
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid action" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Creates folders and sets permissions
 */
function handleInitFolders(groups) {
  const folderUrls = {};
  
  // 指定父資料夾 ID
  const parentFolder = DriveApp.getFolderById("1SCbiEBOdZyUs96Cnv1dCPfqxxzIFF0Vd"); 

  groups.forEach(g => {
    const folderName = `第${g.groupNum}組 - ${g.taskName}`;
    const folder = parentFolder.createFolder(folderName);
    
    // Set permission: Anyone with link can edit
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
    
    folderUrls[g.groupNum] = folder.getUrl();
  });
  
  return { status: "success", urls: folderUrls };
}

/**
 * Appends data to spreadsheet
 */
function handleSubmitResult(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheets()[0]; // Get the first sheet
  
  sheet.appendRow([
    `第${data.group}組`,
    data.taskName,
    data.url,
    new Date() // Timestamp
  ]);
  
  return { status: "success" };
}

/**
 * Fetches all results logged in the spreadsheet
 */
function handleGetResults() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets()[0];
    const values = sheet.getDataRange().getValues();
    const results = [];
    
    if (values.length > 0) {
      let startIdx = 0;
      // Skip header row if it exists
      const firstRow = values[0];
      if (
        firstRow[0] === "組別" || 
        firstRow[0] === "Group" || 
        firstRow[0].toString().includes("組別") ||
        firstRow[0].toString().includes("成員")
      ) {
        startIdx = 1;
      }
      
      for (let i = startIdx; i < values.length; i++) {
        const row = values[i];
        if (row[0] && row[2]) { // Ensure Group and URL columns are not empty
          results.push({
            group: row[0].toString(),
            taskName: row[1] ? row[1].toString() : "",
            url: row[2].toString(),
            timestamp: row[3] ? row[3].toString() : ""
          });
        }
      }
    }
    
    return { status: "success", data: results };
  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}
