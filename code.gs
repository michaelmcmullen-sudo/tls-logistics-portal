const DAILY_SHEET_ID = "14Cl2hzwBq6aH6B0uf-kuq4sGGK-qFVGa_hGXUZmBzPo"; 
const MASTER_SHEET_ID = "1jcI3nRQPAt6O3HAm3k8hBSCsq264_Y1xha_pZehW0V0"; 
const AM3_EXPORT_GID = 2117779177; 

/**
 * 1. CORE TRAFFIC ROUTER (Processes incoming connections from Teams/GitHub HTML)
 */
function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    let responseData = {};

    if (action === "processShipping") {
      processShipping(requestData.chassis, requestData.week, requestData.date);
      responseData = { message: "Outbound logistics update successfully appended." };
      
    } else if (action === "getFilteredCarpetExport") {
      const base64Output = getFilteredCarpetExport(requestData.model, requestData.weekNum);
      responseData = { base64: base64Output };
      
    } else if (action === "syncDemandData") {
      syncDemandData();
      responseData = { message: "Global demand synchronization execution cycle finalized." };
      
    } else if (action === "getTabData") {
      responseData = getTabData(requestData.tabName);
      
    } else if (action === "runValidationScan") {
      responseData = runValidationScan();
    }

    return ContentService.createTextOutput(JSON.stringify(responseData))
                         .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('TLS Logistics Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getTabData(sheetName) {
  let ss;
  try {
    ss = SpreadsheetApp.openById(DAILY_SHEET_ID);
  } catch(e) {
    ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  }
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    try { ss = SpreadsheetApp.openById(MASTER_SHEET_ID); sheet = ss.getSheetByName(sheetName); } catch(err) {}
  }
  if (!sheet) return [];
  return sheet.getDataRange().getValues();
}

/** * 2. CARPET WORKBOOK FILTER COMPILER
 */
function getFilteredCarpetExport(model, weekNum) {
  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    let sheetsToProcess = (model === "AM5") ? ["AP TLS AM5", "AP TLS AM7"] : ["AP TLS " + model];
    
    let combinedFilteredData = [];
    let labelChassis = [["Chassis Number"]];
    let headerWidth = 0;

    sheetsToProcess.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;

      const data = sheet.getDataRange().getValues();
      if (data.length === 0) return;
      
      const headers = data[0];
      headerWidth = headers.length;
      
      const colIndexO = 14; 
      const colIndexA = 0;  
      
      if (combinedFilteredData.length === 0) {
        combinedFilteredData.push(headers); 
      }

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const valO = row[colIndexO] ? row[colIndexO].toString().trim() : "";
        
        if (valO === weekNum) {
          combinedFilteredData.push(row);
          if (row[colIndexA]) {
            labelChassis.push([row[colIndexA].toString().trim()]);
          }
        }
      }
    });

    if (combinedFilteredData.length <= 1) {
      throw new Error("No data matched criteria.");
    }

    const tempTemp = SpreadsheetApp.create("TEMP_EXPORT_" + Utilities.getUuid());
    const mainExportSheet = tempTemp.getActiveSheet().setName("Filtered Dataset");
    mainExportSheet.getRange(1, 1, combinedFilteredData.length, headerWidth).setValues(combinedFilteredData);
    
    const labelSheet = tempTemp.insertSheet("Label Export Run");
    labelSheet.getRange(1, 1, labelChassis.length, 1).setValues(labelChassis);

    DriveApp.getFileById(tempTemp.getId()).setFlattenedFolders(true);
    SpreadsheetApp.flush();
    
    const url = "https://docs.google.com/spreadsheets/d/" + tempTemp.getId() + "/export?format=xlsx";
    const res = UrlFetchApp.fetch(url, {
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    
    DriveApp.getFileById(tempTemp.getId()).setTrashed(true);
    return Utilities.base64Encode(res.getBlob().getBytes());

  } catch (err) {
    throw new Error("Export runtime execution failed: " + err.toString());
  }
}

/** * 3. PROCESS OUTBOUND SHIPPING ENTRIES
 */
function processShipping(rawChassis, weekStr, dateStr) {
  const ss = SpreadsheetApp.openById(DAILY_SHEET_ID);
  const logSheet = ss.getSheetByName("AP TLS Log");
  if (!logSheet) throw new Error("Log worksheet targeted resource array broken.");

  const chassisList = rawChassis.split("\n")
                                .map(r => r.trim())
                                .filter(r => r.length > 0);
  if (chassisList.length === 0) return;

  const timestamp = new Date();
  const parsedDate = new Date(dateStr);
  const formattedDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

  const rowsToAppend = chassisList.map(chassis => [
    chassis,
    weekStr,
    formattedDate,
    timestamp
  ]);

  const lastRow = logSheet.getLastRow();
  logSheet.getRange(lastRow + 1, 1, rowsToAppend.length, 4).setValues(rowsToAppend);
}

/** * 4. OVERWRITE GLOBAL ALLOCATION DEMAND SYNC 
 */
function syncDemandData() {
  const sourceSs = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const targetSs = SpreadsheetApp.openById(DAILY_SHEET_ID);
  
  const sourceSheet = sourceSs.getSheetByName("Master Demand File Tracking");
  const targetSheet = targetSs.getSheetByName("Local Demand Log");
  
  if (!sourceSheet || !targetSheet) throw new Error("Synchronization pathway resource references invalid.");

  const fullDataMatrix = sourceSheet.getDataRange().getValues();
  targetSheet.clearContents();
  targetSheet.getRange(1, 1, fullDataMatrix.length, fullDataMatrix[0].length).setValues(fullDataMatrix);
}

/** * 5. VALIDATION FAULT SCANNING ROUTINE
 */
function runValidationScan() {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  let report = { am3Errors: [], am6Errors: [], am5Errors: [], am7Errors: [], am3Headers: [], am6Headers: [], am5Headers: [], am7Headers: [] };

  function cleanRowData(rowArray) {
    return rowArray.map(cell => {
      if (cell instanceof Date) return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
      return cell;
    });
  }

  // AM3 Scan (Col O)
  const am3Sheet = ss.getSheetByName("AP TLS AM3");
  if (am3Sheet) {
    const am3Data = am3Sheet.getDataRange().getValues();
    if (am3Data.length > 0) {
      report.am3Headers = am3Data[0].map(h => h.toString());
      for (let i = 1; i < am3Data.length; i++) {
        if (am3Data[i][14] && am3Data[i][14].toString().trim().toUpperCase() === "#N/A") {
          report.am3Errors.push(cleanRowData(am3Data[i]));
        }
      }
    }
  }

  // AM6 Scan (Col P)
  const am6Sheet = ss.getSheetByName("AP TLS AM6");
  if (am6Sheet) {
    const am6Data = am6Sheet.getDataRange().getValues();
    if (am6Data.length > 0) {
      report.am6Headers = am6Data[0].map(h => h.toString());
      for (let i = 1; i < am6Data.length; i++) {
        if (am6Data[i][15] && am6Data[i][15].toString().trim().toUpperCase() === "#N/A") {
          report.am6Errors.push(cleanRowData(am6Data[i]));
        }
      }
    }
  }

  // AM5 Scan (Col O)
  const am5Sheet = ss.getSheetByName("AP TLS AM5");
  if (am5Sheet) {
    const am5Data = am5Sheet.getDataRange().getValues();
    if (am5Data.length > 0) {
      report.am5Headers = am5Data[0].map(h => h.toString());
      for (let i = 1; i < am5Data.length; i++) {
        if (am5Data[i][14] && am5Data[i][14].toString().trim().toUpperCase() === "#N/A") {
          report.am5Errors.push(cleanRowData(am5Data[i]));
        }
      }
    }
  }

  // AM7 Scan (Col O)
  const am7Sheet = ss.getSheetByName("AP TLS AM7");
  if (am7Sheet) {
    const am7Data = am7Sheet.getDataRange().getValues();
    if (am7Data.length > 0) {
      report.am7Headers = am7Data[0].map(h => h.toString());
      for (let i = 1; i < am7Data.length; i++) {
        if (am7Data[i][14] && am7Data[i][14].toString().trim().toUpperCase() === "#N/A") {
          report.am7Errors.push(cleanRowData(am7Data[i]));
        }
      }
    }
  }

  return report;
}
