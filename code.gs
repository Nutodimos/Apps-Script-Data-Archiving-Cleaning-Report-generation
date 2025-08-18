
// --- Configuration ---
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId(); // Gets the ID of the current spreadsheet
const RAW_DATA_SHEET_NAME = "Raw Data";
const ARCHIVE_SHEET_NAME = "Archive";
const REPORTS_SHEET_NAME = "Reports";
const HEADER_ROW_INDEX = 1; // Assuming headers are in the first row
// Define headers for your sheets
const HEADERS = ["ID", "Timestamp", "Name", "Email", "Phone", "Order Value", "Status"];

const COL_ID = 0;
const COL_TIMESTAMP = 1;
const COL_NAME = 2;
const COL_EMAIL = 3;
const COL_PHONE = 4;
const COL_VALUE = 5;
const COL_STATUS = 6;

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Data Automation')
    .addItem('Setup Sheets (Run Once)', 'setupSheets')
    .addItem('Open Sidebar UI', 'showSidebar')
    .addSeparator()
    .addItem('Process & Archive Data Now', 'processAndArchiveData')
    .addItem('Generate Daily Report Now', 'generateDailyReport')
    .addToUi();
}

// function showSidebar() {
//  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
//   .setTitle('Data Automation');
// SpreadsheetApp.getUi().showSidebar(html);
//   const output = html.evaluate().setTitle('Data Automation');
//   SpreadsheetApp.getUi().showSidebar(output);
// }
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Data Automation');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getThresholds() {
  const props = PropertiesService.getDocumentProperties();
  const stored = props.getProperty("valueThresholds");
  return stored ? JSON.parse(stored) : { high: 1000, medium: 500, low: 0 };
}

function saveThresholds(thresholds) {
  PropertiesService.getDocumentProperties().setProperty("valueThresholds", JSON.stringify(thresholds));
}


// --- Setup Function (Run Once) ---
/**
 * Sets up the necessary sheets if they don't exist.
 * This function should be run once manually from the custom menu.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // ARCHIVE SHEET
let archiveSheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
if (!archiveSheet) {
  archiveSheet = ss.insertSheet(ARCHIVE_SHEET_NAME);
  archiveSheet.appendRow(HEADERS);
  Logger.log(`Created sheet: ${ARCHIVE_SHEET_NAME}`);
} else {
  const firstRow = archiveSheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  const headersMissing = firstRow.every(cell => cell === "" || cell === null);
  const headersWrong = !firstRow.every((val, i) => val === HEADERS[i]);

  if (headersMissing || headersWrong) {
    archiveSheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    Logger.log(`Set or corrected headers in: ${ARCHIVE_SHEET_NAME}`);
  }
}


    // RAW DATA SHEET
    let rawDataSheet = ss.getSheetByName(RAW_DATA_SHEET_NAME);
    if (!rawDataSheet) {
      rawDataSheet = ss.insertSheet(RAW_DATA_SHEET_NAME);
      rawDataSheet.appendRow(HEADERS);
      Logger.log(`Created sheet: ${RAW_DATA_SHEET_NAME}`);
    }

    // REPORTS SHEET
    let reportsSheet = ss.getSheetByName(REPORTS_SHEET_NAME);
    if (!reportsSheet) {
      reportsSheet = ss.insertSheet(REPORTS_SHEET_NAME);
      reportsSheet.getRange(1, 1).setValue("Daily Data Report");
      Logger.log(`Created sheet: ${REPORTS_SHEET_NAME}`);
    }

    ui.alert('Setup Complete', 'Sheets created or verified successfully.', ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(`Error during setupSheets: ${e.message}`);
    ui.alert('Setup Failed', `An error occurred: ${e.message}`, ui.ButtonSet.OK);
  }
}
//--- Cleans and StandardizeRow-----
function cleanAndStandardizeRow(row) {
  const cleanedRow = [...row];

  // Trim all string cells
  for (let i = 0; i < cleanedRow.length; i++) {
    if (typeof cleanedRow[i] === 'string') {
      cleanedRow[i] = cleanedRow[i].trim();
    }
  }

  // Proper-case Name
  if (cleanedRow[COL_NAME] && typeof cleanedRow[COL_NAME] === 'string') {
    cleanedRow[COL_NAME] = toProperCase(cleanedRow[COL_NAME]);
  }

  // Lowercase Email
  if (cleanedRow[COL_EMAIL] && typeof cleanedRow[COL_EMAIL] === 'string') {
    cleanedRow[COL_EMAIL] = cleanedRow[COL_EMAIL].toLowerCase();
  }
  // Standardize Timestamp
  const ts = cleanedRow[COL_TIMESTAMP];
  if (ts && !(ts instanceof Date)) {
    try {
      cleanedRow[COL_TIMESTAMP] = new Date(ts);
    } catch (e) {
      Logger.log(`Invalid timestamp: ${ts}`);
    }
  }
  return cleanedRow;
}

function toProperCase(str) {
  return str.replace(/\w\S*/g, function (txt) {
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });
}

//---Obtain archive sheet for report generation---
function getArchiveSheetNames() {
  try {
    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    const archiveNames = sheets
      .map(sheet => sheet.getName())
      .filter(name => name.startsWith("Archive"))
      .sort(); // Sort them chronologically
    
    console.log('Found archive sheets:', archiveNames);
    return archiveNames;
  } catch (error) {
    console.error('Error getting archive sheet names:', error);
    return [];
  }
}


// --- Main Processing Function ---
function processAndArchiveData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawDataSheet = ss.getSheetByName(RAW_DATA_SHEET_NAME);
  const today = new Date();
const archiveSheetName = `Archive_${today.getFullYear()}-${(today.getMonth()+1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
let archiveSheet = ss.getSheetByName(archiveSheetName);

if (!archiveSheet) {
  archiveSheet = ss.insertSheet(archiveSheetName);
  archiveSheet.appendRow(HEADERS);
} else {
  Logger.log(`Archive for today (${archiveSheetName}) already exists.`);
}

  const ui = SpreadsheetApp.getUi();

  if (!rawDataSheet) {
    ui.alert('Error', 'Required sheets ("Raw Data" or "Archive") not found. Please run "Setup Sheets" first.', ui.ButtonSet.OK);
    return;
  }

  try {
    const lastRow = rawDataSheet.getLastRow();
    if (lastRow <= HEADER_ROW_INDEX) {
      Logger.log('No new data to process.');
      ui.alert('No New Data', 'The "Raw Data" sheet is empty or only contains headers.', ui.ButtonSet.OK);
      return;
    }

    // STEP 1: Get raw data
    const rawDataRange = rawDataSheet.getRange(
      HEADER_ROW_INDEX + 1,
      1,
      lastRow - HEADER_ROW_INDEX,
      rawDataSheet.getLastColumn()
    );
    let data = rawDataRange.getValues(); // Batch read

    Logger.log(`Fetched ${data.length} rows from Raw Data.`);

    // STEP 2: Clean and standardize
    data = data.map(row => cleanAndStandardizeRow(row));

    // STEP 3: Remove duplicates (based on name + email combo)
    data = removeDuplicates(data, row => `${row[COL_NAME]}|${row[COL_EMAIL]}`);

    // STEP 4: Apply conditional logic (status column)
    data = data.map(row => applyConditionalLogic(row));

    // STEP 5: Read archive and collect existing IDs
    const archiveLastRow = archiveSheet.getLastRow();
    let existingArchiveData = [];
    if (archiveLastRow > HEADER_ROW_INDEX) {
      existingArchiveData = archiveSheet.getRange(
        HEADER_ROW_INDEX + 1,
        1,
        archiveLastRow - HEADER_ROW_INDEX,
        archiveSheet.getLastColumn()
      ).getValues();
    }
    const existingArchiveIds = new Set(existingArchiveData.map(row => row[COL_ID]));

    // STEP 6: Filter only new rows based on ID
    const newDataToArchive = data.filter(row => !existingArchiveIds.has(row[COL_ID]));

    // STEP 7: Append to archive
    if (newDataToArchive.length > 0) {
     const startRow = archiveSheet.getLastRow() < HEADER_ROW_INDEX ? HEADER_ROW_INDEX + 1 : archiveSheet.getLastRow() + 1;

archiveSheet.getRange(
  startRow,
  1,
  newDataToArchive.length,
  newDataToArchive[0].length
).setValues(newDataToArchive);

      Logger.log(`Archived ${newDataToArchive.length} new rows.`);
    } else {
      Logger.log('No new unique data to archive.');
    }

    // STEP 8: Clear original data (keep header)
    rawDataSheet.getRange(
      HEADER_ROW_INDEX + 1,
      1,
      lastRow - HEADER_ROW_INDEX,
      rawDataSheet.getLastColumn()
    ).clearContent();
    Logger.log('Raw Data sheet cleared.');

    ui.alert('Process Complete', `Archived ${newDataToArchive.length} rows and cleared Raw Data.`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log(`Error in processAndArchiveData: ${e.message}`);
    ui.alert('Processing Failed', `Error during processing: ${e.message}`, ui.ButtonSet.OK);
  }
}

function standardizePhoneNumber(str) {
  if (!str) return ''; // Handle null, undefined, or empty cells

  // Convert to string if it's a number
  const cleaned = String(str).replace(/\D/g, '');

  // Optionally format as international
  if (cleaned.startsWith('0')) {
    return '+234' + cleaned.slice(1);  // Nigerian format
  }

  return cleaned;
}

function toProperCase(str) {
  return str.replace(/\w\S*/g, function(txt) {
    return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
  });
}

function removeDuplicates(data, uniqueKeyFn) {
  const seen = new Set();
  const uniqueData = [];

  data.forEach(row => {
    const key = uniqueKeyFn(row);
    if (!seen.has(key)) {
      seen.add(key);
      uniqueData.push(row);
    }
  });

  return uniqueData;
}

function applyConditionalLogic(row) {
  // Get thresholds from the same source as saveThresholds
  const thresholds = getThresholds(); // Use your existing getThresholds function
  const high = Number(thresholds.high) || 1000;
  const medium = Number(thresholds.medium) || 500;
  const low = Number(thresholds.low) || 0;
  
  const value = Number(row[COL_VALUE]) || 0;

  if (value >= high) {
    row[COL_STATUS] = 'High';
  } else if (value >= medium) {
    row[COL_STATUS] = 'Medium';
  } else if (value>= low){
    row[COL_STATUS] = 'Low';
  }
   else {
    // If Value is not a number, set status to 'N/A' or keep existing
    row[COL_STATUS] = "N/A";
  }


  // Example: Flag rows based on a specific category (assuming 'Urgent' is a category)
  // If the category is 'Urgent', add a note or modify status further
  if (row && typeof row === 'string' && row.toLowerCase() === 'urgent') {
    // This example adds to notes, you could modify status if preferred
    row = (row? row + '; ' : '') + 'Urgent Category Flag';
  }

  return row;
}

// --- Reporting Function ---
function generateDailyReport(chartType, archiveName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  //const archiveSheet = ss.getSheetByName(archiveName);
  // Use the passed archiveName parameter, or fall back to today's archive
  let selectedArchiveName = archiveName;
  if (!selectedArchiveName) {
    const today = new Date();
    selectedArchiveName = `Archive_${today.getFullYear()}-${(today.getMonth()+1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
  }

  const archiveSheet = ss.getSheetByName(selectedArchiveName);
  const reportsSheet = ss.getSheetByName(REPORTS_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  if (!archiveSheet || !reportsSheet) {
    ui.alert("Error", `Archive sheet "${selectedArchiveName}" or Reports sheet not found.`, ui.ButtonSet.OK);
    return;
  }

  try {
    // Clear any existing charts first
    const charts = reportsSheet.getCharts();
    for (let i = 0; i < charts.length; i++) {
      reportsSheet.removeChart(charts[i]);
    }

    const archiveLastRow = archiveSheet.getLastRow();
    if (archiveLastRow <= HEADER_ROW_INDEX) {
      reportsSheet.clearContents();
      reportsSheet.getRange(1, 1).setValue(`Daily Data Report - ${selectedArchiveName}`);
      reportsSheet.getRange(2, 1).setValue("No data in archive to report.");
      ui.alert("Report Generated", "Report created but no data was available.", ui.ButtonSet.OK);
      return;
    }

    const headers = archiveSheet.getRange(HEADER_ROW_INDEX, 1, 1, archiveSheet.getLastColumn()).getValues()[0];
    console.log('Headers:', headers);
    console.log('COL_VALUE points to header:', headers[COL_VALUE]);
    
    const data = archiveSheet.getRange(
      HEADER_ROW_INDEX + 1,
      1,
      archiveLastRow - HEADER_ROW_INDEX,
      archiveSheet.getLastColumn()
    ).getValues();

    const thresholds = JSON.parse(PropertiesService.getDocumentProperties().getProperty("valueThresholds") || '{}');
    const highThreshold = thresholds.high != null ? Number(thresholds.high) : 1000;
    const mediumThreshold = thresholds.medium != null ? Number(thresholds.medium) : 500;
    const lowThreshold = thresholds.low != null ? Number(thresholds.low) : 0;

    const valueLevelCounts = { High: 0, Medium: 0, Low: 0 };
    let totalValue = 0;

    data.forEach((row, index) => {
      let rawValue = row[COL_VALUE];
      let value = 0;
      
      if (typeof rawValue === "number") {
        value = rawValue;
      } else if (typeof rawValue === "string") {
        const parsed = parseFloat(rawValue.toString().replace(/[^0-9.-]/g, ''));
        value = isNaN(parsed) ? 0 : parsed;
      } else {
        value = 0;
      }

      let level = "Low";
      if (value >= highThreshold) {
        level = "High";
      } else if (value >= mediumThreshold) {
        level = "Medium";
      }
      
      valueLevelCounts[level]++;
      totalValue += value;
    });

    // Report building
    const reportData = [
      [`Daily Report Summary - ${selectedArchiveName}`],
      [],
      ["Total Order Value", totalValue],
      [],
      ["Value Level Breakdown"],
      ["High", valueLevelCounts["High"]],
      ["Medium", valueLevelCounts["Medium"]],
      ["Low", valueLevelCounts["Low"]],
      [],
      ["Generated At", new Date().toLocaleString()]
    ];

    const normalizedReportData = reportData.map(row => {
      if (row.length === 0) {
        return ["", ""]; 
      } else if (row.length === 1) {
        return [row[0], ""];
      }
      return row; 
    });

    reportsSheet.clearContents();

    if (normalizedReportData.length > 0) {
      reportsSheet.getRange(1, 1, normalizedReportData.length, 2).setValues(normalizedReportData);
    }

    // Chart creation (use passed chartType or get from properties)
    const finalChartType = chartType || PropertiesService.getDocumentProperties().getProperty("chartType");
    
    if (finalChartType && finalChartType !== "none") {
      const breakdownRowIndex = normalizedReportData.findIndex(row => row[0] === "Value Level Breakdown");
      
      if (breakdownRowIndex !== -1) {
        const chartDataStartRow = breakdownRowIndex + 2;
        const chartDataRange = reportsSheet.getRange(chartDataStartRow, 1, 3, 2);
        
        let chartBuilder;
        let chartTitle = "Value Level Distribution";
        
        if (finalChartType === "pie") {
          chartBuilder = reportsSheet.newChart()
            .asPieChart()
            .addRange(chartDataRange)
            .setOption("title", chartTitle)
            .setOption("pieHole", 0.4)
            .setOption("legend", {position: "right", alignment: "center"});
        } else if (finalChartType === "bar") {
          chartBuilder = reportsSheet.newChart()
            .asColumnChart()
            .addRange(chartDataRange)
            .setOption("title", chartTitle)
            .setOption("legend", {position: "none"})
            .setOption("hAxis", {title: "Value Level"})
            .setOption("vAxis", {title: "Count"});
        }
        
        if (chartBuilder) {
          chartBuilder.setPosition(normalizedReportData.length + 2, 1, 0, 0);
          const chart = chartBuilder.build();
          reportsSheet.insertChart(chart);
        }
      }
    }

    ui.alert("Report Generated", `The report for ${selectedArchiveName} has been successfully created.`, ui.ButtonSet.OK);

  } catch (e) {
    Logger.log("Error in generateDailyReport: " + e.message);
    console.error("Error in generateDailyReport: " + e.message);
    ui.alert("Report Generation Failed", `An error occurred: ${e.message}`, ui.ButtonSet.OK);
  }
}
function saveChartPreference(chartType) {
  try {
    PropertiesService.getDocumentProperties().setProperty("chartType", chartType);
    Logger.log(`Chart preference saved: ${chartType}`);
  } catch (error) {
    Logger.log(`Error saving chart preference: ${error.message}`);
    throw error;
  }
}

function getChartPreference() {
  const props = PropertiesService.getDocumentProperties();
  return props.getProperty("chartType") || "none";
}

// --- Trigger Management (Optional, for automation) ---
function createDailyProcessTrigger() {
  // Delete existing triggers to avoid duplicates
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'processAndArchiveData') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }

  // Create a new daily trigger
  ScriptApp.newTrigger('processAndArchiveData')
    .timeBased()
    .everyDays(1) // Runs once every day
    .atHour(2) // You can set a specific hour, e.g., 2 AM
    .create();
  Logger.log('Daily trigger for processAndArchiveData created.');
  SpreadsheetApp.getUi().alert('Automation Set', 'Daily data processing trigger set for 2 AM (or your chosen hour).', SpreadsheetApp.getUi().ButtonSet.OK);
}


 ///Sets up a time-driven trigger to run generateDailyReport daily.
 
function createDailyReportTrigger() {
  // Delete existing triggers to avoid duplicates
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getHandlerFunction() === 'generateDailyReport') {
      ScriptApp.deleteTrigger(allTriggers[i]);
    }
  }

  // Create a new daily trigger
  ScriptApp.newTrigger('generateDailyReport')
    .timeBased()
    .everyDays(1) // Runs once every day
    .atHour(3) // Set it to run after data processing, e.g., 3 AM
    .create();
  Logger.log('Daily trigger for generateDailyReport created.');
  SpreadsheetApp.getUi().alert('Automation Set', 'Daily report generation trigger set for 3 AM (or your chosen hour).', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Deletes all triggers associated with the current project.
 * Useful for cleaning up or resetting automation.
 */
function deleteAllTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  Logger.log('All project triggers deleted.');
  SpreadsheetApp.getUi().alert('Automation Removed', 'All automation triggers for this project have been deleted.', SpreadsheetApp.getUi().ButtonSet.OK);
}

