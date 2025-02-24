/**
 * UAN (Urgent Action Network) Reports Generator
 * 
 * This script processes campaign data from a Google Sheet to generate various reports
 * tracking supporter engagement across different dimensions like country, case number,
 * topics etc. The data is expected to be in a sheet named "processed-export" with 
 * specific columns including:
 * - Campaign ID: Unique identifier for each campaign
 * - Campaign Date: When the campaign action occurred
 * - Supporter ID: Unique identifier for supporters
 * - Supporter Email: Email address of supporters
 * - External Reference fields containing metadata like:
 *   - Country (Reference 6)
 *   - Case Number (Reference 7) 
 *   - Topics (Reference 8)
 *   - Year (Reference 10)
 *   - Type (Reference 10)
 */

/**
 * Creates the menu items in the Google Sheets UI.
 * Adds options to update all reports at once or individual report types.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Update UAN Reports")
    .addItem("Update All Reports", "processCampaignData")
    .addItem("Update All Reports (except by-supporter)", "processCampaignDataExceptSupporter")
    .addSeparator()
    .addItem("Update by-name", "updateCampaignReport")
    .addItem("Update by-case-number", "updateCaseReport") 
    .addItem("Update by-country", "updateCountryReport")
    .addItem("Update by-topic", "updateTopicReport")
    .addItem("Update by-year", "updateYearReport")
    .addItem("Update by-type", "updateTypeReport")
    .addItem("Update by-date", "updateDateReport")
    .addItem("Update by-supporter", "updateSupporterReport")
    .addToUi();
}

/**
 * Updates the date range information in the report tab
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 * @param {string} startDate - Start date or "Mixed"
 * @param {string} endDate - End date or "Mixed"
 */
function updateReportDates(ss, startDate, endDate) {
  var reportSheet = ss.getSheetByName("report");
  if (!reportSheet) {
    reportSheet = ss.insertSheet("report");
  }
  
  // Set labels if they don't exist
  if (reportSheet.getRange("A2").getValue() === "") {
    reportSheet.getRange("A2").setValue("Start Date");
  }
  if (reportSheet.getRange("A3").getValue() === "") {
    reportSheet.getRange("A3").setValue("End Date");
  }
  
  // Set values and alignment
  reportSheet.getRange("B2")
    .setValue(startDate || "")
    .setHorizontalAlignment("right");
  
  reportSheet.getRange("B3")
    .setValue(endDate || "")
    .setHorizontalAlignment("right");
}

/**
 * Main function to process campaign data and generate all reports.
 * Creates or updates sheets for each report type:
 * - by-name: Campaign ID counts
 * - by-case-number: Case number engagement
 * - by-country: Country-wise participation
 * - by-topic: Topic-wise breakdown
 * - by-year: Year-wise analysis
 * - by-type: Type-based categorization
 * - by-date: Monthly trends
 * - by-supporter: Individual supporter engagement
 */
function processCampaignData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var processedSheet = ss.getSheetByName("processed-export");
  var reportByNameSheet = getOrCreateSheet(ss, "by-name");
  var reportByCaseSheet = getOrCreateSheet(ss, "by-case-number");
  var reportByCountrySheet = getOrCreateSheet(ss, "by-country");
  var reportByTopicSheet = getOrCreateSheet(ss, "by-topic");
  var reportByYearSheet = getOrCreateSheet(ss, "by-year");
  var reportByTypeSheet = getOrCreateSheet(ss, "by-type");
  var reportByDateSheet = getOrCreateSheet(ss, "by-date");
  var reportBySupporterSheet = getOrCreateSheet(ss, "by-supporter");

  if (!processedSheet) {
    ui.alert("The 'processed-export' sheet is required.");
    return;
  }

  var startDateInput = ui.prompt("Enter Start Date (YYYY-MM-DD) or leave blank for no limit").getResponseText().trim();
  var endDateInput = ui.prompt("Enter End Date (YYYY-MM-DD) or leave blank for no limit").getResponseText().trim();

  var startDate = startDateInput ? new Date(startDateInput) : null;
  var endDate = endDateInput ? new Date(endDateInput) : null;

  // Update report dates
  updateReportDates(ss, startDateInput, endDateInput);

  var data = processedSheet.getDataRange().getValues();
  var headers = data[0];

  var campaignCol = headers.indexOf("Campaign ID");
  var campaignDateCol = headers.indexOf("Campaign Date");
  var supporterCol = headers.indexOf("Supporter ID");
  var emailCol = headers.indexOf("Supporter Email");

  var columns = {
    "Country": headers.indexOf("External Reference 6 (Country)"),
    "Case Number": headers.indexOf("External Reference 7 (Case Number)"),
    "Topics": headers.indexOf("External Reference 8 (Topics)"),
    "Year": headers.indexOf("External Reference 10 (Year)"),
    "Type": headers.indexOf("External Reference 10 (Type)")
  };

  if (campaignCol === -1 || campaignDateCol === -1 || supporterCol === -1 || emailCol === -1 || Object.values(columns).includes(-1)) {
    ui.alert("One or more required columns are missing.");
    return;
  }

  var campaignCounts = {};
  var caseCounts = {};
  var countryCounts = {};
  var topicCounts = {};
  var yearCounts = {};
  var typeCounts = {};
  var dateCounts = {};
  var supporterCounts = {};

  for (var i = 1; i < data.length; i++) {
    var campaignID = data[i][campaignCol];
    var campaignDateRaw = data[i][campaignDateCol];
    var supporterID = data[i][supporterCol];
    var supporterEmail = data[i][emailCol] || "";
    var campaignDate = "";

    if (campaignDateRaw) {
      campaignDate = campaignDateRaw instanceof Date
        ? Utilities.formatDate(campaignDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd")
        : campaignDateRaw.toString().trim();
    } else {
      continue;
    }

    var campaignDateObj = new Date(campaignDate);
    if (isNaN(campaignDateObj)) continue;

    if ((startDate && campaignDateObj < startDate) || (endDate && campaignDateObj > endDate)) continue;

    if (campaignID) {
      campaignCounts[campaignID] = (campaignCounts[campaignID] || 0) + 1;
    }

    var supporterKey = supporterID + " - " + supporterEmail;
    if (supporterID) {
      supporterCounts[supporterKey] = (supporterCounts[supporterKey] || 0) + 1;
    }

    var monthYear = Utilities.formatDate(campaignDateObj, Session.getScriptTimeZone(), "yyyy-MM");
    dateCounts[monthYear] = (dateCounts[monthYear] || 0) + 1;

    Object.keys(columns).forEach(key => {
      var cellValue = data[i][columns[key]];

      if (key === "Topics" && cellValue) {
        var topics = cellValue.split(",").map(topic => topic.trim()).filter(topic => topic !== "");
        topics.forEach(topic => {
          topicCounts[topic] = (topicCounts[topic] || 0) + 1;
        });
      } else if (cellValue && cellValue.trim() !== "") {
        if (key === "Country") countryCounts[cellValue.trim()] = (countryCounts[cellValue.trim()] || 0) + 1;
        if (key === "Case Number") caseCounts[cellValue.trim()] = (caseCounts[cellValue.trim()] || 0) + 1;
        if (key === "Year") yearCounts[cellValue.trim()] = (yearCounts[cellValue.trim()] || 0) + 1;
        if (key === "Type") typeCounts[cellValue.trim()] = (typeCounts[cellValue.trim()] || 0) + 1;
      }
    });
  }

  // Update this line to include unique supporters for the by-name report
  var uniqueCounts = calculateUniqueSupporters(data, campaignCol, supporterCol, campaignDateCol, startDate, endDate);
  writeSortedData(reportByNameSheet, ["Campaign ID", "Count", "Unique Supporters"], campaignCounts, uniqueCounts);
  
  // Keep the rest of the writeSortedData calls unchanged
  writeSortedData(reportByCaseSheet, ["Case Number", "Count"], caseCounts);
  writeSortedData(reportByCountrySheet, ["Country", "Count"], countryCounts);
  writeSortedData(reportByTopicSheet, ["Topic", "Count"], topicCounts);
  writeSortedData(reportByYearSheet, ["Year", "Count"], yearCounts);
  writeSortedData(reportByTypeSheet, ["Type", "Count"], typeCounts);
  writeSortedData(reportByDateSheet, ["Month", "Count"], dateCounts);
  writeSupporterData(reportBySupporterSheet, supporterCounts);

  ui.alert("✅ Your UAN Reports have been updated!");
}

/**
 * Helper function to get or create a sheet with the specified name.
 * Clears existing content if sheet already exists.
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 * @param {string} sheetName - Name of the sheet to get/create
 * @return {SpreadsheetApp.Sheet} The sheet object
 */
function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  return sheet;
}

/**
 * Writes sorted data to a sheet with consistent formatting.
 * Left-aligns the first column (keys) and right-aligns count columns.
 * @param {SpreadsheetApp.Sheet} sheet - Target sheet for the data
 * @param {Array} headers - Column headers
 * @param {Object} data - Key-value pairs to write
 * @param {Object} [uniqueCounts] - Optional unique counts data
 */
function writeSortedData(sheet, headers, data, uniqueCounts) {
  var sortedData = Object.entries(data).sort(([a], [b]) => a.localeCompare(b));

  sheet.appendRow(headers);
  sheet.getRange(1, 1).setHorizontalAlignment("left"); // Left-align the first header
  
  // Right-align count columns
  for (var i = 2; i <= headers.length; i++) {
    sheet.getRange(1, i).setHorizontalAlignment("right");
  }

  sortedData.forEach(([key, count], index) => {
    var rowData = uniqueCounts ? [key, count, uniqueCounts[key] || 0] : [key, count];
    sheet.appendRow(rowData);
    sheet.getRange(index + 2, 1).setHorizontalAlignment("left"); // Left-align key column
    
    // Right-align count columns
    for (var i = 2; i <= rowData.length; i++) {
      sheet.getRange(index + 2, i).setHorizontalAlignment("right");
    }
  });

  if (sheet.getLastRow() > 1) {
    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length);
    range.sort(1);
  }
}

/**
 * Special handler for supporter data to include both ID and email.
 * Creates a three-column report with supporter ID, email, and action count.
 * @param {SpreadsheetApp.Sheet} sheet - Target sheet for supporter data
 * @param {Object} data - Supporter data with counts
 */
function writeSupporterData(sheet, data) {
  var sortedData = Object.entries(data).sort(([a], [b]) => a.localeCompare(b));

  sheet.clear();
  sheet.appendRow(["Supporter ID", "Supporter Email", "Count"]);
  sheet.getRange(1, 1).setHorizontalAlignment("left"); // Left-align "Supporter ID" header
  sheet.getRange(1, 2).setHorizontalAlignment("left"); // Left-align "Supporter Email" header
  sheet.getRange(1, 3).setHorizontalAlignment("right"); // Right-align "Count" header

  sortedData.forEach(([key, count], index) => {
    var splitKey = key.split(" - ");
    var supporterID = splitKey[0];
    var supporterEmail = splitKey.length > 1 ? splitKey[1] : "";
    sheet.appendRow([supporterID, supporterEmail, count]);
    sheet.getRange(index + 2, 1).setHorizontalAlignment("left"); // Left-align "Supporter ID" cells
    sheet.getRange(index + 2, 2).setHorizontalAlignment("left"); // Left-align "Supporter Email" cells
    sheet.getRange(index + 2, 3).setHorizontalAlignment("right"); // Right-align "Count" cells
  });

  if (sheet.getLastRow() > 1) {
    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3);
    range.sort(1);
  }
}

/**
 * Individual report generation functions.
 * Each function processes data for a specific report type.
 */
function updateCampaignReport() {
  processSpecificReport("by-name", "Campaign ID");
}

function updateCaseReport() {
  processSpecificReport("by-case-number", "Case Number");
}

function updateCountryReport() {
  processSpecificReport("by-country", "Country");
}

function updateTopicReport() {
  processSpecificReport("by-topic", "Topics");
}

function updateYearReport() {
  processSpecificReport("by-year", "Year");
}

function updateTypeReport() {
  processSpecificReport("by-type", "Type");
}

function updateDateReport() {
  processSpecificReport("by-date", "Date");
}

function updateSupporterReport() {
  processSpecificReport("by-supporter", "Supporter");
}

/**
 * Core function for processing individual report types.
 * Handles date filtering and data processing for specific dimensions.
 * @param {string} sheetName - Name of the target sheet
 * @param {string} reportType - Type of report to generate
 */
function processSpecificReport(sheetName, reportType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var processedSheet = ss.getSheetByName("processed-export");
  var reportSheet = getOrCreateSheet(ss, sheetName);

  if (!processedSheet) {
    ui.alert("The 'processed-export' sheet is required.");
    return;
  }

  var startDateInput = ui.prompt("Enter Start Date (YYYY-MM-DD) or leave blank for no limit").getResponseText().trim();
  var endDateInput = ui.prompt("Enter End Date (YYYY-MM-DD) or leave blank for no limit").getResponseText().trim();

  var startDate = startDateInput ? new Date(startDateInput) : null;
  var endDate = endDateInput ? new Date(endDateInput) : null;

  // Update report dates to "Mixed" when running individual reports
  updateReportDates(ss, "Mixed", "Mixed");

  var data = processedSheet.getDataRange().getValues();
  var headers = data[0];

  // Get required column indices
  var campaignDateCol = headers.indexOf("Campaign Date");
  var columns = getRequiredColumns(headers, reportType);
  
  if (campaignDateCol === -1 || Object.values(columns).includes(-1)) {
    ui.alert("One or more required columns are missing.");
    return;
  }

  var counts = processData(data, campaignDateCol, columns, reportType, startDate, endDate);

  // Write the data based on report type
  if (reportType === "Supporter") {
    writeSupporterData(reportSheet, counts);
  } else {
    var headerNames = getHeaderNames(reportType);
    
    // For Campaign ID report, calculate unique supporters
    if (reportType === "Campaign ID") {
      var uniqueCounts = calculateUniqueSupporters(data, columns["Campaign ID"], columns["Supporter ID"], campaignDateCol, startDate, endDate);
      writeSortedData(reportSheet, ["Campaign ID", "Count", "Unique Supporters"], counts, uniqueCounts);
    } else {
      writeSortedData(reportSheet, headerNames, counts);
    }
  }

  ui.alert(`✅ Your ${reportType} Report has been updated!`);
}

/**
 * Maps report types to their corresponding column indices.
 * @param {Array} headers - Sheet headers
 * @param {string} reportType - Type of report
 * @return {Object} Column mapping for the report type
 */
function getRequiredColumns(headers, reportType) {
  var columns = {};
  
  switch(reportType) {
    case "Campaign ID":
      columns["Campaign ID"] = headers.indexOf("Campaign ID");
      columns["Supporter ID"] = headers.indexOf("Supporter ID");
      break;
    case "Case Number":
      columns["Case Number"] = headers.indexOf("External Reference 7 (Case Number)");
      break;
    case "Country":
      columns["Country"] = headers.indexOf("External Reference 6 (Country)");
      break;
    case "Topics":
      columns["Topics"] = headers.indexOf("External Reference 8 (Topics)");
      break;
    case "Year":
      columns["Year"] = headers.indexOf("External Reference 10 (Year)");
      break;
    case "Type":
      columns["Type"] = headers.indexOf("External Reference 10 (Type)");
      break;
    case "Date":
      columns["Campaign Date"] = headers.indexOf("Campaign Date");
      break;
    case "Supporter":
      columns["Supporter ID"] = headers.indexOf("Supporter ID");
      columns["Supporter Email"] = headers.indexOf("Supporter Email");
      break;
  }
  
  return columns;
}

/**
 * Returns appropriate header names for each report type.
 * @param {string} reportType - Type of report
 * @return {Array} Header names for the report
 */
function getHeaderNames(reportType) {
  switch(reportType) {
    case "Campaign ID": return ["Campaign ID", "Count"];
    case "Case Number": return ["Case Number", "Count"];
    case "Country": return ["Country", "Count"];
    case "Topics": return ["Topic", "Count"];
    case "Year": return ["Year", "Count"];
    case "Type": return ["Type", "Count"];
    case "Date": return ["Month", "Count"];
    default: return ["Item", "Count"];
  }
}

/**
 * Processes raw data to generate counts for the specified report type.
 * Handles date filtering and different counting logic for each report type.
 * @param {Array} data - Raw sheet data
 * @param {number} campaignDateCol - Index of campaign date column
 * @param {Object} columns - Required column indices
 * @param {string} reportType - Type of report
 * @param {Date} startDate - Optional start date filter
 * @param {Date} endDate - Optional end date filter
 * @return {Object} Processed counts for the report
 */
function processData(data, campaignDateCol, columns, reportType, startDate, endDate) {
  var counts = {};

  for (var i = 1; i < data.length; i++) {
    var campaignDateRaw = data[i][campaignDateCol];
    if (!campaignDateRaw) continue;

    var campaignDate = campaignDateRaw instanceof Date
      ? Utilities.formatDate(campaignDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : campaignDateRaw.toString().trim();

    var campaignDateObj = new Date(campaignDate);
    if (isNaN(campaignDateObj)) continue;

    if ((startDate && campaignDateObj < startDate) || (endDate && campaignDateObj > endDate)) continue;

    switch(reportType) {
      case "Campaign ID":
      case "Case Number":
      case "Country":
      case "Year":
      case "Type":
        var value = data[i][columns[reportType]];
        if (value && value.trim() !== "") {
          counts[value.trim()] = (counts[value.trim()] || 0) + 1;
        }
        break;
      case "Topics":
        var topics = data[i][columns["Topics"]];
        if (topics) {
          topics.split(",").map(topic => topic.trim()).filter(topic => topic !== "")
            .forEach(topic => {
              counts[topic] = (counts[topic] || 0) + 1;
            });
        }
        break;
      case "Date":
        var monthYear = Utilities.formatDate(campaignDateObj, Session.getScriptTimeZone(), "yyyy-MM");
        counts[monthYear] = (counts[monthYear] || 0) + 1;
        break;
      case "Supporter":
        var supporterId = data[i][columns["Supporter ID"]];
        var supporterEmail = data[i][columns["Supporter Email"]] || "";
        if (supporterId) {
          var supporterKey = supporterId + " - " + supporterEmail;
          counts[supporterKey] = (counts[supporterKey] || 0) + 1;
        }
        break;
    }
  }

  return counts;
}

/**
 * Calculates unique supporters for each campaign.
 * @param {Array} data - Raw sheet data
 * @param {number} campaignCol - Campaign ID column index
 * @param {number} supporterCol - Supporter ID column index
 * @param {number} dateCol - Campaign date column index
 * @param {Date} startDate - Optional start date filter
 * @param {Date} endDate - Optional end date filter
 * @return {Object} Campaign IDs mapped to their unique supporter counts
 */
function calculateUniqueSupporters(data, campaignCol, supporterCol, dateCol, startDate, endDate) {
  var uniqueCounts = {};
  var campaignSupporters = {};

  for (var i = 1; i < data.length; i++) {
    var campaignDateRaw = data[i][dateCol];
    if (!campaignDateRaw) continue;

    var campaignDate = campaignDateRaw instanceof Date
      ? Utilities.formatDate(campaignDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : campaignDateRaw.toString().trim();

    var campaignDateObj = new Date(campaignDate);
    if (isNaN(campaignDateObj)) continue;

    if ((startDate && campaignDateObj < startDate) || (endDate && campaignDateObj > endDate)) continue;

    var campaignID = data[i][campaignCol];
    var supporterID = data[i][supporterCol];

    if (campaignID && supporterID) {
      if (!campaignSupporters[campaignID]) {
        campaignSupporters[campaignID] = new Set();
      }
      campaignSupporters[campaignID].add(supporterID);
    }
  }

  // Convert Sets to counts
  Object.keys(campaignSupporters).forEach(campaignID => {
    uniqueCounts[campaignID] = campaignSupporters[campaignID].size;
  });

  return uniqueCounts;
}

/**
 * Processes all reports except the supporter report.
 * Identical to processCampaignData but skips the by-supporter tab.
 */
function processCampaignDataExceptSupporter() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var processedSheet = ss.getSheetByName("processed-export");
  var reportByNameSheet = getOrCreateSheet(ss, "by-name");
  var reportByCaseSheet = getOrCreateSheet(ss, "by-case-number");
  var reportByCountrySheet = getOrCreateSheet(ss, "by-country");
  var reportByTopicSheet = getOrCreateSheet(ss, "by-topic");
  var reportByYearSheet = getOrCreateSheet(ss, "by-year");
  var reportByTypeSheet = getOrCreateSheet(ss, "by-type");
  var reportByDateSheet = getOrCreateSheet(ss, "by-date");

  if (!processedSheet) {
    ui.alert("The 'processed-export' sheet is required.");
    return;
  }

  var startDateInput = ui.prompt("Enter Start Date (YYYY-MM-DD) or leave blank for no limit").getResponseText().trim();
  var endDateInput = ui.prompt("Enter End Date (YYYY-MM-DD) or leave blank for no limit").getResponseText().trim();

  var startDate = startDateInput ? new Date(startDateInput) : null;
  var endDate = endDateInput ? new Date(endDateInput) : null;

  // Update report dates to "Mixed" since by-supporter isn't included
  updateReportDates(ss, "Mixed", "Mixed");

  var data = processedSheet.getDataRange().getValues();
  var headers = data[0];

  var campaignCol = headers.indexOf("Campaign ID");
  var campaignDateCol = headers.indexOf("Campaign Date");
  var supporterCol = headers.indexOf("Supporter ID");
  var emailCol = headers.indexOf("Supporter Email");

  var columns = {
    "Country": headers.indexOf("External Reference 6 (Country)"),
    "Case Number": headers.indexOf("External Reference 7 (Case Number)"),
    "Topics": headers.indexOf("External Reference 8 (Topics)"),
    "Year": headers.indexOf("External Reference 10 (Year)"),
    "Type": headers.indexOf("External Reference 10 (Type)")
  };

  if (campaignCol === -1 || campaignDateCol === -1 || supporterCol === -1 || emailCol === -1 || Object.values(columns).includes(-1)) {
    ui.alert("One or more required columns are missing.");
    return;
  }

  var campaignCounts = {};
  var caseCounts = {};
  var countryCounts = {};
  var topicCounts = {};
  var yearCounts = {};
  var typeCounts = {};
  var dateCounts = {};

  // Process the data (same as processCampaignData but without supporter counting)
  for (var i = 1; i < data.length; i++) {
    var campaignDateRaw = data[i][campaignDateCol];
    if (!campaignDateRaw) continue;

    var campaignDate = campaignDateRaw instanceof Date
      ? Utilities.formatDate(campaignDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd")
      : campaignDateRaw.toString().trim();

    var campaignDateObj = new Date(campaignDate);
    if (isNaN(campaignDateObj)) continue;

    if ((startDate && campaignDateObj < startDate) || (endDate && campaignDateObj > endDate)) continue;

    var campaignID = data[i][campaignCol];
    if (campaignID) {
      campaignCounts[campaignID] = (campaignCounts[campaignID] || 0) + 1;
    }

    var monthYear = Utilities.formatDate(campaignDateObj, Session.getScriptTimeZone(), "yyyy-MM");
    dateCounts[monthYear] = (dateCounts[monthYear] || 0) + 1;

    Object.keys(columns).forEach(key => {
      var cellValue = data[i][columns[key]];

      if (key === "Topics" && cellValue) {
        var topics = cellValue.split(",").map(topic => topic.trim()).filter(topic => topic !== "");
        topics.forEach(topic => {
          topicCounts[topic] = (topicCounts[topic] || 0) + 1;
        });
      } else if (cellValue && cellValue.trim() !== "") {
        if (key === "Country") countryCounts[cellValue.trim()] = (countryCounts[cellValue.trim()] || 0) + 1;
        if (key === "Case Number") caseCounts[cellValue.trim()] = (caseCounts[cellValue.trim()] || 0) + 1;
        if (key === "Year") yearCounts[cellValue.trim()] = (yearCounts[cellValue.trim()] || 0) + 1;
        if (key === "Type") typeCounts[cellValue.trim()] = (typeCounts[cellValue.trim()] || 0) + 1;
      }
    });
  }

  // Write all reports except supporter
  var uniqueCounts = calculateUniqueSupporters(data, campaignCol, supporterCol, campaignDateCol, startDate, endDate);
  writeSortedData(reportByNameSheet, ["Campaign ID", "Count", "Unique Supporters"], campaignCounts, uniqueCounts);
  writeSortedData(reportByCaseSheet, ["Case Number", "Count"], caseCounts);
  writeSortedData(reportByCountrySheet, ["Country", "Count"], countryCounts);
  writeSortedData(reportByTopicSheet, ["Topic", "Count"], topicCounts);
  writeSortedData(reportByYearSheet, ["Year", "Count"], yearCounts);
  writeSortedData(reportByTypeSheet, ["Type", "Count"], typeCounts);
  writeSortedData(reportByDateSheet, ["Month", "Count"], dateCounts);

  ui.alert("✅ Your UAN Reports have been updated! (except by-supporter)");
}