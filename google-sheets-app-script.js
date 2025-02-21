function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Update UAN Reports")
    .addItem("Run Report", "processCampaignData")
    .addToUi();
}

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

  writeSortedData(reportByNameSheet, ["Campaign ID", "Count"], campaignCounts);
  writeSortedData(reportByCaseSheet, ["Case Number", "Count"], caseCounts);
  writeSortedData(reportByCountrySheet, ["Country", "Count"], countryCounts);
  writeSortedData(reportByTopicSheet, ["Topic", "Count"], topicCounts);
  writeSortedData(reportByYearSheet, ["Year", "Count"], yearCounts);
  writeSortedData(reportByTypeSheet, ["Type", "Count"], typeCounts);
  writeSortedData(reportByDateSheet, ["Month", "Count"], dateCounts);
  writeSupporterData(reportBySupporterSheet, supporterCounts);

  ui.alert("✅ Your UAN Reports have been updated!");
}

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }
  return sheet;
}

function writeSortedData(sheet, headers, data) {
  var sortedData = Object.entries(data).sort(([a], [b]) => a.localeCompare(b));

  sheet.appendRow(headers);
  sheet.getRange(1, 2).setHorizontalAlignment("right"); // Right-align the "Count" header
  sheet.getRange(1, 1).setHorizontalAlignment("left"); // Left-align the "Year" header (or similar columns)

  sortedData.forEach(([key, count], index) => {
    sheet.appendRow([key, count]);
    sheet.getRange(index + 2, 1).setHorizontalAlignment("left"); // Left-align each "Year" cell
    sheet.getRange(index + 2, 2).setHorizontalAlignment("right"); // Right-align each "Count" cell
  });

  if (sheet.getLastRow() > 1) {
    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2);
    range.sort(1);
  }
}

function writeSupporterData(sheet, data) {
  var sortedData = Object.entries(data).sort(([a], [b]) => a.localeCompare(b));

  sheet.clear();
  sheet.appendRow(["Supporter ID", "Supporter Email", "Count"]);
  sheet.getRange(1, 1).setHorizontalAlignment("left"); // Left-align "Supporter ID" header
  sheet.getRange(1, 3).setHorizontalAlignment("right"); // Right-align "Count" header

  sortedData.forEach(([key, count], index) => {
    var splitKey = key.split(" - ");
    var supporterID = splitKey[0];
    var supporterEmail = splitKey.length > 1 ? splitKey[1] : "";
    sheet.appendRow([supporterID, supporterEmail, count]);
    sheet.getRange(index + 2, 1).setHorizontalAlignment("left"); // Left-align "Supporter ID" cells
    sheet.getRange(index + 2, 3).setHorizontalAlignment("right"); // Right-align "Count" cells
  });

  if (sheet.getLastRow() > 1) {
    var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3);
    range.sort(1);
  }
}