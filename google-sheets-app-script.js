function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Update UAN Reports")
    .addItem("Run Report", "processCampaignData")
    .addToUi();
}

function processCampaignData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var exportSheet = ss.getSheetByName("export");
  var reportByNameSheet = ss.getSheetByName("reporting-by-name");
  var reportByERSheet = ss.getSheetByName("reporting-by-er");
  var reportByDateSheet = ss.getSheetByName("reporting-by-date");
  var reportBySupporterSheet = ss.getSheetByName("reporting-by-supporter-id");

  if (!exportSheet) {
    ui.alert("The 'export' sheet is required.");
    return;
  }
  if (!reportByNameSheet) {
    reportByNameSheet = ss.insertSheet("reporting-by-name");
  } else {
    reportByNameSheet.clear();
  }
  if (!reportByERSheet) {
    reportByERSheet = ss.insertSheet("reporting-by-er");
  } else {
    reportByERSheet.clear();
  }
  if (!reportByDateSheet) {
    reportByDateSheet = ss.insertSheet("reporting-by-date");
  } else {
    reportByDateSheet.clear();
  }
  if (!reportBySupporterSheet) {
    reportBySupporterSheet = ss.insertSheet("reporting-by-supporter-id");
  } else {
    reportBySupporterSheet.clear();
  }

  // Prompt for Start and End Dates
  var startDateInput = ui.prompt("Enter Start Date (YYYY-MM-DD) or leave blank for no limit").getResponseText().trim();
  var endDateInput = ui.prompt("Enter End Date (YYYY-MM-DD) or leave blank for no limit").getResponseText().trim();

  var startDate = startDateInput ? new Date(startDateInput) : null;
  var endDate = endDateInput ? new Date(endDateInput) : null;

  var data = exportSheet.getDataRange().getValues();
  var headers = data[0];
  var campaignCol = headers.indexOf("Campaign ID");
  var campaignDateCol = headers.indexOf("Campaign Date");
  var supporterCol = headers.indexOf("Supporter ID");

  var columns = {
    "External Reference 6": headers.indexOf("External Reference 6"),
    "External Reference 7": headers.indexOf("External Reference 7"),
    "External Reference 8": headers.indexOf("External Reference 8"),
    "External Reference 10": headers.indexOf("External Reference 10")
  };

  if (campaignCol === -1 || campaignDateCol === -1 || supporterCol === -1 || Object.values(columns).includes(-1)) {
    ui.alert("One or more required columns are missing.");
    return;
  }

  var campaignCounts = {};
  var results = {};
  var dateCounts = {};
  var supporterCounts = {}; // Stores counts of unique supporters

  for (var i = 1; i < data.length; i++) {
    var campaignID = data[i][campaignCol];
    var campaignDateRaw = data[i][campaignDateCol];
    var supporterID = data[i][supporterCol];
    var campaignDate = "";

    if (campaignDateRaw) {
      if (campaignDateRaw instanceof Date) {
        campaignDate = Utilities.formatDate(campaignDateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        campaignDate = campaignDateRaw.toString().trim();
      }
    } else {
      continue;
    }

    var campaignDateObj = new Date(campaignDate);
    if (isNaN(campaignDateObj)) continue;

    // Apply date filtering
    if ((startDate && campaignDateObj < startDate) || (endDate && campaignDateObj > endDate)) continue;

    // ✅ Only count Campaign ID if it meets the date filter
    if (campaignID) {
      campaignCounts[campaignID] = (campaignCounts[campaignID] || 0) + 1;
    }

    // ✅ Only count Supporter ID if it meets the date filter
    if (supporterID) {
      supporterCounts[supporterID] = (supporterCounts[supporterID] || 0) + 1;
    }

    // Group counts by month (YYYY-MM)
    var monthYear = Utilities.formatDate(campaignDateObj, Session.getScriptTimeZone(), "yyyy-MM");
    dateCounts[monthYear] = (dateCounts[monthYear] || 0) + 1;

    for (var key in columns) {
      var cellValue = data[i][columns[key]];
      if (cellValue && cellValue.includes(":")) {
        var parts = cellValue.split(":");
        var label = parts[0].trim();
        var values = parts[1].split(",").map(val => val.trim()).filter(val => val !== "");

        values.forEach(value => {
          if (key === "External Reference 10" && label === "YearType") {
            if (!isNaN(value)) {
              label = "Year";
            } else {
              label = "Type";
            }
          }
          if (value !== "") {
            var keyName = label + " - " + value;
            results[keyName] = (results[keyName] || 0) + 1;
          }
        });
      }
    }
  }

  // Write Campaign ID Summary (Filtered by Date) to "reporting-by-name"
  reportByNameSheet.appendRow(["Campaign ID", "Count"]);
  Object.entries(campaignCounts).sort().forEach(([campaign, count]) => {
    reportByNameSheet.appendRow([campaign, count]);
  });

  // Write External Reference Data to "reporting-by-er"
  reportByERSheet.appendRow(["Label", "Value", "Count"]);
  Object.entries(results).sort().forEach(([key, count]) => {
    var splitKey = key.split(" - ");
    if (splitKey.length === 2) {
      reportByERSheet.appendRow([splitKey[0], "'" + splitKey[1], count]); // Force text format
    }
  });

  // ✅ Write Date-based Aggregated Report to "reporting-by-date" (Force Text Formatting)
  reportByDateSheet.appendRow(["Month", "Count"]);
  Object.entries(dateCounts)
    .sort(([a], [b]) => a.localeCompare(b)) // Sort by YYYY-MM
    .forEach(([month, count]) => {
      reportByDateSheet.appendRow(["'" + month, "'" + count]); // Force text format to keep left alignment
    });

  // ✅ Write Supporter ID Summary (Filtered by Date) to "reporting-by-supporter-id"
  reportBySupporterSheet.appendRow(["Supporter ID", "Count"]);
  Object.entries(supporterCounts)
    .sort(([a], [b]) => a.localeCompare(b)) // Sort by ID
    .forEach(([supporterID, count]) => {
      reportBySupporterSheet.appendRow(["'" + supporterID, "'" + count]); // Force text format to keep left alignment
    });

  // ✅ Final success message only
  ui.alert("✅ Your UAN Reports have been updated!");
}