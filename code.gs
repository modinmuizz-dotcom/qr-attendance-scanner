/*******************************************************
 * QR ATTENDANCE BACKEND
 * Google Apps Script + Google Sheets
 *
 * Sheet:
 *   Daily Activity
 *
 * Required columns:
 *   Operator/Driver
 *   Start Time
 *   End Time
 *   Date
 *   Operating Hrs
 *******************************************************/


/**
 * Health check
 */
function doGet(e) {
  return ContentService
    .createTextOutput("QR Attendance Backend is running.")
    .setMimeType(ContentService.MimeType.TEXT);
}


/**
 * Receives IN / OUT data from GitHub Pages.
 *
 * The GitHub page sends a normal HTML form POST.
 * This avoids browser CORS problems.
 */
function doPost(e) {

  try {

    if (!e || !e.parameter) {
      throw new Error("No data received.");
    }

    var action = String(e.parameter.action || "").trim().toUpperCase();

    if (action !== "IN" && action !== "OUT") {
      throw new Error("Invalid action. Use IN or OUT.");
    }


    // QR information
    var operator = String(e.parameter.operator || "").trim();
    var qrStartTime = String(e.parameter.startTime || "").trim();
    var qrEndTime = String(e.parameter.endTime || "").trim();
    var operatingHrs = String(e.parameter.operatingHrs || "").trim();
    var qrDate = String(e.parameter.date || "").trim();

    // Time when QR was scanned
    var scanTime = String(e.parameter.scanTime || "").trim();


    // Validate all required QR fields
    if (!operator) {
      throw new Error("Missing Operator/Driver.");
    }

    if (!qrStartTime) {
      throw new Error("Missing Start Time.");
    }

    if (!qrEndTime) {
      throw new Error("Missing End Time.");
    }

    if (!operatingHrs) {
      throw new Error("Missing Operating Hrs.");
    }

    if (!qrDate) {
      throw new Error("Missing Date.");
    }


    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (!spreadsheet) {
      throw new Error("Unable to open the spreadsheet.");
    }


    var sheet = spreadsheet.getSheetByName("Daily Activity");

    if (!sheet) {
      throw new Error(
        'Sheet "Daily Activity" was not found.'
      );
    }


    // Get headers
    var lastColumn = sheet.getLastColumn();

    if (lastColumn === 0) {
      throw new Error("Daily Activity has no headers.");
    }

    var headers = sheet
      .getRange(1, 1, 1, lastColumn)
      .getValues()[0];


    // Dynamically find columns
    var operatorCol = findColumn_(headers, [
      "Operator/Driver",
      "Operator / Driver"
    ]);

    var startCol = findColumn_(headers, [
      "Start Time",
      "StartTime"
    ]);

    var endCol = findColumn_(headers, [
      "End Time",
      "EndTime"
    ]);

    var dateCol = findColumn_(headers, [
      "Date"
    ]);

    var operatingCol = findColumn_(headers, [
      "Operating Hrs",
      "Operating Hours",
      "Operating Hrs."
    ]);


    // Make sure all required columns exist
    if (!operatorCol) {
      throw new Error(
        'Column "Operator/Driver" was not found.'
      );
    }

    if (!startCol) {
      throw new Error(
        'Column "Start Time" was not found.'
      );
    }

    if (!endCol) {
      throw new Error(
        'Column "End Time" was not found.'
      );
    }

    if (!dateCol) {
      throw new Error(
        'Column "Date" was not found.'
      );
    }

    if (!operatingCol) {
      throw new Error(
        'Column "Operating Hrs" was not found.'
      );
    }


    /***************************************************
     * BUILD NEW ROW
     ***************************************************/

    var newRow = new Array(lastColumn).fill("");


    // Operator
    newRow[operatorCol - 1] = operator;


    // Date
    var parsedDate = parseDate_(qrDate);

    if (parsedDate) {
      newRow[dateCol - 1] = parsedDate;
    } else {
      newRow[dateCol - 1] = qrDate;
    }


    // Operating hours
    newRow[operatingCol - 1] = operatingHrs;


    /***************************************************
     * IN
     *
     * Start Time = actual time QR was scanned
     * End Time   = blank
     ***************************************************/

    if (action === "IN") {

      var actualScanTime = parseDateTime_(scanTime);

      if (!actualScanTime) {
        actualScanTime = new Date();
      }

      newRow[startCol - 1] = actualScanTime;
      newRow[endCol - 1] = "";

    }


    /***************************************************
     * OUT
     *
     * Start Time = QR Start Time
     * End Time   = actual OUT button time
     ***************************************************/

    if (action === "OUT") {

      var parsedStartTime = parseDateTime_(qrStartTime);

      if (parsedStartTime) {
        newRow[startCol - 1] = parsedStartTime;
      } else {
        newRow[startCol - 1] = qrStartTime;
      }

      // Actual OUT time
      newRow[endCol - 1] = new Date();

    }


    // Append
    sheet.appendRow(newRow);


    return createResponse_({
      success: true,
      action: action,
      operator: operator,
      message:
        action === "IN"
          ? "IN successfully recorded."
          : "OUT successfully recorded."
    });

  } catch (error) {

    return createResponse_({
      success: false,
      error: error.message
    });
  }
}


/**
 * Dynamically locate a column by header name.
 */
function findColumn_(headers, possibleNames) {

  for (var i = 0; i < headers.length; i++) {

    var header = String(headers[i] || "")
      .trim()
      .toLowerCase();

    for (var j = 0; j < possibleNames.length; j++) {

      var possible = String(possibleNames[j])
        .trim()
        .toLowerCase();

      if (header === possible) {
        return i + 1;
      }
    }
  }

  return null;
}


/**
 * Parse a date.
 */
function parseDate_(value) {

  if (!value) {
    return null;
  }

  var date = new Date(value);

  if (isNaN(date.getTime())) {
    return null;
  }

  return date;
}


/**
 * Parse date/time.
 */
function parseDateTime_(value) {

  if (!value) {
    return null;
  }

  var date = new Date(value);

  if (isNaN(date.getTime())) {
    return null;
  }

  return date;
}


/**
 * Return JSON response.
 */
function createResponse_(data) {

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
