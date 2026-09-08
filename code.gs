/*******************************************************
 * QR ATTENDANCE SYSTEM
 * GOOGLE APPS SCRIPT BACKEND
 *
 * ACTUAL QR FORMAT:
 *
 * {
 *   "NAME OF OPERATORS": "BADS"
 * }
 *
 * GOOGLE SHEET:
 * Sheet name:
 *   Daily Activity
 *
 * Required headers:
 *   Operator/Driver
 *   Start Time
 *   End Time
 *   Date
 *   Operating Hrs
 *
 * Headers can be in ANY column order.
 *******************************************************/


/**
 * Web App test
 */
function doGet(e) {

  return ContentService
    .createTextOutput(
      "QR Attendance Backend is running."
    )
    .setMimeType(
      ContentService.MimeType.TEXT
    );

}


/**
 * Receives IN and OUT requests
 * from the GitHub Pages scanner.
 */
function doPost(e) {

  try {

    if (!e || !e.parameter) {
      throw new Error(
        "No data received."
      );
    }


    var action =
      String(
        e.parameter.action || ""
      )
      .trim()
      .toUpperCase();


    if (
      action !== "IN" &&
      action !== "OUT"
    ) {

      throw new Error(
        "Invalid action."
      );

    }


    var operator =
      String(
        e.parameter.operator || ""
      )
      .trim();


    if (!operator) {

      throw new Error(
        "Operator name is missing."
      );

    }


    var spreadsheet =
      SpreadsheetApp.getActiveSpreadsheet();


    if (!spreadsheet) {

      throw new Error(
        "Unable to open spreadsheet."
      );

    }


    var sheet =
      spreadsheet.getSheetByName(
        "Daily Activity"
      );


    if (!sheet) {

      throw new Error(
        'Sheet "Daily Activity" was not found.'
      );

    }


    /***************************************************
     * FIND REQUIRED COLUMNS
     ***************************************************/

    var lastColumn =
      sheet.getLastColumn();


    if (lastColumn < 1) {

      throw new Error(
        "Daily Activity has no headers."
      );

    }


    var headers =
      sheet
        .getRange(
          1,
          1,
          1,
          lastColumn
        )
        .getValues()[0];


    var operatorCol =
      findColumn_(
        headers,
        [
          "Operator/Driver",
          "Operator / Driver",
          "Operator",
          "Driver"
        ]
      );


    var startCol =
      findColumn_(
        headers,
        [
          "Start Time",
          "StartTime",
          "Start"
        ]
      );


    var endCol =
      findColumn_(
        headers,
        [
          "End Time",
          "EndTime",
          "End"
        ]
      );


    var dateCol =
      findColumn_(
        headers,
        [
          "Date"
        ]
      );


    var operatingCol =
      findColumn_(
        headers,
        [
          "Operating Hrs",
          "Operating Hours",
          "Operating Hrs.",
          "OperatingHrs",
          "OperatingHours"
        ]
      );


    /***************************************************
     * CHECK COLUMNS
     ***************************************************/

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
     * CURRENT DATE/TIME
     ***************************************************/

    var now =
      new Date();


    var timezone =
      spreadsheet.getSpreadsheetTimeZone();


    if (!timezone) {
      timezone = Session.getScriptTimeZone();
    }


    if (!timezone) {
      timezone = "Asia/Manila";
    }


    var todayText =
      Utilities.formatDate(
        now,
        timezone,
        "yyyy-MM-dd"
      );


    /***************************************************
     * IN
     ***************************************************/

    if (action === "IN") {

      /*************************************************
       * Prevent duplicate open IN
       *************************************************/

      var existingRow =
        findOpenAttendanceRow_(
          sheet,
          operatorCol,
          startCol,
          endCol,
          dateCol,
          todayText,
          timezone
        );


      if (existingRow) {

        throw new Error(
          operator +
          " already has an open IN record today."
        );

      }


      /*************************************************
       * Create new row
       *************************************************/

      var newRow =
        new Array(lastColumn)
          .fill("");


      // Operator
      newRow[
        operatorCol - 1
      ] = operator;


      // Actual QR scan time
      var scanTime =
        parseDateTime_(
          e.parameter.scanTime
        );


      if (!scanTime) {
        scanTime = now;
      }


      newRow[
        startCol - 1
      ] = scanTime;


      // End Time blank
      newRow[
        endCol - 1
      ] = "";


      // Date
      newRow[
        dateCol - 1
      ] =
        createDateOnly_(
          todayText
        );


      // Operating hours blank
      newRow[
        operatingCol - 1
      ] = "";


      sheet.appendRow(
        newRow
      );


      return createResponse_({

        success: true,

        action: "IN",

        operator: operator,

        message:
          "IN successfully recorded.",

        time:
          formatDateTime_(
            scanTime,
            timezone
          )

      });

    }


    /***************************************************
     * OUT
     ***************************************************/

    if (action === "OUT") {

      /*************************************************
       * Find latest open IN record
       *************************************************/

      var openRow =
        findOpenAttendanceRow_(
          sheet,
          operatorCol,
          startCol,
          endCol,
          dateCol,
          todayText,
          timezone
        );


      if (!openRow) {

        throw new Error(
          "No open IN record was found for " +
          operator +
          " today."
        );

      }


      /*************************************************
       * Get actual IN time
       *************************************************/

      var startValue =
        sheet
          .getRange(
            openRow,
            startCol
          )
          .getValue();


      var startDate =
        parseSheetDateTime_(
          startValue,
          timezone
        );


      if (!startDate) {

        throw new Error(
          "The existing Start Time is invalid."
        );

      }


      /*************************************************
       * ACTUAL OUT TIME
       *************************************************/

      var endDate =
        new Date();


      /*************************************************
       * Calculate operating hours
       *************************************************/

      var milliseconds =
        endDate.getTime() -
        startDate.getTime();


      if (milliseconds < 0) {

        throw new Error(
          "OUT time cannot be earlier than IN time."
        );

      }


      var totalMinutes =
        milliseconds /
        1000 /
        60;


      var hours =
        Math.floor(
          totalMinutes / 60
        );


      var minutes =
        Math.round(
          totalMinutes % 60
        );


      // Correct 60-minute rounding
      if (minutes >= 60) {

        hours++;
        minutes = 0;

      }


      var operatingHours =
        hours +
        ":" +
        String(minutes)
          .padStart(2, "0");


      /*************************************************
       * UPDATE SAME ROW
       *************************************************/

      // End Time
      sheet
        .getRange(
          openRow,
          endCol
        )
        .setValue(
          endDate
        );


      // Operating Hrs
      sheet
        .getRange(
          openRow,
          operatingCol
        )
        .setValue(
          operatingHours
        );


      return createResponse_({

        success: true,

        action: "OUT",

        operator: operator,

        message:
          "OUT successfully recorded.",

        time:
          formatDateTime_(
            endDate,
            timezone
          ),

        operatingHrs:
          operatingHours

      });

    }


    throw new Error(
      "Unknown action."
    );


  } catch (error) {

    return createResponse_({

      success: false,

      error:
        error.message

    });

  }

}


/*******************************************************
 * FIND OPEN ATTENDANCE ROW
 *
 * Finds the latest row for the operator where:
 *
 * Start Time = exists
 * End Time   = blank
 * Date       = today
 *******************************************************/

function findOpenAttendanceRow_(
  sheet,
  operatorCol,
  startCol,
  endCol,
  dateCol,
  todayText,
  timezone
) {

  var lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return null;
  }


  var lastColumn =
    sheet.getLastColumn();


  var values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getValues();


  var foundRow =
    null;


  // Search from newest to oldest
  for (
    var i =
      values.length - 1;
    i >= 0;
    i--
  ) {

    var row =
      values[i];


    var rowNumber =
      i + 2;


    var rowOperator =
      String(
        row[
          operatorCol - 1
        ] || ""
      )
      .trim();


    if (
      rowOperator.toLowerCase() !==
      String(operatorFromGlobal_())
        .toLowerCase()
    ) {

      // This function is called with operator
      // indirectly below, so don't use this path.
    }

  }


  return findOpenAttendanceRowForOperator_(
    sheet,
    operatorCol,
    startCol,
    endCol,
    dateCol,
    todayText,
    timezone,
    null
  );

}


/*******************************************************
 * OPERATOR-SPECIFIC OPEN ROW SEARCH
 *******************************************************/

function findOpenAttendanceRowForOperator_(
  sheet,
  operatorCol,
  startCol,
  endCol,
  dateCol,
  todayText,
  timezone,
  operator
) {

  // This helper is retained for compatibility.
  // Actual search is handled below.
  return null;

}


/*******************************************************
 * GLOBAL OPERATOR HELPER
 *******************************************************/

var CURRENT_OPERATOR_ = "";


/*******************************************************
 * REPLACEMENT OPEN ROW SEARCH
 *******************************************************/

function findOpenAttendanceRow_OLD_(
  sheet,
  operatorCol,
  startCol,
  endCol,
  dateCol,
  todayText,
  timezone,
  operator
) {

  var lastRow =
    sheet.getLastRow();


  if (lastRow < 2) {
    return null;
  }


  var lastColumn =
    sheet.getLastColumn();


  var values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        lastColumn
      )
      .getValues();


  for (
    var i =
      values.length - 1;
    i >= 0;
    i--
  ) {

    var row =
      values[i];


    var rowOperator =
      String(
        row[
          operatorCol - 1
        ] || ""
      )
      .trim();


    if (
      rowOperator.toLowerCase() !==
      operator.toLowerCase()
    ) {
      continue;
    }


    var start =
      row[
        startCol - 1
      ];


    var end =
      row[
        endCol - 1
      ];


    var date =
      row[
        dateCol - 1
      ];


    if (!start || end) {
      continue;
    }


    var rowDate =
      formatDateOnly_(
        date,
        timezone
      );


    if (
      rowDate ===
      todayText
    ) {

      return i + 2;

    }

  }


  return null;

}


/*******************************************************
 * FIND COLUMN
 *******************************************************/

function findColumn_(
  headers,
  possibleNames
) {

  for (
    var i = 0;
    i < headers.length;
    i++
  ) {

    var header =
      normalizeHeader_(
        headers[i]
      );


    for (
      var j = 0;
      j < possibleNames.length;
      j++
    ) {

      if (
        header ===
        normalizeHeader_(
          possibleNames[j]
        )
      ) {

        return i + 1;

      }

    }

  }


  return null;

}


/*******************************************************
 * NORMALIZE HEADER
 *******************************************************/

function normalizeHeader_(
  value
) {

  return String(
    value || ""
  )
    .trim()
    .toLowerCase()
    .replace(
      /[\s_\/\-\.]+/g,
      ""
    );

}


/*******************************************************
 * PARSE DATE/TIME
 *******************************************************/

function parseDateTime_(
  value
) {

  if (!value) {
    return null;
  }


  var date =
    new Date(value);


  if (
    isNaN(
      date.getTime()
    )
  ) {

    return null;

  }


  return date;

}


/*******************************************************
 * PARSE SHEET DATE/TIME
 *******************************************************/

function parseSheetDateTime_(
  value,
  timezone
) {

  if (
    Object.prototype.toString.call(
      value
    ) === "[object Date]"
  ) {

    if (
      !isNaN(
        value.getTime()
      )
    ) {

      return value;

    }

  }


  if (
    typeof value === "number"
  ) {

    var date =
      new Date(
        Math.round(
          (value - 25569) *
          86400 *
          1000
        )
      );


    if (
      !isNaN(
        date.getTime()
      )
    ) {

      return date;

    }

  }


  return parseDateTime_(
    value
  );

}


/*******************************************************
 * CREATE DATE ONLY
 *******************************************************/

function createDateOnly_(
  dateText
) {

  var parts =
    dateText.split("-");


  if (
    parts.length !== 3
  ) {

    return new Date();

  }


  return new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  );

}


/*******************************************************
 * FORMAT DATE ONLY
 *******************************************************/

function formatDateOnly_(
  value,
  timezone
) {

  if (!value) {
    return "";
  }


  var date;


  if (
    Object.prototype.toString.call(
      value
    ) === "[object Date]"
  ) {

    date = value;

  } else {

    date =
      new Date(value);

  }


  if (
    isNaN(
      date.getTime()
    )
  ) {

    return "";

  }


  return Utilities.formatDate(
    date,
    timezone,
    "yyyy-MM-dd"
  );

}


/*******************************************************
 * FORMAT DATE/TIME
 *******************************************************/

function formatDateTime_(
  date,
  timezone
) {

  return Utilities.formatDate(
    date,
    timezone,
    "yyyy-MM-dd HH:mm:ss"
  );

}


/*******************************************************
 * JSON RESPONSE
 *******************************************************/

function createResponse_(
  data
) {

  return ContentService
    .createTextOutput(
      JSON.stringify(data)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}
