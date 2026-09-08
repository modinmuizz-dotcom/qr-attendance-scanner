/*******************************************************
 * AMANAH CONSTRUCTION SERVICES
 * QR ATTENDANCE SYSTEM
 *
 * QR FORMAT:
 *
 * {
 *   "NAME OF OPERATORS": "BADS"
 * }
 *
 * GOOGLE SHEET:
 * Daily Activity
 *
 * HEADER ROW:
 * Automatically detected
 *
 * REQUIRED HEADERS:
 * Date
 * Operator/Driver
 * Start Time
 * End Time
 * Operating Hrs
 *******************************************************/


function doGet(e) {

  return ContentService
    .createTextOutput(
      JSON.stringify({
        success: true,
        message: "QR Attendance Backend is working."
      })
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}


/*******************************************************
 * RECEIVE IN / OUT FROM GITHUB SCANNER
 *******************************************************/

function doPost(e) {

  try {

    if (!e || !e.parameter) {

      throw new Error(
        "No data was received from the scanner."
      );

    }


    var action =
      String(
        e.parameter.action || ""
      )
      .trim()
      .toUpperCase();


    var operator =
      String(
        e.parameter.operator || ""
      )
      .trim();


    var scanTime =
      String(
        e.parameter.scanTime || ""
      )
      .trim();


    /***************************************************
     * VALIDATE ACTION
     ***************************************************/

    if (
      action !== "IN" &&
      action !== "OUT"
    ) {

      throw new Error(
        "Invalid action: " + action
      );

    }


    /***************************************************
     * VALIDATE OPERATOR
     ***************************************************/

    if (!operator) {

      throw new Error(
        "Operator name is missing."
      );

    }


    /***************************************************
     * GET SPREADSHEET
     ***************************************************/

    var spreadsheet =
      SpreadsheetApp
        .getActiveSpreadsheet();


    if (!spreadsheet) {

      throw new Error(
        "Unable to access the Google Spreadsheet."
      );

    }


    /***************************************************
     * GET DAILY ACTIVITY SHEET
     ***************************************************/

    var sheet =
      spreadsheet
        .getSheetByName(
          "Daily Activity"
        );


    if (!sheet) {

      throw new Error(
        'Sheet "Daily Activity" was not found.'
      );

    }


    /***************************************************
     * FIND HEADER ROW + COLUMNS
     ***************************************************/

    var headerInfo =
      findHeaders_(
        sheet
      );


    if (!headerInfo) {

      throw new Error(
        "Required headers were not found."
      );

    }


    var headerRow =
      headerInfo.row;


    var cols =
      headerInfo.columns;


    /***************************************************
     * TIMEZONE
     ***************************************************/

    var timezone =
      spreadsheet
        .getSpreadsheetTimeZone();


    if (!timezone) {

      timezone =
        "Asia/Manila";

    }


    var now =
      new Date();


    /***************************************************
     * =================================================
     *                    IN
     * =================================================
     ***************************************************/

    if (
      action === "IN"
    ) {

      /*************************************************
       * CHECK FOR EXISTING OPEN IN
       *************************************************/

      var existingOpenRow =
        findOpenRow_(
          sheet,
          headerRow,
          cols,
          operator
        );


      if (existingOpenRow) {

        throw new Error(
          operator +
          " already has an open IN record in row " +
          existingOpenRow +
          ". Please OUT first."
        );

      }


      /*************************************************
       * USE QR SCAN TIME
       *************************************************/

      var startDate =
        parseDateTime_(
          scanTime
        );


      if (!startDate) {

        startDate =
          now;

      }


      /*************************************************
       * NEXT ROW
       *************************************************/

      var newRow =
        Math.max(
          sheet.getLastRow() + 1,
          headerRow + 1
        );


      /*************************************************
       * OPERATOR
       *************************************************/

      sheet
        .getRange(
          newRow,
          cols.operator
        )
        .setValue(
          operator
        );


      /*************************************************
       * START TIME
       *************************************************/

      sheet
        .getRange(
          newRow,
          cols.start
        )
        .setValue(
          startDate
        );


      sheet
        .getRange(
          newRow,
          cols.start
        )
        .setNumberFormat(
          "h:mm:ss AM/PM"
        );


      /*************************************************
       * END TIME
       *************************************************/

      sheet
        .getRange(
          newRow,
          cols.end
        )
        .clearContent();


      /*************************************************
       * DATE
       *************************************************/

      sheet
        .getRange(
          newRow,
          cols.date
        )
        .setValue(
          startDate
        );


      sheet
        .getRange(
          newRow,
          cols.date
        )
        .setNumberFormat(
          "mm/dd/yyyy"
        );


      /*************************************************
       * OPERATING HOURS
       *************************************************/

      sheet
        .getRange(
          newRow,
          cols.operating
        )
        .clearContent();


      /*************************************************
       * FORCE SAVE
       *************************************************/

      SpreadsheetApp
        .flush();


      /*************************************************
       * RESPONSE
       *************************************************/

      return response_({

        success: true,

        action: "IN",

        operator: operator,

        row: newRow,

        time:
          Utilities.formatDate(
            startDate,
            timezone,
            "h:mm:ss a"
          ),

        message:
          "IN successfully recorded for " +
          operator

      });

    }


    /***************************************************
     * =================================================
     *                    OUT
     * =================================================
     ***************************************************/

    if (
      action === "OUT"
    ) {

      /*************************************************
       * IMPORTANT:
       *
       * Find the latest OPEN IN for this operator.
       *
       * We DO NOT restrict this to today's date.
       *
       * This fixes the problem where the IN row has
       * a different Date value.
       *************************************************/

      var openRow =
        findOpenRow_(
          sheet,
          headerRow,
          cols,
          operator
        );


      if (!openRow) {

        throw new Error(
          "No open IN record was found for " +
          operator +
          "."
        );

      }


      /*************************************************
       * GET START TIME
       *************************************************/

      var startValue =
        sheet
          .getRange(
            openRow,
            cols.start
          )
          .getValue();


      var startDate =
        parseSheetDate_(
          startValue
        );


      if (!startDate) {

        throw new Error(
          "The Start Time in row " +
          openRow +
          " is invalid."
        );

      }


      /*************************************************
       * ACTUAL OUT TIME
       *************************************************/

      var endDate =
        new Date();


      /*************************************************
       * CALCULATE DIFFERENCE
       *************************************************/

      var milliseconds =
        endDate.getTime() -
        startDate.getTime();


      if (
        milliseconds < 0
      ) {

        throw new Error(
          "OUT time cannot be earlier than IN time."
        );

      }


      /*************************************************
       * CONVERT TO GOOGLE SHEETS DURATION
       *
       * Google Sheets stores duration as a fraction
       * of one day.
       *************************************************/

      var duration =
        milliseconds /
        86400000;


      /*************************************************
       * WRITE END TIME
       *************************************************/

      sheet
        .getRange(
          openRow,
          cols.end
        )
        .setValue(
          endDate
        );


      sheet
        .getRange(
          openRow,
          cols.end
        )
        .setNumberFormat(
          "h:mm:ss AM/PM"
        );


      /*************************************************
       * WRITE OPERATING HOURS
       *************************************************/

      sheet
        .getRange(
          openRow,
          cols.operating
        )
        .setValue(
          duration
        );


      sheet
        .getRange(
          openRow,
          cols.operating
        )
        .setNumberFormat(
          "[h]:mm"
        );


      /*************************************************
       * FORCE SAVE
       *************************************************/

      SpreadsheetApp
        .flush();


      /*************************************************
       * RESPONSE
       *************************************************/

      return response_({

        success: true,

        action: "OUT",

        operator: operator,

        row: openRow,

        startTime:
          Utilities.formatDate(
            startDate,
            timezone,
            "h:mm:ss a"
          ),

        endTime:
          Utilities.formatDate(
            endDate,
            timezone,
            "h:mm:ss a"
          ),

        operatingHrs:
          formatDuration_(
            milliseconds
          ),

        message:
          "OUT successfully recorded for " +
          operator

      });

    }


    throw new Error(
      "Unknown action."
    );


  }

  catch (error) {

    console.error(
      error
    );


    return response_({

      success: false,

      error:
        error.message ||
        String(error)

    });

  }

}


/*******************************************************
 * FIND HEADER ROW
 *
 * Searches rows 1 through 10.
 *******************************************************/

function findHeaders_(
  sheet
) {

  var maxRows =
    Math.min(
      10,
      sheet.getMaxRows()
    );


  var maxColumns =
    sheet.getLastColumn();


  if (
    maxRows < 1 ||
    maxColumns < 1
  ) {

    return null;

  }


  var values =
    sheet
      .getRange(
        1,
        1,
        maxRows,
        maxColumns
      )
      .getDisplayValues();


  for (
    var r = 0;
    r < values.length;
    r++
  ) {

    var row =
      values[r];


    var columns = {

      date:
        findColumnInRow_(
          row,
          [
            "Date"
          ]
        ),

      operator:
        findColumnInRow_(
          row,
          [
            "Operator/Driver",
            "Operator / Driver",
            "Operator",
            "Driver"
          ]
        ),

      start:
        findColumnInRow_(
          row,
          [
            "Start Time",
            "StartTime",
            "Start"
          ]
        ),

      end:
        findColumnInRow_(
          row,
          [
            "End Time",
            "EndTime",
            "End"
          ]
        ),

      operating:
        findColumnInRow_(
          row,
          [
            "Operating Hrs",
            "Operating Hours",
            "Operating Hrs.",
            "OperatingHrs"
          ]
        )

    };


    if (
      columns.date &&
      columns.operator &&
      columns.start &&
      columns.end &&
      columns.operating
    ) {

      return {

        row:
          r + 1,

        columns:
          columns

      };

    }

  }


  return null;

}


/*******************************************************
 * FIND COLUMN
 *******************************************************/

function findColumnInRow_(
  row,
  names
) {

  for (
    var i = 0;
    i < row.length;
    i++
  ) {

    var current =
      normalizeHeader_(
        row[i]
      );


    for (
      var j = 0;
      j < names.length;
      j++
    ) {

      if (
        current ===
        normalizeHeader_(
          names[j]
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
 * FIND OPEN IN ROW
 *
 * IMPORTANT:
 * Does NOT require today's date.
 *
 * Finds the latest row for this operator where:
 *
 * Start Time = has value
 * End Time   = empty
 *******************************************************/

function findOpenRow_(
  sheet,
  headerRow,
  cols,
  operator
) {

  var lastRow =
    sheet.getLastRow();


  if (
    lastRow <= headerRow
  ) {

    return null;

  }


  var lastColumn =
    sheet.getLastColumn();


  var values =
    sheet
      .getRange(
        headerRow + 1,
        1,
        lastRow - headerRow,
        lastColumn
      )
      .getValues();


  /*
   * Search from bottom to top.
   *
   * Therefore the newest open IN is selected.
   */

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
          cols.operator - 1
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
        cols.start - 1
      ];


    var end =
      row[
        cols.end - 1
      ];


    /*
     * OPEN RECORD:
     *
     * Start exists
     * End is empty
     */

    if (
      start &&
      !end
    ) {

      return (
        headerRow +
        1 +
        i
      );

    }

  }


  return null;

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
    new Date(
      value
    );


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
 * PARSE SHEET DATE
 *******************************************************/

function parseSheetDate_(
  value
) {

  if (
    Object.prototype.toString.call(
      value
    ) ===
    "[object Date]"
  ) {

    if (
      !isNaN(
        value.getTime()
      )
    ) {

      return value;

    }

  }


  return parseDateTime_(
    value
  );

}


/*******************************************************
 * FORMAT DURATION
 *******************************************************/

function formatDuration_(
  milliseconds
) {

  var totalMinutes =
    Math.round(
      milliseconds /
      60000
    );


  var hours =
    Math.floor(
      totalMinutes /
      60
    );


  var minutes =
    totalMinutes %
    60;


  return (
    hours +
    ":" +
    String(
      minutes
    ).padStart(
      2,
      "0"
    )
  );

}


/*******************************************************
 * JSON RESPONSE
 *******************************************************/

function response_(
  data
) {

  return ContentService
    .createTextOutput(
      JSON.stringify(
        data
      )
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}
