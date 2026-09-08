/************************************************************
 * AMANAH CONSTRUCTION SERVICES
 * QR ATTENDANCE SYSTEM
 *
 * GOOGLE SHEET:
 * Daily Activity
 *
 * TIMEZONE:
 * Asia/Manila
 *
 * QR FORMAT:
 *
 * {
 *   "NAME OF OPERATORS": "BADS"
 * }
 *
 * REQUIRED SHEET HEADERS:
 *
 * Date
 * Operator/Driver
 * Start Time
 * End Time
 * Operating Hrs
 ************************************************************/


const SHEET_NAME = "Daily Activity";
const TIMEZONE = "Asia/Manila";


/************************************************************
 * WEB APP TEST
 ************************************************************/

function doGet(e) {

  return ContentService
    .createTextOutput(
      JSON.stringify({
        success: true,
        message: "QR Attendance Backend is working.",
        timezone: TIMEZONE
      })
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );

}


/************************************************************
 * RECEIVE DATA FROM GITHUB SCANNER
 ************************************************************/

function doPost(e) {

  try {

    /******************************************************
     * CHECK REQUEST
     ******************************************************/

    if (!e || !e.parameter) {

      throw new Error(
        "No data was received from the scanner."
      );

    }


    /******************************************************
     * GET DATA
     ******************************************************/

    const action =
      String(
        e.parameter.action || ""
      )
      .trim()
      .toUpperCase();


    const operator =
      String(
        e.parameter.operator || ""
      )
      .trim();


    const scanTime =
      String(
        e.parameter.scanTime || ""
      )
      .trim();


    /******************************************************
     * VALIDATE ACTION
     ******************************************************/

    if (
      action !== "IN" &&
      action !== "OUT"
    ) {

      throw new Error(
        "Invalid action: " + action
      );

    }


    /******************************************************
     * VALIDATE OPERATOR
     ******************************************************/

    if (!operator) {

      throw new Error(
        "Operator name is missing."
      );

    }


    /******************************************************
     * GET SPREADSHEET
     ******************************************************/

    const spreadsheet =
      SpreadsheetApp.getActiveSpreadsheet();


    if (!spreadsheet) {

      throw new Error(
        "Unable to access the Google Spreadsheet."
      );

    }


    /******************************************************
     * FORCE PHILIPPINE TIMEZONE
     ******************************************************/

    spreadsheet.setSpreadsheetTimeZone(
      TIMEZONE
    );


    /******************************************************
     * GET DAILY ACTIVITY SHEET
     ******************************************************/

    const sheet =
      spreadsheet.getSheetByName(
        SHEET_NAME
      );


    if (!sheet) {

      throw new Error(
        'Sheet "' +
        SHEET_NAME +
        '" was not found.'
      );

    }


    /******************************************************
     * FIND REQUIRED COLUMNS
     ******************************************************/

    const headerInfo =
      findHeaders_(
        sheet
      );


    if (!headerInfo) {

      throw new Error(
        "Required columns were not found. " +
        "Please check Date, Operator/Driver, Start Time, " +
        "End Time and Operating Hrs."
      );

    }


    const headerRow =
      headerInfo.row;


    const cols =
      headerInfo.columns;


    /******************************************************
     * SERVER TIME
     ******************************************************/

    const serverNow =
      new Date();


    /******************************************************
     * ====================================================
     *                         IN
     * ====================================================
     ******************************************************/

    if (
      action === "IN"
    ) {

      /****************************************************
       * CHECK FOR EXISTING OPEN IN
       ****************************************************/

      const existingRow =
        findOpenRow_(
          sheet,
          headerRow,
          cols,
          operator
        );


      if (existingRow) {

        throw new Error(
          operator +
          " already has an open IN record in row " +
          existingRow +
          ". Please OUT first."
        );

      }


      /****************************************************
       * GET QR SCAN TIME
       *
       * index.html sends an ISO timestamp when the
       * QR scan succeeds.
       ****************************************************/

      let startDate =
        parseDateTime_(
          scanTime
        );


      /****************************************************
       * FALLBACK TO SERVER TIME
       ****************************************************/

      if (!startDate) {

        startDate =
          serverNow;

      }


      /****************************************************
       * CREATE NEW ROW
       ****************************************************/

      const newRow =
        Math.max(
          sheet.getLastRow() + 1,
          headerRow + 1
        );


      /****************************************************
       * OPERATOR / DRIVER
       ****************************************************/

      sheet
        .getRange(
          newRow,
          cols.operator
        )
        .setValue(
          operator
        );


      /****************************************************
       * DATE
       ****************************************************/

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
          "MM/dd/yyyy"
        );


      /****************************************************
       * START TIME
       ****************************************************/

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


      /****************************************************
       * END TIME
       ****************************************************/

      sheet
        .getRange(
          newRow,
          cols.end
        )
        .clearContent();


      /****************************************************
       * OPERATING HOURS
       ****************************************************/

      sheet
        .getRange(
          newRow,
          cols.operating
        )
        .clearContent();


      /****************************************************
       * SAVE
       ****************************************************/

      SpreadsheetApp.flush();


      /****************************************************
       * RESPONSE
       ****************************************************/

      return response_({

        success: true,

        action: "IN",

        operator: operator,

        row: newRow,

        timezone: TIMEZONE,

        date:
          Utilities.formatDate(
            startDate,
            TIMEZONE,
            "MM/dd/yyyy"
          ),

        startTime:
          Utilities.formatDate(
            startDate,
            TIMEZONE,
            "h:mm:ss a"
          ),

        message:
          "IN successfully recorded for " +
          operator

      });

    }


    /******************************************************
     * ====================================================
     *                         OUT
     * ====================================================
     ******************************************************/

    if (
      action === "OUT"
    ) {

      /****************************************************
       * FIND OPEN IN
       *
       * Searches the entire sheet.
       *
       * It does NOT require the Date to be today.
       ****************************************************/

      const openRow =
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


      /****************************************************
       * GET START TIME
       ****************************************************/

      const startValue =
        sheet
          .getRange(
            openRow,
            cols.start
          )
          .getValue();


      const startDate =
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


      /****************************************************
       * ACTUAL OUT TIME
       *
       * This is the Apps Script server timestamp.
       ****************************************************/

      const endDate =
        serverNow;


      /****************************************************
       * CALCULATE OPERATING TIME
       ****************************************************/

      const milliseconds =
        endDate.getTime() -
        startDate.getTime();


      if (
        milliseconds < 0
      ) {

        throw new Error(
          "OUT time cannot be earlier than IN time."
        );

      }


      /****************************************************
       * GOOGLE SHEETS DURATION
       *
       * One day = 86,400,000 milliseconds.
       ****************************************************/

      const duration =
        milliseconds /
        86400000;


      /****************************************************
       * WRITE END TIME
       ****************************************************/

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


      /****************************************************
       * WRITE OPERATING HOURS
       *
       * IMPORTANT:
       *
       * [h]:mm:ss displays the EXACT duration.
       *
       * Example:
       *
       * 2 minutes 57 seconds
       * = 0:02:57
       ****************************************************/

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
          "[h]:mm:ss"
        );


      /****************************************************
       * SAVE
       ****************************************************/

      SpreadsheetApp.flush();


      /****************************************************
       * RESPONSE
       ****************************************************/

      return response_({

        success: true,

        action: "OUT",

        operator: operator,

        row: openRow,

        timezone: TIMEZONE,

        date:
          Utilities.formatDate(
            startDate,
            TIMEZONE,
            "MM/dd/yyyy"
          ),

        startTime:
          Utilities.formatDate(
            startDate,
            TIMEZONE,
            "h:mm:ss a"
          ),

        endTime:
          Utilities.formatDate(
            endDate,
            TIMEZONE,
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


/************************************************************
 * FIND HEADER ROW AND COLUMNS
 ************************************************************/

function findHeaders_(
  sheet
) {

  const maxRows =
    Math.min(
      10,
      sheet.getMaxRows()
    );


  const maxColumns =
    sheet.getLastColumn();


  if (
    maxRows < 1 ||
    maxColumns < 1
  ) {

    return null;

  }


  const values =
    sheet
      .getRange(
        1,
        1,
        maxRows,
        maxColumns
      )
      .getDisplayValues();


  /******************************************************
   * SEARCH FIRST 10 ROWS
   ******************************************************/

  for (
    let r = 0;
    r < values.length;
    r++
  ) {

    const row =
      values[r];


    const columns = {

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


/************************************************************
 * FIND COLUMN IN HEADER ROW
 ************************************************************/

function findColumnInRow_(
  row,
  names
) {

  for (
    let i = 0;
    i < row.length;
    i++
  ) {

    const current =
      normalizeHeader_(
        row[i]
      );


    for (
      let j = 0;
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


/************************************************************
 * NORMALIZE HEADER
 ************************************************************/

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


/************************************************************
 * FIND OPEN IN RECORD
 *
 * Finds the newest row for this operator where:
 *
 * Start Time = NOT EMPTY
 * End Time   = EMPTY
 *
 * IMPORTANT:
 * Date is NOT checked.
 ************************************************************/

function findOpenRow_(
  sheet,
  headerRow,
  cols,
  operator
) {

  const lastRow =
    sheet.getLastRow();


  if (
    lastRow <= headerRow
  ) {

    return null;

  }


  const lastColumn =
    sheet.getLastColumn();


  const values =
    sheet
      .getRange(
        headerRow + 1,
        1,
        lastRow - headerRow,
        lastColumn
      )
      .getValues();


  /******************************************************
   * SEARCH FROM BOTTOM TO TOP
   *
   * This finds the newest open record.
   ******************************************************/

  for (
    let i =
      values.length - 1;
    i >= 0;
    i--
  ) {

    const row =
      values[i];


    const rowOperator =
      String(
        row[
          cols.operator - 1
        ] || ""
      )
      .trim();


    /****************************************************
     * OPERATOR MATCH
     ****************************************************/

    if (
      rowOperator.toLowerCase() !==
      operator.toLowerCase()
    ) {

      continue;

    }


    const start =
      row[
        cols.start - 1
      ];


    const end =
      row[
        cols.end - 1
      ];


    /****************************************************
     * OPEN RECORD
     ****************************************************/

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


/************************************************************
 * PARSE DATE / TIME
 ************************************************************/

function parseDateTime_(
  value
) {

  if (!value) {

    return null;

  }


  const date =
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


/************************************************************
 * PARSE DATE FROM GOOGLE SHEETS
 ************************************************************/

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


/************************************************************
 * FORMAT DURATION
 ************************************************************/

function formatDuration_(
  milliseconds
) {

  const totalSeconds =
    Math.round(
      milliseconds /
      1000
    );


  const hours =
    Math.floor(
      totalSeconds /
      3600
    );


  const minutes =
    Math.floor(
      (
        totalSeconds %
        3600
      ) /
      60
    );


  const seconds =
    totalSeconds %
    60;


  return (
    hours +
    ":" +
    String(
      minutes
    ).padStart(
      2,
      "0"
    ) +
    ":" +
    String(
      seconds
    ).padStart(
      2,
      "0"
    )
  );

}


/************************************************************
 * JSON RESPONSE
 ************************************************************/

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
