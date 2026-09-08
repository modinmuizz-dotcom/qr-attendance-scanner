function generateOperatorQRCodes() {

  // ==========================================================
  // SETTINGS
  // ==========================================================

  const SHEET_NAME = "DATA DON'T DELETE";
  const FOLDER_NAME = "generated_qr";

  // Possible names for the ID column
  const ID_HEADERS = [
    "ID",
    "Id",
    "id"
  ];

  // Possible names for the operator column
  const OPERATOR_HEADERS = [
    "NAME OF OPERATORS",
    "Name of Operators",
    "NAME OF OPERATOR",
    "Name of Operator"
  ];


  // ==========================================================
  // GET SPREADSHEET
  // ==========================================================

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error(
      'Sheet "' + SHEET_NAME + '" was not found.'
    );
  }


  // ==========================================================
  // GET ALL DATA
  // ==========================================================

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2) {
    throw new Error(
      'The sheet "' + SHEET_NAME + '" does not contain any data.'
    );
  }

  const data = sheet
    .getRange(1, 1, lastRow, lastColumn)
    .getValues();


  // ==========================================================
  // FIND HEADER ROW
  // ==========================================================

  const headers = data[0].map(function(header) {
    return String(header).trim();
  });


  // ==========================================================
  // FIND ID COLUMN DYNAMICALLY
  // ==========================================================

  const idColumn = findColumn(headers, ID_HEADERS);

  if (idColumn === -1) {

    throw new Error(
      'Could not find the "ID" column in "' +
      SHEET_NAME +
      '".\n\n' +
      'Make sure your header is named "ID".'
    );

  }


  // ==========================================================
  // FIND OPERATOR COLUMN DYNAMICALLY
  // ==========================================================

  const operatorColumn = findColumn(
    headers,
    OPERATOR_HEADERS
  );

  if (operatorColumn === -1) {

    throw new Error(
      'Could not find the "NAME OF OPERATORS" column in "' +
      SHEET_NAME +
      '".\n\n' +
      'Make sure your header is named "NAME OF OPERATORS".'
    );

  }


  Logger.log(
    'ID column: ' +
    (idColumn + 1)
  );

  Logger.log(
    'NAME OF OPERATORS column: ' +
    (operatorColumn + 1)
  );


  // ==========================================================
  // GET / CREATE GENERATED_QR FOLDER
  // ==========================================================

  const folder = getOrCreateFolder(FOLDER_NAME);


  // ==========================================================
  // DELETE OLD QR CODES
  // ==========================================================

  deleteEverythingInFolder(folder);


  // ==========================================================
  // GENERATE QR CODES
  // ==========================================================

  let generatedCount = 0;
  let skippedCount = 0;

  for (let row = 1; row < data.length; row++) {

    const id = String(
      data[row][idColumn]
    ).trim();

    const operatorName = String(
      data[row][operatorColumn]
    ).trim();


    // --------------------------------------------------------
    // SKIP EMPTY ROWS
    // --------------------------------------------------------

    if (!id || !operatorName) {

      skippedCount++;

      continue;
    }


    // --------------------------------------------------------
    // JSON DATA
    // --------------------------------------------------------

    const qrJson = {
      "NAME OF OPERATORS": operatorName
    };


    const jsonString = JSON.stringify(qrJson);


    Logger.log(
      'Generating QR for ID: ' +
      id +
      ' | Operator: ' +
      operatorName
    );


    // --------------------------------------------------------
    // GOQR.ME API
    // --------------------------------------------------------

    const apiUrl =
      "https://api.qrserver.com/v1/create-qr-code/" +
      "?size=500x500" +
      "&format=png" +
      "&data=" +
      encodeURIComponent(jsonString);


    try {

      const response = UrlFetchApp.fetch(
        apiUrl,
        {
          method: "get",
          muteHttpExceptions: true
        }
      );


      const responseCode =
        response.getResponseCode();


      // ------------------------------------------------------
      // CHECK API RESPONSE
      // ------------------------------------------------------

      if (responseCode !== 200) {

        Logger.log(
          "FAILED: ID " +
          id +
          " | HTTP " +
          responseCode
        );

        skippedCount++;

        continue;
      }


      // ------------------------------------------------------
      // GET IMAGE
      // ------------------------------------------------------

      const qrBlob = response.getBlob();


      // ------------------------------------------------------
      // FILE NAME
      // ------------------------------------------------------

      qrBlob.setName(
        id + ".png"
      );


      // ------------------------------------------------------
      // SAVE TO DRIVE
      // ------------------------------------------------------

      folder.createFile(qrBlob);


      generatedCount++;


      Logger.log(
        "SUCCESS: " +
        id +
        ".png"
      );


    } catch (error) {

      Logger.log(
        "ERROR generating QR for ID " +
        id +
        ": " +
        error.message
      );

      skippedCount++;
    }


    // --------------------------------------------------------
    // SMALL DELAY
    // Helps avoid sending requests too quickly.
    // --------------------------------------------------------

    Utilities.sleep(100);
  }


  // ==========================================================
  // COMPLETION MESSAGE
  // ==========================================================

  SpreadsheetApp.getUi().alert(
    "QR CODE GENERATION COMPLETE\n\n" +
    "Generated QR Codes: " +
    generatedCount +
    "\n" +
    "Skipped: " +
    skippedCount +
    "\n\n" +
    "Saved to Drive folder:\n" +
    FOLDER_NAME
  );
}


/**
 * ============================================================
 * FIND COLUMN
 * ============================================================
 *
 * Searches the header row for the correct column.
 * Column order does NOT matter.
 */
function findColumn(headers, possibleHeaders) {

  for (let i = 0; i < headers.length; i++) {

    const currentHeader =
      String(headers[i])
        .trim()
        .toLowerCase();

    for (let j = 0; j < possibleHeaders.length; j++) {

      const possibleHeader =
        String(possibleHeaders[j])
          .trim()
          .toLowerCase();

      if (currentHeader === possibleHeader) {

        return i;
      }
    }
  }

  return -1;
}


/**
 * ============================================================
 * GET OR CREATE DRIVE FOLDER
 * ============================================================
 */
function getOrCreateFolder(folderName) {

  const folders =
    DriveApp.getFoldersByName(folderName);


  // Folder already exists
  if (folders.hasNext()) {

    return folders.next();
  }


  // Folder doesn't exist
  return DriveApp.createFolder(folderName);
}


/**
 * ============================================================
 * DELETE EVERYTHING INSIDE FOLDER
 * ============================================================
 *
 * IMPORTANT:
 * This does NOT delete the folder itself.
 *
 * Files are moved to Google Drive Trash.
 */
function deleteEverythingInFolder(folder) {

  const files =
    folder.getFiles();


  while (files.hasNext()) {

    const file =
      files.next();

    file.setTrashed(true);
  }
}

