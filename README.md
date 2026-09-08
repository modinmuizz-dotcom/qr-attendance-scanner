<!DOCTYPE html>
<html lang="en">

<head>

  <meta charset="UTF-8">

  <meta
    name="viewport"
    content="width=device-width, initial-scale=1.0"
  >

  <title>QR Attendance Scanner</title>

  <!-- HTML5 QR CODE -->
  <script src="[https://cdn.jsdelivr.net/npm/html5-qrcode/minified/html5-qrcode.min.js](https://script.google.com/macros/s/AKfycbyCbQgOTaaPtx3Q7D3IlM4ZuK199PtZsQ8UfXFBZ9RGR20vkHpdgX4n_o3oKHlDw2Ez/exec)"></script>

  <style>

    * {
      box-sizing: border-box;
    }

    body {
      margin: 0;
      padding: 20px;
      background: #f3f4f6;
      font-family: Arial, sans-serif;
    }

    .container {
      max-width: 650px;
      margin: 0 auto;
      background: white;
      padding: 25px;
      border-radius: 18px;
      box-shadow: 0 5px 25px rgba(0,0,0,0.10);
    }

    h1 {
      text-align: center;
      margin-top: 0;
      color: #111827;
      font-size: 28px;
    }

    #reader {
      width: 100%;
      max-width: 500px;
      margin: 20px auto;
      overflow: hidden;
      border-radius: 15px;
    }

    #status {
      padding: 14px;
      border-radius: 10px;
      text-align: center;
      margin: 15px 0;
      font-weight: bold;
      display: none;
    }

    .status-info {
      background: #e0f2fe;
      color: #075985;
    }

    .status-success {
      background: #dcfce7;
      color: #166534;
    }

    .status-error {
      background: #fee2e2;
      color: #991b1b;
    }

    #details {
      display: none;
      margin-top: 20px;
    }

    .detail {
      display: flex;
      justify-content: space-between;
      gap: 20px;
      padding: 13px 5px;
      border-bottom: 1px solid #e5e7eb;
    }

    .label {
      font-weight: bold;
      color: #374151;
    }

    .value {
      text-align: right;
      color: #111827;
      word-break: break-word;
    }

    .buttons {
      display: none;
      gap: 15px;
      margin-top: 25px;
    }

    button {
      flex: 1;
      border: none;
      padding: 17px;
      border-radius: 12px;
      font-size: 18px;
      font-weight: bold;
      cursor: pointer;
    }

    #inButton {
      background: #16a34a;
      color: white;
    }

    #outButton {
      background: #dc2626;
      color: white;
    }

    button:disabled {
      opacity: 0.5;
      cursor: not-allowed;
    }

    #scanAgain {
      display: none;
      width: 100%;
      margin-top: 15px;
      background: #2563eb;
      color: white;
    }

    .camera-help {
      text-align: center;
      color: #6b7280;
      font-size: 14px;
      margin-top: 10px;
    }

    @media (max-width: 500px) {

      body {
        padding: 10px;
      }

      .container {
        padding: 18px;
      }

      h1 {
        font-size: 23px;
      }

      .detail {
        display: block;
      }

      .value {
        text-align: left;
        margin-top: 5px;
      }

      .buttons {
        flex-direction: column;
      }

    }

  </style>

</head>


<body>

<div class="container">

  <h1>QR Attendance Scanner</h1>

  <div id="status"></div>

  <div id="reader"></div>

  <div class="camera-help">
    Point the camera at the operator QR code.
  </div>


  <!-- SCANNED INFORMATION -->

  <div id="details">

    <div class="detail">
      <span class="label">Operator / Driver</span>
      <span class="value" id="operator"></span>
    </div>

    <div class="detail">
      <span class="label">Start Time</span>
      <span class="value" id="startTime"></span>
    </div>

    <div class="detail">
      <span class="label">End Time</span>
      <span class="value" id="endTime"></span>
    </div>

    <div class="detail">
      <span class="label">Date</span>
      <span class="value" id="date"></span>
    </div>

    <div class="detail">
      <span class="label">Operating Hrs</span>
      <span class="value" id="operatingHrs"></span>
    </div>

  </div>


  <!-- BUTTONS -->

  <div class="buttons" id="buttons">

    <button
      id="inButton"
      onclick="recordIN()"
    >
      IN
    </button>

    <button
      id="outButton"
      onclick="recordOUT()"
    >
      OUT
    </button>

  </div>


  <button
    id="scanAgain"
    onclick="resetScanner()"
  >
    SCAN AGAIN
  </button>

</div>


<script>

/*******************************************************
 * IMPORTANT
 *
 * PUT YOUR APPS SCRIPT /exec URL HERE
 *******************************************************/

const APPS_SCRIPT_URL =
  "https://modinmuizz-dotcom.github.io/qr-attendance-scanner/";


let scanner = null;
let scannedData = null;
let scanTime = null;
let processing = false;


/*******************************************************
 * START
 *******************************************************/

window.addEventListener("load", function() {

  startScanner();

});


/*******************************************************
 * START CAMERA
 *******************************************************/

async function startScanner() {

  try {

    processing = false;

    showStatus(
      "Requesting camera permission...",
      "info"
    );


    // First request camera permission directly.
    const stream =
      await navigator.mediaDevices.getUserMedia({

        video: {
          facingMode: {
            ideal: "environment"
          }
        },

        audio: false

      });


    // Stop temporary permission stream.
    stream.getTracks().forEach(
      track => track.stop()
    );


    createScanner();

  }

  catch (error) {

    console.error(
      "Camera error:",
      error
    );

    let message =
      "Unable to open camera.";

    if (
      error.name === "NotAllowedError" ||
      error.name === "PermissionDeniedError"
    ) {

      message =
        "Camera permission was denied. Please allow camera access and reload this page.";

    }

    else if (
      error.name === "NotFoundError"
    ) {

      message =
        "No camera was found on this device.";

    }

    else if (
      error.name === "NotReadableError"
    ) {

      message =
        "The camera is being used by another application.";

    }

    else if (
      error.name === "SecurityError"
    ) {

      message =
        "Camera access is blocked by the browser.";

    }


    showStatus(
      message,
      "error"
    );

  }

}


/*******************************************************
 * CREATE QR SCANNER
 *******************************************************/

function createScanner() {

  if (scanner) {

    try {
      scanner.clear();
    }
    catch (e) {}

  }


  scanner =
    new Html5Qrcode("reader");


  const reader =
    document.getElementById("reader");


  const width =
    reader.clientWidth || 300;


  const qrSize =
    Math.min(
      Math.max(width * 0.70, 220),
      350
    );


  scanner.start(

    {
      facingMode: "environment"
    },

    {
      fps: 10,

      qrbox: {
        width: qrSize,
        height: qrSize
      },

      aspectRatio: 1.0

    },

    onScanSuccess,

    onScanFailure

  )

  .then(function() {

    showStatus(
      "Camera ready. Scan a QR code.",
      "success"
    );

  })

  .catch(function(error) {

    console.error(
      "Scanner start error:",
      error
    );

    showStatus(
      "Unable to start QR scanner.",
      "error"
    );

  });

}


/*******************************************************
 * QR SCAN SUCCESS
 *******************************************************/

async function onScanSuccess(decodedText) {

  if (processing) {
    return;
  }

  processing = true;


  // Capture the actual QR scan time.
  scanTime =
    new Date().toISOString();


  // Stop camera
  await stopScanner();


  let data;


  /***************************************************
   * PARSE JSON
   ***************************************************/

  try {

    data = JSON.parse(decodedText);

  }

  catch (error) {

    showStatus(
      "Invalid QR code. The QR code does not contain valid JSON.",
      "error"
    );

    showScanAgain();

    return;

  }


  /***************************************************
   * CHECK OBJECT
   ***************************************************/

  if (
    typeof data !== "object" ||
    data === null ||
    Array.isArray(data)
  ) {

    showStatus(
      "Invalid QR data.",
      "error"
    );

    showScanAgain();

    return;

  }


  /***************************************************
   * REQUIRED FIELDS
   ***************************************************/

  const requiredFields = [

    "Operator/Driver",
    "Start Time",
    "End Time",
    "Operating Hrs",
    "Date"

  ];


  for (
    const field of requiredFields
  ) {

    if (
      data[field] === undefined ||
      data[field] === null ||
      String(data[field]).trim() === ""
    ) {

      showStatus(
        "Invalid QR code. Missing: " + field,
        "error"
      );

      showScanAgain();

      return;

    }

  }


  /***************************************************
   * SAVE SCANNED DATA
   ***************************************************/

  scannedData = data;


  document.getElementById(
    "operator"
  ).textContent =
    data["Operator/Driver"];


  document.getElementById(
    "startTime"
  ).textContent =
    data["Start Time"];


  document.getElementById(
    "endTime"
  ).textContent =
    data["End Time"];


  document.getElementById(
    "date"
  ).textContent =
    data["Date"];


  document.getElementById(
    "operatingHrs"
  ).textContent =
    data["Operating Hrs"];


  document.getElementById(
    "details"
  ).style.display =
    "block";


  document.getElementById(
    "buttons"
  ).style.display =
    "flex";


  showStatus(
    "QR code successfully scanned.",
    "success"
  );

}


/*******************************************************
 * NORMAL QR SCAN FAILURE
 *******************************************************/

function onScanFailure(error) {

  // Ignore normal scanning failures.
}


/*******************************************************
 * STOP SCANNER
 *******************************************************/

async function stopScanner() {

  if (!scanner) {
    return;
  }

  try {

    await scanner.stop();

  }

  catch (error) {

    console.log(
      "Scanner stop:",
      error
    );

  }

  try {

    await scanner.clear();

  }

  catch (error) {}

  scanner = null;

}


/*******************************************************
 * IN
 *******************************************************/

function recordIN() {

  if (!scannedData) {

    showStatus(
      "Please scan a QR code first.",
      "error"
    );

    return;

  }


  disableButtons();


  showStatus(
    "Saving IN...",
    "info"
  );


  submitToAppsScript(
    "IN"
  );

}


/*******************************************************
 * OUT
 *******************************************************/

function recordOUT() {

  if (!scannedData) {

    showStatus(
      "Please scan a QR code first.",
      "error"
    );

    return;

  }


  disableButtons();


  showStatus(
    "Saving OUT...",
    "info"
  );


  submitToAppsScript(
    "OUT"
  );

}


/*******************************************************
 * SEND DATA TO APPS SCRIPT
 *
 * Uses a normal HTML form POST instead of fetch().
 *
 * This avoids the cross-origin/CORS problem between
 * GitHub Pages and Google Apps Script.
 *******************************************************/

function submitToAppsScript(action) {


  const iframeName =
    "appsScriptResponseFrame";


  let iframe =
    document.getElementById(
      iframeName
    );


  if (!iframe) {

    iframe =
      document.createElement("iframe");

    iframe.id =
      iframeName;

    iframe.name =
      iframeName;

    iframe.style.display =
      "none";

    document.body.appendChild(
      iframe
    );

  }


  const form =
    document.createElement("form");


  form.method =
    "POST";

  form.action =
    APPS_SCRIPT_URL;

  form.target =
    iframeName;

  form.style.display =
    "none";


  /***************************************************
   * Add hidden field
   ***************************************************/

  function addField(
    name,
    value
  ) {

    const input =
      document.createElement("input");

    input.type =
      "hidden";

    input.name =
      name;

    input.value =
      value == null
        ? ""
        : String(value);

    form.appendChild(
      input
    );

  }


  addField(
    "action",
    action
  );


  addField(
    "operator",
    scannedData["Operator/Driver"]
  );


  addField(
    "startTime",
    scannedData["Start Time"]
  );


  addField(
    "endTime",
    scannedData["End Time"]
  );


  addField(
    "operatingHrs",
    scannedData["Operating Hrs"]
  );


  addField(
    "date",
    scannedData["Date"]
  );


  // Actual time QR was scanned
  addField(
    "scanTime",
    scanTime
  );


  document.body.appendChild(
    form
  );


  form.submit();


  /***************************************************
   * We cannot reliably read the cross-origin response.
   *
   * Give Apps Script time to process the request,
   * then reset the scanner.
   ***************************************************/

  setTimeout(
    function() {

      showStatus(

        action === "IN"
          ? "✓ IN successfully recorded."
          : "✓ OUT successfully recorded.",

        "success"

      );


      form.remove();


      setTimeout(
        resetScanner,
        1500
      );

    },

    1200

  );

}


/*******************************************************
 * DISABLE BUTTONS
 *******************************************************/

function disableButtons() {

  document.getElementById(
    "inButton"
  ).disabled = true;


  document.getElementById(
    "outButton"
  ).disabled = true;

}


/*******************************************************
 * SHOW SCAN AGAIN
 *******************************************************/

function showScanAgain() {

  document.getElementById(
    "scanAgain"
  ).style.display =
    "block";

}


/*******************************************************
 * RESET
 *******************************************************/

async function resetScanner() {

  await stopScanner();


  scannedData = null;

  scanTime = null;

  processing = false;


  document.getElementById(
    "details"
  ).style.display =
    "none";


  document.getElementById(
    "buttons"
  ).style.display =
    "none";


  document.getElementById(
    "scanAgain"
  ).style.display =
    "none";


  document.getElementById(
    "inButton"
  ).disabled =
    false;


  document.getElementById(
    "outButton"
  ).disabled =
    false;


  document.getElementById(
    "operator"
  ).textContent =
    "";


  document.getElementById(
    "startTime"
  ).textContent =
    "";


  document.getElementById(
    "endTime"
  ).textContent =
    "";


  document.getElementById(
    "date"
  ).textContent =
    "";


  document.getElementById(
    "operatingHrs"
  ).textContent =
    "";


  document.getElementById(
    "reader"
  ).innerHTML =
    "";


  showStatus(
    "Starting camera...",
    "info"
  );


  startScanner();

}


/*******************************************************
 * STATUS
 *******************************************************/

function showStatus(
  message,
  type
) {

  const status =
    document.getElementById(
      "status"
    );


  status.textContent =
    message;


  status.className =
    "";


  status.classList.add(
    "status-" + type
  );


  status.style.display =
    "block";

}

</script>

</body>
</html>
