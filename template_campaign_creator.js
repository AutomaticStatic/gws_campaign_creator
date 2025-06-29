// Google Apps Script template for campaign calendars


// --- Configuration ---
var MAX_MONTH_MENU_ITEMS = 20; // Define a reasonable maximum for dynamic menu items
                               // Ensure you create handler functions up to this number (e.g., handleMonthSheet1 to handleMonthSheet20)

// --- Main Menu Function ---
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = ss.getSheets();
  
  var menu = ui.createMenu('Campaign Creator');
  var subMenu = ui.createMenu('Create Drafts');
  
  var months = ['January', 'February', 'March', 'April', 'May', 'June', 
                'July', 'August', 'September', 'October', 'November', 'December'];
  
  var hasMonthSheets = false;
  var monthSheetCounter = 0; 
  var userProperties = PropertiesService.getUserProperties();

  // Clear any previous properties to avoid conflicts if sheet names change (optional, consider if needed)
  // for (var k = 1; k <= MAX_MONTH_MENU_ITEMS; k++) {
  //   userProperties.deleteProperty('month_sheet_name_' + k);
  // }

  for (var i = 0; i < allSheets.length; i++) {
    if (monthSheetCounter >= MAX_MONTH_MENU_ITEMS) {
      Logger.log("Reached maximum number of dynamic menu items (" + MAX_MONTH_MENU_ITEMS + "). Some sheets may not be added to the menu.");
      break; 
    }

    var currentSheet = allSheets[i];
    var sheetName = currentSheet.getName();
    
    if (currentSheet.isSheetHidden()) {
      continue;
    }
    
    var containsMonth = false;
    for (var j = 0; j < months.length && !containsMonth; j++) {
      if (sheetName.toLowerCase().indexOf(months[j].toLowerCase()) !== -1) { // Case-insensitive check
        containsMonth = true;
      }
    }
    
    if (containsMonth) {
      monthSheetCounter++;
      userProperties.setProperty('month_sheet_name_' + monthSheetCounter, sheetName);
      subMenu.addItem("Create draft campaigns for '" + sheetName + "'", 'handleMonthSheet' + monthSheetCounter);
      hasMonthSheets = true;
    }
  }
  
  if (hasMonthSheets) {
    menu.addSubMenu(subMenu);
  }
  
  menu.addSeparator();
  menu.addToUi();
}

// --- Handler Functions for Dynamic Menu Items ---
// IMPORTANT: Create these functions up to MAX_MONTH_MENU_ITEMS
// For example, if MAX_MONTH_MENU_ITEMS is 20, you need handleMonthSheet1 through handleMonthSheet20

// --- Handler Functions for Dynamic Menu Items ---
// MAX_MONTH_MENU_ITEMS is set to 20, so these handlers go up to handleMonthSheet20.

function handleMonthSheet1() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_1');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 1. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet2() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_2');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 2. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet3() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_3');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 3. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet4() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_4');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 4. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet5() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_5');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 5. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet6() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_6');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 6. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet7() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_7');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 7. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet8() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_8');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 8. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet9() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_9');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 9. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet10() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_10');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 10. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet11() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_11');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 11. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet12() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_12');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 12. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet13() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_13');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 13. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet14() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_14');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 14. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet15() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_15');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 15. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet16() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_16');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 16. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet17() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_17');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 17. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet18() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_18');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 18. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet19() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_19');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 19. Please re-open the spreadsheet.');
  }
}

function handleMonthSheet20() {
  var sheetName = PropertiesService.getUserProperties().getProperty('month_sheet_name_20');
  if (sheetName) {
    createDraftsForSelectedSheet(sheetName);
  } else {
    SpreadsheetApp.getUi().alert('Error: Could not retrieve sheet name for menu item 20. Please re-open the spreadsheet.');
  }
}

// --- Helper function to create self-closing processing dialog HTML ---
function createSelfClosingProcessingDialog(sheetName) {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 20px;
            text-align: center;
          }
          .loading {
            margin: 20px 0;
          }
          .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #3498db;
            border-radius: 50%;
            width: 30px;
            height: 30px;
            animation: spin 1s linear infinite;
            margin: 0 auto;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .complete {
            color: #28a745;
            font-weight: bold;
          }
        </style>
      </head>
      <body>
        <div id="content" class="loading">
          <div class="spinner"></div>
          <p>Processing sheet: <b>${sheetName}</b>... Please wait.</p>
          <p>This may take a moment.</p>
        </div>
        <script>
          // Auto-close after 15 seconds as a safety measure
          setTimeout(function() {
            document.getElementById('content').innerHTML = '<div class="complete">✓ Processing complete!</div><p><small>Closing...</small></p>';
            setTimeout(function() {
              google.script.host.close();
            }, 1000);
          }, 15000);
        </script>
      </body>
    </html>
  `;
}

// --- Core Logic for Creating Drafts ---
function createDraftsForSelectedSheet(sheetName) {
  var ui = SpreadsheetApp.getUi();
  
  var result = ui.alert(
    'Confirm Draft Creation',
    'Are you sure you want to create draft campaigns for the "' + sheetName + '" sheet?',
    ui.ButtonSet.YES_NO
  );
  
  if (result == ui.Button.YES) {
    // Show the self-closing processing dialog
    var htmlOutput = HtmlService.createHtmlOutput(createSelfClosingProcessingDialog(sheetName))
      .setWidth(350)
      .setHeight(150);
    ui.showModalDialog(htmlOutput, "Processing...");

    try {
      var response = invokeLambdaWithSheetInfo(sheetName);

      // The dialog will auto-close after 15 seconds, but we'll show the success message regardless
      if (response && response.success) {
        var numCampaignsCreated = response.created_count;
        var successAlertMessage = "";

        if (numCampaignsCreated !== undefined && numCampaignsCreated !== null) {
          successAlertMessage = numCampaignsCreated + (numCampaignsCreated === 1 ? " campaign" : " campaigns") + " successfully created for sheet '" + sheetName + "'.";
        } else {
          successAlertMessage = "Draft campaigns processed successfully for sheet '" + sheetName + "'.";
        }
        
        if (response.message && (numCampaignsCreated === undefined || numCampaignsCreated === null)) { 
             successAlertMessage += "\nLambda: " + response.message;
        }

        ui.alert('Success', successAlertMessage, ui.ButtonSet.OK);

      } else {
        var errorMessage = "Failed to create draft campaigns for '" + sheetName + "'.";
        if (response && response.message) {
          errorMessage += " Details: " + response.message;
        } else if (response && response.errorDetail) {
          errorMessage += " Details: " + response.errorDetail;
        } else if (response) {
          errorMessage += " Response: " + JSON.stringify(response);
        } else {
          errorMessage += " No response or unknown error from Lambda.";
        }
        ui.alert('Error', errorMessage, ui.ButtonSet.OK);
        Logger.log("Lambda invocation failed or returned error for sheet '" + sheetName + "'. Response: " + JSON.stringify(response));
      }
    } catch (e) {
      Logger.log("Error in createDraftsForSelectedSheet for sheet '" + sheetName + "': " + e.toString() + " Stack: " + e.stack);
      ui.alert("Script Error", "An unexpected error occurred while processing '" + sheetName + "': " + e.message, ui.ButtonSet.OK);
    }
  }
}

// --- AWS Lambda Invocation with Signature V4 ---
function invokeLambdaWithSheetInfo(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetId = ss.getId();
  
  const region = "us-east-2"; 
  const accessKey = "ACCESS_KEY_HERE"; // Access Key for the IAM User apps-script-campaign-creator
  const secretKey = "SECET_KEY_HERE"; // Secret Key for IAM User apps-script-campaign-creator
  const service = "lambda";
  
  const endpoint = "https://uncv3oqcyyjwbzcfwjp6hn55gu0gosaz.lambda-url.us-east-2.on.aws/"; // Function URL for the CampaignCreator lambda
  const host = extractHostFromUrl(endpoint);
  
  const payload = {
    "client_id": "CLIENT_ID_HERE", // e.g. "cor"
    "sheet_id": sheetId,
    "sheet_name": sheetName
  };
  
  const date = new Date();
  const amzDate = date.toISOString().replace(/[:-]|\.\d{3}/g, '');
  const dateStamp = amzDate.slice(0, 8);
  
  const method = "POST";
  const canonicalUri = "/"; 
  const canonicalQueryString = "";
  const payloadStr = JSON.stringify(payload);
  const payloadHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payloadStr);
  const payloadHashHex = toHex(payloadHash);
  
  const canonicalHeaders = 
    "content-type:application/json\n" + 
    "host:" + host + "\n" + 
    "x-amz-date:" + amzDate + "\n";
  const signedHeaders = "content-type;host;x-amz-date";
  
  const canonicalRequest = method + "\n" + 
                          canonicalUri + "\n" + 
                          canonicalQueryString + "\n" + 
                          canonicalHeaders + "\n" + 
                          signedHeaders + "\n" + 
                          payloadHashHex;
  
  const algorithm = "AWS4-HMAC-SHA256";
  const credentialScope = dateStamp + "/" + region + "/" + service + "/aws4_request";
  const stringToSign = algorithm + "\n" + 
                      amzDate + "\n" + 
                      credentialScope + "\n" + 
                      toHex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, canonicalRequest));
  
  // --- Calculate signature with corrected kDate derivation ---
  const kDate = Utilities.computeHmacSha256Signature(
    Utilities.newBlob(dateStamp).getBytes(),
    Utilities.newBlob("AWS4" + secretKey).getBytes() // Corrected line
  );
  
  const kRegion = Utilities.computeHmacSha256Signature(Utilities.newBlob(region).getBytes(), kDate);
  const kService = Utilities.computeHmacSha256Signature(Utilities.newBlob(service).getBytes(), kRegion);
  const kSigning = Utilities.computeHmacSha256Signature(Utilities.newBlob("aws4_request").getBytes(), kService);
  const signatureBytes = Utilities.computeHmacSha256Signature(Utilities.newBlob(stringToSign).getBytes(), kSigning);
  const signatureHex = toHex(signatureBytes);
  
  const authorizationHeader = algorithm + " " + 
                            "Credential=" + accessKey + "/" + credentialScope + ", " + 
                            "SignedHeaders=" + signedHeaders + ", " + 
                            "Signature=" + signatureHex;
  
  const options = {
    "method": method,
    "contentType": "application/json",
    "payload": payloadStr,
    "muteHttpExceptions": true, // Important for parsing error responses
    "headers": {
      "X-Amz-Date": amzDate,
      "Authorization": authorizationHeader
      // "Host": host, // UrlFetchApp adds Host header automatically; it's signed, which is good.
    }
  };
  
  try {
    Logger.log("Invoking Lambda. Endpoint: " + endpoint + " Options: " + JSON.stringify(options));
    const httpResponse = UrlFetchApp.fetch(endpoint, options);
    const responseCode = httpResponse.getResponseCode();
    const responseText = httpResponse.getContentText();
    Logger.log("Lambda Response Code: " + responseCode);
    Logger.log("Lambda Response Text: " + responseText);

    if (responseCode >= 200 && responseCode < 300) {
        if (responseText) {
            return JSON.parse(responseText);
        } else {
            return { success: true, message: "Lambda returned success with no content." }; // Adjust if Lambda should always return content
        }
    } else {
        var errorData = { success: false, message: "Lambda returned HTTP error " + responseCode + "."};
        try {
            var parsedError = JSON.parse(responseText);
            errorData.errorDetail = parsedError.message || responseText; // AWS Lambda errors often have a "message" field
        } catch (e) {
            errorData.errorDetail = responseText;
        }
        Logger.log("Error data from Lambda: " + JSON.stringify(errorData));
        return errorData;
    }
  } catch (error) {
    Logger.log("Error invoking Lambda: " + error.toString() + " Stack: " + error.stack);
    return { success: false, message: "Failed to invoke Lambda due to script error.", errorDetail: error.toString() };
  }
}

// --- Helper Functions ---
function extractHostFromUrl(url) {
  var hostWithPath = url.replace(/^(https?:\/\/)/, '');
  var host = hostWithPath.split('/')[0];
  return host;
}

function toHex(byteArray) {
  return Array.from(byteArray)
    .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2))
    .join('');
}