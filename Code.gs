/**
 * @NotOnlyCurrentDoc
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Wright Finder Motors")
    .addItem("Open App", "openApp")
    .addToUi();
}

/**
 * RUN THIS FUNCTION ONCE MANUALLY FROM THE APPS SCRIPT EDITOR.
 * This will trigger the Google authorization prompt so you can grant Drive and Spreadsheet permissions.
 */
function setupAuthorization() {
  // Creating a dummy file forces Apps Script to request full Drive write permissions
  try {
    var dummy = DriveApp.createFile("auth_test.txt", "test");
    dummy.setTrashed(true); // cleanup
  } catch (e) {}
  SpreadsheetApp.getActiveSpreadsheet();
}

function testCreate() {
  var folderId = "1EUPGHZPwovNhVOsIc-AEEJqUinAqEZKK";
  var folder = DriveApp.getFolderById(folderId);
  folder.createFile("test_direct.txt", "It works!");
}

function doGet(e) {
  var requestId = e.parameter.requestId;
  if (requestId) {
    var template = HtmlService.createTemplateFromFile("UploadDocument");
    template.requestId = requestId;
    template.requestDetails = JSON.stringify(
      getDocumentRequestDetails(requestId),
    );
    return template
      .evaluate()
      .setTitle("Upload Requested Document - Wright Finder Motors")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }

  var id = e.parameter.id || e.parameter.carId;
  if (id) {
    var template = HtmlService.createTemplateFromFile("SignAgreement");
    template.data = getAgreementData(id);
    template.carId = id;
    return template
      .evaluate()
      .setTitle("Sign Purchase Agreement - Wright Finder Motors")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1");
  }
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Wright Finder Motors App")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}
function openApp() {
  var html = HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Wright Finder Motors App")
    .setWidth(1000)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, "Wright Finder Motors");
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function saveNewVehicle(data) {
  var folderId = "1EUPGHZPwovNhVOsIc-AEEJqUinAqEZKK";
  var folder = DriveApp.getFolderById(folderId);

  var mainImageUrls = [];
  if (data.mainImages && data.mainImages.length > 0) {
    for (var i = 0; i < data.mainImages.length; i++) {
      var img = data.mainImages[i];
      var blob = Utilities.newBlob(
        Utilities.base64Decode(img.data),
        img.mimeType,
        img.name,
      );
      var file = DriveApp.createFile(blob);
      file.moveTo(folder);
      mainImageUrls.push(file.getUrl());
    }
  }

  var subImageUrls = [];
  if (data.subImages && data.subImages.length > 0) {
    for (var j = 0; j < data.subImages.length; j++) {
      var imgSub = data.subImages[j];
      var blobSub = Utilities.newBlob(
        Utilities.base64Decode(imgSub.data),
        imgSub.mimeType,
        imgSub.name,
      );
      var fileSub = DriveApp.createFile(blobSub);
      fileSub.moveTo(folder);
      subImageUrls.push(fileSub.getUrl());
    }
  }

  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");

  var cName = data.carName ? data.carName.toString().trim() : "";
  var firstLetter = cName.length > 0 ? cName.charAt(0).toUpperCase() : "X";

  var cModel = data.model ? data.model.toString().trim() : "";
  var secondLetter = cModel.length > 0 ? cModel.charAt(0).toUpperCase() : "X";

  var cYear = data.year ? data.year.toString().trim() : "";
  var yearStr = cYear.length >= 2 ? cYear.slice(-2) : "00";

  var randomNum = Math.floor(Math.random() * 100);
  var randomStr = randomNum < 10 ? "0" + randomNum : randomNum.toString();

  var carId = firstLetter + secondLetter + yearStr + "-" + randomStr;

  if (!sheet) {
    sheet =
      SpreadsheetApp.getActiveSpreadsheet().insertSheet("Vehicle details");
    sheet.appendRow([
      "Timestamp",
      "Car ID",
      "Car Name",
      "Model",
      "Year",
      "Mileage",
      "Price",
      "Discount",
      "VIN",
      "Status",
      "Title",
      "Trim",
      "Body Style",
      "Rent or Sell",
      "Engine",
      "Engine Type/Size",
      "Transmission",
      "Driveline",
      "Fuel Type",
      "Power Options",
      "Drive Condition",
      "Condition",
      "Seat Material",
      "Interior Color",
      "Exterior Color",
      "Interior Features",
      "Main Image URLs",
      "Sub Image URLs",
      "CLIENT NAME",
      "PURCHASE DATE",
      "SOLD DATE",
      "Trade status",
      "IAAI TOTAL PRICE W/ FEES",
      "PAPE PRICE",
      "TRANSPORT FEES",
      "SOLD PRICE",
      "DOWN PAYMENT",
      "CAR PICKUP LOCATION",
      "DRIVER NAME",
      "DRIVER INFORMATION",
      "NOTES",
      "IAAI Price Before Fees",
      "Dispatcher Name",
      "Dispatcher Price",
      "Dispatcher Phone number",
      "Client Phone",
      "Client Email",
      "Driver Phone",
      "Driver Company",
      "Tax Amount",
      "Price on title",
      "Financing status",
      "Tax Responsibility",
      "Trade Value",
      "Trade In",
      "Comment",
    ]);
  }

  sheet.appendRow([
    new Date(),
    carId,
    data.carName || "",
    data.model || "",
    data.year || "",
    data.mileage || "",
    data.price || "",
    data.discount || "",
    data.vin || "",
    data.status || "",
    data.title || "",
    data.trim || "",
    data.bodyStyle || "",
    data.rentOrSell || "",
    data.engine || "",
    data.engineType || "",
    data.transmission || "",
    data.driveline || "",
    data.fuelType || "",
    data.power || "",
    data.driveCondition || "",
    data.condition || "",
    data.seatMaterial || "",
    data.interiorColor || "",
    data.exteriorColor || "",
    data.interiorFeaturesStr || "",
    mainImageUrls.join(", "),
    subImageUrls.join(", "),
    "", // 28: CLIENT NAME
    data.purchaseDate || "", // 29: PURCHASE DATE
    "", // 30: SOLD DATE
    data.tradeStatus || "", // 31: Trade status
    "", // 32
    "", // 33
    "", // 34
    "", // 35
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "", // Dispatcher Name
    "", // Dispatcher Price
    "", // Dispatcher Phone number
    "", // Client Phone
    "", // Client Email
    "", // Driver Phone
    "", // Driver Company
    data["Tax Amount"] || "",
    data["Price on title"] || "",
    data["Financing status"] || "",
    data["Tax Responsibility"] || "",
  ]);

  return "Success";
}

function updateVehicleData(updatedData) {
  var folderId = "1EUPGHZPwovNhVOsIc-AEEJqUinAqEZKK";
  var folder = DriveApp.getFolderById(folderId);

  // Handle Image Updates if provided
  if (updatedData.mainImages && updatedData.mainImages.length > 0) {
    var mainImageUrls = [];
    for (var m = 0; m < updatedData.mainImages.length; m++) {
      var img = updatedData.mainImages[m];
      var mainBlob = Utilities.newBlob(
        Utilities.base64Decode(img.data),
        img.mimeType,
        img.name,
      );
      var mainFile = DriveApp.createFile(mainBlob);
      mainFile.moveTo(folder);
      mainImageUrls.push(mainFile.getUrl());
    }
    updatedData["Main Image URLs"] = mainImageUrls.join(", ");
  }
  delete updatedData.mainImages;

  if (updatedData.subImages && updatedData.subImages.length > 0) {
    var subImageUrls = [];
    for (var s = 0; s < updatedData.subImages.length; s++) {
      var imgSub = updatedData.subImages[s];
      var subBlob = Utilities.newBlob(
        Utilities.base64Decode(imgSub.data),
        imgSub.mimeType,
        imgSub.name,
      );
      var subFile = DriveApp.createFile(subBlob);
      subFile.moveTo(folder);
      subImageUrls.push(subFile.getUrl());
    }
    updatedData["Sub Image URLs"] = subImageUrls.join(", ");
  }
  delete updatedData.subImages;

  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
  if (!sheet) return "Error: Sheet not found";

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return "Error: No data";

  var headers = data[0];
  var carIdIndex = headers.indexOf("Car ID");
  if (carIdIndex === -1) return "Error: Car ID column not found";

  for (var i = 1; i < data.length; i++) {
    if (data[i][carIdIndex] === updatedData["Car ID"]) {
      // Found the row, update it
      for (var key in updatedData) {
        if (updatedData.hasOwnProperty(key) && key !== "Car ID") {
          var colIndex = headers.indexOf(key);
          if (colIndex !== -1) {
            sheet.getRange(i + 1, colIndex + 1).setValue(updatedData[key]);
          }
        }
      }

      // Handle Trade: Set selected inventory vehicle's Trade status to Trading (keep as Available)
      var purchasedCarName = updatedData["Purchased Car Name"] || "";
      if (purchasedCarName) {
        setCarTradeStatusToTrading(purchasedCarName);
      }

      // Automatically set the saved car (current vehicle) status to Sold if it's a Trade transaction
      if (updatedData["Trade status"] === "Trade" && purchasedCarName) {
        var statusCol = headers.indexOf("Status");
        if (statusCol !== -1) {
          sheet.getRange(i + 1, statusCol + 1).setValue("Sold");
          updatedData["Status"] = "Sold"; // Sync locally
        }
      }

      return "Success";
    }
  }
  return "Error: Car ID not found";
}

function setCarTradeStatusToTrading(carId) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  var headers = data[0];
  var carIdIndex = headers.indexOf("Car ID");
  var tradeStatusIndex = headers.indexOf("Trade status");
  if (carIdIndex === -1 || tradeStatusIndex === -1) return;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][carIdIndex]).trim() === String(carId).trim()) {
      sheet.getRange(i + 1, tradeStatusIndex + 1).setValue("Trading");
      break;
    }
  }
}

function getVehicles() {
  ensureAgreementColumns();
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
  if (!sheet) return [];

  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; // Only headers or empty

  var headers = data[0];
  var vehicles = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var vehicle = {};
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      vehicle[header] = row[j];
    }
    vehicles.push(vehicle);
  }

  return vehicles;
}

function saveExpense(data) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All expenses");

  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("All expenses");
    sheet.appendRow([
      "Timestamp",
      "CAR MODEL",
      "CAR ID",
      "Client Name",
      "DESCRIPTION",
      "PAID BY",
      "AMOUNT",
      "EXPENSE DATE",
    ]);
  } else {
    // Check if PAID BY header exists
    var lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
      if (headers.indexOf("PAID BY") === -1) {
        // Find index of DESCRIPTION to insert right after, or append at the end
        var descIndex = headers.indexOf("DESCRIPTION");
        if (descIndex !== -1) {
          sheet.insertColumnAfter(descIndex + 1);
          sheet.getRange(1, descIndex + 2).setValue("PAID BY");
        } else {
          sheet.getRange(1, lastCol + 1).setValue("PAID BY");
        }
      }
    }
  }

  // Get current headers to index map to dynamically build the row array
  var headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  var rowValues = new Array(headers.length);

  // Map our keys to header indices
  var keyMap = {
    Timestamp: new Date(),
    "CAR MODEL": data.carModel || "",
    "CAR ID": data.carId || "",
    "Client Name": data.clientName || "",
    DESCRIPTION: data.description || "",
    "PAID BY": data.paidBy || "",
    AMOUNT: data.amount || "",
    "EXPENSE DATE": data.expenseDate || "",
  };

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    rowValues[i] = keyMap[header] !== undefined ? keyMap[header] : "";
  }

  sheet.appendRow(rowValues);
  return "Success";
}

function getExpenses() {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All expenses");
  if (!sheet) return [];

  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; // Only headers or empty

  var headers = data[0];
  var expenses = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var expense = { rowNumber: i + 1 };
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      expense[header] = row[j];
    }
    expenses.push(expense);
  }

  return expenses;
}

function deleteExpense(rowNumber) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All expenses");
  if (!sheet) return "Error: Sheet not found";

  var row = parseInt(rowNumber);
  if (isNaN(row) || row <= 1) return "Error: Invalid row number";

  sheet.deleteRow(row);
  return "Success";
}

function updateExpense(rowNumber, data) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All expenses");
  if (!sheet) return "Error: Sheet not found";

  var row = parseInt(rowNumber);
  if (isNaN(row) || row <= 1) return "Error: Invalid row number";

  var headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  var rowValues = new Array(headers.length);

  var existingRow = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

  var keyMap = {
    "CAR MODEL": data.carModel || "",
    "CAR ID": data.carId || "",
    "Client Name": data.clientName || "",
    DESCRIPTION: data.description || "",
    "PAID BY": data.paidBy || "",
    AMOUNT: data.amount || "",
    "EXPENSE DATE": data.expenseDate || "",
  };

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    if (keyMap[header] !== undefined) {
      rowValues[i] = keyMap[header];
    } else {
      rowValues[i] = existingRow[i];
    }
  }

  sheet.getRange(row, 1, 1, headers.length).setValues([rowValues]);
  return "Success";
}

function getExpensesForCar(carId) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All expenses");
  if (!sheet) return [];

  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var carIdIndex = headers.indexOf("CAR ID");
  if (carIdIndex === -1) return [];

  var expenses = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][carIdIndex] === carId) {
      var expense = {};
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        expense[header] = data[i][j];
      }
      expenses.push(expense);
    }
  }

  return expenses;
}

function getPaymentsForCar(carId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payment");
  if (!sheet) return [];

  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var carIdIndex = headers.findIndex(function (h) {
    return String(h).toUpperCase() === "CAR ID";
  });
  if (carIdIndex === -1) return [];

  var payments = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][carIdIndex] === carId) {
      var payment = {};
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        payment[header] = data[i][j];
      }
      payments.push(payment);
    }
  }

  return payments;
}

function getTotalExpenseForCar(carId) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("All expenses");
  if (!sheet) return 0;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 0;

  var headers = data[0];
  var carIdIndex = headers.indexOf("CAR ID");
  var amountIndex = headers.indexOf("AMOUNT");

  if (carIdIndex === -1 || amountIndex === -1) return 0;

  var total = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][carIdIndex] === carId) {
      var amt = parseFloat(data[i][amountIndex]);
      if (!isNaN(amt)) {
        total += amt;
      }
    }
  }
  return total;
}

function getTotalPaymentForCar(carId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payment");
  if (!sheet) return 0;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 0;

  var headers = data[0];
  var carIdIndex = headers.findIndex(
    (h) => String(h).toUpperCase() === "CAR ID",
  );
  var amountIndex = headers.findIndex(
    (h) => String(h).toUpperCase() === "AMOUNT",
  );

  if (carIdIndex === -1 || amountIndex === -1) return 0;

  var total = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][carIdIndex] === carId) {
      var amt = parseFloat(
        String(data[i][amountIndex]).replace(/[^0-9.-]+/g, ""),
      );
      if (!isNaN(amt)) {
        total += amt;
      }
    }
  }
  return total;
}

function updateFinancingStatusOnPayment(carId) {
  if (!carId) return;
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  var headers = data[0];
  var carIdIndex = headers.indexOf("Car ID");
  var financingStatusIndex = headers.indexOf("Financing status");

  if (carIdIndex === -1 || financingStatusIndex === -1) return;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][carIdIndex]).trim() === String(carId).trim()) {
      var car = {};
      for (var j = 0; j < headers.length; j++) {
        car[headers[j]] = data[i][j];
      }

      var spRaw = car["SOLD PRICE"] ? car["SOLD PRICE"].toString().trim() : "";
      var spNum = parseFloat(spRaw.replace(/[^0-9.-]+/g, "")) || 0;

      var dpRaw = car["DOWN PAYMENT"]
        ? car["DOWN PAYMENT"].toString().trim()
        : "";
      var dpNum = parseFloat(dpRaw.replace(/[^0-9.-]+/g, "")) || 0;

      var loanAmount = spNum > 0 && dpNum > 0 ? spNum - dpNum : 0;
      if (loanAmount <= 0) return;

      var totalPayments = getTotalPaymentForCar(carId) || 0;
      var remainLoan = loanAmount - totalPayments;

      var currentFinStatus = (car["Financing status"] || "").toString().trim();
      var newFinStatus = currentFinStatus;

      if (remainLoan <= 0) {
        newFinStatus = "Paid off";
      } else {
        newFinStatus = "Loan Active";
      }

      if (currentFinStatus !== newFinStatus) {
        sheet.getRange(i + 1, financingStatusIndex + 1).setValue(newFinStatus);
      }
      break;
    }
  }
}

function syncAllFinancingStatuses() {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
  if (!sheet) return "Error: Vehicle details sheet not found";

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return "No data to sync";

  var headers = data[0];
  var carIdIndex = headers.indexOf("Car ID");
  var financingStatusIndex = headers.indexOf("Financing status");

  if (carIdIndex === -1 || financingStatusIndex === -1)
    return "Error: Required columns not found";

  var updatedCount = 0;

  for (var i = 1; i < data.length; i++) {
    var carId = data[i][carIdIndex];
    if (!carId) continue;

    var car = {};
    for (var j = 0; j < headers.length; j++) {
      car[headers[j]] = data[i][j];
    }

    var spRaw = car["SOLD PRICE"] ? car["SOLD PRICE"].toString().trim() : "";
    var spNum = parseFloat(spRaw.replace(/[^0-9.-]+/g, "")) || 0;

    var dpRaw = car["DOWN PAYMENT"]
      ? car["DOWN PAYMENT"].toString().trim()
      : "";
    var dpNum = parseFloat(dpRaw.replace(/[^0-9.-]+/g, "")) || 0;

    var loanAmount = spNum > 0 && dpNum > 0 ? spNum - dpNum : 0;
    if (loanAmount <= 0) continue;

    var totalPayments = getTotalPaymentForCar(carId) || 0;
    var remainLoan = loanAmount - totalPayments;

    var currentFinStatus = (car["Financing status"] || "").toString().trim();
    var newFinStatus = currentFinStatus;

    if (remainLoan <= 0) {
      newFinStatus = "Paid off";
    } else {
      newFinStatus = "Loan Active";
    }

    if (currentFinStatus !== newFinStatus) {
      sheet.getRange(i + 1, financingStatusIndex + 1).setValue(newFinStatus);
      updatedCount++;
    }
  }

  return "Success: Synced " + updatedCount + " vehicles.";
}

function savePayment(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payment");

  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Payment");
    sheet.appendRow([
      "Timestamp",
      "CAR MODEL",
      "CAR ID",
      "Client Name",
      "PAYMENT OPTION / NOTES",
      "AMOUNT",
      "PAYMENT DATE",
      "NOTES",
    ]);
  } else {
    // Check if NOTES header exists
    var lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
      if (headers.indexOf("NOTES") === -1 && headers.indexOf("Notes") === -1) {
        sheet.getRange(1, lastCol + 1).setValue("NOTES");
      }
    }
  }

  // Get current headers to dynamically build row values
  var headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  var rowValues = new Array(headers.length);

  var keyMap = {
    Timestamp: new Date(),
    "CAR MODEL": data.carModel || "",
    "CAR ID": data.carId || "",
    "Client Name": data.clientName || "",
    "CLIENT NAME": data.clientName || "",
    "PAYMENT OPTION / NOTES": data.paymentOption || "",
    "PAYMENT OPTION": data.paymentOption || "",
    "Payment Option": data.paymentOption || "",
    AMOUNT: data.amount || "",
    "PAYMENT DATE": data.paymentDate || "",
    NOTES: data.notes || "",
    Notes: data.notes || "",
  };

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    rowValues[i] = keyMap[header] !== undefined ? keyMap[header] : "";
  }

  sheet.appendRow(rowValues);
  if (data.carId) {
    updateFinancingStatusOnPayment(data.carId);
  }
  return "Success";
}

function getPayments() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payment");
  if (!sheet) return [];

  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; // Only headers or empty

  var headers = data[0];
  var payments = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var payment = { rowNumber: i + 1 };
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      payment[header] = row[j];
    }
    payments.push(payment);
  }

  return payments;
}

function deletePayment(rowNumber) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payment");
  if (!sheet) return "Error: Sheet not found";

  var row = parseInt(rowNumber);
  if (isNaN(row) || row <= 1) return "Error: Invalid row number";

  var headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  var carIdIndex = headers.indexOf("CAR ID");
  var carId = "";
  if (carIdIndex !== -1) {
    carId = sheet.getRange(row, carIdIndex + 1).getValue();
  }

  sheet.deleteRow(row);

  if (carId) {
    updateFinancingStatusOnPayment(carId);
  }

  return "Success";
}

function updatePayment(rowNumber, data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Payment");
  if (!sheet) return "Error: Sheet not found";

  var row = parseInt(rowNumber);
  if (isNaN(row) || row <= 1) return "Error: Invalid row number";

  var headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  var rowValues = new Array(headers.length);

  var existingRow = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

  var keyMap = {
    "CAR MODEL": data.carModel || "",
    "CAR ID": data.carId || "",
    "Client Name": data.clientName || "",
    "CLIENT NAME": data.clientName || "",
    "PAYMENT OPTION / NOTES": data.paymentOption || "",
    "PAYMENT OPTION": data.paymentOption || "",
    "Payment Option": data.paymentOption || "",
    AMOUNT: data.amount || "",
    "PAYMENT DATE": data.paymentDate || "",
    NOTES: data.notes || "",
    Notes: data.notes || "",
  };

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    if (keyMap[header] !== undefined) {
      rowValues[i] = keyMap[header];
    } else {
      rowValues[i] = existingRow[i];
    }
  }

  sheet.getRange(row, 1, 1, headers.length).setValues([rowValues]);
  var carId = data.carId;
  if (!carId) {
    var carIdIndex = headers.indexOf("CAR ID");
    if (carIdIndex !== -1) {
      carId = existingRow[carIdIndex];
    }
  }
  if (carId) {
    updateFinancingStatusOnPayment(carId);
  }
  return "Success";
}

function getCarListData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Car List");
  if (!sheet) return { names: [], models: {} };

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { names: [], models: {} }; // assuming header row

  var modelsByCar = {};

  for (var i = 1; i < data.length; i++) {
    var carName = data[i][0];
    var carModel = data[i][1];

    if (carName) {
      carName = String(carName).trim();
      carModel = String(carModel || "").trim();

      if (!modelsByCar[carName]) {
        modelsByCar[carName] = [];
      }
      if (carModel && modelsByCar[carName].indexOf(carModel) === -1) {
        modelsByCar[carName].push(carModel);
      }
    }
  }

  var carNames = Object.keys(modelsByCar).sort();

  return { names: carNames, models: modelsByCar };
}

function saveDataInput(data) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
  if (!sheet) return "Error: Sheet not found";

  var cName = data.carName ? data.carName.toString().trim() : "";
  var firstLetter = cName.length > 0 ? cName.charAt(0).toUpperCase() : "X";

  var cModel = data.model ? data.model.toString().trim() : "";
  var secondLetter = cModel.length > 0 ? cModel.charAt(0).toUpperCase() : "X";

  var cYear = data.year ? data.year.toString().trim() : "";
  var yearStr = cYear.length >= 2 ? cYear.slice(-2) : "00";

  var randomNum = Math.floor(Math.random() * 100);
  var randomStr = randomNum < 10 ? "0" + randomNum : randomNum.toString();

  var carId = firstLetter + secondLetter + yearStr + "-" + randomStr;

  // Ensure headers exist
  var currentHeaders = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  var requiredHeaders = [
    "Trade Value",
    "Trade In",
    "Comment",
    "Source Type",
    "Seller Name",
    "Acquisition Date",
    "Purchase Price",
    "Private Notes",
    "Purchased Car Name",
  ];
  requiredHeaders.forEach(function (h) {
    if (currentHeaders.indexOf(h) === -1) {
      sheet.getRange(1, currentHeaders.length + 1).setValue(h);
      currentHeaders.push(h);
    }
  });

  var rowValues = new Array(currentHeaders.length);

  // Set tradeStatus based on sourceType
  var tradeStatusVal =
    data.sourceType === "Trade" ? "Trade" : data.tradeStatus || "";
  var tradeValueVal =
    data.sourceType === "Trade"
      ? data.tradeValue || ""
      : data["Trade Value"] || "";
  var tradeInVal =
    data.sourceType === "Trade" ? data.tradeIn || "" : data["Trade In"] || "";
  var commentVal =
    data.sourceType === "Trade"
      ? data.tradeComment || ""
      : data["Comment"] || "";

  // Base map of columns to data input values
  var keyMap = {
    Timestamp: new Date(),
    "Car ID": carId,
    "Car Name": data.carName || "",
    Model: data.model || "",
    Year: data.year || "",
    Mileage: data.mileage || "",
    Price: "",
    Discount: "",
    VIN: data.vin || "",
    Status: "Available",
    Title: data.title || "",
    Trim: data.trim || "",
    "Body Style": data.bodyStyle || "",
    "Rent or Sell": "",
    Engine: "",
    "Engine Type/Size": "",
    Transmission: "",
    Driveline: "",
    "Fuel Type": data.fuelType || "",
    "Power Options": "",
    "Drive Condition": "",
    Condition: "",
    "Seat Material": "",
    "Interior Color": "",
    "Exterior Color": "",
    "Interior Features": "",
    "Main Image URLs": "",
    "Sub Image URLs": "",
    "CLIENT NAME": data.clientName || "",
    "PURCHASE DATE": data.purchaseDate || "",
    "SOLD DATE": "",
    "Trade status": tradeStatusVal,
    "IAAI TOTAL PRICE W/ FEES":
      data.sourceType === "IAAI" ? data.iaaiTotalPrice || "" : "",
    "PAPE PRICE": data.papePrice || "",
    "TRANSPORT FEES": data.transportFees || "",
    "SOLD PRICE": "",
    "DOWN PAYMENT": "",
    "CAR PICKUP LOCATION": data.pickupLocation || "",
    "DRIVER NAME": data.driverName || "",
    "DRIVER INFORMATION": data.driverInformation || "",
    NOTES: data.notes || "",
    "IAAI Price Before Fees":
      data.sourceType === "IAAI" ? data.iaaiPriceBeforeFees || "" : "",
    "Dispatcher Name": data.dispatcherName || "",
    "Dispatcher Price": data.dispatcherPrice || "",
    "Dispatcher Phone number": data.dispatcherPhone || "",
    "Client Phone": data.clientPhone || "",
    "Client Email": data.clientEmail || "",
    "Driver Phone": data.driverPhone || "",
    "Driver Company": data.driverCompany || "",
    "Tax Amount": data["Tax Amount"] || "",
    "Price on title": data["Price on title"] || "",
    "Financing status": data["Financing status"] || "",
    "Tax Responsibility": data["Tax Responsibility"] || "",
    "Trade Value": tradeValueVal,
    "Trade In": tradeInVal,
    Comment: commentVal,
    "Source Type": data.sourceType || "IAAI",
    "Seller Name":
      data.sourceType === "Private/Partner" ? data.sellerName || "" : "",
    "Acquisition Date":
      data.sourceType === "Private/Partner" ? data.acquisitionDate || "" : "",
    "Purchase Price":
      data.sourceType === "Private/Partner" ? data.purchasePrice || "" : "",
    "Private Notes":
      data.sourceType === "Private/Partner" ? data.privateNotes || "" : "",
    "Purchased Car Name":
      data.sourceType === "Trade" ? data.tradeCarName || "" : "",
  };

  for (var i = 0; i < currentHeaders.length; i++) {
    var header = currentHeaders[i];
    rowValues[i] = keyMap[header] !== undefined ? keyMap[header] : "";
  }

  sheet.appendRow(rowValues);

  // Handle Trade: apply down payment to purchased vehicle and mark as Sold
  if (
    data.sourceType === "Trade" &&
    data.tradeUseAsDownPayment === "true" &&
    data.tradeCarName
  ) {
    applyTradeDownPayment(data.tradeCarName, data.tradeValue);
  } else if (data.sourceType === "Trade" && data.tradeCarName) {
    // Even if not using as down payment, mark the purchased vehicle as Sold
    markCarAsSold(data.tradeCarName);
  }

  return "Success:" + carId;
}

function applyTradeDownPayment(purchasedCarId, tradeValue) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var carIdIndex = headers.indexOf("Car ID");
  var dpIndex = headers.indexOf("DOWN PAYMENT");
  var statusIndex = headers.indexOf("Status");

  if (carIdIndex === -1) return;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][carIdIndex]).trim() === String(purchasedCarId).trim()) {
      if (dpIndex !== -1) {
        sheet.getRange(i + 1, dpIndex + 1).setValue(tradeValue || "");
      }
      if (statusIndex !== -1) {
        sheet.getRange(i + 1, statusIndex + 1).setValue("Sold");
      }
      return;
    }
  }
}

function markCarAsSold(purchasedCarId) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var carIdIndex = headers.indexOf("Car ID");
  var statusIndex = headers.indexOf("Status");

  if (carIdIndex === -1) return;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][carIdIndex]).trim() === String(purchasedCarId).trim()) {
      if (statusIndex !== -1) {
        sheet.getRange(i + 1, statusIndex + 1).setValue("Sold");
      }
      return;
    }
  }
}

function getAvailableCarsForTrade(excludeCarId) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var headers = data[0];
  var carIdIndex = headers.indexOf("Car ID");
  var carNameIndex = headers.indexOf("Car Name");
  var modelIndex = headers.indexOf("Model");
  var yearIndex = headers.indexOf("Year");
  var statusIndex = headers.indexOf("Status");
  var purchasePriceIndex = headers.indexOf("Purchase Price");

  if (carIdIndex === -1) return [];

  var cars = [];
  for (var i = 1; i < data.length; i++) {
    var status =
      statusIndex !== -1 ? String(data[i][statusIndex]).trim() : "Available";
    var carId = String(data[i][carIdIndex]).trim();
    if (status === "Available" && carId !== String(excludeCarId || "").trim()) {
      var carName =
        carNameIndex !== -1 ? String(data[i][carNameIndex]).trim() : "";
      var model = modelIndex !== -1 ? String(data[i][modelIndex]).trim() : "";
      var year = yearIndex !== -1 ? String(data[i][yearIndex]).trim() : "";
      cars.push({
        carId: carId,
        displayName:
          (carId ? "#" + carId + " - " : "") +
          (year ? year + " " : "") +
          (carName ? carName + " " : "") +
          (model || ""),
        carName: carName,
        model: model,
        year: year,
        purchasePrice:
          purchasePriceIndex !== -1 ? data[i][purchasePriceIndex] : "",
      });
    }
  }

  return cars;
}

function getCarListItems() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Car List");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Car List");
    sheet.appendRow([
      "Car Name",
      "Model",
      "Trim",
      "Fuel Type",
      "Body Type",
      "Years Sold",
    ]);
    return [];
  }

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; // only headers or empty

  var items = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    items.push({
      rowNumber: i + 1, // row number in spreadsheet (1-indexed, starts at 2 for first data row)
      carName: row[0] ? String(row[0]).trim() : "",
      model: row[1] ? String(row[1]).trim() : "",
      trim: row[2] ? String(row[2]).trim() : "",
      fuelType: row[3] ? String(row[3]).trim() : "",
      bodyType: row[4] ? String(row[4]).trim() : "",
      yearsSold: row[5] ? String(row[5]).trim() : "",
    });
  }
  return items;
}

function normalizeCommaSeparatedField(value) {
  if (value === null || value === undefined) return "";

  var parts = Array.isArray(value) ? value : String(value).split(",");

  var cleaned = [];
  for (var i = 0; i < parts.length; i++) {
    var part = String(parts[i]).trim();
    if (part && cleaned.indexOf(part) === -1) {
      cleaned.push(part);
    }
  }

  return cleaned.join(", ");
}

function saveCarListItem(itemData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Car List");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Car List");
    sheet.appendRow([
      "Car Name",
      "Model",
      "Trim",
      "Fuel Type",
      "Body Type",
      "Years Sold",
    ]);
  }

  sheet.appendRow([
    itemData.carName ? String(itemData.carName).trim() : "",
    itemData.model ? String(itemData.model).trim() : "",
    itemData.trim ? String(itemData.trim).trim() : "",
    normalizeCommaSeparatedField(itemData.fuelType),
    normalizeCommaSeparatedField(itemData.bodyType),
    itemData.yearsSold ? String(itemData.yearsSold).trim() : "",
  ]);

  return "Success";
}

function updateCarListItem(rowNumber, itemData) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Car List");
  if (!sheet) return "Error: Sheet not found";

  var row = parseInt(rowNumber);
  if (isNaN(row) || row <= 1) return "Error: Invalid row number";

  sheet
    .getRange(row, 1, 1, 6)
    .setValues([
      [
        itemData.carName ? String(itemData.carName).trim() : "",
        itemData.model ? String(itemData.model).trim() : "",
        itemData.trim ? String(itemData.trim).trim() : "",
        normalizeCommaSeparatedField(itemData.fuelType),
        normalizeCommaSeparatedField(itemData.bodyType),
        itemData.yearsSold ? String(itemData.yearsSold).trim() : "",
      ],
    ]);

  return "Success";
}

function deleteCarListItem(rowNumber) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Car List");
  if (!sheet) return "Error: Sheet not found";

  var row = parseInt(rowNumber);
  if (isNaN(row) || row <= 1) return "Error: Invalid row number";

  sheet.deleteRow(row);
  return "Success";
}

function deleteVehicle(carId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Delete from "Vehicle details" (usually only 1 row)
  var vehicleSheet = ss.getSheetByName("Vehicle details");
  var vehicleDeleted = false;
  if (vehicleSheet) {
    var vData = vehicleSheet.getDataRange().getValues();
    if (vData.length > 1) {
      var vHeaders = vData[0];
      var vCarIdIdx = vHeaders.indexOf("Car ID");
      if (vCarIdIdx !== -1) {
        for (var i = vData.length - 1; i >= 1; i--) {
          if (vData[i][vCarIdIdx] === carId) {
            vehicleSheet.deleteRow(i + 1);
            vehicleDeleted = true;
          }
        }
      }
    }
  }

  // 2. Delete from "All expenses" (multiple rows possible)
  var expenseSheet = ss.getSheetByName("All expenses");
  if (expenseSheet) {
    var eData = expenseSheet.getDataRange().getValues();
    if (eData.length > 1) {
      var eHeaders = eData[0];
      var eCarIdIdx = eHeaders.indexOf("CAR ID");
      if (eCarIdIdx !== -1) {
        for (var j = eData.length - 1; j >= 1; j--) {
          if (eData[j][eCarIdIdx] === carId) {
            expenseSheet.deleteRow(j + 1);
          }
        }
      }
    }
  }

  // 3. Delete from "Payment" (multiple rows possible)
  var paymentSheet = ss.getSheetByName("Payment");
  if (paymentSheet) {
    var pData = paymentSheet.getDataRange().getValues();
    if (pData.length > 1) {
      var pHeaders = pData[0];
      var pCarIdIdx = pHeaders.findIndex(function (h) {
        return String(h).toUpperCase() === "CAR ID";
      });
      if (pCarIdIdx !== -1) {
        for (var k = pData.length - 1; k >= 1; k--) {
          if (pData[k][pCarIdIdx] === carId) {
            paymentSheet.deleteRow(k + 1);
          }
        }
      }
    }
  }

  return vehicleDeleted ? "Success" : "Error: Car ID not found";
}

function saveWfmExpense(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "WFM Business Expenses",
  );

  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(
      "WFM Business Expenses",
    );
    sheet.appendRow([
      "Timestamp",
      "DETAILS",
      "VALUE BEFORE TAX",
      "VALUE AFTER TAX",
      "DATE",
      "PAID BY",
    ]);
  }

  // Get current headers to dynamically build row values
  var headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  var rowValues = new Array(headers.length);

  var keyMap = {
    Timestamp: new Date(),
    DETAILS: data.details || "",
    "VALUE BEFORE TAX": data.valueBeforeTax || "",
    "VALUE AFTER TAX": data.valueAfterTax || "",
    DATE: data.date || "",
    "PAID BY": data.paidBy || "",
  };

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    rowValues[i] = keyMap[header] !== undefined ? keyMap[header] : "";
  }

  sheet.appendRow(rowValues);
  return "Success";
}

function getWfmExpenses() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "WFM Business Expenses",
  );
  if (!sheet) return [];

  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; // Only headers or empty

  var headers = data[0];
  var wfmExpenses = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var expense = { rowNumber: i + 1 };
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      expense[header] = row[j];
    }
    wfmExpenses.push(expense);
  }

  return wfmExpenses;
}

function deleteWfmExpense(rowNumber) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "WFM Business Expenses",
  );
  if (!sheet) return "Error: Sheet not found";

  var row = parseInt(rowNumber);
  if (isNaN(row) || row <= 1) return "Error: Invalid row number";

  sheet.deleteRow(row);
  return "Success";
}

function updateWfmExpense(rowNumber, data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "WFM Business Expenses",
  );
  if (!sheet) return "Error: Sheet not found";

  var row = parseInt(rowNumber);
  if (isNaN(row) || row <= 1) return "Error: Invalid row number";

  var headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  var rowValues = new Array(headers.length);

  var existingRow = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

  var keyMap = {
    DETAILS: data.details || "",
    "VALUE BEFORE TAX": data.valueBeforeTax || "",
    "VALUE AFTER TAX": data.valueAfterTax || "",
    DATE: data.date || "",
    "PAID BY": data.paidBy || "",
  };

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    if (keyMap[header] !== undefined) {
      rowValues[i] = keyMap[header];
    } else {
      rowValues[i] = existingRow[i];
    }
  }

  sheet.getRange(row, 1, 1, headers.length).setValues([rowValues]);
  return "Success";
}

function saveOtherExpense(data) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Other Expenses");

  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Other Expenses");
    sheet.appendRow([
      "Timestamp",
      "DETAILS",
      "VALUE BEFORE TAX",
      "VALUE AFTER TAX",
      "DATE",
      "PAID BY",
    ]);
  }

  // Get current headers to dynamically build row values
  var headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  var rowValues = new Array(headers.length);

  var keyMap = {
    Timestamp: new Date(),
    DETAILS: data.details || "",
    "VALUE BEFORE TAX": data.valueBeforeTax || "",
    "VALUE AFTER TAX": data.valueAfterTax || "",
    DATE: data.date || "",
    "PAID BY": data.paidBy || "",
  };

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    rowValues[i] = keyMap[header] !== undefined ? keyMap[header] : "";
  }

  sheet.appendRow(rowValues);
  return "Success";
}

function getOtherExpenses() {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Other Expenses");
  if (!sheet) return [];

  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; // Only headers or empty

  var headers = data[0];
  var otherExpenses = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var expense = { rowNumber: i + 1 };
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      expense[header] = row[j];
    }
    otherExpenses.push(expense);
  }

  return otherExpenses;
}

function deleteOtherExpense(rowNumber) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Other Expenses");
  if (!sheet) return "Error: Sheet not found";

  var row = parseInt(rowNumber);
  if (isNaN(row) || row <= 1) return "Error: Invalid row number";

  sheet.deleteRow(row);
  return "Success";
}

function updateOtherExpense(rowNumber, data) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Other Expenses");
  if (!sheet) return "Error: Sheet not found";

  var row = parseInt(rowNumber);
  if (isNaN(row) || row <= 1) return "Error: Invalid row number";

  var headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  var rowValues = new Array(headers.length);

  var existingRow = sheet.getRange(row, 1, 1, headers.length).getValues()[0];

  var keyMap = {
    DETAILS: data.details || "",
    "VALUE BEFORE TAX": data.valueBeforeTax || "",
    "VALUE AFTER TAX": data.valueAfterTax || "",
    DATE: data.date || "",
    "PAID BY": data.paidBy || "",
  };

  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    if (keyMap[header] !== undefined) {
      rowValues[i] = keyMap[header];
    } else {
      rowValues[i] = existingRow[i];
    }
  }

  sheet.getRange(row, 1, 1, headers.length).setValues([rowValues]);
  return "Success";
}

function sendAgreementEmailToServer(carId) {
  try {
    var vehicles = getVehicles();
    var car = null;
    for (var i = 0; i < vehicles.length; i++) {
      if (vehicles[i]["Car ID"] === carId) {
        car = vehicles[i];
        break;
      }
    }

    if (!car) {
      return { success: false, message: "Vehicle " + carId + " not found." };
    }

    var clientName = car["CLIENT NAME"]
      ? car["CLIENT NAME"].toString().trim()
      : "";
    var clientEmail = car["Client Email"]
      ? car["Client Email"].toString().trim()
      : "";
    var carName = car["Car Name"] || "";
    var model = car["Model"] || "";
    var year = car["Year"] || "";

    if (!clientEmail) {
      return {
        success: false,
        message: "Client email is missing for this vehicle.",
      };
    }

    var scriptUrl = "";
    try {
      scriptUrl = ScriptApp.getService().getUrl();
    } catch (e) {
      Logger.log("Error getting script URL: " + e);
    }
    if (!scriptUrl) {
      scriptUrl = "https://script.google.com/macros/s/AKfycbz_placeholder/exec";
    }
    var signLink = scriptUrl + "?id=" + carId;

    // Creative, premium HTML Email Template
    var htmlBody =
      "<div style=\"font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e1e1e1; border-radius: 12px; background-color: #faf9fd;\">" +
      '<div style="text-align: center; background: linear-gradient(135deg, #170a3d 0%, #3a1f62 100%); padding: 30px; border-radius: 10px 10px 0 0; color: #ffffff;">' +
      '<h1 style="margin: 0; font-size: 24px; font-weight: bold; letter-spacing: 1px;">WRIGHT FINDER MOTORS</h1>' +
      '<p style="margin: 10px 0 0 0; font-size: 14px; opacity: 0.9;">Secure Financing & Agreement Center</p>' +
      "</div>" +
      '<div style="padding: 30px 20px; background-color: #ffffff; border-radius: 0 0 10px 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.02);">' +
      '<h2 style="color: #3a1f62; margin-top: 0; font-size: 20px;">Purchase Agreement Ready for Signature</h2>' +
      '<p style="color: #4a4a4a; font-size: 15px; line-height: 1.6;">Dear ' +
      clientName +
      ",</p>" +
      '<p style="color: #4a4a4a; font-size: 15px; line-height: 1.6;">Thank you for purchasing your vehicle through Wright Finder Motors. We are pleased to inform you that your purchasing agreement for the vehicle listed below is ready to sign.</p>' +
      '<div style="background-color: #f6f4fa; border-left: 4px solid #3a1f62; padding: 15px; margin: 20px 0; border-radius: 4px;">' +
      '<table style="width: 100%; font-size: 14px; border-collapse: collapse; color: #4a4a4a;">' +
      '<tr><td style="padding: 5px 0; font-weight: bold; width: 40%;">Vehicle:</td><td style="padding: 5px 0;">' +
      year +
      " " +
      carName +
      " " +
      model +
      "</td></tr>" +
      '<tr><td style="padding: 5px 0; font-weight: bold;">Stock Number:</td><td style="padding: 5px 0;">' +
      carId +
      "</td></tr>" +
      "</table>" +
      "</div>" +
      '<p style="color: #4a4a4a; font-size: 15px; line-height: 1.6;">Please review the agreement and execute the signature using the link below to finalize the purchasing process:</p>' +
      '<div style="text-align: center; margin: 30px 0;">' +
      '<a href="' +
      signLink +
      '" style="background: linear-gradient(135deg, #3a1f62 0%, #170a3d 100%); color: #ffffff; padding: 12px 30px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 16px; box-shadow: 0 4px 10px rgba(58, 31, 98, 0.3); display: inline-block;">Review & Sign Agreement</a>' +
      "</div>" +
      '<hr style="border: 0; border-top: 1px solid #eeeeee; margin: 30px 0;">' +
      '<p style="color: #888888; font-size: 12px; line-height: 1.5; text-align: center;">If you have any questions or require assistance, please feel free to reach out to our support team.</p>' +
      '<p style="color: #888888; font-size: 12px; text-align: center; margin: 5px 0 0 0;">&copy; ' +
      new Date().getFullYear() +
      " Wright Finder Motors. All rights reserved.</p>" +
      "</div>" +
      "</div>";

    MailApp.sendEmail({
      to: clientEmail,
      subject:
        "Action Required: Sign Your Vehicle Purchase Agreement - " +
        year +
        " " +
        carName +
        " " +
        model,
      htmlBody: htmlBody,
    });

    updateAgreementStatus(carId, "Pending");

    return {
      success: true,
      message:
        "Purchase agreement email sent successfully to " +
        clientName +
        " (" +
        clientEmail +
        ").",
    };
  } catch (e) {
    return { success: false, message: "Failed to send email: " + e.toString() };
  }
}

/**
 * Run this function from the Apps Script editor toolbar (select 'getPermission' and click 'Run')
 * to trigger the authorization prompt for sending emails.
 */
function getPermission() {
  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: "Permission Verification",
    body: "If you are reading this, the MailApp permission has been successfully granted.",
  });
}

function getVehicleById(carId) {
  var vehicles = getVehicles();
  for (var i = 0; i < vehicles.length; i++) {
    if (vehicles[i]["Car ID"] === carId) {
      return vehicles[i];
    }
  }
  return null;
}

function ensureAgreementColumns() {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
  if (!sheet) return;
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var required = [
    "Agreement Status",
    "Installment Frequency",
    "Installment End Date",
    "Agreement PDF Link",
  ];

  for (var i = 0; i < required.length; i++) {
    if (headers.indexOf(required[i]) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(required[i]);
      headers.push(required[i]); // update headers array
    }
  }
}

function getAgreementData(carId, ignoreSignedCheck) {
  var car = getVehicleById(carId);
  if (!car) return null;

  if (!ignoreSignedCheck && car["Agreement Status"] === "Signed") {
    return { alreadySigned: true };
  }

  var soldPriceVal =
    parseFloat(
      (car["SOLD PRICE"] || "").toString().replace(/[^0-9.-]+/g, ""),
    ) || 0;
  var downPaymentVal =
    parseFloat(
      (car["DOWN PAYMENT"] || "").toString().replace(/[^0-9.-]+/g, ""),
    ) || 0;
  var remainLoanVal = soldPriceVal - downPaymentVal;

  var formatMoney = function (num) {
    return (
      "$" +
      num.toLocaleString("en-US", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      })
    );
  };

  var today = new Date();
  var options = { year: "numeric", month: "long", day: "numeric" };
  var formattedDate = today.toLocaleDateString("en-US", options);

  var days = [
    "Sunday",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
  ];
  var todayDayName = days[today.getDay()];

  return {
    car: car,
    todayDate: formattedDate,
    todayDayName: todayDayName,
    soldPrice: formatMoney(soldPriceVal),
    downPayment: formatMoney(downPaymentVal),
    remainLoan: formatMoney(remainLoanVal),
    installmentEndDateFormatted: formatDateString(car["Installment End Date"]),
  };
}

function updateAgreementStatus(carId, status) {
  try {
    var sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var carIdIndex = headers.indexOf("Car ID");
    var statusIndex = headers.indexOf("Agreement Status");

    if (carIdIndex !== -1 && statusIndex !== -1) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][carIdIndex] === carId) {
          sheet.getRange(i + 1, statusIndex + 1).setValue(status);
          break;
        }
      }
    }
  } catch (e) {
    Logger.log("Error updating agreement status: " + e);
  }
}

function submitSignedAgreement(carId, data) {
  try {
    var sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
    if (!sheet)
      return { success: false, message: "Sheet 'Vehicle details' not found" };

    var values = sheet.getDataRange().getValues();
    var headers = values[0];

    var carIdIndex = headers.indexOf("Car ID");
    if (carIdIndex === -1)
      return { success: false, message: "Car ID column not found" };

    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][carIdIndex] === carId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1)
      return { success: false, message: "Vehicle " + carId + " not found" };

    // Generate and save the agreement PDF in Google Drive
    var pdfUrl = generateAndSaveAgreementPdf(carId, data);

    var updates = {
      "Agreement Status": "Signed",
      "Agreement PDF Link": pdfUrl,
    };

    for (var key in updates) {
      var colIndex = headers.indexOf(key);
      if (colIndex !== -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(updates[key]);
      }
    }

    return {
      success: true,
      message: "Agreement successfully signed and verified!",
    };
  } catch (e) {
    return {
      success: false,
      message: "Error signing agreement: " + e.toString(),
    };
  }
}

function generateAndSaveAgreementPdf(carId, data) {
  var agreementData = getAgreementData(carId, true);
  if (!agreementData) {
    throw new Error("Agreement data not found for Car ID: " + carId);
  }

  var car = agreementData.car;
  var clientName = (car["CLIENT NAME"] || "").toString().trim() || "Client";
  var make = (car["Car Name"] || "").toString().trim();
  var model = (car["Model"] || "").toString().trim();
  var year = (car["Year"] || "").toString().trim();

  // Create template from the HTML file
  var template = HtmlService.createTemplateFromFile("AgreementPdfTemplate");

  // Bind template variables
  template.car = car;
  template.todayDate = agreementData.todayDate;
  template.todayDayName = agreementData.todayDayName;
  template.soldPrice = agreementData.soldPrice;
  template.downPayment = agreementData.downPayment;
  template.remainLoan = agreementData.remainLoan;

  template.installmentFrequency = car["Installment Frequency"] || "weekly";
  template.installmentEndDate = formatDateString(car["Installment End Date"]);

  template.buyerName = data.buyerName || clientName;
  template.buyerSignature = data.buyerSignature || "";
  template.witnessName = data.witnessName || "";
  template.witnessSignature = data.witnessSignature || "";

  var now = new Date();
  var signatureDateTime = "";
  try {
    signatureDateTime = Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      "MMMM d, yyyy h:mm a",
    );
  } catch (e) {
    var options = {
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    };
    signatureDateTime = now.toLocaleString("en-US", options);
  }
  template.signatureDateTime = signatureDateTime;

  var htmlContent = template.evaluate().getContent();

  // Get target Google Drive folder
  var folderId = "1Fwz36bo6jnGhC072SxTO3KaoeSN7jmMp";
  var folder = DriveApp.getFolderById(folderId);

  // Convert HTML to PDF by saving a temporary HTML file and converting it
  var tempFile = DriveApp.createFile(
    "temp_agreement_" + carId + ".html",
    htmlContent,
    "text/html",
  );
  var pdfBlob = tempFile.getAs("application/pdf");

  // Rename PDF to: Client name + Make + Model + Year
  var fileName =
    [clientName, make, model, year].filter(Boolean).join(" ") + ".pdf";
  pdfBlob.setName(fileName);

  // Save the PDF file in Google Drive folder
  var pdfFile = folder.createFile(pdfBlob);

  // Delete the temporary file
  tempFile.setTrashed(true);

  // Share file publicly so it can be viewed by anyone with the link
  try {
    pdfFile.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW,
    );
  } catch (e) {
    Logger.log("Could not set sharing permissions on PDF: " + e);
  }

  return pdfFile.getUrl();
}

function formatDateString(dateStr) {
  if (!dateStr) return "";
  try {
    // Check if it's already a Date object
    if (dateStr instanceof Date) {
      var options = { year: "numeric", month: "long", day: "numeric" };
      return dateStr.toLocaleDateString("en-US", options);
    }
    // Check for string date format YYYY-MM-DD
    var parts = dateStr.toString().split("-");
    if (parts.length === 3) {
      // Parts: [YYYY, MM, DD]
      var d = new Date(parts[0], parts[1] - 1, parts[2]);
      var options = { year: "numeric", month: "long", day: "numeric" };
      return d.toLocaleDateString("en-US", options);
    }
    // If it's a standard Date string representation from Apps Script / Google Sheets
    var d2 = new Date(dateStr);
    if (!isNaN(d2.getTime())) {
      var options = { year: "numeric", month: "long", day: "numeric" };
      return d2.toLocaleDateString("en-US", options);
    }
  } catch (e) {}
  return dateStr.toString();
}

function saveInstallmentDetails(carId, frequency, endDate) {
  try {
    ensureAgreementColumns();
    var sheet =
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vehicle details");
    if (!sheet)
      return { success: false, message: "Sheet 'Vehicle details' not found" };

    var values = sheet.getDataRange().getValues();
    var headers = values[0];
    var carIdIndex = headers.indexOf("Car ID");
    if (carIdIndex === -1)
      return { success: false, message: "Car ID column not found" };

    var rowIndex = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][carIdIndex] === carId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1)
      return { success: false, message: "Vehicle " + carId + " not found" };

    var updates = {
      "Installment Frequency": frequency,
      "Installment End Date": endDate,
    };

    for (var key in updates) {
      var colIndex = headers.indexOf(key);
      if (colIndex !== -1) {
        sheet.getRange(rowIndex, colIndex + 1).setValue(updates[key]);
      }
    }

    return {
      success: true,
      message: "Installment terms updated successfully!",
    };
  } catch (e) {
    return { success: false, message: "Error saving details: " + e.toString() };
  }
}

function getDocumentRequestFolder() {
  var folderId = "15-IEb1JvHeAoP_L_aQOudIpieGf2j3Io";
  try {
    return DriveApp.getFolderById(folderId);
  } catch (e) {
    Logger.log("Failed to get Google Drive folder " + folderId + ": " + e);
    return DriveApp.getRootFolder();
  }
}

function getDocumentRequestsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Document Requests");
  if (!sheet) {
    sheet = ss.insertSheet("Document Requests");
    sheet.appendRow([
      "Timestamp",
      "Request ID",
      "Car ID",
      "Client Email",
      "Subject",
      "Message",
      "Sender Attachments",
      "Status",
      "Uploaded Documents",
    ]);
    sheet
      .getRange("A1:I1")
      .setFontWeight("bold")
      .setBackground("#f3f4f6")
      .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  } else {
    // Enforce 9 columns if they don't exist
    if (sheet.getLastColumn() < 9) {
      var headers = [
        "Timestamp",
        "Request ID",
        "Car ID",
        "Client Email",
        "Subject",
        "Message",
        "Sender Attachments",
        "Status",
        "Uploaded Documents",
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sheet;
}

function getDocumentRequestDetails(requestId) {
  try {
    var sheet = getDocumentRequestsSheet();
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (row[1] === requestId) {
        var adminFilesList = [];
        var adminFilesStr = row[6] ? row[6].toString() : "";
        if (adminFilesStr) {
          var links = adminFilesStr.split("\n");
          for (var j = 0; j < links.length; j++) {
            if (links[j]) {
              var parts = links[j].split(" -> ");
              adminFilesList.push({
                name: parts[0] || "Attachment",
                url: parts[1] || links[j],
              });
            }
          }
        }
        return {
          requestId: row[1] ? row[1].toString() : "",
          carId: row[2] ? row[2].toString() : "",
          clientEmail: row[3] ? row[3].toString() : "",
          subject: row[4] ? row[4].toString() : "",
          message: row[5] ? row[5].toString() : "",
          adminFiles: adminFilesList,
          status: row[7] ? row[7].toString() : "",
        };
      }
    }
  } catch (e) {
    Logger.log("Error in getDocumentRequestDetails: " + e);
  }
  return null;
}

function getDocumentRequests() {
  try {
    var sheet = getDocumentRequestsSheet();
    var values = sheet.getDataRange().getValues();
    var results = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      // Convert timestamp to string if it's a Date
      var ts = row[0];
      if (ts instanceof Date) {
        ts = ts.toISOString();
      } else if (ts) {
        ts = ts.toString();
      } else {
        ts = "";
      }
      results.push({
        timestamp: ts,
        requestId: row[1] ? row[1].toString() : "",
        carId: row[2] ? row[2].toString() : "",
        clientEmail: row[3] ? row[3].toString() : "",
        subject: row[4] ? row[4].toString() : "",
        message: row[5] ? row[5].toString() : "",
        senderAttachments: row[6] ? row[6].toString() : "",
        status: row[7] ? row[7].toString() : "",
        uploadedDocuments: row.length > 8 && row[8] ? row[8].toString() : "",
      });
    }
    return results;
  } catch (e) {
    Logger.log("Error in getDocumentRequests: " + e);
    return [];
  }
}

function sendClientDocumentRequestEmail(
  email,
  subject,
  message,
  adminFiles,
  carId,
) {
  try {
    var folder = getDocumentRequestFolder();
    var adminFileLinks = [];
    var mailAttachments = [];

    if (adminFiles && adminFiles.length > 0) {
      for (var i = 0; i < adminFiles.length; i++) {
        var fileData = adminFiles[i];
        var blob = Utilities.newBlob(
          Utilities.base64Decode(fileData.data),
          fileData.mimeType,
          fileData.name,
        );
        var file = folder.createFile(blob);
        try {
          file.setSharing(
            DriveApp.Access.ANYONE_WITH_LINK,
            DriveApp.Permission.VIEW,
          );
        } catch (shareErr) {}

        adminFileLinks.push(fileData.name + " -> " + file.getUrl());
        mailAttachments.push(blob);
      }
    }

    var requestId =
      "req_" + Date.now() + "_" + Math.floor(Math.random() * 1000);

    var scriptUrl = "";
    try {
      scriptUrl = ScriptApp.getService().getUrl();
    } catch (e) {
      Logger.log("Error getting script URL: " + e);
    }
    if (!scriptUrl) {
      scriptUrl = "https://script.google.com/macros/s/AKfycbz_placeholder/exec";
    }

    var uploadLink = scriptUrl + "?requestId=" + requestId;

    var htmlBody =
      "<div style=\"font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e5e7eb; border-radius: 16px; background-color: #faf9fd;\">" +
      '<div style="text-align: center; background: linear-gradient(135deg, #2c1a4d 0%, #512b81 100%); padding: 30px; border-radius: 12px 12px 0 0; color: #ffffff;">' +
      '<h1 style="margin: 0; font-size: 24px; font-weight: 800; letter-spacing: 0.5px;">WRIGHT FINDER MOTORS</h1>' +
      '<p style="margin: 8px 0 0 0; font-size: 14px; opacity: 0.9; font-weight: 400;">Secure Document Portal</p>' +
      "</div>" +
      '<div style="padding: 30px 20px; background-color: #ffffff; border-radius: 0 0 12px 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.02); color: #1f2937;">' +
      '<h2 style="color: #2c1a4d; margin-top: 0; font-size: 18px; font-weight: 700; border-bottom: 2px solid #f3f4f6; padding-bottom: 10px;">Document Request: ' +
      subject +
      "</h2>" +
      '<p style="font-size: 15px; line-height: 1.6; margin-top: 15px;">Dear Client,</p>' +
      '<p style="font-size: 15px; line-height: 1.6;">The team at Wright Finder Motors requires one or more documents to complete your process. Please review the request and submit your files securely using the portal link below:</p>' +
      '<div style="background-color: #f5f3f7; border-left: 4px solid #512b81; padding: 15px; margin: 20px 0; border-radius: 0 8px 8px 0; font-size: 14px; line-height: 1.5; color: #4b5563;">' +
      "<strong>Request details / instructions:</strong><br/>" +
      message.replace(/\n/g, "<br/>") +
      "</div>";

    if (adminFileLinks.length > 0) {
      htmlBody +=
        '<div style="margin: 20px 0;">' +
        '<strong style="font-size: 14px;">Reference documents provided by sender:</strong>' +
        '<ul style="padding-left: 20px; font-size: 14px; margin-top: 5px; color: #512b81;">';
      for (var k = 0; k < adminFileLinks.length; k++) {
        var parts = adminFileLinks[k].split(" -> ");
        htmlBody +=
          '<li><a href="' +
          parts[1] +
          '" style="color: #512b81; font-weight: 600;">' +
          parts[0] +
          "</a></li>";
      }
      htmlBody += "</ul>" + "</div>";
    }

    htmlBody +=
      '<div style="text-align: center; margin: 35px 0;">' +
      '<a href="' +
      uploadLink +
      '" style="background: linear-gradient(135deg, #512b81 0%, #2c1a4d 100%); color: #ffffff; padding: 14px 35px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 16px; box-shadow: 0 6px 15px rgba(81, 43, 129, 0.25); display: inline-block;">Upload Requested Document</a>' +
      "</div>" +
      '<hr style="border: 0; border-top: 1px solid #f3f4f6; margin: 30px 0;">' +
      '<p style="color: #9ca3af; font-size: 12px; line-height: 1.5; text-align: center;">This is a secure connection. Do not share your upload link with anyone.</p>' +
      '<p style="color: #9ca3af; font-size: 12px; text-align: center; margin: 5px 0 0 0;">&copy; ' +
      new Date().getFullYear() +
      " Wright Finder Motors. All rights reserved.</p>" +
      "</div>" +
      "</div>";

    var emailOptions = {
      to: email,
      subject:
        "Required Action: Document request from Wright Finder Motors - " +
        subject,
      htmlBody: htmlBody,
    };

    if (mailAttachments.length > 0) {
      emailOptions.attachments = mailAttachments;
    }

    MailApp.sendEmail(emailOptions);

    var sheet = getDocumentRequestsSheet();
    sheet.appendRow([
      new Date(),
      requestId,
      carId || "N/A",
      email,
      subject,
      message,
      adminFileLinks.join("\n"),
      "Pending",
      "",
    ]);

    return {
      success: true,
      message: "Document request email successfully sent to " + email,
    };
  } catch (e) {
    Logger.log("Error sending document request: " + e);
    return {
      success: false,
      message: "Error sending request: " + e.toString(),
    };
  }
}

function uploadClientDocument(requestId, filesData) {
  try {
    var sheet = getDocumentRequestsSheet();
    var values = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < values.length; i++) {
      if (values[i][1] === requestId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: "Request ID not found." };
    }

    var folder = getDocumentRequestFolder();
    var uploadedLinks = [];

    for (var j = 0; j < filesData.length; j++) {
      var fileInfo = filesData[j];
      var blob = Utilities.newBlob(
        Utilities.base64Decode(fileInfo.data),
        fileInfo.mimeType,
        fileInfo.name,
      );
      var file = folder.createFile(blob);
      try {
        file.setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW,
        );
      } catch (e) {}

      uploadedLinks.push(fileInfo.name + " -> " + file.getUrl());
    }

    sheet.getRange(rowIndex, 8).setValue("Uploaded");

    var existingUploads = sheet.getRange(rowIndex, 9).getValue();
    var allUploads = existingUploads
      ? existingUploads + "\n" + uploadedLinks.join("\n")
      : uploadedLinks.join("\n");
    sheet.getRange(rowIndex, 9).setValue(allUploads);

    return { success: true, message: "Documents uploaded successfully." };
  } catch (e) {
    Logger.log("Error in uploadClientDocument: " + e);
    return {
      success: false,
      message: "Failed to upload files: " + e.toString(),
    };
  }
}

/**
 * Sends a reminder email to all clients who have active loans.
 * Runs on a 2-minute time-driven trigger.
 */
function sendActiveLoanPaymentReminders() {
  try {
    // Workaround for Google Apps Script not allowing everyMinutes(2)
    // Run every minute but execute logic only on even minutes (every 2 minutes)
    var currentMinute = new Date().getMinutes();
    if (currentMinute % 2 !== 0) {
      Logger.log(
        "Skipping this minute to maintain 2-minute interval: " + currentMinute,
      );
      return;
    }

    var vehicles = getVehicles();
    if (!vehicles || vehicles.length === 0) return;

    for (var i = 0; i < vehicles.length; i++) {
      var car = vehicles[i];
      var finStatus = (car["Financing status"] || "")
        .toString()
        .trim()
        .toLowerCase();
      var status = (car["Status"] || "").toString().trim().toLowerCase();
      var clientEmail = car["Client Email"]
        ? car["Client Email"].toString().trim()
        : "";
      var clientName =
        car["CLIENT NAME"] || car["Buyer Name"] || "Valued Customer";

      var spRaw = car["SOLD PRICE"] ? car["SOLD PRICE"].toString().trim() : "";
      var spNum = parseFloat(spRaw.replace(/[^0-9.-]+/g, "")) || 0;
      var dpRaw = car["DOWN PAYMENT"]
        ? car["DOWN PAYMENT"].toString().trim()
        : "";
      var dpNum = parseFloat(dpRaw.replace(/[^0-9.-]+/g, "")) || 0;
      var isSold = status.includes("sold");
      var isForceActiveLoan = isSold && spNum > 0 && dpNum === 0;

      if ((finStatus === "loan active" || isForceActiveLoan) && clientEmail) {
        var loanAmountRaw = car["LOAN AMOUNT"]
          ? car["LOAN AMOUNT"].toString().trim()
          : "";

        var formatMoney = function (num) {
          return (
            "$" +
            num.toLocaleString("en-US", {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            })
          );
        };

        // Fetch total payments to calculate REMAIN LOAN like in Key Specifications
        var totalPayments = getTotalPaymentForCar(car["Car ID"]) || 0;

        var loanAmountVal = spNum > 0 && dpNum > 0 ? spNum - dpNum : 0;
        if (loanAmountRaw) {
          var parsedLA = parseFloat(loanAmountRaw.replace(/[^0-9.-]+/g, ""));
          if (!isNaN(parsedLA) && parsedLA > 0) {
            loanAmountVal = parsedLA;
          }
        }

        var remainLoanVal =
          loanAmountVal > 0 ? loanAmountVal - totalPayments : 0;

        var formattedLoanAmount = formatMoney(loanAmountVal);
        var formattedRemainLoan = formatMoney(remainLoanVal);

        // Fetch total payments history and generate the PDF invoice
        var paymentsHistory = getPaymentsForCar(car["Car ID"]) || [];
        var pdfResult = generateInvoicePdf(
          car,
          clientName,
          clientEmail,
          formatMoney(spNum),
          formatMoney(dpNum),
          formattedLoanAmount,
          formattedRemainLoan,
          paymentsHistory,
        );

        // Build premium, creative HTML table for vehicle payments history in the email
        var paymentsHtml =
          '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse: collapse; font-size: 14px; margin: 20px 0; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">' +
          '<tr style="background-color: #f8fafc; border-bottom: 1px solid #e2e8f0;">' +
          '<th style="padding: 10px 12px; text-align: left; color: #475569; font-weight: 700; font-size: 12px; text-transform: uppercase;">Date</th>' +
          '<th style="padding: 10px 12px; text-align: left; color: #475569; font-weight: 700; font-size: 12px; text-transform: uppercase;">Amount</th>' +
          '<th style="padding: 10px 12px; text-align: left; color: #475569; font-weight: 700; font-size: 12px; text-transform: uppercase;">Method</th>' +
          '<th style="padding: 10px 12px; text-align: left; color: #475569; font-weight: 700; font-size: 12px; text-transform: uppercase;">Notes</th>' +
          "</tr>";

        if (paymentsHistory.length > 0) {
          paymentsHistory.forEach(function (pay) {
            var pDate = pay["PAYMENT DATE"] || pay["Payment Date"] || "";
            var pAmount = pay["AMOUNT"] || "";
            if (pAmount) {
              var pAmtNum =
                parseFloat(String(pAmount).replace(/[^0-9.-]+/g, "")) || 0;
              pAmount =
                "$" +
                pAmtNum.toLocaleString("en-US", {
                  minimumFractionDigits: 2,
                  maximumFractionDigits: 2,
                });
            }
            var pOption =
              pay["PAYMENT OPTION / NOTES"] ||
              pay["Payment Option"] ||
              pay["PAYMENT OPTION"] ||
              "-";
            var pNotes = pay["NOTES"] || pay["Notes"] || "-";

            paymentsHtml +=
              '<tr style="border-bottom: 1px solid #f1f5f9;">' +
              '<td style="padding: 10px 12px; color: #334155;">' +
              pDate +
              "</td>" +
              '<td style="padding: 10px 12px; color: #10b981; font-weight: 600;">' +
              pAmount +
              "</td>" +
              '<td style="padding: 10px 12px; color: #334155;">' +
              pOption +
              "</td>" +
              '<td style="padding: 10px 12px; color: #64748b; font-size: 13px;">' +
              pNotes +
              "</td>" +
              "</tr>";
          });
        } else {
          paymentsHtml +=
            '<tr><td colspan="4" style="padding: 15px; text-align: center; color: #94a3b8; font-style: italic;">No payments recorded yet.</td></tr>';
        }
        paymentsHtml += "</table>";

        // Creative, premium HTML Email Template
        var htmlBody =
          "<div style=\"font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 0; border: none; border-radius: 16px; background-color: #fafbfc; box-shadow: 0 4px 20px rgba(0,0,0,0.05); overflow: hidden;\">" +
          // Header
          '<div style="background: linear-gradient(135deg, #e056fd 0%, #6c5ce7 100%); padding: 35px 20px; text-align: center; color: #ffffff;">' +
          '<h1 style="margin: 0; font-size: 26px; font-weight: 800; letter-spacing: 2px; text-shadow: 0 2px 4px rgba(0,0,0,0.1);">WRIGHT FINDER MOTORS</h1>' +
          '<p style="margin: 8px 0 0 0; font-size: 14px; opacity: 0.9; font-weight: 500; text-transform: uppercase; letter-spacing: 1px;">Loan &amp; Financing Division</p>' +
          "</div>" +
          // Content Card
          '<div style="padding: 40px 30px; background-color: #ffffff; border-radius: 0 0 16px 16px;">' +
          '<div style="margin-bottom: 20px; display: table; width: 100%;">' +
          '<div style="display: table-cell; vertical-align: middle;"><h2 style="color: #2d3436; margin: 0; font-size: 20px; font-weight: 700;">Loan Balance Notice</h2></div>' +
          '<div style="display: table-cell; text-align: right; vertical-align: middle;"><span style="background-color: #f1f5f9; color: #475569; padding: 6px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; border: 1px solid #cbd5e1;">Invoice #' +
          pdfResult.invoiceNumber +
          "</span></div>" +
          "</div>" +
          '<p style="color: #636e72; font-size: 15px; line-height: 1.6;">Dear ' +
          clientName +
          ",</p>" +
          '<p style="color: #636e72; font-size: 15px; line-height: 1.6;">This is an automated notification regarding the active financing agreement for your vehicle. We have attached the official PDF invoice (Invoice #' +
          pdfResult.invoiceNumber +
          ") to this email for your records. Please find your loan status summary and payment history below:</p>" +
          // Side-by-Side Highlighted Badges (Bulletproof Table Layout)
          '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin: 25px 0;">' +
          "<tr>" +
          '<td width="48%" valign="top" style="background-color: #f1f5f9; border: 1px solid #cbd5e1; border-radius: 12px; padding: 20px; text-align: center;">' +
          '<span style="font-size: 10px; text-transform: uppercase; letter-spacing: 1px; color: #64748b; font-weight: 700; display: block; margin-bottom: 6px;">Loan Amount</span>' +
          '<strong style="font-size: 22px; color: #334155; font-weight: 800; display: block;">' +
          formattedLoanAmount +
          "</strong>" +
          "</td>" +
          '<td width="4%">&nbsp;</td>' +
          '<td width="48%" valign="top" style="background-color: #fef2f2; border: 2px solid #ef4444; border-radius: 12px; padding: 20px; text-align: center; box-shadow: 0 4px 15px rgba(239, 68, 68, 0.08);">' +
          '<span style="font-size: 10px; text-transform: uppercase; letter-spacing: 1.2px; color: #b91c1c; font-weight: 700; display: block; margin-bottom: 6px;">🚨 Remain Loan</span>' +
          '<strong style="font-size: 26px; color: #b91c1c; font-weight: 900; display: block;">' +
          formattedRemainLoan +
          "</strong>" +
          "</td>" +
          "</tr>" +
          "</table>" +
          // Details Grid
          '<div style="background-color: #f8f9fa; border: 1px solid #e9ecef; padding: 20px; margin: 25px 0; border-radius: 12px;">' +
          '<h4 style="margin: 0 0 15px 0; color: #2d3436; font-size: 14px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;">Vehicle Details</h4>' +
          '<table style="width: 100%; font-size: 14px; border-collapse: collapse; color: #2d3436;">' +
          '<tr style="border-bottom: 1px solid #e9ecef;"><td style="padding: 8px 0; color: #636e72; font-weight: 500;">Vehicle:</td><td style="padding: 8px 0; font-weight: 600; text-align: right;">' +
          (car["Year"] || "") +
          " " +
          (car["Car Name"] || "") +
          " " +
          (car["Model"] || "") +
          "</td></tr>" +
          '<tr style="border-bottom: 1px solid #e9ecef;"><td style="padding: 8px 0; color: #636e72; font-weight: 500;">Stock ID:</td><td style="padding: 8px 0; font-weight: 600; text-align: right;">' +
          car["Car ID"] +
          "</td></tr>" +
          '<tr style="border-bottom: 1px solid #e9ecef;"><td style="padding: 8px 0; color: #636e72; font-weight: 500;">Original Sold Price:</td><td style="padding: 8px 0; font-weight: 600; text-align: right;">' +
          formatMoney(spNum) +
          "</td></tr>" +
          '<tr><td style="padding: 8px 0 0 0; color: #636e72; font-weight: 500;">Down Payment Received:</td><td style="padding: 8px 0 0 0; font-weight: 600; text-align: right; color: #2ecc71;">' +
          formatMoney(dpNum) +
          "</td></tr>" +
          "</table>" +
          "</div>" +
          // Payments history section
          '<h3 style="color: #2d3436; font-size: 16px; font-weight: 700; margin-top: 30px; margin-bottom: 5px;">Vehicle Payments History</h3>' +
          paymentsHtml +
          '<p style="color: #636e72; font-size: 14px; line-height: 1.6;">Payments can be made in person, or via bank transfer/wire. Once your final installment is processed, your financing status will automatically update to <strong>Paid Off</strong>.</p>' +
          // Divider
          '<hr style="border: 0; border-top: 1px solid #eee; margin: 30px 0;">' +
          // Footer
          '<p style="color: #b2bec3; font-size: 11px; line-height: 1.5; text-align: center; margin: 0;">' +
          "You are receiving this automated update because an active vehicle financing agreement was registered with Wright Finder Motors under this email address. Please ignore if this loan has recently been paid off." +
          "</p>" +
          '<p style="color: #b2bec3; font-size: 11px; text-align: center; margin: 8px 0 0 0;">' +
          "&copy; " +
          new Date().getFullYear() +
          " Wright Finder Motors. All rights reserved." +
          "</p>" +
          "</div>" +
          "</div>";

        MailApp.sendEmail({
          to: clientEmail,
          subject:
            "⚠️ Action Required: Pay Your Active Loan Balance - Wright Finder Motors",
          htmlBody: htmlBody,
          attachments: [pdfResult.blob],
        });

        Logger.log(
          "Active Loan email notification sent to client " +
            clientEmail +
            " for car " +
            car["Car ID"] +
            " with invoice PDF",
        );
      }
    }
  } catch (e) {
    Logger.log("Error in sendActiveLoanPaymentReminders: " + e.toString());
  }
}

/**
 * Helper to initialize or verify the 2-minute time trigger for sending reminders.
 */
function setupActiveLoanPaymentTrigger() {
  var triggerName = "sendActiveLoanPaymentReminders";

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === triggerName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger(triggerName).timeBased().everyMinutes(1).create();

  Logger.log(
    "Trigger for '" +
      triggerName +
      "' created successfully (runs every 1 minute, sending emails on even minutes to achieve a 2-minute interval).",
  );
}

function generateInvoicePdf(
  car,
  clientName,
  clientEmail,
  soldPrice,
  downPayment,
  loanAmount,
  remainLoan,
  paymentsHistory,
) {
  var templateId = "1WhvW1WdQOtrbXZyA11RZUq1US3PvWFFnrC6tiielLjk";

  // Increment and format 5-digit invoice number using script properties
  var scriptProperties = PropertiesService.getScriptProperties();
  var invoiceCounter =
    parseInt(scriptProperties.getProperty("INVOICE_COUNTER") || "0") + 1;
  scriptProperties.setProperty("INVOICE_COUNTER", invoiceCounter.toString());
  var invoiceNumber = String(invoiceCounter).padStart(5, "0");

  // Copy template
  var tempFile = DriveApp.getFileById(templateId).makeCopy(
    "Invoice_" + invoiceNumber + "_" + car["Car ID"],
  );
  var doc = DocumentApp.openById(tempFile.getId());
  var body = doc.getBody();

  // Get current date formatted in spreadsheet's timezone
  var timezone =
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() ||
    Session.getScriptTimeZone();
  var todayDateStr = Utilities.formatDate(new Date(), timezone, "MM/dd/yyyy");

  // Replace standard placeholders
  body.replaceText("{{INVOICE_NUMBER}}", invoiceNumber);
  body.replaceText("{{CLIENT_NAME}}", clientName);
  body.replaceText("{{CLIENT_EMAIL}}", clientEmail);
  body.replaceText("{{SOLD_PRICE}}", soldPrice);
  body.replaceText("{{DOWN_PAYMENT}}", downPayment);
  body.replaceText("{{LOAN_AMOUNT}}", loanAmount);
  body.replaceText("{{REMAIN_LOAN}}", remainLoan);
  body.replaceText("{{TODAY_DATE}}", todayDateStr);

  // Replace vehicle spec placeholders
  body.replaceText("{{CAR_YEAR}}", car["Year"] || "N/A");
  body.replaceText("{{CAR_NAME}}", car["Car Name"] || "N/A");
  body.replaceText("{{CAR_MODEL}}", car["Model"] || "N/A");
  body.replaceText("{{CAR_VIN}}", car["VIN"] || "N/A");
  body.replaceText("{{CAR_MILEAGE}}", car["Mileage"] || "N/A");
  body.replaceText("{{CAR_ENGINE}}", car["Engine"] || "N/A");
  body.replaceText("{{CAR_TRANSMISSION}}", car["Transmission"] || "N/A");
  body.replaceText("{{CAR_EXT_COLOR}}", car["Exterior Color"] || "N/A");
  body.replaceText("{{CAR_INT_COLOR}}", car["Interior Color"] || "N/A");

  // Replace payments table if placeholder exists
  var searchPattern = "{{PAYMENTS_TABLE}}";
  var element = body.findText(searchPattern);
  if (element) {
    var textElement = element.getElement();
    var parent = textElement.getParent();

    var tableData = [
      ["VEHICLE PAYMENTS HISTORY", "", "", ""],
      ["Date", "Amount", "Method", "Notes"],
    ];

    if (paymentsHistory && paymentsHistory.length > 0) {
      paymentsHistory.forEach(function (pay) {
        var pDate = pay["PAYMENT DATE"] || pay["Payment Date"] || "";
        var pAmount = pay["AMOUNT"] || "";
        if (pAmount) {
          var pAmtNum =
            parseFloat(String(pAmount).replace(/[^0-9.-]+/g, "")) || 0;
          pAmount =
            "$" +
            pAmtNum.toLocaleString("en-US", {
              minimumFractionDigits: 2,
              maximumFractionDigits: 2,
            });
        }
        var pOption =
          pay["PAYMENT OPTION / NOTES"] ||
          pay["Payment Option"] ||
          pay["PAYMENT OPTION"] ||
          "";
        var pNotes = pay["NOTES"] || pay["Notes"] || "";
        tableData.push([pDate, pAmount, pOption, pNotes]);
      });
    } else {
      tableData.push(["-", "-", "-", "No payments recorded."]);
    }

    var index = parent.getParent().getChildIndex(parent);
    var table = parent.getParent().insertTable(index + 1, tableData);

    // Merge cells in first row and remove borders
    var row0 = table.getRow(0);
    for (var colIndex = row0.getNumCells() - 1; colIndex > 0; colIndex--) {
      row0.getCell(colIndex).merge();
    }
    table.setBorderWidth(0);

    // Row 0 styling (Header Row)
    var row0 = table.getRow(0);
    var cell0 = row0.getCell(0);
    cell0.setBackgroundColor("#1c0f45");
    var text0 = cell0.getChild(0).asParagraph().editAsText();
    text0.setFontFamily("Poppins");
    text0.setFontSize(9);
    text0.setBold(false);
    text0.setForegroundColor("#ffffff");

    // Row 1 styling (Column Headings Row)
    var row1 = table.getRow(1);
    for (var col = 0; col < row1.getNumCells(); col++) {
      var cell = row1.getCell(col);
      cell.setBackgroundColor("#ff0804");
      var text = cell.getChild(0).asParagraph().editAsText();
      text.setFontFamily("Poppins");
      text.setFontSize(9);
      text.setBold(false);
      text.setForegroundColor("#ffffff");
    }

    // Alternating rows styling (Rows 2+)
    for (var r = 2; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      var bgColor = r % 2 === 0 ? "#ab90ff" : "#ffffff";
      for (var col = 0; col < row.getNumCells(); col++) {
        var cell = row.getCell(col);
        cell.setBackgroundColor(bgColor);
        var text = cell.getChild(0).asParagraph().editAsText();
        text.setFontFamily("Poppins");
        text.setFontSize(9);
        text.setBold(false);
        text.setForegroundColor("#000000");
      }
    }

    parent.removeFromParent();
  }

  doc.saveAndClose();

  var pdfBlob = tempFile
    .getAs(MimeType.PDF)
    .setName("Invoice_" + invoiceNumber + "_" + car["Car ID"] + ".pdf");
  DriveApp.getFileById(tempFile.getId()).setTrashed(true);

  return {
    blob: pdfBlob,
    invoiceNumber: invoiceNumber,
  };
}

function testSendActiveLoanPaymentReminder() {
  var testEmail = "YOUR_TEST_EMAIL@gmail.com"; // Replace with your email, or leave default (uses your Google account email)
  var testCarId = "TR25-68"; // Optional: Specify a Stock # to test a specific vehicle

  // Find a vehicle with active loan status
  var vehicles = getVehicles();
  var car = null;

  if (testCarId) {
    car = vehicles.find(function (c) {
      return String(c["Car ID"]) === String(testCarId);
    });
  } else {
    // Look for the first car with active loan status
    car = vehicles.find(function (c) {
      var finStatus = (c["Financing status"] || "")
        .toString()
        .trim()
        .toLowerCase();
      var status = (c["Status"] || "").toString().trim().toLowerCase();
      var spRaw = c["SOLD PRICE"] ? c["SOLD PRICE"].toString().trim() : "";
      var spNum = parseFloat(spRaw.replace(/[^0-9.-]+/g, "")) || 0;
      var dpRaw = c["DOWN PAYMENT"] ? c["DOWN PAYMENT"].toString().trim() : "";
      var dpNum = parseFloat(dpRaw.replace(/[^0-9.-]+/g, "")) || 0;
      var isSold = status.includes("sold");
      var isForceActiveLoan = isSold && spNum > 0 && dpNum === 0;
      return finStatus === "loan active" || isForceActiveLoan;
    });
  }

  if (!car) {
    Logger.log(
      "No vehicle with active loan status found to test. Please ensure at least one vehicle has 'Financing status' set to 'Loan Active' or status sold with SOLD PRICE > 0.",
    );
    return;
  }

  Logger.log("Running test for Car ID: " + car["Car ID"]);

  // Determine destination email
  var clientEmail =
    testEmail && testEmail !== "YOUR_TEST_EMAIL@gmail.com"
      ? testEmail
      : Session.getActiveUser().getEmail();
  var clientName = car["CLIENT NAME"] || car["Buyer Name"] || "Valued Customer";

  var spRaw = car["SOLD PRICE"] ? car["SOLD PRICE"].toString().trim() : "";
  var spNum = parseFloat(spRaw.replace(/[^0-9.-]+/g, "")) || 0;
  var dpRaw = car["DOWN PAYMENT"] ? car["DOWN PAYMENT"].toString().trim() : "";
  var dpNum = parseFloat(dpRaw.replace(/[^0-9.-]+/g, "")) || 0;

  var loanAmountRaw = car["LOAN AMOUNT"]
    ? car["LOAN AMOUNT"].toString().trim()
    : "";

  var formatMoney = function (num) {
    return (
      "$" +
      num.toLocaleString("en-US", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2,
      })
    );
  };

  var totalPayments = getTotalPaymentForCar(car["Car ID"]) || 0;
  var loanAmountVal = spNum > 0 && dpNum > 0 ? spNum - dpNum : 0;
  if (loanAmountRaw) {
    var parsedLA = parseFloat(loanAmountRaw.replace(/[^0-9.-]+/g, ""));
    if (!isNaN(parsedLA) && parsedLA > 0) {
      loanAmountVal = parsedLA;
    }
  }

  var remainLoanVal = loanAmountVal > 0 ? loanAmountVal - totalPayments : 0;
  var formattedLoanAmount = formatMoney(loanAmountVal);
  var formattedRemainLoan = formatMoney(remainLoanVal);
  var paymentsHistory = getPaymentsForCar(car["Car ID"]) || [];

  // Generate invoice PDF
  var pdfResult = generateInvoicePdf(
    car,
    clientName,
    clientEmail,
    formatMoney(spNum),
    formatMoney(dpNum),
    formattedLoanAmount,
    formattedRemainLoan,
    paymentsHistory,
  );

  // Build table
  var paymentsHtml =
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse: collapse; font-size: 14px; margin: 20px 0; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">' +
    '<tr style="background-color: #f8fafc; border-bottom: 1px solid #e2e8f0;">' +
    '<th style="padding: 10px 12px; text-align: left; color: #475569; font-weight: 700; font-size: 12px; text-transform: uppercase;">Date</th>' +
    '<th style="padding: 10px 12px; text-align: left; color: #475569; font-weight: 700; font-size: 12px; text-transform: uppercase;">Amount</th>' +
    '<th style="padding: 10px 12px; text-align: left; color: #475569; font-weight: 700; font-size: 12px; text-transform: uppercase;">Method</th>' +
    '<th style="padding: 10px 12px; text-align: left; color: #475569; font-weight: 700; font-size: 12px; text-transform: uppercase;">Notes</th>' +
    "</tr>";

  if (paymentsHistory.length > 0) {
    paymentsHistory.forEach(function (pay) {
      var pDate = pay["PAYMENT DATE"] || pay["Payment Date"] || "";
      var pAmount = pay["AMOUNT"] || "";
      if (pAmount) {
        var pAmtNum =
          parseFloat(String(pAmount).replace(/[^0-9.-]+/g, "")) || 0;
        pAmount =
          "$" +
          pAmtNum.toLocaleString("en-US", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
          });
      }
      var pOption =
        pay["PAYMENT OPTION / NOTES"] ||
        pay["Payment Option"] ||
        pay["PAYMENT OPTION"] ||
        "-";
      var pNotes = pay["NOTES"] || pay["Notes"] || "-";

      paymentsHtml +=
        '<tr style="border-bottom: 1px solid #f1f5f9;">' +
        '<td style="padding: 10px 12px; color: #334155;">' +
        pDate +
        "</td>" +
        '<td style="padding: 10px 12px; color: #10b981; font-weight: 600;">' +
        pAmount +
        "</td>" +
        '<td style="padding: 10px 12px; color: #334155;">' +
        pOption +
        "</td>" +
        '<td style="padding: 10px 12px; color: #64748b; font-size: 13px;">' +
        pNotes +
        "</td>" +
        "</tr>";
    });
  } else {
    paymentsHtml +=
      '<tr><td colspan="4" style="padding: 15px; text-align: center; color: #94a3b8; font-style: italic;">No payments recorded yet.</td></tr>';
  }
  paymentsHtml += "</table>";

  var htmlBody =
    "<div style=\"font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 0; border: none; border-radius: 16px; background-color: #fafbfc; box-shadow: 0 4px 20px rgba(0,0,0,0.05); overflow: hidden;\">" +
    '<div style="background: linear-gradient(135deg, #e056fd 0%, #6c5ce7 100%); padding: 35px 20px; text-align: center; color: #ffffff;">' +
    '<h1 style="margin: 0; font-size: 26px; font-weight: 800; letter-spacing: 2px; text-shadow: 0 2px 4px rgba(0,0,0,0.1);">WRIGHT FINDER MOTORS</h1>' +
    '<p style="margin: 8px 0 0 0; font-size: 14px; opacity: 0.9; font-weight: 500; text-transform: uppercase; letter-spacing: 1px;">Loan &amp; Financing Division</p>' +
    "</div>" +
    '<div style="padding: 40px 30px; background-color: #ffffff; border-radius: 0 0 16px 16px;">' +
    '<div style="margin-bottom: 20px; display: table; width: 100%;">' +
    '<div style="display: table-cell; vertical-align: middle;"><h2 style="color: #2d3436; margin: 0; font-size: 20px; font-weight: 700;">Loan Balance Notice</h2></div>' +
    '<div style="display: table-cell; text-align: right; vertical-align: middle;"><span style="background-color: #f1f5f9; color: #475569; padding: 6px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; border: 1px solid #cbd5e1;">Invoice #' +
    pdfResult.invoiceNumber +
    "</span></div>" +
    "</div>" +
    '<p style="color: #636e72; font-size: 15px; line-height: 1.6;">Dear ' +
    clientName +
    ",</p>" +
    '<p style="color: #636e72; font-size: 15px; line-height: 1.6;">This is an automated notification regarding the active financing agreement for your vehicle. We have attached the official PDF invoice (Invoice #' +
    pdfResult.invoiceNumber +
    ") to this email for your records. Please find your loan status summary and payment history below:</p>" +
    '<table cellpadding="0" cellspacing="0" border="0" width="100%" style="margin: 25px 0;">' +
    "<tr>" +
    '<td width="48%" valign="top" style="background-color: #f1f5f9; border: 1px solid #cbd5e1; border-radius: 12px; padding: 20px; text-align: center;">' +
    '<span style="font-size: 10px; text-transform: uppercase; letter-spacing: 1px; color: #64748b; font-weight: 700; display: block; margin-bottom: 6px;">Loan Amount</span>' +
    '<strong style="font-size: 22px; color: #334155; font-weight: 800; display: block;">' +
    formattedLoanAmount +
    "</strong>" +
    "</td>" +
    '<td width="4%">&nbsp;</td>' +
    '<td width="48%" valign="top" style="background-color: #fef2f2; border: 2px solid #ef4444; border-radius: 12px; padding: 20px; text-align: center; box-shadow: 0 4px 15px rgba(239, 68, 68, 0.08);">' +
    '<span style="font-size: 10px; text-transform: uppercase; letter-spacing: 1.2px; color: #b91c1c; font-weight: 700; display: block; margin-bottom: 6px;">🚨 Remain Loan</span>' +
    '<strong style="font-size: 26px; color: #b91c1c; font-weight: 900; display: block;">' +
    formattedRemainLoan +
    "</strong>" +
    "</td>" +
    "</tr>" +
    "</table>" +
    '<div style="background-color: #f8f9fa; border: 1px solid #e9ecef; padding: 20px; margin: 25px 0; border-radius: 12px;">' +
    '<h4 style="margin: 0 0 15px 0; color: #2d3436; font-size: 14px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;">Vehicle Details</h4>' +
    '<table style="width: 100%; font-size: 14px; border-collapse: collapse; color: #2d3436;">' +
    '<tr style="border-bottom: 1px solid #e9ecef;"><td style="padding: 8px 0; color: #636e72; font-weight: 500;">Vehicle:</td><td style="padding: 8px 0; font-weight: 600; text-align: right;">' +
    (car["Year"] || "") +
    " " +
    (car["Car Name"] || "") +
    " " +
    (car["Model"] || "") +
    "</td></tr>" +
    '<tr style="border-bottom: 1px solid #e9ecef;"><td style="padding: 8px 0; color: #636e72; font-weight: 500;">Stock ID:</td><td style="padding: 8px 0; font-weight: 600; text-align: right;">' +
    car["Car ID"] +
    "</td></tr>" +
    "</table>" +
    "</div>" +
    '<h3 style="color: #2d3436; font-size: 16px; font-weight: 700; margin-top: 30px; margin-bottom: 5px;">Vehicle Payments History</h3>' +
    paymentsHtml +
    "</div>" +
    "</div>";

  MailApp.sendEmail({
    to: clientEmail,
    subject:
      "⚠️ Action Required: Pay Your Active Loan Balance - Wright Finder Motors",
    htmlBody: htmlBody,
    attachments: [pdfResult.blob],
  });

  Logger.log("Test email sent successfully to: " + clientEmail);
}
