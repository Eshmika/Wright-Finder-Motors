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
      "Style of Car",
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
      "DISP-PRICE",
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
    data.styleOfCar || "",
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
    "",
    "",
    data.tradeStatus || "",
    "",
    "",
    "",
    "",
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
  ]);

  return "Success";
}

function updateVehicleData(updatedData) {
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
      return "Success";
    }
  }
  return "Error: Car ID not found";
}

function getVehicles() {
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
      "AMOUNT",
      "EXPENSE DATE",
    ]);
  }

  sheet.appendRow([
    new Date(),
    data.carModel || "",
    data.carId || "",
    data.clientName || "",
    data.description || "",
    data.amount || "",
    data.expenseDate || "",
  ]);

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
    var expense = {};
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      expense[header] = row[j];
    }
    expenses.push(expense);
  }

  return expenses;
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
    ]);
  }

  sheet.appendRow([
    new Date(),
    data.carModel || "",
    data.carId || "",
    data.clientName || "",
    data.paymentOption || "",
    data.amount || "",
    data.paymentDate || "",
  ]);

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
    var payment = {};
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      payment[header] = row[j];
    }
    payments.push(payment);
  }

  return payments;
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

  sheet.appendRow([
    new Date(), // 0: Timestamp
    carId, // 1: Car ID
    data.carName || "", // 2
    data.model || "", // 3
    data.year || "", // 4
    data.mileage || "", // 5
    "", // 6: Price
    "", // 7: Discount
    data.vin || "", // 8
    "Available", // 9: Status
    data.title || "", // 10
    "", // 11: Style of Car
    "", // 12: Body Style
    "", // 13: Rent or Sell
    "", // 14: Engine
    "", // 15: Engine Type/Size
    "", // 16: Transmission
    "", // 17: Driveline
    "", // 18: Fuel Type
    "", // 19: Power Options
    "", // 20: Drive Condition
    "", // 21: Condition
    "", // 22: Seat Material
    "", // 23: Interior Color
    "", // 24: Exterior Color
    "", // 25: Interior Features
    "", // 26: Main Image URLs
    "", // 27: Sub Image URLs
    data.clientName || "", // 28: CLIENT NAME
    data.purchaseDate || "", // 29: PURCHASE DATE
    "", // 30: SOLD DATE
    data.tradeStatus || "", // 31: Trade status
    data.iaaiTotalPrice || "", // 32: IAAI TOTAL PRICE W/ FEES
    data.papePrice || "", // 33: PAPE PRICE
    "", // 34: DISP-PRICE
    data.transportFees || "", // 35: TRANSPORT FEES
    "", // 35: SOLD PRICE
    "", // 36: DOWN PAYMENT
    data.pickupLocation || "", // 37: CAR PICKUP LOCATION
    data.driverName || "", // 38: DRIVER NAME
    "", // 39: DRIVER INFORMATION
    data.notes || "", // 40: NOTES
    data.iaaiPriceBeforeFees || "", // 41: IAAI Price Before Fees
    data.dispatcherName || "", // 42: Dispatcher Name
    data.dispatcherPrice || "", // 43: Dispatcher Price
    data.dispatcherPhone || "", // 44: Dispatcher Phone number
    data.clientPhone || "", // 45: Client Phone
    data.clientEmail || "", // 46: Client Email
    data.driverPhone || "", // 47: Driver Phone
    data.driverCompany || "", // 48: Driver Company
  ]);

  return "Success";
}
