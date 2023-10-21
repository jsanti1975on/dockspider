function verifyFuelTransactions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1");

  // Define tolerance (0.20)
  var tolerance = 0.20;

  // Find the last row with data in Column C
  var lastRow = sheet.getRange("C:C").getValues().filter(String).length;

  // Loop through the rows in the data (assuming data starts from row 2)
  for (var i = 2; i <= lastRow; i++) {
    // Get fuel type, amount, and gallons
    var fuelType = sheet.getRange(i, 3).getValue();
    var amount = sheet.getRange(i, 5).getValue();
    var gallons = sheet.getRange(i, 6).getValue();

    // Check if the fuel type is unknown
    if (fuelType !== "REC" && fuelType !== "AV") {
      clearCellBackground(sheet, i, 5);
      continue; // Skip this row and continue to the next one
    }

    // Calculate the expected amount based on the fuel type
    var expectedAmount;
    if (fuelType === "REC") {
      expectedAmount = 5.65 * gallons;
    } else if (fuelType === "AV") {
      expectedAmount = 6.50 * gallons;
    }

    // Check if the actual amount is within the tolerance (+/- 0.20)
    if (Math.abs(amount - expectedAmount) > tolerance) {
      clearCellBackground(sheet, i, 5);
      sheet.getRange(i, 5).setBackgroundRGB(255, 255, 0);
    }
  }
}

function clearCellBackground(sheet, row, col) {
  sheet.getRange(row, col).setBackground(null);
}
