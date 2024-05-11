function updateTenantList() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var bsListSheet = spreadsheet.getSheetByName("bs_list");
  var tenantListSheet = spreadsheet.getSheetByName("tenant_list");
  
  // Clear the existing data in the tenant_list sheet
  tenantListSheet.clear();
  
  var bsData = bsListSheet.getRange("A2:C" + bsListSheet.getLastRow()).getValues();
  var tenantData = [];
  var maxSlipNumber = 0;
  
  for (var i = 0; i < bsData.length; i++) {
    var name = bsData[i][0];
    var slipNumber = bsData[i][1];
    var phoneNumber = bsData[i][2]; // Added line to get phone number from column C
    
    if (name != "" && slipNumber != "") {
      tenantData.push([name, slipNumber, phoneNumber]); // Modified line to include phone number
      
      // Keep track of the maximum slip number
      if (slipNumber > maxSlipNumber) {
        maxSlipNumber = slipNumber;
      }
    }
  }
  
  // Ensure there are rows in tenant_list for each slip number
  for (var slip = 1; slip <= maxSlipNumber; slip++) {
    var found = false;
    for (var j = 0; j < tenantData.length; j++) {
      if (tenantData[j][1] == slip) {
        found = true;
        break;
      }
    }
    if (!found) {
      tenantData.push(["", slip, ""]); // Modified line to include an empty phone number
    }
  }
  
  // Sort the tenantData by slipNumber
  tenantData.sort(function(a, b) {
    return a[1] - b[1];
  });
  
  // Write the sorted tenantData to the tenant_list sheet
  tenantListSheet.getRange(1, 1, tenantData.length, 3).setValues(tenantData); // Modified line to set 3 columns
}
