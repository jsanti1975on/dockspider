// Assuming this is a new script in Office Scripts

function updateTenantList() {
  let workbook = new ExcelScript.Workbook();
  let bsListSheet = workbook.getWorksheet("bs_list");
  let tenantListSheet = workbook.getWorksheet("tenant_list");

  // Clear the existing data in the tenant_list sheet
  tenantListSheet.getRange().clear();

  let bsData = bsListSheet.getRange("A2:B" + bsListSheet.getRange().getRowCount()).getValues();
  let tenantData: any[][] = [];
  let maxSlipNumber = 0;

  for (let i = 0; i < bsData.length; i++) {
    let name = bsData[i][0];
    let slipNumber = bsData[i][1];

    if (name !== "" && slipNumber !== "") {
      tenantData.push([name, slipNumber]);

      // Keep track of the maximum slip number
      if (slipNumber > maxSlipNumber) {
        maxSlipNumber = slipNumber;
      }
    }
  }

  // Ensure there are rows in tenant_list for each slip number
  for (let slip = 1; slip <= maxSlipNumber; slip++) {
    let found = false;
    for (let j = 0; j < tenantData.length; j++) {
      if (tenantData[j][1] === slip) {
        found = true;
        break;
      }
    }
    if (!found) {
      tenantData.push(["", slip]);
    }
  }

  // Sort the tenantData by slipNumber
  tenantData.sort((a, b) => a[1] - b[1]);

  // Write the sorted tenantData to the tenant_list sheet
  tenantListSheet.getRange("A1:B" + tenantData.length).setValues(tenantData);
}
