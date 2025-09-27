function getAllData(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  return values.map(r => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

function getModuleData(module) {
  switch (module) {
    case "crm": return getAllData("CRM_Leads");
    case "inventory": return getAllData("INV_Items");
    case "projects": return getAllData("PROJ_Projects");
    case "billing": return getAllData("BILL_Invoices");
    default: return [];
  }
}
