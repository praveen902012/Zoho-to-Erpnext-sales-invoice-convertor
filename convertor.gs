function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("IYF")
    .addItem("Format to ErpNext", "formatToErpNext")
    .addItem("Generate Invoice Sheet (IGST)", "generateInvoiceStructuredSheet")
    .addItem("Generate Invoice Sheet (CGST_SGST)", "generateInvoiceStructuredSheet_CGST_SGST")
    .addToUi();
}

/* ---------------- FORMAT FUNCTION ---------------- */

function formatToErpNext() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const response = ui.prompt(
    "Format to ErpNext",
    "Enter the SOURCE sheet name:",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Operation cancelled.");
    return;
  }

  const sourceSheetName = response.getResponseText().trim();

  if (!sourceSheetName) {
    ui.alert("Sheet name cannot be empty.");
    return;
  }

  const sourceSheet = ss.getSheetByName(sourceSheetName);

  if (!sourceSheet) {
    ui.alert(`Sheet "${sourceSheetName}" not found!`);
    return;
  }

  const targetSheetName = `Formatted ${sourceSheetName}`;

  const startRow = 2;
  const startCol = 1;
  const numCols = 21; // A → U
  const lastRow = sourceSheet.getLastRow();

  if (lastRow < startRow) {
    ui.alert("No data to process.");
    return;
  }

  // Read header separately
  const headerRange = sourceSheet.getRange(1, startCol, 1, numCols);
  const headerData = headerRange.getValues();

  // Read body data
  const dataRange = sourceSheet.getRange(
    startRow,
    startCol,
    lastRow - startRow + 1,
    numCols
  );
  const data = dataRange.getValues();

  // Process bottom → top in memory
  for (let i = data.length - 1; i > 0; i--) {
    const currentID = data[i][1];     // Column B
    const previousID = data[i - 1][1];

    if (currentID !== "" && currentID === previousID) {
      data[i] = new Array(numCols).fill("");
    }
  }

  // Create or reset target sheet
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
  } else {
    targetSheet.clear();
  }

  // Write header
  targetSheet.getRange(1, startCol, 1, numCols).setValues(headerData);

  // Write processed data
  targetSheet.getRange(startRow, startCol, data.length, numCols).setValues(data);

  targetSheet.autoResizeColumns(startCol, numCols);

  ui.alert(`Formatted data created in: ${targetSheetName}`);
}

/* ---------------- IGST INVOICE FUNCTION ---------------- */

function generateInvoiceStructuredSheet() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    "Generate Invoice Sheet (IGST)",
    "Enter the SOURCE sheet name:",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Operation cancelled.");
    return;
  }

  const SOURCE_SHEET = response.getResponseText().trim();

  if (!SOURCE_SHEET) {
    ui.alert("Sheet name cannot be empty.");
    return;
  }

  const TARGET_SHEET = `Invoice Structured ${SOURCE_SHEET}`;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(SOURCE_SHEET);

  if (!source) {
    ui.alert(`Sheet "${SOURCE_SHEET}" not found`);
    return;
  }

  const data = source.getDataRange().getValues();
  if (data.length <= 1) {
    ui.alert("No data found in source sheet.");
    return;
  }

  const headers = data[0];
  const rows = data.slice(1);

  const colIndex = {
    name: headers.indexOf("name"),
    igstRate: headers.indexOf("IGST Rate %"),
    shippingTax: headers.indexOf("Shipping Charge Tax %"),
    shippingAmount: headers.indexOf("Shipping Charge")
  };

  if (Object.values(colIndex).includes(-1)) {
    ui.alert("Required headers missing in source sheet");
    return;
  }

  let output = [];

  rows.forEach(row => {
    const invoice = row[colIndex.name];
    if (!invoice) return;

    const igstRate = row[colIndex.igstRate] || "";
    const shippingTaxRate = row[colIndex.shippingTax] || "";
    const shippingAmount = row[colIndex.shippingAmount] || "";

    output.push([
      invoice,
      "On Net Total",
      "Output Tax IGST - IYF",
      igstRate,
      "",
      "IGST",
      ""
    ]);

    output.push([
      "",
      "Actual",
      "Shipping - IYF",
      "",
      shippingAmount,
      "Shipping Charges (Net)",
      ""
    ]);

    output.push([
      "",
      "On Previous Row Amount",
      "GST Expense - IYF",
      shippingTaxRate,
      "",
      "GST on Shipping",
      2
    ]);
  });

  let target = ss.getSheetByName(TARGET_SHEET);

  if (!target) {
    target = ss.insertSheet(TARGET_SHEET);
  } else {
    target.clear();
  }

  const header = [
    "name",
    "Type (Sales Taxes and Charges)",
    "Account Head (Sales Taxes and Charges)",
    "Tax Rate (Sales Taxes and Charges)",
    "Amount (Sales Taxes and Charges)",
    "Description (Sales Taxes and Charges)",
    "Reference Row # (Sales Taxes and Charges)"
  ];

  target.getRange(1, 1, 1, header.length).setValues([header]);

  if (output.length > 0) {
    target.getRange(2, 1, output.length, header.length).setValues(output);
  }

  target.autoResizeColumns(1, header.length);

  ui.alert(`Created: ${TARGET_SHEET}`);
}

/* ---------------- CGST + SGST INVOICE FUNCTION ---------------- */

function generateInvoiceStructuredSheet_CGST_SGST() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    "Generate Invoice Sheet (CGST_SGST)",
    "Enter the SOURCE sheet name:",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("Operation cancelled.");
    return;
  }

  const SOURCE_SHEET = response.getResponseText().trim();

  if (!SOURCE_SHEET) {
    ui.alert("Sheet name cannot be empty.");
    return;
  }

  const TARGET_SHEET = `Invoice Structured ${SOURCE_SHEET}`;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(SOURCE_SHEET);

  if (!source) {
    ui.alert(`Source sheet "${SOURCE_SHEET}" not found`);
    return;
  }

  const data = source.getDataRange().getValues();
  if (data.length <= 1) {
    ui.alert("No data found in source sheet.");
    return;
  }

  const headers = data[0];
  const rows = data.slice(1);

  const colIndex = {
    name: headers.indexOf("name"),
    cgstRate: headers.indexOf("CGST Rate %"),
    sgstRate: headers.indexOf("SGST Rate %"),
    shippingTax: headers.indexOf("Shipping Charge Tax %"),
    shippingAmount: headers.indexOf("Shipping Charge")
  };

  if (Object.values(colIndex).includes(-1)) {
    ui.alert("One or more required headers not found in source sheet");
    return;
  }

  let output = [];

  rows.forEach(row => {
    const invoice = row[colIndex.name];
    if (!invoice) return;

    const cgstRate = row[colIndex.cgstRate] || "";
    const sgstRate = row[colIndex.sgstRate] || "";
    const shippingTaxPercent = row[colIndex.shippingTax] || "";
    const shippingAmount = row[colIndex.shippingAmount] || "";

    output.push([
      invoice,
      "On Net Total",
      "Output Tax SGST - IYF",
      sgstRate,
      "",
      "SGST",
      ""
    ]);

    output.push([
      "",
      "On Net Total",
      "Output Tax CGST - IYF",
      cgstRate,
      "",
      "CGST",
      ""
    ]);

    output.push([
      "",
      "Actual",
      "Shipping - IYF",
      "",
      shippingAmount,
      "Shipping Charges (Net)",
      ""
    ]);

    output.push([
      "",
      "On Previous Row Amount",
      "GST Expense - IYF",
      shippingTaxPercent,
      "",
      "GST on Shipping",
      3
    ]);
  });

  let target = ss.getSheetByName(TARGET_SHEET);

  if (!target) {
    target = ss.insertSheet(TARGET_SHEET);
  } else {
    target.clear();
  }

  const header = [
    "name",
    "Type (Sales Taxes and Charges)",
    "Account Head (Sales Taxes and Charges)",
    "Tax Rate (Sales Taxes and Charges)",
    "Amount (Sales Taxes and Charges)",
    "Description (Sales Taxes and Charges)",
    "Reference Row # (Sales Taxes and Charges)"
  ];

  target.getRange(1, 1, 1, header.length).setValues([header]);

  if (output.length > 0) {
    target.getRange(2, 1, output.length, header.length).setValues(output);
  }

  target.autoResizeColumns(1, header.length);

  ui.alert(`Created: ${TARGET_SHEET}`);
}