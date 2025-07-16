function runBidAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check if required sheets exist
  const requiredSheets = ["Takeoff", "MaterialPricing", "LaborPricing", "Settings", "BidWorksheet"];
  requiredSheets.forEach(name => {
    if (!ss.getSheetByName(name)) throw new Error(`Missing sheet: "${name}"`);
  });

  const takeoffSheet = ss.getSheetByName("Takeoff");
  const materialSheet = ss.getSheetByName("MaterialPricing");
  const laborSheet = ss.getSheetByName("LaborPricing");
  const settingsSheet = ss.getSheetByName("Settings");
  const outputSheet = ss.getSheetByName("BidWorksheet");

  outputSheet.clearContents();

  const takeoffData = takeoffSheet.getRange(2, 1, takeoffSheet.getLastRow() - 1, 7).getValues(); // A:G
  if (takeoffData.length === 0) throw new Error("'Takeoff' sheet is empty.");

  const materialData = materialSheet.getRange(2, 1, materialSheet.getLastRow() - 1, 2).getValues();
  const laborData = laborSheet.getRange(2, 1, laborSheet.getLastRow() - 1, 2).getValues();

  const settingsArray = settingsSheet.getRange("A2:B4").getValues();
  const settings = Object.fromEntries(settingsArray);
  const taxRate = parseFloat(settings["TaxRate"]) / 100 || 0;
  const margin = parseFloat(settings["MarginPercent"]) / 100 || 0;
  const overhead = parseFloat(settings["OverheadCost"]) || 0;

  const headers = [
    "Item Name", "Description", "Unit", "Net Qty", "Waste %", "Unit Type", "Number of Units",
    "Material Rate", "Labor Rate", "Material Cost", "Labor Cost", "Line Total"
  ];
  const output = [headers];

  let subtotal = 0;

  takeoffData.forEach(row => {
    const [itemName, description, unit, netQty, wastePercent, unitType, numUnits] = row;

    const qtyVal = parseFloat(numUnits) || 0;
    const materialMatch = materialData.find(m => m[0] === itemName);
    const laborMatch = laborData.find(l => l[0] === itemName);

    const materialRate = materialMatch ? parseFloat(materialMatch[1]) : 0;
    const laborRate = laborMatch ? parseFloat(laborMatch[1]) : 0;

    const materialCost = materialRate * qtyVal;
    const laborCost = laborRate * qtyVal;
    const lineTotal = materialCost + laborCost;

    subtotal += lineTotal;

    output.push([
      itemName, description, unit, netQty, wastePercent, unitType, qtyVal,
      materialRate, laborRate, materialCost, laborCost, lineTotal
    ]);
  });

  // Totals
  const taxAmount = subtotal * taxRate;
  const marginAmount = subtotal * margin;
  const grandTotal = subtotal + taxAmount + marginAmount + overhead;

  // Push totals (fixed row with 12 empty columns)
  output.push(["", "", "", "", "", "", "", "", "", "", "", ""]);
  output.push(["", "", "", "", "", "", "", "", "", "", "Subtotal", subtotal]);
  output.push(["", "", "", "", "", "", "", "", "", "", "Tax", taxAmount]);
  output.push(["", "", "", "", "", "", "", "", "", "", "Margin", marginAmount]);
  output.push(["", "", "", "", "", "", "", "", "", "", "Overhead", overhead]);
  output.push(["", "", "", "", "", "", "", "", "", "", "Grand Total", grandTotal]);

  // Write to BidWorksheet
  outputSheet.getRange(1, 1, output.length, output[0].length).setValues(output);

  // Format currency columns
  const dataRows = output.length - 6;
  const currencyCols = [8, 9, 10, 11, 12]; // Material Rate, Labor Rate, Mat Cost, Lab Cost, Line Total
  currencyCols.forEach(col => {
    outputSheet.getRange(2, col, dataRows - 1).setNumberFormat("$#,##0.00");
  });

  // Format total values
  outputSheet.getRange(dataRows + 2, 12, 5).setNumberFormat("$#,##0.00");

  SpreadsheetApp.getUi().alert("BidWorksheet generated successfully!");
}

// Custom menu
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Bid Automation")
    .addItem("Generate Bid Sheet", "runBidAutomation")
    .addToUi();
}
