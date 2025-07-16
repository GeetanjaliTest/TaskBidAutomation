# Bid Automation for Google Sheets

This Google Apps Script automates the generation of a **Bid Worksheet** in Google Sheets. It calculates material and labor costs, applies tax, margin, and overhead, and outputs a formatted bid summary ready for client presentation or internal use.

---

## Features

- Generates a complete **Bid Worksheet** with a single click.
- Pulls data from:
  - Takeoff
  - Material & Labor Pricing
  - Global Settings
- Automatically calculates:
  - Material & Labor Costs
  - Subtotals
  - Tax, Margin, and Overhead
  - Grand Total
- Applies proper currency formatting.
- Adds a custom menu to your Google Sheet for easy access.

---

## Required Sheet Structure

Create a Google Spreadsheet with the following sheets:

| Sheet Name        | Purpose                                      |
| ----------------- | -------------------------------------------- |
| Takeoff           | Project item breakdown (quantities, units)  |
| MaterialPricing   | Material rates by item name                 |
| LaborPricing      | Labor rates by item name                    |
| Settings          | Tax, margin, and overhead values            |
| BidWorksheet      | Output sheet (cleared and rewritten)        |

---

### Sheet Templates

#### Takeoff (Columns A–G starting from row 2)

| Item Name | Description | Unit | Net Qty | Waste % | Unit Type | Number of Units |

#### MaterialPricing (Columns A–B)

| Item Name | Material Rate |

#### LaborPricing (Columns A–B)

| Item Name | Labor Rate |

#### Settings (Columns A–B)

| Key           | Value  |
| ------------- |--------|
| TaxRate       | e.g. 8 |
| MarginPercent | e.g. 15 |
| OverheadCost  | e.g. 100 |

BidWorksheet can be left blank — it will be auto-filled by the script.

---

## How to Use

### 1. Open Apps Script

- Go to your Google Sheet → **Extensions → Apps Script**
- Paste the script from `bidautomation.gs`

### 2. Save & Reload

- Save the project
- Reload the spreadsheet

### 3. Generate the Bid

- You’ll see a new menu: **Bid Automation**
- Click: **Bid Automation → Generate Bid Sheet**
- Approve permissions
- The script will process and output a formatted sheet in `BidWorksheet`

---

## Output Overview

The generated BidWorksheet will contain:

| Column | Description                         |
|--------|-------------------------------------|
| A–G    | Raw input from Takeoff             |
| H–I    | Pulled Material & Labor Rates      |
| J–K    | Calculated Costs                   |
| L      | Line Total (Material + Labor)      |
| -      | Subtotal, Tax, Margin, Overhead    |
| -      | Grand Total                        |

All cost-related columns are formatted as currency.

---

## Error Handling

The script will alert and stop if:

- A required sheet is missing
- The Takeoff sheet is empty

---

## Customization Tips

To add more settings, modify:

```js
const settingsArray = settingsSheet.getRange("A2:B4").getValues();
```

To change currency format:

```js
.setNumberFormat("$#,##0.00");
```

---

## Script Entry Point

```js
function runBidAutomation() { ... }
```

Adds a UI menu on open:

```js
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Bid Automation")
    .addItem("Generate Bid Sheet", "runBidAutomation")
    .addToUi();
}
```
