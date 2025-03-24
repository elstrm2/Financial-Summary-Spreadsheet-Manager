# Financial Summary Spreadsheet Manager

> A **Google Apps Script** project that automates financial data handling, enforces strict formatting, manages snapshots, and handles real-time currency conversion (fiat & crypto) in **Google Sheets**.

[![License](https://img.shields.io/badge/License-MIT%20Nc-blue.svg)](LICENSE.txt)
[![Language: JavaScript](https://img.shields.io/badge/Language-JavaScript-yellow.svg)](https://developer.mozilla.org/en-US/docs/Web/JavaScript)
[![Google Apps Script](https://img.shields.io/badge/Google-AppsScript-brightgreen.svg)](https://developers.google.com/apps-script)

---

## Table of Contents

1. [Overview](#overview)
2. [Features](#features)
3. [How It Works](#how-it-works)
4. [Getting Started](#getting-started)
5. [Sheet Structure](#sheet-structure)
6. [Formatting & Validation Rules](#formatting--validation-rules)
7. [Snapshots](#snapshots)
8. [Currency Conversion Details](#currency-conversion-details)
9. [Advanced Checks & Error Handling](#advanced-checks--error-handling)
10. [Available Functions](#available-functions)
11. [License](#license)
12. [Contributing](#contributing)
13. [Troubleshooting & FAQ](#troubleshooting--faq)

---

## Overview

**Financial Summary Spreadsheet Manager** is designed to:

- Keep your Google Sheets financial data organized and well-structured.
- Ensure consistent formatting (font size, alignment, cell borders, etc.).
- Automate currency conversion (fiat & crypto) with caching to reduce API calls.
- Save and restore snapshots of your current data for quick backups.

This project uses Google Apps Script—so you can add it directly inside your Google Sheet and harness the powerful features:

1. **Automatic Conversion**: Convert any listed currency (including stablecoins) to a “main currency” found in your `TOTAL:` row.
2. **Validation & Checks**: The script detects formatting issues (e.g., incorrect font sizes, missing bold headers), ensuring a professional and uniform layout.
3. **Snapshots**: Easily save a copy of your financial data for reference, or restore a previous snapshot at any time.

---

## Features

1. **Sheet Management**

   - **Single-Click Clearing**: Remove data while retaining headers.
   - **Example Data Insertion**: Quickly load sample categories, subgroups, and amounts to see how everything works.
   - **Structure Restoration**: Automated fix if your sheet’s columns or formatting are altered accidentally.

2. **Formatting Validation**

   - **Row-by-Row Rules**: Ensures each row is either a “main group,” “sub group,” “subtotal,” or “total” row with correct formatting.
   - **Borders & Alignment**: Each cell should have solid black borders; amounts right-aligned, currencies center-aligned, etc.
   - **Number Format**: Amounts and exchange rates must be at 4 decimal places.
   - **Case Sensitivity**: Currency codes must be uppercase, “TOTAL:” must be uppercase, “Subtotal:” in title case, etc.

3. **Currency Conversion**

   - **Fiat (Exchange Rate API)**: Real-time rates for typical currencies like USD, EUR, RUB, etc.
   - **Crypto (CoinCap API)**: Fetch price data for coins like BTC, ETH, BNB, TON, etc.
   - **Stablecoin Detection**: If a `Notes` field contains “stablecoin,” the script treats its currency as “USD.”
   - **Cache**: 1-hour caching to prevent excessive API calls.

4. **Snapshot Management**

   - **saveSnapshot()**: Automatically appends a new “Snapshot_YYYYMMDDHHmmss” sheet.
   - **loadLastSnapshot()**: Restores the most recent snapshot into the “Financial Summary” sheet.

5. **Debug Logging**
   - A “Debug Logs” sheet is automatically created (when needed) to store script errors and important messages.

---

## How It Works

1. **Checkbox-Triggered Actions**: Each row in the “Button” column has a checkbox that, when checked, calls a specific function (clear sheet, fill data, save snapshot, etc.).
2. **onOpen() Menu**: A custom “Financial Tools” menu is injected on spreadsheet open, giving you quick access to “Clear Cache,” “Check Structure,” or “Restore Structure.”
3. **onEdit() Trigger**: If you check the `Convert to Main Currency` box, the script will:
   - Identify your main currency (based on the `TOTAL:` row).
   - Fetch or retrieve cached rates.
   - Update columns D (Exchange Rate) and E (To Main Currency) for each row.

---

## Getting Started

1. **Download the Example**

   - Grab [`example.xlsx`](./example.xlsx) from this repository.

2. **Import into Google Sheets**

   - Create a blank sheet at [sheets.google.com](https://sheets.google.com).
   - `File > Import > Upload` > select `example.xlsx`.

3. **Insert Checkboxes**

   - In the “Financial Summary” sheet, select the entire “Button” column (column J).
   - Right-click → `Insert checkbox`.

4. **Apps Script Setup**

   - Go to `Extensions > Apps Script`.
   - Copy the contents of [`main.js`](./main.js) into the code editor.

5. **Save & Reload**

   - Click `Save`, then refresh your Google Sheet tab.

6. **Configure Trigger**

   - In Apps Script, open **Triggers** (left sidebar).
   - Add a trigger for `initiateConversion()`, from “Spreadsheet” event, type “On edit.”

7. **Enable Services**

   - In Apps Script, click **Services**.
   - Add **Google Sheets API**.

8. **Update `appsscript.json`**

   - Replace with:
     ```json
     {
       "timeZone": "UTC",
       "runtimeVersion": "V8",
       "exceptionLogging": "STACKDRIVER",
       "dependencies": {
         "enabledAdvancedServices": [
           {
             "userSymbol": "Sheets",
             "version": "v4",
             "serviceId": "sheets"
           }
         ]
       }
     }
     ```
   - Save.

9. **Authorize**
   - The first time you run any function, Google will prompt for permissions. Approve them.

---

## Sheet Structure

**“Financial Summary”** is the main sheet. It must contain the following columns in row 1:

| Column | Header           |
| ------ | ---------------- |
| **A**  | Category         |
| **B**  | Amount           |
| **C**  | Currency         |
| **D**  | Exchange Rate    |
| **E**  | To Main Currency |
| **F**  | Notes            |

In columns **H** through **J** (row 1):

| Column | Header      |
| ------ | ----------- |
| **H**  | Action      |
| **I**  | Description |
| **J**  | Button      |

**Action** and **Description** rows (below the headers) explain the function. The **Button** column has checkboxes that trigger the respective function.

### Additional Rules

- **Column G** is reserved and must remain empty (the script checks this).
- **`TOTAL:` row** must appear exactly once and define your main currency in column C.
- **Rows after `TOTAL:`** must remain empty or be validly formatted as well.

---

## Formatting & Validation Rules

1. **Header Row (1)**
   - 12pt font, bold, centered (both horizontally & vertically).
2. **Data Rows (2–N)**
   - 10pt font, black text, “Arial” family.
   - Correct alignment: amounts typically right-aligned, categories left, etc.
   - Borders: all cells must have solid black borders.
3. **Groups vs. Subgroups**
   - Groups have a label like “Bank Accounts” (bold).
   - Subgroups start with a dash (e.g., “- Bank 1”) with normal text.
4. **Subtotal Rows**
   - Labeled “Subtotal:” in column A, italic font, 10pt.
5. **`TOTAL:` Row**
   - Bold, 10pt in A, with numeric total in B, currency code in C, etc.
6. **4 Decimal Places**
   - Amounts and exchange rates must use four decimal places (e.g., “1.0000”, “3.1415”).
7. **Uppercase Currencies**
   - The script checks for uppercase in column C (e.g., “USD,” not “usd”).

The code also ensures each cell after your `TOTAL:` row is empty (no text, no format, no data validations).

---

## Snapshots

The script allows you to create and restore snapshots of your “Financial Summary” sheet.

1. **saveSnapshot()**

   - Copies the entire sheet into a new one named `Snapshot_YYYYMMDDHHMMSS`.
   - Removes any interactive checkboxes and columns for triggers.
   - Adds a small header with the timestamp.

2. **loadLastSnapshot()**
   - Finds the most recent snapshot by name.
   - Restores that data into “Financial Summary” (overwriting existing rows).

> **Pro Tip**: You can manually keep multiple snapshots if you want different points in time. They remain in your spreadsheet until manually deleted.

---

## Currency Conversion Details

1. **Main Currency**
   - Defined by the row labeled `TOTAL:` in column A.
   - The script reads column C in the same row to decide your main currency, e.g., “USD.”
2. **Stablecoins**
   - If a row’s `Notes` column (F) contains the word “stablecoin,” that currency is treated as USD.
   - A stablecoin cache is stored in `CacheService` to optimize repeated lookups.
3. **Fiat API**
   - [Exchange Rate API](https://open.er-api.com) is used to fetch real-time fiat rates.
   - Rates are cached for 1 hour under a special key. After 1 hour, the script queries again.
4. **Crypto API**
   - [CoinCap](https://api.coincap.io/v2/assets) is queried to get crypto prices in USD.
   - For example, if you hold 1 BNB and BNB is $300 USD, the script calculates `amount * exchangeRate`.
5. **onEdit Trigger**
   - Checking the “Convert to Main Currency” box in the `Action` section calls the function `convertToMainCurrency()`.
   - The script updates column D (Exchange Rate) and calculates column E (To Main Currency).

---

## Advanced Checks & Error Handling

Your code has numerous checks to ensure the sheet is correct and to prevent user errors:

1. **Strict Sheet Names**
   - Only “Financial Summary,” “Debug Logs,” and sheets starting with “Snapshot\_” are allowed. Others get deleted if you run “Restore Structure.”
2. **Column G Must Be Empty**
   - No data, no formatting, no notes, and no images/drawings can exist there.
3. **Single `TOTAL:` Row**
   - If more than one `TOTAL:` row is found, the script throws an error and logs it.
4. **Borders**
   - The script queries the Sheets API to confirm each cell has solid black borders.
5. **Multiple Checks**
   - The script runs “Check Structure” to ensure headings, bold/italic usage, row alignment, column widths (within ±5px of the expected), etc.
6. **Debug Logs**
   - Any major error is appended to a “Debug Logs” sheet with a timestamp, the calling function, and an error message.

---

## Available Functions

All functions are defined in `main.js`. Below are highlights:

- **`clearAllCache()`**  
  Clears the script cache (exchange rates, stablecoin addresses, etc.).
- **`onOpen()`**  
  Adds a custom menu named “Financial Tools” with items:
  1. Clear Cache
  2. Check Structure
  3. Restore Structure
- **`onEdit(e)`**  
  Catches checkbox updates in column J. If it’s “TRUE,” it calls the corresponding function (clear sheet, fill data, etc.).
- **`restoreAllStructure()`**  
  Restores initial structure (deletes invalid sheets, re-adds missing “Financial Summary,” sets up columns, etc.).
- **`checkSheetStructure()`**  
  Runs all validation checks on “Financial Summary.” If any errors are found, it displays them in a modal pop-up.
- **`convertToMainCurrency()`**  
  Retrieves rates for each currency, updates exchange rates, and recalculates columns D/E.

---

## License

This project is licensed under the **MIT License + Non-Commercial Clause**. See [LICENSE.txt](./LICENSE.txt) for more details.

---

## Contributing

1. **Fork** the repo.
2. **Create a feature branch**: `git checkout -b feature/your-feature`.
3. **Commit your changes**: `git commit -m "Add your feature"`.
4. **Push** the branch to your fork: `git push origin feature/your-feature`.
5. **Open a Pull Request** to the main repository.

We welcome PRs for improvements, bug fixes, and new features—please maintain existing code style and add tests where possible.

---

## Troubleshooting & FAQ

1. **I see an error “Multiple TOTAL: rows found!”**

   - Make sure there is only one row with “TOTAL:” in column A. Remove or rename extra rows.

2. **Why aren’t my crypto prices updating?**

   - Ensure the currency code is recognized by CoinCap. For instance, “BNB” is valid, “BSC_BNB” is not.
   - Also confirm you have an internet connection for external API calls.

3. **My sheet was renamed from ‘Financial Summary’ to something else**

   - The script expects the sheet to be named “Financial Summary.” Rename it back or update all references in the code.

4. **Column widths keep resetting**

   - The code enforces approximate column widths as part of the structure. Change the constants in `COLUMN_WIDTHS` if you need different widths.

5. **Is it possible to exclude certain rows from conversion?**
   - Currently, all rows with a “-” prefix are considered subgroups and are converted if they have a numeric `Amount`. You can customize the code to skip certain rows by adding conditions in `convertToMainCurrency()`.

For other issues, open an [Issue](https://github.com/elstrm2/Financial-Summary-Spreadsheet-Manager/issues) or start a [Discussion](https://github.com/elstrm2/Financial-Summary-Spreadsheet-Manager/discussions).

---

**Thank you for using the Financial Summary Spreadsheet Manager!**  
If this project helps you or saves you time, consider leaving a star on [GitHub](https://github.com/elstrm2/Financial-Summary-Spreadsheet-Manager).
