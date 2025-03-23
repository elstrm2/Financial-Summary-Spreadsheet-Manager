# Financial Summary Spreadsheet Manager

A Google Apps Script project for managing financial data in Google Sheets with automated currency conversion, formatting checks, and snapshot management.

## 🌟 Features

- **Data Management**

  - Clear sheet while preserving structure
  - Fill example data for quick testing
  - Create and restore snapshots with timestamps
  - Automated currency conversion

- **Format Validation**

  - Strict formatting rules checking
  - Font size and style validation
  - Cell alignment verification
  - Number format control
  - Case sensitivity checking

- **Currency Handling**
  - Real-time exchange rates
  - Support for both fiat and cryptocurrencies
  - Stablecoin detection and handling
  - Cached exchange rates for performance

## 🚀 Getting Started

1. Download `example.xlsx` from the GitHub repository
2. Open Google Sheets (sheets.google.com)
3. Create a new spreadsheet
4. File > Import > Upload > Select the downloaded `example.xlsx`
5. In the "Button" column, replace all TRUE/FALSE values with Google Sheets checkboxes:
   - Select the entire "Button" column
   - Right-click > Insert checkbox
6. Go to Extensions > Apps Script
7. Copy the code from `main.js` into the script editor
8. Save and reload your spreadsheet
9. Set up the required trigger:
   - In the Apps Script editor, click on "Triggers" in the left sidebar
   - Click "+ Add Trigger" button
   - Configure the trigger:
     - Choose function: `initiateConversion`
     - Select event source: "From spreadsheet"
     - Select event type: "On edit"
     - Click "Save"
10. Enable required API services:
    - In the Apps Script editor, click on "Services" in the left sidebar
    - Click "+ Add Service" button
    - Find and enable "Google Sheets API"
    - Click "Add"
    - Save the script
11. Configure the manifest file:
    - In the Apps Script editor, click on `appsscript.json` in the left sidebar
    - Replace the contents with:
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
12. When first running the script, grant the requested permissions

## 📋 Sheet Structure

The spreadsheet requires specific column headers:

### Main Data Section (A1:F1)

- Category
- Amount
- Currency
- Exchange Rate
- To Main Currency
- Notes

### Action Section (H1:J1)

- Action
- Description
- Button

## 🛠️ Available Functions

### Menu Items

- `Clear Cache`: Removes cached exchange rates and stablecoin data
- `Check Structure`: Validates sheet formatting and structure

### Automated Actions

- `clearSheet()`: Clears data while preserving headers
- `fillExampleData()`: Inserts sample financial data
- `saveSnapshot()`: Creates a timestamped copy
- `loadLastSnapshot()`: Restores the most recent snapshot
- `convertToMainCurrency()`: Updates exchange rates and calculations

## 📝 Format Requirements

### Text Formatting

- Headers: 12pt, bold, center-aligned
- Data rows: 10pt
- Group headers: Bold
- Subtotals: Italic
- Currency codes: Uppercase

### Number Formatting

- Amounts: 2 decimal places
- Exchange rates: 4 decimal places

## 🔄 Currency Conversion

The system supports:

- Real-time fiat currency rates via Exchange Rate API
- Cryptocurrency rates via CoinCap API
- Automatic stablecoin detection
- Rate caching for 1 hour

## ⚠️ Known Issues and Limitations

### Structure Validation Issues

- Inconsistent validation of "Subtotal:" and "TOTAL:" entries - may incorrectly skip variations like "Subtotal" or "TOTA"
- Group structure validation needs improvement. Current structure requirements:
  ```
  Group Name
  - Subitem 1
  - Subitem 2 (Label)
  - Subitem 3
  Subtotal:
  ```
  The validator sometimes fails to properly check this hierarchy

### Planned Improvements

- Add automatic format correction for common structural issues
- Implement stricter group/subgroup relationship validation
- Enhanced subtotal detection and validation
- Smart case correction for standardized entries

---

## 📜 License

This project is licensed under the MIT License with a Non-Commercial Clause - see the [LICENSE.txt](LICENSE.txt) file for details.

## 🤝 Contributing

Feel free to open issues and submit pull requests. Please ensure you follow the existing code style and include appropriate tests.

## ⚠️ Requirements

- Google Sheets
- Google Apps Script
- Active internet connection for exchange rates

## 👤 Author

[elias](https://github.com/elstrm2)

---

_Note: This project is designed for personal use. Commercial use requires explicit permission from the author._
