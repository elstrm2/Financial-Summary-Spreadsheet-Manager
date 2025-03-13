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

1. Open your Google Spreadsheet
2. Go to Extensions > Apps Script
3. Copy the code from `main.js` into the script editor
4. Save and reload your spreadsheet

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
