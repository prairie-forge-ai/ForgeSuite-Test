# Payroll Recorder

Automate payroll journal entry creation and export for accounting systems.

## 🎯 Purpose

Helps accountants record payroll transactions by:
- Normalizing messy payroll data from various sources
- Validating employee and department information
- Generating accurate journal entries
- Exporting to formats compatible with accounting systems

## 📋 Contents

- `index.html` - Module UI
- `app.js` - Bootstraps the compiled bundle after Office initializes
- `app.bundle.js` - Beautified build artifact containing the full module logic
- `styles.css` - Module-specific styles (optional)
- `README.md` - This file
- `src/` - Readable source used to build `app.bundle.js`

### Building

```bash
npm run build:payroll
```

Updates `app.bundle.js` and emits `app.bundle.js.map`. Always edit files under `payroll-recorder/src/` (not the bundle) before running the build.

## ✨ Features

### **Data Processing**
- Import payroll data from Excel sheets
- Normalize dates, currencies, and text formats
- Validate against department lists
- Handle various input formats gracefully

### **Journal Entry Generation**
- Create double-entry accounting transactions
- Map to customer's chart of accounts
- Support multiple accounting systems (QuickBooks, NetSuite, etc.)
- Calculate totals and verify balance

### **Export Options**
- CSV format for import into accounting software
- Customizable column mapping
- Include transaction descriptions
- Add reference numbers

### **Error Handling**
- Identify and flag problematic rows
- Provide clear error messages
- Allow correction without re-processing
- Progressive validation (process what's valid)

## 🧰 Debugging the Bundle

- `app.bundle.js` now ships in a human-readable format so stack traces point to meaningful line numbers.
- The file is still the compiled artifact; edit the readable bundle directly if source modules are unavailable and keep comments noting changes.
- Source maps are not yet generated. If you introduce a build pipeline later, add `//# sourceMappingURL=...` references so the Office task pane can resolve minified stack traces.

## 🔧 Configuration

Each customer has a `config.json` in their folder:

```json
{
  "moduleId": "payroll-recorder",
  "enabled": true,
  
  "accountingSystem": "quickbooks-online",
  
  "departments": [
    "Sales",
    "Marketing",
    "Engineering",
    "Finance",
    "Operations"
  ],
  
  "chartOfAccounts": {
    "salariesExpense": "6010",
    "payrollTaxExpense": "6020",
    "benefitsExpense": "6030",
    "cashAccount": "1000",
    "payrollLiability": "2010"
  },
  
  "journalEntryFormat": "quickbooks-csv",
  
  "validationRules": {
    "requireDepartment": true,
    "requireEmployeeId": true,
    "validateAmounts": true
  }
}
```

## 📊 Expected Input Format

### **Payroll Data Sheet:**
```
| Date       | Employee Name | Dept    | Gross Pay | Taxes  | Net Pay |
|------------|---------------|---------|-----------|--------|---------|
| 11/15/2025 | John Doe      | Sales   | 5000.00   | 750.00 | 4250.00 |
| 11/15/2025 | Jane Smith    | Eng     | 6000.00   | 900.00 | 5100.00 |
```

### **Output Journal Entries:**
```csv
Date,Account,Debit,Credit,Description
11/15/2025,6010,11000.00,,Salaries Expense
11/15/2025,6020,1650.00,,Payroll Tax Expense
11/15/2025,1000,,12650.00,Cash - Payroll Payment
```

## 🎯 Key Functions

### **processPayroll()**
- Reads payroll data from active Excel sheet
- Normalizes all data fields
- Validates against configuration rules
- Generates journal entries
- Reports errors and successes

### **normalizePayrollData(data)**
- Converts Excel dates to standard format
- Parses currency values (removes $, commas)
- Trims whitespace
- Standardizes department names
- Validates required fields

### **generateJournalEntries(processedData)**
- Creates debit/credit entries
- Maps to chart of accounts
- Calculates totals
- Adds descriptions and references
- Ensures balanced entries

### **exportToCSV()**
- Formats journal entries for target system
- Creates downloadable CSV file
- Includes proper headers
- Escapes special characters

## 🧪 Testing

### **Test Cases:**
- [ ] Processes clean payroll data
- [ ] Handles missing dates
- [ ] Handles invalid currency formats
- [ ] Validates department names
- [ ] Generates balanced journal entries
- [ ] Exports valid CSV
- [ ] Shows clear error messages

### **Test Data:**
```javascript
// Valid data
{ date: '11/15/2025', employee: 'John Doe', dept: 'Sales', gross: 5000 }

// Invalid dates
{ date: 'tomorrow', employee: 'Jane Smith', dept: 'Sales', gross: 6000 }

// Invalid currency
{ date: '11/15/2025', employee: 'Bob Jones', dept: 'Eng', gross: '$5,000.00' }

// Missing department
{ date: '11/15/2025', employee: 'Alice Brown', dept: '', gross: 4500 }
```

## 🎨 User Experience

### **Step 1: Select Data**
- User has payroll data in Excel sheet
- Opens Prairie Forge add-in
- Navigates to Payroll Recorder

### **Step 2: Configure**
- Select pay period end date
- Choose department (optional filter)
- Verify settings

### **Step 3: Process**
- Click "Process Payroll Data"
- Module reads active sheet
- Shows progress indicator
- Displays results (success count, errors)

### **Step 4: Review**
- View generated journal entries
- Check for errors
- Verify totals balance

### **Step 5: Export**
- Click "Export Journal Entries"
- Download CSV file
- Import into accounting system

## 📐 Data Flow

```
Excel Sheet
    ↓
Read data (Office.js)
    ↓
Normalize formats
    ↓
Validate against rules
    ↓
Generate journal entries
    ↓
Format for accounting system
    ↓
Export CSV
    ↓
Download to user
```

## 🛡️ Error Handling

### **Common Errors:**

**Invalid Date:**
```
Error: Cannot parse date: "tomorrow"
Fix: Use MM/DD/YYYY format (e.g., 11/15/2025)
```

**Invalid Currency:**
```
Error: Cannot parse amount: "five thousand"
Fix: Use numeric format (e.g., 5000.00)
```

**Unknown Department:**
```
Error: Department "Eng" not found
Fix: Use one of: Sales, Marketing, Engineering, Finance, Operations
```

**Unbalanced Entry:**
```
Error: Debits ($11,650) don't equal Credits ($11,000)
Fix: Check payroll calculations
```

## 🎯 Customization Options

### **Per Customer:**
- Chart of accounts mapping
- Department lists
- Validation rules (required fields)
- Export format (QuickBooks, NetSuite, etc.)
- Date formats
- Currency symbols

### **Per Module:**
- UI labels and language
- Default values
- Additional fields
- Calculation methods

## 📚 Related Documentation

- [Common Utilities](../common/README.md) - Shared functions
- [Customer Configuration](../../customer-files/README.md) - Config structure
- [Office.js Docs](https://docs.microsoft.com/office/dev/add-ins/) - Excel API reference

## 🔄 Version History

See CHANGELOG.md for detailed version history.

**Current Version:** 1.0.0
- Initial release
- Basic payroll processing
- QuickBooks CSV export
- Department validation

---

*Last updated: November 2025*
