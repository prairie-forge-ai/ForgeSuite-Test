# PTO Accrual

Calculate and track employee PTO (Paid Time Off) accruals and balances.

## 🎯 Purpose

Automate PTO tracking by:
- Calculating accruals based on configured rates
- Tracking PTO usage and balances
- Handling carryover rules
- Supporting multiple accrual methods
- Generating PTO reports

## 📋 Contents

- `index.html` - Module UI
- `app.js` - Business logic
- `styles.css` - Module-specific styles (optional)
- `README.md` - This file

## ✨ Features

### **Accrual Calculation**
- Multiple accrual rates (by employee type, tenure)
- Flexible calculation periods (monthly, bi-weekly, annual)
- Automatic calculation based on hire date
- Pro-rated accruals for partial periods

### **Balance Tracking**
- Current PTO balance
- Accrued vs. used
- Available balance
- Historical tracking

### **Carryover Rules**
- Maximum carryover amounts
- Carryover date (typically year-end)
- Use-it-or-lose-it policies
- Unlimited accrual options

### **Reporting**
- Employee PTO balances
- Accrual history
- Usage reports
- Export to CSV

## 🔧 Configuration

Each customer has a `config.json`:

```json
{
  "moduleId": "pto-accrual",
  "enabled": true,
  
  "accrualRates": {
    "standard": {
      "hoursPerPeriod": 6.67,
      "period": "monthly",
      "maxAccrual": 160
    },
    "senior": {
      "hoursPerPeriod": 10,
      "period": "monthly",
      "maxAccrual": 240
    },
    "executive": {
      "hoursPerPeriod": 13.33,
      "period": "monthly",
      "maxAccrual": 320
    }
  },
  
  "carryoverRules": {
    "maxCarryover": 80,
    "carryoverDate": "01-01",
    "resetUnused": false
  },
  
  "calculationMethod": "accrual",
  
  "validationRules": {
    "requireEmployeeId": true,
    "validateBalances": true,
    "preventNegativeBalance": true
  }
}
```

## 📊 PTO Record Structure

```javascript
{
  employeeId: "E001",
  employeeName: "John Doe",
  accrualRate: "standard",
  hireDate: "2020-01-15",
  
  // Current balances
  totalAccrued: 160,      // Total PTO earned
  totalUsed: 80,          // Total PTO taken
  currentBalance: 80,     // Available PTO
  
  // Period calculations
  lastCalculated: "2025-11-15",
  accrualHistory: [
    { date: "2025-11-01", hours: 6.67, type: "accrual" },
    { date: "2025-10-15", hours: -16, type: "usage" },
    // ...
  ]
}
```

## 🎯 Key Functions

### **calculateAccruals()**
- Reads employee data
- Calculates PTO based on accrual rate
- Updates balances
- Handles pro-rating for partial periods
- Returns calculation results

### **calculateAccrualForEmployee(employee, asOfDate)**
- Gets employee's accrual rate
- Calculates time since last accrual
- Computes hours earned
- Checks against maximum
- Returns accrued hours

### **recordPTOUsage(employeeId, hours, date)**
- Validates employee exists
- Checks sufficient balance
- Records usage
- Updates current balance
- Creates history entry

### **applyCarryover()**
- Runs at year-end (or configured date)
- Checks carryover limits
- Resets unused PTO (if configured)
- Updates balances
- Creates audit trail

### **exportPTOData()**
- Formats PTO balances
- Creates CSV report
- Includes all employees
- Downloads to user

## 📐 Calculation Methods

### **Method 1: Accrual-Based (Default)**
```
Hours per period × Periods worked = Total accrued
Total accrued - Total used = Current balance
```

**Example:**
- Rate: 6.67 hours/month
- Months worked: 24
- Total accrued: 160 hours
- Used: 80 hours
- Balance: 80 hours

### **Method 2: Anniversary-Based**
```
On hire date anniversary:
Grant full year's PTO allocation
```

**Example:**
- Hire date: January 15
- Annual allocation: 80 hours
- On Jan 15 each year: +80 hours

### **Method 3: Pro-Rated**
```
(Hours per year / 12) × Months since hire = Accrued
```

**Example:**
- Annual: 80 hours
- Started June 1
- Worked 6 months
- Accrued: 40 hours

## 🧪 Testing

### **Test Cases:**
- [ ] Calculate monthly accrual
- [ ] Calculate bi-weekly accrual
- [ ] Handle new employee (pro-rated)
- [ ] Handle year-end carryover
- [ ] Prevent negative balance
- [ ] Export PTO report
- [ ] Calculate for multiple employee types

### **Test Data:**
```javascript
// New employee (pro-rated)
{
  employeeId: "E001",
  hireDate: "2025-09-01",  // Started mid-year
  accrualRate: "standard",
  asOfDate: "2025-11-15"   // 2.5 months worked
  // Expected: 6.67 × 2.5 = ~16.7 hours
}

// Long-term employee (at max)
{
  employeeId: "E002",
  hireDate: "2020-01-01",
  accrualRate: "standard",
  currentBalance: 160,     // At maximum
  // Expected: No more accrual until PTO used
}

// Carryover scenario
{
  employeeId: "E003",
  currentBalance: 120,     // Current balance
  maxCarryover: 80,        // Limit
  // Expected: After carryover, balance = 80
}
```

## 🎨 User Experience

### **Calculate PTO:**
1. Open PTO Accrual module
2. Select calculation period
3. Choose "As of Date"
4. Click "Calculate PTO Accruals"
5. See results for all employees

### **Record Usage:**
1. Select employee
2. Enter hours used
3. Enter date taken
4. Click "Record PTO Usage"
5. Balance updated automatically

### **Export Report:**
1. Click "Export PTO Report"
2. Choose date range (optional)
3. Download CSV file
4. Use for payroll or reporting

## 🛡️ Validation Rules

### **Balance Validation:**
- Cannot use more PTO than available
- Balances cannot be negative
- Maximum accrual enforced

### **Date Validation:**
- Usage date cannot be in future
- Must be after hire date
- Respects pay periods

### **Employee Validation:**
- Employee must exist in roster
- Must have accrual rate assigned
- Must have hire date

## 🔄 Integration with Other Modules

### **Employee Roster:**
- Gets employee list
- Retrieves hire dates
- Checks employee status
- Uses for accrual eligibility

### **Payroll Recorder:**
- Provides PTO usage for payroll
- Validates PTO taken
- Reports PTO balances for accruals

## 📊 Reports Available

### **PTO Balance Report:**
```
Employee    | Accrued | Used | Balance | Max
------------|---------|------|---------|-----
John Doe    | 160     | 80   | 80      | 160
Jane Smith  | 200     | 120  | 80      | 240
```

### **Accrual History Report:**
```
Date       | Employee  | Type    | Hours | Balance
-----------|-----------|---------|-------|--------
11/01/2025 | John Doe  | Accrual | +6.67 | 86.67
10/15/2025 | John Doe  | Usage   | -16   | 80
```

### **Carryover Report:**
```
Employee    | Before | Carryover Limit | After | Forfeited
------------|--------|-----------------|-------|----------
John Doe    | 120    | 80              | 80    | 40
Jane Smith  | 65     | 80              | 65    | 0
```

## 🎯 Customization Options

### **Per Customer:**
- Accrual rates (by employee type)
- Calculation periods
- Maximum accruals
- Carryover rules
- Reset dates

### **Per Employee:**
- Accrual rate assignment
- Custom adjustments
- Manual balance corrections
- Special allocations

## 🛡️ Error Messages

### **Common Errors:**

**Insufficient Balance:**
```
Error: Employee has 40 hours PTO, cannot use 80 hours
Fix: Reduce hours or wait for more accrual
```

**No Accrual Rate:**
```
Error: Employee "E001" has no accrual rate assigned
Fix: Assign rate in employee roster or config
```

**Invalid Date:**
```
Error: Cannot record PTO usage for future date
Fix: Use current or past date
```

**Maximum Reached:**
```
Error: Employee at maximum accrual (160 hours)
Fix: Use PTO before earning more
```

## 📚 Related Documentation

- [Employee Roster](../employee-roster/README.md) - Employee integration
- [Payroll Recorder](../payroll-recorder/README.md) - Payroll integration
- [Common Utilities](../common/README.md) - Date calculations

## 🔄 Version History

**Current Version:** 1.0.0
- Initial release
- Basic accrual calculation
- Monthly/bi-weekly periods
- Carryover support
- CSV export

**Planned:**
- v1.1.0 - Anniversary-based accrual
- v1.2.0 - Multiple PTO banks (sick, vacation)
- v2.0.0 - Advanced reporting & analytics

---

*Last updated: November 2025*
