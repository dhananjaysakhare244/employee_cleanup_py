# Employee Cleanup Tool

A Python-based Windows desktop application for processing employee Excel records.

The application identifies terminated employees in the Head Count workbook and cleans their matching records in the Classic User workbook.

## Features

- Read employee information from an Excel workbook
- Identify terminated employees
- Select terminated employees where `Employee Status` is `Terminated`
- Match terminated employees by comparing master `Employee ID` with records `Emp_ID`
- Find all records belonging to terminated employees
- Delete all records except the last record
- Keep the records `Id` value in the last record
- Clear all other columns in the last record
- Leave non-terminated employees unchanged
- Create a Processing Log sheet
- Never modify the original input Excel files
- Save the processed data to a new Excel file

## Requirements

- Windows
- Python 3
- VS Code (recommended for developers)

Microsoft Excel is not required for the Python processing itself.

---

# Developer Setup

## 1. Open the project

Open the `Employee_Cleanup_Tool` folder in VS Code.

Open:

```text
Terminal → New Terminal
```

Make sure the terminal is in the project root.

Example:

```text
C:\Projects\Employee_Cleanup_Tool>
```

## 2. Check Python

```powershell
python --version
```

If `python` is not recognized:

```powershell
py --version
```

## 3. Create a virtual environment

Run this once:

```powershell
python -m venv .venv
```

Or:

```powershell
py -m venv .venv
```

## 4. Activate the virtual environment

```powershell
.venv\Scripts\Activate.ps1
```

## 5. Install dependencies

```powershell
pip install -r requirements.txt
```

## 6. Run the application locally

```powershell
python main.py
```

The Employee Cleanup Tool window should open.

## 7. Run without activating the virtual environment

```powershell
.venv\Scripts\python.exe main.py
```

This is useful if PowerShell does not allow virtual-environment activation.

## 8. Install dependencies without activation

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Then:

```powershell
.venv\Scripts\python.exe main.py
```

## Developer Quick Start

After the initial setup:

```powershell
.venv\Scripts\python.exe main.py
```

Or:

```powershell
.venv\Scripts\Activate.ps1
python main.py
```

---

# Running as a User

Normal users should receive the packaged Windows executable rather than the Python source code.

## User Requirements

For the packaged EXE, the user does not need:

- VS Code
- Python
- Python packages
- Virtual environment

## Running the EXE

The user simply:

1. Receives `EmployeeCleanupTool.exe`
2. Double-clicks the EXE
3. Selects the Head Count workbook
4. Selects the Classic User workbook
5. Selects the output location
6. Clicks `PROCESS FILES`

---

# Using the Application

The application requires two Excel files.

## 1. Head Count Workbook

This file contains employee information.

Example:

| Employee ID | Full Name | Employee Status | Email - Primary Work |
|---:|---|---|---|
| 1001 | John Smith | Active | john@example.com |
| 1002 | Jane Doe | Terminated | jane@example.com |
| 1003 | Mike Jones | On Leave | mike@example.com |

The Head Count workbook has a title row in row 1 and its column headers in row 2. Employee data starts on row 3. The application looks for employees where:

```text
Employee Status = Terminated
```

## 2. Classic User Workbook

This file contains one or more records for each employee.

Example:

| Id | User/Email/Profile | Database | Access Level | Email | Emp_ID |
|---:|---|---|---|---|---:|
| 21 | Jane Doe | System A | User | jane@example.com | 1002 |
| 22 | Jane Doe | System B | Admin | jane@example.com | 1002 |
| 23 | Jane Doe | System C | User | jane@example.com | 1002 |

---

# Processing Rules

For every terminated employee:

1. Find all records for the employee.
2. Keep the physically last matching record.
3. Delete all previous matching records.
4. If exactly one row matches, keep that row; if multiple rows match, keep the physically last row and delete the earlier matching rows.
5. Clear every other column in the retained record.
6. Leave non-terminated employees unchanged.

### Before

```text
21 | Jane Doe | System A | User | jane@example.com | 1002
22 | Jane Doe | System B | Admin | jane@example.com | 1002
23 | Jane Doe | System C | User | jane@example.com | 1002
```

### After

```text
23 | | | | | |
```

Only the records `Id` remains.

---

# Matching Logic

The application matches Classic User records by comparing the Head Count `Employee ID` with the Classic User `Emp_ID` value. The comparison ignores letter case and leading or trailing spaces.

```text
Head Count `Employee ID` = Classic User `Emp_ID`
```

Only employees whose `Employee Status` value is `Terminated` are processed. Values such as `Active` and `On Leave` are not processed.

---

# Last Record

Currently, "last record" means the physically last matching row in the Excel worksheet.

Example:

```text
Row 10 → User 1002
Row 11 → User 1002
Row 12 → User 1002
```

Row 12 is retained.

Rows 10 and 11 are deleted.

The `Id` value in row 12 is kept and every other column is cleared.

If the requirement changes to keep the record with the latest date/timestamp, the processing logic should be changed to use the appropriate date/timestamp column.

---

# Non-Terminated Employees

Classic User records belonging to non-terminated employees are not modified.

---

# Output

The original Classic User workbook is never overwritten.

Example:

```text
Classic_User.xlsx
        ↓
Classic_User_Cleaned.xlsx
```

The output workbook also contains a `Processing Log` sheet.

---

# Processing Log

The log contains information such as:

| Metric | Value |
|---|---:|
| Terminated employees in Head Count | 30 |
| Terminated employees found in Classic User | 30 |
| Classic User rows deleted | 199 |
| Classic User rows retained | 30 |
| Terminated employees not found in Classic User | 0 |

---

# Configuration

Default values:

```text
Head Count worksheet:  Sheet1
Classic User worksheet: Profund

Head Count employee ID header:       Employee ID
Head Count status header:            Employee Status
Classic User employee ID header:     Emp_ID
Classic User ID header to keep:      Id
```

These values can be changed from the application UI.

---

# Testing

Dummy test data can be used before processing real employee data.

The test dataset contains:

- 100 employees
- 30 terminated employees
- 70 non-terminated employees
- 5–10 records per employee
- Approximately 770 records

---

# Troubleshooting

## Python is not recognized

Try:

```powershell
py --version
```

If `py` works:

```powershell
py -m venv .venv
```

## PowerShell does not allow activation

Skip activation and run:

```powershell
.venv\Scripts\python.exe main.py
```

Install dependencies without activation:

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## ModuleNotFoundError

For example:

```text
ModuleNotFoundError: No module named 'openpyxl'
```

Run:

```powershell
.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Then:

```powershell
.venv\Scripts\python.exe main.py
```

---

# Quick Start

## Developer

After the initial setup:

```powershell
.venv\Scripts\python.exe main.py
```

## User

For the packaged application:

```text
Double-click EmployeeCleanupTool.exe
```

No Python or VS Code is required for the packaged EXE.
