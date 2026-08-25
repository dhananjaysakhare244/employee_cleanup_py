# Employee Cleanup Tool

A Python-based Windows desktop application for processing employee Excel records.

The application identifies terminated employees from an employee master Excel file and cleans their corresponding records from another Excel file.

## Features

- Read employee information from an Excel workbook
- Identify terminated employees
- Match terminated employees using User ID
- Use Email as a fallback when User ID does not match
- Find all records belonging to terminated employees
- Delete all records except the last record
- Keep the User ID in the last record
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
3. Selects the Employee Master Excel file
4. Selects the Employee Records Excel file
5. Selects the output location
6. Clicks `PROCESS FILES`

---

# Using the Application

The application requires two Excel files.

## 1. Employee Master File

This file contains employee information.

Example:

| User ID | Email | First Name | Last Name | Terminated |
|---:|---|---|---|---|
| 1001 | john@example.com | John | Smith | No |
| 1002 | jane@example.com | Jane | Doe | Yes |
| 1003 | mike@example.com | Mike | Jones | No |

The application looks for employees where:

```text
Terminated = Yes
```

## 2. Employee Records File

This file contains one or more records for each employee.

Example:

| User ID | Email | Date | Status | Description |
|---:|---|---|---|---|
| 1002 | jane@example.com | 2026-01-01 | Active | Record 1 |
| 1002 | jane@example.com | 2026-02-01 | Updated | Record 2 |
| 1002 | jane@example.com | 2026-03-01 | Updated | Record 3 |

---

# Processing Rules

For every terminated employee:

1. Find all records for the employee.
2. Keep the physically last matching record.
3. Delete all previous matching records.
4. Keep the User ID in the retained record.
5. Clear every other column in the retained record.
6. Leave non-terminated employees unchanged.

### Before

```text
1002 | jane@example.com | 2026-01-01 | Active  | Record 1
1002 | jane@example.com | 2026-02-01 | Updated | Record 2
1002 | jane@example.com | 2026-03-01 | Updated | Record 3
```

### After

```text
1002 | | | |
```

Only the User ID remains.

---

# Matching Logic

The application uses the following priority:

```text
User ID
   ↓
If User ID doesn't match
   ↓
Email
```

User ID is the primary matching field.

Email is used as a fallback.

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

The User ID in row 12 is kept and every other column is cleared.

If the requirement changes to keep the record with the latest date/timestamp, the processing logic should be changed to use the appropriate date/timestamp column.

---

# Non-Terminated Employees

Records belonging to non-terminated employees are not modified.

---

# Output

The original records Excel file is never overwritten.

Example:

```text
Employee_Records.xlsx
        ↓
Employee_Records_Cleaned.xlsx
```

The output workbook also contains a `Processing Log` sheet.

---

# Processing Log

The log contains information such as:

| Metric | Value |
|---|---:|
| Terminated employees in employee file | 30 |
| Terminated employees found in records | 30 |
| Records deleted | 199 |
| Records retained | 30 |
| Terminated employees not found in records | 0 |

---

# Configuration

Default values:

```text
Employee sheet:       Sheet1
Records sheet:        Sheet1

Employee User ID:     User ID
Employee Email:       Email
Terminated column:    Terminated
Terminated value:     Yes

Records User ID:      User ID
Records Email:        Email
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
