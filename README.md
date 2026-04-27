# Automated Financial Statement Extraction and PDF Generation System

## Overview
This project automates the extraction of financial data from a web system and generates PDF statements for multiple employees.

It reads employee data from an Excel file, navigates through the system, handles errors, extracts values, and saves structured outputs automatically.

---

## Technologies Used
- Python
- Selenium (Web Automation)
- Pandas (Data Processing)
- PyAutoGUI (Keyboard automation)
- Chrome Print Automation (PDF generation)

---

## How It Works

### 1. Input Data
- Reads an Excel file (`Base.xlsx`)
- Extracts:
  - Employee ID
  - Name
  - PIS number

---

### 2. Web Automation
- Opens browser with custom Chrome settings
- Configures automatic PDF printing
- Logs into the target system
- Navigates through menus

---

### 3. Data Processing Loop
For each employee:
- Inputs PIS number
- Attempts to retrieve account data

---

### 4. Error Handling
Handles specific system errors:
- Error 10039 → Account not found
- Error 144 → Account does not meet criteria

Counters are maintained for both error types.

---

### 5. Data Extraction
- Captures:
  - Financial values
  - Dates
- Updates Excel file dynamically
- Fills next available column automatically

---

### 6. Multiple Accounts Handling
- Detects if employee has multiple accounts
- Iterates through each account
- Extracts and stores all available data

---

### 7. PDF Generation
- Uses browser print feature
- Automatically saves as PDF
- Renames files:
  - "Employee Name.pdf"
  - Handles duplicates with incremental naming

---

### 8. Output
- Updated Excel file (`Base.xlsx`)
- Generated PDF files for each employee
- Summary printed in console:
  - Total errors (10039)
  - Total errors (144)

---

## Key Features
- Full automation of web-based workflow
- Intelligent error handling
- Multi-account data extraction
- Automatic PDF generation and naming
- Integration with Excel datasets

---

## Performance Gain
Manual Process:
- Hours or days depending on volume

Automated Process:
- Fully unattended execution in minutes

---

## How to Run

```bash
pip install selenium pandas pyautogui
python main.py
```

---

## Notes
- ChromeDriver must be installed
- Ensure Excel file structure matches expected columns
- System credentials must be updated in the script
- Screen resolution may affect PyAutoGUI actions

---

## Output Example

- Base.xlsx (updated with extracted values)
- Extrato - Employee Name.pdf
