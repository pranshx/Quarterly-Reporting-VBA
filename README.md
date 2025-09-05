# Quarterly-Reporting-VBA

**Quarterly Log Automation & Region-based Worksheet Copier**  
Excel VBA solution to automate monthly → quarterly reporting tasks: region selection, template duplication, formatted log generation, and batch printing.

---

## What it does
- Prompts user to select a **Region** and **Quarter**.  
- Copies a monthly template into three monthly sheets and renames them to the appropriate months.  
- Inserts a standardized **Log** block with headers (Client Name, Contact Name, Date, Duration, Notes) and formatting.  
- Supports bulk printing of monthly sheets.  
- Basic input validation and retry flow to prevent invalid entries.

---

## Files in this repo
- `Quarterly_Template.xlsm` — (optional) sample macro-enabled workbook with macros. **Contains no real client data.**  
- `vba_modules/MoveCopyPrint.bas` — exported VBA module (readable text) with the macro code.  
- `data/sample_input.csv` — minimal dummy data if needed.  
- `screenshots/run-demo.gif` — short demo of the macro running (recommended).  
- `README.md` — this file.

---

## How to use
1. **Download** the `Sample_data.xlsm`.  
2. Enable macros in Excel.  
3. If using `.bas`, open VBA Editor (Alt+F11) → File → Import File → `vba_module.bas`.  
4. Run the main macro: `MoveCopyPrintDivisionQuarter` (Developer → Macros → Run).  
5. Follow on-screen prompts for Region and Quarter.

---

## Technical notes
- Languages / tech: Excel VBA (macros), basic sample CSV for demonstration.  
- Key functions:
  - `aMovebyInput` — region selection and dispatch to region-specific routines
  - `CreateQuarterlyLog` — formatted log header generator
  - `DuplicateMonthlyLog` — duplicate + rename monthly sheets
  - `PrintTeamQuarter` — batch print routine

---

## Demo
![Demo](Step_Into_rec.gif)  

---


