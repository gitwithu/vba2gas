# 📊 Excel to Google Sheets Automation – GAS Migration Project

This repository contains the Google Apps Script (GAS) code used to migrate and automate a complex Excel-based workbook (`togas.xlsm`) into Google Sheets, replicating its functionality and layout using JavaScript-based scripts.

---

## 🚀 Project Overview

The goal was to recreate the full Excel experience—including navigation, calculations, buttons, and logic—in Google Sheets. This includes:

- Migrating 36 Excel sheets with layout, formulas, and structure
- Converting Excel VBA class modules to Apps Script
- Implementing button-based navigation, printing, and sheet control
- Reproducing checkbox-based triggers and actions
- Providing workarounds for features not natively supported in Google Sheets (e.g., zoom level)

---

## 📁 Folder Structure

vba2gas
- scripts/
  - main.gs (Main application logic).
  - LoanClasses.gs (Converted VBA class modules to GAS)
- reference/
  - vba_output.txt (Original extracted VBA code from Excel)
- README.md


---

## 🛠 Features Implemented

- ✅ Sheet-to-sheet navigation via custom menu and image buttons
- ✅ Automated copying of all sheets and data from Excel into Google Sheets
- ✅ Button functionality for printing, hiding/showing sheets, toggling checkboxes
- ✅ Conversion of class-based VBA to Apps Script classes
- ✅ Scripts designed to mirror original logic, including financial plan handling, CF simulation, and repayment plans

---

## ⚠ Limitations & Notes

- Google Sheets doesn’t support Excel-style zoom control or native VBA UI interactions
- Certain features were replicated using best-possible Apps Script equivalents (e.g., instructional alerts, image buttons)
- Some functionality (checkbox toggles, visual formatting) requires additional enhancement due to Google Sheets constraints

---

## 🧾 Reference

- Original Excel file: `togas.xlsm`
- Extracted VBA reference: `reference/vba_output.txt`

