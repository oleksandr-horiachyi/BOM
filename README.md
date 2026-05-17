# BOM Production Automation

Automation tool for processing multi-sheet BOM and LOG Excel files and generating production-ready schedules and reports.

---

## Overview

This application automates the workflow of:

- reading multi-sheet BOM workbooks
- processing production LOG files
- matching part codes between BOM and LOG
- calculating quantities automatically
- generating formatted Production Schedule Excel files
- generating Combined BOM reports
- validating data consistency and detecting issues

The goal is to reduce manual Excel work and eliminate production planning errors.

---

## Main Features

### BOM Processing
- Reads multi-sheet BOM files
- Combines all part data into a unified structure
- Calculates total quantities
- Creates Combined BOM reports

### LOG Processing
- Reads selected LOG sheets
- Removes empty columns automatically
- Normalizes part codes

### Production Schedule Generation
- Creates separate Production Schedule sheets
- Automatically fills QTY values
- Adds totals and formatting
- Generates schedule headers and company branding

### Validation & Checks
- Detects duplicate part codes
- Detects missing codes
- Generates warning/check reports

### Drawing Codes Report
- Groups parts by drawing code
- Calculates grouped quantities

---

## Technologies

- Python 3
- Tkinter
- OpenPyXL
- Pillow (PIL)

---

## Installation

Install dependencies:

```bash
pip install openpyxl pillow
