# Automatic Excel & CSV Processor

A Python automation tool that cleans, processes and consolidates Excel and CSV files automatically.

## The problem

Businesses often receive spreadsheets containing:

- duplicate rows
- missing values
- multiple Excel worksheets
- multiple CSV/Excel files
- inconsistent data that must be processed manually

This tool automates that workflow.

## Features

- Reads CSV and Excel files automatically
- Processes multiple Excel worksheets
- Removes duplicate rows
- Handles missing values
- Sorts data automatically when an `order_id` column exists
- Creates cleaned Excel files
- Combines processed data into one master dataset
- Generates a processing summary
- Tracks the source file and worksheet of each row

## Example workflow

```text
input/
    sales.csv
    sales_multisheet.xlsx

          ↓

     Python script

          ↓

output/
    sales_clean.xlsx
    sales_multisheet_clean.xlsx
    master_report.xlsx
```

## Master report

The generated `master_report.xlsx` contains two worksheets:

### Cleaned_Data

All cleaned datasets combined into one table.

### Processing_Summary

A report showing:

- original rows
- duplicates removed
- missing values found
- final rows

for every processed file and worksheet.

## Installation

Install the required packages:

```bash
pip install -r requirements.txt
```

## Usage

Place your `.csv` and `.xlsx` files inside the `input` folder.

Then run:

```bash
python3 main.py
```

The processed files will automatically appear inside the `output` folder.

## Technologies

- Python
- pandas
- openpyxl