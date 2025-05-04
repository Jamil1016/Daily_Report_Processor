# 🧾 Daily Report Processor

A Python script to process daily `.xls` POS (Point of Sale) reports by merging, cleaning, and exporting them into a structured Excel file.

---

## 📌 Features

- Merges multiple `.xls` tab-delimited files from a folder
- Cleans and formats dish names
- Splits the data into:
  - A full detailed report
  - A summary of unique transactions
  - A breakdown of individual dish sales
- Exports all reports into a single Excel workbook with multiple sheets

---

## 📂 Input Format

- `.xls` files
- Tab-delimited
- Headers starting from **row 3**
- Encoding: `gbk` or fallback to `utf-8`

---

## 📦 Requirements

Install the required dependencies:

```bash
pip install pandas openpyxl xlsxwriter argparse
