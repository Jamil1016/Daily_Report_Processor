#!/usr/bin/env python3

print("""
Author: Jamil Mendez
Version: 0.1.0
License: MIT
Github: https://github.com/Jamil1016/Daily_Report_Processor

      
Daily Report Processor

Requires: pandas, openpyxl, xlsxwriter, argparse
Install with:
    pip install pandas openpyxl xlsxwriter argparse

----------------------------------------
""")

__version__ = "v.0.1.0"
__author__ = "Jamil Mendez"

import pandas as pd
from pathlib import Path
import argparse


def merge_files(folder: Path) -> pd.DataFrame:
    """
    Merge all .xls files in the folder into a single DataFrame.
    Assumes files are tab-delimited with header starting from row 3.
    """
    all_dfs = []
    for file in folder.glob('*.xls'):
        try:
            df = pd.read_csv(file, header=2, sep='\t', encoding='gbk')
            all_dfs.append(df)
        except UnicodeDecodeError:
            df = pd.read_csv(file, header=2, sep='\t', encoding='utf-8')
            all_dfs.append(df)
        except Exception as e:
            print(f"[ERROR] Could not read '{file.name}': {e}")
    return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()


def clean_dish_name(text: str) -> str:
    """
    Clean the dish text by removing numbers, special characters, and standardizing text.
    """
    if not isinstance(text, str):
        return ""
    replacements = {
        'W/': 'WITH',
        '&': 'AND',
        ' WIT ': ' WITH ',
        '\\n': ' ',
        'PCS': ''
    }
    # Remove digits and selected symbols
    for ch in ['(', ')', '.', '@']:
        text = text.replace(ch, '')
    for key, val in replacements.items():
        text = text.replace(key, val)
    text = ''.join(c for c in text if not c.isdigit())
    return ' '.join(text.upper().split())


def process_report(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Clean and transform the merged data into three DataFrames: full report, transactions, and dishes.
    """
    # Fill forward key fields
    fields_to_fill = ['Date', 'POS Name', 'Cashier Name', 'Transaction No']
    df[fields_to_fill] = df[fields_to_fill].ffill()

    # Drop column if it exists
    df = df.drop(columns=['No data found'], errors='ignore')

    # Add DateTime and format Date
    df.insert(1, 'DateTime', pd.to_datetime(df['Date'], errors='coerce'))
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date

    # Clean dish names
    if 'Dishes' in df.columns:
        df['Dishes'] = df['Dishes'].apply(clean_dish_name)

    # Create transaction-level DataFrame
    df_transactions = df.drop_duplicates('OR No').drop(columns=['Dishes', 'Dish Quantities'], errors='ignore')

    # Create dish-level DataFrame
    dish_columns = ['OR No', 'DateTime', 'Date', 'POS Name', 'Cashier Name', 'Transaction No', 'Dishes', 'Dish Quantities']
    df_dishes = df[dish_columns].copy()

    return df, df_transactions, df_dishes


def export_to_excel(folder: Path, report: pd.DataFrame, transactions: pd.DataFrame, dishes: pd.DataFrame):
    """
    Export the three DataFrames to an Excel file with multiple sheets.
    """
    output_file = folder / "Daily_Report.xlsx"
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        report.to_excel(writer, sheet_name='DailyReport', index=False)
        transactions.to_excel(writer, sheet_name='Transactions', index=False)
        dishes.to_excel(writer, sheet_name='Dishes', index=False)
    print(f"[INFO] Report saved to: {output_file}")


def main():
    parser = argparse.ArgumentParser(description="Process daily POS .xls reports.")
    parser.add_argument(
        'folder',
        nargs='?',
        help='Path to the folder containing .xls files',
        default=input("Input the folder path: ").strip()
    )
    folder = Path(parser.parse_args().folder).resolve()

    if not folder.exists() or not folder.is_dir():
        print(f"[ERROR] Invalid folder path: {folder}")
        return

    df_raw = merge_files(folder)
    if df_raw.empty:
        print("[WARNING] No valid .xls files found or all files failed to load.")
        return

    df_report, df_tx, df_dishes = process_report(df_raw)
    export_to_excel(folder, df_report, df_tx, df_dishes)


if __name__ == "__main__":
    main()
