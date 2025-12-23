from datetime import date
import pandas as pd
from openpyxl import load_workbook

EXCEL_PATH = "data/india_markets_daily.xlsx"

def today_str():
    return date.today().strftime("%Y-%m-%d")

def write_excel(sheet_name, df):
    try:
        book = load_workbook(EXCEL_PATH)
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            writer.book = book
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    except FileNotFoundError:
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
