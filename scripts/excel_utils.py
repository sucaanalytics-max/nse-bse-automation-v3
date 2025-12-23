import pandas as pd
from openpyxl import load_workbook

EXCEL_PATH = "data/nse_bse_business_growth.xlsx"

def append_df(sheet, df):
    book = load_workbook(EXCEL_PATH)

    if sheet in book.sheetnames:
        old = pd.read_excel(EXCEL_PATH, sheet_name=sheet)
        df = pd.concat([old, df], ignore_index=True)

    with pd.ExcelWriter(
        EXCEL_PATH,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        writer.book = book
        df.to_excel(writer, sheet_name=sheet, index=False)
