import pandas as pd
from openpyxl import load_workbook

EXCEL_PATH = "data/nse_bse_business_growth.xlsx"

def append_unique(sheet_name: str, df_new: pd.DataFrame, date_col="date"):
    book = load_workbook(EXCEL_PATH)
    if sheet_name in book.sheetnames:
        df_old = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)
        df = pd.concat([df_old, df_new], ignore_index=True)
        df = df.drop_duplicates(subset=[date_col], keep="last")
    else:
        df = df_new

    with pd.ExcelWriter(
        EXCEL_PATH,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        writer.book = book
        df.to_excel(writer, sheet_name=sheet_name, index=False)
