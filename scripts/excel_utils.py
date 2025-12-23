import pandas as pd
from openpyxl import load_workbook

EXCEL_PATH = "data/nse_bse_business_growth.xlsx"

def append_unique(sheet_name: str, new_df: pd.DataFrame, date_col="Date"):
    book = load_workbook(EXCEL_PATH)

    if sheet_name in book.sheetnames:
        old_df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)

        if date_col in old_df.columns:
            old_dates = set(old_df[date_col].astype(str))
            new_df = new_df[~new_df[date_col].astype(str).isin(old_dates)]

        final_df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        final_df = new_df

    with pd.ExcelWriter(
        EXCEL_PATH,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        writer.book = book
        final_df.to_excel(writer, sheet_name=sheet_name, index=False)
