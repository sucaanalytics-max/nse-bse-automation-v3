from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import date
from scripts.excel_utils import append_to_sheet

SHEET_NAME = "BSE_CASH"
URL = "https://www.bseindia.com/markets/equity/EQReports/MarketWatch.aspx"

def fetch_bse_cash():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(URL, timeout=60000)
        page.wait_for_timeout(10000)

        table = page.locator("table#ContentPlaceHolder1_gvData")
        rows = table.locator("tr").all()

        data = []
        for r in rows[1:]:
            cols = r.locator("td").all_text_contents()
            if cols:
                data.append(cols)

        browser.close()

        df = pd.DataFrame(data)
        df["date"] = date.today()
        return df

if __name__ == "__main__":
    df = fetch_bse_cash()
    append_to_sheet(SHEET_NAME, df)
    print("✅ BSE Cash appended")
