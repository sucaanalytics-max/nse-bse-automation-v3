from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import date
from scripts.excel_utils import append_to_sheet

SHEET_NAME = "NSE_CASH"   # ⚠️ CHANGE ONLY if your Excel uses a different name
URL = "https://www.nseindia.com/market-data/live-equity-market"

def fetch_nse_cash():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(URL, timeout=60000)
        page.wait_for_timeout(8000)

        data = page.evaluate("""
        () => window.__REDUX_STATE__?.marketWatch?.marketData || []
        """)

        browser.close()

        if not data:
            raise RuntimeError("NSE data not loaded")

        df = pd.DataFrame(data)
        df["date"] = date.today()
        return df

if __name__ == "__main__":
    df = fetch_nse_cash()
    append_to_sheet(SHEET_NAME, df)
    print("✅ NSE Cash appended")
