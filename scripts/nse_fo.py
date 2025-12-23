from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import date
from scripts.excel_utils import append_unique

SHEET = "NSE_FO"

def fetch_nse_fo():
    today = date.today()
    month = today.strftime("%b")
    year = today.year

    url = f"https://www.nseindia.com/api/historicalOR/fo/tbg/daily?month={month}&year={year}"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("https://www.nseindia.com", timeout=60000)
        page.wait_for_timeout(3000)

        response = page.request.get(url)
        data = response.json()

        browser.close()

    df = pd.DataFrame(data["data"])
    df["Date"] = today
    return df

if __name__ == "__main__":
    df = fetch_nse_fo()
    append_unique(SHEET, df)
    print("✅ NSE FO appended")
