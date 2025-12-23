from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import date
from scripts.excel_utils import append_unique

SHEET = "BSE_FO"
URL = "https://www.bseindia.com/markets/keystatics/Keystat_turnover_deri.aspx"

def fetch_bse_fo():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(URL, timeout=60000)
        page.wait_for_timeout(5000)

        table = page.locator("table").nth(0)
        rows = table.locator("tr").all()

        data = []
        for r in rows[1:]:
            cols = r.locator("td").all_text_contents()
            if cols:
                data.append(cols)

        browser.close()

    df = pd.DataFrame(data)
    df["Date"] = date.today()
    return df

if __name__ == "__main__":
    df = fetch_bse_fo()
    append_unique(SHEET, df)
    print("✅ BSE FO appended")
