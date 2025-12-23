from playwright.sync_api import sync_playwright
import pandas as pd
from scripts.utils import write_excel, today_str

URL = "https://www.nseindia.com/market-data/live-equity-market"

def fetch_nse():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(URL, timeout=60000)
        page.wait_for_timeout(8000)

        data = page.evaluate("""
        () => {
            return window.__REDUX_STATE__?.marketWatch?.marketData || [];
        }
        """)

        browser.close()

        if not data:
            raise RuntimeError("NSE data not loaded")

        df = pd.DataFrame(data)
        return df

if __name__ == "__main__":
    df = fetch_nse()
    df["date"] = today_str()
    write_excel("NSE_CASH", df)
    print("✅ NSE Cash written")
