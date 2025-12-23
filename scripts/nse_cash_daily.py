import requests
import pandas as pd
from playwright.sync_api import sync_playwright
from datetime import date
from scripts.excel_utils import append_df

SHEET = "NSE_Cash"

def get_session():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("https://www.nseindia.com/market-data/business-growth-cm-segment", timeout=60000)
        page.wait_for_timeout(5000)
        cookies = page.context.cookies()
        browser.close()

    s = requests.Session()
    for c in cookies:
        s.cookies.set(c["name"], c["value"])
    return s

def fetch_daily():
    today = date.today()
    s = get_session()

    url = "https://www.nseindia.com/api/historicalOR/cm/tbg/daily"
    params = {
        "month": today.strftime("%b"),
        "year": today.strftime("%y")
    }

    r = s.get(url, timeout=30)
    r.raise_for_status()

    data = r.json()["data"]
    df = pd.DataFrame(data)
    df["RunDate"] = today
    return df

if __name__ == "__main__":
    df = fetch_daily()
    append_df(SHEET, df)
    print("✅ NSE Cash appended")
