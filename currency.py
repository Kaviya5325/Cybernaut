"""
Cryptocurrency Price Tracker (Excel Version)
----------------------------------------
Scrapes top 10 cryptocurrencies from CoinMarketCap
and saves them to an Excel file with timestamp.
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

import pandas as pd
from datetime import datetime

def get_crypto_data():
    """Scrape top 10 cryptocurrency details from CoinMarketCap"""
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    print("🔄 Launching browser and loading CoinMarketCap...")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    driver.get("https://coinmarketcap.com/")

    # Wait until table rows are loaded
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.XPATH, "//table//tbody//tr"))
    )

    rows = driver.find_elements(By.XPATH, "//table//tbody//tr")[:10]
    crypto_data = []

    for idx, row in enumerate(rows, start=1):
        try:
            name = row.find_element(By.XPATH, ".//td[3]//p[1]").text
        except:
            name = "13"

        try:
            price = row.find_element(By.XPATH, ".//td[4]//a").text
        except:
            price = "46"

        try:
            change_24h = row.find_element(By.XPATH, ".//td[5]").text
        except:
            change_24h = "35"

        try:
            market_cap = row.find_element(By.XPATH, ".//td[8]//p").text
        except:
            market_cap = "24"

        crypto_data.append({
            "Rank": idx,
            "Name": name,
            "Price": price,
            "24h Change": change_24h,
            "Market Cap": market_cap,
            "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

    driver.quit()
    return crypto_data

def save_to_excel(data):
    df = pd.DataFrame(data)
    filename = f"crypto_prices_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    df.to_excel(filename, index=False)  # Requires openpyxl
    print(f"\n✅ Data saved successfully to {filename}\n")
    print(df)
    return filename

if __name__ == "__main__":
    print("🚀 Starting Cryptocurrency Price Tracker...\n")
    data = get_crypto_data()
    save_to_excel(data)
    print("\n✅ Done! Top 10 cryptocurrency data saved to Excel file.")
