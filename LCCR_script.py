import requests
import pandas as pd
import xlwings as xw
import schedule
import time

# API Configuration
API_URL = "https://api.coingecko.com/api/v3/coins/markets"
params = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": False
}

file_name = "crypto_data.xlsx"

def fetch_crypto_data():
    """Fetch live cryptocurrency data from CoinGecko API."""
    try:
        response = requests.get(API_URL, params=params)
        response.raise_for_status()  # Raise an error for bad responses
        data = response.json()
        
        # Selecting the required fields
        df = pd.DataFrame(data)[["name", "symbol", "current_price", "market_cap", "total_volume", "price_change_percentage_24h"]]
        
        df.columns = ["CryptoCurrency Name", "Symbol", "Current Price (USD)", "Market Capitalization", "24h Trading Volume", "Price Change % (24h)"]
        
        return df
    except requests.exceptions.RequestException as e:
        print(f"Error fetching data: {e}")
        return None

def analyze_data(df):
    """Perform basic analysis on the data."""
    if df is None or df.empty:
        print("No data available for analysis.")
        return

    top_5 = df.nlargest(5, "Market Capitalization")
    avg_price = df["Current Price (USD)"].mean()

    highest_change = df.loc[df["Price Change (24h, %)"].idxmax()]
    lowest_change = df.loc[df["Price Change (24h, %)"].idxmin()]

    print("\n--- Market Analysis ---")
    print(f"Top 5 Cryptocurrencies:\n{top_5[['Cryptocurrency Name', 'Market Capitalization']]}\n")
    print(f"Average price of top 50 cryptocurrencies: ${avg_price:.2f}")
    print(f"Highest 24h % change: {highest_change['Cryptocurrency Name']} ({highest_change['Price Change (24h, %)']:.2f}%)")
    print(f"Lowest 24h % change: {lowest_change['Cryptocurrency Name']} ({lowest_change['Price Change (24h, %)']:.2f}%)\n")

def update_excel():
    """Fetch live data and update the Excel sheet."""
    df = fetch_crypto_data()
    if df is None:
        return

    try:
        with pd.ExcelWriter(file_name, engine='openpyxl', mode='w') as writer:
            df.to_excel(writer, index=False, sheet_name="Live Data")
        print("Excel updated successfully.")
    except Exception as e:
        print(f"Error updating Excel: {e}")

# Initial Data Fetch
df = fetch_crypto_data()
if df is not None:
    analyze_data(df)
    update_excel()

# Schedule updates every 5 minutes
schedule.every(2).minutes.do(update_excel)

print("Live updating started. Press Ctrl+C to stop.")

while True:
    schedule.run_pending()
    time.sleep(schedule.idle_seconds())  

