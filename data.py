import requests
import pandas as pd
import schedule
import time

# API URL and Headers
API_URL = "https://api.coingecko.com/api/v3/coins/markets"
PARAMS = {
    "vs_currency": "usd",
    "order": "market_cap_desc",
    "per_page": 50,
    "page": 1,
    "sparkline": False,
}

# Fetch Data from API
def fetch_crypto_data():
    response = requests.get(API_URL, params=PARAMS)
    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error fetching data: {response.status_code}")
        return []

# Save Data to Excel
def save_to_excel(data):
    df = pd.DataFrame(data)
    df = df[[
        "name", "symbol", "current_price", "market_cap",
        "total_volume", "price_change_percentage_24h"
    ]]
    df.columns = [
        "Cryptocurrency Name", "Symbol", "Current Price (USD)",
        "Market Capitalization", "24-hour Trading Volume",
        "Price Change (24-hour, %)"
    ]
    df.to_excel("crypto_data.xlsx", index=False)
    print("Data saved to crypto_data.xlsx")

# Perform Data Analysis
def analyze_data(data):
    df = pd.DataFrame(data)
    df = df[[
        "name", "current_price", "market_cap", "price_change_percentage_24h"
    ]]
    
    # Top 5 Cryptocurrencies by Market Cap
    top_5 = df.nlargest(5, "market_cap")[["name", "market_cap"]]

    # Average Price of Top 50 Cryptocurrencies
    avg_price = df["current_price"].mean()

    # Highest and Lowest Percentage Price Change
    highest_change = df.nlargest(1, "price_change_percentage_24h")[["name", "price_change_percentage_24h"]]
    lowest_change = df.nsmallest(1, "price_change_percentage_24h")[["name", "price_change_percentage_24h"]]

    # Print Analysis
    print("Top 5 Cryptocurrencies by Market Cap:")
    print(top_5)
    print("\nAverage Price of Top 50 Cryptocurrencies:", avg_price)
    print("\nHighest Percentage Price Change in 24h:")
    print(highest_change)
    print("\nLowest Percentage Price Change in 24h:")
    print(lowest_change)

# Main Function
def main():
    data = fetch_crypto_data()
    if data:
        save_to_excel(data)
        analyze_data(data)

# Schedule Task Every 5 Minutes
schedule.every(5).minutes.do(main)

if __name__ == "__main__":
    main()
    while True:
        schedule.run_pending()
        time.sleep(1)
