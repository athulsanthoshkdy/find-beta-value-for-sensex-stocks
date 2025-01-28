import yfinance as yf
import numpy as np
import pandas as pd

# List of companies in Sensex (tickers)
sensex_companies = ["ADANIPORTS.NS", "ASIANPAINT.NS", "AXISBANK.NS", "BAJAJ-AUTO.NS", "BAJFINANCE.NS",
        "BAJAJFINSV.NS", "BHARTIARTL.NS", "BPCL.NS", "CIPLA.NS", "COALINDIA.NS", "DIVISLAB.NS",
        "DRREDDY.NS", "EICHERMOT.NS", "GRASIM.NS", "HCLTECH.NS", "HDFCBANK.NS",
        "HDFCLIFE.NS", "HEROMOTOCO.NS", "HINDALCO.NS", "HINDUNILVR.NS", "ICICIBANK.NS",
        "INDUSINDBK.NS", "INFY.NS", "IOC.NS", "ITC.NS", "JSWSTEEL.NS", "KOTAKBANK.NS",
        "LT.NS", "M&M.NS", "NESTLEIND.NS", "NTPC.NS", "ONGC.NS", "POWERGRID.NS",
        "RELIANCE.NS", "SBILIFE.NS", "SBIN.NS", "SUNPHARMA.NS", "TATAMOTORS.NS",
        "TATASTEEL.NS", "TCS.NS", "TECHM.NS", "TITAN.NS", "UPL.NS",
        "WIPRO.NS", "ABB.NS", "ACC.NS", "ADANIGREEN.NS", "AMBUJACEM.NS",
        "APOLLOHOSP.NS", "AUROPHARMA.NS", "BANDHANBNK.NS", "BANKBARODA.NS", "BEL.NS",
        "BERGEPAINT.NS", "BIOCON.NS", "CHOLAFIN.NS", "CONCOR.NS",
        "DABUR.NS", "DALBHARAT.NS", "DEEPAKNTR.NS", "DLF.NS", "EXIDEIND.NS",
        "GAIL.NS", "GODREJCP.NS", "HAVELLS.NS", "HINDPETRO.NS", "IGL.NS"]  # Add all 30 tickers
benchmark = "^BSESN"  # Sensex Index

# Function to calculate beta
def calculate_beta(stock_returns, benchmark_returns):
    covariance = np.cov(stock_returns, benchmark_returns)[0, 1]
    variance = np.var(benchmark_returns)
    return covariance / variance

# Download historical data
data = yf.download(sensex_companies + [benchmark], start="2018-01-01", end="2024-01-01", interval="1d")
returns = data["Close"].pct_change().dropna()

# Slicing data for 5 years, 3 years, and 1 year
end_date = returns.index[-1]
returns_5y = returns[returns.index >= (end_date - pd.DateOffset(years=5))]
returns_3y = returns[returns.index >= (end_date - pd.DateOffset(years=3))]
returns_1y = returns[returns.index >= (end_date - pd.DateOffset(years=1))]

# Calculate beta for each time period
beta_5y = {}
beta_3y = {}
beta_1y = {}

for company in sensex_companies:
    # Beta for 5 years
    stock_returns_5y = returns_5y[company]
    benchmark_returns_5y = returns_5y[benchmark]
    beta_5y[company] = calculate_beta(stock_returns_5y, benchmark_returns_5y)
    
    # Beta for 3 years
    stock_returns_3y = returns_3y[company]
    benchmark_returns_3y = returns_3y[benchmark]
    beta_3y[company] = calculate_beta(stock_returns_3y, benchmark_returns_3y)
    
    # Beta for 1 year
    stock_returns_1y = returns_1y[company]
    benchmark_returns_1y = returns_1y[benchmark]
    beta_1y[company] = calculate_beta(stock_returns_1y, benchmark_returns_1y)

# Create a DataFrame for results
beta_df = pd.DataFrame({
    "5Y Beta": beta_5y,
    "3Y Beta": beta_3y,
    "1Y Beta": beta_1y
}).transpose()

# Calculate changes in beta
beta_df.loc["Change 5Y-3Y"] = beta_df.loc["5Y Beta"] - beta_df.loc["3Y Beta"]
beta_df.loc["Change 3Y-1Y"] = beta_df.loc["3Y Beta"] - beta_df.loc["1Y Beta"]

# Transpose the DataFrame and save to Excel file
output_file = "beta_values_sensex_transposed.xlsx"
beta_df.T.to_excel(output_file)

print(f"Beta values and changes saved to {output_file} in transposed format")
