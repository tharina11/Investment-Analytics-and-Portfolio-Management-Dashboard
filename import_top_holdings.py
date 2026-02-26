# Import libraries
import yfinance as yf
import pandas as pd

# Function to import ETF holdings data
def get_top_holdings(ticker_list):
    all_holdings = []
    
    for ticker in ticker_list:
        try:
            etf = yf.Ticker(ticker)
            # Fetch holdings (Top 10) and full equity metadata
            holdings = etf.funds_data.top_holdings
            equity_data = etf.funds_data.equity_holdings
            
            if holdings is not None and not holdings.empty:
                # holdings index = Company Name
                # equity_data index = Company Name, 'Symbol' column = Ticker
                if equity_data is not None and 'Symbol' in equity_data.columns:
                    # Map the 'Name' index to the 'Symbol' from equity_data
                    holdings.index = holdings.index.map(equity_data['Symbol'])
                
                # Standardize columns and metadata
                holdings = holdings.reset_index().rename(columns={'index': 'Ticker'})
                holdings['ETF'] = ticker
                all_holdings.append(holdings)
            else:
                print(f"No holdings data found for {ticker}")
                
        except Exception as e:
            print(f"Could not retrieve data for {ticker}: {e}")

    if all_holdings:
        final_df = pd.concat(all_holdings, ignore_index=True)
        final_df = final_df.sort_values(by=['ETF', 'Holding Percent'], ascending=[True, False])
        return final_df
    else:
        return pd.DataFrame()

# Import data for ETFs
tickers = ["VOO", "SPYG", "VTI", "SCHD", "QQQM", "SCHG", "VGT"]
top_holdings_df = get_top_holdings(tickers)

if not top_holdings_df.empty:
    print(top_holdings_df)
    
# Reorder columns by passing a new list
top_holdings_df = top_holdings_df.rename(columns={"Name": "Company Name"})
top_holdings_df = top_holdings_df[["ETF", "Symbol", "Company Name", "Holding Percent"]]

# Save to Excel
top_holdings_df.to_excel("etf_top_holdings.xlsx", index=False)
print("Saved to top_holdings_df.xlsx")
