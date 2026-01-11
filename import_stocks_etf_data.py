import yfinance as yf
import pandas as pd

# Path to Excel file
file_path = "holdings.xlsx"

# Read Excel
df_holdings = pd.read_excel(file_path)

# Create ticker list
tickers = df_holdings["Ticker"].tolist()

# Create dictionary: {Ticker: Number of stocks}
number_of_stocks = dict(
    zip(df_holdings["Ticker"], df_holdings["Number of stocks"])
)

# Drop rows with missing values
df_holdings = df_holdings.dropna(subset=["Ticker", "Number of stocks"])

# Ensure correct types
df_holdings["Ticker"] = df_holdings["Ticker"].astype(str)
df_holdings["Number of stocks"] = df_holdings["Number of stocks"].astype(float)

# Expense ratios (Yahoo does not provide these for ETFs)
expense_ratios = {
    "VOO": 0.03,
    "VTI": 0.03,
    "VGT": 0.10,
    "SPYG": 0.04,
    "QQQM": 0.15,
    "SCHD": 0.06,
    "SCHG": 0.04
}

# Import stocks and ETF prices
def get_price(stock):
    try:
        price = stock.fast_info.get("last_price")
        if price is not None:
            return price
    except:
        pass

    # Fallback: use last close from history
    try:
        hist = stock.history(period="1d")
        if len(hist) > 0:
            return hist["Close"].iloc[0]
    except:
        pass

    return None

# Fetch prices from Yahoo Finance
def fetch_data(ticker):
    stock = yf.Ticker(ticker)
    price = get_price(stock)
    info = stock.info

    return {
        "Ticker": ticker,
        "Price": price,
        "Sector": info.get("sector"),
        "PE Ratio": info.get("trailingPE"),
        "Market Cap": info.get("marketCap"),
        "Beta": info.get("beta"),
        "Expense Ratio": expense_ratios.get(ticker, None),
        
        # NEW COLUMN:
        "Number of Stocks": number_of_stocks.get(ticker, 0)
    }


records = [fetch_data(t) for t in tickers]
df = pd.DataFrame(records)

# Automatically calculate Market Value
df["Market value"] = df["Price"] * df["Number of Stocks"]

# Add empty columns for manual input
df["Growth Class"] = None
df["Type"] = None
df["PE low"] = None
df["PE high"] = None

# Manually fill Growth Class, Type, PE low, PE high using dictionaries
growth_class_dict = {
    "VOO": "Broad Market",
    "VTI": "Broad Market",
    "VGT": "Aggresive Growth",
    "SPYG": "Core Growth",
    "QQQM": "Core Growth",
    "SCHD": "Defensive",
    "SCHG": "Core Growth",
    "O": "Defensive",
    "VICI": "Defensive",
    "MSFT": "Core Growth",
    "NVDA": "Aggresive Growth",
    "AVGO": "Aggresive Growth",
    "KO": "Defensive",
    "UNH": "Defensive",
    "GOOGL": "Aggresive Growth",
    "PLTR": "Aggresive Growth",
    "F": "Defensive",
    "ET": "Defensive",
    "TSM": "Aggresive Growth",
    "ASML": "Aggresive Growth",
    "META": "Aggresive Growth",
    "PG": "Defensive"
}

type_dict = {
    "VOO": "ETF",
    "VTI": "ETF",
    "VGT": "ETF",
    "SPYG": "ETF",
    "QQQM": "ETF",
    "SCHD": "ETF",
    "SCHG": "ETF",
    "O": "REIT",
    "VICI": "REIT",
    "MSFT": "Stock",
    "NVDA": "Stock",
    "AVGO": "Stock",
    "KO": "Stock",
    "UNH": "Stock",
    "GOOGL": "Stock",
    "PLTR": "Stock",
    "F": "Stock",
    "ET": "Stock",
    "TSM": "Stock",
    "ASML": "Stock",
    "META": "Stock",
    "PG": "Stock"
}

pe_low_dict = {
    "MSFT": 28, 
    "NVDA": 35, 
    "AVGO": 24, 
    "KO": 21, 
    "UNH": 18,
    "GOOGL": 22, 
    "F": 7, 
    "TSM": 15, 
    "ASML": 30, 
    "META": 20,
    "PG": 21
}

pe_high_dict = {
    "MSFT": 35, 
    "NVDA": 50, 
    "AVGO": 30, 
    "KO": 25, 
    "UNH": 22,
    "GOOGL": 28, 
    "F": 10, 
    "TSM": 20, 
    "ASML": 40, 
    "META": 26,
    "PG": 24
}

# Apply values from dictionaries


df["Growth Class"] = df["Ticker"].map(growth_class_dict)
df["Type"] = df["Ticker"].map(type_dict)
df["PE low"] = df["Ticker"].map(pe_low_dict)
df["PE high"] = df["Ticker"].map(pe_high_dict)

# (Optional) Fill missing PE ranges with blanks instead of NaN
df.fillna({"PE low": "", "PE high": ""}, inplace=True)


# List of ETF tickers
etf_list = ["VOO", 
            "VTI", 
            "VGT", 
            "SPYG", 
            "QQQM", 
            "SCHD", 
            "SCHG"]

# Assign Sector = "Index" for these tickers
df.loc[df["Ticker"].isin(etf_list), "Sector"] = "Index"

# Calculate Percentage of each holding in the portfolio
df["Percentage"] = (df["Market value"] / df["Market value"].sum())

# Import EPS CAGR for 3 years
def calculate_eps_cagr_3y(ticker):
    """
    Fetches EPS data for a ticker and calculates 3-year EPS CAGR.
    Returns None if data is missing or invalid.
    """
    try:
        stock = yf.Ticker(ticker)

        # Get yearly EPS (most reliable source)
        eps_df = stock.income_stmt

        if eps_df is None or "Diluted EPS" not in eps_df.index:
            return None

        eps_series = eps_df.loc["Diluted EPS"].dropna()

        # Ensure newest -> oldest
        eps_series = eps_series.sort_index(ascending=False)

        if len(eps_series) < 4 or eps_series.iloc[3] <= 0:
            return None

        return (eps_series.iloc[0] / eps_series.iloc[3]) ** (1/3) - 1

    except Exception:
        return None

df["eps_CAGR_3y"] = df["Ticker"].apply(calculate_eps_cagr_3y)


def calculate_fcf_cagr_3y(ticker):
    """
    Calculates 3-year Free Cash Flow CAGR using Yahoo Finance data.
    Returns None if data is missing or invalid.
    """
    try:
        stock = yf.Ticker(ticker)
        cashflow_df = stock.cashflow

        if cashflow_df is None or cashflow_df.empty:
            return None

        # Possible label variations
        ocf_labels = [
            "Operating Cash Flow",
            "Net Cash Provided by Operating Activities",
            "Net Cash from Operating Activities",
            "Cash Flow from Operations"
        ]

        capex_labels = [
            "Capital Expenditures",
            "Capital Expenditure",
            "Purchase of PPE"
        ]

        # Find matching labels
        ocf_label = next((l for l in ocf_labels if l in cashflow_df.index), None)
        capex_label = next((l for l in capex_labels if l in cashflow_df.index), None)

        if ocf_label is None or capex_label is None:
            return None

        # Calculate Free Cash Flow
        fcf_series = (
            cashflow_df.loc[ocf_label]
            + cashflow_df.loc[capex_label]  # CapEx is negative
        ).dropna()

        # Ensure newest → oldest
        fcf_series = fcf_series.sort_index(ascending=False)

        if len(fcf_series) < 4 or fcf_series.iloc[3] <= 0:
            return None

        return (fcf_series.iloc[0] / fcf_series.iloc[3]) ** (1/3) - 1

    except Exception:
        return None


df["fcf_CAGR_3y"] = df["Ticker"].apply(calculate_fcf_cagr_3y)

# 6. Save to Excel

df.to_excel("stock_fundamentals.xlsx", index=False)

print("Saved to stock_fundamentals.xlsx")
