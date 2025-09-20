import yfinance as yf
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import simpledialog, messagebox
import os

def get_price_column(data):
    """Helper to safely extract price series."""
    if isinstance(data.columns, pd.MultiIndex):
        data.columns = [col[0] for col in data.columns]

    if "Adj Close" in data.columns:
        return data["Adj Close"]
    elif "Close" in data.columns:
        return data["Close"]
    else:
        raise ValueError("No valid 'Close' or 'Adj Close' column found in data.")

def fetch_and_calculate(ticker):
    try:
        # Stock data
        data = yf.download(ticker, period="1y")
        price = get_price_column(data)

        df = pd.DataFrame({
            "Date": price.index,
            "Price": price.values
        })

        # Returns
        df["Diff_Return"] = df["Price"].diff()
        df["Relative_Return"] = df["Price"].pct_change()
        df["Log_Return"] = np.log(df["Price"] / df["Price"].shift(1))

        # Risk metrics
        uncertainty = df["Log_Return"].std() * np.sqrt(252)

        # Market data (S&P 500 as proxy)
        market = yf.download("^GSPC", period="1y")
        market_price = get_price_column(market)
        market_ret = market_price.pct_change().dropna()

        stock_ret = df["Relative_Return"].dropna()
        aligned = pd.concat([stock_ret, market_ret], axis=1).dropna()
        aligned.columns = ["Stock", "Market"]

        cov = np.cov(aligned["Stock"], aligned["Market"])[0][1]
        beta = cov / aligned["Market"].var()

        rf = 0.02
        market_rp = 0.05
        required_return = rf + beta * market_rp

        var_95 = np.percentile(df["Relative_Return"].dropna(), 5)
        cvar_95 = df["Relative_Return"].dropna()[df["Relative_Return"] <= var_95].mean()

        mean_ret = df["Relative_Return"].mean()
        std_ret = df["Relative_Return"].std()
        norm_var_95 = mean_ret - 1.65 * std_ret

        # Save Excel
        filename = f"{ticker}_returns.xlsx"
        df.to_excel(filename, index=False)

        # Interpretation
        interpretation = (
            f"Uncertainty (Volatility): {uncertainty:.2%}\n"
            f"Beta vs Market: {beta:.2f}\n"
            f"Required Return (CAPM): {required_return:.2%}\n"
            f"VaR (95%): {var_95:.2%}\n"
            f"CVaR (95%): {cvar_95:.2%}\n"
            f"Normal Approx. VaR (95%): {norm_var_95:.2%}\n\n"
            "Interpretation:\n"
            f"- A Beta of {beta:.2f} means the stock is "
            f"{'more' if beta > 1 else 'less'} volatile than the market.\n"
            "- Required return shows the compensation investors expect for risk.\n"
            "- VaR and CVaR estimate worst expected losses, useful for risk management."
        )

        return interpretation, filename

    except Exception as e:
        return f"Failed to fetch data: {e}", None

# GUI
def run_gui():
    root = tk.Tk()
    root.withdraw()
    ticker = simpledialog.askstring("Stock Input", "Enter Stock Ticker Symbol (e.g., AAPL):")

    if ticker:
        result, file = fetch_and_calculate(ticker)
        if file:
            messagebox.showinfo("Success", f"{result}\n\nResults saved to {os.path.abspath(file)}")
        else:
            messagebox.showerror("Error", result)

if __name__ == "__main__":
    run_gui()
