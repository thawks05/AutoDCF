# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the app

```bash
pip install yfinance pandas openpyxl PyQt5
python "import yfinance as yf.py"
```

The file name is literally `import yfinance as yf.py` — always quote it in shell commands.

The logo is loaded from a hardcoded path (`/Users/timothyh/Desktop/ICDCFAUTO/IC logo.png`); if it's missing the app shows a text fallback and continues normally.

## What this project is

A single-file PyQt5 desktop app ("GIR DCF Autocomplete") that pulls 5 years of annual financials plus LTM (trailing twelve months) from Yahoo Finance (`yfinance`) and writes them into an existing Excel DCF workbook. The user supplies a ticker, an `.xlsx` path, and a sheet name; the tool handles the rest.

## Architecture

Everything lives in one file. The three layers are:

**Data pipeline** — `get_yahoo_financials(ticker, excel_path, sheet_name)`  
Fetches annual statements (income, cash flow, balance sheet) and the last 4 quarters for TTM. Assembles a `years_data` list of dicts — one dict per fiscal year plus one `'LTM'` dict. Derived metrics are computed here:
- `EBT = EBIT − Interest Expense`
- `Net Debt = Total Debt − Cash`
- `FCFF = EBIT*(1−0.375) + D&A + Amort + Capex + PurchaseIntangibles − NonOpInterest − ΔWorkingCapital`
- `FCFE` adds back Interest Expense, Debt Issuance, and Debt Repayment to FCFF

The tax rate in the FCFF/FCFE formula is hardcoded at 37.5%.

Many Yahoo Finance fields have multiple possible names across tickers; `extract_value(df, row_names)` tries each name in order.

**Excel writer** — `write_to_excel(years_data, file_path, sheet_name)`  
Rather than writing to a fixed cell range, this function *scans* the sheet to locate:
1. The earliest year header (e.g. `"2020"` or `"2020A"`) to anchor the column offset
2. Each metric row by case-insensitive label match against `metric_order`
3. An `LTM`/`TTM` column for the trailing-twelve-months data

Internal dict keys don't always match Excel row labels; `write_to_excel` maps them explicitly (`'Capex'` → `'LCapex'`, `'Stock Price'` → `'share price'`, etc.). All monetary values are divided by 1,000,000 before writing; multiples and share prices are written as-is.

**GUI** — `FinancialDataGUI(QMainWindow)` + `DataFetchThread(QThread)`  
Standard PyQt5 form. `DataFetchThread` runs `get_yahoo_financials` in a background thread, redirecting stdout/stderr into `StringIO` buffers so the green console widget can display progress. The `finished` signal carries `(success: bool, output: str, error_output: str)`.

## Key data contract

`years_data` is a `list[dict]` where each dict has:
- `'Year'`: ISO date string like `'2023-12-31'`, or the literal string `'LTM'`
- All financial metric keys as listed in `get_yahoo_financials` (Revenue, COGS, EBITDA, FCFF, FCFE, Net Debt, Stock Price, multiple, …)

The `metric_order` list in `write_to_excel` and `self.metric_mapping` in the GUI must stay in sync with the keys populated in `get_yahoo_financials`.
