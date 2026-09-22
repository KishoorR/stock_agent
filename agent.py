"""
Streamlit portfolio tracker — Excel-as-source-of-truth version.

Run with:
    streamlit run stock_excel_app.py

WHY THIS VERSION EXISTS
-------------------------
The Groww API versions (stock_groww.py / stock_groww_app.py) kept returning
"Access forbidden" on every market-data call (LTP, historical candles) because
that requires an ACTIVE, currently-billed Trading API subscription on the
Groww account — separate from having valid credentials. Holdings/portfolio
reads work without it; live/historical prices don't.

This version sidesteps that entirely:
  - Holdings (symbol, quantity, purchase price) come from an uploaded Excel
    file you maintain yourself (e.g. your broker's holdings statement, or an
    export from Groww's website) — no live API call needed for this part.
  - Current price + 1D/1W/1M/3M/6M/1Y change come from Yahoo Finance via
    yfinance — a free, standard, no-signup data source.

EXCEL FORMAT
-------------
Any of these column names are auto-detected (case-insensitive):
  - Symbol/name column: "trading_symbol", "stock symbol", "symbol",
    "stock name", "scrip name", "instrument"
  - Quantity column: "quantity", "qty", "number of shares", "shares"
  - Purchase price column: "average price", "average buy price",
    "avg price", "purchase price", "buy price"
Optionally add a "ticker" or "yahoo ticker" column with the exact Yahoo
Finance symbol (e.g. "RELIANCE.NS") to override auto-resolution for any row.
An "ISIN" column, if present, is used automatically too (see below).

Ticker resolution priority per row (fully automated — no hand-maintained list):
  1. Explicit ticker/yahoo_ticker column value, if present.
  2. ISIN column match against NSE's own official, publicly downloadable equity
     master list (SYMBOL, company name, ISIN) — the most reliable path, since
     ISIN is a globally unique identifier with no ambiguity. This list is
     fetched live and cached for a day; if the fetch fails (no network, NSE
     blocking the request, etc.), this step is skipped for that run.
  3. If the symbol looks like a short exchange code (no spaces, e.g. "HSCL",
     "TMCV") — try "{symbol}.NS" then "{symbol}.BO".
  4. If it looks like a full company name (has spaces, e.g. "ADANI PORT &
     SEZ LTD") — fuzzy-match it against the same NSE master list's company
     names (using difflib, no extra dependency). Fuzzy matches are logged as
     an [INFO] warning showing the official name matched against, so you can
     sanity-check anything less certain than an exact ISIN match.
  Unresolved names are skipped with a warning listing exactly what wasn't
  matched — add a "ticker" column for those rows to bypass resolution entirely.

PREREQUISITES
--------------
pip install streamlit yfinance pandas openpyxl requests
"""

import difflib
import io
import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
from openpyxl.utils import get_column_letter

try:
    import anthropic

    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

try:
    import openai

    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

try:
    from google import genai
    from google.genai import types as genai_types

    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

try:
    import yfinance as yf
except ImportError:
    st.error("Missing dependency. Install with:\n\n    pip install yfinance")
    st.stop()


# --------------------------------------------------------------------------------------
# Config
# --------------------------------------------------------------------------------------
LOOKBACK_WINDOWS = {
    "1D Change (%)": 1,
    "1W Change (%)": 7,
    "1M Change (%)": 30,
    "3M Change (%)": 90,
    "6M Change (%)": 180,
    "1Y Change (%)": 365,
}
RECENT_BUY_MONTHS = 6

SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_PATH = SCRIPT_DIR / "stock_portfolio.xlsx"

# NSE's own official, publicly downloadable master list of every listed equity —
# SYMBOL, company name, and ISIN. This replaces a hand-maintained name->ticker dict:
# resolution is now automated by looking up your Excel's ISIN (most reliable) or
# fuzzy-matching the company name against this real, current list, instead of only
# working for names someone happened to add to a static dict ahead of time.
NSE_EQUITY_MASTER_URL = "https://nsearchives.nseindia.com/content/equities/EQUITY_L.csv"
NSE_MASTER_CACHE_TTL_SECONDS = 24 * 60 * 60  # the list changes rarely; refetch daily at most


@st.cache_data(ttl=NSE_MASTER_CACHE_TTL_SECONDS, show_spinner=False)
def fetch_nse_equity_master():
    """Download and parse NSE's official equity master list. Returns a DataFrame
    with columns SYMBOL, NAME, ISIN (normalized), or None if the fetch fails —
    callers must handle that by falling back to short-code-only resolution."""
    try:
        resp = requests.get(
            NSE_EQUITY_MASTER_URL,
            headers={"User-Agent": "Mozilla/5.0 (compatible; PortfolioTracker/1.0)"},
            timeout=15,
        )
        resp.raise_for_status()
        df = pd.read_csv(io.StringIO(resp.text))
        df.columns = [c.strip().upper() for c in df.columns]
        df = df.rename(columns={"NAME OF COMPANY": "NAME", "ISIN NUMBER": "ISIN"})
        df["NAME_NORM"] = df["NAME"].astype(str).map(_normalize_company_name)
        df["ISIN"] = df["ISIN"].astype(str).str.strip()
        return df[["SYMBOL", "NAME", "NAME_NORM", "ISIN"]]
    except Exception:  # noqa: BLE001 - any network/parse failure just disables this lookup
        return None


def _normalize_company_name(name):
    """Uppercase, collapse whitespace, and strip common corporate-suffix noise so
    abbreviated broker-statement names ('GMR POW AND URBAN INFRA L') compare more
    fairly against official full names ('GMR Power And Urban Infra Limited')."""
    name = re.sub(r"[.,]", " ", str(name).upper())
    name = re.sub(r"\s+", " ", name).strip()
    for suffix in (" LIMITED", " LTD", " LTD ", " CO LTD", " PVT LTD", " PRIVATE LIMITED"):
        if name.endswith(suffix.strip()):
            name = name[: -len(suffix.strip())].strip()
    return name


def lookup_symbol_by_isin(isin, master_df):
    if master_df is None or not isin or str(isin).strip().lower() == "nan":
        return None
    isin = str(isin).strip()
    match = master_df[master_df["ISIN"] == isin]
    if not match.empty:
        return match.iloc[0]["SYMBOL"]
    return None


def lookup_symbol_by_name(name, master_df, cutoff=0.72):
    """Fuzzy-match a company name against the NSE master list using difflib (no
    extra dependency). Returns (symbol, matched_official_name) or (None, None)."""
    if master_df is None:
        return None, None
    normalized = _normalize_company_name(name)
    choices = master_df["NAME_NORM"].tolist()
    best = difflib.get_close_matches(normalized, choices, n=1, cutoff=cutoff)
    if not best:
        return None, None
    row = master_df[master_df["NAME_NORM"] == best[0]].iloc[0]
    return row["SYMBOL"], row["NAME"]


# --------------------------------------------------------------------------------------
# Column detection & ticker resolution
# --------------------------------------------------------------------------------------
def _find_column(columns, must_include, must_not_include=()):
    for col in columns:
        low = col.lower()
        if all(term in low for term in must_include) and not any(
            term in low for term in must_not_include
        ):
            return col
    return None


def find_header_row(raw_df, max_rows_to_scan=40):
    """Many broker/Groww exports have a metadata preamble (account name, a
    summary block with Invested/Closing Value, etc.) above the real table.
    Scan the first few rows for one that looks like an actual header — i.e.
    contains a quantity-like term plus a price-or-name-like term — and
    return its row index. Returns None if no such row is found (in which
    case the file is probably already header-first)."""
    quantity_terms = ("quantity", "qty", "shares")
    other_terms = ("price", "symbol", "name", "scrip", "instrument")

    rows_to_scan = min(max_rows_to_scan, len(raw_df))
    for i in range(rows_to_scan):
        cells = [str(v).strip().lower() for v in raw_df.iloc[i].tolist() if pd.notna(v)]
        has_quantity = any(any(term in cell for term in quantity_terms) for cell in cells)
        has_other = any(any(term in cell for term in other_terms) for cell in cells)
        if has_quantity and has_other:
            return i
    return None


def load_holdings_table(file_obj, sheet_name):
    """Read a sheet, auto-skipping any metadata preamble above the real table."""
    file_obj.seek(0)
    raw_df = pd.read_excel(file_obj, sheet_name=sheet_name, header=None)
    header_row = find_header_row(raw_df)
    if header_row is None:
        # No preamble detected — assume the file is already header-first.
        file_obj.seek(0)
        return pd.read_excel(file_obj, sheet_name=sheet_name)

    header_values = raw_df.iloc[header_row].tolist()
    data = raw_df.iloc[header_row + 1 :].copy()
    data.columns = [str(v).strip() if pd.notna(v) else f"Unnamed_{i}" for i, v in enumerate(header_values)]
    data = data.dropna(how="all")
    return data.reset_index(drop=True)


def detect_columns(df):
    columns = list(df.columns)

    ticker_col = _find_column(columns, ["ticker"]) or _find_column(columns, ["yahoo"])
    isin_col = _find_column(columns, ["isin"])
    symbol_col = (
        _find_column(columns, ["trading", "symbol"])
        or _find_column(columns, ["stock", "symbol"])
        or _find_column(columns, ["symbol"])
        or _find_column(columns, ["stock", "name"])
        or _find_column(columns, ["scrip"])
        or _find_column(columns, ["instrument"])
    )
    qty_col = (
        _find_column(columns, ["quantity"])
        or _find_column(columns, ["qty"])
        or _find_column(columns, ["number", "shares"])
        or _find_column(columns, ["shares"])
    )
    price_col = (
        _find_column(columns, ["average", "price"])
        or _find_column(columns, ["average", "buy"])
        or _find_column(columns, ["avg", "price"])
        or _find_column(columns, ["purchase", "price"])
        or _find_column(columns, ["buy", "price"], must_not_include=["current", "closing", "close"])
    )
    return ticker_col, symbol_col, qty_col, price_col, isin_col


def looks_like_short_code(symbol):
    """Heuristic: short exchange codes have no spaces and are reasonably short
    (e.g. 'HSCL', 'TMCV', 'GMRP&UI'); full company names have spaces."""
    return " " not in symbol.strip() and len(symbol.strip()) <= 15


def resolve_ticker(symbol, explicit_ticker=None, isin=None, master_df=None):
    """Return (candidates, note) — candidates is a list of Yahoo Finance ticker
    guesses to try in order, or None if unresolvable. `note` explains how the
    match was made (useful to show for anything less certain than an exact ISIN
    or explicit-ticker match, e.g. a fuzzy name match)."""
    if explicit_ticker and str(explicit_ticker).strip() and str(explicit_ticker).lower() != "nan":
        return [str(explicit_ticker).strip()], "explicit ticker column"

    isin_symbol = lookup_symbol_by_isin(isin, master_df)
    if isin_symbol:
        return [f"{isin_symbol}.NS", f"{isin_symbol}.BO"], f"ISIN match ({isin})"

    symbol = str(symbol).strip()
    if not symbol:
        return None, None

    if looks_like_short_code(symbol):
        return [f"{symbol}.NS", f"{symbol}.BO"], "short code"

    fuzzy_symbol, matched_name = lookup_symbol_by_name(symbol, master_df)
    if fuzzy_symbol:
        return [f"{fuzzy_symbol}.NS", f"{fuzzy_symbol}.BO"], f"fuzzy name match -> '{matched_name}'"

    return None, None


# --------------------------------------------------------------------------------------
# Pricing / history (yfinance)
# --------------------------------------------------------------------------------------
def price_n_days_ago(hist, last_date, days):
    target_date = last_date - pd.Timedelta(days=days)
    eligible = hist.index[hist.index <= target_date]
    if len(eligible) == 0:
        return None
    return hist.loc[eligible[-1], "Close"]


def pct_change(current, past):
    if past is None or pd.isna(past) or past == 0:
        return None
    return (current - past) / past * 100


def fetch_price_data(ticker_candidates, log=None):
    """Try each candidate ticker (e.g. .NS then .BO) until one returns data.
    Returns (ticker_used, history_df) or (None, None) if all fail."""
    for ticker in ticker_candidates:
        try:
            hist = yf.Ticker(ticker).history(period="2y")
        except Exception as exc:  # noqa: BLE001 - yfinance can raise various network/parse errors
            if log is not None:
                log.append(f"{ticker}: fetch failed ({exc})")
            continue
        if hist is not None and not hist.empty and len(hist) >= 2:
            return ticker, hist.sort_index()
    return None, None


# --------------------------------------------------------------------------------------
# Purchase-date handling (for the recent-buy sheet split)
# --------------------------------------------------------------------------------------
def load_purchase_dates_from_upload(uploaded_file):
    if uploaded_file is None:
        return {}
    try:
        df = pd.read_csv(uploaded_file, dtype=str)
    except Exception as exc:  # noqa: BLE001
        st.warning(f"Could not read the purchase dates file: {exc}")
        return {}
    if "trading_symbol" not in df.columns or "purchase_date" not in df.columns:
        st.warning("CSV must have columns: trading_symbol, purchase_date")
        return {}
    dates = {}
    for _, row in df.iterrows():
        symbol = str(row.get("trading_symbol") or "").strip()
        date_str = str(row.get("purchase_date") or "").strip()
        if not symbol or not date_str or date_str.lower() == "nan":
            continue
        try:
            parsed = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            st.warning(f"Ignoring bad date '{date_str}' for {symbol}")
            continue
        if symbol not in dates or parsed > dates[symbol]:
            dates[symbol] = parsed
    return dates


def make_template_csv_bytes(symbols):
    df = pd.DataFrame({"trading_symbol": sorted(symbols), "purchase_date": ""})
    buf = io.StringIO()
    df.to_csv(buf, index=False)
    return buf.getvalue().encode("utf-8")


# --------------------------------------------------------------------------------------
# Portfolio construction
# --------------------------------------------------------------------------------------
def build_holding_rows(holdings_df, symbol_col, qty_col, price_col, ticker_col, isin_col, progress=None, log=None):
    rows = []
    unresolved = []
    total = len(holdings_df)

    master_df = fetch_nse_equity_master()
    if master_df is None and log is not None:
        log.append(
            "[INFO] Could not fetch NSE's official equity list (network issue?). "
            "ISIN and company-name matching are unavailable this run — only "
            "short exchange codes and explicit ticker columns will resolve."
        )

    for idx, row_dict in enumerate(holdings_df.to_dict(orient="records"), start=1):
        symbol_raw = row_dict.get(symbol_col)
        shares = row_dict.get(qty_col)
        avg_price = row_dict.get(price_col)
        explicit_ticker = row_dict.get(ticker_col) if ticker_col else None
        isin_raw = row_dict.get(isin_col) if isin_col else None

        if pd.isna(symbol_raw) or pd.isna(shares) or pd.isna(avg_price):
            continue

        symbol_display = str(symbol_raw).strip()
        if progress is not None:
            progress.progress(idx / total, text=f"Fetching {symbol_display} ({idx}/{total})")

        candidates, note = resolve_ticker(symbol_raw, explicit_ticker, isin_raw, master_df)
        if not candidates:
            unresolved.append(symbol_display)
            continue
        if log is not None and note and "fuzzy" in note:
            # Fuzzy name matches are the least certain resolution path — surface
            # them so the person can sanity-check instead of silently trusting it.
            log.append(f"[INFO] {symbol_display}: resolved via {note} — verify this is correct.")

        ticker_used, hist = fetch_price_data(candidates, log=log)
        if ticker_used is None:
            if log is not None:
                log.append(f"Skipping {symbol_display}: no price data from Yahoo Finance ({candidates}).")
            continue

        last_date = hist.index[-1]
        current_price = hist["Close"].iloc[-1]

        change_pcts = {
            col: pct_change(current_price, price_n_days_ago(hist, last_date, days))
            for col, days in LOOKBACK_WINDOWS.items()
        }

        shares = float(shares)
        avg_price = float(avg_price)
        total_value = current_price * shares
        gain_loss_value = (current_price - avg_price) * shares
        gain_loss_percent = pct_change(current_price, avg_price)

        row = {
            "Stock Symbol": symbol_display,
            "Yahoo Ticker": ticker_used,
            "Number of Shares": shares,
            "Purchase Price": round(avg_price, 2),
            "Current Price": round(float(current_price), 2),
        }
        row.update({k: (round(float(v), 2) if v is not None and not pd.isna(v) else None) for k, v in change_pcts.items()})
        row.update(
            {
                "Gain/Loss (Value)": round(gain_loss_value, 2),
                "Gain/Loss (%)": round(gain_loss_percent, 2) if gain_loss_percent is not None else None,
                "Total Value": round(total_value, 2),
            }
        )
        rows.append(row)

    if unresolved and log is not None:
        log.append(
            "[UNRESOLVED] Could not map these to a ticker (add a 'ticker' column "
            f"to your Excel to fix): {', '.join(unresolved)}"
        )

    return rows


def make_df_with_total(rows):
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    total_portfolio_value = df["Total Value"].sum()
    total_gain_loss_value = df["Gain/Loss (Value)"].sum()
    total_gain_loss_percent = (
        (total_gain_loss_value / total_portfolio_value) * 100 if total_portfolio_value else 0
    )
    # Use None (-> NaN) for numeric columns, not "", so they stay float dtype —
    # mixing "" into a numeric column makes it dtype=object, which breaks
    # st.dataframe's interactive column sorting (it sorts as text instead of
    # numerically, or may not sort cleanly at all).
    label_cols = {"Stock Symbol", "Yahoo Ticker"}
    total_row = {col: ("" if col in label_cols else float("nan")) for col in df.columns}
    total_row["Stock Symbol"] = "TOTAL"
    total_row["Gain/Loss (Value)"] = round(total_gain_loss_value, 2)
    total_row["Gain/Loss (%)"] = round(total_gain_loss_percent, 2)
    total_row["Total Value"] = round(total_portfolio_value, 2)
    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def autofit_columns(worksheet, df):
    for i, col in enumerate(df.columns, start=1):
        max_len = max([len(str(col))] + [len(str(v)) for v in df[col].tolist()])
        worksheet.column_dimensions[get_column_letter(i)].width = max_len + 2


def make_sheet_name(label, suffix):
    full = f"{label}{suffix}"
    if len(full) <= 31:
        return full
    keep = max(31 - len(suffix), 0)
    return f"{label[:keep]}{suffix}"


def split_by_recency(rows, purchase_dates):
    cutoff = datetime.now().date() - timedelta(days=RECENT_BUY_MONTHS * 30)
    recent_rows, older_rows = [], []
    for row in rows:
        bought_on = purchase_dates.get(row["Stock Symbol"])
        if bought_on and bought_on >= cutoff:
            recent_rows.append(row)
        else:
            older_rows.append(row)
    return recent_rows, older_rows


def build_excel_bytes(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            autofit_columns(writer.sheets[sheet_name], df)
    buf.seek(0)
    return buf.getvalue()


# --------------------------------------------------------------------------------------
# AI performance analysis & watchlist (optional — uses Claude via the Anthropic API)
# --------------------------------------------------------------------------------------
# IMPORTANT: this analysis is derived ONLY from the price/quantity numbers already in
# your sheet (current price, 1D/1W/1M/3M/6M/1Y % change, gain/loss). It has no access
# to company fundamentals, earnings, news, or analyst research — so it can describe
# *momentum patterns* in the data, but it cannot actually identify "multibaggers" in
# any meaningful sense (that requires business fundamentals, not just recent price
# action). The prompt below is written to keep the model's output honest about that
# limitation rather than presenting speculation as confident advice.
AI_DISCLAIMER = (
    "⚠️ **AI-generated, data-pattern summary — not financial advice.** "
    "This is generated only from the price/quantity data in your sheet (no company "
    "fundamentals, earnings, or news). Past price momentum does not predict future "
    "returns. Please consult a licensed financial advisor before making investment "
    "decisions."
)

AI_SYSTEM_PROMPT = """You are a portfolio data-analysis assistant. You will be given \
tables of stock holdings with derived price-change statistics (current price, \
purchase price, gain/loss %, and 1D/1W/1M/3M/6M/1Y % change). Produce a DETAILED \
Markdown report with exactly these sections, in this order:

## Portfolio Overview
Aggregate stats across all sheets combined: total number of holdings, how many are \
up vs down since purchase, and the 3 largest gainers and 3 largest losers by \
Gain/Loss (%) with their exact numbers.

## Performance Analysis
A thorough breakdown, not a brief summary:
- Discuss momentum across timeframes: which stocks show CONSISTENT direction across \
  1M, 3M, 6M, and 1Y (sustained trend) vs which show conflicting signals between \
  short-term (1D/1W) and long-term (6M/1Y) — call out at least 5-8 individual stocks \
  by name with their specific numbers as evidence, not just the portfolio in general.
- Note any stocks with unusually high volatility (large swings between timeframes).
- Compare the "New <6M" and "Existing" sheets if both are present: is one performing \
  differently from the other in aggregate?
- Every claim must cite the specific number(s) behind it. Do not summarize vaguely —
  name the stock and the number every time you make a claim about it.

## 6-Month Outlook
A speculative, pattern-based extrapolation — NOT a price prediction or guarantee. \
Structure this as:
- Stocks whose sustained multi-timeframe momentum (positive across 3M/6M/1Y together) \
  suggests the current trend has some persistence, purely as a pattern — name them \
  with numbers.
- Stocks whose recent (1D/1W/1M) direction conflicts with their longer-term (6M/1Y) \
  trend, where a reversal or continuation are both plausible — name them with numbers.
- Explicitly and clearly state you have no access to company fundamentals, earnings, \
  news, sector trends, or analyst coverage, and that momentum can reverse without \
  warning — this outlook describes historical price patterns only.

## Watchlist (Top 3 per Sheet)
For EACH sheet given, pick up to 3 stocks showing the strongest *combination* of \
sustained positive momentum across multiple timeframes (e.g. positive 1M, 3M, 6M, and \
1Y together). Use a "### <sheet name>" subheader per sheet. For each pick, write \
2-3 sentences: cite the specific numbers that justify it, AND note what would make \
this pick less compelling (e.g. a recent short-term reversal, high volatility, or \
thin evidence beyond price action). If a sheet has fewer than 3 eligible stocks, \
list what's available and say so plainly.

## Stocks Showing Concerning Momentum
For each sheet, name up to 3 stocks with the weakest combination of signals \
(sustained negative momentum across multiple timeframes, or a sharp recent decline) \
with their specific numbers. This is descriptive, not a sell instruction.

STRICT RULES:
- Do not invent or assume any company fundamentals, news, analyst ratings, sector \
trends, or events not present in the data given to you.
- Never use "guaranteed", "certain", "will definitely", or similar. Frame everything \
as a possibility, never a certainty.
- Never give direct buy/sell/hold instructions. Use language like "shows historical \
momentum worth monitoring" rather than "should buy".
- Do not use the word "multibagger" as a promise — if you reference the idea, frame \
it explicitly as a high-risk, speculative category based purely on price momentum.
- Be specific and numeric throughout — prefer "up 42% over 6M and 18% over 1M" over \
"performing well". Vague, generic statements are not acceptable in this report.
- Output valid Markdown with the headers above, nothing before or after them.
"""


def build_analysis_prompt(sheets):
    """Format each sheet (excluding the TOTAL row) into a compact Markdown table
    for the LLM — only the numbers that were actually computed, nothing else."""
    cols = [
        "Stock Symbol",
        "Current Price",
        "Purchase Price",
        "Gain/Loss (%)",
        "1D Change (%)",
        "1W Change (%)",
        "1M Change (%)",
        "3M Change (%)",
        "6M Change (%)",
        "1Y Change (%)",
    ]
    parts = []
    for sheet_name, df in sheets.items():
        data_rows = df[df["Stock Symbol"] != "TOTAL"]
        available_cols = [c for c in cols if c in data_rows.columns]
        try:
            table_text = data_rows[available_cols].to_markdown(index=False)
        except ImportError:
            # to_markdown needs the optional 'tabulate' package — fall back to CSV
            # text if it's not installed, which the model can still read fine.
            table_text = data_rows[available_cols].to_csv(index=False)
        parts.append(f"### Sheet: {sheet_name}\n\n{table_text}")
    return "\n\n".join(parts)


def run_anthropic_analysis(api_key, model, sheets, max_tokens=4096):
    """Call Claude with the portfolio data and return the Markdown report text."""
    client = anthropic.Anthropic(api_key=api_key)
    data_prompt = build_analysis_prompt(sheets)
    response = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        system=AI_SYSTEM_PROMPT,
        messages=[{"role": "user", "content": data_prompt}],
    )
    text_parts = [block.text for block in response.content if getattr(block, "type", "") == "text"]
    return "\n".join(text_parts)


def run_openai_analysis(api_key, model, sheets, max_tokens=4096):
    """Call OpenAI (basic chat.completions usage) with the portfolio data and
    return the Markdown report text."""
    client = openai.OpenAI(api_key=api_key)
    data_prompt = build_analysis_prompt(sheets)
    response = client.chat.completions.create(
        model=model,
        max_tokens=max_tokens,
        messages=[
            {"role": "system", "content": AI_SYSTEM_PROMPT},
            {"role": "user", "content": data_prompt},
        ],
    )
    return response.choices[0].message.content


def run_gemini_analysis(api_key, model, sheets, max_tokens=4096):
    """Call Gemini (via the google-genai SDK) with the portfolio data and
    return the Markdown report text."""
    client = genai.Client(api_key=api_key)
    data_prompt = build_analysis_prompt(sheets)
    response = client.models.generate_content(
        model=model,
        contents=data_prompt,
        config=genai_types.GenerateContentConfig(
            system_instruction=AI_SYSTEM_PROMPT,
            max_output_tokens=max_tokens,
        ),
    )
    return response.text


def run_ai_analysis(provider, api_key, model, sheets, max_tokens=4096):
    if provider == "OpenAI":
        return run_openai_analysis(api_key, model, sheets, max_tokens=max_tokens)
    if provider == "Gemini":
        return run_gemini_analysis(api_key, model, sheets, max_tokens=max_tokens)
    return run_anthropic_analysis(api_key, model, sheets, max_tokens=max_tokens)


# --------------------------------------------------------------------------------------
# Chat: interactive Q&A grounded in live web/news search, per provider
# --------------------------------------------------------------------------------------
# Each provider grounds responses in current web content differently — there is no
# shared "tools" param that works the same way across all three:
#   - Anthropic: add the web_search_20250305 tool to messages.create.
#   - OpenAI (Chat Completions): only works via dedicated search models
#     (gpt-4o-search-preview, gpt-5-search-api) + web_search_options={}. The regular
#     "tools" mechanism used by the Responses API does NOT work in Chat Completions.
#   - Gemini: add a Tool(google_search=GoogleSearch()) to the request config.
CHAT_DISCLAIMER = (
    "⚠️ **AI chat, grounded in web/news search — not financial advice.** Responses "
    "combine the price data in your sheet with whatever the model finds via live "
    "search. Search results can be incomplete, outdated, or wrong; the model can "
    "still misinterpret them. This is not a recommendation to buy, hold, or sell "
    "anything. Consult a licensed financial advisor before making investment decisions."
)

OPENAI_SEARCH_MODEL_FALLBACK = "gpt-4o-search-preview"


def build_chat_system_prompt(sheets):
    """System prompt for the chat: combines the holdings data with a mandate to
    search for relevant, current news and structure summaries per sheet."""
    sheet_names = list(sheets.keys())
    data_prompt = build_analysis_prompt(sheets)
    return f"""You are a portfolio research assistant with live web search access. \
The person's holdings, with derived price-change statistics, are below. Use web \
search to find CURRENT news, events, and sector/market context relevant to specific \
holdings when it would improve your answer — don't rely on price data alone once \
search is available to you.

HOLDINGS DATA:
{data_prompt}

WHEN ASKED FOR A SUMMARY, OVERVIEW, OR ANALYSIS ACROSS THE WHOLE PORTFOLIO:
Structure your answer with one subsection per sheet, using these exact headers: \
{", ".join(f'"### {name}"' for name in sheet_names)}. Within each, discuss \
performance (grounded in the numbers above) together with any relevant recent news \
you found via search, and note which stocks show the strongest speculative \
"multibagger-style" momentum (sustained gains across multiple timeframes) versus \
which show concerning or conflicting signals.

STRICT RULES:
- When you use information from a search, mention the source (publication/site name \
and, if useful, approximate date) so the person can verify it themselves.
- Never state a future price, target, or return as if it were a fact. Frame all \
forward-looking statements as possibilities based on current momentum and/or news, \
never as certainties.
- Never give direct buy/sell/hold instructions.
- Do not use "multibagger" as a promise — treat it as a high-risk, speculative \
label based on price momentum, and say so explicitly when you use the term.
- If search doesn't turn up anything relevant for a stock, say so plainly rather \
than inventing news to fill the gap.
- For casual/specific questions (e.g. about one stock), you don't need the full \
per-sheet structure — just answer directly, still grounded and sourced.
"""


def chat_anthropic(api_key, model, system_prompt, history, max_tokens=2048):
    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        system=system_prompt,
        messages=[{"role": m["role"], "content": m["content"]} for m in history],
        tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": 5}],
    )
    return "\n".join(block.text for block in response.content if getattr(block, "type", "") == "text")


def chat_openai(api_key, model, system_prompt, history, max_tokens=2048):
    client = openai.OpenAI(api_key=api_key)
    effective_model = model if "search" in model.lower() else OPENAI_SEARCH_MODEL_FALLBACK
    messages = [{"role": "system", "content": system_prompt}]
    messages += [{"role": m["role"], "content": m["content"]} for m in history]
    response = client.chat.completions.create(
        model=effective_model,
        web_search_options={},
        max_tokens=max_tokens,
        messages=messages,
    )
    return response.choices[0].message.content, effective_model


def chat_gemini(api_key, model, system_prompt, history, max_tokens=2048):
    client = genai.Client(api_key=api_key)
    contents = [
        {"role": ("model" if m["role"] == "assistant" else "user"), "parts": [{"text": m["content"]}]}
        for m in history
    ]
    grounding_tool = genai_types.Tool(google_search=genai_types.GoogleSearch())
    response = client.models.generate_content(
        model=model,
        contents=contents,
        config=genai_types.GenerateContentConfig(
            system_instruction=system_prompt,
            tools=[grounding_tool],
            max_output_tokens=max_tokens,
        ),
    )
    return response.text


def run_chat_turn(provider, api_key, model, system_prompt, history, max_tokens=2048):
    """Returns (reply_text, model_actually_used) — model can differ from the
    configured one for OpenAI, since search needs a dedicated model."""
    if provider == "OpenAI":
        return chat_openai(api_key, model, system_prompt, history, max_tokens=max_tokens)
    if provider == "Gemini":
        return chat_gemini(api_key, model, system_prompt, history, max_tokens=max_tokens), model
    return chat_anthropic(api_key, model, system_prompt, history, max_tokens=max_tokens), model


# --------------------------------------------------------------------------------------
# Catalyst & Momentum Scanner — a separate, search-heavy research task. Reuses the
# same per-provider grounded-call plumbing as the chat feature (run_chat_turn), just
# as a single-turn call with a much larger token budget, since this prompt asks for
# many structured tables and a 5-10 stock deep dive.
# --------------------------------------------------------------------------------------
SCANNER_DISCLAIMER = (
    "⚠️ **Speculative short-term research tool — not financial advice.** This scans "
    "for market catalysts using live web/news search and historical pattern framing "
    "only. It explicitly does not give price targets or guarantees (the prompt "
    "itself enforces this), but search results can still be incomplete, stale, or "
    "wrong, and the model can misjudge catalyst quality. Short-term, catalyst-driven "
    "trading carries high risk. Consult a licensed financial advisor before acting "
    "on anything here."
)

SCANNER_SYSTEM_PROMPT = """Act as a short-term Indian equity market research analyst and event-driven market researcher.

The objective is to identify fresh, verifiable catalysts in Indian stocks and ETFs that may lead to significant price movement over the next 1-5 trading sessions, while also studying historical catalyst behaviour to determine whether similar events typically produce:

- 1-day spikes
- 3-5 day momentum
- 6-20 day momentum
- Multi-week repricing

This is not a long-term investment analysis.

The primary objective is to identify situations where:

Fresh catalyst + limited initial price reaction + abnormal/increasing volume + technical confirmation + potential catalyst persistence

may indicate that the market is still digesting new information.

Do not assume that every positive announcement will result in a sustained price increase.

---

PHASE 1 - MARKET-WIDE CATALYST DISCOVERY

Scan the Indian Market. Search the latest available information from:
- NSE corporate announcements
- BSE corporate announcements
- Company investor-relations websites
- SEBI announcements/orders
- Government of India ministries
- RBI
- Other Indian regulators
- Government/PSU tender announcements
- Stock-exchange filings
- Bulk/block deal disclosures
- Promoter buying/selling disclosures
- Institutional activity
- Sector-specific announcements
- Reputable financial news
- Credible business media

Prioritize information published within the last 24-72 hours. For rapidly moving markets, prioritize information from the last trading session and current trading session.

Do not rely on social-media rumours as verified information. If social media or unofficial sources surface a potentially important catalyst, label it: Unverified / Requires Confirmation

---

PHASE 2 - IDENTIFY FRESH CATALYSTS

Look specifically for:

Corporate Catalysts: large orders/contracts, government/PSU orders, new customers, strategic partnerships, acquisitions, mergers, joint ventures, fundraising, preferential issues, qualified institutional placements, promoter buying, institutional buying, bulk/block deals, regulatory approvals, new manufacturing capacity, capex announcements, new product launches, export opportunities, international expansion, strategic investments, debt reduction, debt restructuring, major restructuring, management changes, earnings surprises, guidance upgrades, margin expansion, major contract wins.

Regulatory / Government Catalysts: government policy changes, PLI/incentive changes, tariff changes, import/export restrictions, subsidies, regulatory approvals, new government schemes, sector-specific policy changes, RBI decisions, SEBI decisions, defence procurement, infrastructure spending, energy policy, renewable-energy policy.

Sector Catalysts: identify developments that could affect multiple companies across renewable energy, BESS, defence, railways, infrastructure, semiconductor, AI/data centres, power, banking, pharmaceuticals, chemicals, manufacturing, mining/metals, real estate, consumer, EV, telecom. Also identify whether the catalyst is company-specific or sector-wide.

---

PHASE 3 - NEWS -> PRICE GAP

For every candidate, determine: when was the catalyst first publicly announced; what was the exact announcement; was the information genuinely new; was it already anticipated by the market; had the stock already moved before the announcement; what was the stock price immediately before the catalyst; what was the first trading reaction; how much has the stock moved since; is volume significantly above normal; is the price continuing to move; is the stock consolidating; has another confirming catalyst appeared; does the current price already appear to have priced in the news.

Pay particular attention to: fresh catalyst + limited initial price reaction + increasing volume + continued follow-through. This is more important than simply finding stocks that are already up 10-20%.

---

PHASE 4 - PRICE & VOLUME ANALYSIS

For every candidate collect: current price, previous close, today's % change, intraday high, intraday low, 1-day return, 3-day return, 5-day return, 10-day return, 20-day return, 1-month return, 3-month return, 52-week return, 52-week high, 52-week low, today's volume, average 20-day volume, average 30-day volume, volume ratio, market capitalisation, recent support, recent resistance, breakout level, recent consolidation range.

Calculate Volume Ratio = Today's Volume / Average 20-30 Day Volume. Classify volume as Normal, Elevated, High, or Abnormally high. Determine whether price + volume are confirming each other.

---

PHASE 5 - CATALYST QUALITY

For each catalyst assess Catalyst Strength as Very High, High, Moderate, Low, or Unclear, based on financial materiality, revenue/earnings potential, size of contract, customer quality, recurring vs one-time revenue, strategic importance, regulatory significance, probability of execution, time required to realize financial impact.

Do not confuse a large headline number with actual near-term earnings impact. For example, a Rs 1,000 crore order should not automatically be treated as equivalent to Rs 1,000 crore of immediate revenue. Consider order duration, execution period, margins, working capital, financing requirements, cancellation risk, existing order book, revenue recognition.

---

PHASE 6 - AVOID CHASING

Explicitly flag candidates where: stock already rose substantially before the catalyst; catalyst is old; news is already widely priced in; stock is extremely extended; repeated upper circuits have occurred; price moved without fundamental confirmation; volume is abnormal but catalyst is unclear; announcement is only an MoU; transaction is not yet completed; execution risk is high; financing risk is high; valuation is extremely stretched; liquidity is poor; move appears primarily speculative.

Classify these as: Already Extended / High Chase Risk. Do not assume a strong catalyst automatically makes the current entry attractive.

---

PHASE 7 - HISTORICAL CATALYST ANALYSIS

For each important candidate, investigate similar historical catalyst events where data is available. Track the stock from the first verified public release of the catalyst through at least 20 trading sessions. Do not stop at the initial reaction. Determine: initial reaction, day +1, day +3, day +5, day +10, day +20, peak gain, peak date, days to peak, maximum drawdown, momentum duration, catalyst reinforcement, momentum-ending event.

---

PHASE 8 - CATALYST PERSISTENCE METRICS

For every historical catalyst event calculate: Initial Reaction (return on first trading day after catalyst), Day +1/+3/+5/+10/+20 (cumulative return after N trading sessions), Peak Gain, Peak Date, Days to Peak, Maximum Drawdown (largest decline from post-news high within 20-day window), Momentum Duration (sessions stock remained materially above pre-news price), Catalyst Reinforcement (whether additional positive catalysts appeared), Momentum-Ending Event (first major event associated with reversal/loss of momentum).

Use the pre-news closing price as the primary baseline wherever possible. Clearly distinguish intraday gain, closing-price gain, and cumulative gain. Do not mix these measurements.

---

PHASE 9 - DEFINE MOMENTUM DURATION

Classify each event: 1-Day Event (initial move >=10% or significant abnormal reaction, reverses within 1-2 sessions, no meaningful follow-up catalyst); Short-Term Event (momentum persists 3-5 trading sessions); Medium-Term Event (6-20 trading sessions); Extended Event (beyond 20 trading sessions).

Do not classify a stock as a long-duration catalyst merely because it is higher 20 days later. The key question is: did the original catalyst create persistent repricing, or did subsequent events create the additional gains?

---

PHASE 10 - SEPARATE THE SOURCES OF MOMENTUM

For every major price move, distinguish between: Initial Catalyst, Follow-up Catalyst, Earnings Confirmation, Technical Breakout, Sector Movement, Market Movement (Nifty/Sensex/global), Institutional Activity (promoter/institutional accumulation or block/bulk deals), Speculative/Momentum Buying (continuation without a clearly identifiable new fundamental catalyst).

This separation is mandatory. Do not attribute all subsequent gains to the original announcement.

---

PHASE 11 - CATALYST PERSISTENCE STATISTICS

Across the historical catalyst sample, calculate: % of events still positive after 1/3/5/10/20 days; average and median return at Day +1/+3/+5/+10/+20; median maximum gain; median days to peak; median maximum drawdown.

When sample size is small, explicitly state: "Sample size too small for a reliable statistical conclusion." Never manufacture statistical confidence from a small sample.

---

PHASE 12 - IDENTIFY EARLY-WARNING PATTERNS

Determine whether these characteristics are associated with longer-lasting price moves: strong earnings surprise, large order, government policy, M&A/acquisition, promoter buying, institutional buying, guidance upgrade, regulatory approval, sector-wide catalyst, new 52-week high, abnormal volume, low free float, small/micro-cap status, multiple catalysts within 5 trading sessions, strong price/volume confirmation, consolidation after initial reaction, follow-up company announcements.

Compare these against 1-Day Spike Patterns vs Multi-Day Momentum Patterns vs Multi-Week Repricing Patterns. Do not assume correlation implies causation. Where sample size permits, identify which characteristics appear repeatedly.

---

PHASE 13 - SETUP CLASSIFICATION

Do not rank stocks simply as "best" or "worst." Instead classify each candidate into one of: A. Fresh Catalyst; B. Catalyst + Confirmation; C. Pullback + Catalyst; D. Breakout Watch; E. Already Extended; F. Unexplained Spike; G. Weak Catalyst Reaction; H. Persistent Repricing Candidate.

---

PHASE 14 - MAIN MARKET SCAN TABLE

Produce a Markdown table with columns: Stock | Catalyst | Announcement Date | Current Price | 1D | 5D | Volume vs Avg | Catalyst Strength | Price Confirmation | Setup Type | Key Level | Main Risk. Only include stocks with a credible and verifiable catalyst.

---

PHASE 15 - CATALYST PERSISTENCE TABLE

For historical events, produce a Markdown table with columns: Stock | Catalyst | Initial Move | +1D | +3D | +5D | +10D | +20D | Peak Gain | Peak Date | Days to Peak | Max Drawdown | Momentum Duration | Additional Catalyst? | Momentum-Ending Event. Clearly identify whether returns are cumulative from pre-news close rather than daily returns.

---

PHASE 16 - DEEP DIVE

For the 5-10 most relevant current candidates provide: Catalyst (what happened and why it matters); Market Reaction (pre-news price, initial reaction, current price, 1D/3D/5D performance, accelerating or fading); Volume (whether it confirms the move); Catalyst Persistence (1 day / 3-5 days / 1-2 weeks / longer, using historical evidence where available); Price Structure (recent support/resistance, breakout level, recent high/low, consolidation range); Valuation where relevant (P/E, P/B, EV/EBITDA, market cap, recent growth, earnings trend — do not use valuation alone to predict short-term price movement); Risks (valuation, execution, regulatory, financing, profit-taking, low liquidity, promoter activity, news already priced in, sector weakness, market-wide weakness).

---

PHASE 17 - CURRENT 1-WEEK WATCHLIST

Create a Markdown table: Stock | Setup Type | Why Interesting | Confirmation Needed | Invalidation Signal | Catalyst Horizon | Main Risk. The purpose is to identify what needs to happen next, rather than simply saying what stock looks attractive. Do not provide guaranteed price targets. Do not state that any stock "will rise."

---

PHASE 18 - PREVIOUSLY IDENTIFIED CANDIDATES

If relevant, re-check previously identified stocks, including: Cupid, GK Energy, Pace Digitek, Sterling & Wilson Renewable Energy, Embassy Developments, Waaree Energies, Raymond, Hindustan Copper, GRSE, Protean eGov, Taneja Aerospace, RHI Magnesita.

Do not assume that a stock remains a candidate simply because it appeared in a previous scan. Re-check latest catalyst, latest price, volume, new announcements, catalyst age, technical structure, whether the original catalyst is still relevant. Remove stocks where the catalyst has become stale or the price has already fully reacted.

---

PHASE 19 - CRITICAL ANALYTICAL RULES

Follow strictly: use fresh information (last 24-72 hours where possible); always provide publication dates/times where available; separate verified facts, analysis and speculation; cite every important catalyst; prefer official company/NSE/BSE/SEBI sources for corporate events; cross-check major catalysts with at least one credible independent source; never invent prices, volumes, announcements or historical returns; if historical data is unavailable, explicitly state that; do not treat a large order headline as equivalent to immediate revenue; do not treat a stock being up 10-20% as evidence that it will continue rising; do not assume historical catalyst behaviour will repeat; do not attribute subsequent gains to the original catalyst without evidence; separate company-specific, sector-wide and market-wide effects; flag when the catalyst has already been priced in; flag abnormal price moves where no verified catalyst can be identified; distinguish an MoU from a completed transaction; distinguish an order win from actual revenue/profit realization; consider liquidity and free float when interpreting abnormal price movement; when sample size is small, explicitly state the limitation; do not manufacture statistical significance; do not make guaranteed-return claims; do not present historical performance as a prediction of future performance.

---

PHASE 20 - FINAL OUTPUT

Structure the final answer with these Markdown headers, in order:

## 1. Executive Summary
Summarize the most important fresh catalysts discovered today. Highlight new catalysts, strong price confirmation, weak price reaction, already-extended names, unexplained spikes, sector-level themes.

## 2. Today's Catalyst Scanner
The main market-wide table from Phase 14.

## 3. Historical Catalyst Persistence
The +1D/+3D/+5D/+10D/+20D analysis from Phase 15.

## 4. Early-Warning Patterns
Which catalyst characteristics historically corresponded with 1-day spikes, multi-day momentum, multi-week repricing.

## 5. Detailed Candidate Analysis
The deep-dive into the 5-10 most relevant stocks from Phase 16.

## 6. One-Week Watchlist
Separated into: Fresh Catalyst, Catalyst + Confirmation, Pullback + Catalyst, Breakout Watch, Persistent Repricing Candidate, Already Extended, Unexplained Spike, Weak Catalyst Reaction.

## 7. Key Takeaways
Answer: which types of fresh catalysts are currently appearing in the Indian market, and what evidence exists that these catalysts are producing temporary spikes versus persistent multi-day repricing?

The goal is not to predict the market. The goal is to systematically identify: fresh information -> initial price reaction -> volume confirmation -> catalyst reinforcement -> persistence -> momentum exhaustion, and use historical evidence to understand which patterns tend to persist beyond the first trading session.
"""


def run_scanner_analysis(provider, api_key, model, max_tokens, exclude_symbols=None):
    """Single-turn, search-grounded run of the Catalyst & Momentum Scanner prompt.
    Reuses the same per-provider plumbing as the chat feature."""
    user_prompt = "Run the full scan now, following every phase exactly as specified."
    if exclude_symbols:
        user_prompt += (
            "\n\nThe user already holds the following stocks. This scan is for "
            "finding NEW opportunities outside the current portfolio, so EXCLUDE "
            "all of these entirely from every table, the deep dive, the watchlist, "
            "and the Phase 18 previously-identified-candidates check — even if one "
            "of them currently has a fresh catalyst: " + ", ".join(exclude_symbols)
        )
    history = [{"role": "user", "content": user_prompt}]
    reply, model_used = run_chat_turn(
        provider, api_key, model, SCANNER_SYSTEM_PROMPT, history, max_tokens=max_tokens
    )
    return reply, model_used


# --------------------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------------------
st.set_page_config(page_title="Portfolio Tracker (Excel + Yahoo Finance)", layout="wide")
st.title("📈 Portfolio Tracker")
st.caption(
    "Holdings come from your uploaded Excel file. Prices and 1D/1W/1M/3M/6M/1Y "
    "change come from Yahoo Finance (yfinance) — no API key or subscription needed."
)

with st.sidebar:
    st.header("1. Holdings file")
    st.caption(
        "Upload your broker/Groww holdings export. Columns are auto-detected — "
        "see the app's module docstring for exact names recognized, or just add "
        "a 'ticker' column with the Yahoo symbol to be safe."
    )
    holdings_file = st.file_uploader("Holdings Excel (.xlsx)", type=["xlsx", "xls"])

    sheet_name_choice = None
    if holdings_file is not None:
        try:
            xls = pd.ExcelFile(holdings_file)
            sheet_name_choice = st.selectbox("Sheet", xls.sheet_names)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Could not read the Excel file: {exc}")

    st.divider()
    st.header("2. Purchase dates (optional)")
    st.caption(
        f"Upload a CSV to split stocks bought in the last {RECENT_BUY_MONTHS} months "
        "into their own sheet. Columns: trading_symbol, purchase_date (YYYY-MM-DD)."
    )
    uploaded_dates_file = st.file_uploader("purchase_dates.csv", type=["csv"])

    st.divider()
    st.header("3. Output")
    st.caption("Defaults to the folder this script lives in.")
    output_path_str = st.text_input("Save .xlsx to", value=str(DEFAULT_OUTPUT_PATH))

    account_label = st.text_input("Sheet label", value="Portfolio")

    st.divider()
    st.header("4. AI analysis (optional)")
    st.caption(
        "Summarizes momentum and surfaces a speculative watchlist. Your holdings "
        "numbers (not your identity) are sent to the chosen provider's API for this."
    )
    ai_depth = st.select_slider(
        "Analysis depth",
        options=["Standard", "Detailed", "Very Detailed"],
        value="Detailed",
        help="Controls the output length budget (max_tokens) sent to the model. "
        "More detail = a longer, more thorough report but higher API cost.",
    )
    AI_MAX_TOKENS = {"Standard": 2000, "Detailed": 4096, "Very Detailed": 8192}[ai_depth]
    ai_provider = st.radio("Provider", options=["OpenAI", "Anthropic", "Gemini"], horizontal=True)

    if ai_provider == "OpenAI":
        ai_api_key = st.text_input(
            "OpenAI API key",
            value=os.environ.get("OPENAI_API_KEY", ""),
            type="password",
        )
        ai_model = st.text_input(
            "Model",
            value="gpt-4o-mini",
            help="Basic chat.completions usage — any OpenAI chat model works. "
            "gpt-4o-mini is a cheap, fast default for this kind of summarization task.",
        )
    elif ai_provider == "Gemini":
        ai_api_key = st.text_input(
            "Gemini API key",
            value=os.environ.get("GEMINI_API_KEY", os.environ.get("GOOGLE_API_KEY", "")),
            type="password",
        )
        ai_model = st.text_input(
            "Model",
            value="gemini-2.5-flash",
            help="Any Gemini model works — this is a free-text field since model "
            "names change over time. 'flash' variants are the cheap/fast tier.",
        )
    else:
        ai_api_key = st.text_input(
            "Anthropic API key",
            value=os.environ.get("ANTHROPIC_API_KEY", ""),
            type="password",
        )
        ai_model = st.selectbox(
            "Model",
            options=["claude-sonnet-5", "claude-opus-5", "claude-haiku-4-5-20251001"],
            index=0,
        )


tab1, tab2 = st.tabs(["📈 Portfolio Tracker", "📡 Catalyst & Momentum Scanner"])

with tab1:
    fetch_clicked = st.button("🔄 Fetch Prices & Build Report", type="primary")

    if fetch_clicked:
        if holdings_file is None:
            st.error("Upload a holdings Excel file first.")
            st.stop()

        try:
            holdings_df = load_holdings_table(holdings_file, sheet_name_choice)
        except Exception as exc:  # noqa: BLE001
            st.error(f"Could not read the Excel file: {exc}")
            st.stop()

        ticker_col, symbol_col, qty_col, price_col, isin_col = detect_columns(holdings_df)

        missing = []
        if symbol_col is None:
            missing.append("symbol/name")
        if qty_col is None:
            missing.append("quantity")
        if price_col is None:
            missing.append("purchase price")
        if missing:
            st.error(
                f"Could not auto-detect column(s) for: {', '.join(missing)}. "
                f"Columns found in your file: {list(holdings_df.columns)}. "
                "Rename a column to something recognizable (see the app's docstring) and re-upload."
            )
            st.stop()

        st.caption(
            f"Detected columns — Symbol: `{symbol_col}` · Quantity: `{qty_col}` · "
            f"Purchase Price: `{price_col}`"
            + (f" · Ticker override: `{ticker_col}`" if ticker_col else "")
            + (f" · ISIN: `{isin_col}`" if isin_col else "")
        )

        log = []
        progress = st.progress(0.0, text="Fetching prices...")
        rows = build_holding_rows(
            holdings_df, symbol_col, qty_col, price_col, ticker_col, isin_col, progress=progress, log=log
        )
        progress.empty()

        if not rows:
            st.warning("No rows could be priced. See the log below.")
            for line in log:
                st.text(line)
            st.stop()

        purchase_dates = load_purchase_dates_from_upload(uploaded_dates_file)
        recent_rows, older_rows = split_by_recency(rows, purchase_dates)

        sheets = {}
        if recent_rows:
            sheets[make_sheet_name(account_label, f" (New <{RECENT_BUY_MONTHS}M)")] = make_df_with_total(recent_rows)
        if older_rows:
            sheets[make_sheet_name(account_label, " (Existing)")] = make_df_with_total(older_rows)

        st.session_state["sheets"] = sheets
        st.session_state["all_symbols"] = sorted({r["Stock Symbol"] for r in rows})
        st.session_state["unknown_symbols"] = sorted(
            {r["Stock Symbol"] for r in rows} - set(purchase_dates.keys())
        )
        st.session_state["log"] = log

    # --------------------------------------------------------------------------------------
    # Results
    # --------------------------------------------------------------------------------------
    if "sheets" in st.session_state and st.session_state["sheets"]:
        sheets = st.session_state["sheets"]
        st.success(f"Priced {len(st.session_state['all_symbols'])} holdings.")

        if st.session_state["unknown_symbols"]:
            st.info(
                f"{len(st.session_state['unknown_symbols'])} symbol(s) have no purchase date "
                f"on file and were placed in the 'Existing' sheet: "
                f"{', '.join(st.session_state['unknown_symbols'])}"
            )
            st.download_button(
                "⬇️ Download purchase_dates.csv template",
                data=make_template_csv_bytes(st.session_state["all_symbols"]),
                file_name="purchase_dates.csv",
                mime="text/csv",
            )

        tabs = st.tabs(list(sheets.keys()))
        for tab, (sheet_name, df) in zip(tabs, sheets.items()):
            with tab:
                st.dataframe(df, use_container_width=True)

        excel_bytes = build_excel_bytes(sheets)
        st.download_button(
            "⬇️ Download stock_portfolio.xlsx",
            data=excel_bytes,
            file_name="stock_portfolio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if st.button("💾 Also save to the output path above"):
            try:
                output_path = Path(output_path_str)
                output_path.parent.mkdir(parents=True, exist_ok=True)
                output_path.write_bytes(excel_bytes)
                st.success(f"Saved to {output_path}")
            except OSError as exc:
                st.error(f"Could not save to {output_path_str}: {exc}")

        st.divider()
        st.header("🤖 AI Performance Analysis & 6-Month Outlook")
        st.warning(AI_DISCLAIMER)

        if st.button("Run Performance Analysis & Watchlist"):
            provider_availability = {
                "OpenAI": (OPENAI_AVAILABLE, "openai"),
                "Anthropic": (ANTHROPIC_AVAILABLE, "anthropic"),
                "Gemini": (GEMINI_AVAILABLE, "google-genai"),
            }
            provider_available, pip_name = provider_availability[ai_provider]
            if not provider_available:
                st.error(f"Missing dependency. Install with:\n\n    pip install {pip_name}")
            elif not ai_api_key:
                st.error(f"Enter your {ai_provider} API key in the sidebar first.")
            else:
                with st.spinner(f"Analyzing portfolio momentum ({ai_depth.lower()})..."):
                    try:
                        st.session_state["ai_report"] = run_ai_analysis(
                            ai_provider, ai_api_key, ai_model, sheets, max_tokens=AI_MAX_TOKENS
                        )
                    except Exception as exc:  # noqa: BLE001 - surface any API error to the user
                        st.error(f"AI analysis failed: {exc}")

        if st.session_state.get("ai_report"):
            st.markdown(st.session_state["ai_report"])
            st.download_button(
                "⬇️ Download this analysis (.md)",
                data=st.session_state["ai_report"].encode("utf-8"),
                file_name="portfolio_ai_analysis.md",
                mime="text/markdown",
            )

        st.divider()
        st.header("💬 Chat: Performance & Multibagger Potential (grounded in web/news search)")
        st.warning(CHAT_DISCLAIMER)

        if "chat_messages" not in st.session_state:
            st.session_state["chat_messages"] = []

        col_a, col_b = st.columns([3, 1])
        with col_b:
            if st.button("🗑️ Clear chat"):
                st.session_state["chat_messages"] = []
                st.rerun()
        with col_a:
            if st.button("📋 Run full per-section summary now"):
                st.session_state["chat_messages"].append(
                    {
                        "role": "user",
                        "content": (
                            "Give me a full performance and multibagger-potential summary, "
                            "structured as one section per sheet, using current news where relevant."
                        ),
                    }
                )
                st.session_state["_chat_pending"] = True

        for msg in st.session_state["chat_messages"]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        user_input = st.chat_input("Ask about performance, momentum, or multibagger potential...")
        if user_input:
            st.session_state["chat_messages"].append({"role": "user", "content": user_input})
            st.session_state["_chat_pending"] = True
            st.rerun()

        if st.session_state.get("_chat_pending"):
            provider_availability = {
                "OpenAI": (OPENAI_AVAILABLE, "openai"),
                "Anthropic": (ANTHROPIC_AVAILABLE, "anthropic"),
                "Gemini": (GEMINI_AVAILABLE, "google-genai"),
            }
            provider_available, pip_name = provider_availability[ai_provider]
            if not provider_available:
                st.error(f"Missing dependency. Install with:\n\n    pip install {pip_name}")
                st.session_state["_chat_pending"] = False
            elif not ai_api_key:
                st.error(f"Enter your {ai_provider} API key in the sidebar first.")
                st.session_state["_chat_pending"] = False
            else:
                with st.spinner("Searching and responding..."):
                    try:
                        system_prompt = build_chat_system_prompt(sheets)
                        reply, model_used = run_chat_turn(
                            ai_provider, ai_api_key, ai_model, system_prompt, st.session_state["chat_messages"]
                        )
                        if model_used != ai_model:
                            reply = f"*(used {model_used} for web search — your configured model doesn't support it)*\n\n{reply}"
                        st.session_state["chat_messages"].append({"role": "assistant", "content": reply})
                    except Exception as exc:  # noqa: BLE001 - surface any API error to the user
                        st.session_state["chat_messages"].append(
                            {"role": "assistant", "content": f"⚠️ Request failed: {exc}"}
                        )
                    st.session_state["_chat_pending"] = False
                    st.rerun()

        if st.session_state.get("log"):
            with st.expander("⚠️ Warnings from this fetch"):
                for line in st.session_state["log"]:
                    st.text(line)

with tab2:
    st.header("📡 Indian Stock Short-Term Catalyst & Momentum Scanner")
    st.warning(SCANNER_DISCLAIMER)
    st.caption(
        "Runs a large, multi-phase, search-heavy research prompt: scans NSE/BSE/SEBI "
        "announcements, news, and bulk/block deals from the last 24-72 hours for fresh "
        "catalysts, then cross-checks historical catalyst persistence. Uses the "
        "provider/model/API key configured in the sidebar, with web search forced on "
        "regardless of the toggle used elsewhere, since this prompt is unusable without it."
    )

    holdings_to_exclude = sorted(st.session_state.get("all_symbols", [])) if st.session_state.get("sheets") else []
    if holdings_to_exclude:
        st.caption(
            f"🚫 Excluding {len(holdings_to_exclude)} stock(s) already in your portfolio "
            "(fetched in the first tab) — this scan is for finding new opportunities, "
            "not re-covering what you already hold."
        )

    scanner_depth = st.select_slider(
        "Scanner depth",
        options=["Focused", "Standard", "Comprehensive"],
        value="Standard",
        help="This prompt asks for many tables plus a 5-10 stock deep dive — it "
        "needs a much larger token budget than the portfolio analysis. Comprehensive "
        "costs significantly more per run.",
    )
    SCANNER_MAX_TOKENS = {"Focused": 4096, "Standard": 8192, "Comprehensive": 16384}[scanner_depth]

    run_scanner_clicked = st.button("📡 Run Catalyst & Momentum Scan", type="primary")

    if run_scanner_clicked:
        provider_availability = {
            "OpenAI": (OPENAI_AVAILABLE, "openai"),
            "Anthropic": (ANTHROPIC_AVAILABLE, "anthropic"),
            "Gemini": (GEMINI_AVAILABLE, "google-genai"),
        }
        provider_available, pip_name = provider_availability[ai_provider]
        if not provider_available:
            st.error(f"Missing dependency. Install with:\n\n    pip install {pip_name}")
        elif not ai_api_key:
            st.error(f"Enter your {ai_provider} API key in the sidebar first.")
        else:
            with st.spinner(f"Scanning ({scanner_depth.lower()}) — this involves several live searches..."):
                try:
                    report, model_used = run_scanner_analysis(
                        ai_provider, ai_api_key, ai_model, SCANNER_MAX_TOKENS, holdings_to_exclude
                    )
                    if model_used != ai_model:
                        report = (
                            f"*(used {model_used} for web search — your configured "
                            f"model doesn't support it)*\n\n{report}"
                        )
                    st.session_state["scanner_report"] = report
                except Exception as exc:  # noqa: BLE001 - surface any API error to the user
                    st.error(f"Scan failed: {exc}")

    if st.session_state.get("scanner_report"):
        st.markdown(st.session_state["scanner_report"])
        st.download_button(
            "⬇️ Download this scan (.md)",
            data=st.session_state["scanner_report"].encode("utf-8"),
            file_name="catalyst_momentum_scan.md",
            mime="text/markdown",
        )
