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
import json
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


def get_secret(*names, default=""):
    """Look up a credential by trying, in order: Streamlit Community Cloud's
    st.secrets (set via the app's 'Advanced settings -> Secrets' dialog — this
    is NOT the same as an OS env var), then os.environ (for local runs, Docker,
    or Cloud Run, where env vars are the natural mechanism). Accepts multiple
    candidate names (e.g. GEMINI_API_KEY / GOOGLE_API_KEY) tried in order."""
    for name in names:
        try:
            if name in st.secrets:
                return st.secrets[name]
        except Exception:  # noqa: BLE001 - st.secrets raises if no secrets.toml exists at all
            pass
        if os.environ.get(name):
            return os.environ[name]
    return default


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
    "⚠️ **AI-generated summary, grounded in live web search — not financial advice.** "
    "This combines the price/quantity data in your sheet with news, earnings, and "
    "fundamentals the model finds via search for each stock. Search coverage is "
    "uneven — smaller or thinly-covered stocks may turn up little or nothing, which "
    "is reported honestly rather than filled in. Search results can also be "
    "incomplete, outdated, or wrong, and the model can misread them. Nothing here is "
    "a price prediction, guarantee, or instruction to buy/hold/sell. Please consult a "
    "licensed financial advisor before making investment decisions."
)

AI_SYSTEM_PROMPT = """You are a portfolio research assistant with live web search \
access. You're given stock symbols with price stats (current price, purchase price, \
gain/loss %, 1D/1W/1M/3M/6M/1Y % change) for context. For EACH stock, search the web \
for current news, recent earnings, and fundamentals (P/E, growth, guidance) — don't \
rely on price data alone.

Respond with ONLY a JSON array — no code fences, no text before or after it. One \
object per stock, covering every stock given (no omissions, no extra entries), with \
exactly these keys:

[{"stock": "<exact symbol given>", "momentum_pattern": "<few words, from the timeframe data only, e.g. 'Sustained uptrend', 'Mixed signals', 'High volatility', 'Flat / no clear trend'>", "key_news_fundamentals": "<1-2 sentences on what search found — earnings, announcements, fundamentals — each with source and approx. date; say 'No notable recent news found' if search finds nothing, rather than inventing something>", "outlook_note": "<ONE sentence combining the price pattern and the news/fundamentals, citing a specific number or source; speculative only — never a prediction, guarantee, or buy/sell instruction>"}]

RULES: state nothing as fact unless it's from a search result or the data given — \
never invent news, figures, or fundamentals. Never say "guaranteed", "certain", or \
"will definitely". Never give buy/sell/hold instructions — say "worth monitoring", \
not "should buy". Treat "multibagger" only as a speculative, high-risk label, never \
a promise. No vague claims ("performing well") — cite something specific every time. \
Output ONLY the JSON array — nothing else, no formatting around it.
"""


def compute_portfolio_overview_df(sheets):
    """Portfolio Overview, computed directly from real data — no LLM involved, so
    zero hallucination risk on these numbers. One row per sheet, no total row."""
    rows = []
    for sheet_name, df in sheets.items():
        data_rows = df[df["Stock Symbol"] != "TOTAL"]
        if data_rows.empty:
            continue
        gl = data_rows["Gain/Loss (%)"]
        best = data_rows.loc[gl.idxmax()]
        worst = data_rows.loc[gl.idxmin()]
        rows.append(
            {
                "Sheet": sheet_name,
                "Holdings": len(data_rows),
                "Up": int((gl >= 0).sum()),
                "Down": int((gl < 0).sum()),
                "Best Performer": f"{best['Stock Symbol']} ({best['Gain/Loss (%)']:+.2f}%)",
                "Worst Performer": f"{worst['Stock Symbol']} ({worst['Gain/Loss (%)']:+.2f}%)",
            }
        )
    return pd.DataFrame(rows)


def compute_stock_numeric_df(sheets):
    """All stocks across all sheets (excluding TOTAL rows), flattened into one
    DataFrame with a Status (Up/Down) column — the authoritative numeric source
    that the LLM's qualitative fields get merged into, rather than the other
    way around."""
    frames = [df[df["Stock Symbol"] != "TOTAL"].copy() for df in sheets.values()]
    combined = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if not combined.empty:
        combined["Status"] = combined["Gain/Loss (%)"].apply(
            lambda x: "Up" if pd.notna(x) and x >= 0 else "Down"
        )
    return combined


def _json_array_candidates(text):
    """Yield progressively more permissive guesses at where the JSON array is
    within a messier-than-expected LLM response, in order of preference."""
    yield text

    fence_match = re.search(r"```(?:json)?\s*(\[.*?\])\s*```", text, re.DOTALL)
    if fence_match:
        yield fence_match.group(1)

    start, end = text.find("["), text.rfind("]")
    if start != -1 and end != -1 and end > start:
        yield text[start : end + 1]


def looks_like_truncated_json(parse_error_message, raw_text):
    """Infer truncation directly from the shape of the JSON parse failure,
    independent of whatever the provider's own finish/stop reason said. An
    'Unterminated string' error is only possible when the input ends mid-string
    — complete, valid JSON can never produce it — so it's treated as a certain
    signal on its own. Otherwise, if the reported error position is right at
    the end of the response, that's the same signature."""
    if not parse_error_message:
        return False
    if "Unterminated string" in parse_error_message:
        return True
    match = re.search(r"\(char (\d+)\)", parse_error_message)
    if match and raw_text:
        return int(match.group(1)) >= len(raw_text) - 5
    return False


def parse_ai_qualitative_json(raw_text):
    """Parse the LLM's JSON response into {symbol: {Momentum Pattern, Key News &
    Fundamentals, Outlook Note}}. Returns (qual_map, error_message) — error_message
    is None on success, so callers can show the raw text if parsing fails instead
    of crashing on a malformed response.

    Tries three ways to find the array, since models don't always follow "JSON
    only, no fences" exactly: (1) the whole response as-is, (2) a ```json fenced
    block found anywhere in the text (not just at the very start/end), (3) a
    plain slice from the first '[' to the last ']', which tolerates stray prose
    before/after the array."""
    if not raw_text or not raw_text.strip():
        return {}, "The AI returned an empty response. Try running the analysis again."

    text = raw_text.strip()
    data, last_err = None, None

    for candidate in _json_array_candidates(text):
        try:
            data = json.loads(candidate)
            break
        except json.JSONDecodeError as exc:
            last_err = exc

    if data is None:
        return {}, f"Could not parse the AI's response as JSON ({last_err}). Raw response below."

    qual_map = {}
    for item in data:
        symbol = str(item.get("stock", "")).strip()
        if symbol:
            qual_map[symbol] = {
                "Momentum Pattern": item.get("momentum_pattern", ""),
                "Key News & Fundamentals": item.get("key_news_fundamentals", ""),
                "Outlook Note": item.get("outlook_note", ""),
            }
    return qual_map, None


def merge_ai_performance_df(numeric_df, qual_map):
    """Left-join the app's own numeric data with the LLM's qualitative fields by
    stock symbol. Any stock the LLM didn't return (truncated response, etc.) gets
    a placeholder instead of silently vanishing from the table."""
    placeholder = {
        "Momentum Pattern": "—",
        "Key News & Fundamentals": "No analysis returned for this stock.",
        "Outlook Note": "—",
    }
    rows = []
    for _, row in numeric_df.iterrows():
        symbol = row["Stock Symbol"]
        qual = qual_map.get(symbol, placeholder)
        merged = {
            "Stock": symbol,
            "Status": row["Status"],
            "Current Price": row.get("Current Price"),
            "Gain/Loss (%)": row.get("Gain/Loss (%)"),
        }
        for col in ["1D Change (%)", "1W Change (%)", "1M Change (%)", "3M Change (%)", "6M Change (%)", "1Y Change (%)"]:
            if col in row:
                merged[col] = row[col]
        merged.update(qual)
        rows.append(merged)
    return pd.DataFrame(rows)



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


def run_ai_analysis(provider, api_key, model, sheets, max_tokens=4096):
    """Runs the batch Portfolio Overview + Performance Analysis report, grounded in
    live web search (news/earnings/fundamentals) via the same per-provider search
    plumbing as the chat/scanner features. Returns (report_text, model_used, truncated)."""
    data_prompt = build_analysis_prompt(sheets)
    num_stocks = sum(len(df[df["Stock Symbol"] != "TOTAL"]) for df in sheets.values())
    # One search per stock would be ideal but is impractical for larger portfolios
    # (slow, expensive, and Anthropic's tool caps searches per call anyway) — scale
    # the budget with portfolio size within a sane ceiling instead of a flat default.
    max_search_uses = min(max(num_stocks, 5), 40)
    history = [{"role": "user", "content": data_prompt}]
    return run_chat_turn(
        provider, api_key, model, AI_SYSTEM_PROMPT, history,
        max_tokens=max_tokens, max_search_uses=max_search_uses,
    )


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

OPENAI_SEARCH_MODEL_FALLBACK = "gpt-5-search-api"


def build_chat_system_prompt(exclude_symbols=None, scanner_report=None):
    """System prompt for the Scanner tab's chat: a live-search-grounded research
    companion for finding NEW Indian stock opportunities, explicitly excluding
    whatever the person already holds."""
    exclusion_block = ""
    if exclude_symbols:
        exclusion_block = f"""
THE PERSON ALREADY HOLDS THESE STOCKS — EXCLUDE THEM ENTIRELY:
{", ".join(exclude_symbols)}
Never analyze, recommend, or include any of these in a watchlist or summary, even if \
asked generally about "the market" or "good opportunities". If the person explicitly \
asks about one of these by name, tell them it's excluded because it's already in \
their portfolio, and redirect to other candidates instead.
"""

    context_block = ""
    if scanner_report:
        context_block = f"""
MOST RECENT CATALYST & MOMENTUM SCAN (for context — the person may ask follow-ups \
about specific stocks or tables in this):
{scanner_report}
"""

    return f"""You are a live-search-grounded Indian equity research assistant, for \
finding and discussing stock opportunities OUTSIDE the person's current portfolio. \
{exclusion_block}{context_block}
WHENEVER ASKED FOR ANALYSIS ON THE MARKET OR A SPECIFIC STOCK:
Use web search to gather, together, ALL of: (a) current news and events, (b) recent \
earnings results, (c) fundamentals (P/E, revenue/profit growth, guidance, balance \
sheet health where available), and (d) price/technical behavior. Do not answer from \
price momentum alone, and do not answer from general knowledge alone — combine \
what search actually returns across these areas.

DO NOT HALLUCINATE — THIS IS CRITICAL:
Every factual claim (a news event, an earnings figure, a price level, a date, a \
company fact) must come directly from an actual web search result you retrieved in \
this conversation, or from something the person told you. Never invent, assume, or \
fill in a plausible-sounding fact you did not actually find. If search turns up \
nothing relevant, or you are not confident in what it returned, say so plainly \
("I couldn't find reliable recent information on X") rather than guessing. When you \
do state a fact, mention its source (publication/site name) and approximate date so \
the person can verify it themselves.

STRICT RULES:
- Never state a future price, target, or return as if it were a fact. Frame all \
forward-looking statements as possibilities based on current momentum and/or news, \
never as certainties. Never use "guaranteed", "certain", or "will definitely".
- Never give direct buy/sell/hold instructions.
- Do not use "multibagger" as a promise — treat it as a high-risk, speculative \
label, and say so explicitly when you use the term.
- For casual/specific questions (e.g. about one stock), just answer directly — \
still grounded, sourced, and excluding the person's own holdings as above.
"""


def _reason_indicates_truncation(reason):
    """True if a provider's stop/finish reason indicates the response was cut
    off by the token limit. Deliberately robust rather than an exact string
    match: SDKs often return an enum object whose str() includes a class-name
    prefix (e.g. 'FinishReason.MAX_TOKENS', not just 'MAX_TOKENS'), which an
    exact comparison silently fails to match — this checks a substring of
    whichever text representation is available instead."""
    name = getattr(reason, "name", reason)
    text = str(name).upper()
    return "MAX_TOKEN" in text or "LENGTH" in text


def chat_anthropic(api_key, model, system_prompt, history, max_tokens=2048, max_search_uses=5):
    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model=model,
        max_tokens=max_tokens,
        system=system_prompt,
        messages=[{"role": m["role"], "content": m["content"]} for m in history],
        tools=[{"type": "web_search_20250305", "name": "web_search", "max_uses": max_search_uses}],
    )
    text = "\n".join(block.text for block in response.content if getattr(block, "type", "") == "text")
    if not text.strip():
        stop_reason = getattr(response, "stop_reason", "unknown")
        raise RuntimeError(
            f"Claude returned no text content (stop_reason={stop_reason}). This can "
            "happen if the response was cut off mid-search or hit the token limit "
            "before writing any output — try again, or raise max_tokens."
        )
    truncated = _reason_indicates_truncation(getattr(response, "stop_reason", None))
    return text, truncated


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
    content = response.choices[0].message.content
    finish_reason = getattr(response.choices[0], "finish_reason", "unknown")
    if not content or not content.strip():
        raise RuntimeError(
            f"OpenAI ({effective_model}) returned no text content (finish_reason="
            f"{finish_reason}). This can happen with search-enabled models when the "
            "response is cut off before producing text, or contains only tool-call "
            "data — try again, raise max_tokens, or try a different provider."
        )
    truncated = _reason_indicates_truncation(finish_reason)
    return content, effective_model, truncated


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
    text = getattr(response, "text", None)
    candidates = getattr(response, "candidates", None)
    finish_reason = getattr(candidates[0], "finish_reason", "unknown") if candidates else "unknown"
    if not text or not text.strip():
        raise RuntimeError(
            f"Gemini returned no text content (finish_reason={finish_reason}). This "
            "can happen if the response was blocked, truncated, or contained only "
            "search/tool data — try again, or raise max_tokens."
        )
    truncated = _reason_indicates_truncation(finish_reason)
    return text, truncated


def run_chat_turn(provider, api_key, model, system_prompt, history, max_tokens=2048, max_search_uses=5):
    """Returns (reply_text, model_actually_used, truncated) — model can differ
    from the configured one for OpenAI, since search needs a dedicated model.
    truncated=True means the response was cut off by the token limit, not that
    it finished naturally — callers should surface this rather than presenting
    a cut-off answer as if it were complete."""
    if provider == "OpenAI":
        return chat_openai(api_key, model, system_prompt, history, max_tokens=max_tokens)
    if provider == "Gemini":
        text, truncated = chat_gemini(api_key, model, system_prompt, history, max_tokens=max_tokens)
        return text, model, truncated
    text, truncated = chat_anthropic(
        api_key, model, system_prompt, history, max_tokens=max_tokens, max_search_uses=max_search_uses
    )
    return text, model, truncated


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

SCANNER_SYSTEM_PROMPT_BASE = """Act as a short-term Indian equity market research analyst and event-driven market researcher.

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

COVERAGE TARGET: aim to surface AT LEAST 10 stocks across the Main Market Scan Table (Phase 14) and the One-Week Watchlist (Phase 17). Reach this by broadening the scan across more sectors and catalyst types — including moderate-conviction setups, sector-wide movers, and names that land in "Already Extended" or "Weak Catalyst Reaction" — rather than stopping at the first few strong names. This coverage target NEVER overrides the rule above against inventing catalysts: if genuinely verifiable candidates are fewer than 10 after a thorough search, say so explicitly and list only what you can actually verify. A shorter, honest list is always better than a padded, fabricated one.

---

PHASE 20 - FINAL OUTPUT

{final_output_instructions}
"""

SCANNER_FINAL_OUTPUT_DETAILED = """Structure the final answer with these Markdown headers, in order:

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

The goal is not to predict the market. The goal is to systematically identify: fresh information -> initial price reaction -> volume confirmation -> catalyst reinforcement -> persistence -> momentum exhaustion, and use historical evidence to understand which patterns tend to persist beyond the first trading session."""

SCANNER_FINAL_OUTPUT_SUMMARY = """Structure the final answer with ONLY these two Markdown headers, in order. Do NOT include an Executive Summary, Historical Catalyst Persistence, Early-Warning Patterns, Detailed Candidate Analysis, or Key Takeaways section — omit them entirely, even briefly, even as a single sentence:

## 1. Today's Catalyst Scanner
The main market-wide table from Phase 14.

## 2. One-Week Watchlist
Separated into: Fresh Catalyst, Catalyst + Confirmation, Pullback + Catalyst, Breakout Watch, Persistent Repricing Candidate, Already Extended, Unexplained Spike, Weak Catalyst Reaction.

Still perform the full reasoning from every phase above internally (historical persistence, early-warning patterns, candidate deep-dive) to make these two tables accurate — you are only omitting those sections from the final written output, not skipping the underlying analysis."""


def build_scanner_system_prompt(detail_level="Detailed"):
    final_output = (
        SCANNER_FINAL_OUTPUT_SUMMARY if detail_level == "Summary" else SCANNER_FINAL_OUTPUT_DETAILED
    )
    return SCANNER_SYSTEM_PROMPT_BASE.format(final_output_instructions=final_output)



def run_scanner_analysis(provider, api_key, model, max_tokens, exclude_symbols=None, detail_level="Detailed"):
    """Single-turn, search-grounded run of the Catalyst & Momentum Scanner prompt.
    Reuses the same per-provider plumbing as the chat feature. Returns
    (reply_text, model_used, truncated)."""
    system_prompt = build_scanner_system_prompt(detail_level)
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
    # A moderate, deliberate search budget: enough for light cross-verification
    # across a ~10+ stock coverage target (Phase 19), without eating so much of
    # the same token pool that there's no room left to write the tables out —
    # the generic chat default of 5 undershoots what this task actually needs,
    # while going much higher would worsen the exact truncation this is tuned to avoid.
    return run_chat_turn(
        provider, api_key, model, system_prompt, history, max_tokens=max_tokens, max_search_uses=15
    )


# --------------------------------------------------------------------------------------
# Rendering the Scanner report with sortable tables for two specific sections.
# The report is otherwise free-text Markdown (headers vary in number between Summary
# and Detailed mode), so tables are located by title text, not a fixed heading number.
# --------------------------------------------------------------------------------------
SORTABLE_SCANNER_SECTIONS = ["Today's Catalyst Scanner", "One-Week Watchlist"]


def parse_markdown_table(block_text):
    """Parse Markdown pipe-table(s) found in a block of text into one DataFrame.
    Handles the model splitting a section into several mini-tables that each
    repeat the header + separator row (e.g. one per watchlist category) by
    dropping every repeated header/separator occurrence, not just the first —
    otherwise a repeated header row gets mis-parsed as a literal data row.
    Returns None if no valid table is found (caller should fall back to
    rendering the raw text instead of crashing)."""
    lines = [l for l in block_text.splitlines() if l.strip().startswith("|")]
    if len(lines) < 2:
        return None

    def split_row(line):
        cells = line.strip().strip("|").split("|")
        return [c.strip() for c in cells]

    header = split_row(lines[0])
    header_key = [c.lower() for c in header]

    def is_separator(cells_or_line):
        return bool(re.match(r"^[\s|:-]+$", cells_or_line))

    data_rows = []
    for line in lines[1:]:
        if is_separator(line):
            continue
        cells = split_row(line)
        if [c.lower() for c in cells] == header_key:
            continue  # a repeated header row from a second/third mini-table
        data_rows.append(cells)

    # Guard against rows with a different cell count than the header (a model
    # formatting slip) rather than letting pandas raise on a ragged table.
    data_rows = [r for r in data_rows if len(r) == len(header)]
    if not data_rows:
        return None
    return pd.DataFrame(data_rows, columns=header)


def coerce_numeric_columns(df, min_success_ratio=0.7):
    """Convert columns that are mostly numeric-looking text (e.g. '+12.3%', '2.1x',
    '₹1,234') into real numeric dtype, so st.dataframe sorts them numerically
    instead of alphabetically. Columns that are genuinely textual are left alone."""
    df = df.copy()
    for col in df.columns:
        cleaned = df[col].astype(str).str.replace(r"[₹$,%x]", "", regex=True).str.strip()
        numeric = pd.to_numeric(cleaned, errors="coerce")
        non_empty = df[col].astype(str).str.strip().replace({"": None, "-": None, "N/A": None}).notna()
        if non_empty.sum() == 0:
            continue
        success_ratio = numeric.notna().sum() / non_empty.sum()
        if success_ratio >= min_success_ratio:
            df[col] = numeric
    return df


def split_report_into_sections(report_text):
    """Split a Markdown report into (heading_line_or_None, body_text) chunks at
    each '## ' header, preserving order. The first chunk (before any header) has
    heading=None."""
    parts = re.split(r"(?m)^(## .+)$", report_text)
    sections = []
    if parts[0].strip():
        sections.append((None, parts[0]))
    for i in range(1, len(parts), 2):
        heading = parts[i]
        body = parts[i + 1] if i + 1 < len(parts) else ""
        sections.append((heading, body))
    return sections


def render_scanner_report(report_text):
    """Render the report section-by-section: the two target tables as sortable
    st.dataframe widgets, everything else as plain Markdown, in original order."""
    for heading, body in split_report_into_sections(report_text):
        if heading is None:
            if body.strip():
                st.markdown(body)
            continue

        st.markdown(heading)
        matched_section = next((s for s in SORTABLE_SCANNER_SECTIONS if s in heading), None)
        if matched_section:
            table_df = parse_markdown_table(body)
            if table_df is not None:
                st.dataframe(coerce_numeric_columns(table_df), use_container_width=True, hide_index=True)
                # Show any text in the section besides the table itself (e.g. a
                # one-line note the model added) rather than silently dropping it.
                remainder = re.sub(r"(?m)^\|.*\|\s*$", "", body).strip()
                if remainder:
                    st.markdown(remainder)
                continue
        # Fallback: couldn't confidently parse a table here, or this section
        # isn't one of the two targeted for sorting — render as-is.
        st.markdown(body)



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
    # Always run at the most detailed budget — no control shown, per request.
    # 16384, not 8192 — for a full portfolio, max_search_uses scales up to 40
    # searches, and search steps themselves consume tokens from this same budget
    # before the model ever writes the final JSON. 8192 was too tight and caused
    # truncation mid-response for larger portfolios.
    AI_MAX_TOKENS = 16384
    ai_provider = st.radio("Provider", options=["OpenAI", "Anthropic", "Gemini"], horizontal=True)

    def model_picker(label, options, help_text):
        """Dropdown with a 'Custom...' escape hatch — model names shift often
        enough that a hardcoded list can go stale, so there's always a way to
        type an exact string instead of being stuck with what's listed here."""
        choice = st.selectbox(label, options + ["Custom..."], index=0, help=help_text)
        if choice == "Custom...":
            return st.text_input(f"{label} (custom)", placeholder="Type the exact model name")
        return choice

    if ai_provider == "OpenAI":
        ai_api_key = st.text_input(
            "OpenAI API key",
            value=get_secret("OPENAI_API_KEY"),
            type="password",
        )
        ai_model = model_picker(
            "Model",
            ["gpt-5-search-api", "gpt-5.5", "gpt-5.5-pro", "gpt-4o-mini"],
            "gpt-5-search-api is built for search-grounded tasks like the Scanner. "
            "gpt-5.5-pro is the strongest general option; gpt-4o-mini is cheap/fast "
            "but more prone to cutting off long, complex outputs. Any OpenAI chat "
            "model works if you pick Custom — this list can go stale.",
        )
    elif ai_provider == "Gemini":
        ai_api_key = st.text_input(
            "Gemini API key",
            value=get_secret("GEMINI_API_KEY", "GOOGLE_API_KEY"),
            type="password",
        )
        ai_model = model_picker(
            "Model",
            ["gemini-3.1-pro-preview", "gemini-3-flash-preview", "gemini-2.5-flash-lite"],
            "The 'pro' tier is strongest for long, complex output like the Scanner. "
            "'Flash' tiers are cheaper/faster but more prone to cutting off long "
            "output. Any Gemini model works if you pick Custom — naming here "
            "shifts often, so double-check against ai.google.dev/gemini-api/docs/models "
            "if an option fails.",
        )
    else:
        ai_api_key = st.text_input(
            "Anthropic API key",
            value=get_secret("ANTHROPIC_API_KEY"),
            type="password",
        )
        ai_model = model_picker(
            "Model",
            ["claude-sonnet-5", "claude-opus-5-5", "claude-haiku-4-5-20251001"],
            "Sonnet balances quality and cost; Opus is strongest for long, complex "
            "output like the Scanner; Haiku is cheapest/fastest.",
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
        st.header("🤖 AI Performance Analysis")
        st.warning(AI_DISCLAIMER)
        st.caption(
            "Now grounded in live web search per stock (news, earnings, fundamentals), "
            "not just the price data — this means it can take noticeably longer than "
            "before, especially for larger portfolios."
        )

        if st.button("Run Performance Analysis"):
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
            elif not ai_model:
                st.error("Enter a model name in the sidebar first.")
            else:
                with st.spinner("Researching and analyzing (very detailed) — this involves live searches..."):
                    try:
                        raw_text, model_used, truncated = run_ai_analysis(
                            ai_provider, ai_api_key, ai_model, sheets, max_tokens=AI_MAX_TOKENS
                        )
                        qual_map, parse_error = parse_ai_qualitative_json(raw_text)
                        st.session_state["ai_overview_df"] = compute_portfolio_overview_df(sheets)
                        st.session_state["ai_model_note"] = (
                            f"Used {model_used} for web search — your configured model doesn't support it."
                            if model_used != ai_model
                            else None
                        )
                        # OR with a content-based check: don't rely solely on the provider's
                        # own finish-reason flag, since that's proven unreliable to parse
                        # correctly across SDKs — the shape of the JSON failure itself is
                        # independent evidence.
                        truncated = truncated or looks_like_truncated_json(parse_error, raw_text)
                        st.session_state["ai_truncated"] = truncated
                        if parse_error:
                            if truncated:
                                parse_error = (
                                    "This response was cut off by the token limit before it finished "
                                    "writing — that's why it doesn't parse as valid JSON (a broken string "
                                    "or missing bracket at the cut point), not a formatting mistake. "
                                    "A large portfolio needing many searches can use up most of the "
                                    "budget before writing any output. Try again, or this portfolio may "
                                    "need a higher AI_MAX_TOKENS than the app currently uses."
                                )
                            st.session_state["ai_performance_df"] = None
                            st.session_state["ai_parse_error"] = parse_error
                            st.session_state["ai_raw_response"] = raw_text
                        else:
                            numeric_df = compute_stock_numeric_df(sheets)
                            st.session_state["ai_performance_df"] = merge_ai_performance_df(numeric_df, qual_map)
                            st.session_state["ai_parse_error"] = None
                    except Exception as exc:  # noqa: BLE001 - surface any API error to the user
                        st.error(f"AI analysis failed: {exc}")

        if st.session_state.get("ai_truncated"):
            st.warning(
                "⚠️ The last response hit the token limit before finishing — treat it as "
                "incomplete. See the note below for why and what to try."
            )
        if st.session_state.get("ai_model_note"):
            st.caption(f"*{st.session_state['ai_model_note']}*")

        if st.session_state.get("ai_overview_df") is not None and not st.session_state["ai_overview_df"].empty:
            st.subheader("Portfolio Overview")
            st.dataframe(st.session_state["ai_overview_df"], use_container_width=True, hide_index=True)

        if st.session_state.get("ai_parse_error"):
            st.error(st.session_state["ai_parse_error"])
            with st.expander("Raw AI response"):
                st.text(st.session_state.get("ai_raw_response", ""))

        if st.session_state.get("ai_performance_df") is not None:
            st.subheader("Performance Analysis")
            status_filter = st.radio(
                "Show", ["All", "Up", "Down"], horizontal=True, key="ai_status_filter"
            )
            perf_df = st.session_state["ai_performance_df"]
            if status_filter != "All":
                perf_df_display = perf_df[perf_df["Status"] == status_filter]
            else:
                perf_df_display = perf_df
            st.dataframe(perf_df_display, use_container_width=True, hide_index=True)

            excel_bytes_ai = build_excel_bytes(
                {
                    "Portfolio Overview": st.session_state["ai_overview_df"],
                    "Performance Analysis": perf_df,
                }
            )
            st.download_button(
                "⬇️ Download this analysis (.xlsx)",
                data=excel_bytes_ai,
                file_name="portfolio_ai_analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

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

    # Always run at full depth — this prompt asks for many tables plus a deep
    # dive, and thin token budgets were cutting analysis short. No control
    # shown for this since there's no good reason to run it any shallower.
    # 32768, not 16384 — even Summary mode's two tables (each covering the
    # >=10-stock coverage target) plus the search steps consuming part of the
    # same budget was still running out before finishing the last section.
    SCANNER_MAX_TOKENS = 32768

    scanner_output_level = st.radio(
        "Output length",
        ["Summary", "Detailed"],
        index=1,
        horizontal=True,
        help="Summary keeps only the Catalyst Scanner table and the One-Week "
        "Watchlist. Detailed adds Executive Summary, Historical Catalyst "
        "Persistence, Early-Warning Patterns, Detailed Candidate Analysis, and "
        "Key Takeaways. Both do the same full research — Summary just reports "
        "less of it.",
    )

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
        elif not ai_model:
            st.error("Enter a model name in the sidebar first.")
        else:
            with st.spinner("Scanning (comprehensive) — this involves several live searches..."):
                try:
                    report, model_used, truncated = run_scanner_analysis(
                        ai_provider, ai_api_key, ai_model, SCANNER_MAX_TOKENS,
                        holdings_to_exclude, detail_level=scanner_output_level,
                    )
                    prefix = ""
                    if model_used != ai_model:
                        prefix += (
                            f"*(used {model_used} for web search — your configured "
                            f"model doesn't support it)*\n\n"
                        )
                    if truncated:
                        prefix += (
                            "⚠️ **This scan hit the token limit before finishing — it's cut off, "
                            "not a complete report.** This prompt asks for a lot of output (many "
                            "tables plus a multi-stock deep dive). Try 'Summary' output length "
                            "above, which asks for far less text and is much more likely to "
                            "finish completely, or try again since search-heavy runs vary in length.\n\n"
                        )
                    st.session_state["scanner_report"] = prefix + report
                except Exception as exc:  # noqa: BLE001 - surface any API error to the user
                    st.error(f"Scan failed: {exc}")

    if st.session_state.get("scanner_report"):
        render_scanner_report(st.session_state["scanner_report"])
        st.download_button(
            "⬇️ Download this scan (.md)",
            data=st.session_state["scanner_report"].encode("utf-8"),
            file_name="catalyst_momentum_scan.md",
            mime="text/markdown",
        )

    st.divider()
    st.header("💬 Chat: Performance & Multibagger Potential (grounded in web/news search)")
    st.warning(CHAT_DISCLAIMER)
    if holdings_to_exclude:
        st.caption(
            f"🚫 Excluding {len(holdings_to_exclude)} stock(s) already in your portfolio "
            "from anything discussed here — same as the scan above."
        )

    if "chat_messages" not in st.session_state:
        st.session_state["chat_messages"] = []

    col_a, col_b = st.columns([3, 1])
    with col_b:
        if st.button("🗑️ Clear chat"):
            st.session_state["chat_messages"] = []
            st.rerun()
    with col_a:
        preset_label = (
            "📋 Summarize today's scan findings"
            if st.session_state.get("scanner_report")
            else "📋 Give me today's top opportunities"
        )
        if st.button(preset_label):
            st.session_state["chat_messages"].append(
                {
                    "role": "user",
                    "content": (
                        "Summarize the strongest candidates found, with their key news, "
                        "fundamentals, and price momentum for each — grounded in live search, "
                        "excluding anything I already hold."
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
        elif not ai_model:
            st.error("Enter a model name in the sidebar first.")
            st.session_state["_chat_pending"] = False
        else:
            with st.spinner("Searching and responding..."):
                try:
                    system_prompt = build_chat_system_prompt(
                        exclude_symbols=holdings_to_exclude,
                        scanner_report=st.session_state.get("scanner_report"),
                    )
                    reply, model_used, truncated = run_chat_turn(
                        ai_provider, ai_api_key, ai_model, system_prompt, st.session_state["chat_messages"]
                    )
                    if model_used != ai_model:
                        reply = f"*(used {model_used} for web search — your configured model doesn't support it)*\n\n{reply}"
                    if truncated:
                        reply += "\n\n⚠️ *This reply hit the token limit before finishing — it may be cut off.*"
                    st.session_state["chat_messages"].append({"role": "assistant", "content": reply})
                except Exception as exc:  # noqa: BLE001 - surface any API error to the user
                    st.session_state["chat_messages"].append(
                        {"role": "assistant", "content": f"⚠️ Request failed: {exc}"}
                    )
                st.session_state["_chat_pending"] = False
                st.rerun()
