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
    "⚠️ **Speculative fundamental-screening tool — not financial advice.** This runs "
    "a quantitative discovery framework (Altman Z-Score, Piotroski F-Score, 5 "
    "screeners, technicals, shareholding trends) using live web search for the "
    "underlying data — it does not give price targets, entry/exit levels, or "
    "position sizing as guarantees, even though the framework computes them "
    "explicitly (the prompt frames these as a structured exercise, not advice). "
    "Important limitation: Z-Score and F-Score need precise balance-sheet figures "
    "(EBIT, working capital, total liabilities, OCF, etc.) — general web search is "
    "not built for retrieving structured financial-statement data reliably at scale, "
    "so any score here may be based on incomplete or estimated inputs, which the "
    "model is instructed to flag rather than hide. Verify any number that matters "
    "to you against a primary source (screener.in, the company's own filings) before "
    "acting. Consult a licensed financial advisor before making investment decisions."
)

SCANNER_SYSTEM_PROMPT_BASE = """You are an expert Indian equity research analyst with deep knowledge of
NSE/BSE listed stocks. Run a comprehensive multibagger discovery and analysis
framework. Follow every step in sequence.

DATA-INTEGRITY MANDATE (applies to every step below):
This framework requires precise quantitative inputs (Working Capital, Retained
Earnings, EBIT, Total Assets, Total Liabilities, Revenue, ROA, OCF, debt ratios,
current ratio, share issuance, gross margin, asset turnover, FII/DII/promoter
holding %, RSI, MACD, moving averages, and more) for every candidate stock. You
have live web search, not a structured financial database — search coverage for
line-item-level balance sheet data is uneven, especially for mid/small caps.
NEVER estimate, infer, or invent a number you could not actually verify via
search. For any metric you cannot verify: state "data unavailable" for that
specific field, show the Z-Score or F-Score as incomplete/partial rather than
computing it from assumed values, and do not let a stock advance through a
screener on an unverified pass. Cite a source (publication/site name, approx.
date) for the key figures behind each score wherever you can. A shorter, honest
result set is always preferable to a fully-populated one built on invented data.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 0 — Z-SCORE & F-SCORE PRE-FILTER (Run before everything)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

A. ALTMAN Z-SCORE (Safety gate — eliminates bankruptcy risk)
Z = 1.2(WC/TA) + 1.4(RE/TA) + 3.3(EBIT/TA) + 0.6(MCap/TL) + 1.0(Rev/TA)
-> Z > 3.0 = Safe | 2.6-3.0 = Grey | 1.8-2.6 = Caution | <1.8 = ELIMINATE
Exception: Cyclicals (steel/cement/mining) may structurally show low Z.
Non-manufacturing threshold: Safe > 2.9, Distress < 1.23.

B. PIOTROSKI F-SCORE (9 binary signals — score 1 if true, 0 if false)
Profitability:  F1=ROA>0 | F2=OCF>0 | F3=ROA improving YoY | F4=OCF/TA>ROA
Leverage:       F5=LT debt ratio falling | F6=Current ratio rising | F7=No dilution
Efficiency:     F8=Gross margin rising | F9=Asset turnover rising
-> 8-9=Strong Buy | 6-7=Good | 4-5=Neutral | <4=ELIMINATE
Track trend across 3 quarters — a RISING trend matters more than the number.

C. COMBINED LABEL (assign before screening)
ELITE:  Z>3.0 + F=8-9 -> Maximum conviction, fast-track
STRONG: Z>2.6 + F=7-8 -> High conviction, full analysis
WATCH:  Z>2.0 + F=5-7 -> Monitor, small initial position
AVOID:  Z<2.0 OR F<4  -> No investment, no exceptions

D. 4 PATTERNS TO HUNT (flag if present)
P1 — F-Score Breakout:    F jumps 4->8 while RSI<60 + FII<5% -> Buy before re-rating
P2 — Z-Score Recovery:    Z crosses 2.6 + F=7+ + PE still at discount -> 3-5x potential
P3 — Consistent Quality:  Z>3 + F=8-9 for 3+ qtrs + PE discount -> quality compounder
P4 — Institutional Trigger: F rises 5->8 + FII adds 3%+ + RSI crosses 50 -> highest conviction

Output: | Stock | Z-Score | Zone | F-Score | F-Trend | Pattern | Label |

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 1 — RUN 5 PARALLEL SCREENERS (Nifty 500 — ELITE/STRONG/WATCH only)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

A — QUALITY + VALUE (GARP)
PE > 25% below industry PE | ROE > 15% | D/E < 0.5
3 years consecutive profit | Market cap Rs.500Cr-Rs.50,000Cr | Dividend > 0%

B — TURNAROUND + ACCELERATION
PAT QoQ > 50% for 2+ quarters | Revenue QoQ > 15% for 3+ quarters
OPM expanding QoQ | Recent loss-to-profit shift OR margin expanding >10% YoY

C — INSTITUTIONAL DISCOVERY (Smart Money)
FII up >2% QoQ for 2 quarters OR MF up >3% QoQ | Both buying = strongest signal
Promoter stable/rising, no pledge increase | Mid/small cap preferred (pre-discovery)

D — CASHFLOW QUALITY
OCF positive + growing >20% YoY | OCF/Net Profit > 0.8 | ROCE > 15%
Free cashflow positive | Working capital cycle stable or improving

E — TECHNICAL BREAKOUT
RSI 50-70 | MACD above signal + histogram positive | Price above 50 & 200 DMA
Supertrend bullish 3+ months | 3M price change >10% | Volume >1.5x avg on breakouts

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 2 — CROSS-SCREENER SCORING & TIER ASSIGNMENT
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Base tier from screener count:
1 = Watch | 2 = Tier 3 | 3 = Tier 2 | 4 = Tier 1 | 5 = Tier 1 Elite

F-Score adjustments:
F=8-9: Upgrade 1 Tier | F=4-5: Downgrade 1 Tier | Pattern 4: Auto-promote Elite

Output: | Stock | Mkt Cap | PE vs Ind | ROE | D/E | Z | F | F-Trend | Screeners | Tier |

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 3 — DEEP ANALYSIS (Top 15 stocks scoring 2+ screeners)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

A. VALUATION: PE vs industry (discount%), PB vs industry, EV/EBITDA, PEG ratio
B. GROWTH: Revenue CAGR 1/3/5Y | PAT CAGR 1/3/5Y | OPM & NPM trend 8 quarters
C. BALANCE SHEET: D/E trend 3Y | Interest coverage | Current ratio | Pledge %
D. CASHFLOW: OCF last 4 years | FCF = OCF-Capex | OCF as % of PAT (target >80%)
E. Z & F DEEP READ:
   - All 5 Z components (X1-X5): identify weak vs strong variables
   - All 9 F signals individually: which category is strongest/weakest
   - F-Score trend last 3 quarters: rising/stable/falling
   - Identify which Pattern (P1/P2/P3/P4) applies
F. MOAT: Market position | Sector tailwind | Entry barriers | Client quality

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 4 — SHAREHOLDING ANALYSIS (Last 4-6 quarters)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

PROMOTER: Stable/rising = good | Pledge rising = red flag | Open market buying = strong
FII: Rising sharply = discovery | New entry from 0% = highest signal | Falling = caution
Cross-check: FII buying + F rising = P4 trigger | FII buying + F falling = caution
DII: Both FII+DII buying = maximum conviction | DII buying while FII exits = transition
RETAIL: Rising retail + falling institutions = speculative (not multibagger quality)

Phase labels:
DISCOVERY: FII/MF <10%, just entering
ACCUMULATION: Institutions building, 10-25%
MATURITY: Institutions >25%, re-rating done
EXIT: FII/DII consistently reducing

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 5 — TECHNICAL ANALYSIS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

TREND: Price vs 200 DMA (%) | SMA ladder 10>20>50>100>200 = bull stack
       52W position = (LTP-52WL)/(52WH-52WL)
MOMENTUM: RSI 50-70 = sweet spot | Stochastic + Williams %R for confirmation
MACD: Above signal + histogram expanding = bullish | Divergence = warning
STRENGTH: ADX >25 = trending | Supertrend GREEN = bullish
VOLATILITY: BB position | ATR for stop sizing | Beta for position sizing
SUPPORT/RESISTANCE: Pivot = (H+L+C)/3 | R1/R2/R3 targets | S1/S2/S3 stops
VOLUME: OBV rising = confirmed uptrend | Delivery >50% = conviction buying

Confluence signals:
RSI 50-65 + F=8-9 = Perfect accumulation | RSI>75 + F falling = Sell
RSI<35 + F rising = Strong oversold buy | RSI crosses 50 + F hits 7+ = P4 alert

Technical grade:
A+: RSI 50-65, MACD+, above all SMAs, ST green, breakout
A:  RSI 50-70, MACD+, above 200DMA, ST green
B:  RSI 40-70, mixed, above 200DMA
C:  Below 200DMA OR RSI<40 OR MACD-
D:  All bearish — avoid

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 6 — SECTOR & MACRO THEME MAPPING
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Tier 1 (2025-2030 strongest): Energy Transition | Defense & Aerospace
  AI & Data Infrastructure | Infrastructure & Capex | Pharma CDMO/API

Tier 2 (3-5Y visibility): Cables & Wires | Specialty Chemicals
  Capital Markets | Auto Ancillary | Real Estate & REITs

Tier 3 (moderate): Agri & Food | Healthcare | Consumer Discretionary
  Telecom Infrastructure | Logistics

Rate each: Theme Tier (1/2/3) + Strength (Strong/Moderate/Early)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 7 — RISK SCORING (Base 5, adjust up/down)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Adjustments:
-2: Z>3.0 + F=8-9      +1: F falling 2+ quarters
-1: Z>2.6 OR F=7-8     +1: RSI>75 or >50% above 200DMA
+1: Promoter pledge >10%  +1: Revenue concentration >50% top 3 clients
+2: Z<2.0 OR F<4 (should have been eliminated)

Z-specific flags: X1 negative (liquidity) | X3 falling (efficiency) | X4 falling (debt growing)
F-specific flags: F4=0 (profit not cash) | F7=0 (dilution) | F5+F6=0 (double stress)

Score 1-3 = Low risk (8-10% allocation) | 4-6 = Moderate (5-7%)
      7-8 = High risk (2-4%) | 9-10 = Speculative (1-2%)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 8 — TRADE PLAN
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ENTRY: Staggered — 50% at support/RSI zone, 50% on dip
Avoid entry: RSI>75 | >40% above 200DMA | F-Score just dropped 2+ points
Trigger: Volume breakout + RSI 50-65 + F-Score stable or rising

TARGETS (framed as structured framework output, not a guarantee):
6M: Resistance 1 / Pivot levels
12M: Resistance 2 / PE rerating to industry
24-36M: Forward EPS (2Y) x Industry PE x discount factor
  ELITE stocks: x1.0 | STRONG: x0.85 | WATCH: x0.70

STOP LOSS:
Technical: Below 200 DMA (investors) | Below Support 1 (traders)
Fundamental: Exit immediately if F-Score drops below 4
Time-based: Exit if thesis not playing in 6 months

POSITION SIZING (illustrative allocation based on the framework's own rules,
not personalized financial advice):
Risk score 1-3: up to 10% | 4-6: up to 7% | 7-8: up to 4% | 9-10: up to 2%
Boost +1% if Z>3.0 + F=9 | Cut -1% if F-Score fell last quarter
Hard cap: No single stock >10% | Diversify across 3+ themes

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 9 — MASTER SCORECARD
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

| # | Stock | Cap | Screeners | Z | F | F-Trend | Pattern | Tech | Phase | Theme | Risk | Entry | T-12M | Stars |

5 stars = ELITE + 4-5 screeners + A tech + Discovery phase
4 stars = STRONG + 3 screeners + A/B tech + positive institutional
3 stars = WATCH + 2 screeners + 1 strong signal + macro theme
2 stars = 1 screener + technically sound (watch only)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 10 — MONITORING & EXIT FRAMEWORK
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

MONTHLY: RSI/MACD positive? | Above 50 DMA? | No negative news? | FII stable?

QUARTERLY: Recalculate Z + F fresh | Revenue/PAT vs expectation
PAT margin trend | Debt direction | Promoter pledge change | Guidance

EXIT NOW (any one trigger):
- F-Score drops below 4 in any quarter
- Z-Score drops below 1.8
- Z-Score falling 3 consecutive quarters
- Promoter pledge >20%
- Revenue falling 2 consecutive quarters
- FII selling >5% in one quarter
- RSI<30 + price below 200 DMA simultaneously
- Core thesis invalidated by external event

PARTIAL EXIT 25-50% (any one trigger):
- RSI>80 + stock >50% above 200 DMA
- F-Score drops 2 points in one quarter
- Target hit in under 6 months (book 50%)
- Single quarter earnings miss
- Z-Score drops from Safe to Grey Zone

RE-ENTRY (all conditions together):
- RSI cools to 45-55 | Price at 50 DMA or Support 1
- F-Score recovers to 7+ | Earnings confirm thesis intact

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CONSTRAINTS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

- NSE/BSE only | Min 3Y listing history | Min Rs.500Cr market cap
- Max Rs.75,000Cr (large cap only if Z>3 + F=9)
- No SEBI investigation | No auditor resignation | Data within 1 quarter
- Never include Z<1.8 stocks | Never include F<4 unless rising trend + Z>3

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT FORMAT REQUIREMENTS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

{final_output_instructions}
"""


SCANNER_FINAL_OUTPUT_DETAILED = """Structure the final answer with these Markdown headers, in order:

## 1. Pre-Filter Table
Z-Score + F-Score + Label for all candidates, from Step 0.

## 2. Screening Results
The cross-screener table with qualifying stocks and their Tier, from Step 2.

## 3. Top 15 Stocks — Detailed Analysis
For each stock that scored 2+ screeners (up to 15), cover: Pre-Filter Summary
(Z-Score with all 5 components, F-Score with all 9 signals, F-Score trend,
Pattern label, final label), Multibagger Thesis (1 paragraph), Key Metrics
Table (PE discount, ROE, Debt, Revenue CAGR, Profit CAGR, OCF health),
Shareholding Signal (FII/DII/Promoter trend + phase label, cross-checked
against F-Score trend), Technical Summary (RSI, MACD, SMA position,
Supertrend grade, technical/F-Score confluence), Sector Theme (Theme Tier +
strength + multi-year vision), Risk Scorecard (score 1-10 with the specific
Z-Score, F-Score, technical, and sector risks behind it), and Trade Plan
(entry zone, 6M/12M/24-36M targets, stop loss including the fundamental stop
of F-Score < 4, position sizing % — framed as a structured output of the
rules in Step 8, not a prediction or instruction).

## 4. Master Scorecard Table
The one-line-per-stock summary table from Step 9.

## 5. Portfolio Construction
Suggested allocation % across the qualifying picks, from Step 8's position
sizing rules — a structured illustration of the framework's own rules applied
to this result set, not a personalized recommendation.

## 6. Monitoring Checklist
The monthly and quarterly review items, and the exit/partial-exit/re-entry
triggers, from Step 10."""

SCANNER_FINAL_OUTPUT_SUMMARY = """Structure the final answer with ONLY these Markdown headers, in order. Do NOT include a "Top 15 Stocks — Detailed Analysis" section — omit it entirely, even briefly, even as a single sentence per stock:

## 1. Pre-Filter Table
Z-Score + F-Score + Label for all candidates, from Step 0.

## 2. Screening Results
The cross-screener table with qualifying stocks and their Tier, from Step 2.

## 3. Master Scorecard Table
The one-line-per-stock summary table from Step 9 — this is where the per-stock
Z-Score, F-Score, screener count, tier, risk score, and target/entry columns
live in Summary mode, since the detailed per-stock write-up is omitted.

## 4. Portfolio Construction
Suggested allocation % across the qualifying picks, from Step 8's position
sizing rules.

## 5. Monitoring Checklist
The monthly and quarterly review items, and the exit/partial-exit/re-entry
triggers, from Step 10.

Still perform the full reasoning from every step above internally (fundamental
drill, shareholding deep-dive, full technical analysis, risk assessment) to
make these tables accurate — you are only omitting the written-out per-stock
deep-dive from the final output, not skipping the underlying analysis."""


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
    user_prompt = "Run the full framework now, following every step exactly as specified."
    if exclude_symbols:
        user_prompt += (
            "\n\nThe user already holds the following stocks. This scan is for "
            "finding NEW opportunities outside the current portfolio, so EXCLUDE "
            "all of these entirely from every screener, table, and the per-stock "
            "deep dive — even if one of them would otherwise score well: "
            + ", ".join(exclude_symbols)
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
# Momentum Trading Scanner — a third, separate framework: short-term (1 day-8 weeks)
# catalyst+volume+technical confluence trading, distinct from the long-term
# fundamental Multibagger framework above. Shares the same run_chat_turn plumbing.
# --------------------------------------------------------------------------------------
MOMENTUM_DISCLAIMER = (
    "⚠️ **Speculative short-term trading tool — not financial advice.** This runs a "
    "momentum framework (catalyst + volume + 7-indicator technical confluence) using "
    "live web search for the underlying data. Two limitations matter more here than "
    "elsewhere in this app: (1) technical readings (RSI, MACD, Supertrend, volume "
    "ratio, delivery %, PCR, VIX) require live market-data feeds that general web "
    "search only partially and inconsistently surfaces — the model is instructed to "
    "flag what it can't verify rather than invent it; (2) **every price level here "
    "(entry, stop, target) is only as fresh as whatever search last saw — it is NOT "
    "live data**, and can lag real price by minutes to hours. Short-term trading on a "
    "stale entry trigger is a real way to lose money. Verify the actual current price "
    "and volume yourself before acting on anything here. Nothing here is a "
    "recommendation to buy, hold, or sell. Consult a licensed financial advisor and/or "
    "your own broker terminal before trading."
)

MOMENTUM_SYSTEM_PROMPT_BASE = """You are an expert Indian equity momentum trader with deep knowledge of
NSE/BSE markets, intraday and swing trading setups, and short-term
price action. Run a comprehensive momentum stock discovery framework for
SHORT-TERM profit capture.

Target holding period: 1 day to 8 weeks maximum.
Goal: Identify stocks with the highest probability of a 5-25% move
in the shortest possible time, backed by multiple confirming signals.

Follow every step in sequence. Speed and signal confluence matter
more than deep fundamental analysis here.

DATA-INTEGRITY & FRESHNESS MANDATE (applies to every step below):
This framework asks for exact, current values: volume ratio, delivery %, RSI,
MACD, Supertrend, EMA structure, Bollinger Bands, ADX, VWAP, India VIX,
put-call ratio, and specific entry/stop/target PRICES. You have live web
search, not a real-time market-data feed. Two distinct issues follow:
(1) Verifiability — for any indicator or figure you cannot actually find via
search, state "data unavailable" for that specific field rather than
estimating a plausible-looking value. Do not let a stock pass a filter on an
unverified figure. Cite a source (site name, approx. date/time) for the key
figures behind each score wherever you can.
(2) Freshness — every price, volume, and indicator reading you find via
search reflects whatever moment that source captured it, not the current
live market. Explicitly label prices as "as of [the date/time you found],
verify current price before acting" rather than presenting them as live. This
matters far more here than in a long-horizon analysis: a specific trigger like
"buy above Rs.485" can already be stale by the time it's read.
A shorter, honest result set is always preferable to a fully-populated one
built on invented or unverifiably-fresh data.

Philosophy: Buy what is moving NOW, ride the wave, exit before it stops.
(Contrast with the separate Multibagger tool in this app, which buys great
businesses at fair prices and holds for 12-36 months — this is the opposite:
1 day to 8 weeks, technical/volume/catalyst confluence over fundamentals.)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 0 — LIQUIDITY FILTER (Hard prerequisites — no exceptions)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

MINIMUM REQUIREMENTS:
- Avg daily volume > 5L shares OR Rs.5Cr daily turnover
- Market cap > Rs.1,000Cr | Listed > 1 year | Bid-ask spread < 0.3%
- F&O eligibility preferred (better price discovery + hedging)

UNIVERSE BY TIMEFRAME:
Intraday (same day):   Nifty 50/Next50/Bank only | Vol >20L/day
Swing (2D-4W):         Nifty 500 | Vol >5L/day | Cap >Rs.2,000Cr
Positional (4-8W):     Nifty 500 + SME leaders | Vol >2L/day

ELIMINATE IMMEDIATELY:
- SEBI surveillance/investigation | Circuit hits in last 5 sessions
- LTP < Rs.20 (penny/manipulation risk) | Promoter pledge >50%
- Insolvency/court proceedings pending | Vol <2L/day

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 1 — CATALYST SCANNER (Every move needs a reason)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

TIER 1 — Strongest, most sustained (5-25%+, 3-15 days):
CAT-A: Earnings surprise — PAT >40% YoY + >20% QoQ + margin up 300bps
CAT-B: Major order/contract — value >10% annual revenue (defense/infra = highest)
CAT-C: Policy/sector tailwind — PLI, Budget allocation, regulatory change
CAT-D: FII block deal — >1% shares in one session OR FII net buy >Rs.500Cr/week

TIER 2 — Medium strength (3-15%, 2-8 days):
CAT-E: Analyst upgrade — rating to BUY + target >20% above LTP
CAT-F: Promoter/insider buying — >0.5% shares open market | Buyback announced
CAT-G: 52-week high breakout — new high on volume >2x average (65% success rate)
CAT-H: Index inclusion — passive fund buying creates guaranteed demand 10-15 days

TIER 3 — Weakest, shortest (2-8%, 1-3 days):
CAT-I: Media/social buzz — often already priced in, exit faster
CAT-J: Pure technical breakout — works only in bull market, strict SL required

Score each stock: Tier (1/2/3) + Type (A-J) + Freshness (Hours/1D/2-3D/Stale)
Rule: Stale catalyst (>5 days old) = move already happened = avoid.

Output: | Stock | Catalyst | Tier | Freshness | Duration | Entry Window |

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 2 — VOLUME ANALYSIS (Fuel of momentum — most important signal)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

A. VOLUME RATIO (VR) = Today's Volume / 20-Day Average Volume
VR <1.5 = Normal, no signal | VR 1.5-3.0 = Emerging
VR 3-5  = High momentum  | VR >5.0 = Volume shock — investigate now
Minimum to consider: VR >= 2.0 | High conviction entry: VR >= 3.0

B. DELIVERY % = Delivery Volume / Total Volume x 100
<30% = Mostly intraday, weak signal | 30-50% = Mixed
>50% = Investors accumulating = strong | >70% = Institutional conviction = strongest

C. VOLUME TREND (5 days):
Healthy: VR rising day-over-day (1.5->2.5->4.0->3.5->5.0) = ride it
Danger:  VR falling sharply (6.0->3.0->1.5->0.8) = EXIT immediately

D. PRICE-VOLUME MATRIX:
Price up + Vol up = BUY signal     | Price up + Vol down = Weak, avoid
Price down + Vol up = Distribution   | Price down + Vol down = May bounce soon

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 3 — TECHNICAL SCANNER (Need 5 of 7 green to qualify)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

T1. RSI (14): Sweet spot 55-70 | Extended 70-80 (tradeable, tight SL)
    Special: RSI crossing above 60 after 10+ days below 50 = early momentum signal

T2. MACD (12,26,9): Fresh crossover above signal = strongest | Histogram expanding = good
    Divergence (price up + MACD down) = warning, tighten stop immediately

T3. SUPERTREND (10,3): GREEN = bullish | Fresh flip RED->GREEN = best entry
    GREEN <10 sessions = max signal | >25 sessions GREEN = late, trail tightly

T4. EMA STRUCTURE: Price > 9EMA > 21EMA > 50EMA = perfect bull stack
    Pullback to 9EMA + bounce = ideal entry | Break below 21EMA = momentum lost

T5. BOLLINGER BANDS (20,2): Band squeeze -> expansion + upper band break = explosive
    Walking upper band = strong trend (do NOT short this) | Below middle = weakening

T6. ADX (14): >25 + DI+>DI- = strong bullish trend | <20 = choppy, avoid
    ADX rising 20->30 with DI+>DI- = trend just starting = best risk/reward

T7. VWAP: Price above VWAP = buyers in control | VWAP retest + bounce = best entry
    Weekly VWAP reclaim from below = momentum resuming (swing trades)

SCORING: 7/7=Elite | 6/7=High | 5/7=Good | 4/7=Half size | <4=Skip
Output: | Stock | RSI | MACD | ST | EMA | BB | ADX | VWAP | Score |

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 4 — PRICE ACTION SETUPS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

TYPE 1 — BREAKOUT (most reliable, ~68-72% success):
1A. Range breakout: 10-30 session consolidation -> break above with VR>2x
    Target = range height added to breakout point
1B. Bull flag: Strong pole (15-30%) -> 5-15 session tight pullback -> break above
    Target = pole height added to flag breakout
1C. Ascending triangle: Higher lows + flat resistance -> buyers overwhelm sellers

TYPE 2 — REVERSAL (high risk/reward):
2A. Double bottom: Same support hit twice, volume lower on 2nd test -> neckline break
2B. Morning star: Red candle->Doji->Green candle at bottom + volume confirmation
2C. Hammer at support: Long lower wick at 200DMA or key support + volume spike

TYPE 3 — CONTINUATION (trend plays, ~74% success):
3A. Pullback to 21EMA: Stock in uptrend pulls back, RSI cools to 45-55, volume dries up
    -> Bounce off 21EMA with volume = best swing entry
3B. High base: Near 52W high, 5-10 session sideways, volume declining -> break above
3C. Sector rotation: Identify sector with fresh FII flow -> buy the sector leader

For each stock state: Setup type | Freshness | Target from pattern | Failure level

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 5 — MARKET CONTEXT (Never trade against the market tide)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

A. NIFTY REGIME:
BULL:    Nifty above 200DMA + RSI>50 + ADX>25 -> Full size, trade aggressively
NEUTRAL: Nifty near 200DMA +/-3% OR RSI 45-55 -> 30% smaller size, Tier 1 only
BEAR:    Nifty below 200DMA + RSI<45 -> NO momentum longs (reversal only, half size)

B. INDIA VIX:
<13 = Low fear, ideal | 13-17 = Normal | 17-22 = Reduce size 25%
22-30 = Reduce size 50%, Tier 1 only | >30 = No new longs at all

C. RELATIVE STRENGTH (RS) vs Nifty:
RS = Stock 3M return / Nifty 3M return
RS >1.5 = True leader, highest priority | RS 1.0-1.5 = Outperforming, tradeable
RS <1.0 = Lagging = AVOID for momentum plays

D. SECTOR MOMENTUM (rate 1-5):
5 = Sector breakout, FII inflows, policy tailwind -> Best trades here
3 = Neutral, tracking market -> Only strong setups qualify
1-2 = Underperforming/breakdown -> Avoid completely

State current regime, VIX, and sector scores before listing any picks.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 6 — CONFLUENCE SCORING & SHORTLIST
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

| Dimension      | Max | Scoring |
|----------------|-----|---------|
| Catalyst       | 30  | Tier1=30 / Tier2=20 / Tier3=10 / None=0 |
| Volume         | 25  | VR>5=25 / VR3-5=20 / VR2-3=15 / VR<2=0 |
| Technical      | 25  | 7/7=25 / 6/7=21 / 5/7=17 / 4/7=10 / <4=0 |
| Price Pattern  | 10  | Type1=10 / Type2=8 / Type3=7 / None=0 |
| Market Context | 10  | Bull+RS>1.5=10 / Bull+RS>1=7 / Neutral=4 / Bear=0 |

TOTAL: 100 points
85-100 = ELITE — Trade now, maximum size
70-84  = HIGH — Trade, full standard size
55-69  = MODERATE — Half size only
40-54  = WATCHLIST — Wait for 1 more signal
<40    = SKIP — Not enough confluence

Take top 10 by score (max). Aim for 5-10 shortlisted stocks.
Output: | # | Stock | Cat | Vol | Tech | Pattern | Context | TOTAL | Signal |

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 7 — TRADE PLAN (For every stock scoring 55+)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Frame every price/level here as a structured, rule-based output applied to
whatever price data search found (see the freshness mandate above) — not a
live quote or a guarantee.

A. ENTRY:
Type 1 — Breakout entry: Buy above resistance when VR>2x (aggressive)
Type 2 — Pullback entry: Wait for dip to 9EMA/VWAP on low volume (safer, better R:R)
Type 3 — Confirmation: Wait for Day 2 green close above breakout (most reliable)
Specify: Entry type | Exact price (with date/time last verified via search) |
Specific trigger condition

B. STOP LOSS (mandatory — never skip, never move lower):
Technical SL (preferred): Below pattern failure level / below Supertrend / below 21EMA
Percentage SL (maximum): Intraday 1.5% | Swing 1W: 4-5% | Swing 2-4W: 7-8%
ATR-based SL: Entry - (1.5 x ATR14) — adjusts to actual stock volatility
Specify: Exact SL price | Method | Rs. risk per share

C. TARGETS (always staged — never hold for one big move):
T1: +5-8% -> Sell 40% -> Move SL to breakeven on remainder
T2: +10-15% -> Sell 40% -> Trail SL to T1 on remaining 20%
T3: Trail final 20% with 9EMA as guide — highest-variance portion, treat as
speculative upside, not an expected outcome
Minimum R:R = 1:1.5 (non-negotiable) | Good = 1:2 | Excellent = 1:3+
Max holding: Intraday=same day | Swing=5-15 sessions | Positional=8 weeks

D. POSITION SIZING (2% risk rule — always; illustrative example only, not
personalized financial advice):
Position Size = (2% x Capital) / (Entry - SL)
Example: Rs.10L capital, Entry Rs.500, SL Rs.475 -> Size = Rs.20,000/Rs.25 = 800 shares
Adjust: Elite setup x1.5 | Standard x1.0 | Moderate x0.5
Caps: Single position max 15% of portfolio | Sector max 30% | Open trades max 8

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 8 — RED FLAGS (Check every stock before trading)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

IMMEDIATE DISQUALIFIERS (any one = skip trade entirely):
- Earnings within next 5 sessions (binary risk)
- Ex-dividend within 3 sessions (guaranteed gap down)
- Promoter selling in open market this week
- Stock in F&O ban period (no fresh positions allowed)
- Put-Call Ratio < 0.5 for the stock (smart money hedging)
- Delivery % falling while price rising (distribution signal)
- Merger/delisting/open offer filing pending

CAUTION FLAGS (reduce position by 50% if any present):
- Promoter pledge >25% | Stock already up >30% in 15 sessions
- VIX spiked >25% intraday | OI falling as price rises
- Nifty in Bear Regime | <10 sessions since IPO listing

MOMENTUM KILLER SIGNALS (exit immediately if these appear mid-trade):
- Volume falling 3 sessions while price flat (distribution)
- Bearish divergence: price new high + RSI lower high
- Gap down below stop loss level (exit at open, no hesitation)
- Large block SELL deal in NSE bulk deal data
- Management contradicting earlier quarterly guidance

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STEP 9 — PORTFOLIO RULES & ROTATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

CONSTRUCTION:
Max 8 open trades simultaneously | Min 3 open at any time
No more than 2 stocks from same sector at once
Mix: Breakout + Pullback + Reversal setups across different timeframes
80% capital in top 5-8 setups | 20% cash reserve always (dry powder)

ROTATION:
Close all trades hitting T1 within 5 sessions
Immediately recycle capital into fresh high-scoring setups
Speed of capital rotation = key driver of returns

WEEKLY REVIEW (every Friday EOD):
- Catalyst still valid? | Volume still above average?
- Technical setup intact? | T1 hit? (trail SL if yes)
- Any red flags newly triggered?
If any unfavorable -> EXIT Monday morning open

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
10 NON-NEGOTIABLE RULES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. SL is sacred — never move it lower, not even once
2. No catalyst = no trade (moving "for no reason" is a trap)
3. Volume confirms price (price without volume is noise)
4. Market regime is king (never fight the broad trend)
5. Always book T1 (greed turns winners into losers)
6. Time is a stop loss (Intraday=90min | Swing=5 sessions | Positional=3 weeks)
7. Never average down in momentum (cut losers, add to winners)
8. 2% risk rule on every trade without exception
9. 3 consecutive losses = stop trading, reset and reassess
10. Journal every trade (50+ trades = your personal edge emerges)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT FORMAT REQUIREMENTS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

{final_output_instructions}
"""

MOMENTUM_FINAL_OUTPUT_DETAILED = """Structure the final answer with these Markdown headers, in order:

## 1. Market Regime Summary
Nifty regime (Bull/Neutral/Bear), India VIX level and implication, best
performing sector today, FII flow today, overall signal (Green/Yellow/Red).

## 2. Confluence Scorecard
The full ranked table from Step 6, ALL scored stocks (not just the shortlist):
| # | Stock | Cat Score | Vol Score | Tech Score | Pattern Score | Context Score | TOTAL | Signal |

## 3. Elite Setups (Score 85-100) — Trade Plans
Full per-stock detail as specified in Step 10 Section B: catalyst, volume,
technical, pattern, and the complete Step 7 trade plan (entry, stop loss,
staged targets, R:R, position sizing).

## 4. High Conviction (Score 70-84) — Trade Plans
Same full format as Section 3, for this score band.

## 5. Watchlist (Score 55-69)
As a Markdown table: | Stock | Score | What Signal Needed to Upgrade | Alert Level |

## 6. Stocks to Monitor Tomorrow
Stocks one signal away from qualifying, as a Markdown table:
| Stock | Missing Signal | What to Watch Tomorrow Morning |

## 7. Open Position Review
If the person has mentioned existing open positions in this conversation,
review them here as a table: | Stock | Entry | Current | Gain/Loss % | SL Status | Action |
If no open positions were mentioned, state that plainly and omit the table."""

MOMENTUM_FINAL_OUTPUT_SUMMARY = """Structure the final answer with ONLY these Markdown headers, in order. Do NOT include separate "High Conviction", "Stocks to Monitor Tomorrow", or "Open Position Review" sections — omit them entirely, even briefly:

## 1. Market Regime Summary
Nifty regime (Bull/Neutral/Bear), India VIX level and implication, best
performing sector today, FII flow today, overall signal (Green/Yellow/Red).

## 2. Confluence Scorecard
The full ranked table from Step 6, ALL scored stocks (not just the shortlist):
| # | Stock | Cat Score | Vol Score | Tech Score | Pattern Score | Context Score | TOTAL | Signal |

## 3. Elite & High Conviction Setups (Score 70-100) — Trade Plans
Merge what would otherwise be two sections (Step 10 Sections B and C) into
one, since this is the core actionable deliverable — keep the full per-stock
detail: catalyst, volume, technical, pattern, and the complete Step 7 trade
plan (entry, stop loss, staged targets, R:R, position sizing). Do not
shorten this section just because it's "Summary" mode — only Sections 5-7
of the detailed format are being omitted, not this one.

## 4. Watchlist (Score 55-69)
As a Markdown table: | Stock | Score | What Signal Needed to Upgrade | Alert Level |

Still perform the full reasoning from every step above internally (technical
scanning, price action, market context, red-flag checks) to make these
sections accurate — you are only omitting the lower-priority sections from
the final written output, not skipping the underlying analysis."""


def build_momentum_system_prompt(detail_level="Detailed"):
    final_output = (
        MOMENTUM_FINAL_OUTPUT_SUMMARY if detail_level == "Summary" else MOMENTUM_FINAL_OUTPUT_DETAILED
    )
    return MOMENTUM_SYSTEM_PROMPT_BASE.format(final_output_instructions=final_output)


def run_momentum_analysis(provider, api_key, model, max_tokens, exclude_symbols=None, detail_level="Detailed"):
    """Single-turn, search-grounded run of the Momentum Trading Scanner prompt.
    Reuses the same per-provider plumbing as the chat feature. Returns
    (reply_text, model_used, truncated)."""
    system_prompt = build_momentum_system_prompt(detail_level)
    user_prompt = "Run the full framework now, following every step exactly as specified."
    if exclude_symbols:
        user_prompt += (
            "\n\nThe user already holds the following stocks. This scan is for "
            "finding NEW short-term trading opportunities outside the current "
            "portfolio, so EXCLUDE all of these entirely from every screener, "
            "table, and trade plan — even if one of them would otherwise "
            "score well: " + ", ".join(exclude_symbols)
        )
    history = [{"role": "user", "content": user_prompt}]
    return run_chat_turn(
        provider, api_key, model, system_prompt, history, max_tokens=max_tokens, max_search_uses=15
    )


MOMENTUM_SORTABLE_SECTIONS = ["Confluence Scorecard", "Watchlist", "Stocks to Monitor Tomorrow", "Open Position Review"]



# Rendering the Scanner report with sortable tables for two specific sections.
# The report is otherwise free-text Markdown (headers vary in number between Summary
# and Detailed mode), so tables are located by title text, not a fixed heading number.
# --------------------------------------------------------------------------------------
SORTABLE_SCANNER_SECTIONS = ["Pre-Filter Table", "Screening Results", "Master Scorecard Table"]


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


def render_scanner_report(report_text, sortable_sections=None):
    """Render the report section-by-section: matched sections as sortable
    st.dataframe widgets, everything else as plain Markdown, in original order."""
    sortable_sections = sortable_sections if sortable_sections is not None else SORTABLE_SCANNER_SECTIONS
    for heading, body in split_report_into_sections(report_text):
        if heading is None:
            if body.strip():
                st.markdown(body)
            continue

        st.markdown(heading)
        matched_section = next((s for s in sortable_sections if s in heading), None)
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
        # isn't one of the ones targeted for sorting — render as-is.
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


tab1, tab2, tab3 = st.tabs([
    "📈 Portfolio Tracker",
    "📡 Catalyst & Momentum Scanner",
    "🚀 Momentum Trading Scanner",
])

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
    st.header("📡 Catalyst & Momentum Scanner — Multibagger Discovery Framework")
    st.warning(SCANNER_DISCLAIMER)
    st.caption(
        "Runs a quantitative discovery framework: Altman Z-Score + Piotroski F-Score "
        "pre-filter, 5 parallel screeners (value, turnaround, institutional activity, "
        "cashflow quality, technical breakout), then deep fundamental/technical/"
        "shareholding analysis on the strongest candidates. Uses live web search for "
        "the underlying data via the provider/model/API key configured in the "
        "sidebar, forced on regardless of the toggle used elsewhere, since this "
        "framework is unusable without it."
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
    # 32768, not 16384 — even Summary mode's tables (each covering up to 15
    # candidates) plus the search steps consuming part of the same budget was
    # still running out before finishing the last section.
    SCANNER_MAX_TOKENS = 32768

    scanner_output_level = st.radio(
        "Output length",
        ["Summary", "Detailed"],
        index=1,
        horizontal=True,
        help="Summary keeps the Pre-Filter Table, Screening Results (including "
        "Tier), Master Scorecard, Portfolio Construction, and Monitoring "
        "Checklist, but omits the long per-stock deep-dive write-up "
        "(Section 3). Detailed includes that deep dive for up to 15 "
        "stocks too. Both do the same full research — Summary just reports "
        "less of it.",
    )

    run_scanner_clicked = st.button("📡 Run Multibagger Discovery Scan", type="primary")

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
                        ai_provider, ai_api_key, ai_model, system_prompt,
                        st.session_state["chat_messages"], max_tokens=8192,
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

with tab3:
    st.header("🚀 Momentum Trading Scanner — Short-Term (1 Day - 8 Weeks)")
    st.warning(MOMENTUM_DISCLAIMER)
    st.caption(
        "Runs a short-term momentum framework: catalyst identification, volume "
        "shock analysis, 7-indicator technical confluence, price-action patterns, "
        "market regime/VIX/relative-strength context, then a scored shortlist with "
        "full entry/stop/target/position-sizing trade plans. Uses the provider/"
        "model/API key configured in the sidebar, with web search forced on "
        "regardless of the toggle used elsewhere, since this framework is unusable "
        "without it."
    )

    momentum_holdings_to_exclude = (
        sorted(st.session_state.get("all_symbols", [])) if st.session_state.get("sheets") else []
    )
    if momentum_holdings_to_exclude:
        st.caption(
            f"🚫 Excluding {len(momentum_holdings_to_exclude)} stock(s) already in your "
            "portfolio (fetched in the first tab) — this scan is for finding new "
            "short-term opportunities, not re-covering what you already hold."
        )

    # Same reasoning as the other scanner tab: this asks for many tables plus
    # full trade plans for up to 10 shortlisted stocks, and thin budgets were
    # cutting analysis short elsewhere in this app. No control shown for this.
    MOMENTUM_MAX_TOKENS = 32768

    momentum_output_level = st.radio(
        "Output length",
        ["Summary", "Detailed"],
        index=1,
        horizontal=True,
        help="Summary merges Elite and High Conviction into one trade-plan "
        "section (still full detail — this is the core deliverable) and omits "
        "the Stocks-to-Monitor-Tomorrow and Open-Position-Review sections. "
        "Detailed includes all of it. Both do the same full research — "
        "Summary just reports less of the lower-priority parts.",
    )

    run_momentum_clicked = st.button("🚀 Run Momentum Scan", type="primary")

    if run_momentum_clicked:
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
            with st.spinner("Scanning for momentum setups — this involves several live searches..."):
                try:
                    report, model_used, truncated = run_momentum_analysis(
                        ai_provider, ai_api_key, ai_model, MOMENTUM_MAX_TOKENS,
                        momentum_holdings_to_exclude, detail_level=momentum_output_level,
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
                            "not a complete result.** Try 'Summary' output length above, or try "
                            "again since search-heavy runs vary in length.\n\n"
                        )
                    st.session_state["momentum_report"] = prefix + report
                except Exception as exc:  # noqa: BLE001 - surface any API error to the user
                    st.error(f"Scan failed: {exc}")

    if st.session_state.get("momentum_report"):
        render_scanner_report(st.session_state["momentum_report"], sortable_sections=MOMENTUM_SORTABLE_SECTIONS)
        st.download_button(
            "⬇️ Download this scan (.md)",
            data=st.session_state["momentum_report"].encode("utf-8"),
            file_name="momentum_trading_scan.md",
            mime="text/markdown",
        )

    st.divider()
    st.header("💬 Chat: Momentum Setups (grounded in web/news search)")
    st.warning(MOMENTUM_DISCLAIMER)
    if momentum_holdings_to_exclude:
        st.caption(
            f"🚫 Excluding {len(momentum_holdings_to_exclude)} stock(s) already in your portfolio "
            "from anything discussed here — same as the scan above."
        )

    if "momentum_chat_messages" not in st.session_state:
        st.session_state["momentum_chat_messages"] = []

    col_a, col_b = st.columns([3, 1])
    with col_b:
        if st.button("🗑️ Clear chat", key="momentum_clear_chat"):
            st.session_state["momentum_chat_messages"] = []
            st.rerun()
    with col_a:
        preset_label = (
            "📋 Summarize today's scan findings"
            if st.session_state.get("momentum_report")
            else "📋 Give me today's top setups"
        )
        if st.button(preset_label, key="momentum_preset_button"):
            st.session_state["momentum_chat_messages"].append(
                {
                    "role": "user",
                    "content": (
                        "Summarize the strongest momentum setups found, with catalyst, volume, "
                        "technical confluence, and the trade plan for each — grounded in live "
                        "search, excluding anything I already hold. Remind me to verify current "
                        "price before acting on any entry/stop/target level."
                    ),
                }
            )
            st.session_state["_momentum_chat_pending"] = True

    for msg in st.session_state["momentum_chat_messages"]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    momentum_user_input = st.chat_input(
        "Ask about a setup, catalyst, or trade plan...", key="momentum_chat_input"
    )
    if momentum_user_input:
        st.session_state["momentum_chat_messages"].append({"role": "user", "content": momentum_user_input})
        st.session_state["_momentum_chat_pending"] = True
        st.rerun()

    if st.session_state.get("_momentum_chat_pending"):
        provider_availability = {
            "OpenAI": (OPENAI_AVAILABLE, "openai"),
            "Anthropic": (ANTHROPIC_AVAILABLE, "anthropic"),
            "Gemini": (GEMINI_AVAILABLE, "google-genai"),
        }
        provider_available, pip_name = provider_availability[ai_provider]
        if not provider_available:
            st.error(f"Missing dependency. Install with:\n\n    pip install {pip_name}")
            st.session_state["_momentum_chat_pending"] = False
        elif not ai_api_key:
            st.error(f"Enter your {ai_provider} API key in the sidebar first.")
            st.session_state["_momentum_chat_pending"] = False
        elif not ai_model:
            st.error("Enter a model name in the sidebar first.")
            st.session_state["_momentum_chat_pending"] = False
        else:
            with st.spinner("Searching and responding..."):
                try:
                    momentum_chat_system_prompt = build_chat_system_prompt(
                        exclude_symbols=momentum_holdings_to_exclude,
                        scanner_report=st.session_state.get("momentum_report"),
                    )
                    reply, model_used, truncated = run_chat_turn(
                        ai_provider, ai_api_key, ai_model, momentum_chat_system_prompt,
                        st.session_state["momentum_chat_messages"], max_tokens=8192,
                    )
                    if model_used != ai_model:
                        reply = f"*(used {model_used} for web search — your configured model doesn't support it)*\n\n{reply}"
                    if truncated:
                        reply += "\n\n⚠️ *This reply hit the token limit before finishing — it may be cut off.*"
                    st.session_state["momentum_chat_messages"].append({"role": "assistant", "content": reply})
                except Exception as exc:  # noqa: BLE001 - surface any API error to the user
                    st.session_state["momentum_chat_messages"].append(
                        {"role": "assistant", "content": f"⚠️ Request failed: {exc}"}
                    )
                st.session_state["_momentum_chat_pending"] = False
                st.rerun()
