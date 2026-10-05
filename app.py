"""IBKR trade analysis: parse exports, match fills with FIFO, and review behavior."""

from __future__ import annotations

import csv
import hashlib
import io
import math
import re
import sys
from collections import defaultdict, deque
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from ibkr_service import DEFAULT_QUERY_ID, FlexServiceError, date_windows, fetch_ibkr_trades, hosted_publicly, year_windows

EPS = 1e-8
REVENGE_WINDOW = timedelta(minutes=10)
FILL_GROUP_SECONDS = 5.0
OUTSIZED_LOSS_MULTIPLE = 2.0
FEE_DRAG_RATIO = 0.20
DEFAULT_CAD_PER_USD = 1.38
SUPPORTED_CURRENCIES = {"USD", "CAD"}
FLEX_DIR = Path(__file__).resolve().parent / "data" / "flex"
WEEKDAY_ORDER = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
OUTLIER_SYMBOLS = frozenset({"AMD", "SOXL"})
LEVERAGED_ETFS = frozenset(
    {
        "SOXL",
        "SOXS",
        "TQQQ",
        "SQQQ",
        "UPRO",
        "SPXU",
        "MSTY",
        "TSLY",
        "CONY",
        "AMZY",
        "APLY",
        "UVXY",
    }
)
ASSET_TYPE_OPTIONS = ["Single Stocks", "3x Leveraged ETFs", "Options"]
HOLD_BUCKETS = [
    ("<1 hour", 0.0, 3600.0),
    ("1–24 hours", 3600.0, 86400.0),
    ("1–7 days", 86400.0, 7 * 86400.0),
    (">7 days", 7 * 86400.0, None),
]

GREEN = "#1B8A5A"
RED = "#D64545"
BLUE = "#1F4B99"
SLATE = "#5C6770"

TRADE_SECTIONS = {
    "trades",
    "trade",
    "executions",
    "execution",
    "tradeconfirmations",
    "tradeconfirmation",
    "transactions",
}
KEEP_DETAIL = {"order", "trade", "execution", "exchtrade", "fill"}
SKIP_DETAIL = {
    "closedlot",
    "subtotal",
    "total",
    "summary",
    "symbolsummary",
    "assetcategorysummary",
    "ordersummary",
    "composite",
    "symbol",
    "header",
}

EXEC_COLUMNS = [
    "symbol",
    "trade_time",
    "action",
    "quantity",
    "price",
    "proceeds",
    "commission",
    "multiplier",
    "currency",
    "asset_category",
    "source_row",
    "put_call",
    "strike",
    "expiry",
    "open_close",
    "order_type",
    "order_time",
    "notes_codes",
    "latency_seconds",
]


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------


def _key(value: object) -> str:
    return "".join(ch for ch in str(value).lower() if ch.isalnum())


def _section_key(value: object) -> str:
    return "".join(ch for ch in str(value).lower() if ch.isalpha())


def find_col(columns, aliases: list[str]) -> str | None:
    keyed = {_key(col): col for col in columns}
    for alias in aliases:
        match = keyed.get(_key(alias))
        if match is not None:
            return match
    return None


def find_cols(columns, aliases: set[str]) -> list[str]:
    wanted = {_key(alias) for alias in aliases}
    return [col for col in columns if _key(col) in wanted]


def parse_number(value, decimal_comma: bool = False) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = str(value).strip()
    if text == "" or text.lower() in {"nan", "none", "--", "n/a", "null"}:
        return None
    text = text.replace("$", "").replace(" ", "")
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    if decimal_comma:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", "")
    try:
        number = float(text)
    except ValueError:
        return None
    if math.isnan(number) or math.isinf(number):
        return None
    return number


def parse_datetime(value) -> pd.Timestamp:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return pd.NaT
    if isinstance(value, pd.Timestamp):
        stamp = value
    elif isinstance(value, datetime):
        stamp = pd.Timestamp(value)
    else:
        text = str(value).strip().strip('"')
        if text == "" or text.lower() in {"nan", "none", "nat"}:
            return pd.NaT
        text = text.replace(",", " ").replace(";", " ")
        text = " ".join(text.split())
        compact = None
        digits = text.replace(" ", "")
        if len(digits) == 8 and digits.isdigit():
            compact = pd.to_datetime(digits, format="%Y%m%d", errors="coerce")
        elif len(text) >= 15 and text[:8].isdigit() and text[9:15].isdigit() and text[8] == " ":
            compact = pd.to_datetime(text[:8] + text[9:15], format="%Y%m%d%H%M%S", errors="coerce")
        if compact is not None and not pd.isna(compact):
            stamp = compact
        else:
            try:
                stamp = pd.to_datetime(text, errors="coerce", format="mixed")
            except (TypeError, ValueError):
                stamp = pd.to_datetime(text, errors="coerce")
    if pd.isna(stamp):
        return pd.NaT
    if stamp.tzinfo is not None:
        stamp = stamp.tz_convert("UTC").tz_localize(None)
    return stamp


def detect_delimiter(text: str) -> str:
    sample = "\n".join(text.splitlines()[:40])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        return dialect.delimiter
    except csv.Error:
        counts = {",": sample.count(","), ";": sample.count(";"), "\t": sample.count("\t")}
        return max(counts, key=counts.get)


def detect_decimal_comma(text: str, delimiter: str) -> bool:
    if delimiter != ";":
        return False
    sample = text[:8000]
    comma_decimals = sample.count(",")
    dot_decimals = sample.count(".")
    return comma_decimals > dot_decimals


def decode_upload(name: str, data: bytes) -> str:
    lower = name.lower()
    if lower.endswith(".xls") and not lower.endswith(".xlsx"):
        raise ValueError(
            "Legacy .xls workbooks are not supported. Save the report as .xlsx or .csv and upload it again."
        )
    if lower.endswith((".xlsx", ".xlsm")):
        frame = pd.read_excel(io.BytesIO(data), header=None, dtype=str, engine="openpyxl")
        return frame.to_csv(index=False, header=False)
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            return data.decode(encoding).replace("\x00", "")
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace").replace("\x00", "")


def _is_sectioned(rows: list[list[str]]) -> bool:
    header_hits = 0
    for row in rows[:80]:
        if len(row) >= 2 and row[1].strip().lower() == "header":
            header_hits += 1
    return header_hits > 0


def parse_section_rows(text: str, delimiter: str) -> list[dict]:
    """Pull execution rows out of an Interactive Brokers multi-section CSV."""
    rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
    if not _is_sectioned(rows):
        return []

    section = ""
    headers: list[str] = []
    extracted: list[dict] = []
    for row in rows:
        if len(row) < 2:
            continue
        kind = row[1].strip().lower()
        if kind == "header":
            section = _section_key(row[0])
            headers = [cell.strip() for cell in row[2:]]
            continue
        if kind != "data" or section not in TRADE_SECTIONS or not headers:
            continue
        values = row[2:]
        if len(values) < len(headers):
            values = values + [""] * (len(headers) - len(values))
        record = dict(zip(headers, values))
        record["_section"] = section
        extracted.append(record)
    return extracted


HEADER_KEYS = {
    "symbol",
    "datetime",
    "tradedatetime",
    "buysell",
    "quantity",
    "qty",
    "tradeprice",
    "tprice",
    "price",
    "proceeds",
    "ibcommission",
    "commission",
    "currency",
    "assetclass",
    "assetcategory",
    "multiplier",
    "fxratetobase",
}


def _read_rows(text: str, delimiter: str) -> list[list[str]]:
    return list(csv.reader(io.StringIO(text), delimiter=delimiter))


def _header_score(row: list[str]) -> int:
    keys = {_key(cell) for cell in row if str(cell).strip()}
    if "symbol" not in keys:
        return -1
    return len(keys & HEADER_KEYS)


def _choose_delimiter(text: str) -> str:
    sniffed = detect_delimiter(text)
    candidates = []
    for delimiter in (sniffed, ",", ";", "\t"):
        if delimiter not in candidates:
            candidates.append(delimiter)
    best = sniffed
    best_score = -1
    for delimiter in candidates:
        sample = _read_rows(text, delimiter)[:200]
        if _is_sectioned(sample):
            return delimiter
        score = max((_header_score(row) for row in sample), default=-1)
        if score > best_score:
            best_score = score
            best = delimiter
    return best


def read_flat_frame(text: str, delimiter: str) -> pd.DataFrame:
    """Read a CSV even when rows have more columns than earlier lines.

    Extra fields are ignored. The header is the row that names Symbol and the
    other trade columns, which may sit below a short preamble.
    """
    rows = _read_rows(text, delimiter)
    if not rows:
        return pd.DataFrame()

    best_index = None
    best_score = -1
    best_width = -1
    for index, row in enumerate(rows[:200]):
        score = _header_score(row)
        if score > best_score or (score == best_score and score >= 0 and len(row) > best_width):
            best_score = score
            best_width = len(row)
            best_index = index
    if best_index is None or best_score < 0:
        raise ValueError("Could not find a header row with a Symbol column.")

    header = [cell.strip() for cell in rows[best_index]]
    while header and not header[-1]:
        header.pop()
    width = len(header)
    header_keys = [_key(cell) for cell in header]
    records = []
    for row in rows[best_index + 1 :]:
        if not any(str(cell).strip() for cell in row):
            continue
        values = [cell.strip() for cell in row[:width]]
        if len(values) < width:
            values.extend([""] * (width - len(values)))
        if [_key(cell) for cell in values] == header_keys:
            continue
        records.append(dict(zip(header, values)))
    if not records:
        return pd.DataFrame(columns=header)
    return pd.DataFrame.from_records(records)


def _fee_cost(series: pd.Series, decimal_comma: bool) -> pd.Series:
    """Return fee amounts where a positive number is a cost and a negative number is a credit.

    IBKR usually stores commissions as negative cash. When the column is mostly negative,
    the sign is flipped. Exports that already store fees as positive numbers are left as-is.
    """
    numbers = pd.to_numeric(series.map(lambda value: parse_number(value, decimal_comma)), errors="coerce")
    nonzero = numbers.dropna()
    nonzero = nonzero[nonzero != 0]
    if nonzero.empty:
        return numbers.fillna(0.0)
    if float(nonzero.median()) < 0:
        return (-numbers).fillna(0.0)
    return numbers.fillna(0.0)


def _cell_text(value) -> str:
    if value is None:
        return ""
    try:
        if bool(pd.isna(value)):
            return ""
    except TypeError:
        pass
    text = str(value).strip()
    if text.lower() in {"", "nan", "none"}:
        return ""
    return text


def _is_summary_symbol(symbol: str) -> bool:
    if symbol in {"", "NAN", "NONE", "--", "TOTAL", "SUBTOTAL"}:
        return True
    return symbol.startswith("TOTAL ") or symbol.startswith("SUBTOTAL")


def _infer_multiplier(quantity: float, price: float, proceeds: float | None) -> float:
    if not quantity or not price or proceeds is None:
        return 1.0
    ratio = abs(proceeds) / (abs(quantity) * abs(price))
    if not math.isfinite(ratio) or ratio <= 0:
        return 1.0
    for candidate in (1.0, 100.0):
        if abs(ratio - candidate) / candidate <= 0.15:
            return candidate
    if 1.0 < ratio < 100_000:
        rounded = float(round(ratio))
        if rounded > 0 and abs(ratio - rounded) / rounded <= 0.05:
            return rounded
    return 1.0


def _normalize_action(value, quantity: float | None) -> str | None:
    if value is not None and str(value).strip() and str(value).strip().lower() not in {"nan", "none"}:
        token = _key(value)
        if token.startswith("buy") or token in {"bot", "b", "bought", "cover"}:
            return "BUY"
        if token.startswith("sell") or token in {"sld", "s", "sold", "short"}:
            return "SELL"
    if quantity is None:
        return None
    if quantity < 0:
        return "SELL"
    if quantity > 0:
        return "BUY"
    return None


def _filter_detail_rows(frame: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    column = find_col(frame.columns, ["DataDiscriminator", "LevelOfDetail", "Level Of Detail"])
    if column is None:
        return frame, []
    labels = frame[column].fillna("").map(_key)
    keep_mask = labels.isin(KEEP_DETAIL)
    if bool(keep_mask.any()):
        notes = []
        chosen = next((level for level in ("execution", "order", "trade", "exchtrade", "fill") if labels.isin({level}).any()), None)
        if chosen is None:
            kept = frame.loc[keep_mask].copy()
        else:
            kept = frame.loc[labels.isin({chosen})].copy()
            extra = sorted({label for label in labels[keep_mask].unique() if label and label != chosen})
            if extra:
                notes.append(
                    "Kept "
                    + chosen
                    + " rows and skipped "
                    + ", ".join(extra)
                    + " rows so the same fill is not counted twice."
                )
        if labels.isin({"closedlot"}).any():
            notes.append(
                "Closed-lot summary rows were ignored so realized PnL is calculated from executions only."
            )
        return kept, notes

    skip_mask = labels.isin(SKIP_DETAIL) | labels.str.contains("summary", na=False) | labels.str.contains("total", na=False)
    notes = []
    if labels.isin({"closedlot"}).any():
        notes.append(
            "This file contains closed-lot rows but no order or execution rows. "
            "Export executions (Flex level of detail: Execution or Order). Closed lots were skipped."
        )
    return frame.loc[~skip_mask].copy(), notes


def normalize_executions(frame: pd.DataFrame, decimal_comma: bool = False) -> tuple[pd.DataFrame, list[str]]:
    notes: list[str] = []
    if frame.empty:
        return pd.DataFrame(columns=EXEC_COLUMNS), ["The file has no data rows."]

    frame = frame.copy()
    frame.columns = [str(col).strip() for col in frame.columns]
    filtered, detail_notes = _filter_detail_rows(frame)
    notes.extend(detail_notes)
    frame = filtered
    if frame.empty:
        return pd.DataFrame(columns=EXEC_COLUMNS), notes or ["No execution rows were found."]

    symbol_col = find_col(frame.columns, ["Symbol", "Ticker"])
    when_col = find_col(frame.columns, ["Date/Time", "DateTime", "Trade Date/Time", "TradeDateTime", "Execution Time"])
    date_col = None if when_col else find_col(frame.columns, ["Trade Date", "TradeDate", "Execution Date"])
    time_col = None if when_col else find_col(frame.columns, ["Trade Time", "TradeTime", "Time"])
    action_col = find_col(frame.columns, ["Buy/Sell", "BuySell", "Action", "Side"])
    qty_col = find_col(frame.columns, ["Quantity", "Qty", "Shares", "Trade Quantity", "Filled Quantity"])
    price_col = find_col(frame.columns, ["Trade Price", "TradePrice", "T. Price", "Price", "Exec Price", "Fill Price"])
    proceeds_col = find_col(frame.columns, ["Proceeds"])
    currency_col = find_col(frame.columns, ["Currency", "CurrencyPrimary"])
    asset_col = find_col(frame.columns, ["Asset Class", "AssetClass", "Asset Category", "AssetCategory"])
    multiplier_col = find_col(frame.columns, ["Multiplier"])
    fx_col = find_col(frame.columns, ["FX Rate To Base", "FXRateToBase"])
    comm_ccy_col = find_col(frame.columns, ["IB Commission Currency", "IBCommissionCurrency"])
    put_call_col = find_col(frame.columns, ["Put/Call", "PutCall"])
    strike_col = find_col(frame.columns, ["Strike"])
    expiry_col = find_col(frame.columns, ["Expiry", "Expiration"])
    open_close_col = find_col(frame.columns, ["Open/Close Indicator", "Open/Close", "OpenCloseIndicator"])
    order_type_col = find_col(frame.columns, ["Order Type", "OrderType"])
    order_time_col = find_col(frame.columns, ["Order Time", "OrderTime"])
    notes_col = find_col(frame.columns, ["Notes/Codes", "Notes", "Codes"])
    ib_commission_col = find_col(frame.columns, ["IB Commission", "IBCommission"])
    if ib_commission_col:
        fee_cols = [ib_commission_col]
    else:
        fee_cols = find_cols(
            frame.columns,
            {"Comm/Fee", "Comm/Fees", "Commission", "Commissions", "Fees", "Fee", "Brokerage"},
        )

    missing = []
    if symbol_col is None:
        missing.append("Symbol")
    if qty_col is None:
        missing.append("Quantity")
    if price_col is None:
        missing.append("Price")
        if when_col is None and date_col is None:
            missing.append("Date/Time")
    if missing:
        preview = ", ".join(str(col) for col in list(frame.columns)[:12])
        notes.append(
            "Could not find required columns: "
            + ", ".join(missing)
            + f". Columns seen: {preview}."
        )
        return pd.DataFrame(columns=EXEC_COLUMNS), notes

    frame = frame.reset_index(drop=True)
    fee_cost = pd.Series(0.0, index=frame.index)
    for column in fee_cols:
        fee_cost = fee_cost + _fee_cost(frame[column], decimal_comma)

    records = []
    skipped = 0
    for index, row in enumerate(frame.to_dict(orient="records")):
        symbol = str(row.get(symbol_col, "")).strip().upper()
        if _is_summary_symbol(symbol):
            skipped += 1
            continue
        quantity = parse_number(row.get(qty_col), decimal_comma)
        price = parse_number(row.get(price_col), decimal_comma)
        action = _normalize_action(row.get(action_col) if action_col else None, quantity)
        if when_col is not None:
            trade_time = parse_datetime(row.get(when_col))
        else:
            date_part = "" if date_col is None else str(row.get(date_col, "")).strip()
            time_part = "" if time_col is None else str(row.get(time_col, "")).strip()
            trade_time = parse_datetime(f"{date_part} {time_part}".strip())
        if quantity is None or price is None or price <= 0 or action is None or pd.isna(trade_time):
            skipped += 1
            continue
        if abs(quantity) <= EPS:
            skipped += 1
            continue
        proceeds = parse_number(row.get(proceeds_col), decimal_comma) if proceeds_col else None
        currency = "USD"
        if currency_col:
            currency = str(row.get(currency_col, "")).strip().upper()
            if currency in {"", "NAN", "NONE"}:
                currency = "USD"
        asset = str(row.get(asset_col, "")).strip() if asset_col else ""
        if asset.lower() in {"nan", "none"}:
            asset = ""
        explicit_multiplier = parse_number(row.get(multiplier_col), decimal_comma) if multiplier_col else None
        if explicit_multiplier is None or explicit_multiplier <= 0:
            explicit_multiplier = _infer_multiplier(quantity, price, proceeds)
        comm_currency = _cell_text(row.get(comm_ccy_col)).upper() if comm_ccy_col else ""
        order_time = parse_datetime(row.get(order_time_col)) if order_time_col else pd.NaT
        latency = None
        if not pd.isna(order_time):
            latency = (trade_time - order_time).total_seconds()
        records.append(
            {
                "symbol": symbol,
                "trade_time": trade_time,
                "action": action,
                "quantity": abs(float(quantity)),
                "price": float(price),
                "proceeds": proceeds,
                "commission": float(fee_cost.iloc[index]),
                "multiplier": float(explicit_multiplier),
                "currency": currency,
                "asset_category": asset,
                "source_row": index + 1,
                "put_call": _cell_text(row.get(put_call_col)) if put_call_col else "",
                "strike": _cell_text(row.get(strike_col)) if strike_col else "",
                "expiry": _cell_text(row.get(expiry_col)) if expiry_col else "",
                "open_close": _cell_text(row.get(open_close_col)).upper() if open_close_col else "",
                "order_type": _cell_text(row.get(order_type_col)) if order_type_col else "",
                "order_time": order_time if not pd.isna(order_time) else pd.NaT,
                "notes_codes": _cell_text(row.get(notes_col)) if notes_col else "",
                "latency_seconds": latency,
                "fx_rate": parse_number(row.get(fx_col), decimal_comma) if fx_col else None,
                "comm_currency": comm_currency,
            }
        )

    if skipped:
        notes.append(f"Skipped {skipped} rows that were totals, blanks, or missing a symbol, time, quantity, or price.")
    records = _fold_fx_to_base(records, notes)
    executions = pd.DataFrame.from_records(records, columns=EXEC_COLUMNS)
    if executions.empty:
        notes.append("No usable executions remained after cleaning.")
        return executions, notes

    executions["trade_time"] = pd.to_datetime(executions["trade_time"])
    executions = executions.sort_values(["trade_time", "source_row"], kind="mergesort").reset_index(drop=True)
    currencies = sorted({c for c in executions["currency"].dropna().unique() if c and c != "NAN"})
    if len(currencies) > 1:
        notes.append(
            "Currencies in this file: "
            + ", ".join(currencies)
            + ". USD and CAD amounts are converted with the sidebar rate so every total uses one currency."
        )
    return executions, notes


def load_executions(name: str, data: bytes) -> tuple[pd.DataFrame, list[str]]:
    text = decode_upload(name, data)
    delimiter = _choose_delimiter(text)
    decimal_comma = detect_decimal_comma(text, delimiter)
    section_rows = parse_section_rows(text, delimiter)
    notes: list[str] = []
    if section_rows:
        frame = pd.DataFrame(section_rows)
        executions, notes = normalize_executions(frame, decimal_comma)
        if not executions.empty:
            notes.insert(0, f"Read {len(executions)} executions from the IBKR trades section.")
            return executions, notes
    frame = read_flat_frame(text, delimiter)
    executions, flat_notes = normalize_executions(frame, decimal_comma)
    if not executions.empty:
        flat_notes.insert(0, f"Read {len(executions)} executions from the uploaded table.")
    return executions, notes + flat_notes


# ---------------------------------------------------------------------------
# Fill aggregation, FX, and FIFO matching
# ---------------------------------------------------------------------------


def _median(values: list[float]) -> float:
    ordered = sorted(values)
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) / 2.0


def _detect_base_currency(records: list[dict]) -> str | None:
    near_one: list[str] = []
    rates_by_currency: dict[str, list[float]] = defaultdict(list)
    for record in records:
        rate = record.get("fx_rate")
        if rate is None:
            continue
        currency = str(record.get("currency") or "").upper()
        if not currency:
            continue
        rates_by_currency[currency].append(float(rate))
        if abs(float(rate) - 1.0) <= 0.005:
            near_one.append(currency)
    if near_one:
        return max(set(near_one), key=near_one.count)
    cad_rates = rates_by_currency.get("CAD", [])
    usd_rates = rates_by_currency.get("USD", [])
    if cad_rates and not usd_rates and _median(cad_rates) < 0.95:
        return "USD"
    if usd_rates and not cad_rates and _median(usd_rates) > 1.05:
        return "CAD"
    return None


def _fold_fx_to_base(records: list[dict], notes: list[str]) -> list[dict]:
    """Convert each fill into the account base currency with IBKR's FX Rate To Base."""
    if not records or all(record.get("fx_rate") is None for record in records):
        for record in records:
            record.pop("fx_rate", None)
            record.pop("comm_currency", None)
        return records

    base = _detect_base_currency(records)
    if base is None:
        notes.append(
            "FX Rate To Base is in the file, but the account base currency could not be identified. "
            "Amounts stay in the trade currency."
        )
        for record in records:
            record.pop("fx_rate", None)
            record.pop("comm_currency", None)
        return records

    converted_any = False
    mismatched_commission = False
    for record in records:
        rate = record.get("fx_rate")
        rate = 1.0 if rate is None else float(rate)
        trade_currency = str(record.get("currency") or base).upper()
        if abs(rate - 1.0) > 0.0000001 or trade_currency != base:
            converted_any = True
        record["price"] = float(record["price"]) * rate
        if record.get("proceeds") is not None:
            record["proceeds"] = float(record["proceeds"]) * rate
        commission_currency = str(record.get("comm_currency") or "").upper()
        if commission_currency in {"", trade_currency}:
            record["commission"] = float(record["commission"]) * rate
        elif commission_currency != base:
            record["commission"] = float(record["commission"]) * rate
            mismatched_commission = True
        record["currency"] = base
        record.pop("fx_rate", None)
        record.pop("comm_currency", None)
    if converted_any:
        notes.append(f"Converted prices, proceeds, and commissions to {base} with each row's FX Rate To Base.")
    if mismatched_commission:
        notes.append("Some commissions were in a third currency. Those fees used the trade's FX Rate To Base.")
    return records


def _row_text(row, name: str) -> str:
    return _cell_text(getattr(row, name, ""))


def _contract_label(put_call: str, strike: str, expiry: str) -> str:
    parts = []
    if expiry:
        parts.append(expiry[:10])
    side = put_call[:1].upper() if put_call else ""
    if side == "C":
        parts.append("Call")
    elif side == "P":
        parts.append("Put")
    elif put_call:
        parts.append(put_call)
    if strike:
        parts.append(strike)
    return " ".join(parts)


def _note_flags(text: str) -> list[str]:
    tokens = {token.strip().upper() for token in re.split(r"[;,|\s]+", text) if token.strip()}
    flags = []
    if "W" in tokens:
        flags.append("Wash sale")
    if "A" in tokens:
        flags.append("Assignment")
    if "EX" in tokens or "EXE" in tokens:
        flags.append("Exercise")
    return flags


def _missing_number(value) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except TypeError:
        return False


def aggregate_fills(executions: pd.DataFrame) -> pd.DataFrame:
    """Combine partial fills that share a ticker, side, and a burst of timestamps.

    A fill joins the current group when it has the same symbol, buy/sell side, and
    currency, and it prints within 5 seconds of the previous fill in that group.
    Price is the quantity-weighted average. System order IDs are not used.
    """
    if executions is None or executions.empty:
        return executions.iloc[0:0].copy() if executions is not None else pd.DataFrame(columns=EXEC_COLUMNS)

    ordered = executions.sort_values(["symbol", "action", "currency", "trade_time", "source_row"], kind="mergesort")
    groups: list[list] = []
    current: list = []
    for row in ordered.itertuples(index=False):
        if not current:
            current = [row]
            continue
        previous = current[-1]
        gap = (pd.Timestamp(row.trade_time) - pd.Timestamp(previous.trade_time)).total_seconds()
        same_order = (
            row.symbol == previous.symbol
            and row.action == previous.action
            and str(row.currency or "") == str(previous.currency or "")
            and gap <= FILL_GROUP_SECONDS
        )
        if same_order:
            current.append(row)
        else:
            groups.append(current)
            current = [row]
    if current:
        groups.append(current)

    records = []
    for group in groups:
        quantity = float(sum(row.quantity for row in group))
        if quantity <= EPS:
            continue
        price = float(sum(row.price * row.quantity for row in group) / quantity)
        commission = float(sum(row.commission for row in group))
        proceeds = None
        if all(not _missing_number(row.proceeds) for row in group):
            proceeds = float(sum(row.proceeds for row in group))
        first = min(group, key=lambda row: (pd.Timestamp(row.trade_time), row.source_row))
        asset = next((_row_text(row, "asset_category") for row in group if _row_text(row, "asset_category")), "")
        notes = []
        for row in group:
            note = _row_text(row, "notes_codes")
            if note and note not in notes:
                notes.append(note)
        latencies = [float(row.latency_seconds) for row in group if not _missing_number(getattr(row, "latency_seconds", None))]
        order_times = [pd.Timestamp(row.order_time) for row in group if not _missing_number(getattr(row, "order_time", None))]
        records.append(
            {
                "symbol": first.symbol,
                "trade_time": pd.Timestamp(first.trade_time),
                "action": first.action,
                "quantity": quantity,
                "price": price,
                "proceeds": proceeds,
                "commission": commission,
                "multiplier": float(first.multiplier) if first.multiplier else 1.0,
                "currency": first.currency or "USD",
                "asset_category": asset,
                "source_row": int(first.source_row),
                "put_call": next((_row_text(row, "put_call") for row in group if _row_text(row, "put_call")), ""),
                "strike": next((_row_text(row, "strike") for row in group if _row_text(row, "strike")), ""),
                "expiry": next((_row_text(row, "expiry") for row in group if _row_text(row, "expiry")), ""),
                "open_close": next((_row_text(row, "open_close") for row in group if _row_text(row, "open_close")), ""),
                "order_type": next((_row_text(row, "order_type") for row in group if _row_text(row, "order_type")), ""),
                "order_time": min(order_times) if order_times else pd.NaT,
                "notes_codes": ";".join(notes),
                "latency_seconds": float(sum(latencies) / len(latencies)) if latencies else None,
            }
        )
    aggregated = pd.DataFrame.from_records(records, columns=EXEC_COLUMNS)
    if aggregated.empty:
        return aggregated
    aggregated["trade_time"] = pd.to_datetime(aggregated["trade_time"])
    return aggregated.sort_values(["trade_time", "source_row"], kind="mergesort").reset_index(drop=True)


def fx_multiplier(source: str, target: str, cad_per_usd: float) -> float | None:
    """Scale a source-currency amount into the target currency.

    The quote is CAD per 1 USD. CAD values are divided by that rate to reach USD.
    """
    origin = str(source or "USD").upper()
    destination = str(target or "USD").upper()
    if origin == destination:
        return 1.0
    if origin not in SUPPORTED_CURRENCIES or destination not in SUPPORTED_CURRENCIES:
        return None
    if cad_per_usd <= 0:
        return None
    if origin == "CAD" and destination == "USD":
        return 1.0 / cad_per_usd
    if origin == "USD" and destination == "CAD":
        return cad_per_usd
    return None


def convert_executions(executions: pd.DataFrame, target: str, cad_per_usd: float) -> tuple[pd.DataFrame, list[str]]:
    """Convert prices, proceeds, and commissions into USD or CAD."""
    notes: list[str] = []
    if executions is None or executions.empty:
        empty = executions if executions is not None else pd.DataFrame(columns=EXEC_COLUMNS)
        return empty, notes

    target = target.upper()
    kept_indexes = []
    factors = []
    skipped: set[str] = set()
    for index, source in executions["currency"].fillna("USD").items():
        factor = fx_multiplier(str(source), target, cad_per_usd)
        if factor is None:
            skipped.add(str(source).upper())
            continue
        kept_indexes.append(index)
        factors.append(factor)

    if skipped:
        notes.append(
            "Left out "
            + ", ".join(sorted(skipped))
            + " fills. Conversion covers USD and CAD only."
        )
    converted = executions.loc[kept_indexes].copy()
    if converted.empty:
        return converted.reset_index(drop=True), notes

    factor_series = pd.Series(factors, index=converted.index, dtype=float)
    converted["price"] = converted["price"].astype(float) * factor_series
    converted["commission"] = converted["commission"].astype(float) * factor_series
    proceeds = pd.to_numeric(converted["proceeds"], errors="coerce")
    converted["proceeds"] = proceeds * factor_series
    converted["currency"] = target
    converted = converted.reset_index(drop=True)
    if (factor_series.round(10) != 1).any():
        notes.append(f"Amounts are shown in {target} at {cad_per_usd:.2f} CAD per 1 USD.")
    return converted, notes


@dataclass
class MatchResult:
    closed: pd.DataFrame
    open_lots: pd.DataFrame
    openings: pd.DataFrame


def _empty_match() -> MatchResult:
    closed = pd.DataFrame(
        columns=[
            "symbol",
            "side",
            "quantity",
            "entry_time",
            "exit_time",
            "entry_price",
            "exit_price",
            "multiplier",
            "gross_pnl",
            "commission",
            "net_pnl",
            "hold_seconds",
            "currency",
        ]
    )
    open_lots = pd.DataFrame(
        columns=["symbol", "side", "quantity", "entry_time", "entry_price", "commission", "multiplier", "currency"]
    )
    openings = pd.DataFrame(columns=["symbol", "side", "quantity", "entry_time", "entry_price", "currency"])
    return MatchResult(closed, open_lots, openings)


def match_fifo(executions: pd.DataFrame) -> MatchResult:
    """Match buys and sells per symbol. Long and short cycles both close FIFO."""
    if executions is None or executions.empty:
        return _empty_match()

    ordered = executions.sort_values(["trade_time", "source_row"], kind="mergesort")
    books: dict[str, deque] = defaultdict(deque)
    closed_rows: list[dict] = []
    opening_rows: list[dict] = []

    for record in ordered.itertuples(index=False):
        symbol = record.symbol
        action = record.action
        total_qty = float(record.quantity)
        qty_left = total_qty
        price = float(record.price)
        commission = float(record.commission)
        multiplier = float(record.multiplier) if record.multiplier else 1.0
        trade_time = pd.Timestamp(record.trade_time)
        currency = record.currency or "USD"
        book = books[symbol]

        while qty_left > EPS and book:
            lot = book[0]
            closes_lot = (action == "BUY" and lot["side"] == "short") or (action == "SELL" and lot["side"] == "long")
            if not closes_lot:
                break
            match_qty = min(qty_left, lot["qty"])
            entry_comm = lot["comm_remaining"] * (match_qty / lot["qty"])
            exit_comm = commission * (match_qty / total_qty) if total_qty else 0.0
            if lot["side"] == "long":
                gross = (price - lot["price"]) * match_qty * lot["multiplier"]
            else:
                gross = (lot["price"] - price) * match_qty * lot["multiplier"]
            hold_seconds = (trade_time - lot["time"]).total_seconds()
            exit_notes = _row_text(record, "notes_codes")
            notes = " / ".join(part for part in (lot.get("notes_codes", ""), exit_notes) if part)
            closed_rows.append(
                {
                    "symbol": symbol,
                    "side": lot["side"],
                    "quantity": match_qty,
                    "entry_time": lot["time"],
                    "exit_time": trade_time,
                    "entry_price": lot["price"],
                    "exit_price": price,
                    "multiplier": lot["multiplier"],
                    "gross_pnl": gross,
                    "commission": entry_comm + exit_comm,
                    "net_pnl": gross - entry_comm - exit_comm,
                    "hold_seconds": hold_seconds,
                    "currency": lot["currency"],
                    "asset_category": lot.get("asset_category", ""),
                    "contract": lot.get("contract", ""),
                    "open_close": " → ".join(part for part in (lot.get("open_close", ""), _row_text(record, "open_close")) if part),
                    "order_type": lot.get("order_type", ""),
                    "notes_codes": notes,
                }
            )
            lot["qty"] -= match_qty
            lot["comm_remaining"] -= entry_comm
            qty_left -= match_qty
            if lot["qty"] <= EPS:
                book.popleft()

        if qty_left > EPS:
            side = "long" if action == "BUY" else "short"
            open_comm = commission * (qty_left / total_qty) if total_qty else 0.0
            book.append(
                {
                    "qty": qty_left,
                    "side": side,
                    "price": price,
                    "time": trade_time,
                    "comm_remaining": open_comm,
                    "multiplier": multiplier,
                    "currency": currency,
                    "asset_category": _row_text(record, "asset_category"),
                    "contract": _contract_label(
                        _row_text(record, "put_call"),
                        _row_text(record, "strike"),
                        _row_text(record, "expiry"),
                    ),
                    "open_close": _row_text(record, "open_close"),
                    "order_type": _row_text(record, "order_type"),
                    "notes_codes": _row_text(record, "notes_codes"),
                }
            )
            opening_rows.append(
                {
                    "symbol": symbol,
                    "side": side,
                    "quantity": qty_left,
                    "entry_time": trade_time,
                    "entry_price": price,
                    "currency": currency,
                }
            )

    open_rows = []
    for symbol, book in books.items():
        for lot in book:
            if lot["qty"] <= EPS:
                continue
            open_rows.append(
                {
                    "symbol": symbol,
                    "side": lot["side"],
                    "quantity": lot["qty"],
                    "entry_time": lot["time"],
                    "entry_price": lot["price"],
                    "commission": lot["comm_remaining"],
                    "multiplier": lot["multiplier"],
                    "currency": lot["currency"],
                }
            )

    result = _empty_match()
    if closed_rows:
        result.closed = pd.DataFrame(closed_rows)
        result.closed["entry_time"] = pd.to_datetime(result.closed["entry_time"])
        result.closed["exit_time"] = pd.to_datetime(result.closed["exit_time"])
    if open_rows:
        result.open_lots = pd.DataFrame(open_rows)
        result.open_lots["entry_time"] = pd.to_datetime(result.open_lots["entry_time"])
    if opening_rows:
        result.openings = pd.DataFrame(opening_rows)
        result.openings["entry_time"] = pd.to_datetime(result.openings["entry_time"])
    return result


# ---------------------------------------------------------------------------
# Analytics and behavioral rules
# ---------------------------------------------------------------------------


@dataclass
class Summary:
    net_pnl: float
    commissions: float
    win_rate: float | None
    profit_factor: float | None
    avg_win: float | None
    avg_loss: float | None
    max_drawdown: float
    closed_count: int
    win_count: int
    loss_count: int
    breakeven_count: int
    expectancy: float | None = None
    payoff_ratio: float | None = None
    return_on_drawdown: float | None = None
    commission_drag: float | None = None
    max_consecutive_losses: int = 0


def max_drawdown(pnls: list[float]) -> float:
    cumulative = 0.0
    peak = 0.0
    worst = 0.0
    for pnl in pnls:
        cumulative += pnl
        peak = max(peak, cumulative)
        worst = min(worst, cumulative - peak)
    return worst


def max_consecutive_losses(pnls: list[float]) -> int:
    streak = 0
    longest = 0
    for pnl in pnls:
        if pnl < 0:
            streak += 1
            longest = max(longest, streak)
        else:
            streak = 0
    return longest


def summarize(closed: pd.DataFrame, commissions: float) -> Summary:
    if closed is None or closed.empty:
        return Summary(0.0, commissions, None, None, None, None, 0.0, 0, 0, 0, 0)

    ordered = closed.sort_values(["exit_time", "symbol"], kind="mergesort")
    nets = ordered["net_pnl"].astype(float)
    wins = nets[nets > EPS]
    losses = nets[nets < -EPS]
    breakeven = int((nets.abs() <= EPS).sum())
    decisive = int(len(wins) + len(losses))
    win_rate = float(len(wins) / decisive) if decisive else None
    if len(losses) == 0:
        profit_factor = math.inf if len(wins) else None
    else:
        profit_factor = float(wins.sum() / abs(losses.sum())) if len(wins) else 0.0
    avg_win = float(wins.mean()) if len(wins) else None
    avg_loss = float(losses.mean()) if len(losses) else None
    drawdown = max_drawdown(nets.tolist())
    net_pnl = float(nets.sum())
    gross = net_pnl + float(commissions)
    if avg_win is None or avg_loss is None or abs(avg_loss) <= EPS:
        payoff = None
    else:
        payoff = abs(avg_win / avg_loss)
    recovery = None if abs(drawdown) <= EPS else net_pnl / abs(drawdown)
    drag = None if abs(gross) <= EPS else float(commissions) / gross
    return Summary(
        net_pnl=net_pnl,
        commissions=float(commissions),
        win_rate=win_rate,
        profit_factor=profit_factor,
        avg_win=avg_win,
        avg_loss=avg_loss,
        max_drawdown=drawdown,
        closed_count=int(len(ordered)),
        win_count=int(len(wins)),
        loss_count=int(len(losses)),
        breakeven_count=breakeven,
        expectancy=net_pnl / len(ordered),
        payoff_ratio=payoff,
        return_on_drawdown=recovery,
        commission_drag=drag,
        max_consecutive_losses=max_consecutive_losses(nets.tolist()),
    )


def revenge_trades(closed: pd.DataFrame, openings: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "symbol",
        "loss_exit",
        "loss_quantity",
        "loss_pnl",
        "reentry_time",
        "minutes_later",
        "new_quantity",
        "new_side",
    ]
    if closed is None or closed.empty or openings is None or openings.empty:
        return pd.DataFrame(columns=columns)

    losses = closed[closed["net_pnl"] < -EPS]
    rows = []
    for loss in losses.itertuples(index=False):
        later = openings[
            (openings["symbol"] == loss.symbol)
            & (openings["entry_time"] > loss.exit_time)
            & (openings["entry_time"] <= loss.exit_time + REVENGE_WINDOW)
            & (openings["quantity"] + EPS >= loss.quantity)
        ]
        for opening in later.itertuples(index=False):
            minutes = (opening.entry_time - loss.exit_time).total_seconds() / 60.0
            rows.append(
                {
                    "symbol": loss.symbol,
                    "loss_exit": loss.exit_time,
                    "loss_quantity": loss.quantity,
                    "loss_pnl": loss.net_pnl,
                    "reentry_time": opening.entry_time,
                    "minutes_later": minutes,
                    "new_quantity": opening.quantity,
                    "new_side": opening.side,
                }
            )
    if not rows:
        return pd.DataFrame(columns=columns)
    return _dedupe_revenge(pd.DataFrame(rows))


def _dedupe_revenge(matches: pd.DataFrame) -> pd.DataFrame:
    """Collapse partial-fill re-entries into one alert per re-entry burst."""
    columns = list(matches.columns)
    ordered = matches.sort_values(["symbol", "reentry_time", "loss_exit"], kind="mergesort")
    clusters: list[list] = []
    current: list = []
    cluster_start = None
    cluster_symbol = None
    for row in ordered.itertuples(index=False):
        same_burst = (
            current
            and row.symbol == cluster_symbol
            and (pd.Timestamp(row.reentry_time) - cluster_start).total_seconds() <= FILL_GROUP_SECONDS
        )
        if same_burst:
            current.append(row)
            continue
        if current:
            clusters.append(current)
        current = [row]
        cluster_start = pd.Timestamp(row.reentry_time)
        cluster_symbol = row.symbol
    if current:
        clusters.append(current)

    collapsed = []
    for cluster in clusters:
        openings = []
        seen = set()
        for row in cluster:
            identity = (pd.Timestamp(row.reentry_time), round(float(row.new_quantity), 8), row.new_side)
            if identity in seen:
                continue
            seen.add(identity)
            openings.append(row)
        reentry_time = min(pd.Timestamp(row.reentry_time) for row in openings)
        prior_losses = [row for row in cluster if pd.Timestamp(row.loss_exit) < reentry_time]
        loss = max(prior_losses or list(cluster), key=lambda row: pd.Timestamp(row.loss_exit))
        collapsed.append(
            {
                "symbol": loss.symbol,
                "loss_exit": pd.Timestamp(loss.loss_exit),
                "loss_quantity": loss.loss_quantity,
                "loss_pnl": loss.loss_pnl,
                "reentry_time": reentry_time,
                "minutes_later": (reentry_time - pd.Timestamp(loss.loss_exit)).total_seconds() / 60.0,
                "new_quantity": float(sum(row.new_quantity for row in openings)),
                "new_side": openings[0].new_side,
            }
        )
    if not collapsed:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(collapsed).sort_values(["loss_exit", "symbol"]).reset_index(drop=True)


def outsized_losses(closed: pd.DataFrame) -> pd.DataFrame:
    columns = ["symbol", "exit_time", "net_pnl", "average_loss", "multiple"]
    if closed is None or closed.empty:
        return pd.DataFrame(columns=columns)
    losses = closed[closed["net_pnl"] < -EPS]
    if len(losses) < 2:
        return pd.DataFrame(columns=columns)
    average_loss = abs(float(losses["net_pnl"].mean()))
    if average_loss <= EPS:
        return pd.DataFrame(columns=columns)
    flagged = losses[losses["net_pnl"].abs() > OUTSIZED_LOSS_MULTIPLE * average_loss].copy()
    if flagged.empty:
        return pd.DataFrame(columns=columns)
    flagged["average_loss"] = average_loss
    flagged["multiple"] = flagged["net_pnl"].abs() / average_loss
    return flagged[columns].sort_values("net_pnl").reset_index(drop=True)


def hold_comparison(closed: pd.DataFrame) -> dict | None:
    if closed is None or closed.empty:
        return None
    wins = closed[closed["net_pnl"] > EPS]["hold_seconds"]
    losses = closed[closed["net_pnl"] < -EPS]["hold_seconds"]
    if wins.empty or losses.empty:
        return None
    avg_win = float(wins.mean())
    avg_loss = float(losses.mean())
    return {
        "avg_win_seconds": avg_win,
        "avg_loss_seconds": avg_loss,
        "median_win_seconds": float(wins.median()),
        "median_loss_seconds": float(losses.median()),
        "losers_held_longer": avg_loss > avg_win + 1,
    }


def overtrading_alerts(executions: pd.DataFrame, closed: pd.DataFrame) -> pd.DataFrame:
    columns = ["trade_date", "fills", "median_fills", "fee_ratio", "reasons"]
    if executions is None or executions.empty:
        return pd.DataFrame(columns=columns)

    fills = executions.copy()
    fills["trade_date"] = fills["trade_time"].dt.date
    daily_fills = fills.groupby("trade_date").size()
    daily_fees = fills.groupby("trade_date")["commission"].sum()

    if closed is not None and not closed.empty:
        realized = closed.copy()
        realized["trade_date"] = realized["exit_time"].dt.date
        gross_profit = realized["gross_pnl"].clip(lower=0).groupby(realized["trade_date"]).sum()
    else:
        gross_profit = pd.Series(dtype=float)

    median_fills = float(daily_fills.median()) if len(daily_fills) else 0.0
    mean_fills = float(daily_fills.mean()) if len(daily_fills) else 0.0
    std_fills = float(daily_fills.std(ddof=0)) if len(daily_fills) else 0.0
    count_threshold = max(median_fills * 2.0, mean_fills + 1.5 * std_fills)

    rows = []
    for trade_date, fill_count in daily_fills.items():
        reasons = []
        if len(daily_fills) >= 2 and fill_count >= 5 and fill_count > count_threshold:
            reasons.append(f"{int(fill_count)} fills versus a typical day of {median_fills:.0f}")
        profit = float(gross_profit.get(trade_date, 0.0)) if len(gross_profit) else 0.0
        fees = float(daily_fees.get(trade_date, 0.0))
        ratio = fees / profit if profit > EPS else None
        if ratio is not None and ratio > FEE_DRAG_RATIO:
            reasons.append(f"fees were {ratio:.0%} of that day's gross profits")
        if reasons:
            rows.append(
                {
                    "trade_date": trade_date,
                    "fills": int(fill_count),
                    "median_fills": median_fills,
                    "fee_ratio": ratio,
                    "reasons": "; ".join(reasons),
                }
            )
    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows).sort_values("trade_date").reset_index(drop=True)


# ---------------------------------------------------------------------------
# Sample data
# ---------------------------------------------------------------------------


def _execution(
    source_row: int,
    symbol: str,
    when: str,
    action: str,
    quantity: float,
    price: float,
    commission: float,
) -> dict:
    signed = quantity if action == "SELL" else -quantity
    return {
        "symbol": symbol,
        "trade_time": pd.Timestamp(when),
        "action": action,
        "quantity": quantity,
        "price": price,
        "proceeds": signed * price,
        "commission": commission,
        "multiplier": 1.0,
        "currency": "USD",
        "asset_category": "Stocks",
        "source_row": source_row,
        "put_call": "",
        "strike": "",
        "expiry": "",
        "open_close": "",
        "order_type": "",
        "order_time": pd.NaT,
        "notes_codes": "",
        "latency_seconds": None,
    }


def _sample_round_trip(
    source: int,
    symbol: str,
    when: str,
    quantity: float,
    entry: float,
    exit_price: float,
    commission: float,
    hold_hours: float,
) -> list[dict]:
    opened = pd.Timestamp(when)
    closed = opened + timedelta(hours=hold_hours)
    return [
        _execution(source, symbol, str(opened), "BUY", quantity, entry, commission),
        _execution(source + 1, symbol, str(closed), "SELL", quantity, exit_price, commission),
    ]


def sample_executions() -> pd.DataFrame:
    """Synthetic fills from 2024 through early October 2026.

    The February 2024 block still includes a revenge entry, a fee-heavy day,
    and one outsized loss. Later round trips give each year something to filter.
    """
    rows = [
        _execution(1, "AAPL", "2024-02-01 09:35:00", "BUY", 100, 180.00, 1.00),
        _execution(2, "AAPL", "2024-02-01 10:05:00", "SELL", 100, 183.00, 1.00),
        _execution(3, "AAPL", "2024-02-02 09:40:00", "BUY", 100, 186.00, 1.00),
        _execution(4, "AAPL", "2024-02-02 13:40:00", "SELL", 100, 184.00, 1.00),
        _execution(5, "AAPL", "2024-02-02 13:48:00", "BUY", 200, 183.50, 1.50),
        _execution(6, "AAPL", "2024-02-02 15:48:00", "SELL", 200, 179.00, 1.50),
        _execution(7, "MSFT", "2024-02-05 10:00:00", "BUY", 50, 400.00, 1.00),
        _execution(8, "MSFT", "2024-02-05 16:00:00", "SELL", 50, 399.00, 1.00),
        _execution(9, "NVDA", "2024-02-06 10:00:00", "SELL", 20, 700.00, 1.00),
        _execution(10, "NVDA", "2024-02-06 10:25:00", "BUY", 20, 690.00, 1.00),
        _execution(11, "MSFT", "2024-02-08 10:00:00", "BUY", 10, 410.00, 1.00),
        _execution(12, "MSFT", "2024-02-08 10:15:00", "SELL", 10, 412.00, 1.00),
    ]
    start = pd.Timestamp("2024-02-07 09:30:00")
    source = 13
    for index in range(8):
        opened = start + timedelta(minutes=15 * index)
        closed = opened + timedelta(minutes=8)
        rows.append(_execution(source, "TSLA", str(opened), "BUY", 10, 200.00, 1.25))
        rows.append(_execution(source + 1, "TSLA", str(closed), "SELL", 10, 201.00, 1.25))
        source += 2
    # symbol, open time, quantity, entry, exit, commission, hold hours
    later_trips = [
        ("AAPL", "2024-01-16 10:05:00", 30, 168.00, 171.25, 1.00, 2),
        ("GOOGL", "2024-03-12 10:15:00", 20, 138.00, 141.40, 1.00, 2),
        ("MSFT", "2024-04-18 10:30:00", 15, 415.00, 412.20, 1.00, 6),
        ("NVDA", "2024-05-21 11:00:00", 12, 920.00, 934.00, 1.00, 2),
        ("TSLA", "2024-06-11 10:20:00", 25, 178.00, 182.50, 1.00, 3),
        ("GOOGL", "2024-07-16 13:10:00", 18, 185.00, 182.40, 1.00, 6),
        ("AMZN", "2024-08-08 10:40:00", 16, 167.00, 171.80, 1.00, 2),
        ("META", "2024-09-19 11:05:00", 10, 540.00, 548.00, 1.00, 2),
        ("AAPL", "2024-10-22 10:25:00", 22, 228.00, 225.10, 1.00, 5),
        ("MSFT", "2024-11-14 10:50:00", 12, 425.00, 431.00, 1.00, 2),
        ("NVDA", "2024-12-09 11:15:00", 8, 138.00, 143.20, 1.00, 2),
        ("AAPL", "2025-01-15 10:10:00", 28, 232.00, 236.40, 1.00, 2),
        ("TSLA", "2025-02-20 10:35:00", 20, 355.00, 349.50, 1.00, 6),
        ("GOOGL", "2025-03-11 11:00:00", 14, 168.00, 172.25, 1.00, 2),
        ("NVDA", "2025-04-08 10:20:00", 10, 110.00, 116.80, 1.00, 3),
        ("AMZN", "2025-05-19 13:00:00", 12, 205.00, 201.40, 1.00, 6),
        ("MSFT", "2025-06-12 10:45:00", 11, 455.00, 461.00, 1.00, 2),
        ("META", "2025-07-22 11:10:00", 8, 710.00, 718.50, 1.00, 2),
        ("SOXL", "2025-08-14 10:05:00", 40, 28.50, 30.10, 1.20, 2),
        ("AAPL", "2025-09-09 10:30:00", 18, 226.00, 223.20, 1.00, 5),
        ("TSLA", "2025-10-16 11:20:00", 15, 420.00, 428.00, 1.00, 2),
        ("NVDA", "2025-11-06 10:15:00", 9, 145.00, 151.40, 1.00, 2),
        ("MSFT", "2025-12-18 13:40:00", 10, 480.00, 476.25, 1.00, 6),
        ("AAPL", "2026-01-13 10:10:00", 24, 248.00, 252.80, 1.00, 2),
        ("NVDA", "2026-02-10 10:40:00", 11, 178.00, 184.50, 1.00, 2),
        ("TSLA", "2026-03-17 11:05:00", 16, 390.00, 384.20, 1.00, 6),
        ("GOOGL", "2026-04-21 10:25:00", 13, 175.00, 179.60, 1.00, 2),
        ("AMZN", "2026-05-12 10:55:00", 12, 220.00, 225.40, 1.00, 3),
        ("MSFT", "2026-06-18 13:15:00", 9, 490.00, 486.10, 1.00, 6),
        ("META", "2026-07-14 10:35:00", 7, 735.00, 744.00, 1.00, 2),
        ("AAPL", "2026-08-11 11:00:00", 20, 255.00, 259.75, 1.00, 2),
        ("NVDA", "2026-09-15 10:20:00", 10, 190.00, 186.40, 1.00, 5),
        ("TSLA", "2026-10-01 10:15:00", 14, 410.00, 416.80, 1.00, 2),
        ("MSFT", "2026-10-02 11:30:00", 8, 505.00, 509.50, 1.00, 2),
    ]
    source = 100
    for symbol, when, quantity, entry, exit_price, commission, hold_hours in later_trips:
        rows.extend(_sample_round_trip(source, symbol, when, quantity, entry, exit_price, commission, hold_hours))
        source += 2
    frame = pd.DataFrame(rows, columns=EXEC_COLUMNS)
    frame["trade_time"] = pd.to_datetime(frame["trade_time"])
    return frame.sort_values(["trade_time", "source_row"], kind="mergesort").reset_index(drop=True)


# ---------------------------------------------------------------------------
# Presentation helpers
# ---------------------------------------------------------------------------


def format_money(value: float | None, currency: str = "USD") -> str:
    if value is None or (isinstance(value, float) and not math.isfinite(value)):
        return "—"
    sign = "-" if value < 0 else ""
    if currency == "USD":
        prefix = "$"
    elif currency == "CAD":
        prefix = "C$"
    else:
        prefix = f"{currency} "
    return f"{sign}{prefix}{abs(value):,.2f}"


def format_duration(seconds: float | None) -> str:
    if seconds is None or (isinstance(seconds, float) and math.isnan(seconds)):
        return "—"
    remaining = int(abs(seconds))
    days, remaining = divmod(remaining, 86400)
    hours, remaining = divmod(remaining, 3600)
    minutes = remaining // 60
    parts = []
    if days:
        parts.append(f"{days}d")
    if hours:
        parts.append(f"{hours}h")
    parts.append(f"{minutes}m")
    return " ".join(parts)


def format_factor(value: float | None) -> str:
    if value is None:
        return "—"
    if math.isinf(value):
        return "∞"
    return f"{value:.2f}"


def format_percent(value: float | None) -> str:
    if value is None:
        return "—"
    return f"{value * 100:.1f}%"


def _in_date_range(stamps: pd.Series, start: date, end: date) -> pd.Series:
    days = stamps.dt.date
    return (days >= start) & (days <= end)


EXIT_DATE_PRESETS = (
    "All dates",
    "Year to date",
    "This month",
    "Last month",
    "This quarter",
    "Last quarter",
    "Last 7 days",
    "Last 30 days",
    "Last 90 days",
    "Last 6 months",
    "Last 12 months",
    "Calendar year",
    "Quarter",
    "Month",
    "Custom range",
)
_MONTH_NAMES = (
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)


def _month_bounds(year: int, month: int) -> tuple[date, date]:
    start = date(year, month, 1)
    if month == 12:
        end = date(year, 12, 31)
    else:
        end = date(year, month + 1, 1) - timedelta(days=1)
    return start, end


def _quarter_bounds(year: int, quarter: int) -> tuple[date, date]:
    if quarter not in {1, 2, 3, 4}:
        raise ValueError("Quarter must be 1, 2, 3, or 4.")
    start_month = 3 * (quarter - 1) + 1
    return date(year, start_month, 1), _month_bounds(year, start_month + 2)[1]


def _shift_months(day: date, months: int) -> date:
    month_index = day.month - 1 + months
    year = day.year + month_index // 12
    month = month_index % 12 + 1
    last_day = _month_bounds(year, month)[1].day
    return date(year, month, min(day.day, last_day))


def exit_window_bounds(
    preset: str,
    data_start: date,
    data_end: date,
    today: date | None = None,
    year: int | None = None,
    quarter: int | None = None,
    month: int | None = None,
    custom_start: date | None = None,
    custom_end: date | None = None,
) -> tuple[date, date]:
    """Inclusive exit-date window for a preset.

    Relative presets end on ``today``. A chosen year, quarter, or month uses
    that whole calendar period.
    """
    today = today or date.today()
    if preset == "All dates":
        return data_start, data_end
    if preset == "Year to date":
        return date(today.year, 1, 1), today
    if preset == "This month":
        return date(today.year, today.month, 1), today
    if preset == "Last month":
        if today.month == 1:
            return _month_bounds(today.year - 1, 12)
        return _month_bounds(today.year, today.month - 1)
    quarter_now = (today.month - 1) // 3 + 1
    if preset == "This quarter":
        start, _end = _quarter_bounds(today.year, quarter_now)
        return start, today
    if preset == "Last quarter":
        if quarter_now == 1:
            return _quarter_bounds(today.year - 1, 4)
        return _quarter_bounds(today.year, quarter_now - 1)
    if preset == "Last 7 days":
        return today - timedelta(days=6), today
    if preset == "Last 30 days":
        return today - timedelta(days=29), today
    if preset == "Last 90 days":
        return today - timedelta(days=89), today
    if preset == "Last 6 months":
        return _shift_months(today, -6), today
    if preset == "Last 12 months":
        return _shift_months(today, -12), today
    if preset == "Calendar year":
        if year is None:
            raise ValueError("Choose a year.")
        return date(year, 1, 1), date(year, 12, 31)
    if preset == "Quarter":
        if year is None or quarter is None:
            raise ValueError("Choose a year and a quarter.")
        return _quarter_bounds(year, quarter)
    if preset == "Month":
        if year is None or month is None:
            raise ValueError("Choose a year and a month.")
        return _month_bounds(year, month)
    if preset == "Custom range":
        if custom_start is None or custom_end is None:
            raise ValueError("Select both a start date and an end date.")
        if custom_end < custom_start:
            raise ValueError("The end date must be on or after the start date.")
        return custom_start, custom_end
    raise ValueError(f"Unknown exit date filter: {preset}")


def render_exit_date_filter(min_day: date, max_day: date, filter_key: str) -> tuple[date, date] | None:
    """Draw the exit-date filter and return the inclusive window, or None if a custom range is incomplete."""
    preset_key = f"exit-preset-{filter_key}"
    range_key = f"exit-range-{filter_key}"
    applied_key = f"exit-applied-{filter_key}"
    preset = st.pills(
        "Exit dates",
        EXIT_DATE_PRESETS,
        default="All dates",
        key=preset_key,
        help="Closed trades stay when the exit falls in this window. FIFO still uses earlier fills, so that exit keeps its original entry.",
        width="stretch",
    )
    if preset not in EXIT_DATE_PRESETS:
        preset = "All dates"

    years = list(range(max_day.year, min_day.year - 1, -1))
    year = max_day.year
    quarter = (max_day.month - 1) // 3 + 1
    month = max_day.month
    if preset in {"Calendar year", "Quarter", "Month"}:
        year_col, part_col = st.columns(2)
        with year_col:
            year = int(st.selectbox("Year", years, key=f"exit-year-{filter_key}"))
        if preset == "Quarter":
            with part_col:
                quarter_label = st.selectbox(
                    "Quarter",
                    ["Q1", "Q2", "Q3", "Q4"],
                    index=quarter - 1,
                    key=f"exit-quarter-{filter_key}",
                )
            quarter = int(str(quarter_label)[1])
        elif preset == "Month":
            with part_col:
                month_label = st.selectbox(
                    "Month",
                    _MONTH_NAMES,
                    index=month - 1,
                    key=f"exit-month-{filter_key}",
                )
            month = _MONTH_NAMES.index(str(month_label)) + 1

    if preset == "Custom range":
        signature = ("Custom range",)
    else:
        signature = (preset, year, quarter if preset == "Quarter" else 0, month if preset == "Month" else 0)
        if st.session_state.get(applied_key) != signature or range_key not in st.session_state:
            try:
                st.session_state[range_key] = exit_window_bounds(
                    preset,
                    min_day,
                    max_day,
                    year=year,
                    quarter=quarter,
                    month=month,
                )
            except ValueError as exc:
                st.warning(str(exc))
                return None
            st.session_state[applied_key] = signature

    floor = date(min_day.year, 1, 1)
    ceiling = date(max(max_day.year, date.today().year), 12, 31)
    stored = st.session_state.get(range_key, (min_day, max_day))
    if not isinstance(stored, (list, tuple)) or len(stored) != 2:
        stored = (min_day, max_day)
    start_stored = min(max(stored[0], floor), ceiling)
    end_stored = min(max(stored[1], floor), ceiling)
    if end_stored < start_stored:
        end_stored = start_stored
    st.session_state[range_key] = (start_stored, end_stored)

    range_col, _rest = st.columns([1.35, 2])
    with range_col:
        picked = st.date_input(
            "Exit date range",
            min_value=floor,
            max_value=ceiling,
            key=range_key,
            help="This range is the filter. A preset fills it in. Changing either date keeps your own range.",
        )
    if not isinstance(picked, (list, tuple)) or len(picked) != 2:
        st.warning("Select both a start date and an end date.")
        return None
    start_day, end_day = picked
    if end_day < start_day:
        st.warning("The end date must be on or after the start date.")
        return None
    if preset != "Custom range":
        expected = exit_window_bounds(
            preset,
            min_day,
            max_day,
            year=year,
            quarter=quarter,
            month=month,
        )
        if (start_day, end_day) != expected:
            st.session_state[preset_key] = "Custom range"
            st.session_state[applied_key] = ("Custom range",)
            st.rerun()
    st.caption(f"Showing exits from {start_day.isoformat()} through {end_day.isoformat()}.")
    if end_day < min_day or start_day > max_day:
        st.info(
            f"No exits in this file fall in that window. The file runs {min_day.isoformat()} to {max_day.isoformat()}."
        )
    return start_day, end_day


def cumulative_figure(closed: pd.DataFrame, currency: str) -> go.Figure | None:
    if closed is None or closed.empty:
        return None
    ordered = closed.sort_values(["exit_time", "symbol"], kind="mergesort").copy()
    ordered["cumulative"] = ordered["net_pnl"].cumsum()
    figure = go.Figure()
    figure.add_trace(
        go.Scatter(
            x=ordered["exit_time"],
            y=ordered["cumulative"],
            mode="lines+markers",
            name="Cumulative PnL",
            line={"color": BLUE, "width": 2.4},
            marker={"size": 7, "color": BLUE},
            fill="tozeroy",
            fillcolor="rgba(31, 75, 153, 0.12)",
            hovertemplate="%{x|%Y-%m-%d %H:%M}<br>Cumulative %{y:,.2f}<extra></extra>",
        )
    )
    figure.add_hline(y=0, line_width=1, line_dash="dot", line_color=SLATE)
    figure.update_layout(
        title="Cumulative realized PnL",
        margin=dict(l=80, r=24, t=48, b=40),
        height=420,
        hovermode="x unified",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        xaxis_title=None,
        showlegend=False,
    )
    figure.update_yaxes(
        title_text=currency,
        title_standoff=12,
        automargin=True,
        side="left",
        tickformat=",.2f",
        tickfont={"size": 12},
        zeroline=False,
    )
    return figure


WINNER_BAR = "#2ecc71"
LOSER_BAR = "#ff4d4d"


def ticker_totals(closed: pd.DataFrame) -> pd.DataFrame:
    """One row per symbol from closed FIFO round trips."""
    columns = ["Symbol", "Net PnL", "Trade Count", "Win Rate (%)", "Total Commissions"]
    if closed is None or closed.empty:
        return pd.DataFrame(columns=columns)
    rows = []
    for symbol, group in closed.groupby("symbol", sort=False):
        nets = group["net_pnl"].astype(float)
        wins = int((nets > EPS).sum())
        losses = int((nets < -EPS).sum())
        decisive = wins + losses
        rows.append(
            {
                "Symbol": symbol,
                "Net PnL": float(nets.sum()),
                "Trade Count": int(len(group)),
                "Win Rate (%)": (100.0 * wins / decisive) if decisive else None,
                "Total Commissions": float(group["commission"].sum()) if "commission" in group.columns else 0.0,
            }
        )
    frame = pd.DataFrame(rows, columns=columns)
    frame["_abs"] = frame["Net PnL"].abs()
    frame = frame.sort_values(["_abs", "Symbol"], ascending=[False, True], kind="mergesort").drop(columns="_abs")
    return frame.reset_index(drop=True)


def _signed_money(value: float, currency: str) -> str:
    text = format_money(value, currency)
    if value > EPS and not text.startswith("+"):
        return f"+{text}"
    return text


def top_drivers_figure(totals: pd.DataFrame, top_n: int, currency: str) -> go.Figure | None:
    """Horizontal bars for the largest winning and losing tickers only."""
    if totals is None or totals.empty:
        return None
    winners = totals[totals["Net PnL"] > EPS].nlargest(top_n, "Net PnL").sort_values("Net PnL", ascending=True)
    losers = totals[totals["Net PnL"] < -EPS].nsmallest(top_n, "Net PnL").sort_values("Net PnL", ascending=True)
    selected = pd.concat([losers, winners], ignore_index=True)
    if selected.empty:
        return None
    colors = [WINNER_BAR if value > 0 else LOSER_BAR for value in selected["Net PnL"]]
    labels = [_signed_money(float(value), currency) for value in selected["Net PnL"]]
    figure = go.Figure(
        go.Bar(
            x=selected["Net PnL"],
            y=selected["Symbol"],
            orientation="h",
            marker_color=colors,
            text=labels,
            textposition="outside",
            cliponaxis=False,
            hovertemplate="%{y}<br>%{text}<extra></extra>",
        )
    )
    low = float(selected["Net PnL"].min())
    high = float(selected["Net PnL"].max())
    pad = max(abs(low), abs(high), 1.0) * 0.45
    figure.update_layout(
        title=f"Top {top_n} winners and losers",
        margin=dict(l=100, r=28, t=48, b=40),
        height=max(320, 36 * len(selected) + 90),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        xaxis_title=currency,
        yaxis_title=None,
        showlegend=False,
    )
    figure.update_xaxes(range=[min(0.0, low) - pad, max(0.0, high) + pad], tickformat=",.0f")
    figure.update_yaxes(
        automargin=True,
        tickfont={"size": 12},
        categoryorder="array",
        categoryarray=selected["Symbol"].tolist(),
    )
    return figure


def show_ticker_table(totals: pd.DataFrame, currency: str) -> None:
    money_format = "dollar" if currency == "USD" else "C$%.2f"
    config = {
        "Net PnL": st.column_config.NumberColumn("Net PnL", format=money_format),
        "Trade Count": st.column_config.NumberColumn("Trade Count", format="%d"),
        "Win Rate (%)": st.column_config.NumberColumn("Win Rate (%)", format="%.1f%%"),
        "Total Commissions": st.column_config.NumberColumn("Total Commissions", format=money_format),
    }
    try:
        st.dataframe(totals, width="stretch", hide_index=True, column_config=config, row_height=28)
    except TypeError:
        st.dataframe(totals, use_container_width=True, hide_index=True, column_config=config)


def distribution_figure(summary: Summary) -> go.Figure | None:
    if summary.closed_count == 0:
        return None
    labels = ["Wins", "Losses"]
    values = [summary.win_count, summary.loss_count]
    colors = [GREEN, RED]
    if summary.breakeven_count:
        labels.append("Breakeven")
        values.append(summary.breakeven_count)
        colors.append(SLATE)
    figure = go.Figure(
        go.Bar(
            x=labels,
            y=values,
            marker_color=colors,
            text=values,
            textposition="outside",
            hovertemplate="%{x}: %{y}<extra></extra>",
        )
    )
    figure.update_layout(
        title="Trade distribution",
        margin={"l": 16, "r": 16, "t": 48, "b": 16},
        height=380,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        yaxis_title="Closed trades",
        xaxis_title=None,
    )
    return figure


def underwater_series(closed: pd.DataFrame) -> pd.DataFrame:
    """Peak-to-trough gap on the cumulative realized-PnL curve. Peak starts at zero."""
    ordered = closed.sort_values(["exit_time", "symbol"], kind="mergesort")
    cumulative = 0.0
    peak = 0.0
    rows = []
    for record in ordered.itertuples(index=False):
        cumulative += float(record.net_pnl)
        peak = max(peak, cumulative)
        rows.append({"exit_time": record.exit_time, "underwater": cumulative - peak})
    return pd.DataFrame(rows)


def underwater_figure(closed: pd.DataFrame, currency: str) -> go.Figure | None:
    if closed is None or closed.empty:
        return None
    series = underwater_series(closed)
    figure = go.Figure()
    figure.add_trace(
        go.Scatter(
            x=series["exit_time"],
            y=series["underwater"],
            mode="lines",
            name="Drawdown",
            line={"color": RED, "width": 2.2},
            fill="tozeroy",
            fillcolor="rgba(214, 69, 69, 0.18)",
            hovertemplate="%{x|%Y-%m-%d %H:%M}<br>Below peak %{y:,.2f}<extra></extra>",
        )
    )
    figure.add_hline(y=0, line_width=1, line_dash="dot", line_color=SLATE)
    figure.update_layout(
        title="Drawdown from peak",
        margin=dict(l=80, r=24, t=48, b=40),
        height=320,
        hovermode="x unified",
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        xaxis_title=None,
        showlegend=False,
    )
    figure.update_yaxes(
        title_text=currency,
        title_standoff=12,
        automargin=True,
        side="left",
        tickformat=",.2f",
        tickfont={"size": 12},
        zeroline=False,
    )
    return figure


def asset_type_series(frame: pd.DataFrame) -> pd.Series:
    """Map each fill to Single Stocks, 3x Leveraged ETFs, or Options."""
    symbols = frame["symbol"].astype(str).str.upper()
    options = frame["asset_category"].map(_asset_bucket).eq("Options") if "asset_category" in frame.columns else False
    if "put_call" in frame.columns:
        has_right = frame["put_call"].map(_cell_text).ne("")
        options = has_right if options is False else (options | has_right)
    kinds = pd.Series("Single Stocks", index=frame.index)
    kinds = kinds.mask(symbols.isin(LEVERAGED_ETFS), "3x Leveraged ETFs")
    if options is not False:
        kinds = kinds.mask(options, "Options")
    return kinds


def leveraged_net_pnl(closed: pd.DataFrame) -> float:
    if closed is None or closed.empty or "symbol" not in closed.columns:
        return 0.0
    symbols = closed["symbol"].astype(str).str.upper()
    subset = closed.loc[symbols.isin(LEVERAGED_ETFS), "net_pnl"]
    if subset.empty:
        return 0.0
    return float(subset.astype(float).sum())


def hold_bucket_name(seconds: float) -> str:
    for name, start, end in HOLD_BUCKETS:
        if seconds >= start and (end is None or seconds < end):
            return name
    return HOLD_BUCKETS[-1][0]


def hold_bucket_stats(closed: pd.DataFrame) -> pd.DataFrame:
    columns = ["Bucket", "Net PnL", "Trades"]
    if closed is None or closed.empty or "hold_seconds" not in closed.columns:
        return pd.DataFrame(columns=columns)
    work = closed.dropna(subset=["hold_seconds"]).copy()
    if work.empty:
        return pd.DataFrame(columns=columns)
    work["Bucket"] = work["hold_seconds"].astype(float).map(hold_bucket_name)
    rows = []
    for name, _, _ in HOLD_BUCKETS:
        group = work[work["Bucket"] == name]
        rows.append({"Bucket": name, "Net PnL": float(group["net_pnl"].sum()) if len(group) else 0.0, "Trades": int(len(group))})
    return pd.DataFrame(rows, columns=columns)


def _mean_hold_seconds(closed: pd.DataFrame, winning: bool) -> float | None:
    if closed is None or closed.empty or "hold_seconds" not in closed.columns:
        return None
    nets = closed["net_pnl"].astype(float)
    mask = nets > EPS if winning else nets < -EPS
    held = closed.loc[mask, "hold_seconds"].dropna().astype(float)
    if held.empty:
        return None
    return float(held.mean())


def hold_outcome_figure(closed: pd.DataFrame) -> go.Figure | None:
    if closed is None or closed.empty or "hold_seconds" not in closed.columns:
        return None
    work = closed.dropna(subset=["hold_seconds"]).copy()
    nets = work["net_pnl"].astype(float)
    work = work[(nets > EPS) | (nets < -EPS)]
    if work.empty:
        return None
    work["Hours"] = work["hold_seconds"].astype(float) / 3600.0
    work["Outcome"] = work["net_pnl"].map(lambda value: "Win" if value > EPS else "Loss")
    figure = go.Figure()
    for label, color in (("Win", GREEN), ("Loss", RED)):
        hours = work.loc[work["Outcome"] == label, "Hours"]
        if hours.empty:
            continue
        figure.add_trace(
            go.Box(
                x=hours,
                y=[label] * len(hours),
                name=label,
                orientation="h",
                marker_color=color,
                line_color=color,
                boxmean=True,
                hovertemplate=f"{label}<br>%{{x:.1f}} hours (%{{customdata:.1f}} days)<extra></extra>",
                customdata=hours / 24.0,
            )
        )
    figure.update_layout(
        title="Hold time by outcome",
        margin=dict(l=80, r=24, t=48, b=40),
        height=320,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
        xaxis_title="Hours held",
    )
    figure.update_xaxes(automargin=True, tickformat=",.1f")
    return figure


def hold_bucket_figure(stats: pd.DataFrame, currency: str) -> go.Figure | None:
    if stats is None or stats.empty:
        return None
    colors = [GREEN if value >= 0 else RED for value in stats["Net PnL"]]
    figure = go.Figure(
        go.Bar(
            y=stats["Bucket"],
            x=stats["Net PnL"],
            orientation="h",
            marker_color=colors,
            customdata=stats["Trades"],
            hovertemplate="%{y}<br>Net %{x:,.2f} " + currency + "<br>%{customdata} trades<extra></extra>",
        )
    )
    figure.update_layout(
        title="Net PnL by hold time",
        margin=dict(l=100, r=24, t=48, b=40),
        height=320,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        showlegend=False,
        xaxis_title=currency,
        yaxis_title=None,
    )
    figure.update_xaxes(automargin=True, tickformat=",.0f")
    figure.update_yaxes(categoryorder="array", categoryarray=[name for name, _, _ in HOLD_BUCKETS][::-1])
    return figure


def render_hold_duration(closed: pd.DataFrame, currency: str) -> None:
    st.subheader("Hold duration")
    if closed is None or closed.empty:
        st.info("No closed trades in this date range to measure hold time.")
        return
    winner_hold, loser_hold = st.columns(2)
    winner_hold.metric(
        "Avg winner hold time",
        format_duration(_mean_hold_seconds(closed, True)),
        help="Average time from entry to exit for closed trades with net PnL above zero.",
    )
    loser_hold.metric(
        "Avg loser hold time",
        format_duration(_mean_hold_seconds(closed, False)),
        help="Average time from entry to exit for closed trades with net PnL below zero.",
    )
    left, right = st.columns(2)
    outcome = hold_outcome_figure(closed)
    buckets = hold_bucket_stats(closed)
    with left:
        if outcome is not None:
            show_chart(outcome)
        else:
            st.caption("Need a winning or losing trade to compare hold times.")
    with right:
        chart = hold_bucket_figure(buckets, currency)
        if chart is not None:
            show_chart(chart)


def _asset_bucket(value: object) -> str:
    text = _cell_text(value).upper()
    if "OPT" in text or "OPTION" in text:
        return "Options"
    if any(token in text for token in ("STK", "STOCK", "EQUITY")):
        return "Stocks"
    if not text:
        return "Other"
    return _cell_text(value)


def group_breakdown(closed: pd.DataFrame, labels: pd.Series, order: list[str] | None = None) -> pd.DataFrame:
    work = closed.copy()
    work["_group"] = pd.Series(labels, index=closed.index).map(_cell_text)
    rows = []
    for name, group in work.groupby("_group", sort=False):
        nets = group["net_pnl"].astype(float)
        wins = int((nets > EPS).sum())
        losses = int((nets < -EPS).sum())
        decisive = wins + losses
        rows.append(
            {
                "Group": str(name),
                "Net PnL": float(nets.sum()),
                "Trades": int(len(group)),
                "Win rate": (wins / decisive) if decisive else None,
            }
        )
    frame = pd.DataFrame(rows, columns=["Group", "Net PnL", "Trades", "Win rate"])
    if frame.empty:
        return frame
    if order:
        rank = {name: index for index, name in enumerate(order)}
        frame["_order"] = frame["Group"].map(lambda name: rank.get(name, len(rank)))
        frame = frame.sort_values(["_order", "Group"], kind="mergesort").drop(columns="_order")
    else:
        frame = frame.sort_values("Group", kind="mergesort")
    return frame.reset_index(drop=True)


def group_pnl_figure(stats: pd.DataFrame, title: str, currency: str) -> go.Figure | None:
    if stats is None or stats.empty:
        return None
    colors = [GREEN if value >= 0 else RED for value in stats["Net PnL"]]
    figure = go.Figure(
        go.Bar(
            x=stats["Group"],
            y=stats["Net PnL"],
            marker_color=colors,
            customdata=list(zip(stats["Trades"], stats["Win rate"].map(lambda value: "" if value is None else f"{value * 100:.1f}%"))),
            hovertemplate="%{x}<br>Net %{y:,.2f}<br>%{customdata[0]} trades<br>Win rate %{customdata[1]}<extra></extra>",
        )
    )
    figure.update_layout(
        title=title,
        margin=dict(l=80, r=24, t=48, b=40),
        height=320,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        xaxis_title=None,
        showlegend=False,
    )
    figure.update_yaxes(
        title_text=currency,
        title_standoff=12,
        automargin=True,
        tickformat=",.2f",
        zeroline=True,
        zerolinecolor=SLATE,
    )
    return figure


def format_breakdown(stats: pd.DataFrame, currency: str) -> pd.DataFrame:
    view = stats.copy()
    view["Net PnL"] = view["Net PnL"].map(lambda value: format_money(float(value), currency))
    view["Win rate"] = view["Win rate"].map(format_percent)
    return view


def render_breakdown(closed: pd.DataFrame, currency: str) -> None:
    st.subheader("Breakdown")
    if closed is None or closed.empty:
        st.info("No closed trades in this date range to break down.")
        return
    st.caption("Closed trades by the hour and weekday they were opened, then by side and asset class. Win rate leaves out breakeven trades.")
    hours = closed["entry_time"].dt.hour.map(lambda hour: f"{int(hour):02d}:00")
    weekdays = closed["entry_time"].dt.day_name()
    sides = closed["side"].fillna("").map(lambda value: str(value).title())
    if "asset_category" in closed.columns:
        assets = closed["asset_category"].map(_asset_bucket)
    else:
        assets = pd.Series(["Other"] * len(closed), index=closed.index)
    blocks = [
        (group_breakdown(closed, hours, [f"{hour:02d}:00" for hour in range(24)]), "Net PnL by hour of entry"),
        (group_breakdown(closed, weekdays, WEEKDAY_ORDER), "Net PnL by weekday of entry"),
        (group_breakdown(closed, sides, ["Long", "Short"]), "Long versus short"),
        (group_breakdown(closed, assets, ["Stocks", "Options", "Other"]), "Stocks versus options"),
    ]
    for offset in (0, 2):
        left, right = st.columns(2)
        for column, (stats, title) in zip((left, right), blocks[offset : offset + 2]):
            with column:
                chart = group_pnl_figure(stats, title, currency)
                if chart is not None:
                    show_chart(chart)
                if not stats.empty:
                    show_table(format_breakdown(stats, currency))


def prepare_trade_log(closed: pd.DataFrame) -> pd.DataFrame:
    log = closed.sort_values("exit_time", ascending=False).copy()
    log["Side"] = log["side"].str.title()
    log["Result"] = log["net_pnl"].map(lambda value: "Win" if value > EPS else "Loss" if value < -EPS else "Breakeven")
    log["Hold"] = log["hold_seconds"].map(format_duration)
    log = log.rename(
        columns={
            "symbol": "Symbol",
            "quantity": "Quantity",
            "entry_time": "Entry",
            "exit_time": "Exit",
            "entry_price": "Entry price",
            "exit_price": "Exit price",
            "gross_pnl": "Gross PnL",
            "commission": "Commission",
            "net_pnl": "Net PnL",
            "currency": "Currency",
        }
    )
    columns = [
        "Symbol",
        "Side",
        "Result",
        "Quantity",
        "Entry",
        "Exit",
        "Hold",
        "Entry price",
        "Exit price",
        "Gross PnL",
        "Commission",
        "Net PnL",
        "Currency",
    ]
    if "asset_category" in log.columns and log["asset_category"].map(_cell_text).ne("").any():
        log["Asset"] = log["asset_category"].map(_cell_text)
        columns.insert(1, "Asset")
    if "contract" in log.columns and log["contract"].map(_cell_text).ne("").any():
        log["Contract"] = log["contract"].map(_cell_text)
        columns.insert(2, "Contract")
    if "open_close" in log.columns and log["open_close"].map(_cell_text).ne("").any():
        log["Open/Close"] = log["open_close"].map(_cell_text)
        columns.insert(4, "Open/Close")
    if "order_type" in log.columns and log["order_type"].map(_cell_text).ne("").any():
        log["Order type"] = log["order_type"].map(_cell_text)
        columns.append("Order type")
    if "notes_codes" in log.columns and log["notes_codes"].map(_cell_text).ne("").any():
        log["Notes"] = log["notes_codes"].map(_cell_text)
        columns.append("Notes")
    return log[columns]


def show_chart(figure: go.Figure) -> None:
    try:
        st.plotly_chart(figure, width="stretch", theme=None)
    except TypeError:
        st.plotly_chart(figure, use_container_width=True)


def show_table(frame: pd.DataFrame) -> None:
    try:
        st.dataframe(frame, width="stretch", hide_index=True)
    except TypeError:
        st.dataframe(frame, use_container_width=True, hide_index=True)


def inject_css() -> None:
    st.markdown(
        """
        <style>
        .block-container {padding-top: 1.1rem; padding-bottom: 1.5rem; max-width: 1180px;}
        div[data-testid="stMetric"] {
            background: rgba(128, 128, 128, 0.08);
            border: 1px solid rgba(128, 128, 128, 0.18);
            border-radius: 12px;
            padding: 0.55rem 0.7rem;
        }
        div[data-testid="stMetricValue"] {font-variant-numeric: tabular-nums;}
        section[data-testid="stSidebar"] {
            width: 300px !important;
        }
        section[data-testid="stSidebar"] > div {
            padding-top: 0.6rem;
        }
        section[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
            gap: 0.35rem;
        }
        section[data-testid="stSidebar"] [data-testid="stHeading"] {
            margin: 0.15rem 0 0;
            padding: 0;
        }
        section[data-testid="stSidebar"] [data-testid="stHeading"] h2 {
            font-size: 0.95rem;
        }
        section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {
            font-size: 0.75rem;
            line-height: 1.2;
        }
        section[data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
            padding: 0.35rem 0.5rem;
        }
        section[data-testid="stSidebar"] button[kind="secondary"],
        section[data-testid="stSidebar"] [data-testid="stBaseButton-secondary"] {
            min-height: 2rem;
            padding-top: 0.2rem;
            padding-bottom: 0.2rem;
        }
        section[data-testid="stSidebar"] [data-baseweb="tag"] {
            height: 18px !important;
            max-height: 18px !important;
            margin: 1px 2px !important;
            padding: 0 2px 0 5px !important;
            font-size: 10px !important;
        }
        section[data-testid="stSidebar"] [data-baseweb="tag"] span {
            font-size: 10px !important;
            line-height: 16px !important;
        }
        section[data-testid="stSidebar"] [data-testid="stMultiSelect"] [data-baseweb="select"] > div {
            max-height: 5.5rem;
            overflow-y: auto;
            padding-top: 2px;
            padding-bottom: 2px;
        }
        """
        + (
            """
        [data-testid="stToolbarActions"],
        [data-testid="stToolbarActionButton"],
        [data-testid="stStatusWidget"],
        [data-testid="stMainMenu"],
        [data-testid="stDeployButton"] {
            display: none !important;
            visibility: hidden !important;
            pointer-events: none !important;
        }
        """
            if hosted_publicly()
            else ""
        )
        + """
        </style>
        """,
        unsafe_allow_html=True,
    )


def running_locally() -> bool:
    """The IBKR sync control is only for the copy running on this computer."""
    return not hosted_publicly()


def show_page_title() -> None:
    st.title("IBKR Trade Analysis")
    st.caption("Parse Interactive Brokers exports, match round trips with FIFO, and review realized PnL.")


def render_empty_state() -> None:
    st.info("Drop a CSV, or turn on sample data to explore the dashboard.")
    with st.expander("Exports this app can read", expanded=True):
        st.markdown(
            """
            **Activity statement CSV.** In IBKR, open Reports, then Statements, then Activity, and download CSV.
            The parser looks for the `Trades` section and keeps execution rows (`Order`, `Trade`, or `Execution`).

            **Flex Query CSV.** Required columns are Symbol, Date/Time, Buy/Sell, Quantity, Trade Price, Proceeds, IB Commission, Currency, and Asset Class.
            FX Rate To Base, Multiplier, Put/Call, Strike, Expiry, Open/Close Indicator, Order Type, Order Time, and Notes/Codes are used when present.
            Account, security, and exchange IDs are ignored. Execution time comes from Date/Time, not Order Time or Trade Date.

            **Flat trade CSV.** A normal table also works when it has symbol, date/time, buy/sell or signed quantity,
            quantity, and price. Commission and proceeds are used when those columns exist.

            Realized PnL is recalculated with FIFO. It can differ from the Realized P/L column already in an IBKR file.
            """
        )


def render_dashboard(executions: pd.DataFrame, notes: list[str], source_label: str) -> None:
    symbols = sorted(executions["symbol"].unique().tolist())
    min_day = executions["trade_time"].min().date()
    max_day = executions["trade_time"].max().date()
    filter_key = hashlib.sha1(f"{source_label}|{min_day}|{max_day}|{','.join(symbols)}".encode()).hexdigest()[:12]
    asset_classes = sorted({_cell_text(value) for value in executions["asset_category"].tolist() if _cell_text(value)})

    show_page_title()
    exit_window = render_exit_date_filter(min_day, max_day, filter_key)

    with st.sidebar:
        st.markdown("**Filters**")
        selected = st.multiselect("Tickers", options=symbols, default=symbols, key=f"tickers-{filter_key}")
        selected_assets = asset_classes
        if len(asset_classes) > 1:
            selected_assets = st.multiselect("Asset class", options=asset_classes, default=asset_classes, key=f"assets-{filter_key}")
        exclude_outliers = st.checkbox("Exclude Major Outliers (AMD & SOXL)", key="exclude-outliers")
        selected_types = st.multiselect(
            "Asset Type",
            options=ASSET_TYPE_OPTIONS,
            default=ASSET_TYPE_OPTIONS,
            key="asset-type",
            help="Single stocks stay separate from leveraged and yield ETFs such as SOXL, SOXS, TQQQ, and MSTY. Options are detected from the asset class.",
        )
        st.markdown("**Currency**")
        target_currency = st.radio(
            "Show amounts in",
            ["USD", "CAD"],
            index=1,
            horizontal=True,
            key="display-currency",
            help="CAD rescales every dollar amount by the rate below. USD shows account dollars and does not use that rate.",
        )
        cad_per_usd = st.number_input(
            "CAD per 1 USD",
            min_value=0.50,
            max_value=2.50,
            value=DEFAULT_CAD_PER_USD,
            step=0.01,
            format="%.2f",
            key="cad-per-usd",
            help="Press Enter or the arrows. With CAD selected, every dollar amount is multiplied by this rate.",
        )
        fx_slot = st.empty()
        export_clicked = st.button(
            "Export report for a friend",
            use_container_width=True,
            help="Saves trade_analytics_report.html in this folder. Your friend opens it in a browser. Charts load from the internet.",
        )
        st.caption(source_label)

    if exit_window is None:
        return
    start_day, end_day = exit_window
    if not selected:
        st.warning("Select at least one ticker.")
        return
    if not selected_assets and asset_classes:
        st.warning("Select at least one asset class.")
        return
    if not selected_types:
        st.warning("Select at least one asset type.")
        return

    scoped = executions[executions["symbol"].isin(selected)].copy()
    if asset_classes and selected_assets != asset_classes:
        scoped = scoped[scoped["asset_category"].map(_cell_text).isin(selected_assets)].copy()
    if exclude_outliers and not scoped.empty:
        scoped = scoped[~scoped["symbol"].astype(str).str.upper().isin(OUTLIER_SYMBOLS)].copy()
    if set(selected_types) != set(ASSET_TYPE_OPTIONS) and not scoped.empty:
        scoped = scoped[asset_type_series(scoped).isin(selected_types)].copy()
    aggregated = aggregate_fills(scoped)
    converted, fx_notes = convert_executions(aggregated, target_currency, float(cad_per_usd))
    notes = list(notes) + fx_notes
    matched = match_fifo(converted)
    closed = matched.closed
    if not closed.empty:
        closed = closed[_in_date_range(closed["exit_time"], start_day, end_day)].copy()
    openings = matched.openings
    raw_in_range = scoped[_in_date_range(scoped["trade_time"], start_day, end_day)].copy()
    aggregated_in_range = converted[_in_date_range(converted["trade_time"], start_day, end_day)].copy() if not converted.empty else converted
    commissions = float(aggregated_in_range["commission"].sum()) if not aggregated_in_range.empty else 0.0
    summary = summarize(closed, commissions)
    currency = target_currency
    if export_clicked:
        from export_report import write_friend_report

        report_path = write_friend_report(closed, summary, currency, float(cad_per_usd), source_label)
        st.sidebar.success(f"Saved {report_path.name}. Send that file.")

    revenge = revenge_trades(closed, openings)
    large_losses = outsized_losses(closed)
    holding = hold_comparison(closed)
    busy_days = overtrading_alerts(aggregated_in_range, closed)

    if notes:
        with st.expander("Import notes", expanded=False):
            for note in notes:
                st.write(note)

    latency = median_latency(raw_in_range)
    latency_note = f"  ·  median order-to-fill {format_latency(latency)}" if latency is not None else ""
    if currency == "CAD":
        fx_slot.caption(f"CAD amounts use {float(cad_per_usd):.2f} per 1 USD.")
    else:
        fx_slot.caption(
            f"USD totals skip this rate. At {float(cad_per_usd):.2f}: "
            f"{format_money(summary.net_pnl * float(cad_per_usd), 'CAD')}."
        )
    st.caption(
        f"{source_label}  ·  exits {start_day.isoformat()} to {end_day.isoformat()}  ·  "
        f"{len(raw_in_range)} raw fills  ·  {len(aggregated_in_range)} aggregated trades  ·  "
        f"{summary.closed_count} closed trades  ·  {len(matched.open_lots)} open lots  ·  "
        f"shown in {currency} at {float(cad_per_usd):.2f} CAD per 1 USD  ·  FIFO, net of IB commission{latency_note}"
    )

    if exclude_outliers:
        st.info("⚠️ Outlier Filter Active: Excluding AMD and SOXL")
    if set(selected_types) != set(ASSET_TYPE_OPTIONS):
        st.caption("Asset type: " + ", ".join(selected_types) + ".")

    metric_row_1 = st.columns(4)
    metric_row_1[0].metric(
        "Net realized PnL",
        format_money(summary.net_pnl, currency),
        help="Closed round trips in the selected dates, after commission, in the currency selected in the sidebar.",
    )
    metric_row_1[1].metric(
        "Commissions & fees",
        format_money(summary.commissions, currency),
        help="IB Commission on fills in the selected dates.",
    )
    metric_row_1[2].metric(
        "Win rate",
        format_percent(summary.win_rate),
        help="Winners divided by winners and losers. Breakeven trades are left out.",
    )
    metric_row_1[3].metric(
        "Profit factor",
        format_factor(summary.profit_factor),
        help="Total amount won divided by the total amount lost.",
    )

    metric_row_2 = st.columns(4)
    metric_row_2[0].metric(
        "Expectancy",
        format_money(summary.expectancy, currency),
        help="Net realized PnL divided by the number of closed trades.",
    )
    metric_row_2[1].metric(
        "Average winning trade",
        format_money(summary.avg_win, currency),
        help="Mean net PnL of the winning trades.",
    )
    metric_row_2[2].metric(
        "Average losing trade",
        format_money(summary.avg_loss, currency),
        help="Mean net PnL of the losing trades.",
    )
    metric_row_2[3].metric(
        "Max drawdown",
        format_money(summary.max_drawdown, currency),
        help="Largest peak-to-trough drop on the cumulative realized-PnL curve.",
    )

    st.markdown("**Risk & Efficiency**")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric(
        "Payoff ratio",
        format_factor(summary.payoff_ratio),
        help="Ratio of average win size to average loss size. Values > 1.0 mean your winners are larger than your losers.",
    )
    col2.metric(
        "Return / max drawdown",
        format_factor(summary.return_on_drawdown),
        help="Net Realized PnL divided by Max Drawdown. Measures return efficiency relative to peak drawdown risk.",
    )
    col3.metric(
        "Commission drag",
        format_percent(summary.commission_drag),
        help="Percentage of gross profit consumed by broker commissions and execution fees.",
    )
    col4.metric(
        "Max consecutive losses",
        f"{summary.max_consecutive_losses:,}",
        help="The longest streak of consecutive losing trades in the selected date range.",
    )
    st.columns(4)[0].metric(
        "Leveraged net PnL",
        format_money(leveraged_net_pnl(closed), currency),
        help="Net realized PnL of closed trades in the leveraged and yield ETFs (SOXL, SOXS, TQQQ, SQQQ, UPRO, SPXU, MSTY, TSLY, CONY, AMZY, APLY, UVXY), using the current filters.",
    )

    metric_row_3 = st.columns(4)
    metric_row_3[0].metric(
        "Closed trades",
        f"{summary.closed_count:,}",
        help="Round trips that finished inside the selected dates.",
    )
    metric_row_3[1].metric(
        "Raw fills",
        f"{len(raw_in_range):,}",
        help="Execution rows before partial fills are grouped.",
    )
    metric_row_3[2].metric(
        "Aggregated trades",
        f"{len(aggregated_in_range):,}",
        help="Fills merged when the same ticker and side print within 5 seconds. FIFO uses these.",
    )
    metric_row_3[3].metric(
        "Open lots",
        f"{len(matched.open_lots):,}",
        help="Positions still open after FIFO. They are not included in realized PnL.",
    )
    grouped_away = len(raw_in_range) - len(aggregated_in_range)
    st.columns(4)[0].metric(
        "Fills combined",
        f"{max(grouped_away, 0):,}",
        help="Raw fills minus aggregated trades. This is how many partial prints were folded into a grouped trade.",
    )

    render_hold_duration(closed, currency)

    st.subheader("Performance")
    if closed.empty:
        st.info("No round trips close inside this date range. Open lots are listed further down.")
    else:
        equity = cumulative_figure(closed, currency)
        if equity is not None:
            show_chart(equity)
        drawdown = underwater_figure(closed, currency)
        if drawdown is not None:
            show_chart(drawdown)
        left, right = st.columns(2)
        totals = ticker_totals(closed)
        distribution = distribution_figure(summary)
        with left:
            top_n = st.slider(
                "Top N Tickers",
                min_value=3,
                max_value=10,
                value=5,
                key="top-n-tickers",
                help="How many of the largest winners and how many of the largest losers to plot.",
            )
            drivers = top_drivers_figure(totals, int(top_n), currency)
            if drivers is not None:
                show_chart(drivers)
            else:
                st.caption("No winning or losing tickers in this date range.")
        if distribution is not None:
            with right:
                show_chart(distribution)
        st.caption("Every ticker with a closed round trip in the selected dates. Largest absolute net PnL is listed first.")
        show_ticker_table(totals, currency)

    render_breakdown(closed, currency)

    st.subheader("Mistake identification")
    render_behavior(revenge, large_losses, holding, busy_days, currency)
    render_event_flags(raw_in_range)

    st.subheader("Trade log")
    if not closed.empty:
        log = prepare_trade_log(closed)
        with st.expander(f"Closed trades ({len(log)})", expanded=True):
            show_table(log)
            st.download_button(
                "Download closed trades",
                data=log.to_csv(index=False).encode("utf-8"),
                file_name="fifo_closed_trades.csv",
                mime="text/csv",
            )
    else:
        st.caption("No closed trades to list for this filter.")

    with st.expander(f"Open lots ({len(matched.open_lots)})", expanded=closed.empty and not matched.open_lots.empty):
        if matched.open_lots.empty:
            st.caption("Every fill in the selected tickers is matched.")
        else:
            open_view = matched.open_lots.sort_values("entry_time").copy()
            open_view["side"] = open_view["side"].str.title()
            open_view = open_view.rename(
                columns={
                    "symbol": "Symbol",
                    "side": "Side",
                    "quantity": "Quantity",
                    "entry_time": "Opened",
                    "entry_price": "Entry price",
                    "commission": "Commission remaining",
                    "multiplier": "Multiplier",
                    "currency": "Currency",
                }
            )
            show_table(open_view)
            st.caption("Open lots are not marked to market. Drawdown above uses realized PnL only.")

    with st.expander(f"Raw fills ({len(raw_in_range)})", expanded=False):
        raw = raw_in_range.sort_values("trade_time", ascending=False).rename(
            columns={
                "symbol": "Symbol",
                "trade_time": "Time",
                "action": "Action",
                "quantity": "Quantity",
                "price": "Price",
                "proceeds": "Proceeds",
                "commission": "Commission",
                "multiplier": "Multiplier",
                "currency": "Currency",
                "asset_category": "Asset",
                "put_call": "Put/Call",
                "strike": "Strike",
                "expiry": "Expiry",
                "open_close": "Open/Close",
                "order_type": "Order type",
                "order_time": "Order time",
                "notes_codes": "Notes",
                "latency_seconds": "Order latency (sec)",
            }
        )
        show = [
            "Symbol",
            "Asset",
            "Time",
            "Order time",
            "Order latency (sec)",
            "Action",
            "Open/Close",
            "Order type",
            "Quantity",
            "Price",
            "Proceeds",
            "Commission",
            "Multiplier",
            "Currency",
            "Put/Call",
            "Strike",
            "Expiry",
            "Notes",
        ]
        show = [column for column in show if column in raw.columns]
        show_table(raw[show])


def median_latency(executions: pd.DataFrame) -> float | None:
    if executions is None or executions.empty or "latency_seconds" not in executions.columns:
        return None
    values = pd.to_numeric(executions["latency_seconds"], errors="coerce").dropna()
    values = values[values >= 0]
    if values.empty:
        return None
    return float(values.median())


def format_latency(seconds: float) -> str:
    if seconds < 90:
        return f"{seconds:.0f}s"
    return format_duration(seconds)


def event_flags(executions: pd.DataFrame) -> pd.DataFrame:
    columns = ["Symbol", "Time", "Asset", "Notes", "Flag"]
    if executions is None or executions.empty or "notes_codes" not in executions.columns:
        return pd.DataFrame(columns=columns)
    rows = []
    for record in executions.itertuples(index=False):
        notes = _row_text(record, "notes_codes")
        flags = _note_flags(notes)
        if not flags:
            continue
        rows.append(
            {
                "Symbol": record.symbol,
                "Time": record.trade_time,
                "Asset": _row_text(record, "asset_category"),
                "Notes": notes,
                "Flag": ", ".join(flags),
            }
        )
    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows).sort_values("Time").reset_index(drop=True)


def render_event_flags(executions: pd.DataFrame) -> None:
    flags = event_flags(executions)
    if flags.empty:
        return
    st.warning(f"{len(flags)} fill{'' if len(flags) == 1 else 's'} flagged from Notes/Codes (wash sale, assignment, or exercise).")
    show_table(flags)


def render_behavior(revenge, large_losses, holding, busy_days, currency: str) -> None:
    triggered = False

    if revenge.empty:
        st.caption("Revenge trading: no new position was opened within 10 minutes of a loss, on the same symbol, at equal or larger size.")
    else:
        triggered = True
        st.error(f"{len(revenge)} possible revenge {('entry' if len(revenge) == 1 else 'entries')} after a losing close.")
        view = revenge.copy()
        view["Loss PnL"] = view["loss_pnl"].map(lambda value: format_money(value, currency))
        view["Minutes later"] = view["minutes_later"].map(lambda value: f"{value:.1f}")
        view["New side"] = view["new_side"].str.title()
        view = view.rename(
            columns={
                "symbol": "Symbol",
                "loss_exit": "Loss closed",
                "loss_quantity": "Loss size",
                "reentry_time": "Re-entry",
                "new_quantity": "New size",
            }
        )
        show_table(view[["Symbol", "Loss closed", "Loss size", "Loss PnL", "Re-entry", "Minutes later", "New size", "New side"]])

    if busy_days.empty:
        st.caption("Overtrading: no day combined an unusually high fill count with the rest of this file, and fees stayed within 20% of that day's gross profits.")
    else:
        triggered = True
        st.warning(f"{len(busy_days)} trading {('day' if len(busy_days) == 1 else 'days')} flagged for activity or fee drag.")
        view = busy_days.copy()
        view["Fees / gross profit"] = view["fee_ratio"].map(lambda value: "—" if value is None or (isinstance(value, float) and math.isnan(value)) else f"{value:.0%}")
        view = view.rename(columns={"trade_date": "Date", "fills": "Fills", "reasons": "Why it was flagged"})
        show_table(view[["Date", "Fills", "Fees / gross profit", "Why it was flagged"]])
        st.caption("A high-count day has at least 5 fills and more activity than both twice the median day and the mean plus 1.5 standard deviations.")

    if holding is None:
        st.caption("Hold time: need at least one winning and one losing trade to compare how long each is kept open.")
    else:
        comparison = (
            f"Average winning hold {format_duration(holding['avg_win_seconds'])} "
            f"(median {format_duration(holding['median_win_seconds'])}). "
            f"Average losing hold {format_duration(holding['avg_loss_seconds'])} "
            f"(median {format_duration(holding['median_loss_seconds'])})."
        )
        if holding["losers_held_longer"]:
            triggered = True
            st.warning("Losing trades are held longer than winning trades. " + comparison)
        else:
            st.success("Winning trades are held at least as long as losing trades. " + comparison)

    if large_losses.empty:
        st.caption("Outsized losses: no single loss was larger than twice the average loss in this filter.")
    else:
        triggered = True
        st.error(f"{len(large_losses)} {('loss' if len(large_losses) == 1 else 'losses')} exceeded 2× the average loss.")
        view = large_losses.copy()
        view["Net PnL"] = view["net_pnl"].map(lambda value: format_money(value, currency))
        view["Average loss"] = view["average_loss"].map(lambda value: format_money(-value, currency))
        view["Multiple"] = view["multiple"].map(lambda value: f"{value:.1f}×")
        view = view.rename(columns={"symbol": "Symbol", "exit_time": "Exit"})
        show_table(view[["Symbol", "Exit", "Net PnL", "Average loss", "Multiple"]])

    if not triggered:
        st.success("None of the revenge, overtrading, hold-time, or outsized-loss rules fired for this filter.")


@st.cache_data(show_spinner=False)
def cached_executions(name: str, data: bytes) -> tuple[pd.DataFrame, tuple[str, ...]]:
    executions, notes = load_executions(name, data)
    return executions, tuple(notes)


def save_flex_sync(combined_csv: str, chunks: list[dict], directory: Path = FLEX_DIR) -> list[str]:
    """Replace any previous Flex download with this sync."""
    directory.mkdir(parents=True, exist_ok=True)
    ignore = directory / ".gitignore"
    if not ignore.exists():
        ignore.write_text("*.csv\n", encoding="utf-8")
    for path in directory.glob("*.csv"):
        path.unlink()
    written = ["combined.csv"]
    (directory / "combined.csv").write_text(combined_csv, encoding="utf-8")
    for chunk in chunks:
        csv_text = str(chunk.get("csv") or "")
        if not csv_text.strip():
            continue
        name = f"{chunk['start']}_to_{chunk['end']}.csv"
        (directory / name).write_text(csv_text, encoding="utf-8")
        written.append(name)
    return written


def flex_year_options(today: date | None = None) -> list[str]:
    """Calendar years, newest first, plus the rolling history and a custom range."""
    current = (today or date.today()).year
    return [str(year) for year in range(current, 2009, -1)] + ["Last 5 years", "Custom dates"]


def planned_flex_sync(period: str, picked, today: date | None = None) -> tuple[date | None, date | None]:
    """Return the inclusive dates to request from IBKR.

    ``(None, None)`` means the rolling five-year download. A calendar year runs
    from January 1 through December 31, or through today when that year is still open.
    """
    today = today or date.today()
    if period == "Last 5 years":
        return None, None
    if period == "Custom dates":
        if not isinstance(picked, (list, tuple)) or len(picked) != 2:
            raise ValueError("Select both a start date and an end date, then sync.")
        start, end = picked
        if end < start:
            raise ValueError("The end date must be on or after the start date.")
        if start > today or end > today:
            raise ValueError("IBKR can only return dates through today.")
        return start, end
    year = int(period)
    start = date(year, 1, 1)
    end = min(date(year, 12, 31), today)
    if start > end:
        raise ValueError("That year has not started.")
    return start, end


def flex_sync_caption(period: str, picked, today: date | None = None) -> str:
    """Describe the weekday windows a sync would send to IBKR."""
    start, end = planned_flex_sync(period, picked, today)
    if start is None or end is None:
        windows = year_windows(today=today)
        label = "Last 5 years"
    else:
        windows = date_windows(start, end)
        label = period if period not in {"Custom dates", "Last 5 years"} else f"{start.isoformat()} to {end.isoformat()}"
    if not windows:
        return f"{label} · no weekdays in that range"
    first, last = windows[0][0], windows[-1][1]
    requests = "1 request" if len(windows) == 1 else f"{len(windows)} requests"
    window_text = f"{first.isoformat()} to {last.isoformat()}"
    if label == window_text:
        return f"{window_text} · {requests}"
    return f"{label} · {window_text} · {requests}"


def load_saved_flex(directory: Path = FLEX_DIR) -> tuple[pd.DataFrame, list[str]] | None:
    path = directory / "combined.csv"
    if not path.is_file():
        return None
    executions, notes = load_executions(path.name, path.read_bytes())
    notes = ["Loaded the saved IBKR sync from data/flex. Sync again to replace those files."] + list(notes)
    return executions, notes


@st.dialog("How to make a Flex Query file", width="large")
def show_flex_query_guide() -> None:
    st.markdown(
        """
**Step 1. Open the Flex Queries page**

1. Log in to your Interactive Brokers Client Portal.
2. Go to **Performance & Reports → Flex Queries**.

**Step 2. Create the activity Flex Query**

1. Next to **Activity Flex Query**, click the **+** icon to create a new template.
2. Enter a query name, for example `Trade Analysis Export`.

**Step 3. Select the required columns**

1. Click **Trades** under the Sections list.
2. In the pop-up, select only these 20 columns:

- Currency
- FX Rate To Base
- Asset Class
- Symbol
- Multiplier
- Date/Time
- Quantity
- TradePrice
- Proceeds
- IB Commission
- Open/Close Indicator
- Buy/Sell
- Order Type
- Strike
- Expiry
- Put/Call
- IB Commission Currency
- Notes/Codes
- Order Time
- Level Of Detail

3. Scroll down and click **Save**.

**Step 4. Export each year**

IBKR limits one query to 365 days, so download one year at a time.

1. Under **General Settings**, set **Period** to **Custom Date Range**.
2. For the first file, use **2021/01/01** through **2021/12/31**, set the format to **CSV**, and save the query.
3. Click the yellow **Run** button next to the saved query and download the CSV. Name it `trades_2021.csv`.
4. Click the pencil icon next to that query.
5. Change the dates to the next year, for example **2022/01/01** through **2022/12/31**, then click **Save**.
6. Click **Run** again and save that file as `trades_2022.csv`.
7. Repeat for each year through the current date.

**Step 5. Combine the yearly files into one**

Put the downloaded files in one folder, then use Python or Command Prompt. Upload `Trade_History_Combined.csv` with the file button in this sidebar.
        """
    )
    st.markdown("**Method A. Python**")
    st.caption("Place the yearly CSV files in one folder and run this script there.")
    st.code(
        """import glob
import pandas as pd

csv_files = glob.glob("trades_*.csv")
combined_df = pd.concat([pd.read_csv(f) for f in csv_files], ignore_index=True)
combined_df.to_csv("Trade_History_Combined.csv", index=False)
print(f"Successfully combined {len(csv_files)} yearly files into Trade_History_Combined.csv!")
""",
        language="python",
    )
    st.markdown("**Method B. Windows Command Prompt**")
    st.caption("Open Command Prompt in the folder that holds the yearly files, then run:")
    st.code("copy trades_*.csv Trade_History_Combined.csv", language="text")


def main() -> None:
    st.set_page_config(page_title="IBKR Trade Analysis", page_icon="📈", layout="wide")
    if hosted_publicly():
        st.set_option("client.toolbarMode", "minimal")
    inject_css()

    if "ibkr_bundle" not in st.session_state:
        st.session_state.ibkr_bundle = None
    if "prefer_ibkr" not in st.session_state:
        st.session_state.prefer_ibkr = False
    if "ibkr_error" not in st.session_state:
        st.session_state.ibkr_error = ""
    if "flex_autoload_done" not in st.session_state:
        st.session_state.flex_autoload_done = False
    if not st.session_state.flex_autoload_done and st.session_state.ibkr_bundle is None:
        st.session_state.flex_autoload_done = True
        loaded = load_saved_flex()
        if loaded is not None:
            loaded_executions, loaded_notes = loaded
            st.session_state.ibkr_bundle = {
                "executions": loaded_executions,
                "notes": loaded_notes,
            }
            st.session_state.prefer_ibkr = True

    with st.sidebar:
        st.markdown("**Data**")
        local_app = running_locally()
        sync_clicked = False
        sync_period = "Last 5 years"
        sync_picked = None
        if local_app:
            sync_period = st.selectbox(
                "Year",
                options=flex_year_options(),
                index=0,
                key="ibkr-sync-period",
                help="Sync one calendar year, the last 5 years, or a custom date range. A range longer than 365 days is downloaded in chunks.",
            )
            if sync_period == "Custom dates":
                sync_picked = st.date_input(
                    "Date range",
                    value=(date(date.today().year, 1, 1), date.today()),
                    min_value=date(2010, 1, 1),
                    max_value=date.today(),
                    key="ibkr-sync-dates",
                    help="Choose the start and end date, then click Sync.",
                )
            sync_clicked = st.button("🔄 Sync with IBKR (Live Flex Query)", use_container_width=True)
            saved = (FLEX_DIR / "combined.csv").is_file()
            synced_count = 0
            if st.session_state.prefer_ibkr and st.session_state.ibkr_bundle:
                synced_count = len(st.session_state.ibkr_bundle["executions"])
            try:
                period_note = flex_sync_caption(sync_period, sync_picked)
            except ValueError as exc:
                period_note = str(exc)
            sync_bits = [f"Query {DEFAULT_QUERY_ID}", period_note]
            if saved:
                sync_bits.append("loads from data/flex")
            if synced_count:
                sync_bits.append(f"{synced_count:,} executions")
            st.caption(" · ".join(sync_bits))
            if st.session_state.ibkr_error:
                st.error(st.session_state.ibkr_error)
        uploaded = st.file_uploader(
            "Drop a Flex Query, trade confirmation, or trade CSV",
            type=["csv", "txt", "xlsx", "xlsm"],
            help="A newly chosen file is used for this session.",
        )
        if st.button("How to make Query flex file", use_container_width=True):
            show_flex_query_guide()
        use_sample = st.toggle("Use sample data", value=False)

    if sync_clicked and running_locally():
        try:
            sync_start, sync_end = planned_flex_sync(sync_period, sync_picked)
        except ValueError as exc:
            st.session_state.ibkr_error = str(exc)
            sync_start = sync_end = None
            sync_ready = False
        else:
            sync_ready = True
            st.session_state.ibkr_error = ""
        if sync_ready:
            if sync_start is None or sync_end is None:
                status_label = "Syncing 5 years of IBKR trades in 365-day chunks."
            else:
                status_label = f"Syncing IBKR trades from {sync_start.isoformat()} to {sync_end.isoformat()}."
            with st.status(status_label, expanded=True) as status:
                def on_chunk(index: int, total: int, start, end) -> None:
                    status.write(f"Chunk {index} of {total}: {start.isoformat()} to {end.isoformat()}")

                try:
                    if sync_start is None or sync_end is None:
                        downloaded = fetch_ibkr_trades(progress=on_chunk)
                    else:
                        downloaded = fetch_ibkr_trades(start=sync_start, end=sync_end, progress=on_chunk)
                    csv_text = downloaded.attrs.get("source_csv", "")
                    synced_executions, synced_notes = load_executions("ibkr_flex.csv", csv_text.encode("utf-8"))
                    for warning in downloaded.attrs.get("chunk_warnings") or []:
                        synced_notes.append(warning)
                    written = save_flex_sync(csv_text, list(downloaded.attrs.get("chunks") or []))
                    synced_notes.append("Saved this sync in data/flex. The next sync replaces those files.")
                    st.session_state.ibkr_bundle = {
                        "executions": synced_executions,
                        "notes": synced_notes,
                    }
                    st.session_state.prefer_ibkr = True
                    st.session_state.ibkr_error = ""
                    status.write("Saved " + ", ".join(written))
                    status.update(label=f"Synced {len(synced_executions):,} executions", state="complete")
                except (FlexServiceError, Exception) as exc:
                    message = str(exc)
                    st.session_state.ibkr_error = message
                    status.write(message)
                    status.update(label="Flex sync failed", state="error")

    if uploaded is not None:
        upload_id = f"{uploaded.name}:{getattr(uploaded, 'size', len(uploaded.getvalue()))}"
        previous_upload = st.session_state.get("last_upload_id")
        if previous_upload is None:
            st.session_state.last_upload_id = upload_id
        elif upload_id != previous_upload:
            st.session_state.last_upload_id = upload_id
            if not sync_clicked:
                st.session_state.prefer_ibkr = False

    sample_was_on = bool(st.session_state.get("sample_was_on", False))
    if use_sample and not sample_was_on and not sync_clicked:
        st.session_state.prefer_ibkr = False
    st.session_state.sample_was_on = use_sample

    bundle = st.session_state.ibkr_bundle
    if st.session_state.prefer_ibkr and bundle:
        executions = bundle["executions"]
        notes = list(bundle["notes"])
        source = "IBKR Flex Query"
    elif uploaded is not None:
        try:
            executions, notes_tuple = cached_executions(uploaded.name, uploaded.getvalue())
            notes = list(notes_tuple)
        except Exception as exc:
            show_page_title()
            st.error(f"Could not read that file. {exc}")
            return
        source = uploaded.name
    elif use_sample:
        executions = sample_executions()
        notes = ["Sample data is synthetic and runs from January 2024 through 2 October 2026. It is here so you can click through the dashboard before uploading a report."]
        source = "Sample data"
    else:
        show_page_title()
        render_empty_state()
        return

    if executions.empty:
        show_page_title()
        st.error("No executions were found in that file.")
        for note in notes:
            st.write(note)
        return

    render_dashboard(executions, notes, source)


# ---------------------------------------------------------------------------
# Self-check (python app.py --self-test)
# ---------------------------------------------------------------------------


def _exec(symbol, when, action, quantity, price, commission=0.0, proceeds=None, multiplier=1.0, source_row=1, currency="USD"):
    return {
        "symbol": symbol,
        "trade_time": pd.Timestamp(when),
        "action": action,
        "quantity": quantity,
        "price": price,
        "proceeds": proceeds,
        "commission": commission,
        "multiplier": multiplier,
        "currency": currency,
        "asset_category": "Stocks",
        "source_row": source_row,
    }


def _self_test() -> None:
    long_win = match_fifo(
        pd.DataFrame(
            [
                _exec("AAPL", "2024-01-02 09:30", "BUY", 100, 100, 1, source_row=1),
                _exec("AAPL", "2024-01-02 10:30", "SELL", 100, 110, 1, source_row=2),
            ]
        )
    )
    assert abs(long_win.closed.iloc[0]["gross_pnl"] - 1000) < 1e-6
    assert abs(long_win.closed.iloc[0]["net_pnl"] - 998) < 1e-6

    partial = match_fifo(
        pd.DataFrame(
            [
                _exec("MSFT", "2024-01-03 09:30", "BUY", 100, 10, 10, source_row=1),
                _exec("MSFT", "2024-01-03 10:30", "SELL", 40, 12, 4, source_row=2),
                _exec("MSFT", "2024-01-03 11:30", "SELL", 60, 11, 6, source_row=3),
            ]
        )
    )
    assert list(partial.closed["quantity"]) == [40, 60]
    assert abs(partial.closed.iloc[0]["net_pnl"] - 72) < 1e-6
    assert abs(partial.closed.iloc[1]["net_pnl"] - 48) < 1e-6
    assert partial.open_lots.empty

    short = match_fifo(
        pd.DataFrame(
            [
                _exec("NVDA", "2024-01-04 09:30", "SELL", 10, 100, source_row=1),
                _exec("NVDA", "2024-01-04 10:30", "BUY", 10, 90, source_row=2),
            ]
        )
    )
    assert short.closed.iloc[0]["side"] == "short"
    assert abs(short.closed.iloc[0]["net_pnl"] - 100) < 1e-6

    reversal = match_fifo(
        pd.DataFrame(
            [
                _exec("AMD", "2024-01-05 09:30", "BUY", 10, 100, source_row=1),
                _exec("AMD", "2024-01-05 10:30", "SELL", 15, 110, source_row=2),
                _exec("AMD", "2024-01-05 11:30", "BUY", 5, 100, source_row=3),
            ]
        )
    )
    assert len(reversal.closed) == 2
    assert abs(reversal.closed["net_pnl"].sum() - 150) < 1e-6
    assert reversal.open_lots.empty

    fifo = match_fifo(
        pd.DataFrame(
            [
                _exec("QQQ", "2024-01-06 09:30", "BUY", 100, 10, source_row=1),
                _exec("QQQ", "2024-01-06 09:31", "BUY", 100, 20, source_row=2),
                _exec("QQQ", "2024-01-06 09:32", "SELL", 150, 30, source_row=3),
            ]
        )
    )
    assert abs(fifo.closed["gross_pnl"].sum() - 2500) < 1e-6
    assert abs(fifo.open_lots.iloc[0]["quantity"] - 50) < 1e-6

    option = match_fifo(
        pd.DataFrame(
            [
                _exec("AAPL OPT", "2024-01-07 09:30", "BUY", 1, 2.5, proceeds=-250, multiplier=100, source_row=1),
                _exec("AAPL OPT", "2024-01-07 10:30", "SELL", 1, 3.0, proceeds=300, multiplier=100, source_row=2),
            ]
        )
    )
    assert abs(option.closed.iloc[0]["gross_pnl"] - 50) < 1e-6

    assert abs(max_drawdown([100, 100, -150]) - (-150)) < 1e-6
    assert abs(max_drawdown([-50, 100]) - (-50)) < 1e-6
    assert max_consecutive_losses([20, -10, -10, -10, 5]) == 3
    assert max_consecutive_losses([5, 5]) == 0
    risk_closed = pd.DataFrame(
        {
            "symbol": ["A", "A", "A", "A"],
            "net_pnl": [30.0, -10.0, -10.0, 30.0],
            "exit_time": pd.to_datetime(["2024-02-01", "2024-02-02", "2024-02-03", "2024-02-04"]),
        }
    )
    risk = summarize(risk_closed, 10.0)
    assert abs(risk.payoff_ratio - 3) < 1e-6
    assert abs(risk.return_on_drawdown - 2) < 1e-6
    assert abs(risk.commission_drag - 0.2) < 1e-6
    assert risk.max_consecutive_losses == 2
    assert format_percent(risk.commission_drag) == "20.0%"
    assert asset_type_series(pd.DataFrame({"symbol": ["AMD"], "asset_category": ["STK"], "put_call": [""]})).iloc[0] == "Single Stocks"
    assert asset_type_series(pd.DataFrame({"symbol": ["SOXL"], "asset_category": ["STK"], "put_call": [""]})).iloc[0] == "3x Leveraged ETFs"
    assert asset_type_series(pd.DataFrame({"symbol": ["AAPL"], "asset_category": ["OPT"], "put_call": ["C"]})).iloc[0] == "Options"
    assert abs(leveraged_net_pnl(pd.DataFrame({"symbol": ["SOXL", "AAPL"], "net_pnl": [-50.0, 20.0]})) - (-50.0)) < 1e-6
    assert hold_bucket_name(60) == "<1 hour"
    assert hold_bucket_name(3600) == "1–24 hours"
    assert hold_bucket_name(2 * 86400) == "1–7 days"
    assert hold_bucket_name(8 * 86400) == ">7 days"
    held = pd.DataFrame(
        {
            "symbol": ["A", "A", "B", "B"],
            "net_pnl": [10.0, -4.0, 6.0, -8.0],
            "hold_seconds": [1800.0, 7200.0, 3 * 86400.0, 10 * 86400.0],
            "exit_time": pd.to_datetime(["2024-03-01", "2024-03-02", "2024-03-03", "2024-03-04"]),
        }
    )
    buckets = hold_bucket_stats(held)
    assert list(buckets["Trades"]) == [1, 1, 1, 1]
    assert abs(float(buckets.loc[buckets["Bucket"] == "<1 hour", "Net PnL"].iloc[0]) - 10) < 1e-6
    assert abs(_mean_hold_seconds(held, True) - ((1800 + 3 * 86400) / 2)) < 1e-6
    assert abs(_mean_hold_seconds(held, False) - ((7200 + 10 * 86400) / 2)) < 1e-6

    section = """Statement,Header,Field Name,Field Value
Statement,Data,Period,Year
Trades,Header,DataDiscriminator,Asset Category,Currency,Symbol,Date/Time,Quantity,T. Price,Proceeds,Comm/Fee,Code
Trades,Data,Order,Stocks,USD,AAPL,"2024-03-01, 09:35:00",100,100,-10000,-1.00,O
Trades,Data,Order,Stocks,USD,AAPL,"2024-03-01, 10:00:00",-100,110,11000,-1.00,C
Trades,SubTotal,,Stocks,USD,AAPL,,,,,
Trades,Data,ClosedLot,Stocks,USD,AAPL,"2024-03-01, 10:00:00",100,110,11000,-1,C
"""
    parsed, parse_notes = load_executions("activity.csv", section.encode("utf-8"))
    assert len(parsed) == 2, parse_notes
    assert abs(match_fifo(parsed).closed.iloc[0]["net_pnl"] - 998) < 1e-6
    assert _infer_multiplier(1, 2.5, -250) == 100
    assert _infer_multiplier(100, 10, -1000) == 1

    flat = "Symbol,Trade Date,Buy/Sell,Quantity,Price,Commission\nMSFT,2024-04-01,BUY,10,100,0.5\nMSFT,2024-04-02,SELL,10,120,0.5\n"
    flat_exec, _ = load_executions("trades.csv", flat.encode("utf-8"))
    assert len(flat_exec) == 2
    assert abs(match_fifo(flat_exec).closed.iloc[0]["net_pnl"] - 199) < 1e-6

    demo = sample_executions()
    assert set(demo["trade_time"].dt.year) == {2024, 2025, 2026}
    assert demo["trade_time"].min().date() == date(2024, 1, 16)
    assert demo["trade_time"].max().date() == date(2026, 10, 2)
    demo_match = match_fifo(demo)
    revenge = revenge_trades(demo_match.closed, demo_match.openings)
    assert not revenge.empty
    assert (revenge["symbol"] == "AAPL").any()
    assert revenge["minutes_later"].min() <= 10
    large = outsized_losses(demo_match.closed)
    assert not large.empty
    holding = hold_comparison(demo_match.closed)
    assert holding is not None and holding["losers_held_longer"]
    alerts = overtrading_alerts(demo, demo_match.closed)
    assert (alerts["trade_date"] == date(2024, 2, 7)).any()
    assert alerts["fee_ratio"].dropna().max() > 0.20

    summary = summarize(demo_match.closed, float(demo["commission"].sum()))
    assert summary.win_count > 0 and summary.loss_count > 0
    assert summary.profit_factor is not None and summary.profit_factor > 0

    ibkr_fees = _fee_cost(pd.Series(["-1.00", "-1.50", "0.25"]), False)
    assert abs(float(ibkr_fees.iloc[0]) - 1.0) < 1e-9
    assert abs(float(ibkr_fees.iloc[2]) - (-0.25)) < 1e-9
    positive_fees = _fee_cost(pd.Series(["1.00", "0.50"]), False)
    assert abs(float(positive_fees.sum()) - 1.5) < 1e-9

    kept = normalize_executions(
        pd.DataFrame(
            [
                {"Symbol": "TOTAL", "Date/Time": "2024-01-01 09:30:00", "Quantity": "1", "T. Price": "10"},
                {"Symbol": "BTOTAL", "Date/Time": "2024-01-01 09:30:00", "Quantity": "1", "T. Price": "10", "Buy/Sell": "BUY"},
                {"Symbol": "BTOTAL", "Date/Time": "2024-01-01 10:30:00", "Quantity": "-1", "T. Price": "11"},
            ]
        )
    )[0]
    assert kept["symbol"].tolist() == ["BTOTAL", "BTOTAL"]

    european = (
        "Symbol;Date/Time;Quantity;T. Price;Comm/Fee\n"
        "AAPL;2024-06-01 09:30:00;10;100,50;-1,25\n"
        "AAPL;2024-06-01 10:30:00;-10;101,00;-1,25\n"
    )
    euro_exec, euro_notes = load_executions("flex.csv", european.encode("utf-8"))
    assert len(euro_exec) == 2, euro_notes
    assert abs(euro_exec.iloc[0]["price"] - 100.50) < 1e-9
    assert abs(match_fifo(euro_exec).closed.iloc[0]["net_pnl"] - 2.50) < 1e-6

    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Symbol", "Trade Date", "Buy/Sell", "Quantity", "Price", "Commission"])
    sheet.append(["IBM", "2024-07-01", "BUY", 5, 100, 1])
    sheet.append(["IBM", "2024-07-02", "SELL", 5, 110, 1])
    excel_buffer = io.BytesIO()
    workbook.save(excel_buffer)
    excel_exec, excel_notes = load_executions("book.xlsx", excel_buffer.getvalue())
    assert len(excel_exec) == 2, excel_notes
    assert abs(match_fifo(excel_exec).closed.iloc[0]["net_pnl"] - 48) < 1e-6

    partials = pd.DataFrame(
        [
            _exec("SHOP", "2024-08-01 09:30:00", "BUY", 40, 10, 0.4, source_row=1, currency="CAD"),
            _exec("SHOP", "2024-08-01 09:30:04", "BUY", 60, 12, 0.6, source_row=2, currency="CAD"),
            _exec("SHOP", "2024-08-01 09:30:12", "BUY", 10, 13, 0.1, source_row=3, currency="CAD"),
            _exec("SHOP", "2024-08-01 09:31:00", "SELL", 100, 15, 1.0, source_row=4, currency="CAD"),
        ]
    )
    grouped = aggregate_fills(partials)
    assert len(grouped) == 3
    assert abs(grouped.iloc[0]["quantity"] - 100) < 1e-6
    assert abs(grouped.iloc[0]["price"] - 11.2) < 1e-6
    converted, fx_notes = convert_executions(grouped, "USD", 1.36)
    assert converted["currency"].eq("USD").all()
    assert any("USD" in note for note in fx_notes)
    assert abs(converted.iloc[0]["price"] - (11.2 / 1.36)) < 1e-6
    converted_pnl = match_fifo(converted).closed.iloc[0]["net_pnl"]
    assert abs(converted_pnl - (378 / 1.36)) < 1e-4
    back_to_cad, _ = convert_executions(converted, "CAD", 1.36)
    assert abs(back_to_cad.iloc[0]["price"] - 11.2) < 1e-4

    loss_exit = pd.Timestamp("2024-08-02 10:00:00")
    later_loss = pd.Timestamp("2024-08-02 10:01:00")
    reentry = pd.Timestamp("2024-08-02 10:04:00")
    closed_losses = pd.DataFrame(
        [
            {
                "symbol": "AAPL",
                "side": "long",
                "quantity": 50,
                "entry_time": loss_exit - pd.Timedelta(minutes=5),
                "exit_time": loss_exit,
                "entry_price": 10,
                "exit_price": 9,
                "multiplier": 1,
                "gross_pnl": -50,
                "commission": 0,
                "net_pnl": -50,
                "hold_seconds": 300,
                "currency": "USD",
            },
            {
                "symbol": "AAPL",
                "side": "long",
                "quantity": 50,
                "entry_time": later_loss - pd.Timedelta(minutes=5),
                "exit_time": later_loss,
                "entry_price": 10,
                "exit_price": 9,
                "multiplier": 1,
                "gross_pnl": -40,
                "commission": 0,
                "net_pnl": -40,
                "hold_seconds": 300,
                "currency": "USD",
            },
        ]
    )
    reentries = pd.DataFrame(
        [
            {"symbol": "AAPL", "side": "long", "quantity": 100, "entry_time": reentry, "entry_price": 9, "currency": "USD"},
            {
                "symbol": "AAPL",
                "side": "long",
                "quantity": 80,
                "entry_time": reentry + pd.Timedelta(seconds=2),
                "entry_price": 9,
                "currency": "USD",
            },
            {
                "symbol": "AAPL",
                "side": "long",
                "quantity": 90,
                "entry_time": reentry + pd.Timedelta(seconds=8),
                "entry_price": 9,
                "currency": "USD",
            },
        ]
    )
    revenge_rows = revenge_trades(closed_losses, reentries)
    assert len(revenge_rows) == 2
    burst = revenge_rows.iloc[0]
    assert abs(burst["new_quantity"] - 180) < 1e-6
    assert abs(burst["loss_pnl"] - (-40)) < 1e-6
    assert "MIXED" not in format_money(-12.5, "CAD")
    assert format_money(-12.5, "CAD").startswith("-C$")

    flex_header = [
        "Symbol", "Currency", "FX Rate To Base", "Asset Class", "Date/Time", "Buy/Sell", "Quantity",
        "Trade Price", "Proceeds", "IB Commission", "IB Commission Currency", "Taxes", "Multiplier",
        "Level Of Detail", "IB Order ID", "Put/Call", "Strike", "Expiry", "Open/Close Indicator",
        "Order Type", "Order Time", "Notes/Codes", "Trade Date",
    ]
    flex_rows = [
        ["SHOP", "CAD", "0.74", "STK", "2024-08-01, 09:30:00", "BUY", "40", "10", "-400", "-1.00", "CAD", "-9.99", "1", "EXECUTION", "1001", "", "", "", "O", "LMT", "2024-08-01, 09:29:50", "", "2020-01-01"],
        ["SHOP", "CAD", "0.74", "STK", "2024-08-01, 09:30:03", "BUY", "60", "12", "-720", "-1.50", "CAD", "0", "1", "EXECUTION", "9999", "", "", "", "O", "LMT", "2024-08-01, 09:29:50", "", "2020-01-01"],
        ["SHOP", "CAD", "0.74", "STK", "2024-08-01, 09:30:00", "BUY", "100", "11.2", "-1120", "-2.50", "CAD", "-9.99", "1", "ORDER", "1001", "", "", "", "", "", "", "", "2020-01-01"],
        ["SHOP", "CAD", "0.74", "STK", "2024-08-01, 10:00:00", "SELL", "100", "15", "1500", "-1.00", "CAD", "0", "1", "EXECUTION", "1002", "", "", "", "C", "LMT", "2024-08-01, 09:59:00", "A", "2020-01-01"],
        ["AAPL", "USD", "1", "STK", "2024-08-01, 11:00:00", "BUY", "10", "100", "-1000", "-1", "USD", "0", "1", "EXECUTION", "2001", "", "", "", "", "", "", "", "2020-01-01"],
        ["OPT", "USD", "1", "OPT", "2024-08-01, 12:00:00", "BUY", "1", "2.50", "", "-0.65", "USD", "0", "100", "EXECUTION", "3001", "C", "190", "2024-09-20", "O", "LMT", "2024-08-01, 11:59:00", "W", "2020-01-01"],
    ]
    flex_buffer = io.StringIO()
    flex_writer = csv.writer(flex_buffer)
    flex_writer.writerow(flex_header)
    flex_writer.writerows(flex_rows)
    flex = flex_buffer.getvalue()
    flex_exec, flex_notes = load_executions("flex.csv", flex.encode("utf-8"))
    assert len(flex_exec) == 5, flex_notes
    assert set(flex_exec["currency"]) == {"USD"}
    shop_buys = flex_exec[(flex_exec["symbol"] == "SHOP") & (flex_exec["action"] == "BUY")].sort_values("trade_time")
    assert len(shop_buys) == 2
    assert abs(float(shop_buys["price"].iloc[0]) - 7.4) < 1e-6
    assert shop_buys["trade_time"].iloc[0].year == 2024
    assert abs(float(shop_buys["commission"].sum()) - (2.5 * 0.74)) < 1e-6
    option = flex_exec[flex_exec["symbol"] == "OPT"].iloc[0]
    assert abs(float(option["multiplier"]) - 100) < 1e-6
    assert option["put_call"] == "C"
    assert option["strike"] == "190"
    assert option["expiry"] == "2024-09-20"
    assert option["open_close"] == "O"
    assert option["order_type"] == "LMT"
    assert abs(float(option["latency_seconds"]) - 60) < 1e-6
    assert option["trade_time"].hour == 12
    flex_grouped = aggregate_fills(flex_exec)
    assert len(flex_grouped) == 4
    grouped_buy = flex_grouped[(flex_grouped["symbol"] == "SHOP") & (flex_grouped["action"] == "BUY")].iloc[0]
    assert abs(float(grouped_buy["quantity"]) - 100) < 1e-6
    assert abs(float(grouped_buy["price"]) - 8.288) < 1e-6
    shop_closed = match_fifo(flex_grouped[flex_grouped["symbol"] == "SHOP"]).closed
    assert abs(float(shop_closed.iloc[0]["net_pnl"]) - (376.50 * 0.74)) < 1e-4
    assert "A" in str(shop_closed.iloc[0]["notes_codes"])
    flags = event_flags(flex_exec)
    assert set(flags["Flag"]) == {"Wash sale", "Assignment"}
    assert any("FX Rate To Base" in note for note in flex_notes)
    assert any("order" in note for note in flex_notes)

    wide_header = [
        "Symbol", "Date/Time", "Buy/Sell", "Quantity", "Trade Price", "Proceeds",
        "IB Commission", "Currency", "Asset Class",
    ] + [f"Extra {index}" for index in range(77)]
    wide_buffer = io.StringIO()
    wide_writer = csv.writer(wide_buffer)
    wide_writer.writerow(["Account", "Alias", "Model", "From", "To", "Generated", "When", "Basis", "Title"])
    wide_writer.writerow(["U123", "Main", "A", "2024-01-01", "2024-12-31", "2024-12-31", "00:00:00", "USD", "Activity"])
    wide_writer.writerow(["Notes", "only", "nine", "fields", "in", "this", "preamble", "line", "here"])
    wide_writer.writerow(wide_header)
    wide_writer.writerow(
        ["AMD", "2024-09-02, 10:15:00", "BUY", "10", "160", "-1600", "-1.25", "USD", "STK"]
        + [f"ignore-{index}" for index in range(77)]
        + ["trailing-extra"]
    )
    wide_writer.writerow(
        ["AMD", "2024-09-02, 11:15:00", "SELL", "10", "162", "1620", "-1.25", "USD", "STK"]
        + [""] * 77
    )
    wide_exec, wide_notes = load_executions("wide.csv", wide_buffer.getvalue().encode("utf-8"))
    assert len(wide_exec) == 2, wide_notes
    assert wide_exec.iloc[0]["symbol"] == "AMD"
    assert abs(match_fifo(wide_exec).closed.iloc[0]["net_pnl"] - 17.5) < 1e-6

    from ibkr_service import _frame_from_csv, reference_code_from_response, statement_from_response

    assert reference_code_from_response(
        "<FlexStatementResponse><Status>Success</Status><ReferenceCode>999</ReferenceCode></FlexStatementResponse>"
    ) == "999"
    try:
        reference_code_from_response(
            "<FlexStatementResponse><Status>Error</Status><ErrorMessage>Bad token</ErrorMessage></FlexStatementResponse>"
        )
        raise AssertionError("expected a Flex error")
    except FlexServiceError as exc:
        assert "Bad token" in str(exc)
    waiting_csv, waiting_message = statement_from_response(
        "<FlexStatementResponse><Status>Warn</Status><ErrorCode>1019</ErrorCode>"
        "<ErrorMessage>Statement generation in progress. Please try again shortly.</ErrorMessage></FlexStatementResponse>"
    )
    assert waiting_csv is None
    assert "progress" in waiting_message.lower()
    ready_csv, ready_message = statement_from_response(
        "Symbol,Date/Time,Buy/Sell,Quantity,Trade Price,Proceeds,IB Commission,Currency,Asset Class,FX Rate To Base\n"
        'AMD,"2024-09-02, 10:15:00",BUY,10,160,-1600,-1.25,USD,STK,1\n'
    )
    assert ready_message is None and ready_csv.startswith("Symbol")
    ready_frame = _frame_from_csv(ready_csv)
    assert ready_frame["Date/Time"].iloc[0] == pd.Timestamp("2024-09-02 10:15:00")
    assert abs(float(ready_frame["Quantity"].iloc[0]) - 10) < 1e-9
    assert abs(float(ready_frame["Trade Price"].iloc[0]) - 160) < 1e-9
    assert abs(float(ready_frame["IB Commission"].iloc[0]) - -1.25) < 1e-9
    assert abs(float(ready_frame["FX Rate To Base"].iloc[0]) - 1) < 1e-9
    preamble = (
        "Account,Alias,Model,From,To,Generated,When,Basis,Title\n"
        "U1,Main,A,2024-01-01,2024-12-31,2024-12-31,00:00:00,USD,Activity\n"
        "Symbol,Date/Time,Buy/Sell,Quantity,Trade Price,Proceeds,IB Commission,Currency,Asset Class,FX Rate To Base,Multiplier,Open/Close Indicator,Order Type\n"
        'AMD,"2024-09-02, 10:15:00",BUY,10,160,-1600,-1.25,USD,STK,1,1,O,LMT\n'
    )
    preamble_frame = _frame_from_csv(preamble)
    assert list(preamble_frame["Symbol"]) == ["AMD"]
    assert abs(float(preamble_frame["Multiplier"].iloc[0]) - 1) < 1e-9
    assert preamble_frame["Open/Close Indicator"].iloc[0] == "O"
    assert preamble_frame["Order Type"].iloc[0] == "LMT"

    from datetime import date as date_cls
    from datetime import timedelta as delta_cls

    from ibkr_service import _merge_frames, _send_response, year_windows

    windows = year_windows(5, today=date_cls(2026, 10, 3))
    assert windows[-1][1] == date_cls(2026, 10, 2)
    assert all(chunk_start.weekday() < 5 and chunk_end.weekday() < 5 for chunk_start, chunk_end in windows)
    assert all((chunk_end - chunk_start).days <= 365 for chunk_start, chunk_end in windows)
    assert windows
    for (_left_start, left_end), (right_start, _right_end) in zip(windows, windows[1:]):
        gap = right_start - left_end
        assert delta_cls(days=1) <= gap <= delta_cls(days=3)
    leap_windows = year_windows(1, today=date_cls(2024, 2, 29))
    assert leap_windows[0][0] == date_cls(2023, 2, 28)
    assert leap_windows[-1][1] == date_cls(2024, 2, 29)
    calendar = date_windows(date_cls(2024, 1, 1), date_cls(2024, 12, 31))
    assert calendar
    assert all((chunk_end - chunk_start).days <= 365 for chunk_start, chunk_end in calendar)
    assert calendar[0][0] >= date_cls(2024, 1, 1)
    assert calendar[-1][1] <= date_cls(2024, 12, 31)
    assert all(chunk_start.weekday() < 5 and chunk_end.weekday() < 5 for chunk_start, chunk_end in calendar)
    assert date_windows(date_cls(2026, 10, 3), date_cls(2026, 10, 4)) == []
    long_range = date_windows(date_cls(2023, 1, 1), date_cls(2025, 6, 1))
    assert len(long_range) >= 2
    assert planned_flex_sync("2025", None, today=date_cls(2026, 10, 4)) == (date_cls(2025, 1, 1), date_cls(2025, 12, 31))
    assert planned_flex_sync("2026", None, today=date_cls(2026, 10, 4)) == (date_cls(2026, 1, 1), date_cls(2026, 10, 4))
    assert planned_flex_sync("Last 5 years", None, today=date_cls(2026, 10, 4)) == (None, None)
    assert planned_flex_sync(
        "Custom dates",
        (date_cls(2024, 3, 1), date_cls(2024, 6, 15)),
        today=date_cls(2026, 10, 4),
    ) == (date_cls(2024, 3, 1), date_cls(2024, 6, 15))
    try:
        planned_flex_sync("Custom dates", (date_cls(2026, 1, 1),), today=date_cls(2026, 10, 4))
        raise AssertionError("expected an incomplete custom range")
    except ValueError as exc:
        assert "start date" in str(exc)
    assert flex_year_options(date_cls(2026, 10, 4))[0] == "2026"
    assert flex_year_options(date_cls(2026, 10, 4))[-2:] == ["Last 5 years", "Custom dates"]
    span = (date_cls(2021, 10, 14), date_cls(2026, 10, 2))
    assert exit_window_bounds("All dates", *span, today=date_cls(2026, 10, 4)) == span
    assert exit_window_bounds("Year to date", *span, today=date_cls(2026, 10, 4)) == (date_cls(2026, 1, 1), date_cls(2026, 10, 4))
    assert exit_window_bounds("This month", *span, today=date_cls(2026, 10, 4)) == (date_cls(2026, 10, 1), date_cls(2026, 10, 4))
    assert exit_window_bounds("Last month", *span, today=date_cls(2026, 10, 4)) == (date_cls(2026, 9, 1), date_cls(2026, 9, 30))
    assert exit_window_bounds("Last month", *span, today=date_cls(2026, 1, 15)) == (date_cls(2025, 12, 1), date_cls(2025, 12, 31))
    assert exit_window_bounds("This quarter", *span, today=date_cls(2026, 10, 4)) == (date_cls(2026, 10, 1), date_cls(2026, 10, 4))
    assert exit_window_bounds("Last quarter", *span, today=date_cls(2026, 10, 4)) == (date_cls(2026, 7, 1), date_cls(2026, 9, 30))
    assert exit_window_bounds("Last quarter", *span, today=date_cls(2026, 2, 2)) == (date_cls(2025, 10, 1), date_cls(2025, 12, 31))
    assert exit_window_bounds("Last 7 days", *span, today=date_cls(2026, 10, 4)) == (date_cls(2026, 9, 28), date_cls(2026, 10, 4))
    assert exit_window_bounds("Last 30 days", *span, today=date_cls(2026, 10, 4)) == (date_cls(2026, 9, 5), date_cls(2026, 10, 4))
    assert exit_window_bounds("Last 6 months", *span, today=date_cls(2026, 10, 4)) == (date_cls(2026, 4, 4), date_cls(2026, 10, 4))
    assert exit_window_bounds("Last 12 months", *span, today=date_cls(2024, 2, 29)) == (date_cls(2023, 2, 28), date_cls(2024, 2, 29))
    assert exit_window_bounds("Calendar year", *span, year=2024) == (date_cls(2024, 1, 1), date_cls(2024, 12, 31))
    assert exit_window_bounds("Quarter", *span, year=2024, quarter=1) == (date_cls(2024, 1, 1), date_cls(2024, 3, 31))
    assert exit_window_bounds("Month", *span, year=2024, month=2) == (date_cls(2024, 2, 1), date_cls(2024, 2, 29))
    assert exit_window_bounds("Month", *span, year=2023, month=2) == (date_cls(2023, 2, 1), date_cls(2023, 2, 28))
    assert exit_window_bounds("Custom range", *span, custom_start=date_cls(2024, 3, 1), custom_end=date_cls(2024, 6, 15)) == (
        date_cls(2024, 3, 1),
        date_cls(2024, 6, 15),
    )
    try:
        exit_window_bounds("Custom range", *span, custom_start=date_cls(2024, 6, 15), custom_end=date_cls(2024, 3, 1))
        raise AssertionError("expected a reversed custom range")
    except ValueError as exc:
        assert "end date" in str(exc)

    code, statement_url = _send_response(
        "<FlexStatementResponse><Status>Success</Status><ReferenceCode>999</ReferenceCode>"
        "<url>https://gdcdyn.interactivebrokers.com/Universal/servlet/FlexStatementService.GetStatement</url>"
        "</FlexStatementResponse>"
    )
    assert code == "999"
    assert statement_url.startswith("https://gdcdyn.interactivebrokers.com/")
    try:
        _send_response(
            "<FlexStatementResponse><Status>Warn</Status><ErrorCode>1025</ErrorCode>"
            "<ErrorMessage>Too many failed attempts. Please review your configuration.</ErrorMessage>"
            "</FlexStatementResponse>"
        )
        raise AssertionError("expected error 1025")
    except FlexServiceError as exc:
        assert exc.code == "1025"
        assert "15 minutes" in str(exc)

    first = pd.DataFrame(
        [{"Symbol": "AMD", "Date/Time": "2024-09-02, 10:15:00", "Buy/Sell": "BUY", "Quantity": 10, "Trade Price": 160, "Proceeds": -1600, "IB Commission": -1.25}]
    )
    second = pd.DataFrame(
        [
            {"Symbol": "AMD", "Date/Time": "2024-09-02 10:15:00", "Buy/Sell": "BUY", "Quantity": 10, "Trade Price": 160, "Proceeds": -1600, "IB Commission": -1.25},
            {"Symbol": "AMD", "Date/Time": "2024-09-03, 11:00:00", "Buy/Sell": "SELL", "Quantity": 10, "Trade Price": 162, "Proceeds": 1620, "IB Commission": -1.25},
        ]
    )
    merged = _merge_frames([first, second])
    assert len(merged) == 2
    assert merged["Date/Time"].is_monotonic_increasing
    assert list(merged["Buy/Sell"]) == ["BUY", "SELL"]

    round_trips = match_fifo(
        pd.DataFrame(
            [
                _exec("AAPL", "2024-01-02 09:30", "BUY", 10, 100, source_row=1),
                _exec("AAPL", "2024-01-02 10:30", "SELL", 10, 110, source_row=2),
                _exec("MSFT", "2024-01-03 09:30", "SELL", 5, 50, source_row=3),
                _exec("MSFT", "2024-01-03 15:30", "BUY", 5, 40, source_row=4),
            ]
        )
    )
    round_summary = summarize(round_trips.closed, 0)
    assert abs(round_summary.net_pnl - 150) < 1e-6
    assert abs(round_summary.expectancy - 75) < 1e-6
    dipped = match_fifo(
        pd.DataFrame(
            [
                _exec("AAPL", "2024-01-02 09:30", "BUY", 10, 100, source_row=1),
                _exec("AAPL", "2024-01-02 10:30", "SELL", 10, 110, source_row=2),
                _exec("AAPL", "2024-01-03 09:30", "BUY", 10, 100, source_row=3),
                _exec("AAPL", "2024-01-03 10:30", "SELL", 10, 75, source_row=4),
            ]
        )
    )
    drawdown = underwater_series(dipped.closed)
    assert abs(float(drawdown["underwater"].min()) - summarize(dipped.closed, 0).max_drawdown) < 1e-6
    dipped.closed["asset_category"] = ["STK", "OPT"]
    buckets = group_breakdown(dipped.closed, dipped.closed["asset_category"].map(_asset_bucket), ["Stocks", "Options"])
    assert list(buckets["Group"]) == ["Stocks", "Options"]
    hours = group_breakdown(round_trips.closed, round_trips.closed["entry_time"].dt.hour.map(lambda hour: f"{int(hour):02d}:00"))
    assert set(hours["Group"]) == {"09:00"}

    import tempfile

    with tempfile.TemporaryDirectory() as folder:
        folder_path = Path(folder)
        save_flex_sync(
            "Symbol,Date/Time\nAMD,2024-01-01\n",
            [{"start": "2021-10-04", "end": "2022-10-04", "csv": "Symbol\nAMD\n"}],
            folder_path,
        )
        assert (folder_path / "combined.csv").is_file()
        assert (folder_path / "2021-10-04_to_2022-10-04.csv").is_file()
        save_flex_sync(
            "Symbol,Date/Time\nNVDA,2024-01-02\n",
            [{"start": "2022-10-05", "end": "2023-10-05", "csv": "Symbol\nNVDA\n"}],
            folder_path,
        )
        assert not (folder_path / "2021-10-04_to_2022-10-04.csv").exists()
        assert "NVDA" in (folder_path / "combined.csv").read_text(encoding="utf-8")
        assert (folder_path / "2022-10-05_to_2023-10-05.csv").is_file()

    ticker_rows = pd.DataFrame(
        [
            {"symbol": "BIG", "net_pnl": 1000.0, "commission": 1.0},
            {"symbol": "BIG", "net_pnl": 500.0, "commission": 1.0},
            {"symbol": "SMALL", "net_pnl": 10.0, "commission": 2.0},
            {"symbol": "LOSS", "net_pnl": -800.0, "commission": 3.0},
            {"symbol": "FLAT", "net_pnl": 0.0, "commission": 0.0},
        ]
    )
    by_symbol = ticker_totals(ticker_rows)
    assert list(by_symbol["Symbol"]) == ["BIG", "LOSS", "SMALL", "FLAT"]
    assert int(by_symbol.iloc[0]["Trade Count"]) == 2
    assert abs(float(by_symbol.iloc[0]["Win Rate (%)"]) - 100) < 1e-6
    assert abs(float(by_symbol.iloc[0]["Total Commissions"]) - 2) < 1e-6
    assert abs(float(by_symbol.loc[by_symbol["Symbol"] == "LOSS", "Win Rate (%)"].iloc[0])) < 1e-6
    assert pd.isna(by_symbol.loc[by_symbol["Symbol"] == "FLAT", "Win Rate (%)"].iloc[0])
    assert _signed_money(28500, "USD") == "+$28,500.00"
    assert _signed_money(-16200, "USD") == "-$16,200.00"
    print("self-test ok")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        _self_test()
    else:
        main()
