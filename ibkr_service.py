"""Download an Interactive Brokers Flex Query statement as a table.

IBKR accepts at most 365 days on one SendRequest. A longer history is loaded
as consecutive windows, then concatenated.
"""

from __future__ import annotations

import csv
import io
import os
import time
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from collections.abc import Callable
from datetime import date, timedelta
from pathlib import Path

import pandas as pd


def hosted_publicly() -> bool:
    """True on Streamlit Community Cloud, where this app must not call IBKR."""
    if os.environ.get("STREAMLIT_RUNTIME_ENV", "").lower() == "cloud":
        return True
    if Path("/mount/src").exists():
        return True
    home = os.environ.get("HOME", "").replace("\\", "/").rstrip("/")
    return home == "/home/adminuser"


def flex_token() -> str:
    """Read the Flex token from the environment, a local file, or Streamlit secrets."""
    env = os.environ.get("IBKR_FLEX_TOKEN", "").strip()
    if env:
        return env
    path = Path(__file__).resolve().parent / ".ibkr_token"
    if path.is_file():
        return path.read_text(encoding="utf-8").strip()
    try:
        import streamlit as st

        return str(st.secrets.get("IBKR_FLEX_TOKEN", "")).strip()
    except Exception:
        return ""


DEFAULT_TOKEN = flex_token()
DEFAULT_QUERY_ID = "1659396"
DEFAULT_YEARS = 5

SEND_URL = "https://ndcdyn.interactivebrokers.com/AccountManagement/FlexWebService/SendRequest"
GET_URL = "https://ndcdyn.interactivebrokers.com/AccountManagement/FlexWebService/GetStatement"
POLL_ATTEMPTS = 10
POLL_SECONDS = 3
REQUEST_TIMEOUT = 90
# IBKR allows a from/to span of at most 365 days on one request.
MAX_CHUNK_DAYS = 365

NUMERIC_COLUMNS = ["Quantity", "Trade Price", "Proceeds", "IB Commission", "FX Rate To Base"]
IDENTITY_COLUMNS = ["Symbol", "Date/Time", "Buy/Sell", "Quantity", "Trade Price", "Proceeds", "IB Commission"]

ProgressCallback = Callable[[int, int, date, date], None]
_recent_requests: list[float] = []


class FlexServiceError(Exception):
    """The Flex Web Service returned an error or never produced a statement."""

    def __init__(self, message: str, code: str = ""):
        super().__init__(message)
        self.code = code


# These codes mean "not ready yet" or "slow down", from IBKR's Flex error list.
RETRYABLE_CODES = {"1001", "1004", "1005", "1006", "1007", "1008", "1009", "1018", "1019", "1021"}


def fetch_ibkr_trades(
    token: str = DEFAULT_TOKEN,
    query_id: str = DEFAULT_QUERY_ID,
    years: int = DEFAULT_YEARS,
    progress: ProgressCallback | None = None,
    start: date | None = None,
    end: date | None = None,
) -> pd.DataFrame:
    """Download trades and return one DataFrame.

    Pass ``start`` and ``end`` to sync that inclusive range. Otherwise the last
    ``years`` are downloaded. IBKR accepts at most 365 days per request, so a
    longer range is split into weekday windows. The merged CSV is kept on
    ``frame.attrs["source_csv"]`` so the app can run it through the same cleaner
    used for an uploaded file.
    """
    resolved = (token or flex_token()).strip()
    if hosted_publicly():
        raise FlexServiceError("IBKR sync is turned off on the public site.")
    if not resolved:
        raise FlexServiceError("Set IBKR_FLEX_TOKEN, or save the token in .ibkr_token, before syncing.")
    if start is not None or end is not None:
        if start is None or end is None:
            raise FlexServiceError("A Flex date override needs both a start date and an end date.")
        return fetch_date_range(resolved, query_id, start, end, progress=progress)
    return fetch_multi_year_history(resolved, query_id, years=years, progress=progress)


def fetch_single_flex_statement(
    token: str,
    query_id: str,
    start: date | None = None,
    end: date | None = None,
) -> pd.DataFrame:
    """Execute one Flex request. Pass both dates to override the query's saved range."""
    csv_text = _download_statement(token.strip(), str(query_id).strip(), start, end)
    frame = _frame_from_csv(csv_text)
    frame.attrs["source_csv"] = csv_text
    return frame


def fetch_multi_year_history(
    token: str,
    query_id: str,
    years: int = DEFAULT_YEARS,
    progress: ProgressCallback | None = None,
) -> pd.DataFrame:
    """Fetch history in annual chunks and return one de-duplicated DataFrame."""
    return _fetch_windows(token, query_id, year_windows(years), progress)


def fetch_date_range(
    token: str,
    query_id: str,
    start: date,
    end: date,
    progress: ProgressCallback | None = None,
) -> pd.DataFrame:
    """Fetch one inclusive date range, split into 365-day windows when needed."""
    if end < start:
        raise FlexServiceError("The end date must be on or after the start date.")
    windows = date_windows(start, end)
    if not windows:
        raise FlexServiceError("That date range has no weekdays. IBKR rejects a weekend-only request.")
    return _fetch_windows(token, query_id, windows, progress)


def _fetch_windows(
    token: str,
    query_id: str,
    windows: list[tuple[date, date]],
    progress: ProgressCallback | None = None,
) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    chunks: list[dict[str, str]] = []
    warnings: list[str] = []
    for index, (start, end) in enumerate(windows, start=1):
        if progress is not None:
            progress(index, len(windows), start, end)
        frame = None
        for attempt in range(2):
            try:
                frame = fetch_single_flex_statement(token, query_id, start, end)
                break
            except FlexServiceError as exc:
                if exc.code == "1025" or _is_fatal(str(exc)):
                    raise
                if attempt == 0 and exc.code in RETRYABLE_CODES:
                    time.sleep(15)
                    continue
                warnings.append(f"{start.isoformat()} to {end.isoformat()}: {exc}")
                frame = None
                break
        if frame is not None:
            chunks.append(
                {
                    "start": start.isoformat(),
                    "end": end.isoformat(),
                    "csv": frame.attrs.get("source_csv", ""),
                }
            )
            if not frame.empty:
                frames.append(frame)

    if not frames:
        detail = "; ".join(warnings) if warnings else "IBKR returned no trades."
        raise FlexServiceError(detail)

    merged = _merge_frames(frames)
    merged.attrs["source_csv"] = _combine_statement_text(chunks)
    merged.attrs["chunks"] = chunks
    merged.attrs["chunk_warnings"] = warnings
    merged.attrs["chunk_count"] = len(windows)
    return merged


def year_windows(years: int = DEFAULT_YEARS, today: date | None = None) -> list[tuple[date, date]]:
    """Split ``today`` back ``years`` into windows of at most 365 days.

    IBKR has rejected ranges that end on Saturday or Sunday with error 1025,
    so each window starts and ends on a weekday.
    """
    if years < 1:
        raise ValueError("years must be at least 1")
    end = _previous_weekday(today or date.today())
    try:
        start = end.replace(year=end.year - years)
    except ValueError:
        start = end.replace(year=end.year - years, day=28)
    return date_windows(start, end)


def date_windows(start: date, end: date) -> list[tuple[date, date]]:
    """Split an inclusive range into weekday windows of at most 365 days.

    A Saturday or Sunday endpoint is moved to the nearest weekday inside the
    range, because IBKR has rejected weekend endpoints with error 1025.
    """
    start = _next_weekday(start)
    end = _previous_weekday(end)
    if start > end:
        return []
    windows: list[tuple[date, date]] = []
    cursor = start
    while cursor <= end:
        chunk_end = _previous_weekday(min(cursor + timedelta(days=MAX_CHUNK_DAYS), end))
        if chunk_end < cursor:
            break
        windows.append((cursor, chunk_end))
        cursor = _next_weekday(chunk_end + timedelta(days=1))
    return windows


def _next_weekday(day: date) -> date:
    while day.weekday() >= 5:
        day += timedelta(days=1)
    return day


def _previous_weekday(day: date) -> date:
    while day.weekday() >= 5:
        day -= timedelta(days=1)
    return day


def _download_statement(
    token: str,
    query_id: str,
    start: date | None = None,
    end: date | None = None,
) -> str:
    if not token or not query_id:
        raise FlexServiceError("A Flex token and query ID are required.")
    if (start is None) != (end is None):
        raise FlexServiceError("A Flex date override needs both a start date and an end date.")
    if start is not None and end is not None and (end - start).days > 365:
        raise FlexServiceError("IBKR allows at most 365 days in one Flex request.")

    params = {"t": token, "q": query_id, "v": "3"}
    if start is not None and end is not None:
        params["fd"] = start.strftime("%Y%m%d")
        params["td"] = end.strftime("%Y%m%d")

    request_xml = _http_get(SEND_URL, params)
    reference_code, statement_url = _send_response(request_xml)

    last_waiting_message = "IBKR has not finished generating the statement."
    for attempt in range(POLL_ATTEMPTS):
        if attempt:
            time.sleep(POLL_SECONDS)
        payload = _http_get(statement_url, {"q": reference_code, "t": token, "v": "3"})
        csv_text, waiting_message = statement_from_response(payload)
        if csv_text is not None:
            return csv_text
        if waiting_message:
            last_waiting_message = waiting_message

    raise FlexServiceError(
        f"{last_waiting_message} Tried {POLL_ATTEMPTS} times. Sync again in a minute."
    )


def reference_code_from_response(xml_text: str) -> str:
    code, _statement_url = _send_response(xml_text)
    return code


def _send_response(xml_text: str) -> tuple[str, str]:
    root = _xml_root(xml_text)
    status = _tag_text(root, "Status").lower()
    if status == "success":
        code = _tag_text(root, "ReferenceCode")
        if not code:
            raise FlexServiceError("IBKR accepted the request but did not return a reference code.")
        raw_url = _tag_text(root, "url")
        statement_url = raw_url.split("?")[0] if raw_url.startswith("http") else GET_URL
        return code, statement_url
    _raise_response_error(root)


def _merge_frames(frames: list[pd.DataFrame]) -> pd.DataFrame:
    combined = pd.concat(frames, ignore_index=True)
    if "Date/Time" in combined.columns:
        cleaned = combined["Date/Time"].fillna("").astype(str).str.replace(",", " ", regex=False)
        combined["Date/Time"] = pd.to_datetime(cleaned, errors="coerce", format="mixed")
    for column in NUMERIC_COLUMNS:
        if column in combined.columns:
            combined[column] = pd.to_numeric(combined[column], errors="coerce")
    keys = [column for column in IDENTITY_COLUMNS if column in combined.columns]
    if keys:
        combined = combined.drop_duplicates(subset=keys, keep="first")
    else:
        combined = combined.drop_duplicates()
    if "Date/Time" in combined.columns:
        combined = combined.sort_values("Date/Time", kind="mergesort")
    return combined.reset_index(drop=True)


def _combine_statement_text(chunks: list[dict]) -> str:
    """Keep each downloaded statement, including sections other than trades."""
    parts = [str(chunk.get("csv") or "").strip() for chunk in chunks]
    parts = [part for part in parts if part]
    if not parts:
        return ""
    return "\n".join(parts) + "\n"


def _frame_to_csv(frame: pd.DataFrame) -> str:
    export = frame.copy()
    if "Date/Time" in export.columns:
        export["Date/Time"] = [
            "" if pd.isna(value) else pd.Timestamp(value).strftime("%Y-%m-%d, %H:%M:%S")
            for value in export["Date/Time"]
        ]
    buffer = io.StringIO()
    export.to_csv(buffer, index=False)
    return buffer.getvalue()


def _is_fatal(message: str) -> bool:
    lowered = message.lower()
    return "expired" in lowered or "query id" in lowered or "token is invalid" in lowered or "token has expired" in lowered


def _lockout_message() -> str:
    return (
        "IBKR locked this Flex token after too many failed attempts (error 1025). "
        "Wait about 15 minutes, then click Sync once."
    )


def _raise_response_error(root: ET.Element) -> None:
    code = _tag_text(root, "ErrorCode")
    message = _tag_text(root, "ErrorMessage") or "Flex request failed."
    if code == "1025":
        raise FlexServiceError(_lockout_message(), code="1025")
    raise FlexServiceError(message, code=code)


def _flex_tag(value: object) -> str:
    return str(value or "").strip().strip('"').upper()


def split_flex_tables(text: str, delimiter: str) -> list[tuple[str, list[dict]]]:
    """Split an IBKR Flex CSV envelope (BOF/BOS/EOS) into named tables."""
    rows = list(csv.reader(io.StringIO(text), delimiter=delimiter))
    if not any(row and _flex_tag(row[0]) in {"BOF", "BOS"} for row in rows[:40]):
        return []
    tables: list[tuple[str, list[list[str]]]] = []
    section_name = ""
    bucket: list[list[str]] | None = None
    loose: list[list[str]] = []

    def flush_loose() -> None:
        nonlocal loose
        if loose:
            tables.append(("", loose))
            loose = []

    def flush_section() -> None:
        nonlocal bucket, section_name
        if bucket:
            tables.append((section_name, bucket))
        bucket = None
        section_name = ""

    for row in rows:
        if not row or not any(str(cell).strip() for cell in row):
            continue
        tag = _flex_tag(row[0])
        if tag in {"BOF", "EOF", "BOA", "EOA"}:
            flush_loose()
            continue
        if tag == "BOS":
            flush_loose()
            flush_section()
            section_name = row[2].strip() if len(row) > 2 else ""
            bucket = []
            continue
        if tag == "EOS":
            flush_section()
            continue
        if bucket is not None:
            bucket.append([str(cell).strip() for cell in row])
        else:
            loose.append([str(cell).strip() for cell in row])
    flush_section()
    flush_loose()

    parsed: list[tuple[str, list[dict]]] = []
    for name, body in tables:
        if len(body) < 2:
            continue
        header = body[0]
        records = []
        for values in body[1:]:
            if not values or _flex_tag(values[0]) in {"BOF", "EOF", "BOA", "EOA", "BOS", "EOS"}:
                continue
            padded = values + [""] * (len(header) - len(values))
            records.append(dict(zip(header, padded[: len(header)])))
        if records:
            parsed.append((name, records))
    return parsed


def _flex_trade_frame(csv_text: str) -> pd.DataFrame | None:
    """Trade rows from a Flex envelope. Kept here so Sync does not import app.py."""
    sample = "\n".join(csv_text.splitlines()[:40])
    counts = {",": sample.count(","), ";": sample.count(";"), "\t": sample.count("\t")}
    delimiter = max(counts, key=counts.get)
    frames = []
    for _name, records in split_flex_tables(csv_text, delimiter):
        if not records:
            continue
        keys = {"".join(ch for ch in str(key).lower() if ch.isalnum()) for key in records[0]}
        if "symbol" not in keys:
            continue
        frames.append(pd.DataFrame.from_records(records))
    if not frames:
        return None
    return pd.concat(frames, ignore_index=True)


def statement_from_response(payload: str) -> tuple[str | None, str | None]:
    """Return ``(csv_text, None)`` when ready, or ``(None, message)`` while IBKR is still generating."""
    text = payload.lstrip("\ufeff").lstrip()
    if not text.startswith("<"):
        if not text:
            raise FlexServiceError("IBKR returned an empty statement.")
        return text, None

    root = _xml_root(text)
    code = _tag_text(root, "ErrorCode")
    message = _tag_text(root, "ErrorMessage")
    if code == "1025":
        raise FlexServiceError(_lockout_message(), code="1025")
    if code in RETRYABLE_CODES or "in progress" in message.lower():
        return None, message or "Statement generation is still in progress."
    if code or message:
        raise FlexServiceError(message or "Flex statement request failed.", code=code)
    raise FlexServiceError("IBKR returned XML instead of the CSV statement. Set the Flex Query output format to CSV.")


def _loaded_parser():
    """Use the already-running app module. Importing app.py here re-enters Sync and fails."""
    import sys

    for name in ("__main__", "app"):
        module = sys.modules.get(name)
        if module is not None and hasattr(module, "read_flat_frame") and hasattr(module, "_choose_delimiter"):
            return module
    return None


def _frame_from_csv(csv_text: str) -> pd.DataFrame:
    trade_frame = _flex_trade_frame(csv_text)
    if trade_frame is not None and not trade_frame.empty:
        return trade_frame
    parser = _loaded_parser()
    frame = None
    if parser is not None:
        try:
            frame = parser.read_flat_frame(csv_text, parser._choose_delimiter(csv_text))
        except Exception:
            frame = None
    if frame is None or frame.empty:
        try:
            frame = pd.read_csv(io.StringIO(csv_text), dtype=str, engine="python", on_bad_lines="skip")
        except Exception:
            return pd.DataFrame()
    frame.columns = [str(column).strip() for column in frame.columns]
    if "Date/Time" in frame.columns:
        cleaned = frame["Date/Time"].fillna("").astype(str).str.replace(",", " ", regex=False)
        frame["Date/Time"] = pd.to_datetime(cleaned, errors="coerce", format="mixed")
        frame = frame.sort_values("Date/Time", kind="mergesort").reset_index(drop=True)
    for column in NUMERIC_COLUMNS:
        if column in frame.columns:
            frame[column] = pd.to_numeric(
                frame[column].astype(str).str.replace(",", "", regex=False).str.replace("$", "", regex=False),
                errors="coerce",
            )
    return frame


def _http_get(url: str, params: dict[str, str]) -> str:
    _pace()
    target = url + "?" + urllib.parse.urlencode(params)
    request = urllib.request.Request(target, headers={"User-Agent": "Python/3.14"})
    try:
        with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT) as response:
            raw = response.read()
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        message = _error_message_from_text(detail) or f"IBKR returned HTTP {exc.code}."
        raise FlexServiceError(message) from exc
    except urllib.error.URLError as exc:
        raise FlexServiceError(f"Could not reach IBKR. {exc.reason}") from exc
    return raw.decode("utf-8-sig", errors="replace")


def _xml_root(xml_text: str) -> ET.Element:
    try:
        return ET.fromstring(xml_text.strip())
    except ET.ParseError as exc:
        raise FlexServiceError("IBKR returned a response that is not valid XML or CSV.") from exc


def _tag_text(root: ET.Element, tag: str) -> str:
    wanted = tag.lower()
    for element in root.iter():
        if element.tag.split("}")[-1].lower() == wanted and element.text and element.text.strip():
            return element.text.strip()
    return ""


def _pace() -> None:
    """Stay inside IBKR's Flex limit of 1 request per second and 10 per minute."""
    now = time.monotonic()
    if _recent_requests:
        gap = now - _recent_requests[-1]
        if gap < 1.1:
            time.sleep(1.1 - gap)
    cutoff = time.monotonic() - 60
    recent = [stamp for stamp in _recent_requests if stamp >= cutoff]
    if len(recent) >= 8:
        wait = 60 - (time.monotonic() - recent[0]) + 0.3
        if wait > 0:
            time.sleep(wait)
    _recent_requests.append(time.monotonic())
    expired = time.monotonic() - 120
    while _recent_requests and _recent_requests[0] < expired:
        _recent_requests.pop(0)


def _error_message_from_text(payload: str) -> str:
    text = payload.strip()
    if not text.startswith("<"):
        return text[:300]
    try:
        root = ET.fromstring(text)
    except ET.ParseError:
        return ""
    return _tag_text(root, "ErrorMessage")
