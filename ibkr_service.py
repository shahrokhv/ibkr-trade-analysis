"""Download an Interactive Brokers Flex Query statement as a table.

IBKR accepts at most 365 days on one SendRequest. A longer history is loaded
as consecutive windows, then concatenated.
"""

from __future__ import annotations

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
) -> pd.DataFrame:
    """Download ``years`` of trades in 365-day windows and return one DataFrame.

    The merged CSV is kept on ``frame.attrs["source_csv"]`` so the app can run
    it through the same cleaner used for an uploaded file.
    """
    resolved = (token or flex_token()).strip()
    if not resolved:
        raise FlexServiceError("Set IBKR_FLEX_TOKEN, or save the token in .ibkr_token, before syncing.")
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
    windows = year_windows(years)
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
    merged.attrs["source_csv"] = _frame_to_csv(merged)
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
    start = _next_weekday(start)
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


def _frame_from_csv(csv_text: str) -> pd.DataFrame:
    # Imported here so this module can load before app.py finishes importing it.
    from app import _choose_delimiter, read_flat_frame

    try:
        frame = read_flat_frame(csv_text, _choose_delimiter(csv_text))
    except Exception:
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
