from __future__ import annotations

import atexit
import argparse
import ctypes
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from copy import copy, deepcopy
from dataclasses import dataclass
import csv
from datetime import date, datetime, timedelta
import html
from html.parser import HTMLParser
import io
import json
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import threading
import time
import sys
import tempfile
from typing import Callable
from ctypes import wintypes
from http.cookiejar import Cookie, CookieJar
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, unquote, urljoin, urlparse
from urllib.request import HTTPCookieProcessor, Request, build_opener, urlopen
import xml.etree.ElementTree as ET
import zipfile
from urllib3.exceptions import ReadTimeoutError

from openpyxl import load_workbook
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint, Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.drawing.line import LineProperties
from pypdf import PdfReader
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


TARGET_URL = "https://www.aptldata.ford.com/cgi-bin/entra-query"
SUCCESS_TITLE = "Emissions Utilities"
DEFAULT_CLICK_TARGET = "1.1.1"
DEFAULT_VEHICLE_VIN = "590W960"
DEFAULT_WORKBOOK_TEMPLATE = "Emission summary_automation_templete_2.xlsx"
DEFAULT_RESULT_DIR = "Emission_Summary_Result"
DEFAULT_PDF_DIR = "Emission_Result"
DEFAULT_DOWNLOAD_DIR = "Imported_spreadsheet"
DEFAULT_ATI_DIR = "Imported_spreadsheet/APTL_MF4_download"
DEFAULT_DOWNLOAD_WAIT_SECONDS = 120
DEFAULT_DRIVER_DIR = ".drivers"
DEFAULT_LOG_FILE = "last_run.log"
DEFAULT_STATUS_TEXT_FILE = "run_status.txt"
DEFAULT_STATUS_HTML_FILE = "run_status.html"
DEFAULT_STATUS_JSON_FILE = "run_status.json"
DEFAULT_STATUS_SHELL_FILE = "run_status.hta"
DEFAULT_STATUS_FRAGMENT_FILE = "run_status_fragment.html"
DEFAULT_STATUS_PROFILE_DIR = ".status_browser"
DEFAULT_AUTOMATION_PROFILE_DIR = ".automation_browser"
DEFAULT_STATUS_PID_FILE = "run_status.pid"
DEFAULT_AUTOMATION_PID_FILE = "automation_browser.pid"
DEFAULT_WORKBOOK_OPEN_PID_FILE = "workbook_open.pid"
DEFAULT_RUN_TOKEN_FILE = "current_run.token"
DEFAULT_WINDOW_STATE_FILE = "window_state.json"
DEFAULT_QUERY_SETTINGS_FILE = "query_settings.json"
DEFAULT_STATUS_HOST = "127.0.0.1"
DEFAULT_PLOT_BASE_SHEET = "Emis_Plot_Base"
DEFAULT_PLOT_OBD_SHEET = "Emis_Plot_OBD"
DEFAULT_CATMN_SHEET = "CATMN"
DEFAULT_MONITOR_SOURCE_FILE = "1.0LFoxEU7_ROD_Schedule_Emissions.xlsx"
DEFAULT_MONITOR_SOURCE_LIVE_FILE = "1.0LFoxEU7_ROD_Schedule_Emissions.live.xlsx"
DEFAULT_MONITOR_SOURCE_CLIPBOARD_FILE = "1.0LFoxEU7_ROD_Schedule_Emissions.clipboard.tsv"
DEFAULT_MONITOR_SOURCE_URL = "https://azureford.sharepoint.com/:x:/r/sites/FoxUpgradeEU7Cal/_layouts/15/Doc.aspx?sourcedoc=%7B0901622E-87DE-4C0F-B60A-2B550F7A3DB8%7D&file=1.0LFoxEU7_ROD_Schedule_Emissions.xlsx&action=default&mobileredirect=true&share=cQouYgEJ3ocPTLYKK1UPej24EgUCfoh26uFICj7UwxoI3f-3nQ"
AUTOMATION_WINDOW_SIZE = (1280, 900)
AUTOMATION_WINDOW_POSITION = (620, 40)
STATUS_WINDOW_POSITION = (40, 40)
STATUS_WINDOW_SIZE = (560, 920)
LINKED_DOCUMENT_START_COLUMN = 27
LINKED_DOCUMENT_END_COLUMN = 47

SUMMARY_START_COLUMN = "AZ"
SUMMARY_VALUE_HEADERS = {
    "date": "Date",
    "test_id": "Test ID",
    "monitor": "Monitor",
    "thc": "THC",
    "thc_pct": "THC %",
    "co": "CO",
    "co_pct": "CO %",
    "nox": "NOX",
    "nox_pct": "NOX %",
    "co2": "CO2",
    "ch4": "CH4",
    "nmhc": "NM-HC",
    "nmhc_pct": "NM-HC %",
    "hc_nox": "HC+NOx",
    "soot": "SOOT",
    "fe": "FE",
    "pm": "PM",
    "pm_pct": "PM %",
    "cal_file": "cal File",
    "comments": "Comments",
    "void": "VOID",
}

PERCENT_ALERT_KEYS = {"co_pct", "nox_pct", "nmhc_pct"}
PERCENT_ALERT_FILL = PatternFill(fill_type="solid", start_color="FFFF0000", end_color="FFFF0000")
PERCENT_WARNING_FILL = PatternFill(fill_type="solid", start_color="FFFFFF00", end_color="FFFFFF00")
DEFAULT_FILL = PatternFill()

PLOT_SERIES_COLORS = {
    "FCD-Robot": "70AD47",
    "Overnight Soak-Driver": "4472C4",
    "Overnight Soak-Robot": "ED7D31",
}

CURRENT_STATUS_STAGE = "Starting"
CURRENT_STATUS_DETAIL = "Initializing automation."
CURRENT_STATUS_STATE = "running"
CURRENT_STATUS_PROCESS_INDEX: int | None = None
CURRENT_RUN_TOKEN = ""
STATUS_SERVER_PORT: int | None = None
STATUS_SERVER: ThreadingHTTPServer | None = None
STATUS_SERVER_THREAD: threading.Thread | None = None
STATUS_INPUT_REQUEST: dict[str, str] | None = None
STATUS_INPUT_RESPONSE: dict[str, str] | None = None
STATUS_INPUT_EVENT = threading.Event()
STATUS_INPUT_LOCK = threading.Lock()
STATUS_INPUT_TIMEOUT_SECONDS = 30
LAST_UPDATE_MODE_SELECTION: str | None = None
CURRENT_STATUS_WINDOW_BOUNDS = {
    "left": STATUS_WINDOW_POSITION[0],
    "top": STATUS_WINDOW_POSITION[1],
    "width": STATUS_WINDOW_SIZE[0],
    "height": STATUS_WINDOW_SIZE[1],
}
CURRENT_AUTOMATION_WINDOW_BOUNDS = {
    "left": AUTOMATION_WINDOW_POSITION[0],
    "top": AUTOMATION_WINDOW_POSITION[1],
    "width": AUTOMATION_WINDOW_SIZE[0],
    "height": AUTOMATION_WINDOW_SIZE[1],
}
CURRENT_WORKBOOK_WINDOW_BOUNDS = {
    "left": AUTOMATION_WINDOW_POSITION[0],
    "top": AUTOMATION_WINDOW_POSITION[1],
    "width": AUTOMATION_WINDOW_SIZE[0],
    "height": AUTOMATION_WINDOW_SIZE[1],
}

QUERY_SETTING_FIELDS = [
    ("vehicle_vin", "VIN"),
    ("test_requestor", "Test Requestor"),
    ("start_date", "Start Date"),
    ("sort_by", "Sort By"),
]

RUN_MODE_OPTIONS = {
    "fresh": "Start from template",
    "append": "Append to existing summary",
}

STATUS_PROCESS_STEPS = [
    ("Start And Settings", {"Run Started", "Startup Cleanup"}),
    ("Browser Startup", {"Browser Startup", "Browser Startup Failed"}),
    (
        "Access And Sign-In",
        {
            "Navigating",
            "Waiting For Access",
            "Waiting For Microsoft Sign-In",
            "Proceed Clicked",
            "Access Confirmed",
            "Navigation Ready",
        },
    ),
    (
        "Query Setup",
        {"Menu Selected", "Query Setup", "VIN Entered", "Test Requestor Entered", "Start Date Entered", "End Date Entered", "Sort Selected", "Query Submitted"},
    ),
    ("Spreadsheet Download", {"Import Requested", "Spreadsheet Downloaded"}),
    (
        "ATI And Monitor Downloads",
        {
            "ATI Download Started",
            "ATI Downloaded",
            "ATI Download Complete",
            "ATI Download Skipped",
            "ATI Cache Restored",
            "ATI Files Renamed",
            "ATI Monitor Matched",
            "ATI Missing Files",
            "Monitor Schedule Download Started",
            "Monitor Schedule Downloaded",
            "Monitor Schedule Download Unavailable",
        },
    ),
    ("Compare And Choose", {"Spreadsheet Compared"}),
    ("Workbook Ready", {"Workbook Selected", "Workbook Selection Failed", "Workbook Locked"}),
    ("Spreadsheet Import", {"Spreadsheet Imported"}),
    (
        "PDF Download",
        {
            "Preparing PDF Download",
            "PDF Download Started",
            "PDF Downloaded",
            "PDF Download Finished",
            "PDF Download Skipped",
        },
    ),
    (
        "Workbook And Charts",
        {
            "Workbook Population Started",
            "Workbook Population Complete",
            "Workbook Population Failed",
            "Building Emis_Plot_Base and Emis_Plot_OBD",
            "Workbook Enriched",
            "Workbook Update Complete",
        },
    ),
    (
        "Monitor Update",
        {
            "Monitor Update Started",
            "Monitor Update Complete",
            "Monitor Update Skipped",
        },
    ),
    ("Finish And Open Excel", {"Opening Excel", "Run Complete"}),
]


def _get_download_directory() -> Path:
    download_dir = Path(DEFAULT_DOWNLOAD_DIR)
    download_dir.mkdir(parents=True, exist_ok=True)
    return download_dir.resolve()


def _snapshot_downloads(download_dir: Path) -> set[str]:
    return {path.name for path in download_dir.iterdir() if path.is_file()}


def _clear_previous_imported_spreadsheets(download_dir: Path) -> int:
    removed = 0
    for path in download_dir.iterdir():
        if not path.is_file():
            continue
        if path.suffix.lower() not in {".xlsx", ".xlsm", ".xls", ".csv"}:
            continue
        try:
            path.unlink()
            removed += 1
        except OSError as exc:
            _append_log(
                f"Skipped deleting previous imported spreadsheet in {download_dir}: {path.name} ({exc.__class__.__name__}: {exc})"
            )
    return removed


def _get_automation_profile_directory(browser: str) -> Path:
    profile_dir = Path(DEFAULT_AUTOMATION_PROFILE_DIR) / browser.lower()
    profile_dir.mkdir(parents=True, exist_ok=True)
    return profile_dir.resolve()


def _get_ephemeral_automation_profile_directory(browser: str) -> Path:
    runtime_root = Path(DEFAULT_AUTOMATION_PROFILE_DIR) / "runtime"
    runtime_root.mkdir(parents=True, exist_ok=True)
    return Path(tempfile.mkdtemp(prefix=f"{browser.lower()}_", dir=str(runtime_root))).resolve()


def _mark_browser_profile_clean(profile_dir: Path) -> None:
    candidate_files = [
        profile_dir / "Default" / "Preferences",
        profile_dir / "Local State",
    ]
    for candidate in candidate_files:
        if not candidate.exists():
            continue
        try:
            payload = json.loads(candidate.read_text(encoding="utf-8-sig"))
        except (OSError, json.JSONDecodeError):
            continue
        if not isinstance(payload, dict):
            continue

        profile_state = payload.get("profile")
        if not isinstance(profile_state, dict):
            profile_state = {}
            payload["profile"] = profile_state

        changed = False
        if profile_state.get("exit_type") != "Normal":
            profile_state["exit_type"] = "Normal"
            changed = True
        if profile_state.get("exited_cleanly") is not True:
            profile_state["exited_cleanly"] = True
            changed = True

        if not changed:
            continue
        try:
            candidate.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        except OSError:
            continue


def _get_pdf_directory() -> Path:
    pdf_dir = Path(DEFAULT_PDF_DIR)
    pdf_dir.mkdir(parents=True, exist_ok=True)
    return pdf_dir.resolve()


def _get_ati_directory() -> Path:
    ati_dir = Path(DEFAULT_ATI_DIR)
    ati_dir.mkdir(parents=True, exist_ok=True)
    return ati_dir.resolve()


def _get_ati_archive_directory() -> Path:
    return _get_archive_directory(_get_download_directory())


def _get_result_directory() -> Path:
    result_dir = Path(DEFAULT_RESULT_DIR)
    result_dir.mkdir(parents=True, exist_ok=True)
    return result_dir.resolve()


def _get_archive_directory(directory: Path, archive_folder_name: str = "Archives") -> Path:
    archive_dir = directory / archive_folder_name
    archive_dir.mkdir(parents=True, exist_ok=True)
    return archive_dir.resolve()


def _get_unique_archive_path(archive_dir: Path, file_name: str) -> Path:
    candidate = archive_dir / file_name
    if not candidate.exists():
        return candidate

    stem = Path(file_name).stem
    suffix = Path(file_name).suffix
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    counter = 1

    while True:
        candidate = archive_dir / f"{stem}_{timestamp}_{counter}{suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def _archive_directory_files(directory: Path, archive_folder_name: str = "Archives") -> int:
    archive_dir = _get_archive_directory(directory, archive_folder_name)
    moved = 0

    for path in directory.iterdir():
        if not path.is_file():
            continue
        target = _get_unique_archive_path(archive_dir, path.name)
        try:
            shutil.move(str(path), str(target))
            moved += 1
        except PermissionError as exc:
            _append_log(
                f"Skipped archiving locked file in {directory}: {path.name} ({exc.__class__.__name__}: {exc})"
            )
        except OSError as exc:
            _append_log(
                f"Skipped archiving unreadable file in {directory}: {path.name} ({exc.__class__.__name__}: {exc})"
            )

    return moved


def _prepare_result_workbook() -> Path:
    template_path = Path(DEFAULT_WORKBOOK_TEMPLATE).resolve()
    if not template_path.exists():
        raise FileNotFoundError(f"Workbook template not found: {template_path}")

    result_dir = _get_result_directory()
    archived_count = _archive_directory_files(result_dir, "Archive")
    output_name = f"{datetime.now().strftime('%Y%m%d')}_Emission_Summary.xlsx"
    output_path = result_dir / output_name
    shutil.copy2(template_path, output_path)
    _append_log(
        f"Prepared result workbook from template: {template_path.name} -> {output_path} "
        f"(archived {archived_count} existing file(s) in {result_dir / 'Archive'})."
    )
    return output_path


def _normalize_run_mode(run_mode: str | None) -> str:
    normalized = re.sub(r"[^a-z0-9]+", "", str(run_mode or "").lower())
    if normalized in {"append", "add", "existing", "continue"}:
        return "append"
    return "fresh"


def _get_latest_result_workbook() -> Path | None:
    result_dir = _get_result_directory()
    workbooks = [
        path
        for path in result_dir.glob("*.xlsx")
        if path.is_file() and not path.name.startswith("~$") and _is_valid_result_workbook(path)
    ]
    if not workbooks:
        return None
    return max(workbooks, key=lambda path: (path.stat().st_mtime, path.name.lower()))


def _is_valid_result_workbook(workbook_path: Path) -> bool:
    if not re.match(r"^\d{8}_Emission_Summary(?:_retry_\d{6}_\d{2})?\.xlsx$", workbook_path.name, re.IGNORECASE):
        return False

    try:
        workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    except Exception:
        return False

    try:
        return "QueryResults" in workbook.sheetnames
    finally:
        workbook.close()


def _build_dated_result_workbook_name(file_name: str, run_date: datetime | None = None) -> str:
    current_date = (run_date or datetime.now()).strftime("%Y%m%d")
    match = re.match(r"^\d{8}_(.+)$", file_name)
    suffix = match.group(1) if match else file_name
    suffix = suffix.replace("Emission summary", "Emission_Summary")
    suffix = re.sub(
        r"Emission_Summary(?:_retry_\d{6}_\d{2})?\.xlsx$",
        "Emission_Summary.xlsx",
        suffix,
        flags=re.IGNORECASE,
    )
    return f"{current_date}_{suffix}"


def _prepare_append_result_workbook(existing_workbook: Path) -> Path:
    expected_name = _build_dated_result_workbook_name(existing_workbook.name)
    if existing_workbook.name == expected_name:
        _append_log(f"Using existing result workbook for append mode: {existing_workbook}")
        return existing_workbook.resolve()

    result_dir = _get_result_directory()
    archive_dir = _get_archive_directory(result_dir, "Archive")
    archived_existing = _get_unique_archive_path(archive_dir, existing_workbook.name)
    output_path = result_dir / expected_name

    if output_path.exists():
        archived_output = _get_unique_archive_path(archive_dir, output_path.name)
        shutil.move(str(output_path), str(archived_output))
        _append_log(f"Archived existing same-day result workbook before append mode rename: {archived_output}")

    shutil.copy2(existing_workbook, output_path)
    shutil.move(str(existing_workbook), str(archived_existing))
    _append_log(
        f"Prepared append-mode result workbook for current date: {existing_workbook.name} -> {output_path.name} "
        f"(archived previous workbook to {archived_existing})."
    )
    return output_path.resolve()


def _resolve_run_workbook(run_mode: str, workbook_override: str | None = None) -> Path:
    if workbook_override:
        return Path(workbook_override).resolve()
    if run_mode == "append":
        existing_workbook = _get_latest_result_workbook()
        if existing_workbook is None:
            raise FileNotFoundError(
                f"Append mode requires an existing result workbook in {_get_result_directory()}"
            )
        return _prepare_append_result_workbook(existing_workbook)
    return _prepare_result_workbook()


def _archive_run_directories() -> tuple[int, int]:
    pdf_moved = _archive_directory_files(_get_pdf_directory())
    imported_moved = 0
    return pdf_moved, imported_moved


def _get_driver_directory() -> Path:
    driver_dir = Path(DEFAULT_DRIVER_DIR)
    driver_dir.mkdir(parents=True, exist_ok=True)
    return driver_dir.resolve()


def _append_log(message: str) -> None:
    log_path = Path(DEFAULT_LOG_FILE)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with log_path.open("a", encoding="utf-8") as log_file:
        log_file.write(f"[{timestamp}] {message}\n")
    _render_status_html(CURRENT_STATUS_STAGE, CURRENT_STATUS_DETAIL, CURRENT_STATUS_STATE)


def _default_window_bounds(window_kind: str) -> dict[str, int]:
    if window_kind == "status":
        return {
            "left": STATUS_WINDOW_POSITION[0],
            "top": STATUS_WINDOW_POSITION[1],
            "width": STATUS_WINDOW_SIZE[0],
            "height": STATUS_WINDOW_SIZE[1],
        }
    return {
        "left": AUTOMATION_WINDOW_POSITION[0],
        "top": AUTOMATION_WINDOW_POSITION[1],
        "width": AUTOMATION_WINDOW_SIZE[0],
        "height": AUTOMATION_WINDOW_SIZE[1],
    }


def _sanitize_window_bounds(bounds: object, defaults: dict[str, int]) -> dict[str, int]:
    if not isinstance(bounds, dict):
        return defaults.copy()

    sanitized = defaults.copy()
    for key in ("left", "top", "width", "height"):
        value = bounds.get(key)
        if not isinstance(value, (int, float)):
            continue
        sanitized[key] = int(value)

    sanitized["width"] = max(320, sanitized["width"])
    sanitized["height"] = max(240, sanitized["height"])
    if os.name == "nt":
        try:
            class _Point(ctypes.Structure):
                _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]

            user32 = ctypes.WinDLL("user32", use_last_error=True)
            user32.MonitorFromPoint.argtypes = [_Point, wintypes.DWORD]
            user32.MonitorFromPoint.restype = wintypes.HANDLE
            if not user32.MonitorFromPoint(
                _Point(sanitized["left"], sanitized["top"]),
                0,
            ):
                sanitized["left"] = defaults["left"]
                sanitized["top"] = defaults["top"]
        except (AttributeError, OSError):
            pass
    return sanitized


def _sanitize_workbook_window_bounds(bounds: dict[str, int]) -> dict[str, int]:
    sanitized = bounds.copy()
    defaults = _default_window_bounds("workbook")
    if sanitized["left"] < 0:
        sanitized["left"] = defaults["left"]
    if sanitized["top"] < 0:
        sanitized["top"] = defaults["top"]
    return sanitized


def _load_window_bounds_state() -> dict[str, dict[str, int]]:
    state_path = Path(DEFAULT_WINDOW_STATE_FILE)
    if not state_path.exists():
        return {}
    try:
        payload = json.loads(state_path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(payload, dict):
        return {}
    return {
        key: value
        for key, value in payload.items()
        if key in {"status", "automation", "workbook"} and isinstance(value, dict)
    }


def _save_window_bounds_state(state: dict[str, dict[str, int]]) -> None:
    try:
        Path(DEFAULT_WINDOW_STATE_FILE).write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
    except OSError:
        return


def _load_window_bounds_into_memory() -> None:
    global CURRENT_STATUS_WINDOW_BOUNDS, CURRENT_AUTOMATION_WINDOW_BOUNDS, CURRENT_WORKBOOK_WINDOW_BOUNDS
    state = _load_window_bounds_state()
    CURRENT_STATUS_WINDOW_BOUNDS = _sanitize_window_bounds(
        state.get("status"),
        _default_window_bounds("status"),
    )
    if CURRENT_STATUS_WINDOW_BOUNDS["width"] == 320 and CURRENT_STATUS_WINDOW_BOUNDS["height"] == 240:
        CURRENT_STATUS_WINDOW_BOUNDS = _default_window_bounds("status")
    CURRENT_AUTOMATION_WINDOW_BOUNDS = _sanitize_window_bounds(
        state.get("automation"),
        _default_window_bounds("automation"),
    )
    CURRENT_WORKBOOK_WINDOW_BOUNDS = _sanitize_workbook_window_bounds(
        _sanitize_window_bounds(
            state.get("workbook"),
            _default_window_bounds("workbook"),
        )
    )


def _get_window_bounds(window_kind: str) -> dict[str, int]:
    if window_kind == "status":
        return CURRENT_STATUS_WINDOW_BOUNDS.copy()
    if window_kind == "workbook":
        return CURRENT_WORKBOOK_WINDOW_BOUNDS.copy()
    return CURRENT_AUTOMATION_WINDOW_BOUNDS.copy()


def _set_window_bounds(window_kind: str, bounds: dict[str, int]) -> None:
    global CURRENT_STATUS_WINDOW_BOUNDS, CURRENT_AUTOMATION_WINDOW_BOUNDS, CURRENT_WORKBOOK_WINDOW_BOUNDS
    defaults = _default_window_bounds(window_kind)
    sanitized = _sanitize_window_bounds(bounds, defaults)
    if window_kind == "status" and sanitized["width"] == 320 and sanitized["height"] == 240:
        sanitized = defaults.copy()
    if window_kind == "status":
        CURRENT_STATUS_WINDOW_BOUNDS = sanitized
    elif window_kind == "workbook":
        sanitized = _sanitize_workbook_window_bounds(sanitized)
        CURRENT_WORKBOOK_WINDOW_BOUNDS = sanitized
    else:
        CURRENT_AUTOMATION_WINDOW_BOUNDS = sanitized
    state = _load_window_bounds_state()
    state[window_kind] = sanitized
    _save_window_bounds_state(state)


def _capture_window_bounds_for_pid(pid: int) -> dict[str, int] | None:
    if os.name != "nt":
        return None

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    rects: list[dict[str, int]] = []

    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def enum_callback(hwnd, _lparam):
        process_id = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
        if process_id.value != pid or not user32.IsWindowVisible(hwnd):
            return True

        rect = wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return True

        width = int(rect.right - rect.left)
        height = int(rect.bottom - rect.top)
        if width < 100 or height < 100:
            return True

        rects.append(
            {
                "left": int(rect.left),
                "top": int(rect.top),
                "width": width,
                "height": height,
            }
        )
        return True

    user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
    if not rects:
        return None
    return max(rects, key=lambda item: item["width"] * item["height"])


def _remember_window_bounds(window_kind: str, pid: int) -> None:
    bounds = _capture_window_bounds_for_pid(pid)
    if bounds is not None:
        _set_window_bounds(window_kind, bounds)


def _apply_window_bounds_to_pid(window_kind: str, pid: int) -> bool:
    if os.name != "nt":
        return False

    bounds = _get_window_bounds(window_kind)
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    moved = False
    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def enum_callback(hwnd, _lparam):
        nonlocal moved
        process_id = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
        if process_id.value != pid or not user32.IsWindowVisible(hwnd):
            return True
        if user32.MoveWindow(hwnd, bounds["left"], bounds["top"], bounds["width"], bounds["height"], True):
            moved = True
            return False
        return True

    user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
    return moved


def _workbook_locked_message(workbook_path: Path) -> str:
    return (
        f"Workbook is locked and cannot be updated: {workbook_path.resolve()}. "
        "Close the Excel file and run the automation again."
    )


def _create_locked_workbook_retry_copy(workbook_path: Path) -> Path:
    timestamp = datetime.now().strftime("%H%M%S")
    suffix = workbook_path.suffix or ".xlsx"
    stem = workbook_path.stem
    parent = workbook_path.parent

    for attempt in range(1, 100):
        candidate = parent / f"{stem}_retry_{timestamp}_{attempt:02d}{suffix}"
        if candidate.exists():
            continue
        shutil.copy2(workbook_path, candidate)
        _append_log(f"Created writable retry workbook from locked result workbook: {candidate}")
        return candidate.resolve()

    raise FileExistsError(f"Could not create a retry workbook copy for {workbook_path}")


def _assert_workbook_writable(workbook_path: Path) -> None:
    try:
        with workbook_path.open("rb+"):
            return
    except PermissionError:
        _close_workbook_for_update(workbook_path)

    last_error: PermissionError | None = None
    retry_deadline = time.monotonic() + 30.0
    next_close_retry = time.monotonic() + 5.0
    while time.monotonic() < retry_deadline:
        try:
            with workbook_path.open("rb+"):
                return
        except PermissionError as exc:
            last_error = exc
            if time.monotonic() >= next_close_retry:
                _close_workbook_for_update(workbook_path)
                next_close_retry = time.monotonic() + 5.0
            time.sleep(0.5)

    raise PermissionError(_workbook_locked_message(workbook_path)) from last_error


def _close_workbook_for_update(workbook_path: Path) -> None:
    resolved_path = workbook_path.resolve()
    escaped_path = str(resolved_path).replace("'", "''")
    command = (
        "$target = '" + escaped_path + "'; "
        "$result = 'not-open'; "
        "try { "
        "  $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); "
        "  $workbook = @($excel.Workbooks | Where-Object { $_.FullName -eq $target }) | Select-Object -First 1; "
        "  if ($workbook) { "
        "    $excel.DisplayAlerts = $false; "
        "    $workbook.Save(); "
        "    $workbook.Close($true); "
        "    $result = 'closed'; "
        "  } "
        "} catch { $result = 'failed' }; "
        "Write-Output $result"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError as exc:
        _append_log(f"Workbook close attempt failed for {resolved_path}: {exc.__class__.__name__}: {exc}")
        return

    outcome = result.stdout.strip().lower()
    if outcome == "closed":
        _append_log(f"Saved and closed previous result workbook before update: {resolved_path}")
    elif outcome == "failed":
        _append_log(f"Could not close previous result workbook before update: {resolved_path}")


def _normalize_workbook_selection(workbook) -> None:
    worksheets = list(getattr(workbook, "worksheets", []))
    if not worksheets:
        return

    active_sheet = getattr(workbook, "active", None)
    if active_sheet not in worksheets:
        active_sheet = worksheets[0]
        workbook.active = 0

    active_index = worksheets.index(active_sheet)
    for worksheet in worksheets:
        worksheet.sheet_view.tabSelected = False
    active_sheet.sheet_view.tabSelected = True

    for view in getattr(workbook, "views", []):
        try:
            view.activeTab = active_index
        except Exception:
            pass
        try:
            view.firstSheet = min(getattr(view, "firstSheet", 0) or 0, active_index)
        except Exception:
            pass


def _save_workbook_or_raise(workbook, workbook_path: Path) -> None:
    try:
        _add_ati_mf4_links_to_workbook(workbook)
        _normalize_workbook_selection(workbook)
        _normalize_workbook_number_formats(workbook)
        workbook.save(workbook_path)
    except PermissionError as exc:
        raise PermissionError(_workbook_locked_message(workbook_path)) from exc


def _add_ati_mf4_links_to_workbook(workbook) -> int:
    ati_root_dir = _get_ati_directory()
    if not ati_root_dir.exists():
        return 0

    for sheet in workbook.worksheets:
        for cell in getattr(sheet, "_cells", {}).values():
            target = ""
            if cell.hyperlink is not None:
                target = str(cell.hyperlink.target or cell.hyperlink.location or "")
            value_text = str(cell.value or "")
            if ".mf4" in f"{target} {value_text}".lower() and (
                "aptl_ati" in f"{target} {value_text}".lower()
                or "combined" in f"{target} {value_text}".lower()
            ):
                cell.value = None
                cell.hyperlink = None

    mf4_by_test_id: dict[str, list[Path]] = {}
    for mf4_path in ati_root_dir.rglob("*.mf4"):
        match = re.match(r"^(\d{7})", mf4_path.stem)
        if match:
            mf4_by_test_id.setdefault(match.group(1), []).append(mf4_path)

    linked_count = 0
    for sheet in workbook.worksheets:
        test_id_columns: list[int] = []
        header_cells = [cell for cell in getattr(sheet, "_cells", {}).values() if cell.row == 1]
        for cell in sorted(header_cells, key=lambda candidate: candidate.column):
            column_index = cell.column
            header = re.sub(r"[^a-z0-9]", "", str(cell.value or "").lower())
            if header in {"testid", "testnumber", "testno"}:
                test_id_columns.append(column_index)

        for column_index in test_id_columns:
            for row_index in range(2, sheet.max_row + 1):
                test_id = _normalize_test_id(sheet.cell(row_index, column_index).value)
                if not test_id or test_id not in mf4_by_test_id:
                    continue

                candidates = mf4_by_test_id[test_id]
                selected_path = next(
                    (path for path in candidates if "combined" in path.name.lower()),
                    candidates[0],
                )
                cell = sheet.cell(row_index, column_index)
                cell.hyperlink = selected_path.resolve().as_uri()
                cell.style = "Hyperlink"
                linked_count += 1

    if linked_count:
        _append_log(f"Added or refreshed {linked_count} ATI MF4 hyperlink(s) across workbook sheets.")
    return linked_count


def _refresh_ati_mf4_links(workbook_path: Path) -> None:
    if not workbook_path.exists():
        return
    workbook = load_workbook(workbook_path)
    try:
        _save_workbook_or_raise(workbook, workbook_path)
    finally:
        workbook.close()


def _build_run_token() -> str:
    return f"{os.getpid()}_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}"


def _read_run_token_pid(token: str) -> int | None:
    try:
        pid_text = str(token or "").strip().split("_", 1)[0]
        return int(pid_text) if pid_text else None
    except (TypeError, ValueError):
        return None


def _is_emission_summary_process_running(pid: int) -> bool:
    if pid <= 0:
        return False

    workspace_marker = str(Path(__file__).resolve().parent).replace("'", "''")
    command = (
        f"$targetId = {pid}; "
        "$proc = Get-CimInstance Win32_Process -Filter \"ProcessId=$targetId\" -ErrorAction SilentlyContinue; "
        "if (-not $proc) { 'missing'; exit }; "
        "$cmd = [string]$proc.CommandLine; "
        f"if ($cmd -like '*{workspace_marker}*' -and ($cmd -like '*EM_SUM.py*' -or $cmd -like '*Emission_Summary.py*')) {{ 'running' }} else {{ 'missing' }}"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError:
        return False
    return result.returncode == 0 and result.stdout.strip().lower() == "running"


def _set_current_run_token() -> str:
    global CURRENT_RUN_TOKEN
    CURRENT_RUN_TOKEN = _build_run_token()
    Path(DEFAULT_RUN_TOKEN_FILE).write_text(CURRENT_RUN_TOKEN, encoding="utf-8")
    return CURRENT_RUN_TOKEN


def _release_run_mutex() -> None:
    global CURRENT_RUN_TOKEN
    if not CURRENT_RUN_TOKEN:
        return

    token_path = Path(DEFAULT_RUN_TOKEN_FILE)
    try:
        if token_path.exists() and token_path.read_text(encoding="utf-8").strip() == CURRENT_RUN_TOKEN:
            token_path.unlink(missing_ok=True)
    except OSError:
        pass
    CURRENT_RUN_TOKEN = ""


def _acquire_run_mutex() -> bool:
    global CURRENT_RUN_TOKEN
    token_path = Path(DEFAULT_RUN_TOKEN_FILE)

    for _ in range(2):
        token = _build_run_token()
        try:
            with token_path.open("x", encoding="utf-8") as token_file:
                token_file.write(token)
            CURRENT_RUN_TOKEN = token
            atexit.register(_release_run_mutex)
            return True
        except FileExistsError:
            try:
                existing_token = token_path.read_text(encoding="utf-8").strip()
            except OSError:
                existing_token = ""

            existing_pid = _read_run_token_pid(existing_token)
            if existing_pid is not None and _is_emission_summary_process_running(existing_pid):
                return False

            try:
                token_path.unlink(missing_ok=True)
            except OSError:
                return False

    raise OSError("Could not acquire the Emission Summary run lock.")
    atexit.register(_release_run_mutex)
    return True


def _build_workbook_window_sync_powershell(workbook_path: Path, include_token_gate: bool) -> str:
    resolved_path = workbook_path.resolve()
    escaped_path = str(resolved_path).replace("'", "''")
    state_path = str(Path(DEFAULT_WINDOW_STATE_FILE).resolve()).replace("'", "''")
    pid_path = str(Path(DEFAULT_WORKBOOK_OPEN_PID_FILE).resolve()).replace("'", "''")
    token_path = str(Path(DEFAULT_RUN_TOKEN_FILE).resolve()).replace("'", "''")
    expected_token = CURRENT_RUN_TOKEN.replace("'", "''")
    workbook_bounds = _get_window_bounds("workbook")

    token_guard = "$shouldOpen = $true; "
    if include_token_gate:
        token_guard = (
            f"$token = Get-Content -LiteralPath '{token_path}' -ErrorAction SilentlyContinue; "
            f"$shouldOpen = $token -eq '{expected_token}'; "
        )

    state_load = (
        "$state = @{}; "
        "if (Test-Path -LiteralPath $statePath) { "
        "  try { "
        "    $rawState = Get-Content -LiteralPath $statePath -Raw; "
        "    if ($rawState) { "
        "      $parsedState = $rawState | ConvertFrom-Json; "
        "      if ($parsedState) { "
        "        foreach ($prop in $parsedState.PSObject.Properties) { $state[$prop.Name] = $prop.Value } "
        "      } "
        "    } "
        "  } catch { $state = @{} } "
        "} "
        "if (-not ($state -is [hashtable])) { $state = @{} } "
    )

    return (
        "$ErrorActionPreference = 'SilentlyContinue'; "
        "Add-Type @\"\n"
        "using System;\n"
        "using System.Runtime.InteropServices;\n"
        "public static class CopilotWindowRect {\n"
        "    [StructLayout(LayoutKind.Sequential)]\n"
        "    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }\n"
        "    [DllImport(\"user32.dll\")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);\n"
        "    [DllImport(\"user32.dll\")] public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);\n"
        "}\n"
        "\"@; "
        "$excel = (Get-Command excel.exe -ErrorAction SilentlyContinue).Source; "
        "Start-Sleep -Milliseconds 3000; "
        f"$target = '{escaped_path}'; "
        f"$statePath = '{state_path}'; "
        + token_guard
        + "if (-not $shouldOpen) { exit }; "
        "$excelApp = $null; "
        "$workbook = $null; "
        "try { $excelApp = New-Object -ComObject Excel.Application } catch { $excelApp = $null }; "
        "if ($excelApp) { "
        "  try { $workbook = $excelApp.Workbooks.Open($target) } catch { $workbook = $null } "
        "}; "
        "if ($excelApp -and $workbook) { "
        "  try { $excelApp.Visible = $true } catch { } "
        "  try { $excelApp.WindowState = -4143 } catch { } "
        "  try { $excelApp.ActiveWindow.WindowState = -4143 } catch { } "
        f"  try {{ $excelApp.Left = {workbook_bounds['left']} }} catch {{ }} "
        f"  try {{ $excelApp.Top = {workbook_bounds['top']} }} catch {{ }} "
        f"  try {{ $excelApp.Width = {workbook_bounds['width']} }} catch {{ }} "
        f"  try {{ $excelApp.Height = {workbook_bounds['height']} }} catch {{ }} "
        "  Start-Sleep -Milliseconds 1000; "
        f"  try {{ [CopilotWindowRect]::MoveWindow([IntPtr]$excelApp.Hwnd, {workbook_bounds['left']}, {workbook_bounds['top']}, {workbook_bounds['width']}, {workbook_bounds['height']}, $true) | Out-Null }} catch {{ }} "
        "  $watchDeadline = (Get-Date).AddHours(8); "
        "  $excelHandle = [IntPtr]::Zero; "
        "  try { $excelHandle = [IntPtr]$excelApp.Hwnd } catch { $excelHandle = [IntPtr]::Zero } "
        "  while ((Get-Date) -lt $watchDeadline) { "
        "    try { "
        "      if (-not $excelApp) { break } "
        "      if (@($excelApp.Workbooks | Where-Object { $_.FullName -eq $target }).Count -eq 0) { break } "
        "      if ($excelHandle -eq [IntPtr]::Zero) { try { $excelHandle = [IntPtr]$excelApp.Hwnd } catch { $excelHandle = [IntPtr]::Zero } } "
        "      if ($excelHandle -eq [IntPtr]::Zero) { break } "
        "      $rect = New-Object CopilotWindowRect+RECT; "
        "      if ([CopilotWindowRect]::GetWindowRect($excelHandle, [ref]$rect)) { "
        + state_load
        + "        $state['workbook'] = @{ left = $rect.Left; top = $rect.Top; width = ($rect.Right - $rect.Left); height = ($rect.Bottom - $rect.Top) }; "
        "        $state | ConvertTo-Json -Compress | Set-Content -LiteralPath $statePath -Encoding UTF8; "
        "      } "
        "    } catch { } "
        "    Start-Sleep -Seconds 2; "
        "  } "
        "  try { if (@($excelApp.Workbooks).Count -eq 0) { $excelApp.Quit() } } catch { } "
        "} else { "
        "  try { if ($excel) { Start-Process -FilePath $excel -ArgumentList @($target) | Out-Null } else { Start-Process -FilePath $target | Out-Null } } catch { } "
        "}; "
        f"Remove-Item -LiteralPath '{pid_path}' -ErrorAction SilentlyContinue"
    )


def _find_excel_executable() -> Path | None:
    candidates = [
        Path(r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"),
        Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


def _open_workbook(workbook_path: Path) -> None:
    resolved_path = workbook_path.resolve()
    excel_executable = _find_excel_executable() if os.name == "nt" else None
    if excel_executable is not None:
        try:
            subprocess.Popen(
                [str(excel_executable), str(resolved_path)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            _append_log(f"Requested workbook open via Excel executable: {resolved_path}")
            return
        except OSError as exc:
            _append_log(
                f"Workbook open via Excel executable failed for {resolved_path}: {exc.__class__.__name__}: {exc}"
            )

    if os.name == "nt" and hasattr(os, "startfile"):
        try:
            os.startfile(str(resolved_path))  # type: ignore[attr-defined]
            _append_log(f"Requested workbook open via file association: {resolved_path}")
            return
        except OSError as exc:
            _append_log(
                f"Workbook open via file association failed for {resolved_path}: {exc.__class__.__name__}: {exc}"
            )

    try:
        pid_path = Path(DEFAULT_WORKBOOK_OPEN_PID_FILE).resolve()
        process = subprocess.Popen(
            [
                "powershell",
                "-NoProfile",
                "-WindowStyle",
                "Hidden",
                "-Command",
                _build_workbook_window_sync_powershell(resolved_path, include_token_gate=True),
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        pid_path.write_text(str(process.pid), encoding="utf-8")
        _append_log(f"Scheduled workbook open: {resolved_path}")
    except OSError as exc:
        _append_log(f"Workbook open scheduling failed: {exc.__class__.__name__}: {exc}")


def _build_window_bounds_watch_powershell(window_kind: str, pid: int, watch_hours: int = 8) -> str:
    state_path = str(Path(DEFAULT_WINDOW_STATE_FILE).resolve()).replace("'", "''")
    escaped_kind = window_kind.replace("'", "''")
    state_load = (
        "$state = @{}; "
        "if (Test-Path -LiteralPath $statePath) { "
        "  try { "
        "    $rawState = Get-Content -LiteralPath $statePath -Raw; "
        "    if ($rawState) { "
        "      $parsedState = $rawState | ConvertFrom-Json; "
        "      if ($parsedState) { "
        "        foreach ($prop in $parsedState.PSObject.Properties) { $state[$prop.Name] = $prop.Value } "
        "      } "
        "    } "
        "  } catch { $state = @{} } "
        "} "
        "if (-not ($state -is [hashtable])) { $state = @{} } "
    )
    return (
        "$ErrorActionPreference = 'SilentlyContinue'; "
        "Add-Type @\"\n"
        "using System;\n"
        "using System.Runtime.InteropServices;\n"
        "public static class CopilotWindowRect {\n"
        "    [StructLayout(LayoutKind.Sequential)]\n"
        "    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }\n"
        "    [DllImport(\"user32.dll\")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);\n"
        "}\n"
        "\"@; "
        f"$targetPid = {pid}; "
        f"$windowKind = '{escaped_kind}'; "
        f"$statePath = '{state_path}'; "
        f"$watchDeadline = (Get-Date).AddHours({watch_hours}); "
        "while ((Get-Date) -lt $watchDeadline) { "
        "  try { "
        "    $process = Get-Process -Id $targetPid -ErrorAction SilentlyContinue; "
        "    if (-not $process) { break } "
        "    $handle = [IntPtr]$process.MainWindowHandle; "
        "    if ($handle -ne [IntPtr]::Zero) { "
        "      $rect = New-Object CopilotWindowRect+RECT; "
        "      if ([CopilotWindowRect]::GetWindowRect($handle, [ref]$rect)) { "
        + state_load
        + "        $state[$windowKind] = @{ left = $rect.Left; top = $rect.Top; width = ($rect.Right - $rect.Left); height = ($rect.Bottom - $rect.Top) }; "
        "        $state | ConvertTo-Json -Compress | Set-Content -LiteralPath $statePath -Encoding UTF8; "
        "      } "
        "    } "
        "  } catch { } "
        "  Start-Sleep -Seconds 2; "
        "}"
    )


def _start_window_bounds_watcher(window_kind: str, pid: int) -> None:
    try:
        subprocess.Popen(
            [
                "powershell",
                "-NoProfile",
                "-WindowStyle",
                "Hidden",
                "-Command",
                _build_window_bounds_watch_powershell(window_kind, pid),
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except OSError as exc:
        _append_log(
            f"Window bounds watcher failed for {window_kind} PID {pid}: {exc.__class__.__name__}: {exc}"
        )


def _is_workbook_open_in_excel(workbook_path: Path) -> bool:
    resolved_path = workbook_path.resolve()
    escaped_path = str(resolved_path).replace("'", "''")
    command = (
        "$target = '" + escaped_path + "'; "
        "try { "
        "  $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application'); "
        "  $isOpen = @($excel.Workbooks | Where-Object { $_.FullName -eq $target }).Count -gt 0; "
        "  if ($isOpen) { 'open' } else { 'closed' } "
        "} catch { 'closed' }"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError as exc:
        _append_log(f"Workbook open-state check failed for {resolved_path}: {exc.__class__.__name__}: {exc}")
        return False

    return result.stdout.strip().lower() == "open"


def _ensure_workbook_open_in_excel(workbook_path: Path, description: str = "Workbook") -> None:
    resolved_path = workbook_path.resolve()
    if _is_workbook_open_in_excel(resolved_path):
        _append_log(f"{description} already open in Excel: {resolved_path}")
        return

    try:
        subprocess.Popen(
            [
                "powershell",
                "-NoProfile",
                "-WindowStyle",
                "Hidden",
                "-Command",
                _build_workbook_window_sync_powershell(resolved_path, include_token_gate=False),
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        _append_log(f"Opened {description} in Excel: {resolved_path}")
    except OSError as exc:
        _append_log(f"Could not open {description} in Excel: {resolved_path} ({exc.__class__.__name__}: {exc})")


def _read_recent_log_lines(max_lines: int = 12) -> list[str]:
    log_path = Path(DEFAULT_LOG_FILE)
    if not log_path.exists():
        return []
    lines = log_path.read_text(encoding="utf-8", errors="replace").splitlines()
    return lines[-max_lines:]


def _atomic_write_text(path: Path, content: str) -> None:
    temp_path = path.with_suffix(path.suffix + ".tmp")
    temp_path.write_text(content, encoding="utf-8")
    os.replace(temp_path, path)


def _best_effort_write_text(path: Path, content: str, retries: int = 3) -> None:
    last_error: OSError | None = None
    for _ in range(retries):
        try:
            path.write_text(content, encoding="utf-8")
            return
        except OSError as exc:
            last_error = exc
            time.sleep(0.1)
    if last_error is not None:
        raise last_error


def _format_today_for_query(start_date: str) -> str:
    today = datetime.now()
    formats = ["%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%Y/%m/%d", "%d-%b-%Y", "%d/%m/%Y"]
    for candidate in formats:
        try:
            datetime.strptime(start_date.strip(), candidate)
            return today.strftime(candidate)
        except (TypeError, ValueError, AttributeError):
            continue
    return today.strftime("%Y-%m-%d")


def _build_query_prompt_markup() -> str:
    with STATUS_INPUT_LOCK:
        request = dict(STATUS_INPUT_REQUEST) if STATUS_INPUT_REQUEST else None
    if not request:
        return ""

    def esc(value: str) -> str:
        return html.escape(value or "")

    prompt_type = request.get("prompt_type", "query_settings")

    if prompt_type == "update_mode":
        new_test_ids = [item for item in str(request.get("new_test_ids", "")).split(",") if item]
        updated_test_ids = [item for item in str(request.get("updated_test_ids", "")).split(",") if item]
        append_button_title = "Add New Tests Only"
        append_button_detail = "Keep the current summary workbook as-is, append only new TESTIDs, update only the new Monitor cells, and refresh the charts."
        if updated_test_ids:
            append_button_title = "Apply Spreadsheet Updates"
            append_button_detail = "Keep the current summary workbook as-is, append new TESTIDs if any, update changed existing TESTIDs from the downloaded spreadsheet, and refresh affected charts."
        elif not new_test_ids:
            append_button_title = "No New Tests To Add"
            append_button_detail = "There are no new TESTIDs to append. Keep the current summary workbook as-is, skip adding rows, keep the same append behavior, and refresh the charts."
        try:
            new_test_details = json.loads(str(request.get("new_test_details_json", "[]")))
        except json.JSONDecodeError:
            new_test_details = []
        try:
            updated_test_details = json.loads(str(request.get("updated_test_details_json", "[]")))
        except json.JSONDecodeError:
            updated_test_details = []
        preview = ", ".join(new_test_ids[:20]) if new_test_ids else "None"
        if len(new_test_ids) > 20:
            preview += ", ..."
        updated_preview = ", ".join(updated_test_ids[:20]) if updated_test_ids else "None"
        if len(updated_test_ids) > 20:
            updated_preview += ", ..."
        detail_cards: list[str] = []
        for item in new_test_details[:12]:
            test_id = esc(str(item.get("TESTID", "")))
            test_name = esc(str(item.get("TEST_NAME", "") or ""))
            testnet_title = esc(str(item.get("TESTNET_TITLE", "") or ""))
            tr_number = esc(str(item.get("TR_NUMBER", "") or ""))
            start_date = esc(str(item.get("START_DATE", "") or ""))
            end_date = esc(str(item.get("END_DATE", "") or ""))
            detail_cards.append(
                f'<div class="prompt-test-item"><div class="prompt-test-head"><strong>{test_id or "N/A"}</strong></div><div class="prompt-test-line">TEST_NAME: {test_name or "-"}</div><div class="prompt-test-line">TESTNET_TITLE: {testnet_title or "-"}</div><div class="prompt-test-line">TR_NUMBER: {tr_number or "-"} | START_DATE: {start_date or "-"} | END_DATE: {end_date or "-"}</div></div>'
            )
        if len(new_test_details) > 12:
            detail_cards.append(
                f'<div class="prompt-test-more">Showing 12 of {len(new_test_details)} new tests. Full TESTID list: {esc(preview)}</div>'
            )
        for item in updated_test_details[:12]:
            test_id = esc(str(item.get("TESTID", "")))
            test_name = esc(str(item.get("TEST_NAME", "") or ""))
            testnet_title = esc(str(item.get("TESTNET_TITLE", "") or ""))
            tr_number = esc(str(item.get("TR_NUMBER", "") or ""))
            start_date = esc(str(item.get("START_DATE", "") or ""))
            end_date = esc(str(item.get("END_DATE", "") or ""))
            detail_cards.append(
                f'<div class="prompt-test-item"><div class="prompt-test-head"><strong>{test_id or "N/A"}</strong> <span>(updated)</span></div><div class="prompt-test-line">TEST_NAME: {test_name or "-"}</div><div class="prompt-test-line">TESTNET_TITLE: {testnet_title or "-"}</div><div class="prompt-test-line">TR_NUMBER: {tr_number or "-"} | START_DATE: {start_date or "-"} | END_DATE: {end_date or "-"}</div></div>'
            )
        if len(updated_test_details) > 12:
            detail_cards.append(
                f'<div class="prompt-test-more">Showing 12 of {len(updated_test_details)} updated tests. Full TESTID list: {esc(updated_preview)}</div>'
            )
        return f"""
    <div class=\"card prompt-card\">
        <div class=\"prompt-title\">Summary Update Decision</div>
        <div class=\"prompt-detail\">Compared the downloaded spreadsheet against the current final summary workbook. Review new and updated TESTIDs below, then choose whether to rebuild everything from the template or apply only the spreadsheet differences.</div>
        <div class=\"prompt-summary\">Existing workbook: <strong>{esc(request.get('existing_workbook_name', 'None'))}</strong></div>
        <div class=\"prompt-summary\">New TEST count: <strong>{len(new_test_ids)}</strong></div>
        <div class=\"prompt-summary\">New TESTIDs: <span>{esc(preview)}</span></div>
        <div class=\"prompt-summary\">Updated existing TEST count: <strong>{len(updated_test_ids)}</strong></div>
        <div class=\"prompt-summary\">Updated existing TESTIDs: <span>{esc(updated_preview)}</span></div>
        <div class=\"prompt-test-list\">{''.join(detail_cards) if detail_cards else '<div class="prompt-test-more">No new or updated tests were found.</div>'}</div>
        <form class=\"status-form\" id=\"update-mode-form\">
            <input type=\"hidden\" name=\"prompt_type\" value=\"update_mode\">
            <input type="hidden" name="run_mode" value="">
            <div class=\"prompt-choice-list\">
                <button type=\"button\" class=\"prompt-choice prompt-choice-primary\" data-run-mode=\"append\">
                    <span class="prompt-choice-title">{esc(append_button_title)}</span>
                    <span class="prompt-choice-detail">{esc(append_button_detail)}</span>
                </button>
                <button type=\"button\" class=\"prompt-choice prompt-choice-danger\" data-run-mode=\"fresh\">
                    <span class=\"prompt-choice-title\">Update All From Template</span>
                    <span class=\"prompt-choice-detail\">Archive the current summary workbook, rebuild a new one from the template, then import and regenerate everything.</span>
                </button>
            </div>
            <div class=\"prompt-message\"></div>
        </form>
    </div>
"""

    return f"""
    <div class=\"card prompt-card\">
        <div class=\"prompt-title\">Query Settings</div>
        <div class=\"prompt-detail\">Review the saved values or enter new ones, then click Start Automation.</div>
        <div class=\"auto-start-notice\">If you do not click Start Automation, the saved settings will start automatically in <strong class=\"auto-start-countdown\" data-auto-start-deadline=\"{esc(request.get('auto_start_deadline_ms', '0'))}\">30</strong> seconds.</div>
        <form class=\"status-form\" id=\"query-settings-form\">
            <input type=\"hidden\" name=\"prompt_type\" value=\"query_settings\">
            <label class=\"prompt-field\">VIN
                <input type=\"text\" name=\"vehicle_vin\" value=\"{esc(request.get('vehicle_vin', ''))}\" autocomplete=\"off\" required>
            </label>
            <label class=\"prompt-field\">Test Requestor
                <input type=\"text\" name=\"test_requestor\" value=\"{esc(request.get('test_requestor', ''))}\" autocomplete=\"off\" required>
            </label>
            <label class=\"prompt-field\">Start Date
                <input type=\"text\" name=\"start_date\" value=\"{esc(request.get('start_date', ''))}\" autocomplete=\"off\" required>
            </label>
            <label class=\"prompt-field\">End Date
                <input type=\"text\" value=\"{esc(request.get('end_date', ''))}\" readonly>
            </label>
            <label class=\"prompt-field\">Sort By
                <input type=\"text\" name=\"sort_by\" value=\"{esc(request.get('sort_by', ''))}\" autocomplete=\"off\" required>
            </label>
            <label class="prompt-field">ATI MF4 Download
                <select name="download_ati_mf4">
                    <option value="0" {'selected' if request.get('download_ati_mf4', '1') != '1' else ''}>No - skip ATI MF4 downloads</option>
                    <option value="1" {'selected' if request.get('download_ati_mf4', '1') == '1' else ''}>Yes - download missing ATI MF4 files</option>
                </select>
            </label>
            <div class=\"prompt-actions\">
                <button type=\"submit\" class=\"prompt-button\">Start Automation</button>
            </div>
            <div class=\"prompt-message\"></div>
        </form>
    </div>
"""


def _build_status_fragment_markup(stage: str, detail: str, state: str, current_index: int, recent_logs: list[str]) -> str:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    safe_stage = html.escape(stage)
    safe_detail = html.escape(detail)
    safe_state = html.escape(state.upper())
    state_class = html.escape(state.lower())
    log_html = "<br>".join(html.escape(line) for line in recent_logs) if recent_logs else "No log entries yet."
    step_cells: list[str] = []

    last_index = len(STATUS_PROCESS_STEPS) - 1
    for index, (label, _) in enumerate(STATUS_PROCESS_STEPS):
        item_class = "pending"
        if state == "completed" and current_index >= last_index:
            item_class = "completed"
        elif index < current_index:
            item_class = "completed"
        elif index == current_index:
            if state == "failed":
                item_class = "failed"
            elif state == "completed":
                item_class = "completed"
            else:
                item_class = "current"
        step_cells.append(
            f'<td class="step-cell"><div class="step {item_class}"><span class="step-index">{index + 1}</span><span class="step-label">{html.escape(label)}</span></div></td>'
        )

    step_rows: list[str] = []
    for index in range(0, len(step_cells), 2):
        left_cell = step_cells[index]
        right_cell = step_cells[index + 1] if index + 1 < len(step_cells) else '<td class="step-cell"></td>'
        step_rows.append(f"<tr>{left_cell}{right_cell}</tr>")

    prompt_markup = _build_query_prompt_markup()
    return f"""
    <div class=\"card\">
        <div class="state {state_class}">{safe_state}</div>
        <h1>{safe_stage}</h1>
        <div class=\"detail\">{safe_detail}</div>
        {prompt_markup}
        <div class=\"process-title\">Process</div>
        <table class=\"process-table\"><tbody>{''.join(step_rows)}</tbody></table>
        <div class=\"meta\">Updated: {timestamp}</div>
        <div class=\"meta\">Log file: <code>{html.escape(str(Path(DEFAULT_LOG_FILE).resolve()))}</code></div>
        <div class=\"logs\">
            <div class=\"logs-title\">Recent Log</div>
            <div>{log_html}</div>
        </div>
    </div>
"""


def _build_status_shell_html() -> str:
    status_bounds = _get_window_bounds("status")
    return f"""<!DOCTYPE html>
<html lang=\"en\">
<head>
    <meta charset=\"utf-8\">
    <title>Emission Summary Status</title>
    <script>
        let refreshTimer = null;
        let promptEditing = false;

        function scheduleRefresh(delayMs) {{
            if (refreshTimer !== null) {{
                window.clearTimeout(refreshTimer);
            }}
            refreshTimer = window.setTimeout(refreshStatus, delayMs);
        }}

        async function refreshStatus() {{
            if (promptEditing) {{
                scheduleRefresh(500);
                return;
            }}
            try {{
                const response = await fetch('/api/status', {{ cache: 'no-store' }});
                if (!response.ok) return;
                const payload = await response.json();
                document.getElementById('app').innerHTML = payload.fragment_html;
                updateAutoStartCountdown();
                const nextDelay = payload.state === 'running' ? 500 : 1200;
                scheduleRefresh(nextDelay);
            }} catch (error) {{}}
        }}

        function updateAutoStartCountdown() {{
            const countdown = document.querySelector('.auto-start-countdown');
            if (!countdown) return;
            const deadline = Number(countdown.dataset.autoStartDeadline || 0);
            const remainingSeconds = Math.max(0, Math.ceil((deadline - Date.now()) / 1000));
            countdown.textContent = String(remainingSeconds);
        }}

        document.addEventListener('submit', async function (event) {{
            if (!event.target.classList.contains('status-form')) return;
            event.preventDefault();
            promptEditing = false;
            const message = event.target.querySelector('.prompt-message');
            if (message) message.textContent = 'Submitting...';
            const response = await fetch('/submit-status-input', {{
                method: 'POST',
                headers: {{ 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' }},
                body: new URLSearchParams(new FormData(event.target)).toString(),
            }});
            const payload = await response.json();
            if (message) message.textContent = payload.message || '';
            await refreshStatus();
        }});

        document.addEventListener('input', function (event) {{
            if (event.target.closest && event.target.closest('.status-form')) {{
                promptEditing = true;
            }}
        }});

        document.addEventListener('change', function (event) {{
            if (event.target.closest && event.target.closest('.status-form')) {{
                promptEditing = true;
            }}
        }});

        document.addEventListener('focusin', function (event) {{
            if (event.target.closest && event.target.closest('.status-form')) {{
                promptEditing = true;
            }}
        }});

        document.addEventListener('pointerdown', function (event) {{
            if (event.target.closest && event.target.closest('.status-form')) {{
                promptEditing = true;
            }}
        }});

        document.addEventListener('click', function (event) {{
            const button = event.target.closest('button[data-run-mode]');
            if (!button) return;
            const form = button.closest('form.status-form');
            if (!form) return;
            const hiddenRunMode = form.querySelector('input[name="run_mode"]');
            if (!hiddenRunMode) return;
            hiddenRunMode.value = button.getAttribute('data-run-mode') || '';
            form.requestSubmit();
        }});

        window.onload = function () {{
            try {{
                window.resizeTo({status_bounds['width']}, {status_bounds['height']});
                window.moveTo({status_bounds['left']}, {status_bounds['top']});
            }} catch (error) {{}}
            refreshStatus();
        }};

        document.addEventListener('visibilitychange', function () {{
            if (!document.hidden) {{
                scheduleRefresh(50);
            }}
        }});
    </script>
    <style>
        :root {{ color-scheme: light; }}
        body {{ font-family: Segoe UI, Arial, sans-serif; margin: 10px; background: linear-gradient(180deg, #edf2f7 0%, #f8fbfd 100%); color: #1f2933; }}
        .card {{ max-width: 1020px; background: #ffffff; border-radius: 14px; padding: 14px 16px; box-shadow: 0 10px 28px rgba(31, 41, 51, 0.10); }}
        .prompt-card {{ margin-top: 12px; background: #f8fbfd; border: 1px solid #d9e2ec; box-shadow: none; }}
        .prompt-title {{ font-size: 14px; font-weight: 700; color: #243b53; margin-bottom: 6px; }}
        .prompt-detail {{ font-size: 12px; color: #52606d; margin-bottom: 10px; line-height: 1.4; }}
        .auto-start-notice {{ margin-bottom: 10px; padding: 8px 10px; border-left: 3px solid #0b5cab; background: #e8f4ff; color: #243b53; font-size: 12px; line-height: 1.4; }}
        .auto-start-countdown {{ color: #0b5cab; font-size: 15px; }}
        .prompt-summary {{ font-size: 12px; color: #243b53; margin-bottom: 8px; line-height: 1.4; word-break: break-word; }}
        .prompt-test-list {{ display: grid; gap: 8px; margin-bottom: 10px; }}
        .prompt-test-item {{ border: 1px solid #d9e2ec; border-radius: 10px; background: #ffffff; padding: 8px 10px; }}
        .prompt-test-head {{ font-size: 12px; color: #102a43; margin-bottom: 4px; }}
        .prompt-test-line {{ font-size: 11px; color: #486581; line-height: 1.4; word-break: break-word; }}
        .prompt-test-more {{ font-size: 11px; color: #486581; line-height: 1.4; }}
        .prompt-choice-list {{ display: grid; gap: 10px; margin-top: 12px; }}
        .prompt-choice {{ width: 100%; text-align: left; border: 1px solid transparent; border-radius: 12px; cursor: pointer; padding: 12px 14px; background: #ffffff; transition: transform 120ms ease, box-shadow 120ms ease, border-color 120ms ease; }}
        .prompt-choice:hover {{ transform: translateY(-1px); box-shadow: 0 8px 20px rgba(15, 23, 42, 0.08); }}
        .prompt-choice-title {{ display: block; font-size: 15px; font-weight: 800; line-height: 1.25; }}
        .prompt-choice-detail {{ display: block; margin-top: 5px; font-size: 12px; line-height: 1.45; }}
        .prompt-choice-primary {{ border-color: #7cc0ff; background: linear-gradient(180deg, #e8f4ff 0%, #d8ecff 100%); color: #0b5cab; min-height: 88px; }}
        .prompt-choice-danger {{ border-color: #f3b5b5; background: #fff5f5; color: #b42318; min-height: 64px; padding: 10px 12px; }}
        .prompt-field {{ display: block; margin-bottom: 10px; font-size: 12px; font-weight: 600; color: #243b53; }}
        .prompt-summary {{ font-size: 12px; color: #243b53; margin-bottom: 8px; line-height: 1.4; word-break: break-word; }}
        .prompt-test-list {{ display: grid; gap: 8px; margin-bottom: 10px; }}
        .prompt-test-item {{ border: 1px solid #d9e2ec; border-radius: 10px; background: #ffffff; padding: 8px 10px; }}
        .prompt-test-head {{ font-size: 12px; color: #102a43; margin-bottom: 4px; }}
        .prompt-test-line {{ font-size: 11px; color: #486581; line-height: 1.4; word-break: break-word; }}
        .prompt-test-more {{ font-size: 11px; color: #486581; line-height: 1.4; }}
        .prompt-actions-row {{ display: flex; gap: 8px; flex-wrap: wrap; }}
        .prompt-button-secondary {{ background: #486581; }}
        .prompt-field input, .prompt-field select {{ width: 100%; box-sizing: border-box; margin-top: 4px; padding: 8px 10px; border: 1px solid #bcccdc; border-radius: 8px; font-size: 13px; color: #102a43; background: #ffffff; }}
        .prompt-field input[readonly] {{ background: #eef2f6; color: #52606d; }}
        .prompt-actions {{ margin-top: 4px; }}
        .prompt-actions-row {{ display: flex; gap: 8px; flex-wrap: wrap; }}
        .prompt-button {{ border: 0; background: #0b5cab; color: #ffffff; padding: 9px 12px; border-radius: 8px; font-size: 13px; font-weight: 700; cursor: pointer; }}
        .prompt-button-secondary {{ background: #486581; }}
        .prompt-message {{ margin-top: 8px; min-height: 16px; font-size: 12px; color: #0b5cab; }}
        .state {{ display: inline-block; padding: 5px 10px; border-radius: 999px; font-size: 13px; font-weight: 700; background: #d9ecff; color: #0b5cab; }}
        .state.completed {{ background: #e7f8ee; color: #1f7a3d; }}
        .state.failed {{ background: #ffe3e3; color: #b42318; }}
        h1 {{ margin: 10px 0 6px; font-size: 28px; }}
        .detail {{ font-size: 14px; line-height: 1.45; margin-top: 8px; }}
        .meta {{ margin-top: 10px; font-size: 12px; color: #52606d; }}
        .process-title {{ margin-top: 12px; font-size: 12px; font-weight: 700; letter-spacing: 0.03em; text-transform: uppercase; color: #486581; }}
        .process-table {{ width: 100%; margin-top: 8px; border-collapse: separate; border-spacing: 6px 6px; table-layout: fixed; }}
        .step-cell {{ width: 50%; vertical-align: top; }}
        .step {{ min-height: 36px; padding: 6px 9px; border-radius: 8px; background: #eef2f6; color: #52606d; border: 1px solid transparent; overflow: hidden; }}
        .step-index {{ float: left; width: 24px; height: 24px; border-radius: 999px; display: inline-block; text-align: center; line-height: 24px; font-size: 11px; font-weight: 700; background: rgba(31, 41, 51, 0.08); margin-right: 8px; }}
        .step-label {{ display: block; overflow: hidden; font-size: 12px; font-weight: 600; line-height: 1.2; padding-top: 2px; }}
        .step.current {{ background: #d9ecff; color: #0b5cab; border-color: #7cc0ff; }}
        .step.current .step-index {{ background: #0b5cab; color: #ffffff; }}
        .step.completed {{ background: #e7f8ee; color: #1f7a3d; border-color: #9fe0b1; }}
        .step.completed .step-index {{ background: #1f7a3d; color: #ffffff; }}
        .step.failed {{ background: #ffe3e3; color: #b42318; border-color: #ffb3b3; }}
        .step.failed .step-index {{ background: #b42318; color: #ffffff; }}
        .logs {{ margin-top: 12px; padding: 10px 12px; background: #f8fafc; border-radius: 10px; font-size: 11px; line-height: 1.4; color: #243b53; max-height: 132px; overflow: auto; white-space: normal; }}
        .logs-title {{ font-size: 11px; font-weight: 700; letter-spacing: 0.03em; text-transform: uppercase; color: #486581; margin-bottom: 6px; }}
        code {{ font-family: Consolas, monospace; }}
    </style>
</head>
<body>
    <div id=\"app\"></div>
</body>
</html>
"""


def _start_status_server() -> None:
    global STATUS_SERVER, STATUS_SERVER_PORT, STATUS_SERVER_THREAD
    if STATUS_SERVER is not None:
        return

    class StatusRequestHandler(BaseHTTPRequestHandler):
        def do_GET(self) -> None:
            parsed = urlparse(self.path)
            if parsed.path == "/api/status":
                recent_logs = list(reversed(_read_recent_log_lines()))
                current_index = CURRENT_STATUS_PROCESS_INDEX if CURRENT_STATUS_PROCESS_INDEX is not None else 0
                payload = {
                    "state": CURRENT_STATUS_STATE,
                    "fragment_html": _build_status_fragment_markup(
                        CURRENT_STATUS_STAGE,
                        CURRENT_STATUS_DETAIL,
                        CURRENT_STATUS_STATE,
                        current_index,
                        recent_logs,
                    )
                }
                body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return

            if parsed.path == "/" or parsed.path == "":
                body = _build_status_shell_html().encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Cache-Control", "no-store")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)
                return

            self.send_error(404)

        def do_POST(self) -> None:
            global STATUS_INPUT_RESPONSE, STATUS_INPUT_REQUEST, LAST_UPDATE_MODE_SELECTION
            parsed = urlparse(self.path)
            if parsed.path != "/submit-status-input":
                self.send_error(404)
                return

            content_length = int(self.headers.get("Content-Length", "0"))
            body = self.rfile.read(content_length).decode("utf-8", errors="replace")
            values = parse_qs(body, keep_blank_values=True)
            prompt_type = values.get("prompt_type", ["query_settings"])[0].strip() or "query_settings"
            if prompt_type == "update_mode":
                run_mode = _normalize_run_mode(values.get("run_mode", [""])[0].strip())
                if run_mode not in RUN_MODE_OPTIONS:
                    payload = {"ok": False, "message": "Choose how to update the summary workbook."}
                else:
                    LAST_UPDATE_MODE_SELECTION = run_mode
                    _append_log(f"Status prompt selected summary update mode: {run_mode}")
                    with STATUS_INPUT_LOCK:
                        STATUS_INPUT_RESPONSE = {"run_mode": run_mode}
                        STATUS_INPUT_REQUEST = None
                        STATUS_INPUT_EVENT.set()
                    payload = {"ok": True, "message": "Selection accepted. Continuing automation..."}
            else:
                response = {
                    "vehicle_vin": values.get("vehicle_vin", [""])[0].strip(),
                    "test_requestor": values.get("test_requestor", [""])[0].strip(),
                    "start_date": values.get("start_date", [""])[0].strip(),
                    "sort_by": _resolve_sort_by_value(values.get("sort_by", [""])[0].strip()),
                    "download_ati_mf4": "1" if values.get("download_ati_mf4", ["1"])[0].strip() == "1" else "0",
                }
                missing = [label for field_name, label in QUERY_SETTING_FIELDS if not response.get(field_name)]
                if missing:
                    payload = {"ok": False, "message": f"Required: {', '.join(missing)}"}
                else:
                    with STATUS_INPUT_LOCK:
                        STATUS_INPUT_RESPONSE = response
                        STATUS_INPUT_REQUEST = None
                        STATUS_INPUT_EVENT.set()
                    payload = {"ok": True, "message": "Values accepted. Starting automation..."}

            raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Cache-Control", "no-store")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)

        def log_message(self, format: str, *args: object) -> None:
            return

    STATUS_SERVER = ThreadingHTTPServer((DEFAULT_STATUS_HOST, 0), StatusRequestHandler)
    STATUS_SERVER_PORT = int(STATUS_SERVER.server_address[1])
    STATUS_SERVER_THREAD = threading.Thread(target=STATUS_SERVER.serve_forever, name="status-server", daemon=True)
    STATUS_SERVER_THREAD.start()


def _write_status_json(stage: str, detail: str, state: str, current_index: int, recent_logs: list[str]) -> None:
    payload = {
        "stage": stage,
        "detail": detail,
        "state": state,
        "current_index": current_index,
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "log_file": str(Path(DEFAULT_LOG_FILE).resolve()),
        "recent_logs": recent_logs,
        "steps": [label for label, _ in STATUS_PROCESS_STEPS],
    }
    try:
        _best_effort_write_text(Path(DEFAULT_STATUS_JSON_FILE), json.dumps(payload, ensure_ascii=False))
    except OSError as exc:
        _append_log(f"Status JSON update skipped: {exc.__class__.__name__}: {exc}")


def _write_status_shell() -> None:
    status_bounds = _get_window_bounds("status")
    status_fragment_path = json.dumps(str(Path(DEFAULT_STATUS_FRAGMENT_FILE).resolve()))
    shell_content = f"""<html>
<head>
    <title>Emission Summary Status</title>
    <HTA:APPLICATION APPLICATIONNAME=\"EmissionSummaryStatus\" BORDER=\"thin\" CAPTION=\"yes\" SHOWINTASKBAR=\"yes\" SINGLEINSTANCE=\"yes\" SYSMENU=\"yes\" WINDOWSTATE=\"normal\">
    <meta http-equiv=\"x-ua-compatible\" content=\"IE=9\" />
    <script language=\"javascript\">
        var statusPath = {status_fragment_path};
        var lastMarkup = "";

        function readStatus() {{
            var fso = new ActiveXObject("Scripting.FileSystemObject");
            if (!fso.FileExists(statusPath)) return null;
            var file = fso.OpenTextFile(statusPath, 1, false, -1);
            var text = file.ReadAll();
            file.Close();
            if (!text) return null;
            return text;
        }}

        function render() {{
            try {{
                var markup = readStatus();
                if (!markup) return;
                if (markup !== lastMarkup) {{
                    document.getElementById('app').innerHTML = markup;
                    lastMarkup = markup;
                }}
            }} catch (error) {{}}
        }}

        function init() {{
            try {{
                window.resizeTo({status_bounds['width']}, {status_bounds['height']});
                window.moveTo({status_bounds['left']}, {status_bounds['top']});
            }} catch (error) {{}}
            render();
            window.setInterval(render, 1500);
        }}
    </script>
    <style>
        body {{ font-family: Segoe UI, Arial, sans-serif; margin: 10px; background: #f4f7fb; color: #1f2933; }}
        .card {{ background: #ffffff; border: 1px solid #d9e2ec; border-radius: 12px; padding: 14px; }}
        .state {{ display: inline-block; padding: 5px 10px; border-radius: 999px; font-size: 13px; font-weight: 700; background: #d9ecff; color: #0b5cab; }}
        .state.completed {{ background: #e7f8ee; color: #1f7a3d; }}
        .state.failed {{ background: #ffe3e3; color: #b42318; }}
        h1 {{ margin: 10px 0 6px; font-size: 24px; }}
        .detail {{ font-size: 13px; line-height: 1.4; margin-top: 8px; }}
        .meta {{ margin-top: 8px; font-size: 11px; color: #52606d; word-break: break-all; }}
        .process-title {{ margin-top: 12px; font-size: 12px; font-weight: 700; letter-spacing: 0.03em; text-transform: uppercase; color: #486581; }}
        .process-table {{ width: 100%; margin-top: 8px; border-collapse: separate; border-spacing: 6px; table-layout: fixed; }}
        .step-cell {{ width: 50%; vertical-align: top; }}
        .step {{ min-height: 46px; padding: 8px 10px; border-radius: 8px; background: #eef2f6; color: #52606d; border: 1px solid transparent; overflow: hidden; }}
        .step-index {{ float: left; width: 24px; height: 24px; line-height: 24px; border-radius: 999px; text-align: center; font-size: 11px; font-weight: 700; background: rgba(31, 41, 51, 0.08); margin-right: 8px; }}
        .step-label {{ display: block; overflow: hidden; font-size: 12px; font-weight: 600; line-height: 1.25; padding-top: 3px; }}
        .step.current {{ background: #d9ecff; color: #0b5cab; border-color: #7cc0ff; }}
        .step.current .step-index {{ background: #0b5cab; color: #ffffff; }}
        .step.completed {{ background: #e7f8ee; color: #1f7a3d; border-color: #9fe0b1; }}
        .step.completed .step-index {{ background: #1f7a3d; color: #ffffff; }}
        .step.failed {{ background: #ffe3e3; color: #b42318; border-color: #ffb3b3; }}
        .step.failed .step-index {{ background: #b42318; color: #ffffff; }}
        .prompt-field input, .prompt-field select {{ width: 100%; box-sizing: border-box; margin-top: 4px; padding: 8px 10px; border: 1px solid #bcccdc; border-radius: 8px; font-size: 13px; color: #102a43; background: #ffffff; }}
        .prompt-field input[readonly] {{ background: #eef2f6; color: #52606d; }}
        .logs {{ margin-top: 12px; padding: 10px 12px; background: #f8fafc; border-radius: 10px; font-size: 11px; line-height: 1.4; color: #243b53; max-height: 128px; overflow: auto; }}
        .logs-title {{ font-size: 11px; font-weight: 700; letter-spacing: 0.03em; text-transform: uppercase; color: #486581; margin-bottom: 6px; }}
    </style>
</head>
<body onload=\"init()\">
    <div id=\"app\"></div>
</body>
</html>
"""
    _best_effort_write_text(Path(DEFAULT_STATUS_SHELL_FILE), shell_content)


def _render_status_html(stage: str, detail: str, state: str) -> None:
    current_index = CURRENT_STATUS_PROCESS_INDEX if CURRENT_STATUS_PROCESS_INDEX is not None else 0
    recent_logs = list(reversed(_read_recent_log_lines()))
    fragment_markup = _build_status_fragment_markup(stage, detail, state, current_index, recent_logs)
    status_bounds = _get_window_bounds("status")

    html_content = f"""<!DOCTYPE html>
<html lang=\"en\">
<head>
    <meta charset=\"utf-8\">
    <meta http-equiv=\"refresh\" content=\"1\">
    <title>Emission Summary Status</title>
    <script>
        window.onload = function () {{
            try {{
                window.resizeTo({status_bounds['width']}, {status_bounds['height']});
                window.moveTo({status_bounds['left']}, {status_bounds['top']});
            }} catch (error) {{}}
        }};
    </script>
    <style>
        :root {{ color-scheme: light; }}
        body {{ font-family: Segoe UI, Arial, sans-serif; margin: 10px; background: linear-gradient(180deg, #edf2f7 0%, #f8fbfd 100%); color: #1f2933; }}
        .card {{ max-width: 1020px; background: #ffffff; border-radius: 14px; padding: 14px 16px; box-shadow: 0 10px 28px rgba(31, 41, 51, 0.10); }}
        .state {{ display: inline-block; padding: 5px 10px; border-radius: 999px; font-size: 13px; font-weight: 700; background: #d9ecff; color: #0b5cab; }}
        h1 {{ margin: 10px 0 6px; font-size: 28px; }}
        .detail {{ font-size: 14px; line-height: 1.45; margin-top: 8px; }}
        .meta {{ margin-top: 10px; font-size: 12px; color: #52606d; }}
        .process-title {{ margin-top: 12px; font-size: 12px; font-weight: 700; letter-spacing: 0.03em; text-transform: uppercase; color: #486581; }}
        .process-table {{ width: 100%; margin-top: 8px; border-collapse: separate; border-spacing: 6px 6px; table-layout: fixed; }}
        .step-cell {{ width: 50%; vertical-align: top; }}
        .step {{ min-height: 36px; padding: 6px 9px; border-radius: 8px; background: #eef2f6; color: #52606d; border: 1px solid transparent; overflow: hidden; }}
        .step-index {{ float: left; width: 24px; height: 24px; border-radius: 999px; display: inline-block; text-align: center; line-height: 24px; font-size: 11px; font-weight: 700; background: rgba(31, 41, 51, 0.08); margin-right: 8px; }}
        .step-label {{ display: block; overflow: hidden; font-size: 12px; font-weight: 600; line-height: 1.2; padding-top: 2px; }}
        .step.current {{ background: #d9ecff; color: #0b5cab; border-color: #7cc0ff; }}
        .step.current .step-index {{ background: #0b5cab; color: #ffffff; }}
        .step.completed {{ background: #e7f8ee; color: #1f7a3d; border-color: #9fe0b1; }}
        .step.completed .step-index {{ background: #1f7a3d; color: #ffffff; }}
        .step.failed {{ background: #ffe3e3; color: #b42318; border-color: #ffb3b3; }}
        .step.failed .step-index {{ background: #b42318; color: #ffffff; }}
        .logs {{ margin-top: 12px; padding: 10px 12px; background: #f8fafc; border-radius: 10px; font-size: 11px; line-height: 1.4; color: #243b53; max-height: 132px; overflow: auto; white-space: normal; }}
        .logs-title {{ font-size: 11px; font-weight: 700; letter-spacing: 0.03em; text-transform: uppercase; color: #486581; margin-bottom: 6px; }}
        code {{ font-family: Consolas, monospace; }}
    </style>
</head>
<body>
{fragment_markup}
</body>
</html>
"""
    try:
        _best_effort_write_text(Path(DEFAULT_STATUS_HTML_FILE), html_content)
    except OSError as exc:
        _append_log(f"Status HTML update skipped: {exc.__class__.__name__}: {exc}")
    try:
        _best_effort_write_text(Path(DEFAULT_STATUS_FRAGMENT_FILE), fragment_markup)
    except OSError as exc:
        _append_log(f"Status fragment update skipped: {exc.__class__.__name__}: {exc}")
    _write_status_json(stage, detail, state, current_index, recent_logs)


def _update_status(stage: str, detail: str, state: str = "running") -> None:
        global CURRENT_STATUS_STAGE, CURRENT_STATUS_DETAIL, CURRENT_STATUS_STATE, CURRENT_STATUS_PROCESS_INDEX
        CURRENT_STATUS_STAGE = stage
        CURRENT_STATUS_DETAIL = detail
        CURRENT_STATUS_STATE = state
        resolved_index: int | None = None
        current_index = CURRENT_STATUS_PROCESS_INDEX if CURRENT_STATUS_PROCESS_INDEX is not None else 0
        for index, (_, aliases) in enumerate(STATUS_PROCESS_STEPS):
            if stage in aliases:
                resolved_index = index
                break
        if resolved_index is None and state == "failed":
            resolved_index = current_index
        elif resolved_index is None:
            resolved_index = current_index
        elif resolved_index < current_index and state != "failed":
            resolved_index = current_index
        CURRENT_STATUS_PROCESS_INDEX = resolved_index
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status_text = f"[{timestamp}] {state.upper()} | {stage} | {detail}\n"
        try:
            _best_effort_write_text(Path(DEFAULT_STATUS_TEXT_FILE), status_text)
        except OSError as exc:
            _append_log(f"Status text update skipped: {exc.__class__.__name__}: {exc}")
        _render_status_html(stage, detail, state)


def _open_status_window(headless: bool = False) -> None:
    status_path = Path(DEFAULT_STATUS_HTML_FILE).resolve()
    status_profile_dir = Path(DEFAULT_STATUS_PROFILE_DIR).resolve()
    status_profile_dir.mkdir(parents=True, exist_ok=True)
    status_url = f"http://{DEFAULT_STATUS_HOST}:{STATUS_SERVER_PORT}/" if STATUS_SERVER_PORT else status_path.as_uri()

    _close_previous_status_window()
    if headless:
        _append_log(f"Headless mode: status window suppressed; status remains available at {status_url}")
        return

    status_bounds = _get_window_bounds("status")
    width, height = status_bounds["width"], status_bounds["height"]
    left, top = status_bounds["left"], status_bounds["top"]

    launch_commands: list[tuple[str, list[str]]] = []
    edge_binary = _resolve_browser_binary("edge")
    if edge_binary:
        launch_commands.append(
            (
                "Edge",
                [
                    edge_binary,
                    f"--app={status_url}",
                    "--no-first-run",
                    f"--user-data-dir={status_profile_dir}",
                    f"--window-size={width},{height}",
                    f"--window-position={left},{top}",
                ],
            )
        )

    chrome_binary = _resolve_browser_binary("chrome")
    if chrome_binary:
        launch_commands.append(
            (
                "Chrome",
                [
                    chrome_binary,
                    f"--app={status_url}",
                    "--no-first-run",
                    f"--user-data-dir={status_profile_dir}",
                    f"--window-size={width},{height}",
                    f"--window-position={left},{top}",
                ],
            )
        )

    for browser_name, command in launch_commands:
        try:
            process = subprocess.Popen(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            resolved_pid = _resolve_status_window_pid(status_profile_dir) or process.pid
            _apply_window_bounds_to_pid("status", resolved_pid)
            Path(DEFAULT_STATUS_PID_FILE).write_text(str(resolved_pid), encoding="utf-8")
            _start_window_bounds_watcher("status", resolved_pid)
            _append_log(f"Opened status window with {browser_name}: {status_url}")
            return
        except OSError as exc:
            _append_log(f"Status window launch failed for {browser_name}: {exc.__class__.__name__}: {exc}")

    try:
        os.startfile(status_path)  # type: ignore[attr-defined]
        _append_log(f"Opened status page with default application: {status_path}")
    except OSError as exc:
        _append_log(f"Status page launch failed: {exc.__class__.__name__}: {exc}")


def _close_previous_status_window() -> None:
    pid_path = Path(DEFAULT_STATUS_PID_FILE)
    if not pid_path.exists():
        _close_stray_status_window_processes()
        return

    try:
        pid = int(pid_path.read_text(encoding="utf-8").strip())
    except (OSError, ValueError):
        pid_path.unlink(missing_ok=True)
        return

    try:
        _remember_window_bounds("status", pid)
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        _append_log(f"Closed previous status window process: PID {pid}")
    finally:
        pid_path.unlink(missing_ok=True)

    _close_stray_status_window_processes()


def _schedule_status_window_close(delay_milliseconds: int = 4000) -> None:
    pid_path = Path(DEFAULT_STATUS_PID_FILE)
    if not pid_path.exists():
        return

    try:
        pid = int(pid_path.read_text(encoding="utf-8").strip())
    except (OSError, ValueError):
        pid_path.unlink(missing_ok=True)
        return

    try:
        subprocess.Popen(
            [
                "powershell",
                "-NoProfile",
                "-WindowStyle",
                "Hidden",
                "-Command",
                (
                    f"Start-Sleep -Milliseconds {delay_milliseconds}; "
                    f"taskkill /PID {pid} /T /F > $null 2>&1; "
                    f"Remove-Item -LiteralPath '{pid_path.resolve()}' -ErrorAction SilentlyContinue"
                ),
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        _append_log(f"Scheduled status window close: PID {pid}")
    except OSError as exc:
        _append_log(f"Status window close scheduling failed: {exc.__class__.__name__}: {exc}")


def _get_browser_process_ids_by_profile_dir(profile_dir: Path) -> list[int]:
    profile_root = str(profile_dir.resolve()).replace("'", "''")
    command = (
        "$root = '" + profile_root + "'; "
        "$targets = Get-CimInstance Win32_Process | Where-Object { "
        "  ($_.Name -eq 'msedge.exe' -or $_.Name -eq 'chrome.exe') -and "
        "  $_.CommandLine -and $_.CommandLine.Contains($root) "
        "} | Select-Object -ExpandProperty ProcessId; "
        "if ($targets) { $targets | Sort-Object -Unique }"
    )

    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError as exc:
        _append_log(f"Browser process lookup failed for {profile_dir}: {exc.__class__.__name__}: {exc}")
        return []

    return [int(line.strip()) for line in result.stdout.splitlines() if line.strip().isdigit()]


def _get_descendant_browser_process_ids(service_pid: int, process_names: list[str]) -> list[int]:
    quoted_process_names = ", ".join(f"'{name}'" for name in process_names)
    command = (
        "$names = @("
        + quoted_process_names
        + "); "
        "$all = @(Get-CimInstance Win32_Process); "
        "$queue = New-Object System.Collections.Generic.Queue[int]; "
        f"$queue.Enqueue({service_pid}); "
        "$seen = New-Object 'System.Collections.Generic.HashSet[int]'; "
        "$matches = New-Object System.Collections.Generic.List[int]; "
        "while ($queue.Count -gt 0) { "
        "  $current = $queue.Dequeue(); "
        "  if (-not $seen.Add($current)) { continue } "
        "  $children = $all | Where-Object { $_.ParentProcessId -eq $current }; "
        "  foreach ($child in $children) { "
        "    $childPid = [int]$child.ProcessId; "
        "    $queue.Enqueue($childPid); "
        "    if ($names -contains $child.Name) { [void]$matches.Add($childPid) } "
        "  } "
        "}; "
        "if ($matches.Count -gt 0) { $matches | Sort-Object -Unique }"
    )

    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError:
        return []

    return [int(line.strip()) for line in result.stdout.splitlines() if line.strip().isdigit()]


def _remember_window_bounds_for_pids(window_kind: str, pids: list[int]) -> None:
    best_bounds: dict[str, int] | None = None
    best_area = -1
    for pid in pids:
        bounds = _capture_window_bounds_for_pid(pid)
        if bounds is None:
            continue
        area = bounds["width"] * bounds["height"]
        if area > best_area:
            best_bounds = bounds
            best_area = area
    if best_bounds is not None:
        _set_window_bounds(window_kind, best_bounds)


def _resolve_windowed_pid(pids: list[int]) -> int | None:
    best_pid: int | None = None
    best_area = -1
    for pid in pids:
        bounds = _capture_window_bounds_for_pid(pid)
        if bounds is None:
            continue
        area = bounds["width"] * bounds["height"]
        if area > best_area:
            best_pid = pid
            best_area = area
    if best_pid is not None:
        return best_pid
    if pids:
        return pids[0]
    return None


def _close_stray_status_window_processes() -> None:
    status_profile_dir = Path(DEFAULT_STATUS_PROFILE_DIR).resolve()
    status_pids = _get_browser_process_ids_by_profile_dir(status_profile_dir)
    if not status_pids:
        return

    _remember_window_bounds_for_pids("status", status_pids)
    for pid in status_pids:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
    _append_log("Closed stray status window process(es) by status profile path: " + ", ".join(str(pid) for pid in status_pids))


def _resolve_status_window_pid(status_profile_dir: Path, timeout_seconds: float = 5.0) -> int | None:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        status_pids = _get_browser_process_ids_by_profile_dir(status_profile_dir)
        if status_pids:
            return _resolve_windowed_pid(status_pids)
        time.sleep(0.2)
    return None


def _record_automation_pid(pid: int) -> None:
    Path(DEFAULT_AUTOMATION_PID_FILE).write_text(str(pid), encoding="utf-8")


def _clear_automation_pid() -> None:
    Path(DEFAULT_AUTOMATION_PID_FILE).unlink(missing_ok=True)


def _close_stray_automation_browser_processes() -> None:
    automation_root = str(Path(DEFAULT_AUTOMATION_PROFILE_DIR).resolve()).replace("'", "''")
    command = (
        "$root = '" + automation_root + "'; "
        "$targets = Get-CimInstance Win32_Process | Where-Object { "
        "  ($_.Name -eq 'msedge.exe' -or $_.Name -eq 'chrome.exe') -and "
        "  $_.CommandLine -and $_.CommandLine.Contains($root) "
        "} | Select-Object -ExpandProperty ProcessId; "
        "if ($targets) { $targets | Sort-Object -Unique }"
    )

    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            check=False,
        )
    except OSError as exc:
        _append_log(f"Stray automation browser cleanup failed: {exc.__class__.__name__}: {exc}")
        return

    cleaned_pids = [int(line.strip()) for line in result.stdout.splitlines() if line.strip().isdigit()]
    if cleaned_pids:
        _remember_window_bounds_for_pids("automation", cleaned_pids)
        for pid in cleaned_pids:
            subprocess.run(
                ["taskkill", "/PID", str(pid), "/T", "/F"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
            )
        _append_log(
            "Closed stray automation browser process(es) by automation profile path: "
            + ", ".join(str(pid) for pid in cleaned_pids)
        )


def _close_previous_automation_window() -> None:
    pid_path = Path(DEFAULT_AUTOMATION_PID_FILE)
    if pid_path.exists():
        try:
            pid = int(pid_path.read_text(encoding="utf-8").strip())
        except (OSError, ValueError):
            pid_path.unlink(missing_ok=True)
        else:
            try:
                _remember_window_bounds("automation", pid)
                subprocess.run(
                    ["taskkill", "/PID", str(pid), "/T", "/F"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                )
                _append_log(f"Closed previous automation browser process: PID {pid}")
            finally:
                pid_path.unlink(missing_ok=True)

    _close_stray_automation_browser_processes()


def _get_browser_process_id(driver: WebDriver) -> int | None:
    service_process = getattr(getattr(driver, "service", None), "process", None)
    service_pid = getattr(service_process, "pid", None)
    if not service_pid:
        return None

    browser_name = str(driver.capabilities.get("browserName", "")).lower()
    process_names = ["msedge.exe"] if "edge" in browser_name else ["chrome.exe"]

    deadline = time.monotonic() + 5
    while time.monotonic() < deadline:
        candidate_pids = _get_descendant_browser_process_ids(service_pid, process_names)
        resolved_pid = _resolve_windowed_pid(candidate_pids)
        if resolved_pid is not None:
            return resolved_pid
        time.sleep(0.2)

    return None


def _preserve_automation_browser(driver: WebDriver) -> None:
    browser_pid = _get_browser_process_id(driver)
    if browser_pid:
        _record_automation_pid(browser_pid)
        _start_window_bounds_watcher("automation", browser_pid)
        _append_log(f"Tracking automation browser process: PID {browser_pid}")
    else:
        _append_log("Could not resolve automation browser PID; next run may not auto-close the previous browser window.")

    service = getattr(driver, "service", None)
    if service is None:
        return

    try:
        service.stop()
        _append_log("Stopped Selenium driver service and left automation browser open.")
    except Exception as exc:
        _append_log(f"Driver service stop failed: {exc.__class__.__name__}: {exc}")


def _cancel_pending_workbook_open() -> None:
    pid_path = Path(DEFAULT_WORKBOOK_OPEN_PID_FILE)
    if not pid_path.exists():
        return

    try:
        pid = int(pid_path.read_text(encoding="utf-8").strip())
    except (OSError, ValueError):
        pid_path.unlink(missing_ok=True)
        return

    try:
        subprocess.run(
            ["taskkill", "/PID", str(pid), "/T", "/F"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=False,
        )
        _append_log(f"Cancelled pending workbook open process: PID {pid}")
    finally:
        pid_path.unlink(missing_ok=True)


def _resolve_browser_binary(browser: str) -> str | None:
    candidates = {
        "edge": [
            Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
            Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
        ],
        "chrome": [
            Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
            Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
        ],
    }

    for candidate in candidates[browser]:
        if candidate.exists():
            return str(candidate)
    return None


def _resolve_local_driver(browser: str) -> str | None:
    env_var_names = {
        "edge": ["EDGE_DRIVER_PATH", "MSEDGEDRIVER_PATH"],
        "chrome": ["CHROME_DRIVER_PATH", "CHROMEDRIVER_PATH"],
    }
    driver_names = {
        "edge": "msedgedriver.exe",
        "chrome": "chromedriver.exe",
    }

    for env_var_name in env_var_names[browser]:
        env_value = os.environ.get(env_var_name)
        if env_value and Path(env_value).exists():
            _append_log(f"Using {browser} driver from environment variable {env_var_name}: {env_value}")
            return env_value

    driver_path = _get_driver_directory() / browser / driver_names[browser]
    if driver_path.exists():
        return str(driver_path)
    return None


def _get_windows_file_version(executable_path: str) -> str:
    command = [
        "powershell",
        "-NoProfile",
        "-Command",
        f"(Get-Item '{executable_path}').VersionInfo.ProductVersion",
    ]
    result = subprocess.run(command, capture_output=True, text=True, check=True)
    version = result.stdout.strip()
    if not version:
        raise RuntimeError(f"Could not determine browser version for {executable_path}")
    return version


def _get_major_version(version: str) -> str:
    return version.split(".", 1)[0].strip()


def _read_text_from_url(url: str) -> str:
    request = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(request) as response:
        return response.read().decode("utf-8").strip()


def _download_bytes(url: str) -> bytes:
    request = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(request) as response:
        return response.read()


def _ensure_edge_driver() -> str:
    edge_binary = _resolve_browser_binary("edge")
    if not edge_binary:
        raise RuntimeError("Microsoft Edge is not installed.")

    browser_version = _get_windows_file_version(edge_binary)
    major_version = _get_major_version(browser_version)
    _append_log(f"Detected Edge version: {browser_version}")

    existing = _resolve_local_driver("edge")
    stale_existing: str | None = None
    if existing:
        try:
            driver_version = _get_windows_file_version(existing)
        except (OSError, subprocess.SubprocessError) as exc:
            _append_log(
                f"Could not inspect cached EdgeDriver {existing}: {exc.__class__.__name__}: {exc}"
            )
        else:
            if _get_major_version(driver_version) == major_version:
                _append_log(f"Using cached EdgeDriver: {existing} ({driver_version})")
                return existing
            _append_log(
                f"Cached EdgeDriver is incompatible with Edge {browser_version}: "
                f"driver {driver_version}. It will be replaced."
            )
            stale_existing = existing

    release_urls = [
        f"https://msedgedriver.microsoft.com/LATEST_RELEASE_{major_version}_WINDOWS",
        f"https://msedgedriver.azureedge.net/LATEST_RELEASE_{major_version}_WINDOWS",
        f"https://msedgedriver.microsoft.com/LATEST_RELEASE_{major_version}",
        f"https://msedgedriver.azureedge.net/LATEST_RELEASE_{major_version}",
    ]

    driver_version: str | None = None
    for release_url in release_urls:
        try:
            _append_log(f"Checking EdgeDriver release URL: {release_url}")
            driver_version = _read_text_from_url(release_url)
            if driver_version:
                _append_log(f"Resolved EdgeDriver version: {driver_version}")
                break
        except Exception as exc:
            _append_log(
                f"Release lookup failed for {release_url}: {exc.__class__.__name__}: {exc}"
            )
            continue

    if not driver_version:
        driver_version = browser_version
        _append_log(
            "EdgeDriver release endpoint lookup failed; falling back to installed "
            f"Edge version: {driver_version}"
        )

    zip_urls = [
        f"https://msedgedriver.microsoft.com/{driver_version}/edgedriver_win64.zip",
        f"https://msedgedriver.azureedge.net/{driver_version}/edgedriver_win64.zip",
    ]

    zip_bytes: bytes | None = None
    for zip_url in zip_urls:
        try:
            _append_log(f"Downloading EdgeDriver from: {zip_url}")
            zip_bytes = _download_bytes(zip_url)
            if zip_bytes:
                break
        except Exception as exc:
            _append_log(f"Driver download failed for {zip_url}: {exc.__class__.__name__}: {exc}")
            continue

    if not zip_bytes:
        if stale_existing:
            _append_log(
                "Using incompatible cached EdgeDriver because the matching driver "
                "could not be downloaded."
            )
            return stale_existing
        raise RuntimeError(
            f"Unable to download EdgeDriver {driver_version} for Edge {browser_version}."
        )

    target_dir = _get_driver_directory() / "edge"
    target_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as archive:
        for member in archive.namelist():
            if member.lower().endswith("msedgedriver.exe"):
                archive.extract(member, target_dir)
                extracted_path = target_dir / member
                final_path = target_dir / "msedgedriver.exe"
                if extracted_path != final_path:
                    final_path.write_bytes(extracted_path.read_bytes())
                    extracted_path.unlink(missing_ok=True)
                _append_log(f"Extracted EdgeDriver to: {final_path}")
                return str(final_path)

    raise RuntimeError("Downloaded EdgeDriver archive did not contain msedgedriver.exe")


def _build_edge_options(profile_dir: Path, headless: bool = False) -> EdgeOptions:
    automation_bounds = _get_window_bounds("automation")
    options = EdgeOptions()
    options.add_argument("--new-window")
    options.add_argument(f"--window-size={automation_bounds['width']},{automation_bounds['height']}")
    options.add_argument(f"--window-position={automation_bounds['left']},{automation_bounds['top']}")
    options.add_argument(f"--user-data-dir={profile_dir}")
    options.add_argument("--remote-debugging-port=0")
    options.add_argument("--no-first-run")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-component-update")
    options.add_argument("--disable-domain-reliability")
    options.add_argument("--disable-sync")
    options.add_argument("--disable-usb-keyboard-detect")
    options.add_argument("--disable-logging")
    options.add_argument("--log-level=3")
    if headless:
        options.add_argument("--headless=new")
    else:
        options.add_experimental_option("detach", True)
    edge_binary = _resolve_browser_binary("edge")
    if edge_binary:
        options.binary_location = edge_binary
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": str(_get_download_directory()),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        },
    )
    return options


def build_edge_driver(headless: bool = False) -> WebDriver:
    edge_profile_dir = _get_automation_profile_directory("edge")
    _mark_browser_profile_clean(edge_profile_dir)
    primary_options = _build_edge_options(edge_profile_dir, headless)
    edge_driver_path: str | None = None
    try:
        edge_driver_path = _ensure_edge_driver()
    except RuntimeError as exc:
        _append_log(
            "Custom EdgeDriver resolution failed; falling back to Selenium Manager: "
            f"{exc.__class__.__name__}: {exc}"
        )
        edge_service = EdgeService(log_output=os.devnull)
    else:
        edge_service = EdgeService(
            executable_path=edge_driver_path,
            service_args=["--disable-build-check"],
            log_output=os.devnull,
        )
    try:
        return webdriver.Edge(service=edge_service, options=primary_options)
    except Exception as exc:
        message = str(exc)
        if "DevToolsActivePort" not in message and "user data directory is already in use" not in message.lower():
            raise
        fallback_profile_dir = _get_ephemeral_automation_profile_directory("edge")
        _append_log(
            "Retrying Edge startup with an isolated runtime profile after the persistent profile failed: "
            f"{edge_profile_dir} -> {fallback_profile_dir}"
        )
        fallback_options = _build_edge_options(fallback_profile_dir, headless)
        if edge_driver_path is not None:
            fallback_service = EdgeService(
                executable_path=edge_driver_path,
                service_args=["--disable-build-check"],
                log_output=os.devnull,
            )
        else:
            fallback_service = EdgeService(log_output=os.devnull)
        return webdriver.Edge(service=fallback_service, options=fallback_options)


def build_chrome_driver(headless: bool = False) -> WebDriver:
    automation_bounds = _get_window_bounds("automation")
    options = ChromeOptions()
    chrome_profile_dir = _get_automation_profile_directory("chrome")
    _mark_browser_profile_clean(chrome_profile_dir)
    options.add_argument("--new-window")
    options.add_argument(f"--window-size={automation_bounds['width']},{automation_bounds['height']}")
    options.add_argument(f"--window-position={automation_bounds['left']},{automation_bounds['top']}")
    options.add_argument(f"--user-data-dir={chrome_profile_dir}")
    options.add_argument("--no-first-run")
    options.add_argument("--disable-background-networking")
    options.add_argument("--disable-component-update")
    options.add_argument("--disable-domain-reliability")
    options.add_argument("--disable-sync")
    options.add_argument("--disable-usb-keyboard-detect")
    options.add_argument("--disable-logging")
    options.add_argument("--log-level=3")
    if headless:
        options.add_argument("--headless=new")
    else:
        options.add_experimental_option("detach", True)
    options.add_experimental_option("excludeSwitches", ["enable-logging"])
    chrome_binary = _resolve_browser_binary("chrome")
    if chrome_binary:
        options.binary_location = chrome_binary
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": str(_get_download_directory()),
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        },
    )
    chrome_driver_path = _resolve_local_driver("chrome")
    if chrome_driver_path:
        return webdriver.Chrome(
            service=ChromeService(executable_path=chrome_driver_path, log_output=os.devnull),
            options=options,
        )
    return webdriver.Chrome(service=ChromeService(log_output=os.devnull), options=options)


def create_driver(browser: str, headless: bool = False) -> WebDriver:
    factories: list[tuple[str, Callable[[], WebDriver]]]

    if browser == "edge":
        factories = [("Edge", lambda: build_edge_driver(headless))]
    elif browser == "chrome":
        factories = [("Chrome", lambda: build_chrome_driver(headless))]
    else:
        factories = [
            ("Edge", lambda: build_edge_driver(headless)),
            ("Chrome", lambda: build_chrome_driver(headless)),
        ]

    errors: list[str] = []
    for browser_name, factory in factories:
        try:
            _append_log(f"Starting browser factory: {browser_name}")
            driver = factory()
            automation_bounds = _get_window_bounds("automation")
            driver.set_window_rect(
                automation_bounds["left"],
                automation_bounds["top"],
                automation_bounds["width"],
                automation_bounds["height"],
            )
            print(f"Opened external browser with {browser_name}.")
            _append_log(f"Browser started successfully: {browser_name} at {automation_bounds}.")
            return driver
        except Exception as exc:
            _append_log(f"Browser startup failed for {browser_name}: {exc.__class__.__name__}: {exc}")
            errors.append(f"{browser_name}: {exc.__class__.__name__}: {exc}")

    error_text = "\n".join(errors) if errors else "No browser could be started."
    raise RuntimeError(
        "Failed to start a supported browser. Install Microsoft Edge or Google Chrome, "
        "or allow the script to download a matching Edge driver.\n"
        f"Details:\n{error_text}"
    )


def wait_for_access(driver: WebDriver, timeout_seconds: int) -> bool:
    sign_in_reported = False
    proceed_reported = False

    def has_app_frames(current: WebDriver) -> bool:
        frame_names = {
            (frame.get_attribute("name") or "").strip().lower()
            for frame in current.find_elements(By.CSS_SELECTOR, "frame, iframe")
        }
        return {"contents", "main"}.issubset(frame_names)

    def click_proceed_button(current: WebDriver) -> bool:
        proceed_targets = current.find_elements(
            By.XPATH,
            "//input[(translate(@type, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='button'"
            " or translate(@type, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')='submit')"
            " and contains(@value, 'Proceed to APTL Emissions data website')]"
            " | //button[contains(normalize-space(.), 'Proceed to APTL Emissions data website')]"
            " | //a[contains(normalize-space(.), 'Proceed to APTL Emissions data website')]"
            " | //a[contains(@href, 'cgi-bin/entra-query')]",
        )
        if not proceed_targets:
            return False

        target = proceed_targets[0]

        try:
            current.execute_script("arguments[0].scrollIntoView({block: 'center'});", target)
        except WebDriverException:
            pass

        click_attempts = [
            lambda: target.click(),
            lambda: current.execute_script("arguments[0].click();", target),
            lambda: current.execute_script(
                "arguments[0].dispatchEvent(new MouseEvent('click', {bubbles: true, cancelable: true}));",
                target,
            ),
        ]

        for attempt in click_attempts:
            try:
                attempt()
                return True
            except WebDriverException:
                continue

        href = target.get_attribute("href")
        if href:
            current.get(href)
            return True

        try:
            enclosing_form = target.find_element(By.XPATH, "ancestor::form[1]")
        except WebDriverException:
            enclosing_form = None

        if enclosing_form is not None:
            try:
                enclosing_form.submit()
                return True
            except WebDriverException:
                return False

        return False

    def access_ready(current: WebDriver) -> bool:
        nonlocal sign_in_reported, proceed_reported

        if SUCCESS_TITLE in current.title or has_app_frames(current):
            return True

        if (
            not sign_in_reported
            and "login.microsoftonline.com" in current.current_url.lower()
        ):
            print("Microsoft sign-in page detected. Complete sign-in in the browser to continue.")
            _append_log(f"Sign-in page detected: title='{current.title}', url='{current.current_url}'")
            _update_status("Waiting For Microsoft Sign-In", "Approve the login and Authenticator prompt in the browser.")
            sign_in_reported = True

        if not proceed_reported and "welcomeaptldataentra" in current.current_url.lower():
            _append_log(f"Proceed page detected: title='{current.title}', url='{current.current_url}'")
            _update_status("Waiting For Access", "TEST SYSTEM UTILITIES screen detected. Proceeding automatically.")
            proceed_reported = True

        if click_proceed_button(current):
            try:
                print("Clicked: Proceed to APTL Emissions data website")
                _append_log(
                    f"Clicked Proceed button. Current title='{current.title}', url='{current.current_url}'"
                )
                _update_status("Proceed Clicked", "Moved from welcome page to Microsoft sign-in or app page.")
            except WebDriverException:
                return False

        return SUCCESS_TITLE in current.title or has_app_frames(current)

    wait = WebDriverWait(driver, timeout_seconds)
    wait.until(access_ready)
    _append_log(f"Access ready. Final title='{driver.title}', url='{driver.current_url}'")
    _update_status("Access Confirmed", "Reached the Emissions Utilities application page.")
    return True


def switch_to_frame(driver: WebDriver, frame_name: str, timeout_seconds: int) -> None:
    driver.switch_to.default_content()
    WebDriverWait(driver, timeout_seconds).until(
        EC.frame_to_be_available_and_switch_to_it((By.NAME, frame_name))
    )


def _scroll_current_frame_to_bottom(driver: WebDriver) -> None:
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")


def _scroll_element_into_bottom_view(driver: WebDriver, element: object) -> None:
    driver.execute_script("arguments[0].scrollIntoView(false);", element)


def _restore_query_results_bottom_view(
    driver: WebDriver,
    vehicle_vin: str,
    test_requestor: str,
    start_date: str,
    sort_by: str,
    timeout_seconds: int,
) -> None:
    restore_timeout = max(2, min(timeout_seconds, 8))
    try:
        switch_to_frame(driver, "main", restore_timeout)
    except TimeoutException:
        driver.switch_to.default_content()
        _append_log("Skipped browser bottom-view restore because the main frame was not available.")
        return

    try:
        _scroll_current_frame_to_bottom(driver)
        wait = WebDriverWait(driver, restore_timeout)
        bottom_anchor = wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//input[@type='button' and contains(@value, 'Run program')]"
                    " | //input[@type='button' and normalize-space(@value)='Clear']"
                    " | //a[contains(normalize-space(.), 'Import the results') and contains(normalize-space(.), 'spreadsheet')]",
                )
            )
        )
        _scroll_element_into_bottom_view(driver, bottom_anchor)
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        _append_log("Restored browser to the Reprint Report results page at the bottom controls.")
    except TimeoutException:
        _append_log("Skipped browser bottom-view restore because the bottom controls did not appear in time.")
    finally:
        driver.switch_to.default_content()


def _load_query_settings() -> dict[str, str]:
    settings_path = Path(DEFAULT_QUERY_SETTINGS_FILE)
    if not settings_path.exists():
        return {}
    try:
        payload = json.loads(settings_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(payload, dict):
        return {}
    return {str(key): str(value) for key, value in payload.items() if value not in (None, "")}


def _save_query_settings(settings: dict[str, str]) -> None:
    settings_path = Path(DEFAULT_QUERY_SETTINGS_FILE)
    serializable = {key: value for key, value in settings.items() if value}
    _best_effort_write_text(settings_path, json.dumps(serializable, indent=2, ensure_ascii=True) + "\n")


def _prompt_non_empty(prompt_label: str, default_value: str | None = None) -> str:
    while True:
        suffix = f" [{default_value}]" if default_value else ""
        entered = input(f"{prompt_label}{suffix}: ").strip()
        if entered:
            return entered
        if default_value:
            return default_value
        print(f"{prompt_label} is required.")


def _resolve_sort_by_value(sort_by: str) -> str:
    normalized = re.sub(r"[^a-z0-9]+", "", str(sort_by).lower())
    aliases = {
        "vehicleid": "vehicleid",
        "vehicle": "vehicleid",
        "vin": "vehicleid",
        "testrequestor": "requestor",
        "requestor": "requestor",
        "startdate": "startdate",
        "date": "startdate",
        "testid": "testid",
    }
    return aliases.get(normalized, normalized)


def _collect_query_settings(args: argparse.Namespace, headless: bool = False) -> dict[str, str]:
    saved = _load_query_settings()
    has_saved_values = all(saved.get(field_name) for field_name, _ in QUERY_SETTING_FIELDS)
    current = {
        "vehicle_vin": str((getattr(args, "vehicle_vin", "") or "") if getattr(args, "vehicle_vin", "") else saved.get("vehicle_vin") or ""),
        "test_requestor": str(getattr(args, "test_requestor", "") or saved.get("test_requestor") or ""),
        "start_date": str(getattr(args, "start_date", "") or saved.get("start_date") or ""),
        "sort_by": str(getattr(args, "sort_by", "") or saved.get("sort_by") or ""),
        "download_ati_mf4": "1",
    }

    if headless:
        if not has_saved_values:
            raise RuntimeError("Headless mode requires complete query settings in query_settings.json or command-line arguments.")
        _save_query_settings(current)
        _update_status("Run Started", "Using saved query settings in headless mode.")
        return current

    if STATUS_SERVER_PORT:
        current["sort_by"] = _resolve_sort_by_value(current["sort_by"] or "vehicleid")
        with STATUS_INPUT_LOCK:
            global STATUS_INPUT_REQUEST, STATUS_INPUT_RESPONSE
            STATUS_INPUT_REQUEST = {
                "prompt_type": "query_settings",
                "vehicle_vin": current["vehicle_vin"],
                "test_requestor": current["test_requestor"],
                "start_date": current["start_date"],
                "end_date": _format_today_for_query(current["start_date"]),
                "sort_by": current["sort_by"] or "vehicleid",
                "download_ati_mf4": current["download_ati_mf4"],
                "auto_start_deadline_ms": str(int((time.time() + STATUS_INPUT_TIMEOUT_SECONDS) * 1000)),
            }
            STATUS_INPUT_RESPONSE = None
            STATUS_INPUT_EVENT.clear()
        _update_status("Run Started", "Review VIN, Test Requestor, Start Date, End Date, and Sort By in the status window, then click Start Automation.")
        prompt_received = STATUS_INPUT_EVENT.wait(timeout=STATUS_INPUT_TIMEOUT_SECONDS)
        with STATUS_INPUT_LOCK:
            response = dict(STATUS_INPUT_RESPONSE or {})
            STATUS_INPUT_REQUEST = None
            STATUS_INPUT_RESPONSE = None
        if prompt_received and response:
            _save_query_settings({field_name: response[field_name] for field_name, _ in QUERY_SETTING_FIELDS})
            return response
        _append_log("Status query-settings prompt timed out or returned no values.")
        if has_saved_values:
            _append_log("Continuing with saved query settings after status prompt timeout.")
            _update_status(
                "Run Started",
                "Status prompt timed out. Continuing with saved VIN, requestor, start date, and sort settings.",
            )
            _save_query_settings(current)
            return current

    if has_saved_values:
        print("Confirm query settings before starting:")
        for field_name, label in QUERY_SETTING_FIELDS:
            print(f"- {label}: {current[field_name]}")
        answer = input("Are these values correct? [Y/n]: ").strip().lower()
        if answer in ("", "y", "yes"):
            _save_query_settings(current)
            return current

    print("Enter query settings before starting:")
    collected = {
        "vehicle_vin": _prompt_non_empty("VIN", current["vehicle_vin"] or None),
        "test_requestor": _prompt_non_empty("Test Requestor", current["test_requestor"] or None),
        "start_date": _prompt_non_empty("Start Date", current["start_date"] or None),
        "sort_by": _prompt_non_empty("Sort By", current["sort_by"] or None),
        "download_ati_mf4": "1",
    }
    collected["sort_by"] = _resolve_sort_by_value(collected["sort_by"])
    _save_query_settings(collected)
    return collected


def _read_workbook_test_ids(workbook_path: Path) -> tuple[set[str], str | None]:
    workbook = load_workbook(workbook_path, data_only=False, read_only=True)
    try:
        sheet = workbook["QueryResults"]
        return _get_existing_queryresults_test_ids(sheet), _get_last_queryresults_test_id(sheet)
    finally:
        workbook.close()


@dataclass
class ImportedSpreadsheetComparison:
    new_test_details: list[dict[str, str]]
    updated_test_details: list[dict[str, str]]


@dataclass
class ImportedSpreadsheetApplyResult:
    appended_rows: list[int]
    updated_rows: list[int]
    appended_test_ids: list[str]
    updated_test_ids: list[str]

    @property
    def touched_rows(self) -> list[int]:
        return [*self.appended_rows, *self.updated_rows]


@dataclass
class AtiDownloadJob:
    future: Future[list[Path]]
    executor: ThreadPoolExecutor

    def wait(self) -> list[Path]:
        try:
            return self.future.result()
        finally:
            self.executor.shutdown(wait=True)


def _stringify_cell(value: object) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def _normalize_import_compare_value(value: object) -> object:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, datetime):
        return value.replace(microsecond=0)
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return round(value, 12)
    return str(value).strip()


def _normalize_import_compare_row(values: list[object]) -> tuple[object, ...]:
    normalized = _normalize_row_length(values, 51)
    return tuple(_normalize_import_compare_value(value) for value in normalized)


def _import_row_has_changes(imported_row: list[object], existing_row: tuple[object, ...]) -> bool:
    normalized_imported = _normalize_import_compare_row(imported_row)
    for column_index, (imported_value, existing_value) in enumerate(zip(normalized_imported, existing_row), start=1):
        # The source export leaves its final column blank; preserve the existing workbook value.
        if column_index == 51 and imported_value is None:
            continue
        if imported_value != existing_value:
            return True
    return False


def _get_existing_queryresults_row_map(sheet) -> dict[str, int]:
    existing_rows: dict[str, int] = {}
    for row_index in range(2, sheet.max_row + 1):
        test_id = _normalize_test_id(sheet.cell(row_index, 1).value)
        if test_id and test_id not in existing_rows:
            existing_rows[test_id] = row_index
    return existing_rows


def _collect_queryresults_existing_rows(sheet, max_columns: int = 51) -> tuple[dict[str, int], dict[str, tuple[object, ...]]]:
    existing_rows: dict[str, int] = {}
    existing_row_values: dict[str, tuple[object, ...]] = {}
    for row_index, row_values in enumerate(
        sheet.iter_rows(min_row=2, max_col=max_columns, values_only=True),
        start=2,
    ):
        normalized_row = list(row_values)
        test_id = _normalize_test_id(normalized_row[0] if normalized_row else None)
        if not test_id or test_id in existing_rows:
            continue
        existing_rows[test_id] = row_index
        existing_row_values[test_id] = _normalize_import_compare_row(normalized_row)
    return existing_rows, existing_row_values


def _build_imported_test_detail(row_values: list[object], header_map: dict[str, int]) -> dict[str, str] | None:
    test_id = _normalize_test_id(row_values[0] if row_values else None)
    if test_id is None:
        return None
    return {
        "TESTID": test_id,
        "TEST_NAME": _stringify_cell(row_values[header_map["TEST_NAME"]]) if "TEST_NAME" in header_map and len(row_values) > header_map["TEST_NAME"] else "",
        "TESTNET_TITLE": _stringify_cell(row_values[header_map["TESTNET_TITLE"]]) if "TESTNET_TITLE" in header_map and len(row_values) > header_map["TESTNET_TITLE"] else "",
        "TR_NUMBER": _stringify_cell(row_values[header_map["TR_NUMBER"]]) if "TR_NUMBER" in header_map and len(row_values) > header_map["TR_NUMBER"] else "",
        "START_DATE": _stringify_cell(row_values[header_map["START_DATE"]]) if "START_DATE" in header_map and len(row_values) > header_map["START_DATE"] else "",
        "END_DATE": _stringify_cell(row_values[header_map["END_DATE"]]) if "END_DATE" in header_map and len(row_values) > header_map["END_DATE"] else "",
    }


def _compare_imported_spreadsheet(imported_file_path: Path, workbook_path: Path) -> ImportedSpreadsheetComparison:
    rows = read_imported_spreadsheet(imported_file_path)
    header_row_index = _find_header_row(rows)
    headers = [str(value).strip() if value not in (None, "") else "" for value in rows[header_row_index]]
    header_map = {header: index for index, header in enumerate(headers) if header}
    data_rows = rows[header_row_index + 1 :]
    data_rows = [row for row in data_rows if any(value not in (None, "") for value in row[:51])]

    workbook = load_workbook(workbook_path, data_only=False, read_only=True)
    try:
        sheet = workbook["QueryResults"]
        existing_row_map, existing_row_values = _collect_queryresults_existing_rows(sheet)
    finally:
        workbook.close()

    new_test_details: list[dict[str, str]] = []
    updated_test_details: list[dict[str, str]] = []
    seen_new_test_ids: set[str] = set()
    seen_updated_test_ids: set[str] = set()

    for row_values in data_rows:
        detail = _build_imported_test_detail(row_values, header_map)
        if detail is None:
            continue
        test_id = detail["TESTID"]

        if test_id not in existing_row_values:
            if test_id in seen_new_test_ids:
                continue
            new_test_details.append(detail)
            seen_new_test_ids.add(test_id)
            continue

        if _import_row_has_changes(row_values, existing_row_values[test_id]) and test_id not in seen_updated_test_ids:
            updated_test_details.append(detail)
            seen_updated_test_ids.add(test_id)

    return ImportedSpreadsheetComparison(
        new_test_details=new_test_details,
        updated_test_details=updated_test_details,
    )


def _normalize_header_text(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def _normalize_date_key(value: object) -> str | None:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 20000 <= float(value) <= 80000:
            try:
                converted = datetime(1899, 12, 30) + timedelta(days=float(value))
                return converted.date().isoformat()
            except OverflowError:
                return None
    text = str(value).strip()
    if not text:
        return None
    for candidate in (
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m-%d-%Y",
        "%Y/%m/%d",
        "%d-%b-%Y",
        "%d/%m/%Y",
        "%Y%m%d",
    ):
        try:
            return datetime.strptime(text, candidate).date().isoformat()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text.replace("Z", "")).date().isoformat()
    except ValueError:
        return None


def _normalize_req_no(value: object) -> str | None:
    if value in (None, ""):
        return None
    text = str(value).strip()
    if not text:
        return None
    return text.upper()


def _looks_like_html_document(path: Path) -> bool:
    try:
        prefix = path.read_bytes()[:512].lower()
    except OSError:
        return False
    return b"<html" in prefix or b"<!doctype html" in prefix


def _looks_like_html_payload(payload: bytes) -> bool:
    prefix = payload[:1024].lower()
    return b"<html" in prefix or b"<!doctype html" in prefix


def _is_xml_spreadsheet_workbook(path: Path) -> bool:
    try:
        prefix = path.read_bytes()[:4096].lower()
    except OSError:
        return False
    return b"<workbook" in prefix and b"urn:schemas-microsoft-com:office:spreadsheet" in prefix


def _is_valid_excel_workbook(path: Path) -> bool:
    if _is_xml_spreadsheet_workbook(path):
        return True
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
        workbook.close()
        return True
    except Exception:
        return False


def _resolve_monitor_source_download_urls() -> list[str]:
    urls: list[str] = []
    parsed_url = urlparse(DEFAULT_MONITOR_SOURCE_URL)
    source_doc = parse_qs(parsed_url.query).get("sourcedoc", [""])[0].strip()
    source_doc = source_doc.strip("{}")
    if source_doc:
        urls.append(
            f"{parsed_url.scheme}://{parsed_url.netloc}/sites/FoxUpgradeEU7Cal/_layouts/15/download.aspx?UniqueId={source_doc}"
        )

    if "download=1" in DEFAULT_MONITOR_SOURCE_URL:
        urls.append(DEFAULT_MONITOR_SOURCE_URL)
    else:
        separator = "&" if "?" in DEFAULT_MONITOR_SOURCE_URL else "?"
        urls.append(f"{DEFAULT_MONITOR_SOURCE_URL}{separator}download=1")

    deduped_urls: list[str] = []
    seen: set[str] = set()
    for url in urls:
        if not url or url in seen:
            continue
        deduped_urls.append(url)
        seen.add(url)
    return deduped_urls


def _click_if_present(driver: WebDriver, selectors: list[tuple[str, str]], timeout_seconds: int = 5) -> bool:
    end_time = time.monotonic() + timeout_seconds
    while time.monotonic() < end_time:
        for by, selector in selectors:
            try:
                element = WebDriverWait(driver, 1).until(EC.element_to_be_clickable((by, selector)))
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
                element.click()
                return True
            except Exception:
                continue
    return False


def _wait_for_download_if_any(download_dir: Path, before_files: set[str], timeout_seconds: int) -> Path | None:
    try:
        return _wait_for_new_download(download_dir, before_files, timeout_seconds)
    except TimeoutException:
        return None


def _build_cookie_jar_from_driver(driver: WebDriver, request_url: str) -> CookieJar:
    parsed_url = urlparse(request_url)
    request_host = parsed_url.hostname or ""
    cookie_jar = CookieJar()
    for cookie in driver.get_cookies():
        domain = (cookie.get("domain") or request_host or "").strip()
        if not domain:
            continue
        cookie_jar.set_cookie(
            Cookie(
                version=0,
                name=str(cookie["name"]),
                value=str(cookie["value"]),
                port=None,
                port_specified=False,
                domain=domain,
                domain_specified=bool(cookie.get("domain")),
                domain_initial_dot=domain.startswith("."),
                path=str(cookie.get("path") or "/"),
                path_specified=True,
                secure=bool(cookie.get("secure")),
                expires=cookie.get("expiry"),
                discard=False,
                comment=None,
                comment_url=None,
                rest={"HttpOnly": bool(cookie.get("httpOnly"))},
                rfc2109=False,
            )
        )
    return cookie_jar


def _download_monitor_schedule_with_authenticated_session(
    driver: WebDriver,
    target_path: Path,
    timeout_seconds: int = 60,
) -> Path | None:
    browser_download_urls = _resolve_monitor_source_download_urls()
    for download_url in browser_download_urls:
        try:
            opener = build_opener(HTTPCookieProcessor(_build_cookie_jar_from_driver(driver, download_url)))
            request = Request(
                download_url,
                headers={
                    "User-Agent": "Mozilla/5.0",
                    "Accept": "*/*",
                    "Referer": DEFAULT_MONITOR_SOURCE_URL,
                },
            )
            with opener.open(request, timeout=timeout_seconds) as response:
                payload = response.read()
                content_type = (response.headers.get("Content-Type") or "").lower()

            if not payload:
                continue
            if "text/html" in content_type or _looks_like_html_payload(payload):
                continue

            target_path.write_bytes(payload)
            if _is_valid_excel_workbook(target_path):
                _append_log(f"Downloaded monitor schedule workbook through authenticated session: {target_path}")
                return target_path
            target_path.unlink(missing_ok=True)
        except Exception as exc:
            _append_log(f"Authenticated monitor download failed for {download_url}: {exc.__class__.__name__}: {exc}")
    return None


def _open_url_in_new_tab(driver: WebDriver, page_url: str) -> str:
    original_handles = set(driver.window_handles)
    driver.switch_to.new_window("tab")
    new_handles = [handle for handle in driver.window_handles if handle not in original_handles]
    new_handle = driver.current_window_handle
    if new_handles:
        new_handle = new_handles[-1]
        driver.switch_to.window(new_handle)
    driver.get(page_url)
    return new_handle


def _download_monitor_schedule_with_browser(driver: WebDriver, target_path: Path, timeout_seconds: int = 45) -> Path | None:
    download_dir = _get_download_directory()
    original_window = driver.current_window_handle
    original_url = driver.current_url
    browser_download_urls = _resolve_monitor_source_download_urls()
    try:
        for browser_download_url in browser_download_urls:
            attempt_before_files = _snapshot_downloads(download_dir)
            _append_log(f"Attempting browser monitor download URL: {browser_download_url}")
            _open_url_in_new_tab(driver, browser_download_url)

            if "login.microsoftonline.com" in driver.current_url.lower():
                _append_log(
                    f"Monitor schedule requires Microsoft sign-in: title='{driver.title}', url='{driver.current_url}'"
                )
                _update_status(
                    "Monitor Update Started",
                    "Sign in to Microsoft in the automation browser to download the monitor schedule workbook.",
                )
                try:
                    WebDriverWait(driver, 120).until(
                        lambda current_driver: "login.microsoftonline.com" not in current_driver.current_url.lower()
                    )
                    _append_log(
                        f"Monitor schedule sign-in completed: title='{driver.title}', url='{driver.current_url}'"
                    )
                except TimeoutException:
                    raise TimeoutException(
                        "Timed out waiting for Microsoft sign-in before monitor schedule download."
                    )

            current_url = driver.current_url.lower()
            initial_wait_seconds = timeout_seconds if "download.aspx" in current_url else min(timeout_seconds, 12)
            downloaded_file = _wait_for_download_if_any(download_dir, attempt_before_files, initial_wait_seconds)
            if downloaded_file is None and "download.aspx" not in current_url:
                _click_if_present(
                    driver,
                    [
                        (By.XPATH, "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'save as')]"),
                        (By.XPATH, "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download a copy')]"),
                        (By.XPATH, "//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download a copy')]"),
                        (By.XPATH, "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download')]"),
                        (By.XPATH, "//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download')]"),
                        (By.XPATH, "//*[@aria-label and contains(translate(@aria-label, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download')]"),
                        (By.XPATH, "//*[@title and contains(translate(@title, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download')]"),
                    ],
                    timeout_seconds=6,
                )
                downloaded_file = _wait_for_download_if_any(download_dir, attempt_before_files, 8)

            if downloaded_file is None and "download.aspx" not in current_url:
                opened_file_menu = _click_if_present(
                    driver,
                    [
                        (By.XPATH, "//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'save as')]"),
                    ],
                    timeout_seconds=4,
                )
                time.sleep(1)
                if opened_file_menu:
                    time.sleep(1)
                    _click_if_present(
                        driver,
                        [
                            (By.XPATH, "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'save as')]"),
                            (By.XPATH, "//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'save as')]"),
                        ],
                        timeout_seconds=4,
                    )
                    time.sleep(1)
                    _click_if_present(
                        driver,
                        [
                            (By.XPATH, "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download a copy')]"),
                            (By.XPATH, "//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download a copy')]"),
                            (By.XPATH, "//button[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download')]"),
                            (By.XPATH, "//a[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'download')]"),
                        ],
                        timeout_seconds=6,
                    )
                    downloaded_file = _wait_for_download_if_any(download_dir, attempt_before_files, timeout_seconds)

            if downloaded_file is not None:
                break

            try:
                driver.close()
            except Exception:
                pass
            driver.switch_to.window(original_window)

        if downloaded_file is None:
            raise TimeoutException(f"Timed out waiting for a browser download of {target_path.name}")

        if _looks_like_html_document(downloaded_file) or not _is_valid_excel_workbook(downloaded_file):
            _append_log(f"Browser monitor download returned an unusable file: {downloaded_file}")
            downloaded_file.unlink(missing_ok=True)
            return None

        target_path.unlink(missing_ok=True)
        shutil.move(str(downloaded_file), str(target_path))
        _append_log(f"Downloaded monitor schedule workbook through browser: {target_path}")
        return target_path
    except Exception as exc:
        _append_log(f"Browser monitor download failed: {exc.__class__.__name__}: {exc}")
        return None
    finally:
        try:
            if len(driver.window_handles) > 1:
                driver.close()
        except Exception:
            pass
        try:
            driver.switch_to.window(original_window)
        except Exception:
            pass
        try:
            if driver.current_url != original_url:
                driver.get(original_url)
        except Exception:
            pass


def _resolve_monitor_schedule_workbook(driver: WebDriver | None = None) -> Path | None:
    local_path = Path(DEFAULT_MONITOR_SOURCE_FILE).resolve()
    clipboard_path = Path(DEFAULT_MONITOR_SOURCE_CLIPBOARD_FILE).resolve()
    if driver is not None:
        live_source = _capture_live_monitor_schedule_source(driver)
        if live_source is not None and live_source.exists():
            try:
                if live_source == clipboard_path:
                    _load_monitor_schedule_entries_from_tsv(live_source)
                    _append_log(f"Using live SharePoint monitor schedule clipboard capture: {live_source}")
                else:
                    _load_monitor_schedule_entries(live_source)
                    _append_log(f"Using live SharePoint monitor schedule workbook download: {live_source}")
                return live_source
            except Exception as exc:
                _append_log(f"Live SharePoint monitor source validation failed: {exc.__class__.__name__}: {exc}")
                live_source.unlink(missing_ok=True)

    if clipboard_path.exists():
        try:
            _load_monitor_schedule_entries_from_tsv(clipboard_path)
            _append_log(f"Using cached monitor schedule clipboard capture: {clipboard_path}")
            return clipboard_path
        except Exception:
            clipboard_path.unlink(missing_ok=True)

    if local_path.exists() and _is_valid_excel_workbook(local_path):
        _append_log(f"Using fallback local monitor schedule workbook: {local_path}")
        return local_path

    return None


def _collect_links_from_page_in_temp_window(
    driver: WebDriver,
    page_url: str,
    collector: Callable[[WebDriver], dict[str, str]],
    timeout_seconds: int = 10,
) -> dict[str, str]:
    original_window = driver.current_window_handle
    original_handles = set(driver.window_handles)
    new_window: str | None = None

    try:
        driver.switch_to.default_content()
        _open_url_in_new_tab(driver, page_url)
        new_handles = [handle for handle in driver.window_handles if handle not in original_handles]
        if not new_handles:
            return {}
        new_window = new_handles[-1]
        driver.switch_to.window(new_window)
        try:
            WebDriverWait(driver, timeout_seconds).until(
                lambda current_driver: current_driver.execute_script("return document.readyState") == "complete"
            )
        except TimeoutException:
            pass
        time.sleep(0.5)
        return collector(driver)
    finally:
        if new_window is not None:
            try:
                driver.close()
            except Exception:
                pass
        try:
            driver.switch_to.window(original_window)
        except Exception:
            pass


def _visit_frames_for_element(driver: WebDriver, locator: tuple[str, str], max_depth: int = 6) -> bool:
    def search(depth: int) -> bool:
        if driver.find_elements(*locator):
            return True
        if depth >= max_depth:
            return False

        frames = driver.find_elements(By.CSS_SELECTOR, "iframe, frame")
        for index in range(len(frames)):
            try:
                driver.switch_to.frame(frames[index])
            except Exception:
                continue
            if search(depth + 1):
                return True
            driver.switch_to.parent_frame()
        return False

    driver.switch_to.default_content()
    return search(0)


def _switch_to_monitor_schedule_grid_context(driver: WebDriver, timeout_seconds: int = 30) -> None:
    grid_locator = (By.XPATH, "//div[@role='textbox' and contains(@aria-label, 'Actual Schedule')]")
    tab_locator = (By.XPATH, "//a[@role='tab' and normalize-space()='Actual Schedule']")

    end_time = time.time() + timeout_seconds
    while time.time() < end_time:
        if _visit_frames_for_element(driver, tab_locator):
            tabs = driver.find_elements(*tab_locator)
            if tabs:
                try:
                    tabs[0].click()
                except Exception:
                    driver.execute_script("arguments[0].click();", tabs[0])
                time.sleep(0.5)
        if _visit_frames_for_element(driver, grid_locator):
            return
        time.sleep(0.5)

    raise TimeoutException("Could not locate the Actual Schedule grid in the live monitor workbook.")


def _read_system_clipboard_text(retries: int = 10, delay_seconds: float = 0.25) -> str:
    for _ in range(retries):
        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", "Get-Clipboard -Raw"],
                capture_output=True,
                text=True,
                check=False,
            )
        except OSError:
            return ""

        if result.returncode == 0:
            clipboard_text = result.stdout or ""
            if clipboard_text:
                return clipboard_text
        time.sleep(delay_seconds)
    return ""


def _read_browser_clipboard_text(driver: WebDriver) -> str:
    clipboard_result = driver.execute_async_script(
        """
const done = arguments[0];
navigator.clipboard.readText()
  .then(text => done({ok: true, text}))
  .catch(error => done({ok: false, error: String(error)}));
"""
    )
    if not clipboard_result or not clipboard_result.get("ok"):
        return ""
    return str(clipboard_result.get("text") or "")


def _copy_monitor_schedule_tsv_with_browser(driver: WebDriver, target_path: Path) -> Path | None:
    original_window = driver.current_window_handle
    original_url = driver.current_url
    try:
        _open_url_in_new_tab(driver, DEFAULT_MONITOR_SOURCE_URL)

        if "login.microsoftonline.com" in driver.current_url.lower():
            _append_log(
                f"Monitor schedule requires Microsoft sign-in for clipboard fallback: title='{driver.title}', url='{driver.current_url}'"
            )
            _update_status(
                "Monitor Update Started",
                "Sign in to Microsoft in the automation browser to read the monitor schedule sheet.",
            )
            WebDriverWait(driver, 120).until(
                lambda current_driver: "login.microsoftonline.com" not in current_driver.current_url.lower()
            )

        _switch_to_monitor_schedule_grid_context(driver, timeout_seconds=120)
        grid = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//div[@role='textbox' and contains(@aria-label, 'Actual Schedule')]"))
        )
        ActionChains(driver).move_to_element(grid).click(grid).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).perform()
        time.sleep(0.5)
        ActionChains(driver).key_down(Keys.CONTROL).send_keys("c").key_up(Keys.CONTROL).perform()
        time.sleep(1.0)

        clipboard_text = _read_system_clipboard_text()
        if "Test ID" not in clipboard_text or "Monitor" not in clipboard_text:
            clipboard_text = _read_browser_clipboard_text(driver)
        if "Test ID" not in clipboard_text or "Monitor" not in clipboard_text:
            raise ValueError("Clipboard capture did not contain the Actual Schedule monitor grid.")

        target_path.write_text(clipboard_text, encoding="utf-8")
        _append_log(f"Captured monitor schedule TSV through browser clipboard: {target_path}")
        return target_path
    except Exception as exc:
        _append_log(f"Browser monitor clipboard capture failed: {exc.__class__.__name__}: {exc}")
        return None
    finally:
        try:
            if len(driver.window_handles) > 1:
                driver.close()
        except Exception:
            pass
        try:
            driver.switch_to.window(original_window)
        except Exception:
            pass
        try:
            if driver.current_url != original_url:
                driver.get(original_url)
        except Exception:
            pass


def _resolve_schedule_block_date(rows: list[list[object]], header_row_index: int, start_column: int, end_column: int) -> str | None:
    for candidate_row in range(max(0, header_row_index - 2), header_row_index):
        row_values = rows[candidate_row] if candidate_row < len(rows) else []
        for column_index in range(start_column, end_column + 1):
            value = row_values[column_index] if column_index < len(row_values) else None
            date_key = _normalize_date_key(value)
            if date_key:
                return date_key
    return None


def _open_live_share_workbook_tab(driver: WebDriver) -> bool:
    try:
        _open_url_in_new_tab(driver, DEFAULT_MONITOR_SOURCE_URL)
        WebDriverWait(driver, 30).until(lambda current_driver: current_driver.current_url and current_driver.current_url != "about:blank")
        _append_log(f"Opened Live Share workbook tab in automation browser: {driver.current_url}")
        return True
    except Exception as exc:
        _append_log(f"Could not open Live Share workbook tab in automation browser: {exc.__class__.__name__}: {exc}")
        return False


def _find_monitor_schedule_blocks(rows: list[list[object]]) -> list[dict[str, int | str]]:
    test_id_aliases = {"testid", "test_id", "testno", "testnumber"}
    blocks: list[dict[str, int | str]] = []

    for row_index, row_values in enumerate(rows):
        row_headers = {
            column_index: _normalize_header_text(value)
            for column_index, value in enumerate(row_values)
        }
        for column_index in range(max(0, len(row_values) - 3)):
            req_header = row_headers.get(column_index, "")
            test_id_header = row_headers.get(column_index + 1, "")
            cycle_header = row_headers.get(column_index + 2, "")
            monitor_header = row_headers.get(column_index + 3, "")
            if req_header != "reqno":
                continue
            if test_id_header not in test_id_aliases or cycle_header != "cycle" or monitor_header != "monitor":
                continue

            date_key = _resolve_schedule_block_date(rows, row_index, column_index, column_index + 3)
            if not date_key:
                continue

            blocks.append(
                {
                    "header_row": row_index,
                    "start_column": column_index,
                    "end_column": column_index + 3,
                    "req_no_column": column_index,
                    "test_id_column": column_index + 1,
                    "monitor_column": column_index + 3,
                    "date_key": date_key,
                }
            )

    return blocks


def _sheet_to_row_matrix(sheet) -> list[list[object]]:
    return [list(row) for row in sheet.iter_rows(values_only=True)]


def _extract_monitor_text_from_row(row_values: list[object], block: dict[str, int | str], blocks: list[dict[str, int | str]]) -> str:
    monitor_column = int(block["monitor_column"])
    next_block_start = min(
        [int(other_block["start_column"]) for other_block in blocks if int(other_block["start_column"]) > monitor_column],
        default=len(row_values),
    )
    text_parts: list[str] = []
    for column_index in range(monitor_column, min(next_block_start, len(row_values))):
        text = _stringify_cell(row_values[column_index])
        if not text:
            if column_index == monitor_column:
                continue
            break
        text_parts.append(text)
    return " ".join(text_parts).strip()


def _merge_monitor_schedule_entries(*entry_sets: dict[str, str]) -> dict[str, str]:
    merged: dict[str, str] = {}
    for entries in entry_sets:
        for req_no, monitor_text in entries.items():
            if not monitor_text:
                continue
            current = merged.get(req_no, "")
            if not current or len(monitor_text.strip()) >= len(current.strip()):
                merged[req_no] = monitor_text
    return merged


def _load_monitor_schedule_entries_from_rows(
    rows: list[list[object]],
    source_label: str,
    sheet_name: str | None = None,
) -> dict[str, str]:
    blocks = _find_monitor_schedule_blocks(rows)
    if not blocks:
        return {}

    entries: dict[str, str] = {}
    for index, block in enumerate(blocks):
        start_row = int(block["header_row"]) + 1
        end_row = len(rows)
        for later_block in blocks[index + 1 :]:
            if int(later_block["start_column"]) == int(block["start_column"]):
                end_row = int(later_block["header_row"])
                break

        for row_index in range(start_row, end_row):
            row_values = rows[row_index] if row_index < len(rows) else []
            req_value = row_values[int(block["req_no_column"])] if int(block["req_no_column"]) < len(row_values) else None
            req_no = _normalize_req_no(req_value)
            monitor_text = _extract_monitor_text_from_row(row_values, block, blocks)
            if not req_no or not monitor_text:
                continue
            entries[req_no] = monitor_text

    label = f" sheet {sheet_name}" if sheet_name else ""
    _append_log(
        f"Scanned {len(blocks)} monitor block(s) from {source_label}{label}; cumulative matches so far: {len(entries)}."
    )
    return entries


def _load_monitor_schedule_entries_from_tsv(schedule_path: Path) -> dict[str, str]:
    text = schedule_path.read_text(encoding="utf-8")
    rows = list(csv.reader(io.StringIO(text), delimiter="\t"))
    entries = _load_monitor_schedule_entries_from_rows(rows, schedule_path.name)
    if entries:
        _append_log(f"Loaded {len(entries)} monitor schedule row(s) from {schedule_path.name}.")
        return entries
    raise ValueError(f"Could not find Actual Schedule monitor blocks in {schedule_path.name}")


def _load_xml_spreadsheet_row_matrix(schedule_path: Path) -> dict[str, list[list[object]]]:
    namespace = "{urn:schemas-microsoft-com:office:spreadsheet}"
    workbook_root = ET.parse(schedule_path).getroot()
    sheets: dict[str, list[list[object]]] = {}

    for worksheet in workbook_root.findall(f"{namespace}Worksheet"):
        sheet_name = worksheet.attrib.get(f"{namespace}Name") or ""
        table = worksheet.find(f"{namespace}Table")
        if not sheet_name or table is None:
            continue

        rows: list[list[object]] = []
        current_row_index = 1
        for row_element in table.findall(f"{namespace}Row"):
            indexed_row = row_element.attrib.get(f"{namespace}Index")
            if indexed_row:
                target_row_index = int(indexed_row)
                while current_row_index < target_row_index:
                    rows.append([])
                    current_row_index += 1

            row_values: list[object] = []
            current_column_index = 1
            for cell_element in row_element.findall(f"{namespace}Cell"):
                indexed_column = cell_element.attrib.get(f"{namespace}Index")
                if indexed_column:
                    current_column_index = int(indexed_column)

                while len(row_values) < current_column_index - 1:
                    row_values.append(None)

                data_element = cell_element.find(f"{namespace}Data")
                cell_text = "" if data_element is None else "".join(data_element.itertext())
                row_values.append(cell_text)

                merge_across = int(cell_element.attrib.get(f"{namespace}MergeAcross", "0") or "0")
                for _ in range(merge_across):
                    row_values.append(None)

                current_column_index = len(row_values) + 1

            rows.append(row_values)
            current_row_index += 1

        sheets[sheet_name] = rows

    return sheets


def _load_monitor_schedule_entries(schedule_path: Path) -> dict[str, str]:
    if schedule_path.suffix.lower() == ".tsv":
        return _load_monitor_schedule_entries_from_tsv(schedule_path)

    if _is_xml_spreadsheet_workbook(schedule_path):
        sheet_order = sorted(
            _load_xml_spreadsheet_row_matrix(schedule_path).items(),
            key=lambda item: (0 if _normalize_header_text(item[0]) == "actualschedule" else 1, item[0].lower()),
        )
        entries: dict[str, str] = {}
        for sheet_name, rows in sheet_order:
            entries.update(_load_monitor_schedule_entries_from_rows(rows, schedule_path.name, sheet_name))

        if entries:
            _append_log(f"Loaded {len(entries)} monitor schedule row(s) from {schedule_path.name}.")
            return entries
        raise ValueError(f"Could not find Actual Schedule monitor blocks in {schedule_path.name}")

    workbook = load_workbook(schedule_path, read_only=False, data_only=True)
    try:
        sheet_order = sorted(
            workbook.worksheets,
            key=lambda sheet: (0 if _normalize_header_text(sheet.title) == "actualschedule" else 1, sheet.title.lower()),
        )
        entries: dict[str, str] = {}
        for sheet in sheet_order:
            rows = _sheet_to_row_matrix(sheet)
            entries.update(_load_monitor_schedule_entries_from_rows(rows, schedule_path.name, sheet.title))

        if entries:
            _append_log(f"Loaded {len(entries)} monitor schedule row(s) from {schedule_path.name}.")
            return entries
        raise ValueError(f"Could not find Actual Schedule monitor blocks in {schedule_path.name}")
    finally:
        workbook.close()


def _capture_live_monitor_schedule_source(driver: WebDriver) -> Path | None:
    live_workbook_path = Path(DEFAULT_MONITOR_SOURCE_FILE).resolve().with_name(DEFAULT_MONITOR_SOURCE_LIVE_FILE)
    live_workbook = _download_monitor_schedule_with_browser(driver, live_workbook_path)
    if live_workbook is not None and live_workbook.exists():
        return live_workbook

    live_workbook = _download_monitor_schedule_with_authenticated_session(driver, live_workbook_path)
    if live_workbook is not None and live_workbook.exists():
        return live_workbook

    clipboard_path = Path(DEFAULT_MONITOR_SOURCE_CLIPBOARD_FILE).resolve()
    return _copy_monitor_schedule_tsv_with_browser(driver, clipboard_path)


def _load_monitor_schedule_entries_with_live_fallback(
    schedule_path: Path,
    driver: WebDriver | None = None,
) -> dict[str, str]:
    fallback_entries: dict[str, str] = {}
    if schedule_path.exists():
        fallback_entries = _load_monitor_schedule_entries(schedule_path)

    if driver is None or schedule_path.name.lower() == DEFAULT_MONITOR_SOURCE_LIVE_FILE.lower():
        return fallback_entries

    clipboard_path = Path(DEFAULT_MONITOR_SOURCE_CLIPBOARD_FILE).resolve()
    live_entries: dict[str, str] = {}
    try:
        captured_path = _capture_live_monitor_schedule_source(driver)
        if captured_path is not None and captured_path.exists():
            if captured_path == clipboard_path:
                live_entries = _load_monitor_schedule_entries_from_tsv(captured_path)
            else:
                live_entries = _load_monitor_schedule_entries(captured_path)
    except Exception as exc:
        _append_log(f"Live monitor schedule load failed: {exc.__class__.__name__}: {exc}")

    if live_entries:
        supplemented_entries = _merge_monitor_schedule_entries(live_entries, fallback_entries)
        supplemented_count = sum(1 for req_no in fallback_entries if req_no not in live_entries)
        _append_log(
            f"Using live SharePoint monitor schedule entries: live={len(live_entries)}, fallback_only={supplemented_count}, total={len(supplemented_entries)}."
        )
        return supplemented_entries

    if fallback_entries:
        _append_log(
            f"Live SharePoint monitor schedule unavailable; falling back to local source with {len(fallback_entries)} entries."
        )
    return fallback_entries


def _get_monitor_match_candidates(sheet, row: int, queryresults_header_map: dict[str, int]) -> list[str]:
    req_numbers: list[str] = []
    for candidate in (
        sheet.cell(row, queryresults_header_map["TR_NUMBER"]).value if "TR_NUMBER" in queryresults_header_map else None,
        sheet.cell(row, 3).value,
    ):
        normalized = _normalize_req_no(candidate)
        if normalized and normalized not in req_numbers:
            req_numbers.append(normalized)

    return req_numbers

def update_monitor_values(
    workbook_path: Path,
    schedule_path: Path,
    rows: list[int] | None = None,
    driver: WebDriver | None = None,
) -> int:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")
    if not schedule_path.exists():
        raise FileNotFoundError(f"Monitor schedule workbook not found: {schedule_path}")
    _assert_workbook_writable(workbook_path)

    schedule_entries = _load_monitor_schedule_entries_with_live_fallback(schedule_path, driver=driver)
    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook["QueryResults"]
        summary_value_columns = _get_summary_value_columns(sheet)
        queryresults_header_map = _get_queryresults_header_map(sheet)
        monitor_column = summary_value_columns.get("monitor")
        if monitor_column is None:
            _append_log(f"Monitor column not found in workbook {workbook_path.name}; skipped monitor update.")
            return 0

        _update_status("Monitor Update Started", f"Updating monitor values in {workbook_path.name} using {schedule_path.name}.")
        updated_count = 0
        target_rows = rows if rows is not None else list(range(2, sheet.max_row + 1))
        for row_index in target_rows:
            if row_index < 2 or row_index > sheet.max_row:
                continue
            if _normalize_test_id(sheet.cell(row_index, 1).value) is None:
                continue
            matched_monitor: str | None = None
            for candidate_key in _get_monitor_match_candidates(sheet, row_index, queryresults_header_map):
                matched_monitor = schedule_entries.get(candidate_key)
                if matched_monitor is not None:
                    break
            if matched_monitor is None:
                continue
            cell = sheet.cell(row_index, monitor_column)
            if cell.value == matched_monitor:
                continue
            cell.value = matched_monitor
            updated_count += 1

        _save_workbook_or_raise(workbook, workbook_path)
        _append_log(f"Updated monitor values for {updated_count} row(s) using {schedule_path.name}.")
        _update_status("Monitor Update Complete", f"Updated monitor values for {updated_count} row(s) using {schedule_path.name}.")
        return updated_count
    finally:
        workbook.close()


def _collect_new_test_details(imported_file_path: Path, workbook_path: Path) -> list[dict[str, str]]:
    return _compare_imported_spreadsheet(imported_file_path, workbook_path).new_test_details


def _prompt_update_mode(existing_workbook_path: Path, comparison: ImportedSpreadsheetComparison) -> str:
    new_test_details = comparison.new_test_details
    updated_test_details = comparison.updated_test_details
    new_test_ids = [item.get("TESTID", "") for item in new_test_details if item.get("TESTID")]
    updated_test_ids = [item.get("TESTID", "") for item in updated_test_details if item.get("TESTID")]
    preview = ", ".join(new_test_ids[:20]) if new_test_ids else "None"
    if len(new_test_ids) > 20:
        preview += ", ..."
    updated_preview = ", ".join(updated_test_ids[:20]) if updated_test_ids else "None"
    if len(updated_test_ids) > 20:
        updated_preview += ", ..."

    _append_log(
        f"Comparison with existing workbook {existing_workbook_path.name}: {len(new_test_ids)} new TESTID(s) -> {preview}; {len(updated_test_ids)} updated existing TESTID(s) -> {updated_preview}"
    )

    if STATUS_SERVER_PORT:
        with STATUS_INPUT_LOCK:
            global STATUS_INPUT_REQUEST, STATUS_INPUT_RESPONSE
            STATUS_INPUT_REQUEST = {
                "prompt_type": "update_mode",
                "existing_workbook_name": existing_workbook_path.name,
                "new_test_ids": ",".join(new_test_ids),
                "new_test_details_json": json.dumps(new_test_details, ensure_ascii=False),
                "updated_test_ids": ",".join(updated_test_ids),
                "updated_test_details_json": json.dumps(updated_test_details, ensure_ascii=False),
            }
            STATUS_INPUT_RESPONSE = None
            STATUS_INPUT_EVENT.clear()
        _update_status(
            "Spreadsheet Compared",
            f"Compared with {existing_workbook_path.name}: {len(new_test_ids)} new TESTID(s) ({preview}), {len(updated_test_ids)} updated existing TESTID(s) ({updated_preview}). Choose full refresh or append in the status window.",
        )
        STATUS_INPUT_EVENT.wait()
        with STATUS_INPUT_LOCK:
            response = dict(STATUS_INPUT_RESPONSE or {})
            STATUS_INPUT_REQUEST = None
            STATUS_INPUT_RESPONSE = None
        selected_mode = _normalize_run_mode(response.get("run_mode", "fresh"))
        _append_log(f"Resolved summary update mode from status prompt: {selected_mode}")
        return selected_mode

    print(f"Existing workbook: {existing_workbook_path.name}")
    print(f"New TEST count: {len(new_test_ids)}")
    print(f"New TESTIDs: {preview}")
    print(f"Updated existing TEST count: {len(updated_test_ids)}")
    print(f"Updated existing TESTIDs: {updated_preview}")
    for item in new_test_details[:12]:
        print(
            f"- TESTID={item.get('TESTID', '')} | TEST_NAME={item.get('TEST_NAME', '')} | "
            f"TESTNET_TITLE={item.get('TESTNET_TITLE', '')} | TR_NUMBER={item.get('TR_NUMBER', '')} | "
            f"START_DATE={item.get('START_DATE', '')}"
        )
    while True:
        selected = _normalize_run_mode(input("Update all from template or append new tests only? [fresh/append]: ").strip())
        if selected in RUN_MODE_OPTIONS:
            _append_log(f"Resolved summary update mode from console prompt: {selected}")
            return selected
        print("Enter 'fresh' or 'append'.")


def _find_query_field(driver: WebDriver, timeout_seconds: int, candidates: list[str], keywords: list[str]):
    lowered_keywords = [keyword.lower() for keyword in keywords]

    for candidate in candidates:
        try:
            matches = driver.find_elements(By.NAME, candidate)
            if matches:
                return matches[0]
        except Exception:
            pass
        try:
            matches = driver.find_elements(By.ID, candidate)
            if matches:
                return matches[0]
        except Exception:
            pass

    script = """
const keywords = arguments[0];
const elements = Array.from(document.querySelectorAll('input, select, textarea'));
for (const element of elements) {
  const haystack = [element.name, element.id, element.getAttribute('aria-label'), element.placeholder, element.title]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();
  if (keywords.every(keyword => haystack.includes(keyword))) {
    return element;
  }
}
return null;
"""
    element = driver.execute_script(script, lowered_keywords)
    if element is None:
        raise ValueError(f"Could not find query field using candidates={candidates} keywords={keywords}")
    return element


def _set_query_text_field(driver: WebDriver, timeout_seconds: int, value: str, candidates: list[str], keywords: list[str], label: str) -> None:
    element = _find_query_field(driver, timeout_seconds, candidates, keywords)
    element.clear()
    element.send_keys(value)
    print(f"Entered {label}: {value}")


def _select_sort_option(driver: WebDriver, timeout_seconds: int, sort_by: str) -> None:
    normalized = _resolve_sort_by_value(sort_by)
    selector_candidates = [
        f'input[type="radio"][name="aptl_sort"][value="{normalized}"]',
        f'input[type="radio"][value="{normalized}"]',
    ]
    wait = WebDriverWait(driver, timeout_seconds)
    for selector in selector_candidates:
        try:
            option = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
            option.click()
            print(f"Selected Sort by: {normalized}")
            return
        except TimeoutException:
            continue

    xpath = (
        "//label[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "
        f"'{normalized}')]/preceding-sibling::input[@type='radio'][1]"
        " | //label[contains(translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), "
        f"'{normalized}')]/following-sibling::input[@type='radio'][1]"
    )
    try:
        option = driver.find_element(By.XPATH, xpath)
        driver.execute_script("arguments[0].click();", option)
        print(f"Selected Sort by: {normalized}")
        return
    except Exception as exc:
        raise ValueError(f"Could not find sort option '{sort_by}'.") from exc


def click_target_path(driver: WebDriver, target: str, timeout_seconds: int) -> None:
    if target != DEFAULT_CLICK_TARGET:
        raise ValueError(f"Unsupported click target: {target}")

    switch_to_frame(driver, "contents", timeout_seconds)
    menu_radio = WebDriverWait(driver, timeout_seconds).until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="radio"][name="OPERATION"][onclick*="aptl.html?1"]'))
    )
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", menu_radio)
    menu_radio.click()
    print("Clicked: 1.1.1 Reprint Report")
    _update_status("Menu Selected", "Clicked 1.1.1 Reprint Report.")
    driver.switch_to.default_content()


def populate_reprint_report_query(
    driver: WebDriver,
    vehicle_vin: str,
    test_requestor: str,
    start_date: str,
    sort_by: str,
    timeout_seconds: int,
) -> None:
    switch_to_frame(driver, "main", timeout_seconds)
    wait = WebDriverWait(driver, timeout_seconds)
    end_date = _format_today_for_query(start_date)

    wait.until(
        EC.text_to_be_present_in_element((By.TAG_NAME, "body"), "ELMS Emissions Test Report Query")
    )

    vehicle_vin_input = wait.until(
        EC.presence_of_element_located((By.NAME, "aptl_vehicleid"))
    )
    vehicle_vin_input.clear()
    vehicle_vin_input.send_keys(vehicle_vin)
    print(f"Entered Vehicle VIN: {vehicle_vin}")
    _update_status("VIN Entered", f"Entered Vehicle VIN {vehicle_vin}.")

    _set_query_text_field(
        driver,
        timeout_seconds,
        test_requestor,
        ["aptl_requestor", "requestor", "test_requestor", "aptl_testrequestor"],
        ["requestor"],
        "Test Requestor",
    )
    _update_status("Test Requestor Entered", f"Entered Test Requestor {test_requestor}.")

    _set_query_text_field(
        driver,
        timeout_seconds,
        start_date,
        ["aptl_startdate", "startdate", "start_date", "aptl_begindate"],
        ["start", "date"],
        "Start Date",
    )
    _update_status("Start Date Entered", f"Entered Start Date {start_date}.")

    _set_query_text_field(
        driver,
        timeout_seconds,
        end_date,
        ["aptl_enddate", "enddate", "end_date", "aptl_todate"],
        ["end", "date"],
        "End Date",
    )
    _update_status("End Date Entered", f"Entered End Date {end_date}.")

    _select_sort_option(driver, timeout_seconds, sort_by)
    _update_status("Sort Selected", f"Selected Sort by {sort_by}.")

    submit_button = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, 'input[type="submit"][value*="Submit Query"]'))
    )
    submit_button.click()
    print("Clicked: Submit Query")
    _update_status("Query Submitted", "Submitted the Reprint Report query.")
    driver.switch_to.default_content()


def _is_local_webdriver_read_timeout(exc: BaseException) -> bool:
    if isinstance(exc, ReadTimeoutError):
        return True
    message = str(exc)
    return "HTTPConnectionPool(host='localhost'" in message and "Read timed out" in message


def _is_retryable_navigation_error(exc: BaseException) -> bool:
    if _is_local_webdriver_read_timeout(exc):
        return True
    if not isinstance(exc, WebDriverException):
        return False

    message = str(exc)
    retryable_markers = (
        "net::ERR_CONNECTION_TIMED_OUT",
        "net::ERR_TIMED_OUT",
        "net::ERR_NAME_NOT_RESOLVED",
        "net::ERR_INTERNET_DISCONNECTED",
        "net::ERR_NETWORK_CHANGED",
        "net::ERR_CONNECTION_RESET",
    )
    return any(marker in message for marker in retryable_markers)


class AptlNavigationError(RuntimeError):
    pass


def _open_target_url(driver: WebDriver, url: str, timeout_seconds: int) -> None:
    for attempt in range(2):
        try:
            driver.get(url)
            return
        except Exception as exc:
            is_retryable = _is_retryable_navigation_error(exc)
            if attempt == 1:
                if is_retryable:
                    raise AptlNavigationError(
                        "Could not reach the APTL Emissions server after retrying. "
                        "The company server or network connection appears unavailable right now. "
                        "Please try again later."
                    ) from exc
                raise
            if not is_retryable:
                raise

            _append_log(
                "Initial APTL navigation failed with a retryable browser/network error; resetting the tab and retrying once. "
                f"{exc.__class__.__name__}: {exc}"
            )
            _update_status(
                "Navigation Retry",
                "The first APTL page load timed out. Resetting the browser tab and retrying once.",
            )
            try:
                driver.get("about:blank")
            except Exception:
                pass
            time.sleep(min(2.0, max(0.5, timeout_seconds / 60)))


def _prepare_and_submit_reprint_report_query(
    driver: WebDriver,
    vehicle_vin: str,
    test_requestor: str,
    start_date: str,
    sort_by: str,
    timeout_seconds: int,
) -> None:
    for attempt in range(2):
        try:
            click_target_path(driver, DEFAULT_CLICK_TARGET, timeout_seconds)
            print("Target click sequence completed.")

            populate_reprint_report_query(
                driver,
                vehicle_vin,
                test_requestor,
                start_date,
                sort_by,
                timeout_seconds,
            )
            print("Query form submitted.")
            return
        except Exception as exc:
            if attempt == 1 or not _is_local_webdriver_read_timeout(exc):
                raise

            _append_log(
                "Query setup hit a local WebDriver read timeout; reloading the APTL page and retrying once. "
                f"{exc.__class__.__name__}: {exc}"
            )
            _update_status(
                "Query Setup Retry",
                "The browser stopped responding during query setup. Reloading the APTL page and retrying once.",
            )
            try:
                driver.switch_to.default_content()
            except Exception:
                pass
            _open_target_url(driver, TARGET_URL, timeout_seconds)
            wait_for_access(driver, timeout_seconds)
            _update_status("Navigation Ready", f"Access reconfirmed after retry. Continuing to {DEFAULT_CLICK_TARGET}.")


def _wait_for_new_download(download_dir: Path, before_files: set[str], timeout_seconds: int) -> Path:
    deadline = time.monotonic() + timeout_seconds
    candidate: Path | None = None
    last_size = -1
    stable_checks = 0

    while time.monotonic() < deadline:
        files = [
            path for path in download_dir.iterdir()
            if path.is_file() and path.name not in before_files and not path.name.endswith((".crdownload", ".part", ".tmp"))
        ]

        if files:
            candidate = max(files, key=lambda path: path.stat().st_mtime)
            current_size = candidate.stat().st_size
            if current_size > 0 and current_size == last_size:
                stable_checks += 1
            else:
                stable_checks = 0
                last_size = current_size

            if stable_checks >= 2:
                return candidate

        time.sleep(1)

    raise TimeoutException(f"Timed out waiting for a downloaded spreadsheet in {download_dir}")


def download_imported_spreadsheet(
    driver: WebDriver,
    timeout_seconds: int,
    headless: bool = False,
) -> Path:
    download_dir = _get_download_directory()
    removed_count = _clear_previous_imported_spreadsheets(download_dir)
    before_files = _snapshot_downloads(download_dir)

    switch_to_frame(driver, "main", timeout_seconds)
    wait = WebDriverWait(driver, timeout_seconds)

    import_link = wait.until(
        EC.element_to_be_clickable(
            (
                By.XPATH,
                "//a[contains(normalize-space(.), 'Import the results') and contains(normalize-space(.), 'spreadsheet')]"
                " | //input[contains(translate(@value, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'import the results')"
                " and contains(translate(@value, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'spreadsheet')]"
                " | //button[contains(normalize-space(.), 'Import the results') and contains(normalize-space(.), 'spreadsheet')]",
            )
        )
    )

    if headless:
        import_href = (import_link.get_attribute("href") or "").strip()
        if import_href and not import_href.lower().startswith("javascript:"):
            direct_path = download_dir / "ho3.xls"
            try:
                _download_file_with_session(driver, urljoin(driver.current_url, import_href), direct_path)
                read_imported_spreadsheet(direct_path)
                print(f"Downloaded spreadsheet through authenticated session: {direct_path}")
                _update_status("Spreadsheet Downloaded", f"Downloaded spreadsheet {direct_path.name}.")
                return direct_path
            except Exception as exc:
                direct_path.unlink(missing_ok=True)
                _append_log(
                    f"Direct headless spreadsheet download failed; falling back to browser click: {exc.__class__.__name__}: {exc}"
                )

    _scroll_current_frame_to_bottom(driver)
    _scroll_element_into_bottom_view(driver, import_link)
    import_link.click()
    _scroll_current_frame_to_bottom(driver)
    print(f"Clicked import link. Download directory: {download_dir}")
    _update_status(
        "Import Requested",
        f"Requested spreadsheet download into {download_dir} after deleting {removed_count} previous spreadsheet file(s).",
    )
    driver.switch_to.default_content()
    downloaded_file = _wait_for_new_download(download_dir, before_files, DEFAULT_DOWNLOAD_WAIT_SECONDS)
    print(f"Downloaded spreadsheet: {downloaded_file}")
    _update_status("Spreadsheet Downloaded", f"Downloaded spreadsheet {downloaded_file.name}.")
    return downloaded_file


def _read_xlsx_rows(file_path: Path) -> list[list[object]]:
    workbook = load_workbook(file_path, data_only=False)
    try:
        sheet = workbook.active
        return [list(row) for row in sheet.iter_rows(values_only=True)]
    finally:
        workbook.close()


def _read_csv_rows(file_path: Path) -> list[list[object]]:
    with file_path.open("r", encoding="utf-8-sig", newline="") as handle:
        return [row for row in csv.reader(handle)]


def _read_xls_rows(file_path: Path) -> list[list[object]]:
    import xlrd

    workbook = xlrd.open_workbook(file_path)
    sheet = workbook.sheet_by_index(0)
    rows: list[list[object]] = []
    for row_index in range(sheet.nrows):
        rows.append(sheet.row_values(row_index))
    return rows


def _read_xlsx_hyperlinks(file_path: Path) -> dict[tuple[int, int], str]:
    workbook = load_workbook(file_path, data_only=False)
    try:
        sheet = workbook.active
        hyperlinks: dict[tuple[int, int], str] = {}

        for row_index, row in enumerate(sheet.iter_rows(), start=0):
            for column_index in range(LINKED_DOCUMENT_START_COLUMN, LINKED_DOCUMENT_END_COLUMN + 1):
                cell = row[column_index - 1]
                if cell.hyperlink is None:
                    continue
                target = cell.hyperlink.target or cell.hyperlink.location
                if target:
                    hyperlinks[(row_index, column_index - 1)] = target

        return hyperlinks
    finally:
        workbook.close()


def _read_xls_hyperlinks(file_path: Path) -> dict[tuple[int, int], str]:
    import xlrd

    workbook = xlrd.open_workbook(file_path, formatting_info=True)
    sheet = workbook.sheet_by_index(0)
    hyperlink_map = getattr(sheet, "hyperlink_map", None) or {}
    hyperlinks: dict[tuple[int, int], str] = {}

    for key, link in hyperlink_map.items():
        row_index, column_index = key
        if not (LINKED_DOCUMENT_START_COLUMN - 1 <= column_index <= LINKED_DOCUMENT_END_COLUMN - 1):
            continue
        target = getattr(link, "url_or_path", None)
        if target:
            hyperlinks[(row_index, column_index)] = target

    return hyperlinks


def _read_excel_xml_rows(file_path: Path) -> list[list[object]]:
    namespaces = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
    root = ET.parse(file_path).getroot()
    worksheet = root.find("ss:Worksheet", namespaces)
    if worksheet is None:
        raise ValueError(f"No worksheet found in XML spreadsheet: {file_path}")

    table = worksheet.find("ss:Table", namespaces)
    if table is None:
        raise ValueError(f"No table found in XML spreadsheet: {file_path}")

    rows: list[list[object]] = []
    for row_element in table.findall("ss:Row", namespaces):
        row_values: list[object] = []
        for cell in row_element.findall("ss:Cell", namespaces):
            index_attr = cell.get(f"{{{namespaces['ss']}}}Index")
            if index_attr:
                target_index = int(index_attr) - 1
                while len(row_values) < target_index:
                    row_values.append(None)

            data = cell.find("ss:Data", namespaces)
            if data is None:
                row_values.append(None)
                continue

            cell_type = data.get(f"{{{namespaces['ss']}}}Type", "String")
            text = data.text or ""
            if cell_type == "Number":
                if "." in text:
                    row_values.append(float(text))
                else:
                    row_values.append(int(text))
            else:
                row_values.append(text)

        rows.append(row_values)

    return rows


def _read_excel_xml_hyperlinks(file_path: Path) -> dict[tuple[int, int], str]:
    namespaces = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
    root = ET.parse(file_path).getroot()
    worksheet = root.find("ss:Worksheet", namespaces)
    if worksheet is None:
        raise ValueError(f"No worksheet found in XML spreadsheet: {file_path}")

    table = worksheet.find("ss:Table", namespaces)
    if table is None:
        raise ValueError(f"No table found in XML spreadsheet: {file_path}")

    hyperlinks: dict[tuple[int, int], str] = {}
    for row_index, row_element in enumerate(table.findall("ss:Row", namespaces)):
        column_index = 0
        for cell in row_element.findall("ss:Cell", namespaces):
            index_attr = cell.get(f"{{{namespaces['ss']}}}Index")
            if index_attr:
                column_index = int(index_attr) - 1

            href = cell.get(f"{{{namespaces['ss']}}}HRef")
            if href and LINKED_DOCUMENT_START_COLUMN - 1 <= column_index <= LINKED_DOCUMENT_END_COLUMN - 1:
                hyperlinks[(row_index, column_index)] = href

            column_index += 1

    return hyperlinks


def _looks_like_xml_spreadsheet(file_path: Path) -> bool:
    with file_path.open("rb") as handle:
        prefix = handle.read(256).lstrip()
    return prefix.startswith(b"<?xml") or prefix.startswith(b"<Workbook")


def read_imported_spreadsheet(file_path: Path) -> list[list[object]]:
    suffix = file_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return _read_xlsx_rows(file_path)
    if suffix == ".csv":
        return _read_csv_rows(file_path)
    if suffix == ".xls":
        if _looks_like_xml_spreadsheet(file_path):
            return _read_excel_xml_rows(file_path)
        return _read_xls_rows(file_path)
    raise ValueError(f"Unsupported imported spreadsheet format: {file_path.suffix}")


def read_imported_hyperlinks(file_path: Path) -> dict[tuple[int, int], str]:
    suffix = file_path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return _read_xlsx_hyperlinks(file_path)
    if suffix == ".csv":
        return {}
    if suffix == ".xls":
        if _looks_like_xml_spreadsheet(file_path):
            return _read_excel_xml_hyperlinks(file_path)
        return _read_xls_hyperlinks(file_path)
    raise ValueError(f"Unsupported imported spreadsheet format: {file_path.suffix}")


def _normalize_row_length(values: list[object], width: int) -> list[object]:
    padded = list(values[:width])
    if len(padded) < width:
        padded.extend([None] * (width - len(padded)))
    return padded


def _normalize_test_id(value: object) -> str | None:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return format(value, "f").rstrip("0").rstrip(".")

    text = str(value).strip()
    if not text:
        return None
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def _get_last_queryresults_test_id(sheet) -> str | None:
    for row_index in range(sheet.max_row, 1, -1):
        test_id = _normalize_test_id(sheet.cell(row_index, 1).value)
        if test_id:
            return test_id
    return None


def _get_last_nonempty_row(sheet, key_column: int = 1) -> int:
    for row_index in range(sheet.max_row, 1, -1):
        value = sheet.cell(row_index, key_column).value
        if value not in (None, ""):
            return row_index
    return 1


def _delete_queryresults_orphan_link_rows(sheet) -> int:
    orphan_rows: list[int] = []
    for row_index in range(sheet.max_row, 1, -1):
        test_id = sheet.cell(row_index, 1).value
        if test_id not in (None, ""):
            continue

        linked_values = [sheet.cell(row_index, column_index).value for column_index in range(LINKED_DOCUMENT_START_COLUMN, LINKED_DOCUMENT_END_COLUMN + 1)]
        if not any(isinstance(value, str) and value.lower().startswith("http") for value in linked_values):
            continue

        other_values = [
            sheet.cell(row_index, column_index).value
            for column_index in range(1, sheet.max_column + 1)
            if column_index < LINKED_DOCUMENT_START_COLUMN or column_index > LINKED_DOCUMENT_END_COLUMN
        ]
        if any(value not in (None, "") for value in other_values):
            continue

        orphan_rows.append(row_index)

    for row_index in orphan_rows:
        sheet.delete_rows(row_index, 1)

    return len(orphan_rows)


def _to_float(value: object) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def _find_header_index(headers: list[object], header_name: str) -> int:
    for index, value in enumerate(headers, start=1):
        if value == header_name:
            return index
    raise ValueError(f"Could not find required QueryResults header: {header_name}")


def _classify_plot_series(test_name: object, testnet_title: object, test_procedure: object) -> str:
    test_name_text = str(test_name or "").lower()
    testnet_title_text = str(testnet_title or "").lower()
    test_procedure_text = str(test_procedure or "").lower()

    if "fcd" in testnet_title_text:
        return "FCD-Robot"
    if "robot" in test_name_text or "robot" in test_procedure_text:
        return "Overnight Soak-Robot"
    return "Overnight Soak-Driver"


def _normalize_plot_row_value(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    return str(value).strip()


def _capture_column_dimensions(sheet) -> dict[str, object]:
    return {key: copy(value) for key, value in sheet.column_dimensions.items()}


def _restore_column_dimensions(sheet, dimensions: dict[str, object]) -> None:
    for key, value in dimensions.items():
        sheet.column_dimensions[key] = copy(value)


def _apply_plot_sheet_template(plot_sheet, template_sheet) -> None:
    plot_sheet._charts = [deepcopy(chart) for chart in getattr(template_sheet, "_charts", [])]
    _restore_column_dimensions(plot_sheet, _capture_column_dimensions(template_sheet))
    plot_sheet.freeze_panes = template_sheet.freeze_panes


def _normalize_queryresults_font_size(
    sheet,
    font_size: float = 10.0,
    font_name: str = "Aptos Narrow",
) -> None:
    if sheet.title != "QueryResults":
        return

    for row_index in range(1, sheet.max_row + 1):
        for column_index in range(1, sheet.max_column + 1):
            cell = sheet.cell(row_index, column_index)
            updated_font = copy(cell.font)
            updated_font.sz = font_size
            updated_font.name = font_name
            cell.font = updated_font


def _normalize_plot_sheet_font(
    sheet,
    target_width: int,
    font_size: float = 10.0,
    font_name: str = "Aptos Narrow",
) -> None:
    if sheet.title not in {DEFAULT_PLOT_BASE_SHEET, DEFAULT_PLOT_OBD_SHEET}:
        return

    for row_index in range(1, sheet.max_row + 1):
        for column_index in range(1, target_width + 1):
            cell = sheet.cell(row_index, column_index)
            updated_font = copy(cell.font)
            updated_font.sz = font_size
            updated_font.name = font_name
            cell.font = updated_font


def _normalize_workbook_number_formats(workbook) -> int:
    updated_count = 0
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if not isinstance(cell.value, (int, float)) or isinstance(cell.value, bool):
                    continue
                current_format = str(cell.number_format or "")
                if "%" in current_format or any(token in current_format.lower() for token in ("yy", "dd", "mm", "h", "s")):
                    continue
                if current_format == "0.00":
                    continue
                cell.number_format = "0.00"
                updated_count += 1
    return updated_count


def _get_sheet_header_map(sheet) -> dict[str, int]:
    header_map: dict[str, int] = {}
    for column_index in range(1, sheet.max_column + 1):
        header_value = sheet.cell(1, column_index).value
        if header_value in (None, ""):
            continue
        header_text = str(header_value)
        if header_text not in header_map:
            header_map[header_text] = column_index
    return header_map


def _get_summary_block_header_map(sheet) -> dict[str, int]:
    header_map: dict[str, int] = {}
    start_column = sheet[f"{SUMMARY_START_COLUMN}1"].column
    for column_index in range(start_column, sheet.max_column + 1):
        header_value = sheet.cell(1, column_index).value
        if header_value in (None, ""):
            continue
        header_text = str(header_value).strip()
        if header_text not in header_map:
            header_map[header_text] = column_index
    return header_map


def _get_summary_value_columns(sheet) -> dict[str, int]:
    summary_header_map = _get_summary_block_header_map(sheet)
    resolved: dict[str, int] = {}
    for key, header_text in SUMMARY_VALUE_HEADERS.items():
        column_index = summary_header_map.get(header_text)
        if column_index is not None:
            resolved[key] = column_index
    return resolved


def _get_summary_block_headers(sheet) -> list[object]:
    start_column = sheet[f"{SUMMARY_START_COLUMN}1"].column
    headers: list[object] = []
    for column_index in range(start_column, sheet.max_column + 1):
        header_value = sheet.cell(1, column_index).value
        if header_value in (None, ""):
            continue
        headers.append(header_value)
    return headers


def _ensure_plot_sheet_columns(plot_sheet, target_headers: list[object]) -> dict[str, int]:
    header_map = _get_sheet_header_map(plot_sheet)
    next_column = max(plot_sheet.max_column, len(target_headers)) + 1

    for default_column, header in enumerate(target_headers, start=1):
        header_text = str(header)
        if header_text in header_map:
            continue

        while plot_sheet.cell(1, next_column).value not in (None, ""):
            next_column += 1

        plot_sheet.cell(1, next_column).value = header
        header_map[header_text] = next_column
        next_column += 1

        if plot_sheet.max_column < default_column and plot_sheet.cell(1, default_column).value in (None, ""):
            plot_sheet.cell(1, default_column).value = header
            header_map[header_text] = default_column

    return header_map


def _get_or_reset_plot_sheet(workbook, sheet_name: str, target_headers: list[object], template_sheet=None):
    if sheet_name in workbook.sheetnames:
        plot_sheet = workbook[sheet_name]
    else:
        plot_sheet = workbook.create_sheet(sheet_name)

    if template_sheet is not None:
        _apply_plot_sheet_template(plot_sheet, template_sheet)

    _ensure_plot_sheet_columns(plot_sheet, target_headers)

    return plot_sheet


def _clear_plot_sheet_data(plot_sheet, target_columns: list[int]) -> None:
    for row_index in range(2, plot_sheet.max_row + 1):
        for column_index in target_columns:
            cell = plot_sheet.cell(row_index, column_index)
            cell.value = None
            cell.hyperlink = None


def _collect_plot_row_signatures(plot_sheet, plot_header_map: dict[str, int], target_headers: list[object]) -> set[tuple[str, ...]]:
    signatures: set[tuple[str, ...]] = set()
    target_columns = [plot_header_map[str(header)] for header in target_headers if str(header) in plot_header_map]
    if not target_columns:
        return signatures

    key_column = target_columns[0]
    for row_index in range(2, plot_sheet.max_row + 1):
        if plot_sheet.cell(row_index, key_column).value in (None, ""):
            continue
        signature = tuple(
            _normalize_plot_row_value(plot_sheet.cell(row_index, column_index).value)
            for column_index in target_columns
        )
        signatures.add(signature)
    return signatures


def _snapshot_plot_sheet_rows(plot_sheet, visible_column_count: int = 32) -> list[tuple[str, ...]]:
    snapshot: list[tuple[str, ...]] = []
    for row_index in range(2, plot_sheet.max_row + 1):
        if plot_sheet.cell(row_index, 1).value in (None, ""):
            continue
        snapshot.append(
            tuple(
                _normalize_plot_row_value(plot_sheet.cell(row_index, column_index).value)
                for column_index in range(1, visible_column_count + 1)
            )
        )
    return snapshot


def _assert_append_preserved_plot_rows(before_rows: list[tuple[str, ...]], after_rows: list[tuple[str, ...]], sheet_name: str) -> None:
    if len(after_rows) < len(before_rows):
        raise RuntimeError(f"Append mode removed plot rows from {sheet_name}.")
    if after_rows[: len(before_rows)] != before_rows:
        raise RuntimeError(f"Append mode changed existing plot row order or values in {sheet_name}.")


def _is_limit_catalyst(plot_identifier: str) -> bool:
    normalized_identifier = " ".join(plot_identifier.lower().split())
    return any(
        marker in normalized_identifier
        for marker in ("threshold baseline", "catmn dcms thres data collection")
    )


def _select_plot_sheet_name(plot_identifier: str) -> str:
    normalized_identifier = plot_identifier.lower()
    if _is_limit_catalyst(plot_identifier):
        return DEFAULT_PLOT_OBD_SHEET
    if "baseline" in normalized_identifier:
        return DEFAULT_PLOT_BASE_SHEET
    return DEFAULT_PLOT_OBD_SHEET


def _build_catmn_sheet(workbook, query_rows: list[int] | None = None, clear_existing: bool = True) -> None:
    query_sheet = workbook["QueryResults"]
    header_map = _get_queryresults_header_map(query_sheet)
    summary_columns = _get_summary_value_columns(query_sheet)
    monitor_column = summary_columns.get("monitor")
    sheet_exists = DEFAULT_CATMN_SHEET in workbook.sheetnames
    catmn_sheet = workbook[DEFAULT_CATMN_SHEET] if sheet_exists else workbook.create_sheet(DEFAULT_CATMN_SHEET)
    source_column_count = max(header_map.values(), default=1)
    test_id_column = header_map.get("TESTID") or summary_columns.get("test_id")
    if test_id_column is None:
        raise ValueError("Missing TESTID or Test ID column in QueryResults.")

    if clear_existing:
        for cells in catmn_sheet.iter_rows(min_row=2, max_col=source_column_count):
            for cell in cells:
                cell.value = None
                cell.hyperlink = None

    for source_cell in query_sheet[1][:source_column_count]:
        target_cell = catmn_sheet.cell(1, source_cell.column, source_cell.value)
        target_cell._style = copy(source_cell._style)
        column_letter = get_column_letter(source_cell.column)
        if column_letter in query_sheet.column_dimensions:
            catmn_sheet.column_dimensions[column_letter].width = query_sheet.column_dimensions[column_letter].width

    existing_ids = {
        _normalize_plot_row_value(catmn_sheet.cell(row, test_id_column).value)
        for row in range(2, catmn_sheet.max_row + 1)
        if catmn_sheet.cell(row, test_id_column).value not in (None, "")
    }
    next_row = _get_last_nonempty_row(catmn_sheet, test_id_column) + 1
    source_rows = query_rows if query_rows is not None and sheet_exists and not clear_existing else range(2, query_sheet.max_row + 1)
    for row in source_rows:
        if row < 2 or row > query_sheet.max_row:
            continue
        identifier = _build_row_plot_identifier(query_sheet, row, header_map)
        monitor_text = str(query_sheet.cell(row, monitor_column).value or "") if monitor_column is not None else ""
        if not _is_limit_catalyst(f"{identifier} {monitor_text}"):
            continue
        test_id = _normalize_plot_row_value(query_sheet.cell(row, test_id_column).value)
        if not test_id or test_id in existing_ids:
            continue
        for source_cell in query_sheet[row][:source_column_count]:
            target_cell = catmn_sheet.cell(next_row, source_cell.column, source_cell.value)
            target_cell._style = copy(source_cell._style)
            if source_cell.hyperlink:
                target_cell.hyperlink = copy(source_cell.hyperlink)
            if source_cell.comment:
                target_cell.comment = copy(source_cell.comment)
        existing_ids.add(test_id)
        next_row += 1
    catmn_sheet.freeze_panes = "A2"
    catmn_sheet.auto_filter.ref = f"A1:{get_column_letter(source_column_count)}{max(1, next_row - 1)}"


def _build_emis_plot_sheets(workbook, standards: dict[str, float], query_rows: list[int] | None = None, clear_existing: bool = True) -> None:
    query_sheet = workbook["QueryResults"]
    _ = standards
    template_workbook = None
    query_headers = [query_sheet.cell(1, column).value for column in range(1, query_sheet.max_column + 1)]
    query_header_map = _get_sheet_header_map(query_sheet)
    summary_value_columns = _get_summary_value_columns(query_sheet)
    if not {"date", "test_id", "co", "co_pct", "nox", "nox_pct", "nmhc", "nmhc_pct"}.issubset(summary_value_columns):
        missing = sorted({"date", "test_id", "co", "co_pct", "nox", "nox_pct", "nmhc", "nmhc_pct"} - set(summary_value_columns))
        raise ValueError(f"Missing summary columns in QueryResults: {', '.join(missing)}")
    optional_plot_headers = ("TESTNET_TITLE",)
    tr_number_column = _find_header_index(query_headers, "TR_NUMBER")
    test_name_column = _find_header_index(query_headers, "TEST_NAME")
    test_procedure_column = _find_header_index(query_headers, "TEST_PROCEDURE")
    testnet_title_column = _find_header_index(query_headers, "TESTNET_TITLE")
    source_headers = _get_summary_block_headers(query_sheet)
    source_columns = [_get_summary_block_header_map(query_sheet)[str(header).strip()] for header in source_headers]
    target_headers = [source_headers[0], "TR_NUMBER", *source_headers[1:]]
    target_width = len(target_headers)
    template_path = Path(DEFAULT_WORKBOOK_TEMPLATE)
    if clear_existing and template_path.exists():
        template_workbook = load_workbook(template_path)
    plot_sheets = {
        DEFAULT_PLOT_BASE_SHEET: _get_or_reset_plot_sheet(
            workbook,
            DEFAULT_PLOT_BASE_SHEET,
            target_headers,
            template_workbook[DEFAULT_PLOT_BASE_SHEET] if template_workbook and DEFAULT_PLOT_BASE_SHEET in template_workbook.sheetnames else None,
        ),
        DEFAULT_PLOT_OBD_SHEET: _get_or_reset_plot_sheet(
            workbook,
            DEFAULT_PLOT_OBD_SHEET,
            target_headers,
            template_workbook[DEFAULT_PLOT_OBD_SHEET] if template_workbook and DEFAULT_PLOT_OBD_SHEET in template_workbook.sheetnames else None,
        ),
    }
    plot_header_maps = {
        sheet_name: _ensure_plot_sheet_columns(plot_sheet, target_headers)
        for sheet_name, plot_sheet in plot_sheets.items()
    }

    if "Emis_Plot" in workbook.sheetnames:
        del workbook["Emis_Plot"]

    workbook.active = workbook.sheetnames.index(DEFAULT_PLOT_BASE_SHEET)

    existing_signatures: dict[str, set[tuple[str, ...]]] = {
        sheet_name: set() for sheet_name in plot_sheets
    }
    next_rows: dict[str, int] = {}
    for sheet_name, plot_sheet in plot_sheets.items():
        target_columns = [plot_header_maps[sheet_name][str(header)] for header in target_headers]
        target_columns.extend(
            plot_header_maps[sheet_name][header]
            for header in optional_plot_headers
            if header in plot_header_maps[sheet_name]
        )
        if clear_existing:
            _clear_plot_sheet_data(plot_sheet, target_columns)
            next_rows[sheet_name] = 2
        else:
            existing_signatures[sheet_name] = _collect_plot_row_signatures(plot_sheet, plot_header_maps[sheet_name], target_headers)
            next_rows[sheet_name] = _get_last_nonempty_row(plot_sheet, plot_header_maps[sheet_name][str(target_headers[0])]) + 1

    target_date_header = str(target_headers[0])
    target_co_header = str(query_sheet.cell(1, summary_value_columns["co"]).value)
    target_co_pct_header = str(query_sheet.cell(1, summary_value_columns["co_pct"]).value)
    target_nox_header = str(query_sheet.cell(1, summary_value_columns["nox"]).value)
    target_nox_pct_header = str(query_sheet.cell(1, summary_value_columns["nox_pct"]).value)
    target_nmhc_header = str(query_sheet.cell(1, summary_value_columns["nmhc"]).value)
    target_nmhc_pct_header = str(query_sheet.cell(1, summary_value_columns["nmhc_pct"]).value)
    source_row_indexes = query_rows if query_rows is not None else list(range(2, query_sheet.max_row + 1))
    for row_index in source_row_indexes:
        if row_index < 2 or row_index > query_sheet.max_row:
            continue
        az_value = query_sheet.cell(row_index, summary_value_columns["date"]).value
        ba_value = query_sheet.cell(row_index, summary_value_columns["test_id"]).value
        if az_value in (None, "") or ba_value in (None, ""):
            continue

        plot_sheet_name = (
            DEFAULT_PLOT_BASE_SHEET
            if _row_uses_base_standards(query_sheet, row_index, query_header_map, summary_value_columns)
            else DEFAULT_PLOT_OBD_SHEET
        )
        plot_sheet = plot_sheets[plot_sheet_name]

        row_values = [
            query_sheet.cell(row_index, source_columns[0]).value,
            query_sheet.cell(row_index, tr_number_column).value,
            *[query_sheet.cell(row_index, column_index).value for column_index in source_columns[1:]],
        ]
        row_value_map = {
            str(header): value
            for header, value in zip(target_headers, row_values)
        }
        row_signature = tuple(_normalize_plot_row_value(value) for value in row_values)
        if row_signature in existing_signatures[plot_sheet_name]:
            continue

        next_row = next_rows[plot_sheet_name]
        standard_type = "base" if plot_sheet_name == DEFAULT_PLOT_BASE_SHEET else "obd"
        plot_header_map = plot_header_maps[plot_sheet_name]
        target_date_column = plot_header_map[target_date_header]
        target_co_column = plot_header_map[target_co_header]
        target_co_pct_column = plot_header_map[target_co_pct_header]
        target_nox_column = plot_header_map[target_nox_header]
        target_nox_pct_column = plot_header_map[target_nox_pct_header]
        target_nmhc_column = plot_header_map[target_nmhc_header]
        target_nmhc_pct_column = plot_header_map[target_nmhc_pct_header]
        plot_standard_cells = _get_plot_standard_cells(plot_sheet_name)
        for header, value in zip(target_headers, row_values):
            cell = plot_sheet.cell(next_row, plot_header_map[str(header)])
            cell.value = value
            if cell.column == target_date_column:
                cell.number_format = "yyyy-mm-dd"

        for header in optional_plot_headers:
            if header not in plot_header_map or header not in query_header_map:
                continue
            plot_sheet.cell(next_row, plot_header_map[header]).value = query_sheet.cell(
                row_index,
                query_header_map[header],
            ).value

        plot_sheet.cell(next_row, target_co_pct_column).value = _build_plot_percentage_formula(
            target_co_column,
            plot_standard_cells["co"],
            next_row,
        )
        plot_sheet.cell(next_row, target_nox_pct_column).value = _build_plot_percentage_formula(
            target_nox_column,
            plot_standard_cells["nox"],
            next_row,
        )
        plot_sheet.cell(next_row, target_nmhc_pct_column).value = _build_plot_percentage_formula(
            target_nmhc_column,
            plot_standard_cells["nmhc"],
            next_row,
        )
        plot_sheet.cell(next_row, target_co_pct_column).number_format = "0.00%"
        plot_sheet.cell(next_row, target_nox_pct_column).number_format = "0.00%"
        plot_sheet.cell(next_row, target_nmhc_pct_column).number_format = "0.00%"

        existing_signatures[plot_sheet_name].add(row_signature)
        _apply_ratio_alert_fill(
            plot_sheet.cell(next_row, target_co_pct_column),
            (row_value_map.get(target_co_header) / standards[standard_type]["co"])
            if isinstance(row_value_map.get(target_co_header), (int, float)) and standards[standard_type]["co"] not in (None, 0)
            else None,
        )
        _apply_ratio_alert_fill(
            plot_sheet.cell(next_row, target_nox_pct_column),
            (row_value_map.get(target_nox_header) / standards[standard_type]["nox"])
            if isinstance(row_value_map.get(target_nox_header), (int, float)) and standards[standard_type]["nox"] not in (None, 0)
            else None,
        )
        _apply_ratio_alert_fill(
            plot_sheet.cell(next_row, target_nmhc_pct_column),
            (row_value_map.get(target_nmhc_header) / standards[standard_type]["nmhc"])
            if isinstance(row_value_map.get(target_nmhc_header), (int, float)) and standards[standard_type]["nmhc"] not in (None, 0)
            else None,
        )
        next_rows[plot_sheet_name] += 1

    for sheet_name, plot_sheet in plot_sheets.items():
        _trim_plot_charts(plot_sheet)
        _restore_plot_chart_titles(plot_sheet)
        _extend_plot_chart_ranges(plot_sheet, next_rows[sheet_name] - 1)
        standard_type = "base" if sheet_name == DEFAULT_PLOT_BASE_SHEET else "obd"
        for chart in getattr(plot_sheet, "_charts", []):
            _style_threshold_chart(plot_sheet, chart, next_rows[sheet_name] - 1, standard_type, standards)
        if clear_existing:
            _normalize_plot_sheet_font(plot_sheet, target_width)
        plot_sheet.freeze_panes = "A2"
        target_date_column = plot_header_maps[sheet_name][target_date_header]
        target_co_pct_column = plot_header_maps[sheet_name][target_co_pct_header]
        target_nox_pct_column = plot_header_maps[sheet_name][target_nox_pct_header]
        target_nmhc_pct_column = plot_header_maps[sheet_name][target_nmhc_pct_header]
        format_start_row = 2 if clear_existing else max(2, next_rows[sheet_name] - len(source_row_indexes))
        for row_index in range(format_start_row, plot_sheet.max_row + 1):
            plot_sheet.cell(row_index, target_date_column).number_format = "yyyy-mm-dd"
            plot_sheet.cell(row_index, target_co_pct_column).number_format = "0.00%"
            plot_sheet.cell(row_index, target_nox_pct_column).number_format = "0.00%"
            plot_sheet.cell(row_index, target_nmhc_pct_column).number_format = "0.00%"

    if template_workbook is not None:
        template_workbook.close()

    _build_catmn_sheet(workbook, query_rows=query_rows, clear_existing=clear_existing)


def _get_existing_queryresults_test_ids(sheet) -> set[str]:
    existing_test_ids: set[str] = set()
    for row_index in range(2, sheet.max_row + 1):
        test_id = _normalize_test_id(sheet.cell(row_index, 1).value)
        if test_id:
            existing_test_ids.add(test_id)
    return existing_test_ids


def _filter_rows_after_test_id(data_rows: list[list[object]], last_test_id: str | None) -> list[list[object]]:
    if not last_test_id:
        return data_rows

    if not last_test_id.isdigit():
        return []

    threshold = int(last_test_id)
    filtered_rows: list[list[object]] = []
    for row_values in data_rows:
        test_id = _normalize_test_id(row_values[0] if row_values else None)
        if test_id is None or not test_id.isdigit():
            continue
        if int(test_id) > threshold:
            filtered_rows.append(_normalize_row_length(row_values, 51))

    return filtered_rows


def _select_rows_after_test_id(data_rows: list[list[object]], last_test_id: str | None) -> list[int]:
    if not last_test_id:
        return list(range(len(data_rows)))

    if not last_test_id.isdigit():
        return []

    threshold = int(last_test_id)
    selected_rows: list[int] = []
    for index, row_values in enumerate(data_rows):
        test_id = _normalize_test_id(row_values[0] if row_values else None)
        if test_id is None or not test_id.isdigit():
            continue
        if int(test_id) > threshold:
            selected_rows.append(index)

    return selected_rows


def _select_rows_with_missing_test_ids(data_rows: list[list[object]], existing_test_ids: set[str]) -> list[int]:
    selected_rows: list[int] = []
    seen_new_test_ids: set[str] = set()

    for index, row_values in enumerate(data_rows):
        test_id = _normalize_test_id(row_values[0] if row_values else None)
        if test_id is None:
            continue
        if test_id in existing_test_ids or test_id in seen_new_test_ids:
            continue
        selected_rows.append(index)
        seen_new_test_ids.add(test_id)

    return selected_rows


def _find_header_row(rows: list[list[object]]) -> int:
    for index, row in enumerate(rows):
        normalized = [str(value).strip() if value is not None else "" for value in row[:5]]
        if normalized[:5] == ["TESTID", "VIN", "TR_NUMBER", "START_DATE", "END_DATE"]:
            return index
    raise ValueError("Could not find the expected header row in the imported spreadsheet.")


def _clean_pdf_text(text: str) -> str:
    return text.replace("\u2212", "-").replace("\u2013", "-")


def _find_line(text: str, prefix: str) -> str:
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith(prefix):
            return stripped
    match = re.search(rf"{re.escape(prefix)}\s+(?:-?\d+(?:\.\d+)?\s*)+", text)
    if match:
        return match.group(0).strip()
    raise ValueError(f"Could not find line starting with '{prefix}'.")


def _parse_numeric_line(line: str) -> list[float]:
    return [float(value) for value in re.findall(r"-?\d+(?:\.\d+)?", line)]


def _extract_post_test_remarks(text: str) -> str | None:
    lines = [line.strip() for line in text.splitlines()]
    for index, line in enumerate(lines):
        if line == "Post Test Remarks:" and index + 1 < len(lines):
            remark = lines[index + 1].strip()
            return remark or None
    return None


def _extract_void_flag(text: str) -> int | None:
    match = re.search(r"APTL_VOIDORRESTART\s*=\s*(\d+)", text)
    if not match:
        return None
    return int(match.group(1))


def _extract_cal_file(text: str) -> str | None:
    match = re.search(r"([A-Za-z0-9_\-]+\.cal)", text, re.IGNORECASE)
    if not match:
        return None
    return match.group(1)


def parse_summary_pdf(pdf_path: Path) -> dict[str, object]:
    reader = PdfReader(str(pdf_path))
    cleaned_pages = [_clean_pdf_text(page.extract_text() or "") for page in reader.pages]
    summary_pages = [page_text for page_text in cleaned_pages if "SUMMARY REPORT" in page_text]

    if len(summary_pages) < 2:
        raise ValueError(f"Expected summary pages in {pdf_path.name}, found {len(summary_pages)}.")

    summary_text = "\n".join(summary_pages)
    full_text = "\n".join(cleaned_pages)

    test_id_match = re.search(r"Test Id\s*=\s*(\d+)", summary_text)
    if not test_id_match:
        raise ValueError(f"Could not find test id in {pdf_path.name}.")

    test_date_match = re.search(r"Test Init Start\s*=\s*(.+?)\s+Fuel Calculation Type", summary_text)
    if not test_date_match:
        raise ValueError(f"Could not find test date in {pdf_path.name}.")

    test_date = datetime.strptime(test_date_match.group(1).strip(), "%d %B %Y %H:%M:%S").date()

    total_mass_line = _find_line(summary_text, "|T. mg+")
    total_distance_line = _find_line(summary_text, "|mg/km+")

    total_mass_values = _parse_numeric_line(total_mass_line)
    total_distance_values = _parse_numeric_line(total_distance_line)

    if len(total_mass_values) < 9 or len(total_distance_values) < 8:
        raise ValueError(f"Unexpected summary totals in {pdf_path.name}.")

    return {
        "date": test_date,
        "test_id": test_id_match.group(1),
        "thc": total_distance_values[0],
        "co": total_distance_values[1],
        "nox": total_distance_values[2],
        "co2": total_distance_values[3],
        "ch4": total_distance_values[4],
        "nmhc": total_distance_values[5],
        "hc_nox": total_distance_values[6],
        "soot": total_distance_values[7],
        "fe": total_mass_values[8],
        "pm": None,
        "cal_file": _extract_cal_file(full_text),
        "comments": _extract_post_test_remarks(cleaned_pages[0]) if cleaned_pages else None,
        "void": _extract_void_flag(full_text),
    }


def get_eu7_standards(workbook_path: Path) -> dict[str, dict[str, float | None]]:
    workbook = load_workbook(workbook_path, data_only=True)
    try:
        sheet = workbook["EU7 Regulation"]
        return {
            "base": {
                "co": float(sheet["D2"].value),
                "nox": float(sheet["E2"].value),
                "nmhc": float(sheet["F2"].value),
            },
            "obd": {
                "co": float(sheet["D4"].value),
                "nox": float(sheet["E4"].value),
                "nmhc": float(sheet["F4"].value),
            },
        }
    finally:
        workbook.close()


def calculate_percentage(value: float | None, standard: float | None) -> float | None:
    if value is None or standard in (None, 0):
        return None
    return round((value / standard) * 100, 2)


def _apply_percentage_alert_fill(cell) -> None:
    value = cell.value
    if isinstance(value, (int, float)) and value >= 100:
        cell.fill = PERCENT_ALERT_FILL
        return
    if isinstance(value, (int, float)) and value >= 90:
        cell.fill = PERCENT_WARNING_FILL
        return
    cell.fill = DEFAULT_FILL


def _apply_ratio_alert_fill(cell, ratio: float | None) -> None:
    if ratio is None:
        cell.fill = DEFAULT_FILL
        return
    if ratio >= 1:
        cell.fill = PERCENT_ALERT_FILL
        return
    if ratio >= 0.9:
        cell.fill = PERCENT_WARNING_FILL
        return
    cell.fill = DEFAULT_FILL


def _build_plot_percentage_formula(value_column: int, regulation_cell: str, row_index: int) -> str:
    regulation_column = regulation_cell[0]
    regulation_row = regulation_cell[1:]
    return f"={get_column_letter(value_column)}{row_index}/'EU7 Regulation'!${regulation_column}${regulation_row}"


def _get_plot_standard_cells(plot_sheet_name: str) -> dict[str, str]:
    if plot_sheet_name == DEFAULT_PLOT_BASE_SHEET:
        return {"co": "D2", "nox": "E2", "nmhc": "F2"}
    return {"co": "D4", "nox": "E4", "nmhc": "F4"}


def _update_chart_range_formula(formula: str, sheet_name: str, last_row: int) -> str:
    escaped_sheet_name = re.escape(sheet_name)
    pattern = rf"({escaped_sheet_name}!\$[A-Z]+\$)\d+(:\$[A-Z]+\$)\d+"
    return re.sub(pattern, rf"\g<1>2\g<2>{last_row}", formula)


def _extend_plot_chart_ranges(plot_sheet, last_row: int) -> None:
    target_last_row = max(last_row, 2)
    for chart in getattr(plot_sheet, "_charts", []):
        for series in getattr(chart, "ser", []):
            x_value_ref = getattr(getattr(series, "xVal", None), "numRef", None)
            if x_value_ref is not None and x_value_ref.f:
                x_value_ref.f = _update_chart_range_formula(x_value_ref.f, plot_sheet.title, target_last_row)
                x_value_ref.numCache = None

            y_value_ref = getattr(getattr(series, "yVal", None), "numRef", None)
            if y_value_ref is not None and y_value_ref.f:
                y_value_ref.f = _update_chart_range_formula(y_value_ref.f, plot_sheet.title, target_last_row)
                y_value_ref.numCache = None


def _trim_plot_charts(plot_sheet, max_charts: int = 6) -> None:
    charts = list(getattr(plot_sheet, "_charts", []))
    if len(charts) <= max_charts:
        return
    plot_sheet._charts = charts[:max_charts]


def _get_series_column_from_formula(formula: str) -> int | None:
    match = re.search(r"\$([A-Z]+)\$\d+(?::\$[A-Z]+\$\d+)?$", formula)
    if not match:
        return None
    return column_index_from_string(match.group(1))


def _get_percent_chart_standard_key(y_header: object) -> str | None:
    normalized_header = str(y_header or "").strip().upper()
    mapping = {
        "CO %": "co",
        "NOX %": "nox",
        "NM-HC %": "nmhc",
    }
    return mapping.get(normalized_header)


def _get_chart_standard_key(y_header: object) -> tuple[str | None, bool]:
    normalized_header = str(y_header or "").strip().upper()
    mapping = {
        "CO": ("co", False),
        "CO %": ("co", True),
        "NOX": ("nox", False),
        "NOX %": ("nox", True),
        "NM-HC": ("nmhc", False),
        "NM-HC %": ("nmhc", True),
    }
    return mapping.get(normalized_header, (None, False))


def _ensure_threshold_helper_column(plot_sheet, last_row: int, helper_key: str, threshold_value: float) -> int:
    helper_columns = {
        "co_raw_threshold": 16384,
        "co_raw_warning": 16383,
        "co_raw_exceed": 16382,
        "co_raw_ninety": 16381,
        "co_pct_threshold": 16380,
        "co_pct_warning": 16379,
        "co_pct_exceed": 16378,
        "co_pct_ninety": 16377,
        "nox_raw_threshold": 16376,
        "nox_raw_warning": 16375,
        "nox_raw_exceed": 16374,
        "nox_raw_ninety": 16373,
        "nox_pct_threshold": 16372,
        "nox_pct_warning": 16371,
        "nox_pct_exceed": 16370,
        "nox_pct_ninety": 16369,
        "nmhc_raw_threshold": 16368,
        "nmhc_raw_warning": 16367,
        "nmhc_raw_exceed": 16366,
        "nmhc_raw_ninety": 16365,
        "nmhc_pct_threshold": 16364,
        "nmhc_pct_warning": 16363,
        "nmhc_pct_exceed": 16362,
        "nmhc_pct_ninety": 16361,
    }
    helper_column = helper_columns[helper_key]
    plot_sheet.cell(1, helper_column).value = helper_key
    target_last_row = max(last_row, 2)
    for row_index in range(2, max(plot_sheet.max_row, target_last_row) + 1):
        plot_sheet.cell(row_index, helper_column).value = threshold_value if row_index <= target_last_row else None
    plot_sheet.column_dimensions[get_column_letter(helper_column)].hidden = True
    return helper_column


def _style_threshold_chart(plot_sheet, chart, last_row: int, standard_type: str, standards: dict[str, dict[str, float | None]]) -> None:
    if not getattr(chart, "ser", None):
        return

    chart.scatterStyle = "marker"
    chart.visible_cells_only = False
    chart.display_blanks = "span"

    primary_series = chart.ser[0]
    y_value_ref = getattr(getattr(primary_series, "yVal", None), "numRef", None)
    if y_value_ref is None or not y_value_ref.f:
        return

    y_column = _get_series_column_from_formula(y_value_ref.f)
    if y_column is None:
        return

    standard_key, is_percent_chart = _get_chart_standard_key(plot_sheet.cell(1, y_column).value)
    if standard_key is None:
        return

    if is_percent_chart:
        chart.y_axis.numFmt.formatCode = "0%"
        chart.y_axis.numFmt.sourceLinked = False
    else:
        chart.y_axis.numFmt.formatCode = "0"
        chart.y_axis.numFmt.sourceLinked = False

    standard_value = standards[standard_type].get(standard_key)
    if is_percent_chart:
        chart.y_axis.scaling.max = 1.5
    elif standard_value not in (None, 0):
        chart.y_axis.scaling.max = float(standard_value) * 1.5

    del chart.ser[1:]
    primary_series.dPt = []
    primary_series.spPr = GraphicalProperties(ln=LineProperties(noFill=True))
    primary_series.marker = Marker(symbol="circle", size=5)

    threshold_value = 1.0 if is_percent_chart else float(standard_value)
    helper_column = _ensure_threshold_helper_column(
        plot_sheet,
        last_row,
        f"{standard_key}_{'pct' if is_percent_chart else 'raw'}_threshold",
        threshold_value,
    )
    threshold_series = Series(
        Reference(plot_sheet, min_col=helper_column, min_row=2, max_row=max(last_row, 2)),
        xvalues=Reference(plot_sheet, min_col=3, min_row=2, max_row=max(last_row, 2)),
        title="100%" if is_percent_chart else str(int(threshold_value) if float(threshold_value).is_integer() else threshold_value),
    )
    threshold_series.marker = Marker(symbol="none")
    threshold_series.spPr = GraphicalProperties(
        ln=LineProperties(w=28575, solidFill="000000", prstDash="sysDot")
    )
    chart.ser.append(threshold_series)

    ninety_value = threshold_value * 0.9
    ninety_column = _ensure_threshold_helper_column(
        plot_sheet,
        last_row,
        f"{standard_key}_{'pct' if is_percent_chart else 'raw'}_ninety",
        ninety_value,
    )
    ninety_series = Series(
        Reference(plot_sheet, min_col=ninety_column, min_row=2, max_row=max(last_row, 2)),
        xvalues=Reference(plot_sheet, min_col=3, min_row=2, max_row=max(last_row, 2)),
        title=None,
    )
    ninety_series.marker = Marker(symbol="none")
    ninety_series.spPr = GraphicalProperties(
        ln=LineProperties(w=28575, solidFill="7F7F7F", prstDash="sysDot")
    )
    chart.ser.append(ninety_series)

    warning_helper_key = f"{standard_key}_{'pct' if is_percent_chart else 'raw'}_warning"
    exceed_helper_key = f"{standard_key}_{'pct' if is_percent_chart else 'raw'}_exceed"
    warning_column = _ensure_threshold_helper_column(plot_sheet, last_row, warning_helper_key, None)
    exceed_column = _ensure_threshold_helper_column(plot_sheet, last_row, exceed_helper_key, None)
    hidden_primary_points: list[DataPoint] = []

    for point_index, row_index in enumerate(range(2, last_row + 1)):
        raw_value = _to_float(plot_sheet.cell(row_index, y_column - 1).value) if is_percent_chart else _to_float(plot_sheet.cell(row_index, y_column).value)
        if raw_value is None or standard_value in (None, 0):
            plot_sheet.cell(row_index, warning_column).value = None
            plot_sheet.cell(row_index, exceed_column).value = None
            continue

        ratio = raw_value / standard_value
        if ratio >= 1:
            plot_sheet.cell(row_index, warning_column).value = None
            plot_sheet.cell(row_index, exceed_column).value = plot_sheet.cell(row_index, y_column).value
            hidden_primary_points.append(
                DataPoint(
                    idx=point_index,
                    marker=Marker(symbol="none"),
                    spPr=GraphicalProperties(ln=LineProperties(noFill=True)),
                )
            )
        elif ratio >= 0.9:
            plot_sheet.cell(row_index, warning_column).value = plot_sheet.cell(row_index, y_column).value
            plot_sheet.cell(row_index, exceed_column).value = None
            hidden_primary_points.append(
                DataPoint(
                    idx=point_index,
                    marker=Marker(symbol="none"),
                    spPr=GraphicalProperties(ln=LineProperties(noFill=True)),
                )
            )
        else:
            plot_sheet.cell(row_index, warning_column).value = None
            plot_sheet.cell(row_index, exceed_column).value = None

    primary_series.dPt = hidden_primary_points

    for helper_column_value, marker_color in ((warning_column, "FFD966"), (exceed_column, "FF0000")):
        marker_series = Series(
            Reference(plot_sheet, min_col=helper_column_value, min_row=2, max_row=max(last_row, 2)),
            xvalues=Reference(plot_sheet, min_col=3, min_row=2, max_row=max(last_row, 2)),
            title=None,
        )
        marker_series.spPr = GraphicalProperties(solidFill=marker_color, ln=LineProperties(noFill=True))
        marker_series.marker = Marker(
            symbol="circle",
            size=6,
            spPr=GraphicalProperties(solidFill=marker_color, ln=LineProperties(noFill=True)),
        )
        chart.ser.append(marker_series)


def _restore_plot_chart_titles(plot_sheet) -> None:
    for chart in getattr(plot_sheet, "_charts", []):
        title = None
        series_list = getattr(chart, "ser", [])
        if series_list:
            y_value_ref = getattr(getattr(series_list[0], "yVal", None), "numRef", None)
            if y_value_ref is not None and y_value_ref.f:
                match = re.search(r"\$([A-Z]+)\$\d+(?::\$[A-Z]+\$\d+)?$", y_value_ref.f)
                if match:
                    title = plot_sheet.cell(1, column_index_from_string(match.group(1))).value

        if title in (None, ""):
            title = "Chart"

        chart.title = str(title)


def _get_queryresults_header_map(sheet) -> dict[str, int]:
    return {
        str(value): column_index
        for column_index, value in enumerate((sheet.cell(1, column).value for column in range(1, sheet.max_column + 1)), start=1)
        if value not in (None, "")
    }


def _build_row_plot_identifier(sheet, row: int, header_map: dict[str, int]) -> str:
    return " ".join(
        str(value)
        for value in (
            sheet.cell(row, header_map[header_name]).value
            for header_name in ("TESTNET_TITLE", "TEST_NAME", "TEST_PROCEDURE")
            if header_name in header_map
        )
        if value not in (None, "")
    )


def _row_uses_base_standards(
    sheet,
    row: int,
    header_map: dict[str, int],
    summary_value_columns: dict[str, int] | None = None,
) -> bool:
    plot_identifier = _build_row_plot_identifier(sheet, row, header_map)
    monitor_column: int | None = None
    if summary_value_columns is not None:
        monitor_column = summary_value_columns.get("monitor")
    if monitor_column is None:
        monitor_column = _get_summary_value_columns(sheet).get("monitor")
    monitor_text = str(sheet.cell(row, monitor_column).value or "") if monitor_column is not None else ""
    return _select_plot_sheet_name(f"{plot_identifier} {monitor_text}") == DEFAULT_PLOT_BASE_SHEET


def _get_emission_standard_type(sheet, row: int, header_map: dict[str, int]) -> str:
    if _row_uses_base_standards(sheet, row, header_map):
        return "base"
    return "obd"


def _apply_emission_percentages(
    summary: dict[str, object],
    standard_type: str,
    standards: dict[str, dict[str, float | None]],
) -> None:
    selected_standards = standards[standard_type]
    summary["co_pct"] = calculate_percentage(summary.get("co"), selected_standards.get("co"))
    summary["nox_pct"] = calculate_percentage(summary.get("nox"), selected_standards.get("nox"))
    summary["nmhc_pct"] = calculate_percentage(summary.get("nmhc"), selected_standards.get("nmhc"))
    summary["pm_pct"] = None


def _write_summary_to_row(sheet, row: int, summary: dict[str, object]) -> None:
    for key, column_index in _get_summary_value_columns(sheet).items():
        if key not in summary:
            continue
        cell = sheet.cell(row, column_index)
        cell.value = summary.get(key)
        if key in PERCENT_ALERT_KEYS:
            _apply_percentage_alert_fill(cell)


def apply_pdf_summary_to_row(sheet, row: int, pdf_dir: Path, standards: dict[str, dict[str, float | None]]) -> bool:
    pdf_name = sheet[f"AA{row}"].value
    if not pdf_name:
        return False

    pdf_path = pdf_dir / str(pdf_name)
    if not pdf_path.exists():
        return False
    if not _has_expected_pdf_page_count(pdf_path):
        return False

    summary = parse_summary_pdf(pdf_path)
    header_map = _get_queryresults_header_map(sheet)
    standard_type = _get_emission_standard_type(sheet, row, header_map)
    _apply_emission_percentages(summary, standard_type, standards)

    sheet[f"AY{row}"] = summary.get("comments")

    _write_summary_to_row(sheet, row, summary)

    return True


def get_rows_with_pdf_and_csv(workbook_path: Path) -> list[int]:
    workbook = load_workbook(workbook_path, data_only=False, read_only=True)
    try:
        sheet = workbook["QueryResults"]
        rows: list[int] = []

        for row in range(2, sheet.max_row + 1):
            if sheet[f"AA{row}"].value and sheet[f"AU{row}"].value:
                rows.append(row)

        return rows
    finally:
        workbook.close()


def get_subset_rows_with_pdf_and_csv(workbook_path: Path, rows: list[int]) -> list[int]:
    if not rows:
        return []

    workbook = load_workbook(workbook_path, data_only=False, read_only=True)
    try:
        sheet = workbook["QueryResults"]
        matched_rows: list[int] = []
        for row in rows:
            if sheet[f"AA{row}"].value and sheet[f"AU{row}"].value:
                matched_rows.append(row)
        return matched_rows
    finally:
        workbook.close()


def get_rows_with_incomplete_summary_pdf_and_csv(workbook_path: Path) -> list[int]:
    workbook = load_workbook(workbook_path, data_only=False, read_only=True)
    try:
        sheet = workbook["QueryResults"]
        summary_value_columns = _get_summary_value_columns(sheet)
        required_summary_columns = [
            summary_value_columns[key]
            for key in ("date", "test_id", "co", "nox", "nmhc")
            if key in summary_value_columns
        ]
        rows: list[int] = []
        max_column = max([47, *required_summary_columns], default=47)

        for row_index, row_values in enumerate(sheet.iter_rows(min_row=2, max_col=max_column, values_only=True), start=2):
            if not row_values[26] or not row_values[46]:
                continue
            if any(row_values[column - 1] in (None, "") for column in required_summary_columns):
                rows.append(row_index)

        return rows
    finally:
        workbook.close()


def append_imported_results(workbook_path: Path, imported_file_path: Path) -> ImportedSpreadsheetApplyResult:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")
    if not imported_file_path.exists():
        raise FileNotFoundError(f"Imported spreadsheet not found: {imported_file_path}")
    _assert_workbook_writable(workbook_path)

    rows = read_imported_spreadsheet(imported_file_path)
    hyperlinks = read_imported_hyperlinks(imported_file_path)
    header_row_index = _find_header_row(rows)
    data_rows = rows[header_row_index + 1 :]
    data_rows = [row for row in data_rows if any(value not in (None, "") for value in row[:51])]

    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook["QueryResults"]
        removed_orphan_rows = _delete_queryresults_orphan_link_rows(sheet)
        last_existing_test_id = _get_last_queryresults_test_id(sheet)
        existing_row_map, existing_row_values = _collect_queryresults_existing_rows(sheet)
        rows_to_append: list[list[object]] = []
        rows_to_append_source_indexes: list[int] = []
        appended_test_ids: list[str] = []
        updated_rows: list[int] = []
        updated_test_ids: list[str] = []
        seen_new_test_ids: set[str] = set()
        seen_updated_test_ids: set[str] = set()

        if last_existing_test_id:
            _append_log(f"Last QueryResults TESTID: {last_existing_test_id}")
        _append_log(f"Existing QueryResults TESTID count: {len(existing_row_map)}")
        if removed_orphan_rows:
            _append_log(f"Removed {removed_orphan_rows} orphan QueryResults link row(s) before import.")

        for source_row_index, row_values in enumerate(data_rows):
            test_id = _normalize_test_id(row_values[0] if row_values else None)
            if test_id is None:
                continue
            normalized = _normalize_row_length(row_values, 51)

            existing_row_index = existing_row_map.get(test_id)
            if existing_row_index is None:
                if test_id in seen_new_test_ids:
                    continue
                rows_to_append.append(normalized)
                rows_to_append_source_indexes.append(source_row_index)
                appended_test_ids.append(test_id)
                seen_new_test_ids.add(test_id)
                continue

            if not _import_row_has_changes(normalized, existing_row_values[test_id]) or test_id in seen_updated_test_ids:
                continue

            for column_index, value in enumerate(normalized, start=1):
                if column_index == 51 and value in (None, ""):
                    continue
                cell = sheet.cell(existing_row_index, column_index)
                cell.value = value

                hyperlink_target = hyperlinks.get((header_row_index + 1 + source_row_index, column_index - 1))
                if hyperlink_target and LINKED_DOCUMENT_START_COLUMN <= column_index <= LINKED_DOCUMENT_END_COLUMN:
                    cell.hyperlink = hyperlink_target
                    cell.style = "Hyperlink"
                elif LINKED_DOCUMENT_START_COLUMN <= column_index <= LINKED_DOCUMENT_END_COLUMN:
                    cell.hyperlink = None

            updated_rows.append(existing_row_index)
            updated_test_ids.append(test_id)
            seen_updated_test_ids.add(test_id)
            print(f"Updated imported row for existing TESTID {test_id} at QueryResults row {existing_row_index}")

        if not rows_to_append and not updated_rows:
            if removed_orphan_rows:
                _save_workbook_or_raise(workbook, workbook_path)
                print(f"Workbook saved after cleaning orphan QueryResults rows: {workbook_path}")
            if last_existing_test_id:
                detail = f"No new or changed imported rows were found. Last QueryResults row TESTID is {last_existing_test_id}."
            else:
                detail = "No imported rows found to append or update."
            print(detail)
            _append_log(detail)
            _update_status("Spreadsheet Imported", detail, "completed")
            return ImportedSpreadsheetApplyResult([], [], [], [])

        next_row = _get_last_nonempty_row(sheet, 1) + 1
        appended_rows: list[int] = []

        for source_row_index, row_values in zip(rows_to_append_source_indexes, rows_to_append):
            normalized = _normalize_row_length(row_values, 51)
            for column_index, value in enumerate(normalized, start=1):
                cell = sheet.cell(next_row, column_index)
                cell.value = value

                hyperlink_target = hyperlinks.get((header_row_index + 1 + source_row_index, column_index - 1))
                if hyperlink_target and LINKED_DOCUMENT_START_COLUMN <= column_index <= LINKED_DOCUMENT_END_COLUMN:
                    cell.hyperlink = hyperlink_target
                    cell.style = "Hyperlink"

            print(f"Appended imported row to QueryResults row {next_row}")
            appended_rows.append(next_row)
            next_row += 1

        _normalize_queryresults_font_size(sheet)
        _save_workbook_or_raise(workbook, workbook_path)
        print(f"Workbook saved after importing spreadsheet: {workbook_path}")
        print(f"Imported rows: {len(appended_rows)}")
        print(f"Updated existing rows: {len(updated_rows)}")
        if appended_test_ids:
            _append_log(
                f"Appended missing TESTIDs from imported spreadsheet: {appended_test_ids[0]} to {appended_test_ids[-1]} "
                f"({len(appended_test_ids)} row(s))."
            )
            _append_log(f"Last appended TESTID after import: {appended_test_ids[-1]}")
        if updated_test_ids:
            _append_log(
                f"Updated existing TESTIDs from imported spreadsheet: {updated_test_ids[0]} to {updated_test_ids[-1]} "
                f"({len(updated_test_ids)} row(s))."
            )
        if last_existing_test_id:
            _update_status(
                "Spreadsheet Imported",
                f"Imported {len(appended_rows)} new row(s) and updated {len(updated_rows)} existing row(s) in QueryResults; last existing row TESTID was {last_existing_test_id}.",
            )
        else:
            _update_status("Spreadsheet Imported", f"Imported {len(appended_rows)} new row(s) and updated {len(updated_rows)} existing row(s) in QueryResults.")
        return ImportedSpreadsheetApplyResult(appended_rows, updated_rows, appended_test_ids, updated_test_ids)
    finally:
        workbook.close()


def _get_cookie_header(driver: WebDriver) -> str:
    return "; ".join(f"{cookie['name']}={cookie['value']}" for cookie in driver.get_cookies())


def _download_file_with_session(driver: WebDriver, file_url: str, target_path: Path) -> None:
    request = Request(file_url, headers={"Cookie": _get_cookie_header(driver), "User-Agent": "Mozilla/5.0"})
    with urlopen(request) as response:
        target_path.write_bytes(response.read())


def _download_file_with_cookie_header(cookie_header: str, file_url: str, target_path: Path) -> None:
    request = Request(file_url, headers={"Cookie": cookie_header, "User-Agent": "Mozilla/5.0"})
    with urlopen(request) as response:
        target_path.write_bytes(response.read())


def _has_expected_pdf_page_count(pdf_path: Path, minimum_pages: int = 14) -> bool:
    try:
        reader = PdfReader(str(pdf_path))
    except Exception:
        return False
    return len(reader.pages) >= minimum_pages


def _collect_pdf_links_from_current_page(driver: WebDriver) -> dict[str, str]:
    elements = driver.find_elements(By.XPATH, "//a[contains(translate(@href, 'PDF', 'pdf'), '.pdf') or contains(translate(normalize-space(.), 'PDF', 'pdf'), '.pdf')]")
    links: dict[str, str] = {}
    for element in elements:
        href = element.get_attribute("href")
        text = (element.text or "").strip()
        candidate = href or text
        if not candidate:
            continue
        name = Path(candidate.split("?")[0]).name
        if name.lower().endswith(".pdf") and href:
            links[name] = href
    return links


def _collect_named_page_urls(driver: WebDriver, timeout_seconds: int, link_text: str) -> list[str]:
    collected_pages: list[str] = []

    def collect_from_current_context() -> None:
        links = driver.find_elements(
            By.XPATH,
            f"//a[contains(normalize-space(.), '{link_text}')] | //button[contains(normalize-space(.), '{link_text}')]",
        )
        for link in links:
            href = (link.get_attribute("href") or "").strip()
            if href and not href.lower().startswith("javascript:"):
                collected_pages.append(urljoin(driver.current_url, href))
                continue

            onclick = (link.get_attribute("onclick") or "").strip()
            if not onclick:
                continue
            match = re.search(r"['\"]([^'\"]+)['\"]", onclick)
            if not match:
                continue
            collected_pages.append(urljoin(driver.current_url, match.group(1)))

    for frame_name in ("contents", "main"):
        try:
            switch_to_frame(driver, frame_name, timeout_seconds)
        except TimeoutException:
            driver.switch_to.default_content()
            continue
        collect_from_current_context()
        driver.switch_to.default_content()

    collect_from_current_context()
    driver.switch_to.default_content()
    return list(dict.fromkeys(collected_pages))


def _open_downloadable_files_pages(driver: WebDriver, timeout_seconds: int) -> list[str]:
    return _collect_named_page_urls(driver, timeout_seconds, "Downloadable Files")


def _open_ati_files_pages(driver: WebDriver, timeout_seconds: int) -> list[str]:
    return _collect_named_page_urls(driver, timeout_seconds, "ATI Files in GTDM")


def _build_ati_detail_page_urls(
    driver: WebDriver,
    test_ids: list[str],
    vehicle_vin: str,
) -> list[str]:
    base_url = urljoin(driver.current_url, "bldgtdm.cgi")
    return [
        f"{base_url}?testid={test_id}&vehid={vehicle_vin}&ati=y"
        for test_id in test_ids
        if test_id and vehicle_vin
    ]


def _collect_mf4_links_from_current_page(driver: WebDriver) -> dict[str, str]:
    links: dict[str, str] = {}

    def extract_file_names(value: str) -> list[str]:
        return re.findall(r"(?i)([^/?&=\"'\s]+\.mf4)\b", unquote(value or ""))

    def add_link(raw_url: str, text: str) -> None:
        if not raw_url or raw_url.lower().startswith("javascript:"):
            return
        file_names = extract_file_names(raw_url) + extract_file_names(text)
        if not file_names:
            return
        resolved_url = urljoin(driver.current_url, raw_url)
        for file_name in file_names:
            links[file_name] = resolved_url

    def collect_from_current_context(depth: int = 0) -> None:
        elements = driver.find_elements(By.CSS_SELECTOR, "a[href], [data-href], [data-url], [onclick]")
        for element in elements:
            text = (element.text or "").strip()
            for attribute_name in ("href", "data-href", "data-url"):
                add_link((element.get_attribute(attribute_name) or "").strip(), text)
            onclick = (element.get_attribute("onclick") or "").strip()
            for raw_url in re.findall(r"['\"]([^'\"]+)['\"]", onclick):
                add_link(raw_url, text)

        if depth >= 6:
            return
        frames = driver.find_elements(By.CSS_SELECTOR, "iframe, frame")
        for frame in frames:
            switched_to_frame = False
            try:
                driver.switch_to.frame(frame)
                switched_to_frame = True
                collect_from_current_context(depth + 1)
            except Exception:
                pass
            finally:
                if switched_to_frame:
                    driver.switch_to.parent_frame()

    try:
        collect_from_current_context()
    finally:
        driver.switch_to.default_content()
    return links


class _Mf4AnchorParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: list[tuple[str, str]] = []
        self._href = ""
        self._text: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "a":
            return
        attributes = dict(attrs)
        self._href = attributes.get("href") or ""
        self._text = []

    def handle_data(self, data: str) -> None:
        if self._href:
            self._text.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "a" and self._href:
            self.links.append((self._href, " ".join(self._text).strip()))
            self._href = ""
            self._text = []


def _collect_mf4_links_from_authenticated_page(
    driver: WebDriver,
    page_url: str,
) -> dict[str, str]:
    return _collect_mf4_links_with_cookie_header(page_url, _get_cookie_header(driver))


def _collect_mf4_links_with_cookie_header(
    page_url: str,
    cookie_header: str,
) -> dict[str, str]:
    request = Request(
        page_url,
        headers={"Cookie": cookie_header, "User-Agent": "Mozilla/5.0"},
    )
    with urlopen(request, timeout=30) as response:
        page_html = response.read().decode(response.headers.get_content_charset() or "utf-8", errors="replace")

    parser = _Mf4AnchorParser()
    parser.feed(page_html)
    links: dict[str, str] = {}
    for raw_url, text in parser.links:
        values = (unquote(raw_url), unquote(text))
        file_names = re.findall(r"(?i)([^/?&=\"'\s]+\.mf4)\b", " ".join(values))
        if not file_names:
            continue
        resolved_url = urljoin(page_url, raw_url)
        for file_name in file_names:
            links[file_name] = resolved_url
    return links


def _matches_test_id_filename(file_name: str, test_id: str) -> bool:
    return re.search(rf"(?<!\d){re.escape(test_id)}(?!\d)", file_name) is not None


def _get_ati_file_suffix(file_name: str) -> str | None:
    match = re.search(r"_(\d{3})\.mf4$", Path(file_name).name, re.IGNORECASE)
    return match.group(1) if match else None


def _is_requested_ati_filename(file_name: str) -> bool:
    return _get_ati_file_suffix(file_name) is not None


def _get_existing_ati_test_ids(ati_root_dir: Path) -> set[str]:
    return {
        match.group(1)
        for path in ati_root_dir.rglob("*.mf4")
        if _is_requested_ati_filename(path.name)
        and (match := re.match(r"^(\d{7})", path.stem))
    }


def _get_known_ati_test_ids(ati_root_dir: Path) -> set[str]:
    archive_test_ids = _get_existing_ati_test_ids(_get_ati_archive_directory())
    target_test_ids = _get_existing_ati_test_ids(ati_root_dir)
    return archive_test_ids | target_test_ids


def _collect_imported_test_ids(imported_file_path: Path) -> list[str]:
    rows = read_imported_spreadsheet(imported_file_path)
    header_row_index = _find_header_row(rows)
    data_rows = rows[header_row_index + 1 :]
    ordered_test_ids: list[str] = []
    seen_test_ids: set[str] = set()

    for row in data_rows:
        if not any(value not in (None, "") for value in row[:51]):
            continue
        test_id = _normalize_test_id(row[0] if row else None)
        if test_id is None or test_id in seen_test_ids:
            continue
        ordered_test_ids.append(test_id)
        seen_test_ids.add(test_id)

    return ordered_test_ids


def _get_summary_monitor_by_test_id(workbook_path: Path | None) -> dict[str, str]:
    if workbook_path is None or not workbook_path.exists():
        return {}

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        if "QueryResults" not in workbook.sheetnames:
            return {}
        sheet = workbook["QueryResults"]
        monitor_column = _get_summary_value_columns(sheet).get("monitor")
        if monitor_column is None:
            return {}
        return {
            test_id: str(sheet.cell(row_index, monitor_column).value).strip()
            for row_index in range(2, sheet.max_row + 1)
            if (test_id := _normalize_test_id(sheet.cell(row_index, 1).value))
            and sheet.cell(row_index, monitor_column).value not in (None, "")
        }
    finally:
        workbook.close()


def _get_summary_test_ids(workbook_path: Path | None) -> list[str]:
    if workbook_path is None or not workbook_path.exists():
        return []

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        if "QueryResults" not in workbook.sheetnames:
            return []
        sheet = workbook["QueryResults"]
        return list(
            dict.fromkeys(
                test_id
                for row_index in range(2, sheet.max_row + 1)
                if (test_id := _normalize_test_id(sheet.cell(row_index, 1).value))
            )
        )
    finally:
        workbook.close()


def _get_missing_ati_test_ids(test_ids: list[str], ati_root_dir: Path) -> list[str]:
    existing_test_ids = _get_existing_ati_test_ids(ati_root_dir)
    return [
        test_id
        for test_id in test_ids
        if test_id not in existing_test_ids
    ]


def _get_ati_monitor_by_test_id(
    imported_file_path: Path,
    schedule_path: Path | None,
    summary_workbook_path: Path | None,
) -> dict[str, str]:
    monitor_by_test_id = _get_summary_monitor_by_test_id(summary_workbook_path)
    schedule_entries: dict[str, str] = {}
    if schedule_path is not None:
        try:
            schedule_entries = _load_monitor_schedule_entries(schedule_path)
        except Exception as exc:
            _append_log(f"Could not load monitor names for ATI MF4 filenames: {exc.__class__.__name__}: {exc}")

    rows = read_imported_spreadsheet(imported_file_path)
    header_row_index = _find_header_row(rows)
    header_map = {
        _normalize_header_text(value): index
        for index, value in enumerate(rows[header_row_index])
        if value not in (None, "")
    }
    test_id_column = header_map.get("testid")
    tr_number_column = header_map.get("trnumber")
    if test_id_column is None or tr_number_column is None:
        return monitor_by_test_id

    for row_values in rows[header_row_index + 1 :]:
        test_id = _normalize_test_id(row_values[test_id_column] if test_id_column < len(row_values) else None)
        tr_number = _normalize_req_no(row_values[tr_number_column] if tr_number_column < len(row_values) else None)
        if test_id and tr_number and schedule_entries.get(tr_number):
            monitor_by_test_id[test_id] = schedule_entries[tr_number]
    return monitor_by_test_id


def _get_ati_monitor_keywords(monitor_text: str) -> list[str]:
    normalized = str(monitor_text or "").upper()
    keywords: list[str] = []
    if re.search(r"\bDEMO\b", normalized):
        keywords.append("EM")
    elif re.search(r"\bPREP\b", normalized):
        keywords.append("Prep")

    patterns = (
        (r"\bAFIMN\s*[_ -]?\s*([RL])\s*(\d+)", lambda match: f"AFIMN_{match.group(1)}{match.group(2)}"),
        (r"\bFAOS\s*[_ -]?\s*R\s*(\d+)", lambda match: f"FASO_R{match.group(1)}"),
        (r"\bFAOS\s*[_ -]?\s*L\s*(\d+)", lambda match: f"FASOL_L{match.group(1)}"),
        (r"\bFUELMN\s*[_ -]?\s*([RL])\s*(\d+)", lambda match: f"FUELMN_{match.group(1)}{match.group(2)}"),
        (r"\bEGOMN\s*[_ -]?\s*P\s*(\d+)\s*(\d+)\s*MS", lambda match: f"EGOMN_P{match.group(1)}_{match.group(2)}ms"),
        (r"\bVCTMN\s*(\d+)\s*DEG\s*INT\b", lambda match: f"VCTMN_INT_{match.group(1)}deg"),
        (r"\bVCTMN\s*EXH\s*(\d+)\s*S\b", lambda match: f"VCT_MN_EXH_{match.group(1)}s"),
        (r"(?:^|[^A-Z0-9])BASELINE(?:$|[^A-Z0-9])", lambda _match: "Baseline"),
    )
    for pattern, formatter in patterns:
        match = re.search(pattern, normalized)
        if match:
            keywords.append(formatter(match))
            break
    return keywords


def _build_ati_mf4_file_name(test_id: str, source_name: str, monitor_text: str) -> str:
    _ = test_id, monitor_text
    return Path(source_name).name


def _rename_ati_mf4_files(ati_root_dir: Path, monitor_by_test_id: dict[str, str]) -> int:
    renamed_count = 0
    for source_path in sorted(ati_root_dir.rglob("*.mf4")):
        test_id = next((candidate for candidate in monitor_by_test_id if _matches_test_id_filename(source_path.name, candidate)), None)
        if test_id is None:
            continue
        target_path = ati_root_dir / _build_ati_mf4_file_name(test_id, source_path.name, monitor_by_test_id.get(test_id, ""))
        if target_path == source_path:
            continue
        try:
            if target_path.exists():
                source_path.unlink()
                renamed_count += 1
                continue
            source_path.rename(target_path)
            renamed_count += 1
        except PermissionError as exc:
            _append_log(f"Skipped locked ATI MF4 filename update: {source_path.name} ({exc})")
    return renamed_count


def _normalize_ati_combined_filenames(ati_root_dir: Path) -> int:
    renamed_count = 0
    for source_path in sorted(ati_root_dir.rglob("*.mf4")):
        if not _is_requested_ati_filename(source_path.name):
            continue
        target_path = source_path
        if target_path == source_path:
            continue
        try:
            if target_path.exists():
                source_path.unlink()
            else:
                source_path.rename(target_path)
            renamed_count += 1
        except PermissionError as exc:
            _append_log(f"Skipped locked ATI MF4 combined filename update: {source_path.name} ({exc})")
    return renamed_count


def _restore_cached_combined_ati_mf4s(
    test_ids: list[str],
    ati_root_dir: Path,
    monitor_by_test_id: dict[str, str],
) -> tuple[list[Path], int]:
    cached_root = _get_ati_archive_directory()
    restored: list[Path] = []
    skipped_existing = 0

    for source_path in sorted(cached_root.rglob("*.mf4")):
        if source_path.parent == ati_root_dir:
            continue
        if not _is_requested_ati_filename(source_path.name):
            continue
        if not any(_matches_test_id_filename(source_path.name, test_id) for test_id in test_ids):
            continue

        test_id = next(test_id for test_id in test_ids if _matches_test_id_filename(source_path.name, test_id))
        target_path = ati_root_dir / _build_ati_mf4_file_name(
            test_id,
            source_path.name,
            monitor_by_test_id.get(test_id, ""),
        )
        if target_path.exists():
            skipped_existing += 1
            continue
        shutil.copy2(source_path, target_path)
        restored.append(target_path)

    return restored, skipped_existing


def _download_ati_files_worker(
    cookie_header: str,
    download_tasks: list[tuple[str, str, str, Path]],
    cached_files: list[Path],
    ati_root_dir: Path,
    selected_count: int,
    skipped_existing: int,
) -> list[Path]:
    downloaded = list(cached_files)
    matched_names = [file_name for _, file_name, _, _ in download_tasks]
    with ThreadPoolExecutor(max_workers=min(6, max(1, len(download_tasks)))) as executor:
        future_map = {
            executor.submit(_download_file_with_cookie_header, cookie_header, file_url, target_path): (test_id, file_name, target_path)
            for test_id, file_name, file_url, target_path in download_tasks
        }
        for future in as_completed(future_map):
            test_id, file_name, target_path = future_map[future]
            try:
                future.result()
            except Exception as exc:
                _append_log(f"ATI MF4 download failed for {test_id} ({file_name}): {exc.__class__.__name__}: {exc}")
                continue
            downloaded.append(target_path)
            _update_status(
                "ATI Downloaded",
                f"Downloaded {len(downloaded) - len(cached_files)} of {len(download_tasks)} ATI MF4 file(s); latest file: {file_name}.",
            )

    if matched_names:
        preview = ", ".join(matched_names[:10])
        if len(matched_names) > 10:
            preview += ", ..."
        _append_log(
            f"ATI Files in GTDM matched {len(matched_names)} numbered MF4 file(s) across {selected_count} TESTID(s): {preview}"
        )
    else:
        _append_log("ATI Files in GTDM contained no numbered MF4 file names matching the current TESTIDs")

    _update_status(
        "ATI Download Complete",
        f"Prepared {len(downloaded)} numbered MF4 file(s) in {ati_root_dir}: restored {len(cached_files)} from cache and downloaded {len(downloaded) - len(cached_files)} from ATI; skipped {skipped_existing} existing file(s).",
        "completed",
    )
    return downloaded


def _run_ati_discovery_and_download(
    cookie_header: str,
    candidate_pages: list[str],
    selected_test_ids: list[str],
    missing_test_ids: list[str],
    ati_root_dir: Path,
    archive_dir: Path,
    monitor_by_test_id: dict[str, str],
    cached_files: list[Path],
    skipped_existing: int,
) -> list[Path]:
    downloaded = list(cached_files)
    matched_names: list[str] = []
    missing_set = set(selected_test_ids)
    scheduled_file_keys: set[tuple[str, str | None]] = set()
    page_link_count = 0
    completed_pages = 0
    download_executor = ThreadPoolExecutor(max_workers=min(6, max(1, len(missing_test_ids))))
    download_futures: dict[Future[None], tuple[str, str, Path, Path]] = {}

    try:
        with ThreadPoolExecutor(max_workers=min(6, max(1, len(candidate_pages)))) as page_executor:
            page_futures = {
                page_executor.submit(_collect_mf4_links_with_cookie_header, page_url, cookie_header): page_url
                for page_url in candidate_pages
            }
            for page_future in as_completed(page_futures):
                completed_pages += 1
                try:
                    page_links = page_future.result()
                except Exception as exc:
                    _append_log(
                        f"Could not read ATI detail page in parallel: {page_futures[page_future]} "
                        f"({exc.__class__.__name__}: {exc})"
                    )
                    page_links = {}

                page_link_count += len(page_links)
                _update_status(
                    "Reading ATI Detail Pages",
                    f"Reading ATI detail page {completed_pages}/{len(candidate_pages)}; found {page_link_count} MF4 link(s) so far.",
                )

                for file_name, file_url in page_links.items():
                    if not _is_requested_ati_filename(file_name):
                        continue
                    test_id = next(
                        (candidate for candidate in missing_set if _matches_test_id_filename(file_name, candidate)),
                        None,
                    )
                    file_key = (test_id, _get_ati_file_suffix(file_name)) if test_id else None
                    if test_id is None or file_key in scheduled_file_keys:
                        continue

                    target_path = ati_root_dir / _build_ati_mf4_file_name(
                        test_id,
                        file_name,
                        monitor_by_test_id.get(test_id, ""),
                    )
                    if target_path.exists():
                        skipped_existing += 1
                        scheduled_file_keys.add(file_key)
                        continue

                    scheduled_file_keys.add(file_key)
                    matched_names.append(file_name)
                    archive_path = archive_dir / file_name
                    download_futures[download_executor.submit(
                        _download_file_with_cookie_header,
                        cookie_header,
                        file_url,
                        archive_path,
                    )] = (test_id, file_name, archive_path, target_path)
                    _update_status(
                        "ATI Download Queued",
                        f"Found {file_name}; queued TESTID {test_id} for download while detail pages continue. Downloads queued: {len(download_futures)}.",
                    )

        for download_future in as_completed(download_futures):
            test_id, file_name, archive_path, target_path = download_futures[download_future]
            try:
                download_future.result()
                shutil.copy2(archive_path, target_path)
            except Exception as exc:
                _append_log(
                    f"ATI MF4 download failed for {test_id} ({file_name}): "
                    f"{exc.__class__.__name__}: {exc}"
                )
                continue
            downloaded.append(target_path)
            _update_status(
                "ATI Downloaded",
                f"Downloaded and copied {len(downloaded) - len(cached_files)} of {len(download_futures)} ATI MF4 file(s); latest file: {file_name}.",
            )
    finally:
        download_executor.shutdown(wait=True)

    if matched_names:
        preview = ", ".join(matched_names[:10])
        if len(matched_names) > 10:
            preview += ", ..."
        _append_log(
            f"ATI Files in GTDM matched {len(matched_names)} numbered MF4 file(s) across {len(selected_test_ids)} TESTID(s): {preview}"
        )
    else:
        _append_log("ATI Files in GTDM contained no numbered MF4 file names matching the current TESTIDs")

    _update_status(
        "ATI Download Complete",
        f"Prepared {len(downloaded)} numbered MF4 file(s) in {ati_root_dir}: restored {len(cached_files)} from cache and downloaded {len(downloaded) - len(cached_files)} from ATI; skipped {skipped_existing} existing file(s).",
        "completed",
    )
    return downloaded


def download_combined_ati_mf4s(
    driver: WebDriver,
    test_ids: list[str],
    ati_root_dir: Path,
    timeout_seconds: int,
    monitor_by_test_id: dict[str, str] | None = None,
    vehicle_vin: str = "",
) -> list[Path] | AtiDownloadJob:
    selected_test_ids = [test_id for test_id in test_ids if test_id]
    monitor_by_test_id = monitor_by_test_id or {}
    ati_root_dir.mkdir(parents=True, exist_ok=True)
    archive_dir = _get_ati_archive_directory()
    archive_dir.mkdir(parents=True, exist_ok=True)
    if not selected_test_ids:
        _update_status("ATI Download Skipped", "No TESTIDs were available for ATI GTDM downloads.", "completed")
        return []

    _update_status(
        "ATI Download Started",
        f"Checking Imported_spreadsheet\\Archives for {len(selected_test_ids)} TESTID(s) and downloading numbered MF4 files.",
    )
    cached_files, skipped_existing = _restore_cached_combined_ati_mf4s(
        selected_test_ids,
        ati_root_dir,
        monitor_by_test_id,
    )
    if cached_files:
        _update_status(
            "ATI Cache Restored",
            f"Restored {len(cached_files)} matching numbered MF4 file(s) from the {Path(DEFAULT_DOWNLOAD_DIR)} cache into {ati_root_dir}.",
        )
    known_test_ids = _get_known_ati_test_ids(ati_root_dir)
    missing_test_ids = [test_id for test_id in selected_test_ids if test_id not in known_test_ids]
    existing_test_count = len(selected_test_ids) - len(missing_test_ids)
    _append_log(
        f"ATI local check: {existing_test_count} TESTID(s) found in Imported_spreadsheet\\Archives or {ati_root_dir}; "
        f"{len(missing_test_ids)} TESTID(s) are absent and need download."
    )
    _update_status(
        "ATI Download Started",
        f"Archives check complete: {existing_test_count} TESTID(s) already present, {len(missing_test_ids)} need download.",
    )

    candidate_pages = _build_ati_detail_page_urls(driver, selected_test_ids, vehicle_vin)
    candidate_pages = list(dict.fromkeys(page_url for page_url in candidate_pages if page_url))

    if not candidate_pages:
        _append_log("No ATI Files in GTDM page links were found on the current APTL page.")
        preview = ", ".join(missing_test_ids[:12]) or "None"
        _update_status(
            "ATI Download Skipped",
            f"ATI Files in GTDM link was unavailable. {len(missing_test_ids)} TESTID(s) still need MF4 download: {preview}.",
            "completed",
        )
        return cached_files

    executor = ThreadPoolExecutor(max_workers=1)
    future = executor.submit(
        _run_ati_discovery_and_download,
        _get_cookie_header(driver),
        candidate_pages,
        selected_test_ids,
        missing_test_ids,
        ati_root_dir,
        archive_dir,
        monitor_by_test_id,
        cached_files,
        skipped_existing,
    )
    _append_log(
        f"Started parallel ATI detail-page checking for {len(candidate_pages)} page(s); "
        "MF4 downloads will start immediately as links are found."
    )
    return AtiDownloadJob(future, executor)


def download_query_result_pdfs(
    driver: WebDriver,
    expected_pdf_names: set[str],
    pdf_dir: Path,
    timeout_seconds: int,
) -> list[Path]:
    pdf_dir.mkdir(parents=True, exist_ok=True)
    downloaded: dict[str, Path] = {}
    _update_status("PDF Download Started", f"Looking for {len(expected_pdf_names)} expected PDF file(s).")

    candidate_pages = [driver.current_url]
    candidate_pages.extend(_open_downloadable_files_pages(driver, timeout_seconds))

    for page_url in candidate_pages:
        page_links = _collect_links_from_page_in_temp_window(driver, page_url, _collect_pdf_links_from_current_page)

        for pdf_name in expected_pdf_names:
            href = page_links.get(pdf_name)
            if not href or pdf_name in downloaded:
                continue

            target_path = pdf_dir / pdf_name
            _download_file_with_session(driver, href, target_path)

            if not _has_expected_pdf_page_count(target_path):
                target_path.unlink(missing_ok=True)
                print(f"Skipped PDF with unexpected page count: {pdf_name}")
                continue

            downloaded[pdf_name] = target_path
            print(f"Downloaded PDF: {target_path}")
            _update_status(
                "PDF Downloaded",
                f"Downloaded {len(downloaded)} of {len(expected_pdf_names)} expected PDF file(s).",
            )
    _update_status("PDF Download Finished", f"Downloaded {len(downloaded)} PDF file(s) into {pdf_dir}.")
    return list(downloaded.values())


def _get_pdf_download_targets_for_rows(workbook_path: Path, rows: list[int]) -> dict[str, str]:
    workbook = load_workbook(workbook_path, data_only=False)
    try:
        sheet = workbook["QueryResults"]
        download_targets: dict[str, str] = {}

        for row in rows:
            pdf_cell = sheet[f"AA{row}"]
            csv_cell = sheet[f"AU{row}"]
            pdf_name = pdf_cell.value
            if not pdf_name or not csv_cell.value or pdf_cell.hyperlink is None:
                continue

            hyperlink_target = pdf_cell.hyperlink.target or pdf_cell.hyperlink.location
            if not hyperlink_target:
                continue

            download_targets[str(pdf_name)] = hyperlink_target

        return download_targets
    finally:
        workbook.close()


def _restore_archived_pdfs(pdf_names: set[str], pdf_dir: Path) -> int:
    archive_dir = _get_archive_directory(pdf_dir)
    restored = 0
    for pdf_name in pdf_names:
        target_path = pdf_dir / pdf_name
        if target_path.exists():
            continue
        archived_matches = list(archive_dir.rglob(pdf_name))
        for archived_path in reversed(archived_matches):
            if not archived_path.is_file() or not _has_expected_pdf_page_count(archived_path):
                continue
            try:
                shutil.copy2(archived_path, target_path)
            except OSError as exc:
                _append_log(f"Could not restore archived PDF {pdf_name}: {exc.__class__.__name__}: {exc}")
                continue
            restored += 1
            break
    return restored


def download_missing_linked_pdfs(
    driver: WebDriver,
    workbook_path: Path,
    rows: list[int],
    pdf_dir: Path,
) -> list[Path]:
    pdf_dir.mkdir(parents=True, exist_ok=True)
    download_targets = _get_pdf_download_targets_for_rows(workbook_path, rows)

    if not download_targets:
        _update_status("PDF Download Skipped", "No AA PDF links matched rows with AU CSV values.", "completed")
        return []

    restored_archived = _restore_archived_pdfs(set(download_targets), pdf_dir)
    if restored_archived:
        _append_log(
            f"Restored {restored_archived} valid PDF file(s) from {pdf_dir.name}\\Archives before download check."
        )

    downloaded: list[Path] = []
    skipped_existing = 0
    skipped_invalid = 0
    _update_status(
        "PDF Download Started",
        f"Checking {len(download_targets)} linked PDF file(s) for rows with AU CSV values; only PDFs with 14 or more pages will be kept.",
    )

    for pdf_name, href in download_targets.items():
        target_path = pdf_dir / pdf_name
        if target_path.exists():
            if _has_expected_pdf_page_count(target_path):
                skipped_existing += 1
                print(f"Skipped existing valid PDF: {target_path}")
                continue

            target_path.unlink(missing_ok=True)
            skipped_invalid += 1
            _append_log(f"Removed existing PDF with unexpected page count: {target_path}")

        _download_file_with_session(driver, href, target_path)
        if not _has_expected_pdf_page_count(target_path):
            target_path.unlink(missing_ok=True)
            skipped_invalid += 1
            print(f"Skipped PDF with fewer than 14 pages: {pdf_name}")
            _append_log(f"Skipped PDF with fewer than 14 pages: {pdf_name}")
            continue

        downloaded.append(target_path)
        print(f"Downloaded linked PDF: {target_path}")
        _update_status(
            "PDF Downloaded",
            f"Downloaded {len(downloaded)} valid PDF(s) with 14 or more pages of {len(download_targets)} candidate linked PDF file(s).",
        )

    _update_status(
        "PDF Download Finished",
        f"Downloaded {len(downloaded)} valid PDF file(s) with 14 or more pages into {pdf_dir}; skipped {skipped_existing} existing valid file(s); rejected {skipped_invalid} file(s) with fewer than 14 pages.",
        "completed",
    )
    return downloaded


def enrich_rows_from_local_pdfs(workbook_path: Path, rows: list[int], pdf_dir: Path, rebuild_plots: bool = True) -> int:
    if not rows:
        return 0
    _assert_workbook_writable(workbook_path)

    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook["QueryResults"]
        standards = get_eu7_standards(workbook_path)
        enriched = 0
        skipped: list[str] = []

        for row in rows:
            try:
                if apply_pdf_summary_to_row(sheet, row, pdf_dir, standards):
                    enriched += 1
                    print(f"PDF-enriched QueryResults row {row}")
                else:
                    pdf_name = sheet[f"AA{row}"].value or "<missing pdf name>"
                    pdf_path = pdf_dir / str(pdf_name) if pdf_name != "<missing pdf name>" else None
                    if pdf_path is not None and pdf_path.exists() and not _has_expected_pdf_page_count(pdf_path):
                        _append_log(f"Skipped AY:BS enrichment for QueryResults row {row} ({pdf_name}): PDF has fewer than 14 pages")
            except Exception as exc:
                pdf_name = sheet[f"AA{row}"].value or "<missing pdf name>"
                skipped.append(str(pdf_name))
                _append_log(
                    f"Skipped PDF enrichment for QueryResults row {row} ({pdf_name}): "
                    f"{exc.__class__.__name__}: {exc}"
                )
                print(f"Skipped PDF enrichment for QueryResults row {row} ({pdf_name}): {exc}")

        if rebuild_plots:
            _update_status("Building Emis_Plot_Base and Emis_Plot_OBD", "Creating the Emis_Plot_Base and Emis_Plot_OBD worksheets.")
            _build_emis_plot_sheets(workbook, standards)
        _normalize_queryresults_font_size(sheet)
        _save_workbook_or_raise(workbook, workbook_path)
        print(f"Workbook saved after PDF enrichment: {workbook_path}")
        detail = f"Updated {enriched} row(s) from local PDF summaries."
        if skipped:
            detail += f" Skipped {len(skipped)} PDF(s): {', '.join(skipped[:5])}."
        _update_status("Workbook Enriched", detail)
        return enriched
    finally:
        workbook.close()


def rebuild_workbook_plots(workbook_path: Path) -> None:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")
    _assert_workbook_writable(workbook_path)

    standards = get_eu7_standards(workbook_path)
    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook["QueryResults"]
        _update_status("Building Emis_Plot_Base and Emis_Plot_OBD", "Refreshing the Emis_Plot_Base and Emis_Plot_OBD worksheets from the current workbook data.")
        _build_emis_plot_sheets(workbook, standards)
        _normalize_queryresults_font_size(sheet)
        _save_workbook_or_raise(workbook, workbook_path)
        print(f"Workbook plots rebuilt: {workbook_path}")
    finally:
        workbook.close()


def append_workbook_plots(workbook_path: Path, query_rows: list[int], update_status: bool = True) -> None:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")
    _assert_workbook_writable(workbook_path)

    standards = get_eu7_standards(workbook_path)
    workbook = load_workbook(workbook_path)
    try:
        before_snapshots = {
            DEFAULT_PLOT_BASE_SHEET: _snapshot_plot_sheet_rows(workbook[DEFAULT_PLOT_BASE_SHEET]),
            DEFAULT_PLOT_OBD_SHEET: _snapshot_plot_sheet_rows(workbook[DEFAULT_PLOT_OBD_SHEET]),
        }
        action_detail = (
            "Refreshing all plot charts from the current Base/OBD sheet layout."
            if not query_rows
            else "Appending new QueryResults rows to Emis_Plot_Base and Emis_Plot_OBD, then refreshing all plot charts."
        )
        if update_status:
            _update_status(
                "Building Emis_Plot_Base and Emis_Plot_OBD",
                action_detail,
            )
        _build_emis_plot_sheets(workbook, standards, query_rows=query_rows, clear_existing=False)
        _assert_append_preserved_plot_rows(
            before_snapshots[DEFAULT_PLOT_BASE_SHEET],
            _snapshot_plot_sheet_rows(workbook[DEFAULT_PLOT_BASE_SHEET]),
            DEFAULT_PLOT_BASE_SHEET,
        )
        _assert_append_preserved_plot_rows(
            before_snapshots[DEFAULT_PLOT_OBD_SHEET],
            _snapshot_plot_sheet_rows(workbook[DEFAULT_PLOT_OBD_SHEET]),
            DEFAULT_PLOT_OBD_SHEET,
        )
        _save_workbook_or_raise(workbook, workbook_path)
        print(f"Workbook plots appended: {workbook_path}")
    finally:
        workbook.close()


def get_pdf_names_for_rows(workbook_path: Path, rows: list[int]) -> set[str]:
    workbook = load_workbook(workbook_path, data_only=False)
    try:
        sheet = workbook["QueryResults"]
        pdf_names: set[str] = set()
        for row in rows:
            value = sheet[f"AA{row}"].value
            if value:
                pdf_names.add(str(value))
        return pdf_names
    finally:
        workbook.close()


def populate_workbook_from_pdfs(workbook_path: Path, pdf_dir: Path) -> int:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")
    if not pdf_dir.exists():
        raise FileNotFoundError(f"PDF directory not found: {pdf_dir}")
    _assert_workbook_writable(workbook_path)

    standards = get_eu7_standards(workbook_path)
    workbook = load_workbook(workbook_path)
    try:
        sheet = workbook["QueryResults"]
        header_map = _get_queryresults_header_map(sheet)

        row_lookup: dict[str, int] = {}
        for row in range(2, sheet.max_row + 1):
            test_id = sheet[f"A{row}"].value
            pdf_name = sheet[f"AA{row}"].value
            if test_id is not None:
                row_lookup[str(test_id)] = row
            if pdf_name:
                row_lookup[str(pdf_name)] = row

        updated_rows = 0
        for pdf_path in sorted(pdf_dir.glob("*.pdf")):
            try:
                summary = parse_summary_pdf(pdf_path)
            except Exception as exc:
                print(f"Skipped {pdf_path.name}: {exc}")
                continue

            row = row_lookup.get(str(summary["test_id"])) or row_lookup.get(pdf_path.name)

            if row is None:
                print(f"Skipped {pdf_path.name}: no matching QueryResults row.")
                continue

            standard_type = _get_emission_standard_type(sheet, row, header_map)
            _apply_emission_percentages(summary, standard_type, standards)

            _write_summary_to_row(sheet, row, summary)

            updated_rows += 1
            print(f"Updated row {row} from {pdf_path.name}")

        _update_status("Building Emis_Plot_Base and Emis_Plot_OBD", "Creating the Emis_Plot_Base and Emis_Plot_OBD worksheets.")
        _build_emis_plot_sheets(workbook, standards)
        _normalize_queryresults_font_size(sheet)
        _save_workbook_or_raise(workbook, workbook_path)
        print(f"Workbook saved: {workbook_path}")
        return updated_rows
    finally:
        workbook.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Open Emissions Utilities in an external browser and verify access after login."
    )
    parser.add_argument(
        "--browser",
        choices=["auto", "edge", "chrome"],
        default="auto",
        help="Browser to launch. Default tries Microsoft Edge first, then falls back to Google Chrome.",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run the automation browser without displaying a browser window.",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=300,
        help="Seconds to wait for login and successful page load.",
    )
    parser.add_argument(
        "--close-on-success",
        action="store_true",
        help="Close the browser immediately after confirming access.",
    )
    parser.add_argument(
        "--click-target",
        default=DEFAULT_CLICK_TARGET,
        help="Hierarchical menu label to click after login. Default is 1.1.1.",
    )
    parser.add_argument(
        "--run-mode",
        choices=["fresh", "append"],
        default="",
        help="Optional override for how to update the summary workbook after comparison.",
    )
    parser.add_argument(
        "--vehicle-vin",
        default="",
        help="Vehicle VIN value to enter in the Reprint Report query form.",
    )
    parser.add_argument(
        "--test-requestor",
        default="",
        help="Test Requestor value to enter in the Reprint Report query form.",
    )
    parser.add_argument(
        "--start-date",
        default="",
        help="Start Date value to enter in the Reprint Report query form.",
    )
    parser.add_argument(
        "--sort-by",
        default="",
        help="Sort By option to select in the Reprint Report query form.",
    )
    parser.add_argument(
        "--populate-workbook",
        action="store_true",
        help="Populate QueryResults columns AZ:BV from PDF summary data in Emission_Result.",
    )
    parser.add_argument(
        "--workbook-path",
        default="",
        help="Optional workbook path override. When omitted, the app copies the template into Emission_Summary_Result as YYYYMMDD_Emission_Summary.xlsx.",
    )
    parser.add_argument(
        "--pdf-dir",
        default=DEFAULT_PDF_DIR,
        help="PDF directory used with --populate-workbook.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if not _acquire_run_mutex():
        print("Emission Summary is already running. Wait for the current run to finish.", file=sys.stderr)
        return 1

    workbook_path: Path | None = None
    workbook_to_open: Path | None = None
    existing_result_workbook: Path | None = None
    monitor_schedule_path: Path | None = None
    ati_root_dir: Path | None = None
    ati_monitor_by_test_id: dict[str, str] = {}
    downloaded_ati_files: list[Path] = []
    ati_download_job: AtiDownloadJob | None = None
    renamed_ati_files = 0
    result_code = 0
    keep_automation_browser_open = False
    live_share_tab_opened = False

    _set_current_run_token()
    _load_window_bounds_into_memory()
    Path(DEFAULT_LOG_FILE).write_text("", encoding="utf-8")
    Path(DEFAULT_STATUS_TEXT_FILE).write_text("", encoding="utf-8")
    Path(DEFAULT_STATUS_HTML_FILE).write_text("", encoding="utf-8")
    Path(DEFAULT_STATUS_JSON_FILE).write_text("", encoding="utf-8")
    Path(DEFAULT_STATUS_FRAGMENT_FILE).write_text("", encoding="utf-8")
    _write_status_shell()
    _start_status_server()
    _append_log(f"Started run with args: {sys.argv[1:]}")
    _update_status(
        "Run Started",
        f"Automation started. Status files: {Path(DEFAULT_STATUS_TEXT_FILE).resolve()} and {Path(DEFAULT_STATUS_HTML_FILE).resolve()}.",
    )
    _open_status_window(args.headless)
    _cancel_pending_workbook_open()
    _close_previous_automation_window()

    query_settings = _collect_query_settings(args, args.headless)
    args.vehicle_vin = query_settings["vehicle_vin"]
    args.test_requestor = query_settings["test_requestor"]
    args.start_date = query_settings["start_date"]
    args.sort_by = query_settings["sort_by"]
    args.download_ati_mf4 = query_settings.get("download_ati_mf4", "0") == "1"

    existing_result_workbook = _get_latest_result_workbook()
    if existing_result_workbook is not None:
        _append_log(f"Found existing result workbook for comparison: {existing_result_workbook}")
    else:
        _append_log("No existing result workbook found for comparison.")

    if args.populate_workbook:
        args.run_mode = _normalize_run_mode(args.run_mode or "fresh")
        try:
            workbook_path = _resolve_run_workbook(args.run_mode, args.workbook_path)
        except FileNotFoundError as exc:
            _append_log(f"Workbook selection failed: {exc}")
            _update_status("Workbook Selection Failed", str(exc), "failed")
            print(str(exc), file=sys.stderr)
            return 1
        args.workbook_path = str(workbook_path)
        _close_workbook_for_update(workbook_path)
        try:
            _assert_workbook_writable(workbook_path)
        except PermissionError as exc:
            if args.run_mode == "append":
                try:
                    retry_workbook_path = _create_locked_workbook_retry_copy(workbook_path)
                    _assert_workbook_writable(retry_workbook_path)
                    workbook_path = retry_workbook_path
                    args.workbook_path = str(workbook_path)
                    _append_log(
                        f"Workbook remained locked; continuing append run with retry workbook copy: {workbook_path}"
                    )
                    _update_status(
                        "Workbook Locked",
                        f"{exc} Continuing with unlocked workbook copy {workbook_path.name} instead.",
                    )
                except Exception as retry_exc:
                    keep_automation_browser_open = not args.headless
                    _append_log(f"Workbook write check failed: {exc}")
                    _append_log(
                        f"Retry workbook copy failed: {retry_exc.__class__.__name__}: {retry_exc}"
                    )
                    _update_status(
                        "Workbook Locked",
                        f"{exc} The Emissions Utilities browser will remain open so you can close Excel and retry.",
                        "failed",
                    )
                    print(str(exc), file=sys.stderr)
                    return 1
            else:
                keep_automation_browser_open = not args.headless
                _append_log(f"Workbook write check failed: {exc}")
                _update_status(
                    "Workbook Locked",
                    f"{exc} The Emissions Utilities browser will remain open so you can close Excel and retry.",
                    "failed",
                )
                print(str(exc), file=sys.stderr)
                return 1
        try:
            _update_status("Workbook Population Started", f"Reading PDFs from {Path(args.pdf_dir).resolve()} into {workbook_path.name}.")
            updated_rows = populate_workbook_from_pdfs(workbook_path, Path(args.pdf_dir))
        except Exception as exc:
            _append_log(f"Workbook population failed: {exc.__class__.__name__}: {exc}")
            _update_status("Workbook Population Failed", str(exc), "failed")
            print(str(exc), file=sys.stderr)
            return 1

        print(f"Updated {updated_rows} row(s) from PDF summaries.")
        _update_status("Workbook Population Complete", f"Updated {updated_rows} row(s) from PDF summaries. Excel will open next and the status window will stay open.", "completed")
        workbook_to_open = workbook_path
        _open_workbook(workbook_to_open)
        return 0

    archived_pdf_count, archived_import_count = _archive_run_directories()
    _append_log(
        "Archived startup files: "
        f"{archived_pdf_count} file(s) from {Path(DEFAULT_PDF_DIR).resolve()}, "
        f"{archived_import_count} file(s) from {Path(DEFAULT_DOWNLOAD_DIR).resolve()} (import downloads kept in place)."
    )
    _update_status(
        "Startup Cleanup",
        f"Archived {archived_pdf_count} file(s) from {DEFAULT_PDF_DIR}; kept files in {DEFAULT_DOWNLOAD_DIR}.",
    )

    try:
        browser_mode = "headless" if args.headless else "visible"
        _update_status("Browser Startup", f"Launching {args.browser} browser in {browser_mode} mode.")
        driver = create_driver(args.browser, args.headless)
    except Exception as exc:
        _append_log(f"Driver creation failed: {exc.__class__.__name__}: {exc}")
        _update_status("Browser Startup Failed", str(exc), "failed")
        print(str(exc), file=sys.stderr)
        return 1

    try:
        print(f"Navigating to: {TARGET_URL}")
        _update_status("Navigating", f"Opening {TARGET_URL}.")
        _open_target_url(driver, TARGET_URL, args.timeout)
        print("If you see a Microsoft sign-in page, complete the login in the browser window.")
        print(f"Waiting up to {args.timeout} seconds for '{SUCCESS_TITLE}'...")
        _update_status("Waiting For Access", "Watching for welcome page, sign-in, and final app load.")

        wait_for_access(driver, args.timeout)

        print("Access confirmed.")
        print(f"Title: {driver.title}")
        print(f"URL:   {driver.current_url}")
        print(f"Clicking through target: {args.click_target}")
        _update_status("Navigation Ready", f"Access confirmed. Continuing to {args.click_target}.")

        _prepare_and_submit_reprint_report_query(
            driver,
            args.vehicle_vin,
            args.test_requestor,
            args.start_date,
            args.sort_by,
            args.timeout,
        )

        downloaded_file = download_imported_spreadsheet(driver, args.timeout, args.headless)
        print("Requested spreadsheet download.")
        imported_ati_test_ids = _collect_imported_test_ids(downloaded_file)
        _update_status(
            "Monitor Schedule Download Started",
            "Downloading the current Live Share monitor schedule workbook.",
        )
        monitor_schedule_path = _resolve_monitor_schedule_workbook(driver)
        if monitor_schedule_path is None:
            _append_log("Monitor schedule workbook unavailable; monitor update will be skipped.")
            _update_status(
                "Monitor Schedule Download Unavailable",
                "The Live Share monitor schedule workbook was unavailable; monitor values will not be updated.",
                "completed",
            )
        else:
            _update_status(
                "Monitor Schedule Downloaded",
                f"Downloaded the Live Share monitor schedule workbook: {monitor_schedule_path.name}.",
                "completed",
            )
        ati_monitor_by_test_id = _get_ati_monitor_by_test_id(
            downloaded_file,
            monitor_schedule_path,
            existing_result_workbook,
        )
        if args.download_ati_mf4:
            ati_root_dir = _get_ati_directory()
            ati_test_ids = list(dict.fromkeys(imported_ati_test_ids + _get_summary_test_ids(existing_result_workbook)))
            monitor_named_test_ids = [
                test_id
                for test_id in ati_test_ids
                if _get_ati_monitor_keywords(ati_monitor_by_test_id.get(test_id, ""))
            ]
            prep_test_ids = [
                test_id
                for test_id in ati_test_ids
                if _get_ati_monitor_keywords(ati_monitor_by_test_id.get(test_id, ""))[:1] == ["Prep"]
            ]
            _update_status(
                "ATI Monitor Matched",
                f"Matched Monitor text for {len(ati_monitor_by_test_id)} TESTID(s); {len(monitor_named_test_ids)} current TESTID(s) have filename keywords, including {len(prep_test_ids)} Prep test(s).",
            )
            missing_ati_test_ids = [
                test_id
                for test_id in ati_test_ids
                if test_id not in _get_known_ati_test_ids(ati_root_dir)
            ]
            _update_status(
                "ATI Missing Files",
                f"Checking ATI for {len(missing_ati_test_ids)} missing TESTID(s), including Summary-only tests, for _001/_002 MF4 files.",
            )
            ati_download_result = download_combined_ati_mf4s(
                driver,
                ati_test_ids,
                ati_root_dir,
                args.timeout,
                ati_monitor_by_test_id,
                args.vehicle_vin,
            )
            if isinstance(ati_download_result, AtiDownloadJob):
                ati_download_job = ati_download_result
            else:
                downloaded_ati_files = ati_download_result
        else:
            _append_log("ATI _001/_002 MF4 download disabled by default; skipping ATI download steps.")
            _update_status(
                "ATI Download Skipped",
                "ATI _001/_002 MF4 download is disabled. Continuing without opening ATI Files in GTDM.",
                "completed",
            )
            _update_status(
                "ATI Files Renamed",
                    f"Renamed {renamed_ati_files} ATI MF4 file(s) after downloads completed.",
            )

        if existing_result_workbook is not None:
            _update_status(
                "Spreadsheet Compared",
                f"Comparing downloaded spreadsheet with {existing_result_workbook.name}. This can take a little while when many existing rows must be checked.",
            )
            comparison = _compare_imported_spreadsheet(downloaded_file, existing_result_workbook)
            new_test_ids = [item.get("TESTID", "") for item in comparison.new_test_details if item.get("TESTID")]
            updated_test_ids = [item.get("TESTID", "") for item in comparison.updated_test_details if item.get("TESTID")]
            preview = ", ".join(new_test_ids[:20]) if new_test_ids else "None"
            if len(new_test_ids) > 20:
                preview += ", ..."
            updated_preview = ", ".join(updated_test_ids[:20]) if updated_test_ids else "None"
            if len(updated_test_ids) > 20:
                updated_preview += ", ..."
            _append_log(
                f"Compared downloaded spreadsheet with {existing_result_workbook.name}: {len(new_test_ids)} new TESTID(s) -> {preview}; {len(updated_test_ids)} updated existing TESTID(s) -> {updated_preview}"
            )
            _update_status(
                "Spreadsheet Compared",
                f"Compared download with {existing_result_workbook.name}; found {len(new_test_ids)} new TESTID(s): {preview}; updated existing TESTID(s): {updated_preview}.",
            )
            args.run_mode = _normalize_run_mode(args.run_mode or _prompt_update_mode(existing_result_workbook, comparison))
            if LAST_UPDATE_MODE_SELECTION == "append" and args.run_mode != "append":
                _append_log(
                    "Corrected summary update mode to append because the status prompt selection recorded append."
                )
                args.run_mode = "append"
            _append_log(f"Final summary update mode before workbook selection: {args.run_mode}")
        else:
            args.run_mode = "fresh"
            _update_status(
                "Spreadsheet Compared",
                "No existing final workbook was found, so the run will continue with a full refresh from the template.",
            )

        try:
            workbook_path = _resolve_run_workbook(args.run_mode, args.workbook_path)
        except FileNotFoundError as exc:
            _append_log(f"Workbook selection failed: {exc}")
            _update_status("Workbook Selection Failed", str(exc), "failed")
            print(str(exc), file=sys.stderr)
            return 1
        args.workbook_path = str(workbook_path)
        mode_label = RUN_MODE_OPTIONS.get(args.run_mode, args.run_mode)
        _update_status("Workbook Selected", f"Using workbook {workbook_path.name} in {workbook_path.parent} with mode: {mode_label}.")

        _close_workbook_for_update(workbook_path)
        try:
            _assert_workbook_writable(workbook_path)
        except PermissionError as exc:
            if args.run_mode == "append":
                try:
                    retry_workbook_path = _create_locked_workbook_retry_copy(workbook_path)
                    _assert_workbook_writable(retry_workbook_path)
                    workbook_path = retry_workbook_path
                    args.workbook_path = str(workbook_path)
                    _append_log(
                        f"Workbook remained locked; continuing append run with retry workbook copy: {workbook_path}"
                    )
                    _update_status(
                        "Workbook Locked",
                        f"{exc} Continuing with unlocked workbook copy {workbook_path.name} instead.",
                    )
                except Exception as retry_exc:
                    keep_automation_browser_open = not args.headless
                    _append_log(f"Workbook write check failed: {exc}")
                    _append_log(
                        f"Retry workbook copy failed: {retry_exc.__class__.__name__}: {retry_exc}"
                    )
                    _update_status(
                        "Workbook Locked",
                        f"{exc} The Emissions Utilities browser will remain open so you can close Excel and retry.",
                        "failed",
                    )
                    print(str(exc), file=sys.stderr)
                    return 1
            else:
                keep_automation_browser_open = not args.headless
                _append_log(f"Workbook write check failed: {exc}")
                _update_status(
                    "Workbook Locked",
                    f"{exc} The Emissions Utilities browser will remain open so you can close Excel and retry.",
                    "failed",
                )
                print(str(exc), file=sys.stderr)
                return 1

        import_result = append_imported_results(workbook_path, downloaded_file)
        imported_rows = import_result.touched_rows
        print(
            f"Imported {len(import_result.appended_rows)} new row(s) and updated {len(import_result.updated_rows)} existing row(s) into the workbook."
        )

        monitor_updated_rows = 0
        if monitor_schedule_path is not None:
            monitor_updated_rows = update_monitor_values(workbook_path, monitor_schedule_path, driver=driver)
        else:
            _update_status("Monitor Update Skipped", "Monitor schedule workbook was not available; skipped monitor update.")

        pdf_dir = _get_pdf_directory()
        incomplete_summary_rows: list[int] = []
        if args.run_mode == "append":
            workbook_rows_with_pdf_and_csv = get_subset_rows_with_pdf_and_csv(workbook_path, imported_rows)
            incomplete_summary_rows = get_rows_with_incomplete_summary_pdf_and_csv(workbook_path)
            if incomplete_summary_rows:
                original_candidate_count = len(workbook_rows_with_pdf_and_csv)
                workbook_rows_with_pdf_and_csv = list(dict.fromkeys(workbook_rows_with_pdf_and_csv + incomplete_summary_rows))
                _append_log(
                    f"Added {len(workbook_rows_with_pdf_and_csv) - original_candidate_count} existing PDF+CSV row(s) with incomplete summary data to append-mode PDF enrichment."
                )
        else:
            workbook_rows_with_pdf_and_csv = get_rows_with_pdf_and_csv(workbook_path)
        _update_status(
            "Preparing PDF Download",
            f"Found {len(workbook_rows_with_pdf_and_csv)} QueryResults row(s) with both PDF and CSV entries.",
        )
        downloaded_pdfs = download_missing_linked_pdfs(
            driver,
            workbook_path,
            workbook_rows_with_pdf_and_csv,
            pdf_dir,
        )
        print(f"Downloaded {len(downloaded_pdfs)} missing linked PDF file(s) into {pdf_dir}.")

        enriched_rows = enrich_rows_from_local_pdfs(
            workbook_path,
            workbook_rows_with_pdf_and_csv,
            pdf_dir,
            rebuild_plots=args.run_mode != "append",
        )

        if args.run_mode == "append":
            if import_result.updated_rows or monitor_updated_rows or incomplete_summary_rows:
                rebuild_workbook_plots(workbook_path)
            else:
                append_workbook_plots(workbook_path, imported_rows)
        elif monitor_updated_rows:
            rebuild_workbook_plots(workbook_path)

        if ati_download_job is not None:
            _update_status("ATI Download Finishing", "Waiting for background ATI MF4 downloads before finalizing the run.")
            downloaded_ati_files = ati_download_job.wait()
            ati_download_job = None
        renamed_ati_files = _rename_ati_mf4_files(ati_root_dir, ati_monitor_by_test_id) if args.download_ati_mf4 else 0
        if renamed_ati_files:
            _append_log(f"Renamed {renamed_ati_files} ATI MF4 file(s) from Monitor values after downloads completed.")
        if args.download_ati_mf4:
            _refresh_ati_mf4_links(workbook_path)

        print(f"PDF-enriched {enriched_rows} row(s) with both PDF and CSV entries.")
        try:
            _restore_query_results_bottom_view(
                driver,
                args.vehicle_vin,
                args.test_requestor,
                args.start_date,
                args.sort_by,
                args.timeout,
            )
        except Exception as exc:
            _append_log(
                f"Could not restore browser to query results bottom view: {exc.__class__.__name__}: {exc}"
            )
        if monitor_schedule_path is not None and monitor_schedule_path.name == DEFAULT_MONITOR_SOURCE_LIVE_FILE:
            live_share_tab_opened = _open_live_share_workbook_tab(driver)
        keep_automation_browser_open = not args.headless
        _update_status(
            "Workbook Update Complete",
            f"Imported {len(import_result.appended_rows)} new row(s), updated {len(import_result.updated_rows)} existing row(s), downloaded {len(downloaded_pdfs)} PDF(s), prepared {len(downloaded_ati_files)} ATI _001/_002 MF4 file(s), renamed {renamed_ati_files} existing ATI MF4 file(s), enriched {enriched_rows} row(s), and updated monitor values for {monitor_updated_rows} row(s). The automation browser is {'hidden and will close' if args.headless else 'visible and will stay open'}; this status window will stay open.",
        )
        workbook_to_open = workbook_path
        result_code = 0
    except AptlNavigationError as exc:
        keep_automation_browser_open = not args.headless
        _append_log(f"APTL navigation failed: {exc}")
        _update_status("APTL Server Unavailable", str(exc), "failed")
        print(str(exc), file=sys.stderr)
        result_code = 1
    except TimeoutException:
        _append_log("Timed out waiting for Emissions Utilities page.")
        _update_status("Run Timed Out", f"Timed out while waiting at title '{driver.title}' and URL '{driver.current_url}'.", "failed")
        print("Timed out before the Emissions Utilities page finished loading.", file=sys.stderr)
        print(f"Current title: {driver.title}", file=sys.stderr)
        print(f"Current URL:   {driver.current_url}", file=sys.stderr)
        result_code = 2
    except Exception as exc:
        _append_log(f"Run failed: {exc.__class__.__name__}: {exc}")
        _update_status("Run Failed", str(exc), "failed")
        raise
    finally:
        if ati_download_job is not None:
            try:
                downloaded_ati_files = ati_download_job.wait()
                if ati_root_dir is not None:
                    renamed_ati_files = _rename_ati_mf4_files(ati_root_dir, ati_monitor_by_test_id)
            except Exception as exc:
                _append_log(f"Background ATI MF4 download cleanup failed: {exc.__class__.__name__}: {exc}")
            ati_download_job = None
        if keep_automation_browser_open:
            _append_log("Preserving automation browser window for post-run review.")
            _preserve_automation_browser(driver)
        else:
            _append_log("Closing browser driver.")
            _clear_automation_pid()
            driver.quit()

    if result_code == 0 and workbook_to_open is not None:
        _update_status(
            "Opening Excel",
            f"Opening {workbook_to_open.name}. The automation browser is {'hidden and closed' if args.headless else 'open'}; the status window will remain open.",
            "completed",
        )
        _open_workbook(workbook_to_open)
        time.sleep(0.75)
        completion_detail = (
            f"Completed run and opened {workbook_to_open.name}. "
            "The Emissions Utilities browser and status window remain open until the next run resets them."
        )
        if live_share_tab_opened:
            completion_detail = (
                f"Completed run and opened {workbook_to_open.name}, and opened the Live Share workbook in another tab of the Emissions Utilities browser. "
                "The Emissions Utilities browser and status window remain open until the next run resets them."
            )
        _update_status("Run Complete", completion_detail, "completed")
        time.sleep(0.5)

    return result_code


if __name__ == "__main__":
    raise SystemExit(main())