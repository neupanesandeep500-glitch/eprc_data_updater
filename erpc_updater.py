import os
import re
import io
import json
import sys
import time
from datetime import datetime
from urllib.parse import urljoin

import requests
import gspread
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from google.oauth2.service_account import Credentials

ERPC_URL = "https://www.erpc.gov.in/dsm-rras-fras-and-sced-accts"
GOOGLE_SHEET_ID = os.environ["GOOGLE_SHEET_ID"]
SERVICE_ACCOUNT_JSON = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]

NEA_SHEET = "NEA-Bihar"
NVVN_SHEET = "NVVN-Nepal"
CONTROL_SHEET = "_CONTROL"

session = requests.Session()
session.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/151.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,"
    "image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.erpc.gov.in/",
})


def connect_google_sheet():
    info = json.loads(SERVICE_ACCOUNT_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds).open_by_key(GOOGLE_SHEET_ID)


def get_or_create_sheet(spreadsheet, name):
    try:
        return spreadsheet.worksheet(name)
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(title=name, rows=1000, cols=100)


def prepare_control_sheet(spreadsheet):
    sheet = get_or_create_sheet(spreadsheet, CONTROL_SHEET)
    headers = [
        "Settlement Period", "Account Type", "ERPC Account",
        "ERPC Start Date", "ERPC End Date", "Data Files 1 URL",
        "NEA Rows", "NVVN Rows", "Status", "Imported At"
    ]
    values = sheet.get_all_values()
    if not values:
        sheet.update("A1", [headers])
    elif values[0] != headers:
        sheet.update("A1", [headers])
    return sheet


DATE_PATTERN = re.compile(
    r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})"
    r"\s*(?:to|\-|\–|\—)\s*"
    r"(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})",
    re.I,
)


def extract_date_range(text):
    match = DATE_PATTERN.search(text or "")
    if not match:
        return None
    start = datetime(int(match.group(3)), int(match.group(2)), int(match.group(1)))
    end = datetime(int(match.group(6)), int(match.group(5)), int(match.group(4)))
    return start, end


def is_relevant_account(text):
    t = " ".join((text or "").split()).lower()
    if "dsm" not in t or "account" not in t:
        return False
    if "revised" in t:
        return False
    if re.search(r"\bsced\s+account\b", t):
        return False
    return True


def fetch_with_retry(url, max_retries=3, **kwargs):
    """Fetch URL with exponential backoff retry."""
    for attempt in range(max_retries):
        try:
            response = session.get(url, timeout=60, **kwargs)
            response.raise_for_status()
            return response
        except Exception as e:
            print(f"Attempt {attempt + 1}/{max_retries} failed: {e}")
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)
            else:
                raise


def find_latest_account():
    print("Opening ERPC page:", ERPC_URL)
    response = fetch_with_retry(ERPC_URL)
    print("ERPC HTTP status:", response.status_code)

    soup = BeautifulSoup(response.text, "html.parser")
    candidates = []

    for link in soup.find_all("a", href=True):
        link_text = " ".join(link.get_text(" ", strip=True).split())
        if not re.search(r"data\s*files?\s*1", link_text, re.I):
            continue

        href = urljoin(ERPC_URL, link["href"])
        containers = []
        node = link

        for _ in range(6):
            node = node.parent
            if not node:
                break
            txt = " ".join(node.get_text(" ", strip=True).split())
            if txt:
                containers.append(txt)

        account_text = None
        for txt in containers:
            if is_relevant_account(txt) and extract_date_range(txt):
                account_text = txt
                break

        if not account_text:
            node = link
            for _ in range(20):
                node = node.find_previous()
                if not node:
                    break
                if hasattr(node, 'get_text'):
                    txt = " ".join(node.get_text(" ", strip=True).split())
                    if txt and len(txt) < 1200:
                        if is_relevant_account(txt) and extract_date_range(txt):
                            account_text = txt
                            break

        if not account_text:
            continue

        dates = extract_date_range(account_text)
        if dates:
            candidates.append({
                "title": account_text,
                "start": dates[0],
                "end": dates[1],
                "data_url": href,
            })

    if not candidates:
        # Debug: print all links found to help diagnose page structure changes
        print("DEBUG: All links on page:")
        for link in soup.find_all("a", href=True):
            txt = " ".join(link.get_text(" ", strip=True).split())
            if txt:
                print(f"  - {txt[:100]} -> {link['href'][:100]}")
        raise RuntimeError(
            "No normal DSM settlement account with 'Data Files 1' "
            "was found on the ERPC page."
        )

    unique = {}
    for item in candidates:
        key = (
            item["start"].date().isoformat(),
            item["end"].date().isoformat(),
            item["data_url"],
        )
        unique[key] = item

    candidates = sorted(
        unique.values(),
        key=lambda x: (x["start"], x["end"]),
        reverse=True,
    )
    latest = candidates[0]

    print("Latest account:", latest["title"])
    print("Data Files 1:", latest["data_url"])
    return latest


def download_excel(url):
    print("Downloading:", url)
    response = fetch_with_retry(url, timeout=120)
    print("Excel HTTP status:", response.status_code)
    if len(response.content) < 100:
        raise RuntimeError("Downloaded file is unexpectedly small.")
    return response.content


def read_excel_sheet(excel_bytes, target_name):
    # FIX 1: Removed read_only=True because data_only=True has NO EFFECT
    # in read-only mode. All formula cells would return None, breaking
    # header detection and data extraction.
    try:
        wb = load_workbook(io.BytesIO(excel_bytes), data_only=True)
    except Exception as e:
        print(f"Warning: Failed to load with data_only=True: {e}")
        wb = load_workbook(io.BytesIO(excel_bytes))

    target = None
    normalized_target = re.sub(r"[\s_\-]", "", target_name.lower())

    for name in wb.sheetnames:
        normalized = re.sub(r"[\s_\-]", "", name.lower())
        if normalized == normalized_target:
            target = name
            break

    if target is None:
        for name in wb.sheetnames:
            normalized = re.sub(r"[\s_\-]", "", name.lower())
            if normalized_target in normalized or normalized in normalized_target:
                target = name
                break

    if target is None:
        names = ", ".join(wb.sheetnames)
        wb.close()
        raise RuntimeError(
            f"Excel sheet '{target_name}' not found. Available sheets: {names}"
        )

    print("Reading Excel sheet:", target)
    rows = []
    for row in wb[target].iter_rows(values_only=True):
        values = list(row)
        if any(v is not None and str(v).strip() != "" for v in values):
            rows.append(values)
    wb.close()
    return rows


def find_header_row(rows):
    keywords = [
        "date", "time", "block", "frequency", "schedule",
        "actual", "deviation", "mcp", "amount", "energy",
    ]
    for i, row in enumerate(rows[:60]):
        text = " ".join(str(v).lower() for v in row if v is not None)
        if sum(k in text for k in keywords) >= 2:
            return i
    return 0


def prepare_data(rows, account):
    if not rows:
        return [], []

    header_i = find_header_row(rows)
    source_headers = rows[header_i]
    data_rows = rows[header_i + 1:]

    headers = [
        "Settlement Period", "ERPC Start Date",
        "ERPC End Date", "ERPC Account",
    ] + ["" if h is None else str(h) for h in source_headers]

    period = (
        account["start"].strftime("%d.%m.%Y")
        + " to " + account["end"].strftime("%d.%m.%Y")
    )

    output = []
    for row in data_rows:
        if not any(v is not None and str(v).strip() != "" for v in row):
            continue
        output.append([
            period,
            account["start"].strftime("%d.%m.%Y"),
            account["end"].strftime("%d.%m.%Y"),
            account["title"],
        ] + list(row))

    return headers, output


def existing_periods(sheet):
    values = sheet.col_values(1)
    return {str(v).strip() for v in values[1:] if str(v).strip()}


def ensure_header(sheet, headers):
    values = sheet.get_all_values()
    if not values:
        sheet.update("A1", [headers])
    elif len(values) == 1 and values[0] == ["Settlement Period"]:
        sheet.clear()
        sheet.update("A1", [headers])


def insert_new_rows(sheet, headers, rows):
    if not rows:
        return 0

    ensure_header(sheet, headers)

    # FIX 2: Replaced empty-row insertion + update with direct padded-row insertion.
    # The old code: sheet.insert_rows([[] for _ in rows], row=2) often fails
    # because Google Sheets API rejects rows with zero columns.
    end_col = max(len(headers), max(len(r) for r in rows))
    padded_rows = [r + [""] * (end_col - len(r)) for r in rows]

    # Batch insert to stay within Google API limits
    BATCH_SIZE = 500
    total_inserted = 0
    for i in range(0, len(padded_rows), BATCH_SIZE):
        batch = padded_rows[i:i + BATCH_SIZE]
        sheet.insert_rows(batch, row=2)
        total_inserted += len(batch)

    return total_inserted


def add_control_record(control, account, nea_count, nvvn_count, status):
    period = (
        account["start"].strftime("%d.%m.%Y")
        + " to " + account["end"].strftime("%d.%m.%Y")
    )
    control.append_row([
        period,
        "DSM-SRAS-TRAS-SCUC",
        account["title"],
        account["start"].strftime("%d.%m.%Y"),
        account["end"].strftime("%d.%m.%Y"),
        account["data_url"],
        nea_count,
        nvvn_count,
        status,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    ])


def main():
    print("=" * 60)
    print("ERPC LATEST DSM UPDATE")
    print("=" * 60)

    try:
        spreadsheet = connect_google_sheet()
        nea_sheet = get_or_create_sheet(spreadsheet, NEA_SHEET)
        nvvn_sheet = get_or_create_sheet(spreadsheet, NVVN_SHEET)
        control = prepare_control_sheet(spreadsheet)

        account = find_latest_account()
        period = (
            account["start"].strftime("%d.%m.%Y")
            + " to " + account["end"].strftime("%d.%m.%Y")
        )

        print("Settlement period:", period)

        if period in set(control.col_values(1)):
            print("Already imported. No action required.")
            return

        excel = download_excel(account["data_url"])

        nea_raw = read_excel_sheet(excel, "NEA-Bihar")
        nvvn_raw = read_excel_sheet(excel, "NVVN-Nepal")

        nea_headers, nea_rows = prepare_data(nea_raw, account)
        nvvn_headers, nvvn_rows = prepare_data(nvvn_raw, account)

        nea_count = 0
        if period not in existing_periods(nea_sheet):
            nea_count = insert_new_rows(nea_sheet, nea_headers, nea_rows)

        nvvn_count = 0
        if period not in existing_periods(nvvn_sheet):
            nvvn_count = insert_new_rows(nvvn_sheet, nvvn_headers, nvvn_rows)

        add_control_record(
            control, account, nea_count, nvvn_count, "SUCCESS"
        )

        print("=" * 60)
        print("UPDATE SUCCESSFUL")
        print("Period:", period)
        print("NEA rows:", nea_count)
        print("NVVN rows:", nvvn_count)
        print("=" * 60)

    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
