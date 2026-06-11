import os
import re
import json
import sys
import requests
from dotenv import load_dotenv, find_dotenv

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
DOWNLOADED_RFQS_DIR = os.path.join(os.path.dirname(__file__), "Downloaded_RFQs")
TIMEOUT = 30

dotenv_path = find_dotenv()
if not dotenv_path:
    sys.exit(".env file not found — run from the RFQ scraper directory.")
load_dotenv(dotenv_path, override=True)

ERPNEXT_URL        = os.getenv("ERPNEXT_URL", "").rstrip("/")
ERPNEXT_API_KEY    = os.getenv("ERPNEXT_API_KEY", "")
ERPNEXT_API_SECRET = os.getenv("ERPNEXT_API_SECRET", "")

if not ERPNEXT_URL:
    sys.exit("Missing ERPNEXT_URL in .env")
if not ERPNEXT_API_KEY or not ERPNEXT_API_SECRET:
    sys.exit("Missing ERPNEXT_API_KEY or ERPNEXT_API_SECRET in .env — paste your credentials first.")

HEADERS = {
    "Authorization": f"token {ERPNEXT_API_KEY}:{ERPNEXT_API_SECRET}",
    "Accept":        "application/json",
}

# Timestamp suffix pattern at the end of every folder name: _YYYYMMDD-HHMMSS
_TS_PATTERN = re.compile(r"_\d{8}-\d{6}$")


# ---------------------------------------------------------------------------
# RFQ number extraction
# ---------------------------------------------------------------------------
def extract_rfq_number(entry: str) -> str | None:
    """
    Returns the first meaningful RFQ number found in a folder-name entry,
    or None if no numeric RFQ ID can be determined.

    Priority order:
    1. RFQ#NUMBER  or  RFQ_#_NUMBER  (the explicit client RFQ reference)
    2. A_-_NUMBER_-_                 (A-series folder format)
    3. RFQ_NUMBER at start           (RFQ-prefixed folder without explicit #)
    4. First standalone 6-7 digit number (general fallback)

    The trailing timestamp (_YYYYMMDD-HHMMSS) is stripped first so that the
    6-digit time component (e.g. 090344) is never mistaken for an RFQ number.
    """
    clean = _TS_PATTERN.sub("", entry)

    # 1) Explicit RFQ# reference: RFQ#2156036, RFQ_#_1020272, RFQ#_1019980
    m = re.search(r"RFQ[\s_]*#[\s_]*(\d{5,8})", clean, re.IGNORECASE)
    if m:
        return m.group(1)

    # 2) A-series: A_-_1061487_-_...
    m = re.search(r"\bA_-_(\d{5,8})_-_", clean, re.IGNORECASE)
    if m:
        return m.group(1)

    # 3) RFQ_NUMBER prefix: RFQ_1057244_-_..., RFQ_1020701_-_...
    m = re.search(r"\bRFQ_(\d{5,8})(?:_-_|\b)", clean, re.IGNORECASE)
    if m:
        return m.group(1)

    # 4) Any standalone 6-7 digit number (avoids 8-digit dates)
    numbers = re.findall(r"\b(\d{6,7})\b", clean)
    if numbers:
        return numbers[0]

    return None


# ---------------------------------------------------------------------------
# ERPNext search
# ---------------------------------------------------------------------------
def _search_opportunity(filters: list) -> bool:
    """Return True if at least one Opportunity matches the given filters."""
    try:
        r = requests.get(
            f"{ERPNEXT_URL}/api/resource/Opportunity",
            headers=HEADERS,
            params={
                "filters":           json.dumps(filters),
                "fields":            json.dumps(["name"]),
                "limit_page_length": 1,
            },
            timeout=TIMEOUT,
        )
        if r.status_code == 200:
            return bool(r.json().get("data"))
        if r.status_code == 401:
            sys.exit("ERPNext returned 401 — check ERPNEXT_API_KEY and ERPNEXT_API_SECRET.")
        print(f"  [WARN] ERP returned {r.status_code}: {r.text[:120]}")
        return False
    except requests.RequestException as exc:
        print(f"  [WARN] Network error: {exc}")
        return False


def find_opportunity_by_rfq(rfq_number: str) -> bool:
    """
    Two-pass search:
      1. custom_client_rfq_number LIKE %rfq_number%
      2. Opportunity name          LIKE %rfq_number%
    Returns True if found by either pass.
    """
    if _search_opportunity(
        [["Opportunity", "custom_client_rfq_number", "like", f"%{rfq_number}%"]]
    ):
        return True

    return _search_opportunity(
        [["Opportunity", "name", "like", f"%{rfq_number}%"]]
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    if not os.path.isdir(DOWNLOADED_RFQS_DIR):
        sys.exit(f"Directory not found: {DOWNLOADED_RFQS_DIR}")

    folders = sorted(
        name for name in os.listdir(DOWNLOADED_RFQS_DIR)
        if os.path.isdir(os.path.join(DOWNLOADED_RFQS_DIR, name))
    )

    print(f"Found {len(folders)} folders in Downloaded_RFQs\n")

    found:   list[tuple[str, str]] = []
    missing: list[tuple[str, str]] = []
    skipped: list[str]             = []

    for folder in folders:
        rfq_number = extract_rfq_number(folder)

        if rfq_number is None:
            print(f"[SKIP]    {folder}")
            skipped.append(folder)
            continue

        in_erp = find_opportunity_by_rfq(rfq_number)
        if in_erp:
            print(f"[FOUND]   RFQ#{rfq_number}  →  {folder}")
            found.append((rfq_number, folder))
        else:
            print(f"[MISSING] RFQ#{rfq_number}  →  {folder}")
            missing.append((rfq_number, folder))

    print(f"\n{'='*60}")
    print(f"Total folders : {len(folders)}")
    print(f"Found in ERP  : {len(found)}")
    print(f"Missing in ERP: {len(missing)}")
    print(f"Skipped       : {len(skipped)}  (no RFQ# extractable)")

    if missing:
        print("\nMissing RFQs (not in ERPNext):")
        for rfq_num, folder in missing:
            print(f"  - {folder}  (RFQ#{rfq_num})")

    if skipped:
        print("\nSkipped folders (no extractable RFQ number):")
        for s in skipped:
            print(f"  - {s}")


if __name__ == "__main__":
    main()
