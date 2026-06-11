import os
import re
import json
import sys
import requests
from urllib.parse import quote
from dotenv import load_dotenv, find_dotenv

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
ERP_PROCESSED_PATH = os.path.join(
    os.path.dirname(__file__), "Downloaded_RFQs", "erp_processed.json"
)
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
        # Any other HTTP error: treat as not found but warn
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
    # Pass 1 — dedicated RFQ field
    if _search_opportunity(
        [["Opportunity", "custom_client_rfq_number", "like", f"%{rfq_number}%"]]
    ):
        return True

    # Pass 2 — fallback: opportunity document name contains the number
    return _search_opportunity(
        [["Opportunity", "name", "like", f"%{rfq_number}%"]]
    )


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    if not os.path.exists(ERP_PROCESSED_PATH):
        sys.exit(f"File not found: {ERP_PROCESSED_PATH}")

    with open(ERP_PROCESSED_PATH, "r", encoding="utf-8") as fh:
        entries: list[str] = json.load(fh)

    print(f"Loaded {len(entries)} entries from erp_processed.json\n")

    kept:    list[str] = []
    removed: list[str] = []
    skipped: list[str] = []

    for entry in entries:
        rfq_number = extract_rfq_number(entry)

        if rfq_number is None:
            print(f"[SKIP]   {entry}")
            skipped.append(entry)
            kept.append(entry)
            continue

        found = find_opportunity_by_rfq(rfq_number)
        if found:
            print(f"[KEEP]   RFQ#{rfq_number}  →  {entry}")
            kept.append(entry)
        else:
            print(f"[REMOVE] RFQ#{rfq_number}  →  {entry}")
            removed.append(entry)

    # Write cleaned file
    with open(ERP_PROCESSED_PATH, "w", encoding="utf-8") as fh:
        json.dump(kept, fh, indent=2, ensure_ascii=False)

    print(f"\n{'='*60}")
    print(f"Done.")
    print(f"  Kept    : {len(kept)}")
    print(f"  Removed : {len(removed)}")
    print(f"  Skipped : {len(skipped)}  (no RFQ# — kept)")

    if removed:
        print("\nRemoved entries:")
        for r in removed:
            print(f"  - {r}")

    if skipped:
        print("\nSkipped entries (no extractable RFQ number — kept as-is):")
        for s in skipped:
            print(f"  - {s}")


if __name__ == "__main__":
    main()
