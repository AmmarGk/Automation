import os
import json
import time
import imaplib
import email as email_lib
import email.utils as eutils
import datetime as dt
import random
import re
import requests
import ctypes
import atexit
import base64

from dotenv import load_dotenv, find_dotenv

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

# ==========================================================
# WINDOWS 11: PREVENT PC SLEEP WHILE SCRIPT RUNS
# ==========================================================
def prevent_sleep_windows():
    ES_CONTINUOUS        = 0x80000000
    ES_SYSTEM_REQUIRED   = 0x00000001
    ES_AWAYMODE_REQUIRED = 0x00000040
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_AWAYMODE_REQUIRED
        )
    except Exception:
        pass

    def restore():
        try:
            ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        except Exception:
            pass
    atexit.register(restore)


# ==========================================================
# ENV
# ==========================================================
dotenv_path = find_dotenv()
if not dotenv_path:
    raise SystemExit(".env file not found.")
load_dotenv(dotenv_path, override=True)

ERPNEXT_URL        = os.getenv("ERPNEXT_URL", "").rstrip("/")
ERPNEXT_API_KEY    = os.getenv("ERPNEXT_API_KEY", "")
ERPNEXT_API_SECRET = os.getenv("ERPNEXT_API_SECRET", "")
TIMEOUT            = int(os.getenv("TIMEOUT", "60"))

KT_USER = os.getenv("KTENDERING_USER_ID")
KT_PASS = os.getenv("KTENDERING_PASSWORD")

# --- IMAP credentials (Hostinger) — OTP fetch only ---
IMAP_SERVER    = os.getenv("IMAP_SERVER",    "imap.hostinger.com")
IMAP_PORT      = int(os.getenv("IMAP_PORT",  "993"))
EMAIL_USER     = os.getenv("EMAIL_USER",     "info@gkepoxy.com")
EMAIL_APP_PASS = os.getenv("EMAIL_APP_PASS")

if not all([EMAIL_USER, EMAIL_APP_PASS]):
    raise SystemExit("Missing EMAIL_USER or EMAIL_APP_PASS in .env — needed for OTP retrieval.")

if not all([ERPNEXT_URL, ERPNEXT_API_KEY, ERPNEXT_API_SECRET]):
    raise SystemExit("Missing ERP env vars: ERPNEXT_URL / ERPNEXT_API_KEY / ERPNEXT_API_SECRET")
if not all([KT_USER, KT_PASS]):
    raise SystemExit("Missing KT env vars: KTENDERING_USER_ID / KTENDERING_PASSWORD")

HEADERS = {
    "Authorization": f"token {ERPNEXT_API_KEY}:{ERPNEXT_API_SECRET}",
    "Accept":        "application/json",
    "Content-Type":  "application/json",
}

# ==========================================================
# CONFIG
# ==========================================================
RFX_FIELDNAME              = "custom_client_rfq_number"
SALES_STATUS_FIELDNAME     = "sales_stage"
EXPECTED_CLOSING_FIELDNAME = "expected_closing"

SALES_STAGE_BID_WON              = "Bid won"
SALES_STAGE_COMPLETED            = "Completed"
SALES_STAGE_BID_LOST             = "Bid Lost"
SALES_STAGE_NP                   = "N/P"
SALES_STAGE_AWARDED              = "Awarded"
SALES_STAGE_we_Awarded           = "We Awarded"
SALES_STAGE_Clarification_Pending = "Clarification Pending"
INCLUDE_ALL_RFQS         = True
INCLUDE_ALL_SALES_STAGES = False

SKIP_SALES_STAGES = {
    SALES_STAGE_BID_WON,
    SALES_STAGE_COMPLETED,
    SALES_STAGE_BID_LOST,
    SALES_STAGE_NP,
    SALES_STAGE_we_Awarded,
    SALES_STAGE_AWARDED,
    SALES_STAGE_Clarification_Pending
}

OPP_PAGE_SIZE = int(os.getenv("OPP_PAGE_SIZE", "250"))

COMPANY_NAME_ALIASES = {
    "TWENTY FOUR SEVEN CO FOR GENERAL TRADING":      "Twenty-Four Seven Co for General Trading",
    "TWENTY-FOUR SEVEN CO FOR GENERAL TRADING":      "Twenty-Four Seven Co for General Trading",
    "TWENTY FOUR SEVEN COMPANY FOR GENERAL TRADING": "Twenty-Four Seven Co for General Trading",
    "GLOBALKEMYA INDIA LLP":                         "Globalkemya India LLP",
}

OUR_COMPANY_CANONICAL = [
    "Globalkemya India LLP",
    "Twenty-Four Seven Co for General Trading",
    "Globalkemya",
]

COMPETITOR_DOCTYPE    = "Competitor"
COMPANY_DOCTYPE       = "Company"
COMPETITOR_TYPE_VALUE = "Competitor"
COMPANY_TYPE_VALUE    = "Company"

BID_RESULT_CURRENCY = "KWD"
MAIN_URL  = "https://ktendering.com.kw/esop/kuw-kpc-host/public/ktendering/web/login.html"
SPLIT_MSG = "ATTENTION: Review required – query is subject to SPLIT."

CURRENCY_TO_KWD = {
    "KWD": 1.0,      "USD": 1/3.259,   "EUR": 1/2.815,   "GBP": 1/2.476,   "AED": 1/11.969,
    "SAR": 1/12.222, "QAR": 1/84.080,  "OMR": 1/1.254,   "BHD": 1/1.225,   "CNY": 1/23.214,
    "JPY": 1/498.843,"INR": 1/289.089, "CAD": 1/4.587,   "AUD": 1/5.036,   "NZD": 1/5.806,
    "SGD": 1/4.243,  "CHF": 1/2.621,   "SEK": 1/31.085,  "NOK": 1/33.111,  "CZK": 1/68.426,
    "DKK": 1/21.025, "HKD": 1/25.348,  "TWD": 1/101.057, "THB": 1/105.644, "PHP": 1/192.393,
    "PLN": 1/11.946, "RON": 1/14.321,  "ZAR": 1/56.471,  "MXN": 1/60.309,  "BRL": 1/17.426,
    "TRY": 1/137.581,"RUB": 1/263.817, "PKR": 1/920.326, "LKR": 1/992.989, "NPR": 1/462.760,
    "MYR": 1/13.611, "BWP": 1/45.920,  "CLP": 1/3082.820,"COP": 1/12324.414,"ARS": 1/4667.700,
    "IDR": 1/54423.319, "IRR": 1/137209.743
}

# ==========================================================
# HTTP RETRY LAYER
# ==========================================================
RETRY_STATUS_CODES   = {502, 503, 504, 520, 521, 522, 524}
MAX_RETRIES          = int(os.getenv("ERP_HTTP_MAX_RETRIES", "6"))
BACKOFF_BASE_SECONDS = float(os.getenv("ERP_HTTP_BACKOFF", "1.5"))

def request_with_retry(method, url, *, headers=None, params=None, data=None, timeout=60):
    last_exc = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = requests.request(
                method=method, url=url, headers=headers,
                params=params, data=data, timeout=timeout
            )
            if r.status_code in RETRY_STATUS_CODES:
                sleep_s = (BACKOFF_BASE_SECONDS ** (attempt - 1)) + random.uniform(0, 0.7)
                print(f"[WARN] {method} {url} -> {r.status_code}. Retry {attempt}/{MAX_RETRIES} in {sleep_s:.1f}s")
                time.sleep(sleep_s)
                continue
            return r
        except (requests.exceptions.Timeout, requests.exceptions.ConnectionError,
                requests.exceptions.ChunkedEncodingError) as e:
            last_exc = e
            sleep_s  = (BACKOFF_BASE_SECONDS ** (attempt - 1)) + random.uniform(0, 0.7)
            print(f"[WARN] network error: {e}. Retry {attempt}/{MAX_RETRIES} in {sleep_s:.1f}s")
            time.sleep(sleep_s)
    raise RuntimeError(f"{method} {url} failed after {MAX_RETRIES} retries. Last: {last_exc}")


# ==========================================================
# OTP HELPER  (Hostinger IMAP)
# ==========================================================
def fetch_otp_from_email(sent_after_timestamp: float, max_wait: int = 90, poll_interval: int = 5) -> str:
    ts_str = time.strftime('%H:%M:%S', time.localtime(sent_after_timestamp))
    print(f"\n🔐 Waiting for OTP via {IMAP_SERVER}:{IMAP_PORT} "
          f"(sent after {ts_str}, timeout {max_wait}s)…")
    deadline = time.time() + max_wait

    while time.time() < deadline:
        try:
            mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
            mail.login(EMAIL_USER, EMAIL_APP_PASS)
            mail.select("INBOX")

            _, data = mail.search(None, '(UNSEEN FROM "noreply@jaggaer.com")')
            mail_ids = data[0].split()
            if not mail_ids:
                _, data = mail.search(None, "UNSEEN")
                mail_ids = data[0].split()

            for mid in reversed(mail_ids):
                _, msg_data = mail.fetch(mid, "(RFC822)")
                raw_email   = msg_data[0][1]
                msg         = email_lib.message_from_bytes(raw_email)

                sender = msg.get("From", "")
                if "jaggaer" not in sender.lower():
                    continue

                date_str = msg.get("Date", "")
                try:
                    email_ts = eutils.parsedate_to_datetime(date_str).timestamp()
                except Exception:
                    email_ts = 0.0

                if email_ts < (sent_after_timestamp - 30):
                    print(f"   ⏭  Skipping old OTP email (sent {time.strftime('%H:%M:%S', time.localtime(email_ts))})")
                    continue

                body = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        ct = part.get_content_type()
                        if ct == "text/plain":
                            body += part.get_payload(decode=True).decode(errors="ignore")
                        elif ct == "text/html" and not body:
                            body += part.get_payload(decode=True).decode(errors="ignore")
                else:
                    body = msg.get_payload(decode=True).decode(errors="ignore")

                subject   = msg.get("Subject", "")
                full_text = subject + "\n" + body
                print(f"   📨 Email from: {sender[:60]}")
                print(f"   Subject: {subject[:80]}")

                otp_match = re.search(
                    r"verification\s+code\s+is[:\s]+([A-Za-z0-9]{4,12})",
                    full_text, re.IGNORECASE
                )
                if not otp_match:
                    otp_match = re.search(
                        r"(?:^|\b)(?:code|token|otp)[:\s]+([A-Za-z0-9]{4,12})\b",
                        full_text, re.IGNORECASE
                    )
                if not otp_match:
                    for candidate in re.findall(r"\b([A-Za-z0-9]{4,12})\b", body):
                        if re.search(r"[A-Za-z]", candidate) and re.search(r"[0-9]", candidate):
                            class _M:
                                def __init__(self, v): self._v = v
                                def group(self, n): return self._v
                            otp_match = _M(candidate)
                            break

                if otp_match:
                    otp = otp_match.group(1)
                    print(f"✅ OTP found: {otp}  (email at {time.strftime('%H:%M:%S', time.localtime(email_ts))})")
                    mail.store(mid, "+FLAGS", "\\Seen")
                    mail.logout()
                    return otp
                else:
                    print(f"   ⚠️  Email matched but OTP not extracted. Body: {body[:100]}")

            mail.logout()

        except Exception as e:
            print(f"   IMAP error: {e}")

        remaining = int(deadline - time.time())
        if remaining > 0:
            print(f"   Retrying in {poll_interval}s… ({remaining}s left)")
            time.sleep(poll_interval)

    raise TimeoutError("❌ OTP email from JAGGAER did not arrive in time.")


# ==========================================================
# ERP HELPERS
# ==========================================================
def q(s):
    return requests.utils.quote(str(s))

def erp_get(path, params=None):
    url = f"{ERPNEXT_URL}{path}"
    r   = request_with_retry("GET", url, headers=HEADERS, params=params, timeout=TIMEOUT)
    if r.status_code >= 400:
        raise RuntimeError(f"GET {url} [{r.status_code}]: {r.text}")
    return r.json()

def erp_post(path, payload):
    url = f"{ERPNEXT_URL}{path}"
    r   = request_with_retry("POST", url, headers=HEADERS, data=json.dumps(payload), timeout=TIMEOUT)
    if r.status_code >= 400:
        raise RuntimeError(f"POST {url} [{r.status_code}]: {r.text}")
    return r.json()

def erp_put(path, payload):
    url = f"{ERPNEXT_URL}{path}"
    r   = request_with_retry("PUT", url, headers=HEADERS, data=json.dumps(payload), timeout=TIMEOUT)
    if r.status_code >= 400:
        raise RuntimeError(f"PUT {url} [{r.status_code}]: {r.text}")
    return r.json()

def erp_try_get_resource(doctype, name):
    url = f"{ERPNEXT_URL}/api/resource/{q(doctype)}/{q(name)}"
    r   = request_with_retry("GET", url, headers=HEADERS, timeout=TIMEOUT)
    if r.status_code == 200:
        return True, (r.json().get("data") or {})
    if r.status_code == 404:
        return False, None
    if r.status_code >= 400:
        raise RuntimeError(f"GET {url} [{r.status_code}]: {r.text}")
    return False, None

def get_doctype_definition(doctype):
    return erp_get(f"/api/resource/DocType/{q(doctype)}").get("data", {}) or {}

def find_bid_result_note_field(field_label_contains):
    meta   = get_doctype_definition("Bid Result")
    fields = meta.get("fields") or []
    kw     = field_label_contains.strip().lower()
    for f in fields:
        if kw in (f.get("label") or "").strip().lower():
            return f.get("fieldname")
    return None

def company_exists(name):
    exists, _ = erp_try_get_resource(COMPANY_DOCTYPE, name)
    return bool(exists)

def build_min_create_payload_for_competitor(name, context=None):
    context    = context or {}
    meta       = get_doctype_definition(COMPETITOR_DOCTYPE)
    fields     = meta.get("fields") or []
    fieldnames = {f.get("fieldname") for f in fields}
    payload    = {"doctype": COMPETITOR_DOCTYPE}

    autoname = (meta.get("autoname") or "").strip()
    if autoname.startswith("field:"):
        fn = autoname.split("field:", 1)[1].strip()
        if fn:
            payload[fn] = name

    for fn in ["competitor_name", "company_name", "party_name", "participant_name"]:
        if fn in fieldnames:
            payload.setdefault(fn, name)

    for f in fields:
        if not f.get("reqd"):
            continue
        fn = f.get("fieldname")
        if not fn or fn in payload:
            continue
        ft   = f.get("fieldtype")
        opts = f.get("options") or ""
        if fn in context and context[fn]:
            payload[fn] = context[fn]
        elif ft in ("Data", "Small Text", "Text", "Text Editor"):
            payload[fn] = name
        elif ft in ("Float", "Currency", "Int"):
            payload[fn] = 0
        elif ft == "Check":
            payload[fn] = 0
        elif ft == "Select":
            choices = [x.strip() for x in opts.splitlines() if x.strip()]
            if choices:
                payload[fn] = choices[0]
    return payload

def ensure_competitor_exists(name, context=None):
    name = (name or "").strip()
    if not name:
        return
    exists, _ = erp_try_get_resource(COMPETITOR_DOCTYPE, name)
    if exists:
        return
    erp_post(f"/api/resource/{q(COMPETITOR_DOCTYPE)}",
             build_min_create_payload_for_competitor(name, context=context))
    print(f"[OK] Auto-created Competitor: {name}")

def ensure_sales_stage_exists(stage_name):
    exists, _ = erp_try_get_resource("Sales Stage", stage_name)
    if not exists:
        raise SystemExit(
            f"❌ Sales Stage '{stage_name}' does not exist in ERP.\n"
            f"Create it: Settings → CRM → Sales Stage"
        )

# ==========================================================
# NORMALIZATION
# ==========================================================
def _norm_match(s):
    return re.sub(r"[^A-Za-z0-9]+", "", (s or "")).upper()

OUR_MATCH_KEYS = {_norm_match(x) for x in OUR_COMPANY_CANONICAL}

def normalize_vendor_to_company_name(vendor):
    v = (vendor or "").strip()
    return COMPANY_NAME_ALIASES.get(v.upper(), v)

def is_our_company(vendor):
    return _norm_match(normalize_vendor_to_company_name(vendor)) in OUR_MATCH_KEYS

# ==========================================================
# MARGIN
# ==========================================================
def compute_margin_fields(sorted_bids_kwd, our_bid_kwd):
    if not sorted_bids_kwd:
        return 0.0, 0.0
    l1 = float(sorted_bids_kwd[0]["value_kwd"])
    l2 = float(sorted_bids_kwd[1]["value_kwd"]) if len(sorted_bids_kwd) > 1 else None
    margin = max(0.0, l2 - l1) if l2 is not None else (
             max(0.0, float(our_bid_kwd) - l1) if our_bid_kwd is not None else 0.0)
    return float(margin), float((margin / l1 * 100.0) if l1 > 0 else 0.0)

def force_bid_result_fields(bid_result_name, fields):
    if not bid_result_name:
        return
    try:
        erp_put(f"/api/resource/Bid Result/{q(bid_result_name)}", fields)
        print(f"[OK] Forced Bid Result fields ({bid_result_name})")
    except Exception as e:
        print(f"[WARN] Could not force Bid Result fields: {e}")

# ==========================================================
# BUSINESS LOGIC
# ==========================================================
def normalize_rank(n):
    return f"L{n}" if 1 <= n <= 5 else "L6+"

def count_bid_results_for_opportunity(opportunity_name):
    rows = erp_get("/api/resource/Bid Result", params={
        "fields":            json.dumps(["name", "result_date", "creation"]),
        "filters":           json.dumps([["Bid Result", "opportunity", "=", opportunity_name]]),
        "order_by":          "creation asc",
        "limit_page_length": 50,
    }).get("data") or []
    names = [r["name"] for r in rows]
    return len(names), names


def fetch_kuwait_opportunities(limit_page_length=250):
    filters = [["Opportunity", "customer_name", "like", "%Kuwait%"]]
    fields  = [
        "name", "customer_name", "territory", "transaction_date",
        "company", "currency", RFX_FIELDNAME, SALES_STATUS_FIELDNAME,
        "_assign", EXPECTED_CLOSING_FIELDNAME,
    ]
    all_rows, start = [], 0
    while True:
        batch = erp_get("/api/resource/Opportunity", params={
            "fields": json.dumps(fields), "filters": json.dumps(filters),
            "limit_page_length": limit_page_length, "limit_start": start,
            "order_by": "modified desc"
        }).get("data", []) or []
        if not batch:
            break
        all_rows.extend(batch)
        start += len(batch)
        if len(batch) < limit_page_length:
            break
    return all_rows

def update_opportunity_sales_stage(opportunity_name, stage_value):
    return erp_put(f"/api/resource/Opportunity/{q(opportunity_name)}",
                   {SALES_STATUS_FIELDNAME: stage_value})

def map_sales_stage(scraped):
    if not scraped.get("our_present"):
        return SALES_STAGE_NP
    rank = scraped.get("our_rank", "L6+")
    if rank == "L1":
        return SALES_STAGE_BID_WON
    if rank in ("L2", "L3", "L4", "L5"):
        return SALES_STAGE_COMPLETED
    return SALES_STAGE_BID_LOST

def fetch_opportunity_items(opportunity_name):
    try:
        doc      = erp_get(f"/api/resource/Opportunity/{q(opportunity_name)}")
        data     = doc.get("data") or {}
        raw_items = data.get("items") or data.get("opportunity_item") or []
        if not raw_items:
            print(f"  [INFO] No items found on Opportunity {opportunity_name}")
            return []
        result = []
        for row in raw_items:
            item_code  = (row.get("item_code")  or "").strip()
            item_name  = (row.get("item_name")  or "").strip()
            qty        = float(row.get("qty") or row.get("quantity") or 1)
            item_group = (row.get("item_group") or "").strip()
            if not item_code:
                continue
            result.append({"item_code": item_code, "item_name": item_name,
                            "qty": qty, "item_group": item_group})
        print(f"  [INFO] Fetched {len(result)} item(s) from Opportunity {opportunity_name}")
        return result
    except Exception as e:
        print(f"  [WARN] Could not fetch items from Opportunity {opportunity_name}: {e}")
        return []

# ==========================================================
# BUILD HTML COMMENT
# ==========================================================
def build_comment_text(rfq_number, scraped, round_number=1):
    result_type   = scraped.get("result_type", "")
    our_present   = scraped.get("our_present", False)
    participants  = scraped.get("participants", [])
    winner_name   = scraped.get("winner_name", "")
    winner_amount = scraped.get("winner_amount_kwd", 0.0)
    total         = scraped.get("total_participants", len(participants))
    is_split      = scraped.get("is_split", False)
    today_str     = dt.date.today().strftime("%d %b %Y")

    if not our_present:
        badge_color, badge_icon, badge_label = "#6c757d", "⚪", "Not Participated"
    elif result_type == "Won":
        badge_color, badge_icon, badge_label = "#1a7a3f", "🏆", "Bid Won"
    else:
        badge_color, badge_icon, badge_label = "#c0392b", "❌", "Bid Lost"

    round_label = f"Round {round_number}" if round_number > 1 else ""
    round_html  = (
        f'<span style="background:#1a4fa0;color:#fff;font-size:11px;'
        f'padding:2px 8px;border-radius:3px;margin-left:10px;">🔁 {round_label}</span>'
        if round_label else ""
    )

    html = f"""
<div style="font-family:Arial,sans-serif;font-size:13px;color:#1a1a1a;max-width:680px;">
  <table width="100%" cellpadding="0" cellspacing="0"
         style="border-bottom:2px solid {badge_color};margin-bottom:12px;">
    <tr>
      <td style="padding:6px 0;">
        <span style="font-size:15px;font-weight:bold;color:{badge_color};">
          {badge_icon} &nbsp; RFQ {rfq_number} &mdash; {badge_label}
        </span>
        {round_html}
      </td>
      <td align="right" style="font-size:11px;color:#888;padding:6px 0;">{today_str}</td>
    </tr>
  </table>
"""
    if not our_present:
        html += '<p style="color:#555;font-style:italic;">&#9432;&nbsp; We did not participate.</p>\n'
    if is_split:
        html += (f'<div style="background:#fff8e1;border-left:4px solid #f0a500;'
                 f'padding:8px 12px;margin-bottom:12px;">'
                 f'<b>⚠️ SPLIT AWARD DETECTED</b><br>'
                 f'<span style="font-size:12px;">{SPLIT_MSG}</span></div>\n')

    top5      = [p for p in participants if p.get("rank") in ("L1","L2","L3","L4","L5")]
    remaining = total - len(top5)
    RANK_COLORS = {"L1":"#1a7a3f","L2":"#1a4fa0","L3":"#555555","L4":"#555555","L5":"#555555"}

    html += ('<table width="100%" cellpadding="0" cellspacing="0" '
             'style="border-collapse:collapse;margin-bottom:14px;">\n')
    html += ('<tr style="background:#f0f0f0;">'
             '<td width="60" style="padding:7px 10px;font-weight:bold;border-bottom:1px solid #ccc;">Rank</td>'
             '<td style="padding:7px 10px;font-weight:bold;border-bottom:1px solid #ccc;">Company</td>'
             '<td width="150" style="padding:7px 10px;font-weight:bold;border-bottom:1px solid #ccc;'
             'text-align:right;">Amount (KWD)</td></tr>\n')

    for i, p in enumerate(top5):
        rank   = p.get("rank","")
        name   = p.get("name","")
        amount = p.get("amount_kwd", 0.0)
        ours   = is_our_company(name)
        row_bg = "#f9fff9" if ours else ("#ffffff" if i % 2 == 0 else "#fafafa")
        badge  = ('&nbsp;<span style="background:#1a7a3f;color:#fff;font-size:10px;'
                  'padding:1px 5px;border-radius:3px;">⭐ OUR COMPANY</span>' if ours else "")
        html += (f'<tr style="background:{row_bg};">'
                 f'<td style="padding:8px 10px;border-bottom:1px solid #eee;">'
                 f'<b style="color:{RANK_COLORS.get(rank,"#333")};">{rank}</b></td>'
                 f'<td style="padding:8px 10px;border-bottom:1px solid #eee;">{name}{badge}</td>'
                 f'<td style="padding:8px 10px;border-bottom:1px solid #eee;text-align:right;'
                 f'font-weight:bold;">{amount:,.3f}</td></tr>\n')

    if remaining > 0:
        html += (f'<tr><td colspan="3" style="padding:6px 10px;font-size:11px;color:#888;'
                 f'border-top:1px solid #ddd;font-style:italic;">'
                 f'+ {remaining} more participant{"s" if remaining > 1 else ""} not shown</td></tr>\n')
    html += "</table>\n"

    margin_val = scraped.get("winning_margin_kwd", 0)
    margin_pct = scraped.get("winning_margin_pct", 0)
    margin_row = ""
    if margin_val and margin_val > 0:
        margin_row = f"""
  <tr>
    <td style="padding:6px 10px;color:#333;font-size:13px;border-bottom:1px solid #eee;width:40%;">
      <b>Winning Margin</b>
    </td>
    <td style="padding:6px 10px;color:#1a1a1a;font-size:13px;font-weight:bold;border-bottom:1px solid #eee;">
      {margin_val:,.3f} KWD
      <span style="color:#555;font-size:12px;font-weight:normal;">({margin_pct:.2f}%)</span>
    </td>
  </tr>"""

    html += f"""
<table width="100%" cellpadding="0" cellspacing="0"
       style="border-collapse:collapse;border-top:2px solid #ddd;margin-top:4px;background:#f9f9f9;">
  <tr>
    <td style="padding:6px 10px;color:#333;font-size:13px;border-bottom:1px solid #eee;width:40%;">
      <b>Winner</b>
    </td>
    <td style="padding:6px 10px;font-size:13px;font-weight:bold;border-bottom:1px solid #eee;color:{badge_color};">
      {winner_name} &mdash; {winner_amount:,.3f} KWD
    </td>
  </tr>
  {margin_row}
  <tr>
    <td style="padding:6px 10px;color:#333;font-size:13px;width:40%;">
      <b>Total Participants</b>
    </td>
    <td style="padding:6px 10px;color:#1a1a1a;font-size:13px;font-weight:bold;">
      {total}
    </td>
  </tr>
</table>
</div>
"""
    return html

# ==========================================================
# POST COMMENT TO ERP
# ==========================================================
def post_comment_to_erp(doctype, docname, comment_html):
    try:
        result       = erp_post("/api/resource/Comment", {
            "doctype": "Comment", "comment_type": "Comment",
            "reference_doctype": doctype, "reference_name": docname,
            "content": comment_html,
        })
        comment_name = (result.get("data") or {}).get("name", "")
        print(f"[OK] Comment posted on {doctype} '{docname}' → {comment_name}")
        return comment_name
    except Exception as e:
        print(f"[WARN] Comment failed on {doctype} '{docname}': {e}")
        return None

# ==========================================================
# UPLOAD SCREENSHOT
# ==========================================================
def upload_screenshot_to_opportunity(png_bytes, opportunity_name, rfq_number):
    if not png_bytes:
        return None
    try:
        filename       = f"bid_result_{rfq_number}_{dt.date.today().isoformat()}.png"
        upload_headers = {
            "Authorization": f"token {ERPNEXT_API_KEY}:{ERPNEXT_API_SECRET}",
            "Accept":        "application/json",
        }
        r = requests.post(
            f"{ERPNEXT_URL}/api/method/upload_file",
            headers=upload_headers,
            files={"file": (filename, png_bytes, "image/png")},
            data={"is_private": "1", "doctype": "Opportunity",
                  "docname": opportunity_name, "folder": "Home/Attachments"},
            timeout=TIMEOUT
        )
        if r.status_code == 200:
            file_url = (r.json().get("message") or {}).get("file_url", "")
            print(f"[OK] Screenshot uploaded → {file_url}")
            return file_url
        print(f"[WARN] Upload failed [{r.status_code}]: {r.text[:200]}")
    except Exception as e:
        print(f"[WARN] Screenshot upload error: {e}")
    return None

# ==========================================================
# SELENIUM — start driver
# ==========================================================
def start_driver():
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_experimental_option("prefs", {
        "profile.managed_default_content_settings.images":      2,
        "profile.managed_default_content_settings.stylesheets": 2,
        "profile.managed_default_content_settings.fonts":       2,
    })
    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=chrome_options
    )


# ==========================================================
# SMART LOGIN  (detects OTP page vs direct portal access)
# ==========================================================
def ktender_login_and_open_bid_results(driver):
    wait = WebDriverWait(driver, 30)

    driver.get(MAIN_URL)
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    supplier = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(.,'Supplier Access')]")))
    driver.execute_script("arguments[0].click();", supplier)
    time.sleep(1.5)
    driver.switch_to.window(driver.window_handles[-1])

    def _has(by, val, timeout=6):
        try:
            WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, val)))
            return True
        except TimeoutException:
            return False

    if _has(By.XPATH, "//a[contains(.,'Supplier Reports')]", timeout=6):
        print("[OK] Session still active — skipped login entirely.")

    elif _has(By.ID, "username", timeout=6):
        username_field = driver.find_element(By.ID, "username")
        username_field.clear()
        username_field.send_keys(EMAIL_USER)
        print(f"[OK] Entered username: {EMAIL_USER}")

        password_field = driver.find_element(By.ID, "password")
        password_field.clear()
        password_field.send_keys(KT_PASS)
        print("[OK] Entered password")

        login_clicked_at = time.time()

        login_btn = wait.until(EC.element_to_be_clickable((By.ID, "kc-login")))
        driver.execute_script("arguments[0].click();", login_btn)
        print("[OK] Clicked Login — detecting next page…")

        otp_detected    = False
        portal_detected = False

        deadline = time.time() + 15
        while time.time() < deadline:
            page_src = driver.page_source
            if 'id="code"' in page_src or 'name="otp"' in page_src:
                otp_detected = True
                break
            if "Supplier Reports" in page_src:
                portal_detected = True
                break
            time.sleep(0.5)

        if otp_detected:
            print("[OK] OTP page detected — fetching code from mailbox…")
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "code")))
            except TimeoutException:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "otp")))

            otp_code = fetch_otp_from_email(
                sent_after_timestamp=login_clicked_at,
                max_wait=90,
                poll_interval=5
            )

            try:
                otp_field = driver.find_element(By.ID, "code")
            except NoSuchElementException:
                otp_field = driver.find_element(By.NAME, "otp")
            otp_field.clear()
            otp_field.send_keys(otp_code)
            print(f"[OK] Entered OTP: {otp_code}")

            try:
                submit_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[@id='kc-login' and @data-action-type='login']")
                ))
            except TimeoutException:
                submit_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[@type='submit' and (@value='Submit' or @id='kc-login')]")
                ))
            driver.execute_script("arguments[0].click();", submit_btn)
            print("[OK] Submitted OTP — waiting for portal…")

            wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            time.sleep(2)

        elif portal_detected:
            print("[OK] Logged in directly — no OTP required.")
            time.sleep(1)

        else:
            driver.save_screenshot("login_debug.png")
            raise RuntimeError(
                "❌ Could not detect OTP page or portal after login. "
                "Screenshot saved as login_debug.png"
            )

    else:
        driver.save_screenshot("login_debug.png")
        raise RuntimeError(
            "❌ Could not find login form or portal. "
            "Screenshot saved as login_debug.png"
        )

    print("✅ Login complete!")

    try:
        popup = driver.find_element(By.XPATH, "//input[@class='btn-close' and @value='x']")
        driver.execute_script("arguments[0].click();", popup)
    except Exception:
        pass

    reports = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//a[contains(.,'Supplier Reports')]")
    ))
    driver.execute_script("arguments[0].click();", reports)
    time.sleep(0.5)

    bid_results = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//a[contains(.,'Bid Results - Past 3 Months')]")
    ))
    driver.execute_script("arguments[0].click();", bid_results)

    time.sleep(1.5)
    driver.switch_to.window(driver.window_handles[-1])
    main_window = driver.current_window_handle

    found = False
    try:
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.ID, "searchby"))
        )
        found = True
        print("[OK] Bid Results page ready (no iframe).")
    except TimeoutException:
        pass

    if not found:
        for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(iframe)
                WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.ID, "searchby"))
                )
                found = True
                print("[OK] Bid Results page ready (inside iframe).")
                break
            except (TimeoutException, Exception):
                continue

    if not found:
        raise RuntimeError("❌ Could not find Bid Results search form after login.")

    return main_window


# ==========================================================
# _ensure_search_page
# ==========================================================
def _ensure_search_page(driver, main_window, label=""):
    for handle in list(driver.window_handles):
        if handle != main_window:
            try:
                driver.switch_to.window(handle)
                driver.close()
                print(f"[WARN] Closed stray tab: {handle}")
            except Exception:
                pass
    driver.switch_to.window(main_window)

    try:
        driver.find_element(By.ID, "searchby")
        return True
    except NoSuchElementException:
        pass

    try:
        driver.switch_to.default_content()
        driver.find_element(By.ID, "searchby")
        return True
    except NoSuchElementException:
        pass

    driver.switch_to.default_content()
    for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(iframe)
            WebDriverWait(driver, 4).until(
                EC.presence_of_element_located((By.ID, "searchby"))
            )
            return True
        except (TimeoutException, Exception):
            continue

    print(f"[WARN] Search page lost{' for ' + label if label else ''} — re-navigating to Bid Results…")
    driver.switch_to.default_content()
    try:
        w = WebDriverWait(driver, 15)
        try:
            bid_link = w.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(.,'Bid Results - Past 3 Months')]")
            ))
        except TimeoutException:
            reports = w.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(.,'Supplier Reports')]")
            ))
            driver.execute_script("arguments[0].click();", reports)
            time.sleep(0.5)
            bid_link = w.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(.,'Bid Results - Past 3 Months')]")
            ))

        driver.execute_script("arguments[0].click();", bid_link)
        time.sleep(1.5)
        driver.switch_to.window(driver.window_handles[-1])

        try:
            WebDriverWait(driver, 8).until(
                EC.presence_of_element_located((By.ID, "searchby"))
            )
            return True
        except TimeoutException:
            for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(iframe)
                    WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.ID, "searchby"))
                    )
                    return True
                except (TimeoutException, Exception):
                    continue
    except Exception as e:
        print(f"[ERROR] Re-navigation failed: {e}")

    return False


# ==========================================================
# SCRAPE ONE RFQ
# ==========================================================
def scrape_one_rfq(driver, main_window, rfq):
    wait       = WebDriverWait(driver, 20)
    short_wait = WebDriverWait(driver, 5)

    if not _ensure_search_page(driver, main_window, label=f"RFQ {rfq}"):
        print(f"[WARN] Could not reach search page for RFQ {rfq} — skipping.")
        return None, None

    Select(driver.find_element(By.ID, "searchby")).select_by_visible_text("RFQ Title")
    search_input = driver.find_element(By.ID, "searchvalue")
    search_input.clear()
    search_input.send_keys(rfq)

    try:
        search_btn = driver.find_element(By.XPATH, "//button[contains(.,'Search')]")
        driver.execute_script("arguments[0].click();", search_btn)
    except NoSuchElementException:
        search_input.send_keys("\n")

    time.sleep(2)

    try:
        link = short_wait.until(EC.element_to_be_clickable((
            By.XPATH,
            f"//td/a[contains(@href,'viewSuppliers.jsp') and contains(text(),'{rfq}')]"
        )))
    except TimeoutException:
        return None, None

    driver.execute_script("arguments[0].setAttribute('target','_blank');", link)
    driver.execute_script("arguments[0].click();", link)

    wait.until(lambda d: len(d.window_handles) > 1)
    driver.switch_to.window(driver.window_handles[-1])
    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
    time.sleep(1.5)

    rows = driver.find_elements(By.XPATH, "//table//tr[td]")
    bids = []
    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        if len(cells) < 4:
            continue
        vendor   = normalize_vendor_to_company_name(cells[0].text.strip())
        currency = cells[1].text.strip().upper()
        try:
            value = float(cells[2].text.strip().replace(",", ""))
            items = int(cells[3].text.strip())
        except Exception:
            continue
        rate = CURRENCY_TO_KWD.get(currency)
        if rate is None:
            continue
        bids.append({"vendor": vendor, "value_kwd": value * rate, "items": items})

    screenshot_png = None
    try:
        screenshot_png = driver.get_screenshot_as_png()
        print(f"[OK] Screenshot captured for RFQ {rfq}")
    except Exception as e:
        print(f"[WARN] Screenshot failed: {e}")

    try:
        driver.close()
    except Exception:
        pass
    driver.switch_to.window(main_window)

    if not bids:
        return None, None

    is_split    = len({b["items"] for b in bids}) > 1
    sorted_bids = sorted(bids, key=lambda x: x["value_kwd"])
    our_bids    = [b for b in sorted_bids if is_our_company(b["vendor"])]
    our_lowest  = min(our_bids, key=lambda x: x["value_kwd"]) if our_bids else None
    our_present = our_lowest is not None

    participants = [
        {"rank": normalize_rank(i), "name": b["vendor"],
         "amount_kwd": b["value_kwd"], "is_winner": int(i == 1), "items": b["items"]}
        for i, b in enumerate(sorted_bids, start=1)
    ]

    winner        = sorted_bids[0]["vendor"]
    winner_amount = sorted_bids[0]["value_kwd"]

    if not our_present:
        result_type, our_rank, our_amount = "No Decision", "L6+", 0.0
    else:
        our_pos = next(
            (i for i, b in enumerate(sorted_bids, start=1)
             if b["vendor"] == our_lowest["vendor"] and b["value_kwd"] == our_lowest["value_kwd"]), 99
        )
        our_rank    = normalize_rank(our_pos)
        our_amount  = our_lowest["value_kwd"]
        result_type = "Won" if our_rank == "L1" else "Lost"

    margin_val, margin_pct = compute_margin_fields(sorted_bids, our_amount if our_present else None)

    return {
        "result_type":        result_type,
        "our_rank":           our_rank,
        "our_amount_kwd":     our_amount,
        "our_present":        our_present,
        "winner_name":        winner,
        "winner_amount_kwd":  winner_amount,
        "participants":       participants,
        "total_participants": len(participants),
        "is_split":           is_split,
        "winning_margin_kwd": margin_val,
        "winning_margin_pct": margin_pct,
    }, screenshot_png


# ==========================================================
# CREATE BID RESULT IN ERP
# ==========================================================
def create_bid_result_in_erp(opp, scraped, rfq_number, opp_items=None, round_number=1):
    today = dt.date.today().isoformat()
    ctx   = {"company": opp.get("company"), "territory": opp.get("territory")}

    loss_analysis_field = find_bid_result_note_field("loss analysis")
    win_factors_field   = find_bid_result_note_field("win factors")
    winner_name         = scraped["winner_name"]

    if company_exists(winner_name):
        winner_type_value = COMPANY_TYPE_VALUE
    else:
        ensure_competitor_exists(winner_name, context=ctx)
        winner_type_value = COMPETITOR_TYPE_VALUE

    for p in scraped["participants"]:
        if not company_exists(p["name"]):
            ensure_competitor_exists(p["name"], context=ctx)

    rfq_line           = f"RFQ: {rfq_number}"
    loss_analysis_text = rfq_line if scraped["result_type"] != "Won" else ""
    win_factors_text   = rfq_line if scraped["result_type"] == "Won"  else ""

    if scraped.get("is_split"):
        if scraped["result_type"] == "Won":
            win_factors_text   = (win_factors_text + "\n\n" + SPLIT_MSG).strip()
        else:
            loss_analysis_text = (loss_analysis_text + "\n\n" + SPLIT_MSG).strip()

    payload = {
        "doctype":            "Bid Result",
        "opportunity":        opp["name"],
        "result_date":        today,
        "bid_date":           opp.get("transaction_date") or today,
        "territory":          opp.get("territory"),
        "company":            opp.get("company"),
        "result_type":        scraped["result_type"],
        "our_rank":           scraped["our_rank"],
        "our_quoted_amount":  round(float(scraped["our_amount_kwd"]), 3),
        "currency":           BID_RESULT_CURRENCY,
        "winner_type":        winner_type_value,
        "winner":             winner_name,
        "winning_amount":     round(float(scraped["winner_amount_kwd"]), 3),
        "total_participants": int(scraped["total_participants"]),
        "round":              round_number,
        "bid_participants":   [],
        "custom_bid_items":   [],
    }

    if loss_analysis_field and loss_analysis_text:
        payload[loss_analysis_field] = loss_analysis_text
    if win_factors_field and win_factors_text:
        payload[win_factors_field] = win_factors_text

    for p in scraped["participants"]:
        nm    = p["name"]
        ptype = COMPANY_TYPE_VALUE if company_exists(nm) else COMPETITOR_TYPE_VALUE
        payload["bid_participants"].append({
            "doctype":          "Bid Participant",
            "participant_type": ptype,
            "participant":      nm,
            "participant_name": nm,
            "rank":             p["rank"],
            "quoted_amount":    round(float(p["amount_kwd"]), 3),
            "is_winner":        int(p["is_winner"]),
        })

    if opp_items:
        for item in opp_items:
            payload["custom_bid_items"].append({
                "doctype":    "Bid Item",
                "item_code":  item.get("item_code", ""),
                "item_name":  item.get("item_name", ""),
                "qty":        item.get("qty", 1),
                "item_group": item.get("item_group", ""),
            })
        print(f"  [INFO] Added {len(opp_items)} item(s) to Bid Result payload (Round {round_number})")
    else:
        print(f"  [INFO] No items to add to Bid Result (Round {round_number})")

    created  = erp_post("/api/resource/Bid Result", payload)
    bid_name = (created.get("data") or {}).get("name")

    if bid_name:
        codes  = " ".join(i.get("item_code",  "") for i in (opp_items or []) if i.get("item_code"))
        groups = " ".join(i.get("item_group", "") for i in (opp_items or []) if i.get("item_group"))

        force_fields = {
            "currency":                  BID_RESULT_CURRENCY,
            "winning_margin":            round(float(scraped.get("winning_margin_kwd", 0.0)), 3),
            "winning_margin_percentage": round(float(scraped.get("winning_margin_pct", 0.0)), 3),
            "round":                     round_number,
        }
        if codes:
            force_fields["custom_search_item_code"]  = codes
        if groups:
            force_fields["custom_search_item_group"] = groups

        force_bid_result_fields(bid_name, force_fields)

    return created


# ==========================================================
# MAIN
# ==========================================================
def main():
    prevent_sleep_windows()

    print("🔌 Connecting to ERPNext…")
    opps = fetch_kuwait_opportunities(limit_page_length=OPP_PAGE_SIZE)
    print(f"[OK] Fetched {len(opps)} Kuwait opportunities from ERP.")

    candidates = []
    for opp in opps:
        rfq = (opp.get(RFX_FIELDNAME) or "").strip()
        if not rfq:
            continue
        stage = (opp.get(SALES_STATUS_FIELDNAME) or "").strip()
        if not INCLUDE_ALL_SALES_STAGES and stage in SKIP_SALES_STAGES:
            continue
        candidates.append(opp)

    print(f"[OK] {len(candidates)} opportunities to process.")
    if not candidates:
        print("Nothing to do.")
        return

    driver     = start_driver()
    main_window = ktender_login_and_open_bid_results(driver)

    processed = skipped = errors = 0

    for opp in candidates:
        rfq_number = (opp.get(RFX_FIELDNAME) or "").strip()
        opp_name   = opp["name"]
        print(f"\n{'='*60}")
        print(f"Processing: {rfq_number}  |  Opportunity: {opp_name}")
        print(f"{'='*60}")

        try:
            scraped, screenshot_png = scrape_one_rfq(driver, main_window, rfq_number)

            if scraped is None:
                print(f"[SKIP] No bid result found on ktendering for RFQ {rfq_number}")
                skipped += 1
                continue

            round_count, _ = count_bid_results_for_opportunity(opp_name)
            round_number   = round_count + 1

            opp_items = fetch_opportunity_items(opp_name)

            created  = create_bid_result_in_erp(
                opp, scraped, rfq_number,
                opp_items=opp_items, round_number=round_number
            )
            bid_name = (created.get("data") or {}).get("name", "")
            print(f"[OK] Bid Result created: {bid_name}  (Round {round_number})")

            new_stage = map_sales_stage(scraped)
            update_opportunity_sales_stage(opp_name, new_stage)
            print(f"[OK] Sales stage → {new_stage}")

            if screenshot_png:
                upload_screenshot_to_opportunity(screenshot_png, opp_name, rfq_number)

            comment_html = build_comment_text(rfq_number, scraped, round_number=round_number)
            post_comment_to_erp("Opportunity", opp_name, comment_html)

            processed += 1

        except Exception as e:
            import traceback
            print(f"[ERROR] {rfq_number}: {e}")
            traceback.print_exc()
            errors += 1

    driver.quit()
    print(f"\n✅ Done — processed: {processed}, skipped: {skipped}, errors: {errors}")


if __name__ == "__main__":
    main()
