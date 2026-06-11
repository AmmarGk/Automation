# -*- coding: utf-8 -*-
import os
import sys
import io
import time
import re
import shutil
import mimetypes
import random
import pickle
import json
import imaplib
import email as email_lib
import email.utils as eutils

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from dotenv import load_dotenv

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# --- GOOGLE DRIVE IMPORTS ---
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError
import zipfile

# Load environment variables from .env file
load_dotenv()

# ==========================================================
# 1. SETUP AND CONFIGURATION
# ==========================================================

# Ktendering credentials
LOGIN    = os.getenv("KTENDER_USERNAME")
PASSWORD = os.getenv("KTENDER_PASSWORD")

# Google Drive config
PARENT_FOLDER_ID    = os.getenv("PARENT_FOLDER_ID")
SERVICE_ACCOUNT_FILE = "service_account.json"
SCOPES               = ["https://www.googleapis.com/auth/drive"]

# --- IMAP credentials (Hostinger) — OTP fetch only ---
IMAP_SERVER    = os.getenv("IMAP_SERVER",    "imap.hostinger.com")
IMAP_PORT      = int(os.getenv("IMAP_PORT",  "993"))
EMAIL_USER     = os.getenv("EMAIL_USER",     "info@gkepoxy.com")
EMAIL_APP_PASS = os.getenv("EMAIL_APP_PASS", "")

# --- Validate required credentials ---
if not LOGIN or not PASSWORD:
    raise ValueError("Error: KTENDER_USERNAME or KTENDER_PASSWORD missing in .env")
if not PARENT_FOLDER_ID or PARENT_FOLDER_ID.startswith("YOUR_"):
    raise ValueError("Error: PARENT_FOLDER_ID is not set correctly in .env")
if not os.path.exists(SERVICE_ACCOUNT_FILE):
    raise FileNotFoundError(f"Service account file not found: {SERVICE_ACCOUNT_FILE}")

# URLs
LOGIN_URL = "https://ktendering.com.kw/esop/guest/login.do?j1pidm=true&internal=false&userAct=changeLangIndex&language=en_GB&_ncp=1758864349514.1369049-1"

# Local paths
DOWNLOAD_DIR     = os.path.join(os.getcwd(), "Downloaded_RFQs")
LOG_FILE         = os.path.join(os.getcwd(), "RFQ_Log.xlsx")
SESSION_FILE     = os.path.join(os.getcwd(), "ktender_session.pkl")
TITLES_JSON_FILE = os.path.join(os.getcwd(), "RFQ_Titles_History.json")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)


# ==========================================================
# 2. EMAIL / OTP HELPERS
# ==========================================================

def fetch_otp_from_email(sent_after_timestamp: float, max_wait: int = 90, poll_interval: int = 5) -> str:
    """
    Polls the Hostinger IMAP inbox for an OTP email from jaggaer.com
    that arrived after `sent_after_timestamp`. Returns the OTP string.

    OTP locator strategy (in order):
      1. "verification code is: XXXXXX"
      2. "code: XXXXXX" / "otp: XXXXXX" / "token: XXXXXX"
      3. Any 4-12 char alphanumeric token containing both letters and digits
    """
    ts_str = time.strftime('%H:%M:%S', time.localtime(sent_after_timestamp))
    print(f"\n🔐 Polling {IMAP_SERVER}:{IMAP_PORT} for OTP (sent after {ts_str}, timeout {max_wait}s)…")
    deadline = time.time() + max_wait

    while time.time() < deadline:
        try:
            mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
            mail.login(EMAIL_USER, EMAIL_APP_PASS)
            mail.select("INBOX")

            # Try targeted search first, fall back to all unseen
            _, data = mail.search(None, '(UNSEEN FROM "noreply@jaggaer.com")')
            mail_ids = data[0].split()
            if not mail_ids:
                _, data = mail.search(None, "UNSEEN")
                mail_ids = data[0].split()

            for mid in reversed(mail_ids):
                _, msg_data = mail.fetch(mid, "(RFC822)")
                msg = email_lib.message_from_bytes(msg_data[0][1])

                sender = msg.get("From", "")
                # Only process emails from jaggaer (ktendering OTP sender)
                if "jaggaer" not in sender.lower():
                    continue

                # Skip emails that arrived before the login click
                date_str = msg.get("Date", "")
                try:
                    email_ts = eutils.parsedate_to_datetime(date_str).timestamp()
                except Exception:
                    email_ts = 0.0
                if email_ts < (sent_after_timestamp - 30):
                    print(f"   ⏭  Skipping old OTP email ({time.strftime('%H:%M:%S', time.localtime(email_ts))})")
                    continue

                # Extract body
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
                print(f"   📨 From: {sender[:60]}")
                print(f"   Subject: {subject[:80]}")

                # --- OTP extraction patterns ---
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
                    # Fallback: any mixed alphanumeric token in body
                    for candidate in re.findall(r"\b([A-Za-z0-9]{4,12})\b", body):
                        if re.search(r"[A-Za-z]", candidate) and re.search(r"[0-9]", candidate):
                            class _M:
                                def __init__(self, v): self._v = v
                                def group(self, _): return self._v
                            otp_match = _M(candidate)
                            break

                if otp_match:
                    otp = otp_match.group(1)
                    print(f"✅ OTP found: {otp}")
                    mail.store(mid, "+FLAGS", "\\Seen")
                    mail.logout()
                    return otp
                else:
                    print(f"   ⚠️  Email matched but OTP not extracted. Body snippet: {body[:120]}")

            mail.logout()

        except Exception as e:
            print(f"   IMAP error: {e}")

        remaining = int(deadline - time.time())
        if remaining > 0:
            print(f"   Retrying in {poll_interval}s… ({remaining}s left)")
            time.sleep(poll_interval)

    raise TimeoutError("❌ OTP email from JAGGAER did not arrive within the timeout period.")


# ==========================================================
# 3. HELPER FUNCTIONS
# ==========================================================

def extract_rfq_title_from_detail_page(wait):
    """
    Extracts the RFQ title from <span class="mainTitle"> on the detail page.
    Returns dict with full_title, event_id, rfq_number, description.
    """
    try:
        main_title_element = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "span.mainTitle"))
        )
        full_title = main_title_element.text.strip()
        print(f"\n📋 Extracted Title: {full_title}")

        result = {
            'full_title': full_title,
            'event_id':   'N/A',
            'rfq_number': 'N/A',
            'description':'N/A'
        }

        event_match = re.search(r'event_(\d+)', full_title)
        if event_match:
            result['event_id'] = f"event_{event_match.group(1)}"

        rfq_match = re.search(r'RFQ#(\d+)', full_title)
        if rfq_match:
            result['rfq_number'] = f"RFQ#{rfq_match.group(1)}"

        desc_match = re.search(r'RFQ#\d+_(.+)$', full_title)
        if desc_match:
            result['description'] = desc_match.group(1).strip()

        print(f"   📌 Event ID: {result['event_id']}")
        print(f"   📌 RFQ Number: {result['rfq_number']}")
        print(f"   📌 Description: {result['description']}")
        return result

    except TimeoutException:
        print("⚠️  Could not find mainTitle element on detail page")
    except Exception as e:
        print(f"⚠️  Error extracting mainTitle: {e}")

    return {'full_title': 'N/A', 'event_id': 'N/A', 'rfq_number': 'N/A', 'description': 'N/A'}


def wait_for_download_completion(download_dir, files_before, max_wait_seconds=120):
    """
    Waits for a browser download to complete using pure filesystem polling.
    Monitors both download_dir AND ~/Downloads as a fallback.

    Returns: (success: bool, new_files: set, failure_reason: str)
             failure_reason: 'virus_scan' | 'timeout' | ''
    """
    print(f"\n⬇️  Waiting for download to complete...")

    sys_dl_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    if os.path.isdir(sys_dl_dir) and os.path.abspath(sys_dl_dir) != os.path.abspath(download_dir):
        sys_dl_before = set(os.listdir(sys_dl_dir))
        print(f"   (Also watching: {sys_dl_dir})")
    else:
        sys_dl_dir    = None
        sys_dl_before = set()

    start          = time.time()
    had_crdownload = False

    while time.time() - start < max_wait_seconds:
        current_files = set(os.listdir(download_dir))
        crdownloads   = [f for f in current_files if f.endswith(".crdownload")]
        new_real      = {f for f in (current_files - files_before) if not f.endswith(".crdownload")}

        if crdownloads:
            had_crdownload = True
            print(f"   ⬇️  In progress: {crdownloads[0]}")

        if new_real:
            print(f"\n📁 File(s) detected: {new_real}")
            print("   ⏳ Waiting 2 s for virus scan...")
            time.sleep(2)
            stable = {f for f in new_real
                      if os.path.exists(os.path.join(download_dir, f))
                      and os.path.getsize(os.path.join(download_dir, f)) > 0}
            if stable:
                print(f"✅ Download confirmed: {stable}")
                return True, stable, ""
            else:
                print("🦠 File vanished after download — virus scan / SmartScreen removed it")
                return False, set(), "virus_scan"

        if sys_dl_dir:
            sys_current    = set(os.listdir(sys_dl_dir))
            sys_crdownloads = [f for f in sys_current if f.endswith(".crdownload")]
            sys_new_real   = {f for f in (sys_current - sys_dl_before) if not f.endswith(".crdownload")}

            if sys_crdownloads:
                had_crdownload = True
                print(f"   ⬇️  In progress (fallback dir): {sys_crdownloads[0]}")

            if sys_new_real:
                print(f"\n📁 File(s) detected in fallback dir: {sys_new_real}")
                print("   ⏳ Waiting 3 s for browser to finish writing...")
                time.sleep(3)
                stable = set()
                for f in sys_new_real:
                    src = os.path.join(sys_dl_dir, f)
                    dst = os.path.join(download_dir, f)
                    if os.path.exists(src) and os.path.getsize(src) > 0:
                        try:
                            shutil.move(src, dst)
                            print(f"   📦 Moved from fallback dir → {download_dir}: {f}")
                            stable.add(f)
                        except Exception as mv_err:
                            print(f"   ⚠️  Could not move {f}: {mv_err}")
                if stable:
                    print(f"✅ Download confirmed (from fallback): {stable}")
                    return True, stable, ""

        if had_crdownload and not crdownloads and not new_real:
            fallback_clear = True
            if sys_dl_dir:
                sys_current  = set(os.listdir(sys_dl_dir))
                fallback_new = {f for f in (sys_current - sys_dl_before) if not f.endswith(".crdownload")}
                fallback_clear = not fallback_new and not any(f.endswith(".crdownload") for f in sys_current)
            if fallback_clear:
                print("🦠 .crdownload disappeared with no completed file — virus scan blocked download")
                return False, set(), "virus_scan"

        time.sleep(1)

    elapsed = int(time.time() - start)
    print(f"\n❌ TIMEOUT after {elapsed}s")

    for check_dir, before in [(download_dir, files_before),
                               (sys_dl_dir, sys_dl_before) if sys_dl_dir else (None, None)]:
        if check_dir is None:
            continue
        found = {
            f for f in (set(os.listdir(check_dir)) - before)
            if not f.endswith(".crdownload")
            and os.path.getsize(os.path.join(check_dir, f)) > 0
        }
        if found:
            if check_dir == sys_dl_dir:
                moved = set()
                for f in found:
                    try:
                        shutil.move(os.path.join(sys_dl_dir, f), os.path.join(download_dir, f))
                        moved.add(f)
                    except Exception:
                        pass
                if moved:
                    print(f"📁 Found after timeout in fallback dir, moved: {moved}")
                    return True, moved, ""
            else:
                print(f"📁 Found after timeout: {found}")
                return True, found, ""

    return False, set(), "timeout"


def attempt_download_with_retry(driver, wait, download_dir, max_retries=5):
    """
    Clicks Mass Download → Download Selected Files → Confirm, retrying up to max_retries times.
    Returns: (success: bool, new_files: set)
    """
    print(f"\n{'='*70}")
    print(f"🎯 DOWNLOAD TARGET FOLDER: {download_dir}")
    print(f"{'='*70}")
    print(f"\n{'='*60}")
    print(f"Selecting files for download (Mass Download)")
    print(f"{'='*60}")

    try:
        mass_download = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Mass Download')]")
        ))
        mass_download.click()
        print("✅ Clicked 'Mass Download'")
        time.sleep(2)

        try:
            confirm_wait = WebDriverWait(driver, 3)
            confirm_warning_button = confirm_wait.until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Confirm')]"))
            )
            confirm_warning_button.click()
            print("✅ Clicked warning 'Confirm' button")
            time.sleep(2)
        except TimeoutException:
            print("ℹ️  No warning dialog")

    except Exception as e:
        print(f"❌ Error during Mass Download: {e}")
        return False, set()

    for attempt in range(1, max_retries + 1):
        print(f"\n{'='*60}")
        print(f"📥 DOWNLOAD ATTEMPT {attempt}/{max_retries}")
        print(f"{'='*60}")

        try:
            files_before = set(os.listdir(download_dir))
            print(f"\n📂 Files in folder BEFORE attempt {attempt}: {len(files_before)}")

            print("🖱️  Looking for 'Download Selected Files' button...")
            dl_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Download Selected Files')]")
            ))
            dl_btn.click()
            print("✅ Clicked 'Download Selected Files'")
            time.sleep(1)

            try:
                print("🖱️  Looking for 'Confirm' popup button...")
                confirm_download_wait = WebDriverWait(driver, 8)
                confirm_download_button = confirm_download_wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Confirm')]"))
                )
                confirm_download_button.click()
                print("✅ Clicked 'Confirm' on download popup")
                time.sleep(1)
            except TimeoutException:
                print("⚠️ No Confirm button found - continuing anyway")

            _, new_files, failure_reason = wait_for_download_completion(
                download_dir, files_before)

            if new_files:
                print(f"\n{'='*60}")
                print(f"🎉 SUCCESS! Downloaded {len(new_files)} file(s):")
                for f in new_files:
                    file_path = os.path.join(download_dir, f)
                    size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
                    print(f"   ✓ {f} ({size:,} bytes)")
                print(f"{'='*60}")
                return True, new_files

            if failure_reason == "virus_scan":
                print(f"\n{'='*60}")
                print(f"🦠 ATTEMPT {attempt} BLOCKED BY VIRUS SCAN / SMARTSCREEN")
                print(f"   Edge removed the file after download.")
                print(f"{'='*60}")
            else:
                print(f"\n{'='*60}")
                print(f"❌ ATTEMPT {attempt} FAILED - NO FILES APPEARED")
                print(f"{'='*60}")

            if attempt < max_retries:
                print(f"\n⏳ Waiting 5 seconds before retry {attempt + 1}...")
                time.sleep(5)
                print("🧹 Cleaning up failed download files...")
                for f in os.listdir(download_dir):
                    if f.endswith('.crdownload') or f.endswith('.tmp'):
                        try:
                            os.remove(os.path.join(download_dir, f))
                            print(f"   🗑️  Removed: {f}")
                        except Exception as clean_err:
                            print(f"   ⚠️ Could not remove {f}: {clean_err}")
                print(f"\n🔄 RETRYING... (Attempt {attempt + 1})")

        except Exception as e:
            print(f"\n❌ EXCEPTION during attempt {attempt}: {e}")
            import traceback
            traceback.print_exc()
            if attempt < max_retries:
                print(f"\n⏳ Waiting 5 seconds before retry...")
                time.sleep(5)

    print(f"\n{'='*60}")
    print(f"❌ COMPLETE FAILURE: All {max_retries} attempts unsuccessful")
    print(f"{'='*60}")
    print(f"📂 Final folder contents: {os.listdir(download_dir)}")
    return False, set()


_JUNK_NAMES = {"__macosx", "thumbs.db", ".ds_store"}

def extract_and_flatten(zip_path, dest_dir):
    """Extracts zip (and nested zips) into dest_dir, flattens single-folder nesting."""
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(dest_dir)
    os.remove(zip_path)

    while True:
        inner_zips = [
            os.path.join(root, f)
            for root, _, files in os.walk(dest_dir)
            for f in files
            if f.lower().endswith(".zip")
        ]
        if not inner_zips:
            break
        for inner_zip in inner_zips:
            inner_dest = os.path.splitext(inner_zip)[0]
            os.makedirs(inner_dest, exist_ok=True)
            with zipfile.ZipFile(inner_zip) as z:
                z.extractall(inner_dest)
            os.remove(inner_zip)

    changed = True
    while changed:
        changed = False
        items = os.listdir(dest_dir)
        real  = [i for i in items if i.lower() not in _JUNK_NAMES and not i.startswith('_')]
        if len(real) == 1:
            only_item = os.path.join(dest_dir, real[0])
            if os.path.isdir(only_item):
                for child in os.listdir(only_item):
                    shutil.move(os.path.join(only_item, child), dest_dir)
                shutil.rmtree(only_item)
                for name in items:
                    if name.lower() in _JUNK_NAMES:
                        p = os.path.join(dest_dir, name)
                        shutil.rmtree(p) if os.path.isdir(p) else os.remove(p)
                changed = True


def upload_folder_to_drive(service, folder_path, parent_folder_id):
    """Recursively uploads folder_path into Google Drive under parent_folder_id."""
    if not os.path.isdir(folder_path):
        raise ValueError(f"Folder does not exist: {folder_path}")

    def _create_drive_folder(name, parent_id):
        meta = {"name": name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
        return service.files().create(
            body=meta, fields="id, webViewLink", supportsAllDrives=True
        ).execute()

    def _upload_file(local_path, parent_id):
        mime_type, _ = mimetypes.guess_type(local_path)
        meta  = {"name": os.path.basename(local_path), "parents": [parent_id]}
        media = MediaFileUpload(local_path, mimetype=mime_type or "application/octet-stream")
        print(f"Uploading -> {local_path}")
        attempts = 0
        while attempts < 5:
            try:
                file = service.files().create(
                    body=meta, media_body=media,
                    fields="id, webViewLink", supportsAllDrives=True
                ).execute()
                print(f"✅ Uploaded: {file['id']} ({file['webViewLink']})")
                return file.get("id")
            except HttpError as e:
                code = getattr(e, "status_code", None) or getattr(e.resp, "status", None)
                print(f"⚠️ Upload failed (attempt {attempts+1}), code={code}: {e}")
                attempts += 1
                time.sleep((2 ** attempts) + random.uniform(0, 0.5))
        print(f"❌ Giving up on: {local_path}")
        return None

    folder_name = os.path.basename(folder_path.rstrip(os.sep))
    print(f"Uploading '{folder_name}' to Google Drive...")
    root     = _create_drive_folder(folder_name, parent_folder_id)
    root_id  = root["id"]
    web_link = root.get("webViewLink", "")
    print(f"✅ Created Drive folder: {root_id}")

    dir_to_id = {os.path.abspath(folder_path): root_id}

    for current_dir, subdirs, files in os.walk(folder_path):
        abs_dir   = os.path.abspath(current_dir)
        parent_id = dir_to_id[abs_dir]
        for sub in subdirs:
            sub_local = os.path.join(abs_dir, sub)
            created   = _create_drive_folder(sub, parent_id)
            dir_to_id[os.path.abspath(sub_local)] = created["id"]
            print(f"  📁 {os.path.relpath(sub_local, folder_path)} -> {created['id']}")
        for fname in files:
            if fname.endswith(".crdownload") or fname.endswith(".part"):
                continue
            _upload_file(os.path.join(abs_dir, fname), parent_id)

    print("✅ Upload complete.")
    return web_link


def save_session(driver, filename=SESSION_FILE):
    try:
        with open(filename, 'wb') as f:
            pickle.dump(driver.get_cookies(), f)
        print(f"✅ Session saved to {filename}")
    except Exception as e:
        print(f"⚠️  Could not save session: {e}")


def load_session(driver, filename=SESSION_FILE):
    if not os.path.exists(filename):
        print("ℹ️  No saved session found")
        return False
    try:
        with open(filename, 'rb') as f:
            cookies = pickle.load(f)
        driver.get("https://ktendering.com.kw")
        time.sleep(1)
        for cookie in cookies:
            try:
                driver.add_cookie(cookie)
            except Exception:
                print(f"⚠️  Could not add cookie: {cookie.get('name')}")
        print("✅ Session cookies loaded")
        return True
    except Exception as e:
        print(f"⚠️  Could not load session: {e}")
        return False


# ==========================================================
# SMART LOGIN WITH AUTO OTP FETCH
# ==========================================================

def handle_otp_login(driver, wait, username, password):
    """
    Smart login:
      1. Enter credentials and click Login.
      2. Poll page source for up to 15 s to detect what appears next:
           - OTP page  → auto-fetch OTP from Hostinger IMAP, enter it, submit.
           - Portal    → already logged in, done.
           - Neither   → raise error.

    Ktendering OTP page locators (Keycloak):
      OTP field  : id="otp"  (primary)  |  id="code" (fallback)
      Submit btn : id="kc-login" + data-action-type="login"  (primary)
                   //button[@type='submit'] (fallback)
    """
    try:
        username_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
        password_field = wait.until(EC.presence_of_element_located((By.ID, "password")))
        print("✅ Found username and password fields")

        username_field.clear()
        password_field.clear()
        username_field.send_keys(username)
        password_field.send_keys(password)
        print(f"✅ Entered username: {username[:3]}***")
        print("✅ Entered password: ***")

        login_button = wait.until(EC.element_to_be_clickable((By.ID, "kc-login")))
        login_clicked_at = time.time()
        login_button.click()
        print("✅ Clicked Login — detecting next page…")

        # ── Poll for up to 15 s: OTP page OR portal ──────────────────────────
        otp_detected    = False
        portal_detected = False
        deadline        = time.time() + 15

        while time.time() < deadline:
            src = driver.page_source
            # ktendering OTP page has id="otp" (Keycloak standard) or id="code"
            if 'id="otp"' in src or 'id="code"' in src or 'name="otp"' in src:
                otp_detected = True
                break
            # Portal detected: RFQ list nav link is present
            if "rfq/public" in src or "list-tbody" in src:
                portal_detected = True
                break
            time.sleep(0.5)

        # ── Branch: OTP required ──────────────────────────────────────────────
        if otp_detected:
            print("[OK] OTP page detected — fetching code from mailbox…")

            # Wait for the OTP field to be interactable
            # Primary locator: id="otp"  (ktendering standard)
            # Fallback:        id="code" (some Keycloak versions)
            try:
                otp_field = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "otp"))
                )
                print("[OK] OTP field found: id='otp'")
            except TimeoutException:
                try:
                    otp_field = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "code"))
                    )
                    print("[OK] OTP field found: id='code'")
                except TimeoutException:
                    otp_field = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.NAME, "otp"))
                    )
                    print("[OK] OTP field found: name='otp'")

            # Fetch OTP automatically from Hostinger IMAP
            otp_code = fetch_otp_from_email(
                sent_after_timestamp=login_clicked_at,
                max_wait=90,
                poll_interval=5
            )

            otp_field.clear()
            otp_field.send_keys(otp_code)
            print(f"[OK] Entered OTP: {otp_code}")

            # Submit OTP
            # Primary: id="kc-login" with data-action-type="login"
            # Fallback: any submit button on the page
            try:
                submit_btn = wait.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[@id='kc-login' and @data-action-type='login']")
                ))
                print("[OK] Submit button found: id='kc-login' data-action-type='login'")
            except TimeoutException:
                try:
                    submit_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[@type='submit' and @id='kc-login']")
                    ))
                    print("[OK] Submit button found: id='kc-login' type='submit'")
                except TimeoutException:
                    submit_btn = wait.until(EC.element_to_be_clickable(
                        (By.XPATH, "//button[@type='submit']")
                    ))
                    print("[OK] Submit button found: generic type='submit'")

            driver.execute_script("arguments[0].click();", submit_btn)
            print("[OK] Submitted OTP — waiting for portal…")

            # Wait for successful login after OTP
            deadline_post = time.time() + 30
            while time.time() < deadline_post:
                src = driver.page_source
                if "rfq/public" in src or "list-tbody" in src:
                    print("✅ OTP verification successful! (Found portal/RFQ list)")
                    return True
                # OTP field gone = submitted successfully
                if 'id="otp"' not in src and 'id="code"' not in src and 'name="otp"' not in src:
                    print("✅ OTP verification successful! (OTP field gone)")
                    return True
                time.sleep(1)

            print("❌ OTP submitted but portal not detected within 30 s")
            return False

        # ── Branch: already logged in (no OTP needed) ────────────────────────
        elif portal_detected:
            print("✅ Login successful — no OTP required.")
            return True

        # ── Branch: neither detected ──────────────────────────────────────────
        else:
            # Last-chance check for error messages
            try:
                err = driver.find_element(By.CLASS_NAME, "alert-danger")
                print(f"❌ Login failed: {err.text}")
                return False
            except NoSuchElementException:
                pass
            print("⚠️  Login outcome unclear after 15 s — proceeding cautiously")
            return True

    except TimeoutError as te:
        print(f"❌ OTP fetch timed out: {te}")
        return False
    except Exception as e:
        print(f"❌ Login error: {e}")
        import traceback
        traceback.print_exc()
        return False


def debug_current_page(driver):
    print("\n" + "="*60)
    print("🔍 DEBUG: Current Page State")
    print("="*60)
    print(f"URL: {driver.current_url}")
    print(f"Title: {driver.title}")
    for name, by, value in [
        ("Username field", By.ID, "username"),
        ("Password field", By.ID, "password"),
        ("OTP field",      By.ID, "otp"),
        ("RFQ List",       By.CLASS_NAME, "list-tbody"),
        ("RFQ Link",       By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/public']"),
    ]:
        try:
            driver.find_element(by, value)
            print(f"✅ Found: {name}")
        except Exception:
            print(f"❌ Not found: {name}")
    print("="*60 + "\n")


def is_redirect_error(driver):
    try:
        src   = driver.page_source.lower()
        title = driver.title.lower()
        for indicator in [
            "redirected you too many times", "err_too_many_redirects",
            "session is invalid", "session expired", "session invalid",
            "too many redirects",
        ]:
            if indicator in src or indicator in title:
                print(f"⚠️  Session/redirect error detected: '{indicator}'")
                return True
        return False
    except Exception:
        return False


def handle_session_warning(driver, wait):
    try:
        dialog = driver.find_element(By.ID, "dijit_Dialog_0")
        if not dialog.is_displayed():
            return False
        print("\n" + "="*60)
        print("⚠️  SESSION WARNING DIALOG DETECTED")
        print("="*60)
        try:
            dialog.find_element(By.XPATH, ".//button[@title='Main Page']").click()
            print("✅ Clicked 'Main Page' on session warning dialog")
            time.sleep(2)
        except Exception:
            try:
                dialog.find_element(By.CLASS_NAME, "dijitDialogCloseIcon").click()
                time.sleep(1)
            except Exception:
                pass
        recover_session(driver, wait)
        return True
    except Exception:
        return False


def _recover_to_rfq_list(driver, wait):
    print("🔄 Recovery: attempting to return to RFQ list...")
    try:
        handle_session_warning(driver, wait)

        try:
            cancel_btn = WebDriverWait(driver, 4).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[title='Cancel']")))
            cancel_btn.click()
            time.sleep(1)
            try:
                confirm = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., 'Confirm') or contains(., 'Yes') or contains(., 'OK')]")))
                confirm.click()
                time.sleep(1)
            except TimeoutException:
                pass
            print("   ✅ Clicked Cancel")
        except TimeoutException:
            pass

        try:
            back_btn = WebDriverWait(driver, 4).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[title='Back to List']")))
            back_btn.click()
            time.sleep(1)
            print("   ✅ Clicked Back to List")
        except TimeoutException:
            pass

        try:
            WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
            print("✅ Recovery successful — on RFQ list.")
            return True
        except TimeoutException:
            pass

        print("   ↩️  Navigating directly to RFQ list URL...")
        driver.get("https://ktendering.com.kw/esop/guest/go/neg/rfq/public")
        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
        print("✅ Recovery successful — on RFQ list (via direct URL).")
        return True

    except Exception as e:
        print(f"❌ Recovery failed: {e}")
        return False


def recover_session(driver, wait):
    print("\n" + "="*60)
    print("🔄 SESSION RECOVERY — clearing cookies and re-logging in...")
    print("="*60)
    try:
        driver.delete_all_cookies()
        print("✅ Cleared browser cookies")
    except Exception as e:
        print(f"⚠️  Could not clear browser cookies: {e}")

    if os.path.exists(SESSION_FILE):
        try:
            os.remove(SESSION_FILE)
            print(f"✅ Deleted stale session file: {SESSION_FILE}")
        except Exception as e:
            print(f"⚠️  Could not delete session file: {e}")

    driver.get(LOGIN_URL)
    time.sleep(2)

    if not handle_otp_login(driver, wait, LOGIN, PASSWORD):
        raise Exception("Session recovery failed — could not re-authenticate")

    save_session(driver)
    time.sleep(2)

    try:
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "btn-close"))
        ).click()
        time.sleep(1)
    except TimeoutException:
        pass

    print("✅ Session recovery complete!")


def append_to_titles_json(titles_file, rfq_data):
    if os.path.exists(titles_file):
        try:
            with open(titles_file, 'r', encoding='utf-8') as f:
                titles_history = json.load(f)
            if not isinstance(titles_history, list):
                titles_history = []
        except Exception as e:
            print(f"⚠️  Could not read existing JSON, starting fresh: {e}")
            titles_history = []
    else:
        titles_history = []

    titles_history.append({
        "timestamp":  time.strftime("%Y-%m-%d %H:%M:%S"),
        "full_title": rfq_data.get('full_title',  'N/A'),
        "event_id":   rfq_data.get('event_id',    'N/A'),
        "rfq_number": rfq_data.get('rfq_number',  'N/A'),
        "description":rfq_data.get('description', 'N/A'),
        "folder_path":rfq_data.get('folder_path', 'N/A'),
        "status":     rfq_data.get('status',      'Processed'),
    })

    try:
        with open(titles_file, 'w', encoding='utf-8') as f:
            json.dump(titles_history, f, indent=2, ensure_ascii=False)
        print(f"✅ Updated titles history: {titles_file}")
        print(f"   Total RFQs in history: {len(titles_history)}")
    except Exception as e:
        print(f"❌ Failed to write titles JSON: {e}")


# ==========================================================
# 4. SELENIUM SETUP
# ==========================================================

edge_options = Options()
edge_options.add_experimental_option("prefs", {"download.default_directory": DOWNLOAD_DIR})

driver = webdriver.Edge(service=Service(), options=edge_options)
wait   = WebDriverWait(driver, 20)
print("Browser opened.")

try:
    # --- LOGIN with Session Persistence and Smart OTP ---
    session_valid = False

    if load_session(driver):
        print("\n♻️  Attempting to reuse existing session...")
        driver.get(LOGIN_URL)
        time.sleep(2)
        if is_redirect_error(driver):
            print("❌ Redirect/session error on load — will do fresh login")
        else:
            try:
                WebDriverWait(driver, 5).until(EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/public']")
                ))
                print("✅ Session still valid! Skipping login.")
                session_valid = True
            except TimeoutException:
                print("❌ Session expired, need fresh login")

    if not session_valid:
        print("\n🔐 Starting fresh login...")
        driver.delete_all_cookies()
        if os.path.exists(SESSION_FILE):
            os.remove(SESSION_FILE)
        driver.get(LOGIN_URL)
        time.sleep(2)

        if not handle_otp_login(driver, wait, LOGIN, PASSWORD):
            raise Exception("Login failed - could not authenticate")

        save_session(driver)
        time.sleep(2)

    try:
        wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-close"))).click()
        print("✅ Closed popup after login")
        time.sleep(1)
    except TimeoutException:
        print("⚠️  No popup to close")

    # Navigate to RFQ list
    print("\n📋 Navigating to RFQ list...")
    try:
        driver.find_element(By.CLASS_NAME, "list-tbody")
        print("✅ Already on RFQ list page")
    except Exception:
        rfq_link = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/public']")
        ))
        driver.execute_script("arguments[0].scrollIntoView(true);", rfq_link)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", rfq_link)
        print("✅ Clicked RFQ link")

    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
    print("✅ RFQ list loaded successfully")

    # Authenticate Drive
    creds         = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    drive_service = build("drive", "v3", credentials=creds)
    print("✅ Google Drive authentication successful.")

    # ==========================================================
    # MAIN LOOP
    # ==========================================================
    while True:
        log_data = {
            "RFQ_ID":            "N/A",
            "RFQ_Title":         "N/A",
            "Event_ID":          "N/A",
            "RFQ_Number":        "N/A",
            "Description":       "N/A",
            "Start_Timestamp":   time.strftime("%Y-%m-%d %H:%M:%S"),
            "Status":            "In Progress",
            "Progress/Comments": "Looking for RFQ.",
            "Local_Folder_Path": "N/A",
            "Google_Drive_URL":  "N/A",
        }
        rfq_processed = False

        try:
            handle_session_warning(driver, wait)

            if is_redirect_error(driver):
                recover_session(driver, wait)
                rfq_link = wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/public']")))
                driver.execute_script("arguments[0].click();", rfq_link)
                wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))

            short_wait     = WebDriverWait(driver, 5)
            first_rfq_link = short_wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "detailLink")))
            rfq_processed  = True

            rfq_title_from_list = first_rfq_link.text
            print(f"\n{'='*70}")
            print(f"🎯 Processing RFQ from list: {rfq_title_from_list}")
            print(f"{'='*70}")

            first_rfq_link.click()
            wait.until(EC.visibility_of_element_located((By.NAME, "detailRfqForm")))
            print("✅ RFQ detail page loaded")

            rfq_title_data = extract_rfq_title_from_detail_page(wait)

            log_data["RFQ_Title"]   = rfq_title_data['full_title']
            log_data["Event_ID"]    = rfq_title_data['event_id']
            log_data["RFQ_Number"]  = rfq_title_data['rfq_number']
            log_data["Description"] = rfq_title_data['description']

            match = re.search(r'RFQ#(\d+)', rfq_title_data['full_title'])
            if match:
                log_data["RFQ_ID"] = match.group(1)

            safe_name = re.sub(r'[\\/*?:"<>|]', "", rfq_title_data['description']).replace(" ", "_")
            if not safe_name or safe_name == "NA":
                safe_name = re.sub(r'[\\/*?:"<>|]', "", rfq_title_from_list).replace(" ", "_")

            rfq_num_clean = re.sub(r'[^\d]', '', rfq_title_data['rfq_number'])
            if not rfq_num_clean:
                m = re.search(r'\d+', rfq_title_from_list)
                rfq_num_clean = m.group(0) if m else "unknown"

            timestamp     = time.strftime("%Y%m%d-%H%M%S")
            unique_folder = f"RFQ_{rfq_num_clean}_-_{safe_name}_{timestamp}"
            rfq_folder    = os.path.join(DOWNLOAD_DIR, unique_folder)
            os.makedirs(rfq_folder, exist_ok=True)
            log_data["Local_Folder_Path"] = rfq_folder

            append_to_titles_json(TITLES_JSON_FILE, {
                'full_title':  rfq_title_data['full_title'],
                'event_id':    rfq_title_data['event_id'],
                'rfq_number':  rfq_title_data['rfq_number'],
                'description': rfq_title_data['description'],
                'folder_path': rfq_folder,
                'status':      'Processing'
            })

            metadata_file = os.path.join(rfq_folder, "_RFQ_Metadata.txt")
            with open(metadata_file, "w", encoding="utf-8") as f:
                f.write(f"Full Title: {rfq_title_data['full_title']}\n")
                f.write(f"Event ID: {rfq_title_data['event_id']}\n")
                f.write(f"RFQ Number: {rfq_title_data['rfq_number']}\n")
                f.write(f"Description: {rfq_title_data['description']}\n")
                f.write(f"Extracted At: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            print(f"✅ Created metadata file: {metadata_file}")

            # Express interest
            express_btn = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class, 'toolbar-secondSide')]//button[contains(., 'Express Interest')]")))
            driver.execute_script("arguments[0].click();", express_btn)
            confirm_btn = wait.until(EC.element_to_be_clickable(
                (By.ID, "esop_dialog_for_confirms_confirmButton_BUTTON")))
            confirm_btn.click()
            buyer_tab = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a[href='/esop/toolkit/negotiation/rfq/detailRfqAttachments.do']")))
            buyer_tab.click()

            # Download
            download_success, new_files = attempt_download_with_retry(
                driver, wait, DOWNLOAD_DIR, max_retries=5)

            drive_link = "N/A"

            if download_success and new_files:
                for filename in new_files:
                    print(f"\n📦 Organizing new file: {filename}")
                    moved_file = os.path.join(rfq_folder, filename)
                    shutil.move(os.path.join(DOWNLOAD_DIR, filename), moved_file)
                    if moved_file.lower().endswith(".zip"):
                        print(f"📂 Unzipping and flattening {filename}...")
                        extract_and_flatten(moved_file, rfq_folder)

                drive_link = upload_folder_to_drive(drive_service, rfq_folder, PARENT_FOLDER_ID)
                log_data["Google_Drive_URL"] = drive_link

                if drive_link:
                    with open(os.path.join(rfq_folder, "_GoogleDriveLink.txt"), "w") as f:
                        f.write(drive_link)
                    print(f"✅ Created link file")

                with open(os.path.join(rfq_folder, "_finished.txt"), "w") as f:
                    pass
                print(f"✅ Created sentinel file")

                log_data["Status"]            = "Success"
                log_data["Progress/Comments"] = "Cycle completed successfully."

            else:
                print("\n⚠️  Download failed after all retries — skipping upload, logging failure.")
                log_data["Status"] = "Failed"
                leftover = {
                    f for f in os.listdir(rfq_folder)
                    if not f.startswith("_") and not f.endswith(".crdownload")
                }
                if leftover:
                    log_data["Progress/Comments"] = (
                        "Download flagged as failed but files found in folder — check manually.")
                    print(f"⚠️  Unexpected files in rfq_folder: {leftover}")
                else:
                    log_data["Progress/Comments"] = "Download failed after all retries."

            # Navigate back
            print("\n🔙 Navigating back to RFQ list...")
            handle_session_warning(driver, wait)

            cancel_button = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[title='Cancel']")))
            cancel_button.click()

            try:
                confirm_cancel = WebDriverWait(driver, 3).until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., 'Confirm') or contains(., 'Yes') or contains(., 'OK')]")))
                confirm_cancel.click()
                print("✅ Confirmed cancel dialog")
                time.sleep(1)
            except TimeoutException:
                pass

            wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[title='Back to List']"))).click()
            wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
            print("✅ Returned to RFQ list.")

        except TimeoutException:
            if not rfq_processed:
                print("\n⏹️ No more RFQs found. Exiting loop.")
                log_data["Status"]            = "Idle"
                log_data["Progress/Comments"] = "No RFQs available."
                break
            else:
                print("\n⚠️ Timeout mid-processing — attempting recovery to continue loop...")
                log_data["Status"]            = "Failed"
                log_data["Progress/Comments"] = "Timeout error during RFQ processing."
                if not _recover_to_rfq_list(driver, wait):
                    break

        except HttpError as e:
            log_data["Status"]            = "Failed"
            log_data["Progress/Comments"] = f"Drive error: {e}"
            break

        except Exception as e:
            log_data["Status"]            = "Failed"
            log_data["Progress/Comments"] = f"Error: {e}"
            print(f"\n❌ Unexpected error: {e}")
            import traceback
            traceback.print_exc()
            if not _recover_to_rfq_list(driver, wait):
                break

        finally:
            if rfq_processed:
                try:
                    df = pd.read_excel(LOG_FILE) if os.path.exists(LOG_FILE) else pd.DataFrame()
                    df = pd.concat([df, pd.DataFrame([log_data])], ignore_index=True)
                    df.to_excel(LOG_FILE, index=False)
                    print("✅ Log updated.")
                except Exception as log_e:
                    print(f"❌ Log write failed: {log_e}")

    print("\n✅ All cycles completed.")

except Exception as e:
    print(f"\n❌ Major error: {e}")
    import traceback
    traceback.print_exc()

finally:
    print("\n🔚 Closing browser.")
    if "driver" in locals():
        driver.quit()
