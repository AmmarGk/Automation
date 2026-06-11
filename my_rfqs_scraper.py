import os
import re
import shutil
import mimetypes
import random
import time
import pickle
import zipfile
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from dotenv import load_dotenv

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

load_dotenv()

# --- CONFIGURATION ---
LOGIN = os.getenv("KTENDER_USERNAME")
PASSWORD = os.getenv("KTENDER_PASSWORD")
PARENT_FOLDER_ID = os.getenv("PARENT_FOLDER_ID")
SERVICE_ACCOUNT_FILE = "service_account.json"
SCOPES = ["https://www.googleapis.com/auth/drive"]

if not LOGIN or not PASSWORD:
    raise ValueError("Error: KTENDER_USERNAME or KTENDER_PASSWORD missing in .env")
if not PARENT_FOLDER_ID or PARENT_FOLDER_ID.startswith("YOUR_"):
    raise ValueError("Error: PARENT_FOLDER_ID is not set correctly in .env")
if not os.path.exists(SERVICE_ACCOUNT_FILE):
    raise FileNotFoundError(f"Service account file not found: {SERVICE_ACCOUNT_FILE}")

LOGIN_URL = "https://ktendering.com.kw/esop/guest/login.do?j1pidm=true&internal=false&userAct=changeLangIndex&language=en_GB&_ncp=1758864349514.1369049-1"
SESSION_FILE = os.path.join(os.getcwd(), "ktender_session.pkl")
DOWNLOAD_DIR = os.path.join(os.getcwd(), "Downloaded_RFQs")
LOG_FILE = os.path.join(os.getcwd(), "MyRFQ_Log.xlsx")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

RFQ_NUMBERS = [
    "1060882", "1060987", "1061105",
    "1061152", "1061167", "1061294", "1061347", "1061429", "1061477",
    "1061536", "1061591", "1061597", "1061598", "1061638", "1061640",
    "1020484", "1020533", "1020328", "1020170", "1020664", "1020715",
]


# --- HELPER FUNCTIONS ---

def wait_for_download_completion(driver, download_dir, files_before, max_wait_seconds=120):
    print(f"\n⬇️  Waiting for download to complete...")
    sys_dl_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    if os.path.isdir(sys_dl_dir) and os.path.abspath(sys_dl_dir) != os.path.abspath(download_dir):
        sys_dl_before = set(os.listdir(sys_dl_dir))
        print(f"   (Also watching: {sys_dl_dir})")
    else:
        sys_dl_dir = None
        sys_dl_before = set()

    start = time.time()
    had_crdownload = False

    while time.time() - start < max_wait_seconds:
        current_files = set(os.listdir(download_dir))
        crdownloads = [f for f in current_files if f.endswith(".crdownload")]
        new_real = {f for f in (current_files - files_before) if not f.endswith(".crdownload")}

        if crdownloads:
            had_crdownload = True
            print(f"   ⬇️  In progress: {crdownloads[0]}")

        if new_real:
            print(f"\n📁 File(s) detected: {new_real}")
            stable = set()
            for fname in new_real:
                fpath = os.path.join(download_dir, fname)
                last_size = -1
                stable_ticks = 0
                print(f"   ⏳ Waiting for {fname} to finish writing...")
                for _ in range(60):  # up to ~120 s
                    try:
                        sz = os.path.getsize(fpath)
                        if sz > 0 and sz == last_size:
                            stable_ticks += 1
                            if stable_ticks >= 3:
                                stable.add(fname)
                                break
                        else:
                            stable_ticks = 0
                        last_size = sz
                    except OSError:
                        stable_ticks = 0
                    time.sleep(2)
                else:
                    print(f"   ⚠️  {fname} never stabilised — may be incomplete")
            if stable:
                print(f"✅ Download confirmed: {stable}")
                return True, stable, ""
            else:
                print("🦠 File vanished after download — virus scan / SmartScreen removed it")
                return False, set(), "virus_scan"

        if sys_dl_dir:
            sys_current = set(os.listdir(sys_dl_dir))
            sys_crdownloads = [f for f in sys_current if f.endswith(".crdownload")]
            sys_new_real = {f for f in (sys_current - sys_dl_before) if not f.endswith(".crdownload")}

            if sys_crdownloads:
                had_crdownload = True
                print(f"   ⬇️  In progress (fallback dir): {sys_crdownloads[0]}")

            if sys_new_real:
                print(f"\n📁 File(s) detected in fallback dir: {sys_new_real}")
                stable = set()
                for f in sys_new_real:
                    src = os.path.join(sys_dl_dir, f)
                    dst = os.path.join(download_dir, f)
                    last_size = -1
                    stable_ticks = 0
                    print(f"   ⏳ Waiting for {f} to finish writing...")
                    for _ in range(60):
                        try:
                            sz = os.path.getsize(src)
                            if sz > 0 and sz == last_size:
                                stable_ticks += 1
                                if stable_ticks >= 3:
                                    break
                            else:
                                stable_ticks = 0
                            last_size = sz
                        except OSError:
                            stable_ticks = 0
                        time.sleep(2)
                    if os.path.exists(src) and os.path.getsize(src) > 0:
                        try:
                            shutil.move(src, dst)
                            stable.add(f)
                        except Exception as mv_err:
                            print(f"   ⚠️  Could not move {f}: {mv_err}")
                if stable:
                    print(f"✅ Download confirmed (from fallback): {stable}")
                    return True, stable, ""

        if had_crdownload and not crdownloads and not new_real:
            fallback_clear = True
            if sys_dl_dir:
                sys_current = set(os.listdir(sys_dl_dir))
                fallback_new = {f for f in (sys_current - sys_dl_before) if not f.endswith(".crdownload")}
                fallback_clear = not fallback_new and not any(f.endswith(".crdownload") for f in sys_current)
            if fallback_clear:
                print("🦠 .crdownload disappeared with no completed file — virus scan blocked download")
                return False, set(), "virus_scan"

        time.sleep(1)

    elapsed = int(time.time() - start)
    print(f"\n❌ TIMEOUT after {elapsed}s")
    for check_dir, before in [(download_dir, files_before), (sys_dl_dir, sys_dl_before) if sys_dl_dir else (None, None)]:
        if check_dir is None:
            continue
        found = {f for f in (set(os.listdir(check_dir)) - before) if not f.endswith(".crdownload") and os.path.getsize(os.path.join(check_dir, f)) > 0}
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
                    return True, moved, ""
            else:
                return True, found, ""
    return False, set(), "timeout"


def attempt_download_with_retry(driver, wait, download_dir, max_retries=5):
    print(f"\n{'='*70}")
    print(f"🎯 DOWNLOAD TARGET FOLDER: {download_dir}")
    print(f"{'='*70}")

    try:
        mass_download = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Mass Download')]")
        ))
        mass_download.click()
        print("✅ Clicked 'Mass Download'")
        time.sleep(2)
        try:
            confirm_warning_button = WebDriverWait(driver, 3).until(
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
            dl_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Download Selected Files')]")
            ))
            dl_btn.click()
            print("✅ Clicked 'Download Selected Files'")
            time.sleep(1)
            try:
                confirm_download_button = WebDriverWait(driver, 8).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Confirm')]"))
                )
                confirm_download_button.click()
                print("✅ Clicked 'Confirm' on download popup")
                time.sleep(1)
            except TimeoutException:
                print("⚠️ No Confirm button found - continuing anyway")

            success, new_files, failure_reason = wait_for_download_completion(driver, download_dir, files_before)

            if new_files:
                print(f"\n🎉 SUCCESS! Downloaded {len(new_files)} file(s)")
                return True, new_files

            if attempt < max_retries:
                time.sleep(5)
                for f in os.listdir(download_dir):
                    if f.endswith('.crdownload') or f.endswith('.tmp'):
                        try:
                            os.remove(os.path.join(download_dir, f))
                        except Exception:
                            pass
        except Exception as e:
            print(f"\n❌ EXCEPTION during attempt {attempt}: {e}")
            if attempt < max_retries:
                time.sleep(5)

    print(f"\n❌ COMPLETE FAILURE: All {max_retries} attempts unsuccessful")
    return False, set()


def extract_and_flatten(zip_path, dest_dir):
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(dest_dir)
    os.remove(zip_path)
    for root, _, files in os.walk(dest_dir):
        for f in files:
            if f.lower().endswith(".zip"):
                inner_zip = os.path.join(root, f)
                inner_dest = os.path.splitext(inner_zip)[0]
                os.makedirs(inner_dest, exist_ok=True)
                extract_and_flatten(inner_zip, inner_dest)
    items = os.listdir(dest_dir)
    if len(items) == 1:
        only_item = os.path.join(dest_dir, items[0])
        if os.path.isdir(only_item):
            for child in os.listdir(only_item):
                shutil.move(os.path.join(only_item, child), dest_dir)
            shutil.rmtree(only_item)


def upload_folder_to_drive(service, folder_path, parent_folder_id):
    if not os.path.isdir(folder_path):
        raise ValueError(f"Folder does not exist: {folder_path}")

    def _create_drive_folder(name, parent_id):
        meta = {"name": name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
        return service.files().create(body=meta, fields="id, webViewLink", supportsAllDrives=True).execute()

    def _upload_file(local_path, parent_id):
        mime_type, _ = mimetypes.guess_type(local_path)
        meta = {"name": os.path.basename(local_path), "parents": [parent_id]}
        media = MediaFileUpload(local_path, mimetype=mime_type or "application/octet-stream")
        attempts = 0
        while attempts < 5:
            try:
                file = service.files().create(body=meta, media_body=media, fields="id, webViewLink", supportsAllDrives=True).execute()
                print(f"✅ Uploaded: {file['id']}")
                return file.get("id")
            except HttpError as e:
                attempts += 1
                time.sleep((2 ** attempts) + random.uniform(0, 0.5))
        print(f"❌ Giving up on file: {local_path}")
        return None

    folder_name = os.path.basename(folder_path.rstrip(os.sep))
    root = _create_drive_folder(folder_name, parent_folder_id)
    root_id = root["id"]
    web_link = root.get("webViewLink", "")
    print(f"✅ Created Drive folder: {root_id}")
    dir_to_id = {os.path.abspath(folder_path): root_id}

    for current_dir, subdirs, files in os.walk(folder_path):
        abs_dir = os.path.abspath(current_dir)
        parent_id = dir_to_id[abs_dir]
        for sub in subdirs:
            sub_local = os.path.join(abs_dir, sub)
            created = _create_drive_folder(sub, parent_id)
            dir_to_id[os.path.abspath(sub_local)] = created["id"]
        for fname in files:
            if fname.endswith(".crdownload") or fname.endswith(".part"):
                continue
            _upload_file(os.path.join(abs_dir, fname), parent_id)

    print("✅ Upload complete.")
    return web_link


def save_session(driver):
    try:
        cookies = driver.get_cookies()
        with open(SESSION_FILE, 'wb') as f:
            pickle.dump(cookies, f)
        print(f"✅ Session saved")
    except Exception as e:
        print(f"⚠️  Could not save session: {e}")


def load_session(driver):
    if not os.path.exists(SESSION_FILE):
        print("ℹ️  No saved session found")
        return False
    try:
        with open(SESSION_FILE, 'rb') as f:
            cookies = pickle.load(f)
        driver.get("https://ktendering.com.kw")
        time.sleep(1)
        for cookie in cookies:
            try:
                driver.add_cookie(cookie)
            except Exception:
                pass
        print("✅ Session cookies loaded")
        return True
    except Exception as e:
        print(f"⚠️  Could not load session: {e}")
        return False


def is_redirect_error(driver):
    try:
        src = driver.page_source.lower()
        title = driver.title.lower()
        for indicator in ["redirected you too many times", "err_too_many_redirects",
                          "session is invalid", "session expired", "session invalid", "too many redirects"]:
            if indicator in src or indicator in title:
                print(f"⚠️  Session/redirect error detected: '{indicator}'")
                return True
        return False
    except Exception:
        return False


def handle_otp_login(driver, wait, username, password, max_otp_wait=120):
    try:
        username_field = wait.until(EC.presence_of_element_located((By.ID, "username")))
        password_field = wait.until(EC.presence_of_element_located((By.ID, "password")))
        username_field.clear()
        password_field.clear()
        username_field.send_keys(username)
        password_field.send_keys(password)
        print(f"✅ Entered username: {username[:3]}***")
        login_button = wait.until(EC.element_to_be_clickable((By.ID, "kc-login")))
        login_button.click()
        print("✅ Clicked Login button")
        time.sleep(2)

        try:
            WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "otp")))
            print("\n" + "="*60)
            print("🔐 OTP REQUIRED — enter in browser (2 min)")
            print("="*60 + "\n")
            start_time = time.time()
            while time.time() - start_time < max_otp_wait:
                try:
                    driver.find_element(By.ID, "otp")
                except Exception:
                    print("\n✅ OTP verification successful!")
                    return True
                remaining = int(max_otp_wait - (time.time() - start_time))
                print(f"⏳ Waiting for OTP... ({remaining}s remaining)", end='\r')
                time.sleep(2)
            print("\n❌ OTP timeout")
            return False
        except TimeoutException:
            print("ℹ️  No OTP required — checking login status...")
            try:
                WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/my']"))
                )
                print("✅ Login successful")
                return True
            except TimeoutException:
                try:
                    error_msg = driver.find_element(By.CLASS_NAME, "alert-danger")
                    print(f"❌ Login failed: {error_msg.text}")
                    return False
                except Exception:
                    print("⚠️  Login status unclear — proceeding cautiously")
                    return True
    except Exception as e:
        print(f"❌ Login error: {e}")
        import traceback
        traceback.print_exc()
        return False


def navigate_to_my_rfqs(driver, wait):
    """Recovery fallback: go back to My RFQs list via Back to List or nav link."""
    try:
        back_btn = WebDriverWait(driver, 4).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[title='Back to List']"))
        )
        back_btn.click()
        print("   ✅ Clicked 'Back to List'")
        wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
        return
    except TimeoutException:
        pass
    nav_link = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/my']")
    ))
    driver.execute_script("arguments[0].click();", nav_link)
    print("   ✅ Clicked 'My RFQs' nav link")
    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))


# --- SELENIUM SETUP ---
edge_options = Options()
prefs = {"download.default_directory": DOWNLOAD_DIR}
edge_options.add_experimental_option("prefs", prefs)

driver = webdriver.Edge(service=Service(), options=edge_options)
wait = WebDriverWait(driver, 20)
print("Browser opened.")

try:
    session_valid = False

    if load_session(driver):
        print("\n♻️  Attempting to reuse existing session...")
        driver.get(LOGIN_URL)
        time.sleep(2)
        if is_redirect_error(driver):
            print("❌ Redirect/session error — will do fresh login")
        else:
            try:
                WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/my']"))
                )
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

    # Navigate to My RFQs / Tenders / Auctions
    print("\n📋 Navigating to My RFQs / Tenders / Auctions...")
    my_rfqs_link = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/my']")
    ))
    driver.execute_script("arguments[0].scrollIntoView(true);", my_rfqs_link)
    time.sleep(0.5)
    driver.execute_script("arguments[0].click();", my_rfqs_link)
    print("✅ Clicked 'My RFQs / Tenders / Auctions' link")

    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
    print("✅ My RFQs / Tenders / Auctions page loaded successfully")

    # --- Authenticate Google Drive ---
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    drive_service = build("drive", "v3", credentials=creds)
    print("✅ Google Drive authentication successful.")

    print(f"\n📂 {len(RFQ_NUMBERS)} RFQ numbers to process")

    # --- MAIN LOOP ---
    for idx, rfq_num in enumerate(RFQ_NUMBERS, start=1):
        print(f"\n{'='*70}")
        print(f"[{idx}/{len(RFQ_NUMBERS)}] Processing RFQ: {rfq_num}")
        print(f"{'='*70}")

        log_data = {
            "RFQ_Number": rfq_num,
            "RFQ_Title": "N/A",
            "Start_Timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "Status": "In Progress",
            "Progress/Comments": "Starting.",
            "Local_Folder_Path": "N/A",
            "Google_Drive_URL": "N/A",
        }

        try:
            # Step 1: Select "RFQ Descriptions" filter and enter RFQ number
            combobox_input = wait.until(EC.element_to_be_clickable((By.ID, "filterPickerSelect")))
            combobox_input.clear()
            combobox_input.send_keys("RFQ Descriptions")
            time.sleep(0.5)
            combobox_input.send_keys(Keys.ARROW_DOWN)
            time.sleep(0.3)
            combobox_input.send_keys(Keys.RETURN)
            print("✅ Selected 'RFQ Descriptions' via keyboard")
            time.sleep(1)

            desc_field = wait.until(EC.element_to_be_clickable((By.ID, "RfqDescription_FILTER")))
            desc_field.clear()
            desc_field.send_keys(rfq_num)
            print(f"✅ Entered '{rfq_num}' into RFQ Description filter")
            time.sleep(0.5)

            # Step 2: Click Search — wait for DOM to refresh before reading results
            try:
                old_tbody = driver.find_element(By.CLASS_NAME, "list-tbody")
            except Exception:
                old_tbody = None
            search_btn = wait.until(EC.element_to_be_clickable((By.ID, "filterSearchButton_BUTTON")))
            search_btn.click()
            print("✅ Clicked Search")
            if old_tbody:
                wait.until(EC.staleness_of(old_tbody))
            wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
            print("✅ Results refreshed")

            # Step 3: Click Title column link (with validation)
            first_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "td.col_TITLE a.detailLink"))
            )
            link_text = first_link.text.strip()
            log_data["RFQ_Title"] = link_text

            if rfq_num not in link_text:
                print(f"⚠️  Search returned '{link_text}' — does not contain '{rfq_num}'. Skipping.")
                log_data["Status"] = "Skipped"
                log_data["Progress/Comments"] = f"Search result mismatch: got '{link_text}'"
                driver.back()
                wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
                continue

            print(f"✅ Found title link: '{link_text}'")
            driver.execute_script("arguments[0].click();", first_link)

            # Step 4: Wait for detail page
            wait.until(EC.visibility_of_element_located((By.NAME, "detailRfqForm")))
            print(f"✅ Detail page loaded")

            # Step 5: Click Buyer Attachments tab
            buyer_tab = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a[href='/esop/toolkit/negotiation/rfq/detailRfqAttachments.do']")
            ))
            buyer_tab.click()
            print("✅ Clicked 'Buyer Attachments'")
            time.sleep(1)

            # Step 6: Create local folder for this RFQ
            safe_title = re.sub(r'[\\/*?:"<>|]', "", link_text).replace(" ", "_")
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            rfq_folder = os.path.join(DOWNLOAD_DIR, f"RFQ_{safe_title}_{timestamp}")
            os.makedirs(rfq_folder, exist_ok=True)
            log_data["Local_Folder_Path"] = rfq_folder

            # Step 7: Mass Download
            download_success, new_files = attempt_download_with_retry(driver, wait, DOWNLOAD_DIR, max_retries=5)

            if download_success and new_files:
                for filename in new_files:
                    print(f"\n📦 Organizing: {filename}")
                    moved_file = os.path.join(rfq_folder, filename)
                    shutil.move(os.path.join(DOWNLOAD_DIR, filename), moved_file)
                    if moved_file.lower().endswith(".zip"):
                        print(f"📂 Unzipping {filename}...")
                        extract_and_flatten(moved_file, rfq_folder)

                # Step 8: Upload to Drive
                drive_link = upload_folder_to_drive(drive_service, rfq_folder, PARENT_FOLDER_ID)
                log_data["Google_Drive_URL"] = drive_link

                if drive_link:
                    with open(os.path.join(rfq_folder, "_GoogleDriveLink.txt"), "w") as f:
                        f.write(drive_link)

                with open(os.path.join(rfq_folder, "_finished.txt"), "w") as f:
                    pass

                log_data["Status"] = "Success"
                log_data["Progress/Comments"] = "Completed successfully."
            else:
                print("\n⚠️  Download failed after all retries.")
                log_data["Status"] = "Failed"
                log_data["Progress/Comments"] = "Download failed after all retries."

            # Step 9: Navigate back to My RFQs list (exactly 3 backs, wait for full page load each time)
            print("\n🔙 Navigating back to My RFQs list...")
            for _ in range(3):
                driver.back()
                WebDriverWait(driver, 15).until(
                    lambda d: d.execute_script("return document.readyState") == "complete"
                )
                if "joinRfq/list.si" in driver.current_url:
                    print(f"✅ Back on My RFQs list ({driver.current_url})")
                    break
            else:
                print("⚠️  Not on list after 3 backs — navigating via nav link")
                my_rfqs_link = wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/my']")
                ))
                driver.execute_script("arguments[0].click();", my_rfqs_link)
                wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
                print("✅ Back on My RFQs list via nav link")
            time.sleep(1)

        except TimeoutException as e:
            print(f"⚠️  Timeout processing RFQ {rfq_num}: {e}")
            log_data["Status"] = "Failed"
            log_data["Progress/Comments"] = f"Timeout: {e}"
            try:
                for _ in range(3):
                    driver.back()
                    WebDriverWait(driver, 15).until(
                        lambda d: d.execute_script("return document.readyState") == "complete"
                    )
                    if "joinRfq/list.si" in driver.current_url:
                        print(f"✅ Recovered — back on My RFQs list ({driver.current_url})")
                        break
                else:
                    my_rfqs_link = wait.until(EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/my']")
                    ))
                    driver.execute_script("arguments[0].click();", my_rfqs_link)
                    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
                    print("✅ Recovered — back on My RFQs list via nav link")
            except Exception:
                print("❌ Could not recover to list page")

        except HttpError as e:
            log_data["Status"] = "Failed"
            log_data["Progress/Comments"] = f"Drive error: {e}"
            print(f"❌ Drive error for RFQ {rfq_num}: {e}")
            try:
                for _ in range(3):
                    driver.back()
                    WebDriverWait(driver, 15).until(
                        lambda d: d.execute_script("return document.readyState") == "complete"
                    )
                    if "joinRfq/list.si" in driver.current_url:
                        print(f"✅ Recovered — back on My RFQs list ({driver.current_url})")
                        break
                else:
                    my_rfqs_link = wait.until(EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/my']")
                    ))
                    driver.execute_script("arguments[0].click();", my_rfqs_link)
                    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
                    print("✅ Recovered — back on My RFQs list via nav link")
            except Exception:
                pass

        except Exception as e:
            print(f"❌ Error processing RFQ {rfq_num}: {e}")
            import traceback
            traceback.print_exc()
            log_data["Status"] = "Failed"
            log_data["Progress/Comments"] = f"Error: {e}"
            try:
                for _ in range(3):
                    driver.back()
                    WebDriverWait(driver, 15).until(
                        lambda d: d.execute_script("return document.readyState") == "complete"
                    )
                    if "joinRfq/list.si" in driver.current_url:
                        print(f"✅ Recovered — back on My RFQs list ({driver.current_url})")
                        break
                else:
                    my_rfqs_link = wait.until(EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/my']")
                    ))
                    driver.execute_script("arguments[0].click();", my_rfqs_link)
                    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
                    print("✅ Recovered — back on My RFQs list via nav link")
            except Exception:
                pass

        finally:
            try:
                df = pd.read_excel(LOG_FILE) if os.path.exists(LOG_FILE) else pd.DataFrame()
                df = pd.concat([df, pd.DataFrame([log_data])], ignore_index=True)
                df.to_excel(LOG_FILE, index=False)
                print("✅ Log updated.")
            except Exception as log_e:
                print(f"❌ Log write failed: {log_e}")

    print(f"\n{'='*70}")
    print(f"✅ All {len(RFQ_NUMBERS)} RFQs processed.")
    print(f"{'='*70}")

except Exception as e:
    print(f"\n❌ Fatal error: {e}")
    import traceback
    traceback.print_exc()

finally:
    print("\n🔚 Closing browser.")
    if "driver" in locals():
        driver.quit()
