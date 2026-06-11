import os
import time
import re
import shutil
import mimetypes
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from dotenv import load_dotenv
# -*- coding: utf-8 -*-
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# --- GOOGLE DRIVE IMPORTS ---
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError
import zipfile

# Load environment variables from .env file
load_dotenv()

# --- 1. SETUP AND CONFIGURATION ---

# Get credentials securely from .env file
LOGIN = os.getenv("KTENDER_USERNAME")
PASSWORD = os.getenv("KTENDER_PASSWORD")

# Google Drive config
PARENT_FOLDER_ID = os.getenv("PARENT_FOLDER_ID")
SERVICE_ACCOUNT_FILE = "service_account.json"
SCOPES = ["https://www.googleapis.com/auth/drive"]

# Validate inputs
if not LOGIN or not PASSWORD:
    raise ValueError("Error: KTENDER_USERNAME or KTENDER_PASSWORD missing in .env")
if not PARENT_FOLDER_ID or PARENT_FOLDER_ID.startswith("YOUR_"):
    raise ValueError("Error: PARENT_FOLDER_ID is not set correctly in .env")
if not os.path.exists(SERVICE_ACCOUNT_FILE):
    raise FileNotFoundError(f"Service account file not found: {SERVICE_ACCOUNT_FILE}")

# URLs
LOGIN_URL = "https://ktendering.com.kw/esop/guest/login.do?j1pidm=true&internal=false&userAct=changeLangIndex&language=en_GB&_ncp=1758864349514.1369049-1"

# Local paths
DOWNLOAD_DIR = os.path.join(os.getcwd(), "Downloaded_RFQs")
LOG_FILE = os.path.join(os.getcwd(), "RFQ_Log.xlsx")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)


# --- 2. HELPER FUNCTIONS ---

def wait_for_download_completion(download_dir, files_before, max_wait_seconds=60, check_interval=1):
    """
    Waits for Chrome downloads to complete by monitoring for new files.
    Returns True only if NEW FILES actually appear in the directory.
    
    Returns:
        tuple: (success: bool, new_files: set)
    """
    print(f"\n🔍 MONITORING FOLDER: {download_dir}")
    print(f"📂 Files BEFORE download: {len(files_before)}")
    if files_before:
        print(f"   Existing: {list(files_before)[:3]}...")
    
    elapsed_time = 0
    last_file_count = len(files_before)
    
    while elapsed_time < max_wait_seconds:
        # Check current state
        current_files = set(os.listdir(download_dir))
        current_count = len(current_files)
        crdownload_files = [f for f in current_files if f.endswith('.crdownload')]
        new_files = current_files - files_before
        new_actual_files = [f for f in new_files if not f.endswith('.crdownload')]
        
        # Detailed status every check
        print(f"⏱️  {elapsed_time}s | Total: {current_count} | .crdownload: {len(crdownload_files)} | New: {len(new_files)} | New actual: {len(new_actual_files)}")
        
        # Success condition: We have NEW files that are NOT .crdownload
        if new_actual_files:
            print(f"\n✅ SUCCESS! Found {len(new_actual_files)} new file(s):")
            for f in new_actual_files:
                print(f"   📦 {f}")
            return True, set(new_actual_files)
        
        # If .crdownload files exist, download is in progress
        if crdownload_files:
            print(f"   ⬇️  Downloading: {crdownload_files}")
        
        time.sleep(check_interval)
        elapsed_time += check_interval
    
    # Timeout
    print(f"\n❌ TIMEOUT after {elapsed_time}s")
    print(f"📂 Files AFTER: {len(os.listdir(download_dir))}")
    final_files = set(os.listdir(download_dir))
    new_files = final_files - files_before
    print(f"🔍 New files found: {new_files if new_files else 'NONE'}")
    
    return False, new_files


def attempt_download_with_retry(driver, wait, download_dir, max_retries=5):
    """
    Attempts to download files with aggressive retry logic.
    Retries until files are actually downloaded, regardless of failure reason.
    
    Returns:
        tuple: (success: bool, new_files: set)
    """
    print(f"\n{'='*70}")
    print(f"🎯 DOWNLOAD TARGET FOLDER: {download_dir}")
    print(f"{'='*70}")
    
    # First, do Mass Download to select files (only once)
    print(f"\n{'='*60}")
    print(f"Selecting files for download (Mass Download)")
    print(f"{'='*60}")
    
    try:
        # Click Mass Download button
        mass_download = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(., 'Mass Download')]")
        ))
        mass_download.click()
        print("✅ Clicked 'Mass Download'")
        time.sleep(2)
        
        # Handle potential warning confirmation
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
    
    # Now repeatedly try the actual download until it succeeds
    for attempt in range(1, max_retries + 1):
        print(f"\n{'='*60}")
        print(f"📥 DOWNLOAD ATTEMPT {attempt}/{max_retries}")
        print(f"{'='*60}")
        
        try:
            # Record files before this download attempt
            files_before = set(os.listdir(download_dir))
            print(f"\n📂 Files in folder BEFORE attempt {attempt}: {len(files_before)}")
            
            # Click Download Selected Files button
            print("🖱️  Looking for 'Download Selected Files' button...")
            dl_btn = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Download Selected Files')]")
            ))
            dl_btn.click()
            print("✅ Clicked 'Download Selected Files'")
            time.sleep(1)
            
            # Handle the download confirmation popup
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
            
            # Wait for download to complete
            print(f"\n⏳ Monitoring {download_dir} for new files...")
            success, new_files = wait_for_download_completion(download_dir, files_before, max_wait_seconds=45)
            
            # If we got files, we're done!
            if success and new_files:
                print(f"\n{'='*60}")
                print(f"🎉 SUCCESS! Downloaded {len(new_files)} file(s):")
                for f in new_files:
                    file_path = os.path.join(download_dir, f)
                    size = os.path.getsize(file_path) if os.path.exists(file_path) else 0
                    print(f"   ✓ {f} ({size:,} bytes)")
                print(f"{'='*60}")
                return True, new_files
            
            # Download failed - prepare for retry
            print(f"\n{'='*60}")
            print(f"❌ ATTEMPT {attempt} FAILED - NO FILES APPEARED")
            print(f"{'='*60}")
            
            # If we have more retries, wait and try again
            if attempt < max_retries:
                wait_time = 5
                print(f"\n⏳ Waiting {wait_time} seconds before retry {attempt + 1}...")
                time.sleep(wait_time)
                
                # Clean up any failed download remnants
                print("🧹 Cleaning up failed download files...")
                current_files = os.listdir(download_dir)
                for f in current_files:
                    if f.endswith('.crdownload') or f.endswith('.tmp'):
                        try:
                            file_to_remove = os.path.join(download_dir, f)
                            os.remove(file_to_remove)
                            print(f"   🗑️  Removed: {f}")
                        except Exception as clean_err:
                            print(f"   ⚠️ Could not remove {f}: {clean_err}")
                
                print(f"\n🔄 RETRYING... (Attempt {attempt + 1})")
            
        except Exception as e:
            print(f"\n❌ EXCEPTION during attempt {attempt}:")
            print(f"   Error: {e}")
            import traceback
            traceback.print_exc()
            
            if attempt < max_retries:
                print(f"\n⏳ Waiting 5 seconds before retry...")
                time.sleep(5)
    
    # All retries exhausted
    print(f"\n{'='*60}")
    print(f"❌ COMPLETE FAILURE: All {max_retries} attempts unsuccessful")
    print(f"{'='*60}")
    print(f"📂 Final folder contents: {os.listdir(download_dir)}")
    print("💡 TIPS:")
    print("   - Check if Chrome is blocking downloads")
    print("   - Check Chrome downloads page (chrome://downloads/)")
    print("   - Files might need manual approval")
    return False, set()


def extract_and_flatten(zip_path, dest_dir):
    """Extracts zip (and nested zips) into dest_dir, flattens single-folder nesting."""
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(dest_dir)
    os.remove(zip_path)

    # Recursively handle inner zips
    for root, _, files in os.walk(dest_dir):
        for f in files:
            if f.lower().endswith(".zip"):
                inner_zip = os.path.join(root, f)
                inner_dest = os.path.splitext(inner_zip)[0]
                os.makedirs(inner_dest, exist_ok=True)
                extract_and_flatten(inner_zip, inner_dest)

    # Flatten if only one subfolder exists
    items = os.listdir(dest_dir)
    if len(items) == 1:
        only_item = os.path.join(dest_dir, items[0])
        if os.path.isdir(only_item):
            for child in os.listdir(only_item):
                shutil.move(os.path.join(only_item, child), dest_dir)
            shutil.rmtree(only_item)


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
        """Upload one file with retries and logging."""
        mime_type, _ = mimetypes.guess_type(local_path)
        meta = {"name": os.path.basename(local_path), "parents": [parent_id]}
        media = MediaFileUpload(local_path, mimetype=mime_type or "application/octet-stream")

        print(f"Uploading -> {local_path}")
        attempts = 0
        while attempts < 5:
            try:
                file = service.files().create(
                    body=meta,
                    media_body=media,
                    fields="id, webViewLink",
                    supportsAllDrives=True
                ).execute()
                print(f"✅ Uploaded to Drive: {file['id']} (link: {file['webViewLink']})")
                return file.get("id")
            except HttpError as e:
                code = getattr(e, "status_code", None) or getattr(e.resp, "status", None)
                print(f"⚠️ Upload failed (attempt {attempts+1}), code={code}, error={e}")
                attempts += 1
                time.sleep((2 ** attempts) + random.uniform(0, 0.5))

        print(f"❌ Giving up on file after {attempts} attempts: {local_path}")
        return None

    # Create root folder
    folder_name = os.path.basename(folder_path.rstrip(os.sep))
    print(f"Uploading '{folder_name}' to Google Drive...")
    root = _create_drive_folder(folder_name, parent_folder_id)
    root_id = root["id"]
    web_link = root.get("webViewLink", "")
    print(f"✅ Created Google Drive folder: {root_id}")

    # Map local dir → Drive folder id
    dir_to_id = {os.path.abspath(folder_path): root_id}

    for current_dir, subdirs, files in os.walk(folder_path):
        abs_dir = os.path.abspath(current_dir)
        parent_id = dir_to_id[abs_dir]

        # Create subfolders
        for sub in subdirs:
            sub_local = os.path.join(abs_dir, sub)
            created = _create_drive_folder(sub, parent_id)
            dir_to_id[os.path.abspath(sub_local)] = created["id"]
            print(f"  📁 {os.path.relpath(sub_local, folder_path)} -> {created['id']}")

        # Upload files
        for fname in files:
            if fname.endswith(".crdownload") or fname.endswith(".part"):
                continue
            local_file = os.path.join(abs_dir, fname)
            _upload_file(local_file, parent_id)

    print("✅ Upload complete.")
    return web_link


# --- 3. SELENIUM SETUP ---

edge_options = Options()
prefs = {"download.default_directory": DOWNLOAD_DIR}
edge_options.add_experimental_option("prefs", prefs)

driver = webdriver.Edge(service=Service(), options=edge_options)
wait = WebDriverWait(driver, 20)
print("Browser opened.")

try:
    # --- LOGIN ---
    driver.get(LOGIN_URL)
    print("On login page.")
    username_field = wait.until(EC.presence_of_element_located((By.ID, "login")))
    password_field = wait.until(EC.presence_of_element_located((By.ID, "password")))
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Submit')]")))
    username_field.send_keys(LOGIN)
    password_field.send_keys(PASSWORD)
    login_button.click()
    print("✅ Login successful!")

    # Close any popup/modal after login
    try:
        close_button = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "btn-close")))
        close_button.click()
        print("✅ Closed popup after login")
        time.sleep(1)  # Give page a moment to settle
    except TimeoutException:
        print("⚠️ No popup to close")
    
    # Wait for the RFQ link to be clickable and visible
    print("Looking for RFQ link...")
    rfq_link = wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, "a[href='/esop/guest/go/neg/rfq/public']")
    ))
    
    # Scroll into view to ensure it's visible
    driver.execute_script("arguments[0].scrollIntoView(true);", rfq_link)
    time.sleep(0.5)
    
    print(f"✅ Found RFQ link: '{rfq_link.text.strip()}'")
    
    # Click using JavaScript to avoid any overlay issues
    driver.execute_script("arguments[0].click();", rfq_link)
    print("✅ Clicked RFQ link")
    
    # Wait for RFQ list page to load
    wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
    print("✅ RFQ list loaded.")

    # --- AUTHENTICATE DRIVE ---
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    drive_service = build("drive", "v3", credentials=creds)
    print("✅ Google Drive authentication successful.")

    # --- MAIN LOOP ---
    while True:
        log_data = {
            "RFQ_ID": "N/A", "RFQ_Title": "N/A", "Start_Timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "Status": "In Progress", "Progress/Comments": "Looking for RFQ.",
            "Local_Folder_Path": "N/A", "Google_Drive_URL": "N/A",
        }
        rfq_processed = False

        try:
            short_wait = WebDriverWait(driver, 5)
            first_rfq_link = short_wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "detailLink")))
            rfq_processed = True

            rfq_title = first_rfq_link.text
            log_data["RFQ_Title"] = rfq_title
            match = re.search(r"^\d+", rfq_title)
            if match:
                log_data["RFQ_ID"] = match.group(0)

            safe_name = re.sub(r'[\\/*?:"<>|]', "", rfq_title).replace(" ", "_")
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            unique_folder = f"RFQ_{safe_name}_{timestamp}"
            rfq_folder = os.path.join(DOWNLOAD_DIR, unique_folder)
            os.makedirs(rfq_folder, exist_ok=True)
            log_data["Local_Folder_Path"] = rfq_folder

            # Open RFQ
            first_rfq_link.click()
            wait.until(EC.visibility_of_element_located((By.NAME, "detailRfqForm")))

            # Express interest
            express_btn = wait.until(EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class, 'toolbar-secondSide')]//button[contains(., 'Express Interest')]")))
            driver.execute_script("arguments[0].click();", express_btn)
            confirm_btn = wait.until(EC.element_to_be_clickable((By.ID, "esop_dialog_for_confirms_confirmButton_BUTTON")))
            confirm_btn.click()
            buyer_tab = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "a[href='/esop/toolkit/negotiation/rfq/detailRfqAttachments.do']")))
            buyer_tab.click()

            # Download files with retry logic (5 attempts)
            download_success, new_files = attempt_download_with_retry(driver, wait, DOWNLOAD_DIR, max_retries=5)

            if not download_success or not new_files:
                print("\n⚠️ WARNING: No files were downloaded after all retry attempts!")
                log_data["Progress/Comments"] = "Download failed after retries"
            else:
                # Process downloaded files
                for filename in new_files:
                    print(f"\n📦 Organizing new file: {filename}")
                    moved_file = os.path.join(rfq_folder, filename)
                    shutil.move(os.path.join(DOWNLOAD_DIR, filename), moved_file)
                    if moved_file.lower().endswith(".zip"):
                        print(f"📂 Unzipping and flattening {filename}...")
                        extract_and_flatten(moved_file, rfq_folder)

                # Upload to Drive
                drive_link = upload_folder_to_drive(drive_service, rfq_folder, PARENT_FOLDER_ID)
                log_data["Google_Drive_URL"] = drive_link

                # Create a text file with the Google Drive link
                if drive_link:
                    link_file_path = os.path.join(rfq_folder, "_GoogleDriveLink.txt")
                    with open(link_file_path, "w") as f:
                        f.write(drive_link)
                    print(f"✅ Created link file: {link_file_path}")

                # Create sentinel file to signal completion
                sentinel_file_path = os.path.join(rfq_folder, "_finished.txt")
                with open(sentinel_file_path, "w") as f:
                    pass  # This creates an empty file
                print(f"✅ Created sentinel file: {sentinel_file_path}")

            # Navigate back
            print("\n🔙 Navigating back to RFQ list...")
            cancel_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[title='Cancel']")))
            cancel_button.click()

            back_to_list_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[title='Back to List']")))
            back_to_list_button.click()

            wait.until(EC.visibility_of_element_located((By.CLASS_NAME, "list-tbody")))
            print("✅ Returned to RFQ list.")

            log_data["Status"] = "Success" if download_success else "Partial Success"
            if download_success:
                log_data["Progress/Comments"] = "Cycle completed successfully."

        except TimeoutException:
            print("\n⏹️ No more RFQs found. Exiting loop.")
            log_data["Status"] = "Idle"
            log_data["Progress/Comments"] = "No RFQs available."
            break
        except HttpError as e:
            log_data["Status"] = "Failed"
            log_data["Progress/Comments"] = f"Drive error: {e}"
            break
        except Exception as e:
            log_data["Status"] = "Failed"
            log_data["Progress/Comments"] = f"Error: {e}"
            print(f"\n❌ Unexpected error: {e}")
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

finally:
    print("\n🔚 Closing browser.")
    if "driver" in locals():
        driver.quit()