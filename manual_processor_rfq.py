import os
import time
import re
import shutil
import mimetypes
import random
import zipfile
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError
# -*- coding: utf-8 -*-
import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Load environment variables
load_dotenv()

# --- CONFIGURATION ---
DOWNLOAD_DIR = os.path.join(os.getcwd(), "Downloaded_RFQs")
PARENT_FOLDER_ID = os.getenv("PARENT_FOLDER_ID")
SERVICE_ACCOUNT_FILE = "service_account.json"
SCOPES = ["https://www.googleapis.com/auth/drive"]

# Validate configuration
if not PARENT_FOLDER_ID or PARENT_FOLDER_ID.startswith("YOUR_"):
    raise ValueError("Error: PARENT_FOLDER_ID is not set correctly in .env")
if not os.path.exists(SERVICE_ACCOUNT_FILE):
    raise FileNotFoundError(f"Service account file not found: {SERVICE_ACCOUNT_FILE}")

os.makedirs(DOWNLOAD_DIR, exist_ok=True)


# --- HELPER FUNCTIONS ---

def extract_and_flatten(zip_path, dest_dir):
    """Extracts zip (and nested zips) into dest_dir, flattens single-folder nesting."""
    print(f"📂 Extracting: {os.path.basename(zip_path)}")
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(dest_dir)
    os.remove(zip_path)
    print(f"   ✅ Extracted and removed zip file")

    # Recursively handle inner zips
    for root, _, files in os.walk(dest_dir):
        for f in files:
            if f.lower().endswith(".zip"):
                inner_zip = os.path.join(root, f)
                inner_dest = os.path.splitext(inner_zip)[0]
                os.makedirs(inner_dest, exist_ok=True)
                print(f"   📦 Found nested zip: {f}")
                extract_and_flatten(inner_zip, inner_dest)

    # Flatten if only one subfolder exists
    items = os.listdir(dest_dir)
    if len(items) == 1:
        only_item = os.path.join(dest_dir, items[0])
        if os.path.isdir(only_item):
            print(f"   📁 Flattening single subfolder: {items[0]}")
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

        print(f"   📤 Uploading: {os.path.basename(local_path)}")
        attempts = 0
        while attempts < 5:
            try:
                file = service.files().create(
                    body=meta,
                    media_body=media,
                    fields="id, webViewLink",
                    supportsAllDrives=True
                ).execute()
                print(f"      ✅ Uploaded: {file['id']}")
                return file.get("id")
            except HttpError as e:
                code = getattr(e, "status_code", None) or getattr(e.resp, "status", None)
                print(f"      ⚠️ Upload failed (attempt {attempts+1}), code={code}")
                attempts += 1
                time.sleep((2 ** attempts) + random.uniform(0, 0.5))

        print(f"      ❌ Upload failed after {attempts} attempts: {os.path.basename(local_path)}")
        return None

    # Create root folder
    folder_name = os.path.basename(folder_path.rstrip(os.sep))
    print(f"\n☁️  Creating Google Drive folder: {folder_name}")
    root = _create_drive_folder(folder_name, parent_folder_id)
    root_id = root["id"]
    web_link = root.get("webViewLink", "")
    print(f"   ✅ Created folder: {root_id}")

    # Map local dir → Drive folder id
    dir_to_id = {os.path.abspath(folder_path): root_id}

    # Upload all files and subfolders
    for current_dir, subdirs, files in os.walk(folder_path):
        abs_dir = os.path.abspath(current_dir)
        parent_id = dir_to_id[abs_dir]

        # Create subfolders
        for sub in subdirs:
            sub_local = os.path.join(abs_dir, sub)
            created = _create_drive_folder(sub, parent_id)
            dir_to_id[os.path.abspath(sub_local)] = created["id"]
            print(f"   📁 Created subfolder: {sub}")

        # Upload files
        for fname in files:
            if fname.endswith(".crdownload") or fname.endswith(".part") or fname.startswith("_"):
                continue
            local_file = os.path.join(abs_dir, fname)
            _upload_file(local_file, parent_id)

    print(f"   ✅ Upload complete!")
    return web_link


def find_zip_files_in_directory(directory):
    """Find all zip files directly in the specified directory (not in subfolders)."""
    zip_files = []
    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        if os.path.isfile(item_path) and item.lower().endswith('.zip'):
            zip_files.append(item_path)
    return zip_files


def process_zip_file(zip_path, drive_service):
    """Process a single zip file: extract, upload to Drive, create link file."""
    print(f"\n{'='*70}")
    print(f"🎯 PROCESSING: {os.path.basename(zip_path)}")
    print(f"{'='*70}")
    
    # Extract RFQ title from filename or use generic name
    zip_filename = os.path.basename(zip_path)
    base_name = os.path.splitext(zip_filename)[0]
    
    # Create unique folder for this RFQ
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    safe_name = re.sub(r'[\\/*?:"<>|]', "", base_name).replace(" ", "_")
    unique_folder = f"RFQ_{safe_name}_{timestamp}"
    rfq_folder = os.path.join(DOWNLOAD_DIR, unique_folder)
    os.makedirs(rfq_folder, exist_ok=True)
    
    print(f"📁 Created folder: {unique_folder}")
    
    try:
        # Move zip to the new folder
        moved_zip = os.path.join(rfq_folder, zip_filename)
        shutil.move(zip_path, moved_zip)
        print(f"✅ Moved zip to folder")
        
        # Extract and flatten
        print(f"\n📦 EXTRACTING ZIP FILE...")
        extract_and_flatten(moved_zip, rfq_folder)
        
        # Count extracted files
        file_count = sum([len(files) for _, _, files in os.walk(rfq_folder)])
        print(f"✅ Extracted {file_count} file(s)")
        
        # Upload to Google Drive
        print(f"\n☁️  UPLOADING TO GOOGLE DRIVE...")
        drive_link = upload_folder_to_drive(drive_service, rfq_folder, PARENT_FOLDER_ID)
        
        if drive_link:
            print(f"\n🔗 Google Drive Link: {drive_link}")
            
            # Create link text file
            link_file_path = os.path.join(rfq_folder, "_GoogleDriveLink.txt")
            with open(link_file_path, "w") as f:
                f.write(drive_link)
            print(f"✅ Created link file: _GoogleDriveLink.txt")
            
            # Create sentinel file
            sentinel_file_path = os.path.join(rfq_folder, "_finished.txt")
            with open(sentinel_file_path, "w") as f:
                f.write(f"Processed: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Original file: {zip_filename}\n")
                f.write(f"Google Drive: {drive_link}\n")
            print(f"✅ Created sentinel file: _finished.txt")
        
        print(f"\n{'='*70}")
        print(f"✅ SUCCESSFULLY PROCESSED: {zip_filename}")
        print(f"{'='*70}")
        return True
        
    except Exception as e:
        print(f"\n❌ ERROR processing {zip_filename}: {e}")
        import traceback
        traceback.print_exc()
        return False


# --- MAIN EXECUTION ---

def main():
    print(f"\n{'='*70}")
    print(f"🚀 MANUAL DOWNLOAD PROCESSOR")
    print(f"{'='*70}")
    print(f"📂 Working directory: {DOWNLOAD_DIR}")
    print(f"☁️  Google Drive parent: {PARENT_FOLDER_ID}")
    print(f"{'='*70}\n")
    
    # Authenticate with Google Drive
    print("🔐 Authenticating with Google Drive...")
    try:
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        drive_service = build("drive", "v3", credentials=creds)
        print("✅ Google Drive authentication successful\n")
    except Exception as e:
        print(f"❌ Failed to authenticate with Google Drive: {e}")
        return
    
    # Find all zip files in the download directory
    zip_files = find_zip_files_in_directory(DOWNLOAD_DIR)
    
    if not zip_files:
        print("ℹ️  No zip files found in the directory")
        print(f"   Looking in: {DOWNLOAD_DIR}")
        print("\n💡 TIP: Place your manually downloaded zip files in this directory")
        return
    
    print(f"📦 Found {len(zip_files)} zip file(s) to process:\n")
    for idx, zf in enumerate(zip_files, 1):
        print(f"   {idx}. {os.path.basename(zf)}")
    
    # Process each zip file
    success_count = 0
    failure_count = 0
    
    for zip_path in zip_files:
        if process_zip_file(zip_path, drive_service):
            success_count += 1
        else:
            failure_count += 1
        print()  # Empty line between files
    
    # Summary
    print(f"\n{'='*70}")
    print(f"📊 PROCESSING SUMMARY")
    print(f"{'='*70}")
    print(f"✅ Successful: {success_count}")
    print(f"❌ Failed: {failure_count}")
    print(f"📁 Total processed: {len(zip_files)}")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Process interrupted by user")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        print("\n🔚 Script finished")