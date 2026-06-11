import os
import re
import shutil

DOWNLOAD_DIR = os.path.join(os.getcwd(), "Downloaded_RFQs")
RFQ_N_DIR = os.path.join(DOWNLOAD_DIR, "RFQ_N")


def parse_title(full_title):
    """Extract (rfq_number, description) from a full title string."""
    # Format: "RFQ: event_XXXXX - RFQ # 1020433 Supply of spares"
    m = re.search(r'RFQ\s*#\s*(\d+)\s+(.+)', full_title)
    if m:
        return m.group(1).strip(), m.group(2).strip()

    # Format: "RFQ: event_XXXXX - 1061621 - Gloves Cotton Dotted, Brush Paint"
    m = re.search(r'event_\d+\s+-\s+(?:\S+\s+-\s+)?(\d+)\s+-\s+(.+)', full_title)
    if m:
        return m.group(1).strip(), m.group(2).strip()

    # Fallback: any standalone number in the title
    m = re.search(r'\b(\d{5,})\b', full_title)
    if m:
        number = m.group(1)
        desc = full_title.split(number, 1)[-1].strip().lstrip('-').strip()
        return number, desc or "N_A"

    return None, None


def safe_folder_name(text):
    return re.sub(r'[\\/*?:"<>|]', "", text).replace(" ", "_")


def fix_rfq_n():
    if not os.path.isdir(RFQ_N_DIR):
        print(f"Folder not found: {RFQ_N_DIR}")
        return

    subfolders = [
        f for f in os.listdir(RFQ_N_DIR)
        if os.path.isdir(os.path.join(RFQ_N_DIR, f))
    ]

    print(f"Found {len(subfolders)} folders inside RFQ_N\n")

    moved = 0
    skipped = 0

    for subfolder in subfolders:
        src = os.path.join(RFQ_N_DIR, subfolder)
        metadata_path = os.path.join(src, "_RFQ_Metadata.txt")

        # Extract timestamp from subfolder name e.g. "A_NA_20260309-083259"
        ts_match = re.search(r'(\d{8}-\d{6})$', subfolder)
        timestamp = ts_match.group(1) if ts_match else subfolder

        # Read metadata
        full_title = None
        if os.path.exists(metadata_path):
            with open(metadata_path, "r", encoding="utf-8") as f:
                for line in f:
                    if line.startswith("Full Title:"):
                        full_title = line.split("Full Title:", 1)[1].strip()
                        break

        if not full_title or full_title == "N/A":
            print(f"  [SKIP] {subfolder} — no parseable Full Title")
            skipped += 1
            continue

        rfq_number, description = parse_title(full_title)

        if not rfq_number:
            print(f"  [SKIP] {subfolder} — could not parse RFQ number from: {full_title}")
            skipped += 1
            continue

        safe_desc = safe_folder_name(description)
        new_folder_name = f"RFQ_{rfq_number}_-_{safe_desc}_{timestamp}"
        dest = os.path.join(DOWNLOAD_DIR, new_folder_name)

        if os.path.exists(dest):
            print(f"  [SKIP] {subfolder} — destination already exists: {new_folder_name}")
            skipped += 1
            continue

        shutil.move(src, dest)
        print(f"  [OK] {subfolder}")
        print(f"    -> {new_folder_name}")
        moved += 1

    print(f"\nDone. Moved: {moved}, Skipped: {skipped}")

    # Remove RFQ_N if now empty
    remaining = os.listdir(RFQ_N_DIR)
    if not remaining:
        os.rmdir(RFQ_N_DIR)
        print("Removed empty RFQ_N folder.")
    else:
        print(f"RFQ_N still has {len(remaining)} items — not removed.")


if __name__ == "__main__":
    fix_rfq_n()
