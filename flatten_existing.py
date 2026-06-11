# -*- coding: utf-8 -*-
import os
import shutil

DOWNLOAD_DIR = os.path.join(os.path.dirname(__file__), "Downloaded_RFQs")
_JUNK_NAMES  = {"__macosx", "thumbs.db", ".ds_store"}

def move_or_merge(src, dst_dir):
    """Copy src into dst_dir (merging dirs), then delete source. Skips locked files."""
    dst = os.path.join(dst_dir, os.path.basename(src))
    if os.path.isdir(src):
        os.makedirs(dst, exist_ok=True)
        for child in os.listdir(src):
            move_or_merge(os.path.join(src, child), dst)
        try:
            os.rmdir(src)   # only succeeds if now empty
        except OSError:
            pass
    else:
        if not os.path.exists(dst):
            try:
                shutil.copy2(src, dst)
                os.remove(src)
            except PermissionError as e:
                print(f"    LOCKED (skipped): {src} — {e}")
        else:
            # Destination already exists — discard duplicate source
            try:
                os.remove(src)
            except PermissionError:
                pass


def flatten_folder(dest_dir):
    """Repeatedly collapse single-child wrapper dirs, ignoring _ files and junk."""
    rounds = 0
    changed = True
    while changed:
        changed = False
        items = os.listdir(dest_dir)
        real  = [i for i in items
                 if i.lower() not in _JUNK_NAMES and not i.startswith('_')]
        if len(real) == 1:
            only_item = os.path.join(dest_dir, real[0])
            if os.path.isdir(only_item):
                for child in os.listdir(only_item):
                    move_or_merge(os.path.join(only_item, child), dest_dir)
                shutil.rmtree(only_item, ignore_errors=True)
                for name in items:
                    if name.lower() in _JUNK_NAMES:
                        p = os.path.join(dest_dir, name)
                        shutil.rmtree(p) if os.path.isdir(p) else os.remove(p)
                rounds += 1
                changed = True
    return rounds


rfq_folders = [
    os.path.join(DOWNLOAD_DIR, d)
    for d in os.listdir(DOWNLOAD_DIR)
    if os.path.isdir(os.path.join(DOWNLOAD_DIR, d))
]

skipped = flattened = already_flat = 0

for folder in sorted(rfq_folders):
    name = os.path.basename(folder)
    items = os.listdir(folder)
    real  = [i for i in items
             if i.lower() not in _JUNK_NAMES and not i.startswith('_')]

    if not real:
        print(f"  SKIP (empty content): {name}")
        skipped += 1
        continue

    rounds = flatten_folder(folder)
    if rounds:
        print(f"  FLATTENED ({rounds} wrapper(s) removed): {name}")
        flattened += 1
    else:
        already_flat += 1

print(f"\nDone. Flattened: {flattened}  |  Already flat: {already_flat}  |  Skipped: {skipped}")
