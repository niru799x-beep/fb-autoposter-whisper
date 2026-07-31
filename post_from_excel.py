"""
Facebook Auto-Poster — Excel Sheet Version
============================================
Content Calendar (Excel) পড়ে scheduled পোস্ট করে।

USAGE:
  python post_from_excel.py

ফোল্ডার structure:
  fb-autoposter/
  ├── post_from_excel.py
  ├── facebook_content_calendar.xlsx
  └── posts/
      ├── images/          ← সব ছবি এখানে
      │   ├── 001_morning.jpg
      │   └── 003_quote.png
      └── videos/          ← সব ভিডিও এখানে
          └── 002_reel.mp4

Excel-এ ফাইলের নাম কলামে লিখবে:
  images/001_morning.jpg
  videos/002_reel.mp4
  (খালি রাখলে text-only পোস্ট)
"""

import os
import sys
import requests
import openpyxl
from pathlib import Path
from datetime import datetime

# ── Config ────────────────────────────────────────────
PAGE_ID       = os.environ.get("FACEBOOK_PAGE_ID",      "YOUR_PAGE_ID")
ACCESS_TOKEN  = os.environ.get("FACEBOOK_ACCESS_TOKEN", "YOUR_TOKEN")
EXCEL_FILE    = Path("facebook_content_calendar.xlsx")
POSTS_DIR     = Path("posts")          # root — subfolder path আসবে Excel থেকে
SHEET_NAME    = "Content Calendar"

# Excel column index (1-based)
COL_ID        = 1
COL_FILENAME  = 2
COL_TYPE      = 3
COL_CAPTION   = 4
COL_SCHEDULE  = 5
COL_STATUS    = 6
COL_POST_ID   = 7
COL_NOTE      = 8

BASE_URL = f"https://graph.facebook.com/v19.0/{PAGE_ID}"

# ── Load Excel ────────────────────────────────────────

def load_sheet():
    if not EXCEL_FILE.exists():
        print(f"❌ Excel file পাওয়া যায়নি: {EXCEL_FILE}")
        sys.exit(1)
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb[SHEET_NAME]
    return wb, ws

def save_sheet(wb):
    wb.save(EXCEL_FILE)

# ── Find due posts ────────────────────────────────────

def get_due_rows(ws):
    now = datetime.now()
    due = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        status   = row[COL_STATUS - 1].value
        schedule = row[COL_SCHEDULE - 1].value

        if status and str(status).strip().lower() == "pending":
            if schedule and isinstance(schedule, datetime) and schedule <= now:
                due.append(row)
    return due

# ── Resolve file path ────────────────────────────────
# Excel-এ "images/001.jpg" বা "videos/002.mp4" লেখা থাকবে
# POSTS_DIR / filename = posts/images/001.jpg

def resolve_path(filename):
    """
    Excel-এ যা লেখা আছে (যেমন images/001.jpg বা videos/002.mp4)
    সেটা posts/ এর ভেতরে খোঁজো।
    """
    return POSTS_DIR / filename

# ── Post functions ────────────────────────────────────

def post_image(file_path, caption):
    print(f"  📸 Image পোস্ট করছি: {file_path.name}")
    with open(file_path, "rb") as f:
        resp = requests.post(
            f"{BASE_URL}/photos",
            params={"access_token": ACCESS_TOKEN},
            files={"source": f},
            data={"caption": caption or "", "published": "true"}
        ).json()
    if "id" in resp:
        return resp["id"]
    print(f"  ❌ Image error: {resp.get('error', {}).get('message', resp)}")
    return None

def post_video(file_path, caption):
    print(f"  🎬 Video পোস্ট করছি: {file_path.name}")
    file_size = file_path.stat().st_size

    # Step 1: Init
    init = requests.post(
        f"{BASE_URL}/video_reels",
        data={"upload_phase": "start", "access_token": ACCESS_TOKEN}
    ).json()
    if "error" in init:
        print(f"  ❌ Video init error: {init['error']['message']}")
        return None

    video_id   = init["video_id"]
    upload_url = init["upload_url"]

    # Step 2: Upload
    with open(file_path, "rb") as f:
        up = requests.post(
            upload_url,
            headers={
                "Authorization": f"OAuth {ACCESS_TOKEN}",
                "offset": "0",
                "file_size": str(file_size),
            },
            data=f
        ).json()
    if not up.get("success"):
        print(f"  ❌ Upload error: {up}")
        return None

    # Step 3: Publish
    pub = requests.post(
        f"{BASE_URL}/video_reels",
        data={
            "upload_phase": "finish",
            "access_token": ACCESS_TOKEN,
            "video_id": video_id,
            "video_state": "PUBLISHED",
            "description": caption or "",
        }
    ).json()
    if "error" in pub:
        print(f"  ❌ Publish error: {pub['error']['message']}")
        return None
    return video_id

def post_text(caption):
    print(f"  📝 Text পোস্ট করছি...")
    resp = requests.post(
        f"{BASE_URL}/feed",
        data={"message": caption or "", "access_token": ACCESS_TOKEN}
    ).json()
    if "id" in resp:
        return resp["id"]
    print(f"  ❌ Text error: {resp.get('error', {}).get('message', resp)}")
    return None

# ── Mark row as Done ──────────────────────────────────

def mark_done(ws, row, post_id):
    from openpyxl.styles import PatternFill, Font
    status_cell          = row[COL_STATUS - 1]
    post_id_cell         = row[COL_POST_ID - 1]
    status_cell.value    = "Done"
    status_cell.fill     = PatternFill("solid", fgColor="D4EDDA")
    status_cell.font     = Font(name="Arial", size=10, bold=True, color="155724")
    post_id_cell.value   = str(post_id)

# ── Main ──────────────────────────────────────────────

def main():
    print("=" * 55)
    print("  Facebook Auto-Poster (Excel Mode)")
    print(f"  সময়: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print("=" * 55)

    wb, ws = load_sheet()
    due_rows = get_due_rows(ws)

    if not due_rows:
        print("ℹ️  এই মুহূর্তে কোনো scheduled পোস্ট নেই।")
        return

    print(f"\n📋 {len(due_rows)}টি পোস্ট পাওয়া গেছে।\n")
    posted_count = 0

    for row in due_rows:
        row_num   = row[0].row
        filename  = str(row[COL_FILENAME - 1].value or "").strip()
        post_type = str(row[COL_TYPE - 1].value or "").strip().lower()
        caption   = str(row[COL_CAPTION - 1].value or "").strip()
        schedule  = row[COL_SCHEDULE - 1].value

        print(f"── Row {row_num}: {filename or '(text only)'} [{post_type}]")

        post_id = None

        if post_type == "text" or not filename:
            post_id = post_text(caption)

        elif post_type == "image":
            file_path = resolve_path(filename)
            if not file_path.exists():
                print(f"  ⚠️  ফাইল নেই: {file_path}")
                print(f"      নিশ্চিত করো posts/images/ ফোল্ডারে আছে কিনা।")
                continue
            post_id = post_image(file_path, caption)

        elif post_type == "video":
            file_path = resolve_path(filename)
            if not file_path.exists():
                print(f"  ⚠️  ফাইল নেই: {file_path}")
                print(f"      নিশ্চিত করো posts/videos/ ফোল্ডারে আছে কিনা।")
                continue
            post_id = post_video(file_path, caption)

        else:
            print(f"  ⚠️  অচেনা ধরন: '{post_type}' — Image / Video / Text লিখতে হবে")
            continue

        if post_id:
            mark_done(ws, row, post_id)
            print(f"  ✅ সফল! Post ID: {post_id}")
            posted_count += 1
        else:
            print(f"  ❌ ব্যর্থ।")

    save_sheet(wb)
    print(f"\n{'='*55}")
    print(f"  ✅ {posted_count}/{len(due_rows)} পোস্ট সফল।")
    print(f"  Excel আপডেট হয়েছে।")
    print(f"{'='*55}")

if __name__ == "__main__":
    main()
