"""
Telegram File ID Bot
=====================
এই script চালালে Bot তোমার পাঠানো ছবি/ভিডিওর File ID reply করবে।

USAGE:
  python telegram_file_id_bot.py

তারপর Telegram-এ Bot-এ ছবি/ভিডিও পাঠাও।
Bot File ID reply করবে — সেটা Excel-এ লিখো।

PC-তে একবার চালাতে হবে।
"""

import os
import time
import requests

TG_BOT_TOKEN = "8900286614:AAGToz-Q6cCqWMB32hpn_lAO4bOOu2iR_mI"
TG_API = f"https://api.telegram.org/bot{TG_BOT_TOKEN}"

print("=" * 50)
print("  Telegram File ID Bot চালু হয়েছে")
print("  Bot-এ ছবি বা ভিডিও পাঠাও")
print("  বন্ধ করতে Ctrl+C চাপো")
print("=" * 50)

last_update_id = 0

while True:
    try:
        resp = requests.get(
            f"{TG_API}/getUpdates",
            params={"offset": last_update_id + 1, "timeout": 30},
            timeout=35
        ).json()

        for update in resp.get("result", []):
            last_update_id = update["update_id"]
            msg = update.get("message", {})
            chat_id = msg.get("chat", {}).get("id")

            # ছবি
            if "photo" in msg:
                file_id = msg["photo"][-1]["file_id"]
                print(f"\n📸 ছবি পাওয়া গেছে!")
                print(f"   File ID: {file_id}")
                requests.post(f"{TG_API}/sendMessage", data={
                    "chat_id": chat_id,
                    "text": f"📸 Image File ID:\n`tg:{file_id}`\n\nExcel-এ ফাইলের নাম কলামে এটা লিখো।",
                    "parse_mode": "Markdown"
                })

            # ভিডিও
            elif "video" in msg:
                file_id = msg["video"]["file_id"]
                print(f"\n🎬 ভিডিও পাওয়া গেছে!")
                print(f"   File ID: {file_id}")
                requests.post(f"{TG_API}/sendMessage", data={
                    "chat_id": chat_id,
                    "text": f"🎬 Video File ID:\n`tg:{file_id}`\n\nExcel-এ ফাইলের নাম কলামে এটা লিখো।",
                    "parse_mode": "Markdown"
                })

            # Document (ভিডিও file হিসেবে পাঠালে)
            elif "document" in msg:
                file_id = msg["document"]["file_id"]
                fname = msg["document"].get("file_name", "")
                print(f"\n📄 Document পাওয়া গেছে: {fname}")
                print(f"   File ID: {file_id}")
                requests.post(f"{TG_API}/sendMessage", data={
                    "chat_id": chat_id,
                    "text": f"📄 File ID:\n`tg:{file_id}`\n\nExcel-এ ফাইলের নাম কলামে এটা লিখো।",
                    "parse_mode": "Markdown"
                })

    except KeyboardInterrupt:
        print("\nBot বন্ধ হয়েছে।")
        break
    except Exception as e:
        print(f"Error: {e}")
        time.sleep(5)
