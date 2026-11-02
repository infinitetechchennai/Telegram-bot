import os
import sys
import time
import re
import random
import openpyxl
import pandas as pd
from dotenv import load_dotenv

if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8")

from telethon.sync import TelegramClient
from telethon.tl.functions.contacts import ImportContactsRequest, DeleteContactsRequest
from telethon.tl.types import InputPhoneContact
from telethon.errors import FloodWaitError

load_dotenv()
try:
    API_ID = int(os.getenv("API_ID", "0"))
except ValueError:
    API_ID = 0
API_HASH = os.getenv("API_HASH", "")

EXCEL_FILE = "2US Calling - Arivudainambi Jul 2026.xlsx"
SHEET_NAME = "Software"
BATCH_SIZE = 30
DELAY_BETWEEN_BATCHES = 3  # seconds between batch requests


def clean_phone_number(raw_phone: str) -> str:
    if not raw_phone:
        return ""
    text = str(raw_phone).strip().lower().lstrip("=")
    text = re.split(r"[\s,]*(?:x|ext)\.?\s*\d+", text)[0]
    if text.endswith(".0") and text.count(".") == 1:
        text = text[:-2]
    digits = re.sub(r"[^0-9+]", "", text)
    if not digits:
        return ""
    if digits.startswith("+"):
        return digits
    if len(digits) == 10:
        return "+1" + digits
    elif len(digits) == 11 and digits.startswith("1"):
        return "+" + digits
    return "+" + digits


def scan_all_numbers(start_idx=0):
    if API_ID == 0 or not API_HASH:
        print("❌ Error: API_ID / API_HASH not found in .env")
        return

    print("==================================================")
    print(f"📊 Loading '{SHEET_NAME}' sheet from '{EXCEL_FILE}'...")
    wb = openpyxl.load_workbook(EXCEL_FILE, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    rows = ws.iter_rows(values_only=True)
    header = [str(c or "").strip().lower() for c in next(rows)]

    phone_idx = header.index("area code and phone")
    fn_idx = header.index("firstname") if "firstname" in header else None
    ln_idx = header.index("lastname") if "lastname" in header else None
    biz_idx = header.index("business name") if "business name" in header else None

    # Collect unique phone numbers to avoid redundant checks
    unique_contacts = {}  # clean_phone -> dict of details
    total_rows = 0

    for r_idx, row in enumerate(rows, start=2):
        total_rows += 1
        raw_p = row[phone_idx]
        if raw_p:
            clean_p = clean_phone_number(raw_p)
            if clean_p and len(clean_p) >= 11:
                if clean_p not in unique_contacts:
                    first = str(row[fn_idx] or "").strip() if fn_idx is not None else ""
                    last = str(row[ln_idx] or "").strip() if ln_idx is not None else ""
                    biz = str(row[biz_idx] or "").strip() if biz_idx is not None else ""
                    name = f"{first} {last}".strip() or biz or f"Lead {r_idx}"
                    unique_contacts[clean_p] = {
                        "row": r_idx,
                        "name": name,
                        "business": biz,
                        "raw_phone": raw_p,
                        "clean_phone": clean_p
                    }

    print(f"✅ Total leads in sheet: {total_rows}")
    print(f"✅ Unique phone numbers to scan: {len(unique_contacts)}")
    print(f"🚀 Scanning in batches of {BATCH_SIZE} numbers...")
    print("==================================================")

    out_file = "registered_telegram_software_leads.xlsx"
    registered_leads = []
    if os.path.exists(out_file):
        try:
            old_df = pd.read_excel(out_file)
            registered_leads = old_df.to_dict("records")
            print(f"📂 Loaded {len(registered_leads)} previously found Telegram leads from '{out_file}'.")
        except Exception:
            pass

    phone_items = list(unique_contacts.values())
    total_scanned = start_idx

    with TelegramClient("bulk_user_session", API_ID, API_HASH) as client:
        for b_idx in range(start_idx, len(phone_items), BATCH_SIZE):
            batch = phone_items[b_idx : b_idx + BATCH_SIZE]
            input_contacts = []

            for item in batch:
                input_contacts.append(
                    InputPhoneContact(
                        client_id=random.randint(1000, 999999),
                        phone=item["clean_phone"],
                        first_name=item["name"][:30],
                        last_name=""
                    )
                )

            try:
                res = client(ImportContactsRequest(input_contacts))
                found_in_batch = len(res.users)
                total_scanned += len(batch)

                if found_in_batch > 0:
                    for u in res.users:
                        matched = [item for item in batch if item["clean_phone"].endswith(str(u.phone or ""))]
                        m_item = matched[0] if matched else {}
                        lead_info = {
                            "Row": m_item.get("row", ""),
                            "Name": m_item.get("name", u.first_name),
                            "Business": m_item.get("business", ""),
                            "Phone": m_item.get("clean_phone", u.phone),
                            "Telegram_ID": u.id,
                            "Username": f"@{u.username}" if u.username else "",
                            "Telegram_Name": f"{u.first_name or ''} {u.last_name or ''}".strip()
                        }
                        registered_leads.append(lead_info)
                        print(f"  🎉 FOUND ON TELEGRAM! {lead_info['Phone']} | {lead_info['Name']} | {lead_info['Username']}")

                    try:
                        client(DeleteContactsRequest(id=res.users))
                    except Exception:
                        pass

                pct = int((total_scanned / len(phone_items)) * 100)
                print(f"[{total_scanned}/{len(phone_items)}] ({pct}%) Checked... Found so far: {len(registered_leads)}")

            except FloodWaitError as e:
                print(f"\n⏳ Telegram FloodWait: Must pause {e.seconds}s before continuing...")
                time.sleep(e.seconds + 2)
                # Retry this batch
                try:
                    res = client(ImportContactsRequest(input_contacts))
                    total_scanned += len(batch)
                    if res.users:
                        for u in res.users:
                            matched = [item for item in batch if item["clean_phone"].endswith(str(u.phone or ""))]
                            m_item = matched[0] if matched else {}
                            lead_info = {
                                "Row": m_item.get("row", ""),
                                "Name": m_item.get("name", u.first_name),
                                "Business": m_item.get("business", ""),
                                "Phone": m_item.get("clean_phone", u.phone),
                                "Telegram_ID": u.id,
                                "Username": f"@{u.username}" if u.username else "",
                                "Telegram_Name": f"{u.first_name or ''} {u.last_name or ''}".strip()
                            }
                            registered_leads.append(lead_info)
                        client(DeleteContactsRequest(id=res.users))
                except Exception as retry_e:
                    print(f"⚠️ Retry error: {retry_e}")

            except Exception as e:
                print(f"⚠️ Error on batch: {e}")
                time.sleep(3)

            time.sleep(DELAY_BETWEEN_BATCHES)

    print("\n==================================================")
    print(f"🏁 100% SCAN COMPLETE ACROSS ENTIRE SOFTWARE SHEET!")
    print(f"• Total Unique Numbers Scanned: {total_scanned} / {len(phone_items)}")
    print(f"• Total Contacts on Telegram: {len(registered_leads)}")
    print("==================================================")

    if registered_leads:
        df = pd.DataFrame(registered_leads)
        df.drop_duplicates(subset=["Phone"], inplace=True)
        df.to_excel(out_file, index=False)
        print(f"💾 Saved {len(df)} active Telegram leads to '{out_file}'!")
    else:
        print("ℹ️ None of the scanned numbers have active Telegram accounts.")


if __name__ == "__main__":
    start_at = 0
    if len(sys.argv) > 1 and sys.argv[1].isdigit():
        start_at = int(sys.argv[1])
    scan_all_numbers(start_idx=start_at)
