import os
import sys
import time
import re
import openpyxl
import pandas as pd
from dotenv import load_dotenv

# Ensure UTF-8 output in Windows console
if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8")

from telethon.sync import TelegramClient
from telethon.tl.functions.contacts import ImportContactsRequest
from telethon.tl.types import InputPhoneContact
from telethon.errors import FloodWaitError, UserPrivacyRestrictedError

import argparse

# Load credentials from .env file
load_dotenv()
try:
    API_ID = int(os.getenv("API_ID", "0"))
except ValueError:
    API_ID = 0
API_HASH = os.getenv("API_HASH", "")

# Batch & Safety Controls
# Adjust these defaults or override them with command line arguments:
# e.g.: python send_bulk.py --limit 20 --start 11 --delay 15
DEFAULT_LIMIT = 10         # Number of contacts to process in one run (0 = all 8500)
DEFAULT_START = 1          # 1-based contact index to start from
DEFAULT_DELAY = 20         # Delay in seconds between each message

# Outreach message containing your bot link
MESSAGE_TEXT = (
    "👋 Hello!\n\n"
    "Welcome to Infinite Tech. Looking for custom AI, automation solutions, or joining our tech community?\n\n"
    "Connect with our assistant on Telegram to get started:\n"
    "👉 https://t.me/InfiniteTech_LeadBot"
)


def clean_phone_number(raw_phone: str) -> str:
    """Sanitizes phone number by removing =, dots, dashes, spaces, and office extensions."""
    if not raw_phone:
        return ""
    text = str(raw_phone).strip().lower().lstrip("=")

    # Strip extensions like 'x376' or 'ext 227'
    text = re.split(r"[\s,]*(?:x|ext)\.?\s*\d+", text)[0]

    # Strip trailing Excel float notation like '.0'
    if text.endswith(".0") and text.count(".") == 1:
        text = text[:-2]

    # Remove all non-digit and non-plus characters
    digits = re.sub(r"[^0-9+]", "", text)

    if not digits:
        return ""

    if digits.startswith("+"):
        return digits

    # North American 10-digit number -> default to +1
    if len(digits) == 10:
        return "+1" + digits
    elif len(digits) == 11 and digits.startswith("1"):
        return "+" + digits
    elif len(digits) == 12 and digits.startswith("91"):
        return "+" + digits

    return "+" + digits


def get_contacts():
    """Loads contacts (Name, Phone) from verified leads file or original Excel sheets."""
    # Priority 1: Verified registered leads from Software sheet
    verified_file = "registered_telegram_software_leads.xlsx"
    if os.path.exists(verified_file):
        try:
            df = pd.read_excel(verified_file)
            if not df.empty and "Phone" in df.columns:
                contacts = []
                for _, row in df.iterrows():
                    name = str(row.get("Name", "Lead")).strip()
                    phone = str(row.get("Phone", "")).strip()
                    if phone:
                        contacts.append((name, phone))
                print(f"✅ Found {len(contacts)} verified Telegram contacts in '{verified_file}'")
                return contacts, verified_file
        except Exception as e:
            print(f"Error reading {verified_file}: {e}")

    # Priority 2: Full Excel workbooks
    excel_files = [
        "2US Calling - Arivudainambi Jul 2026.xlsx",
        "US Calling - Arivudainambi Jul 2026.xlsx",
    ]

    for file_name in excel_files:
        if os.path.exists(file_name):
            try:
                wb = openpyxl.load_workbook(file_name, read_only=True, data_only=True)
                target_sheet = None
                # Check for 'software' first, then 'canada'
                for sname in wb.sheetnames:
                    if "software" in sname.lower():
                        target_sheet = sname
                        break
                if not target_sheet:
                    for sname in wb.sheetnames:
                        if "canada" in sname.lower():
                            target_sheet = sname
                            break

                if target_sheet:
                    ws = wb[target_sheet]
                    rows = ws.iter_rows(values_only=True)
                    header = [str(c or "").strip().lower() for c in next(rows)]

                    phone_idx = None
                    fn_idx = None
                    ln_idx = None
                    name_idx = None

                    for idx, col in enumerate(header):
                        if "phone" in col or "mobile" in col:
                            phone_idx = idx
                        elif "firstname" in col or col == "first name":
                            fn_idx = idx
                        elif "lastname" in col or col == "last name":
                            ln_idx = idx
                        elif "agency" in col or "name" in col or "business" in col:
                            if name_idx is None:
                                name_idx = idx

                    if phone_idx is not None:
                        contacts = []
                        for r_idx, row in enumerate(rows, start=2):
                            raw_phone = row[phone_idx]
                            if raw_phone:
                                raw_str = str(raw_phone).strip()
                                if raw_str and "not provided" not in raw_str.lower():
                                    if fn_idx is not None and ln_idx is not None:
                                        first = str(row[fn_idx] or "").strip()
                                        last = str(row[ln_idx] or "").strip()
                                        biz = str(row[name_idx] or "").strip() if name_idx is not None else ""
                                        name_val = f"{first} {last}".strip() or biz or f"Lead {r_idx-1}"
                                    elif name_idx is not None:
                                        name_val = str(row[name_idx] or f"Lead {r_idx-1}").strip()
                                    else:
                                        name_val = f"Lead {r_idx-1}"
                                    contacts.append((name_val, raw_str))

                        print(f"✅ Found {len(contacts)} contacts in '{file_name}' (sheet: '{target_sheet}')")
                        return contacts, f"{file_name} [{target_sheet}]"
            except Exception as e:
                print(f"Error reading {file_name} with openpyxl: {e}")

    # Fallback to other files
    files_to_check = ["telegram phone number.xlsx", "phone_numbers.xlsx"]
    for file_name in files_to_check:
        if os.path.exists(file_name):
            try:
                df = pd.read_excel(file_name)
                phone_col = [c for c in df.columns if "phone" in str(c).lower() or "mobile" in str(c).lower() or "number" in str(c).lower()]
                name_col = [c for c in df.columns if "name" in str(c).lower()]

                if phone_col:
                    pcol = phone_col[0]
                    ncol = name_col[0] if name_col else None
                    contacts = []
                    for idx, row in df.iterrows():
                        p = row[pcol]
                        if pd.notna(p) and str(p).strip():
                            n = str(row[ncol]).strip() if ncol and pd.notna(row[ncol]) else f"Lead {idx+1}"
                            contacts.append((n, str(p).strip()))

                    print(f"✅ Found {len(contacts)} contacts in '{file_name}' (column: '{pcol}')")
                    return contacts, file_name
            except Exception as e:
                print(f"Error reading {file_name}: {e}")

    return [], None


def send_bulk_messages(limit=DEFAULT_LIMIT, start=DEFAULT_START, delay=DEFAULT_DELAY):
    if API_ID == 0 or not API_HASH:
        print("❌ Error: Please open your .env file and set your API_ID and API_HASH first!")
        return

    contacts, source_file = get_contacts()
    if not contacts:
        print("❌ Error: Could not find any contacts in the Excel files.")
        return

    total_contacts = len(contacts)
    start_idx = max(1, start) - 1
    if limit and limit > 0:
        end_idx = min(total_contacts, start_idx + limit)
    else:
        end_idx = total_contacts

    batch = contacts[start_idx:end_idx]

    print("==================================================")
    print(f"Loaded {total_contacts} contacts from '{source_file}'")
    print(f"Batch config: Processing {len(batch)} contacts (Index {start_idx + 1} to {end_idx})")
    print(f"Delay: {delay}s between messages")
    print("==================================================")

    if not batch:
        print("⚠️ No contacts found in the specified start/limit range.")
        return

    sent_count = 0
    not_registered_count = 0
    skipped_count = 0

    with TelegramClient("bulk_user_session", API_ID, API_HASH) as client:
        for i, (name, raw_phone) in enumerate(batch, start=1):
            overall_idx = start_idx + i
            phone = clean_phone_number(raw_phone)

            # Skip invalid phone numbers
            if not phone or len(phone) < 10:
                print(f"\n[{i}/{len(batch)}] (#{overall_idx} overall) ⚠️ Skipping invalid number: '{raw_phone}' ({name})")
                skipped_count += 1
                continue

            print(f"\n[{i}/{len(batch)}] (#{overall_idx} overall) Processing: {phone} | {name}")
            try:
                contact = InputPhoneContact(client_id=0, phone=phone, first_name=name[:30], last_name="")
                result = client(ImportContactsRequest([contact]))

                if result.users:
                    target_user = result.users[0]
                    client.send_message(target_user, MESSAGE_TEXT)
                    sent_count += 1
                    print(f"  ✅ Sent successfully to {phone} (@{target_user.username or target_user.first_name})")
                else:
                    not_registered_count += 1
                    print(f"  ⚠️ {phone} is not registered on Telegram.")

            except FloodWaitError as e:
                print(f"  ⏳ Telegram FloodWait: Must wait {e.seconds} seconds before sending again.")
                time.sleep(e.seconds)
            except UserPrivacyRestrictedError:
                print(f"  🔒 {phone} privacy settings do not allow receiving messages from non-contacts.")
            except Exception as e:
                print(f"  ❌ Failed to send to {phone}: {e}")

            if i < len(batch):
                print(f"  ⏳ Waiting {delay} seconds before next contact...")
                time.sleep(delay)

    print("\n==================================================")
    print(f"🎉 Bulk messaging batch completed!")
    print(f"• Successfully Sent: {sent_count}")
    print(f"• Not Registered: {not_registered_count}")
    print(f"• Skipped/Invalid: {skipped_count}")
    if end_idx < total_contacts:
        print(f"💡 Next batch can start from index {end_idx + 1}:")
        print(f"   python send_bulk.py --start {end_idx + 1} --limit {limit or 10}")
    print("==================================================")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Send bulk Telegram outreach messages.")
    parser.add_argument("--limit", type=int, default=DEFAULT_LIMIT, help=f"Max contacts to process in this run (default: {DEFAULT_LIMIT}, 0 for all)")
    parser.add_argument("--start", type=int, default=DEFAULT_START, help=f"Start from contact index (default: {DEFAULT_START})")
    parser.add_argument("--delay", type=int, default=DEFAULT_DELAY, help=f"Delay between messages in seconds (default: {DEFAULT_DELAY})")
    args = parser.parse_args()

    send_bulk_messages(limit=args.limit, start=args.start, delay=args.delay)
