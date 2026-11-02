# 🤖 Infinite Tech — Telegram Lead Generation Bot

> **Automate lead capture, bulk outreach, and Telegram registration scanning — all in one place.**

---

## 📋 Table of Contents

- [Project Overview](#-project-overview)
- [Folder Structure](#-folder-structure)
- [Scripts](#-scripts)
- [Prerequisites](#-prerequisites)
- [Installation](#-installation)
- [Configuration](#-configuration)
- [Usage](#-usage)
- [Google Sheets Integration](#-google-sheets-integration)
- [Security Notes](#-security-notes)

---

## 🌐 Project Overview

This repository contains **three Python scripts** that work together to power Infinite Tech's Telegram-based lead generation pipeline:

| Script | Purpose |
|---|---|
| `bot.py` | Conversational Telegram bot that captures lead info (name, phone, email, industry, pain points, desired AI solution) and logs it to Google Sheets + Excel |
| `send_bulk.py` | Bulk outreach tool — reads a lead list and sends a personalised Telegram message to every registered contact |
| `check_software_sheet_telegram.py` | Pre-screening scanner — checks which phone numbers from your Excel sheet are registered on Telegram, then saves verified leads to a file |

### Typical Workflow

```
check_software_sheet_telegram.py  →  registered_telegram_software_leads.xlsx
         ↓
send_bulk.py  (reads verified leads, sends bot link to each)
         ↓
bot.py  (user clicks link, goes through 6-step lead form)
         ↓
Google Sheet + user_responses.xlsx  (captured leads)
```

---

## 📁 Folder Structure

```
telegrambot/
│
├── bot.py                          # Main Telegram lead-gen bot
├── send_bulk.py                    # Bulk outreach via Telethon
├── check_software_sheet_telegram.py# Telegram registration scanner
│
├── requirements.txt                # Python dependencies
├── .env.example                    # Environment variable template (safe to commit)
├── .env                            # ⚠️ Real secrets — NEVER committed (gitignored)
│
├── data/                           # ⚠️ All Excel files live here at runtime (gitignored)
│   └── .gitkeep
│
├── logs/                           # ⚠️ Runtime log files (gitignored)
│   └── .gitkeep
│
├── .gitignore
└── README.md
```

> **Note:** Excel files (`.xlsx`) and Telegram session files (`.session`) are intentionally excluded from git via `.gitignore`. Place your data files in the `data/` folder at runtime.

---

## 📜 Scripts

### `bot.py` — Lead Generation Bot

A 6-step conversational Telegram bot built with **python-telegram-bot**.

**Conversation Flow:**
1. `/start` → Welcome message + Yes / Not Now buttons
2. ✅ Yes → Asks **Full Name**
3. → Asks **Mobile / WhatsApp Number**
4. → Asks **Work Email**
5. → Asks **Industry** (button menu or free text)
6. → Asks **Biggest Challenge** (button menu or free text)
7. → Asks **Desired AI Solution** (button menu or free text)
8. → Saves to Google Sheet + Excel → Shows Telegram community invite button

**Run:**
```bash
python bot.py
```

---

### `send_bulk.py` — Bulk Outreach

Reads a verified lead list and sends a personalised outreach message containing the bot link to every Telegram-registered contact.

**Source Priority:**
1. `data/registered_telegram_software_leads.xlsx` (pre-screened list)
2. `data/2US Calling - Arivudainambi Jul 2026.xlsx` → `Software` sheet
3. `data/US Calling - Arivudainambi Jul 2026.xlsx` → `Canada` sheet
4. `data/telegram phone number.xlsx`
5. `data/phone_numbers.xlsx`

**Run:**
```bash
# Default: 10 contacts, starting from index 1, 20s delay
python send_bulk.py

# Custom batch
python send_bulk.py --limit 50 --start 101 --delay 15
```

**Arguments:**

| Argument | Default | Description |
|---|---|---|
| `--limit` | `10` | Number of contacts to process (`0` = all) |
| `--start` | `1` | 1-based index to start from |
| `--delay` | `20` | Seconds between each message |

---

### `check_software_sheet_telegram.py` — Telegram Scanner

Scans all phone numbers in the `Software` sheet of your Excel file and identifies which ones have active Telegram accounts. Saves results to `data/registered_telegram_software_leads.xlsx`.

**Run:**
```bash
# Scan from the beginning
python check_software_sheet_telegram.py

# Resume from a specific index (e.g., resume after interruption at index 300)
python check_software_sheet_telegram.py 300
```

---

## ✅ Prerequisites

- Python **3.10+**
- A **Telegram Bot Token** from [@BotFather](https://t.me/BotFather)
- A **Telegram API ID & Hash** from [my.telegram.org](https://my.telegram.org) (for bulk scripts)
- A **Google Service Account JSON** key with Sheets + Drive access (for Google Sheets logging)

---

## 🛠️ Installation

```bash
# 1. Clone the repository
git clone <your-repo-url>
cd telegrambot

# 2. Create and activate a virtual environment
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt
```

---

## ⚙️ Configuration

```bash
# Copy the template
cp .env.example .env   # (or copy manually on Windows)
```

Open `.env` and fill in your real values:

```env
# Telegram Bot (from @BotFather)
BOT_TOKEN=your_bot_token_here

# Telegram Group Invite Link
GROUP_INVITE_LINK=https://t.me/+YourGroupInviteLink

# Telethon User API (from https://my.telegram.org)
API_ID=your_api_id_here
API_HASH=your_api_hash_here

# Google Sheets Integration
GOOGLE_SHEET_ID=your_google_sheet_id_here
GOOGLE_CREDENTIALS_FILE=your_service_account_credentials.json
```

> **Place your Google credentials JSON file** in the project root (it is gitignored via `*.json`).

---

## 📊 Google Sheets Integration

The bot logs every lead automatically to a Google Sheet.

**Setup:**
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project → Enable **Google Sheets API** and **Google Drive API**
3. Create a **Service Account** → Download the JSON key
4. Share your target Google Sheet with the service account email (Editor access)
5. Set `GOOGLE_SHEET_ID` and `GOOGLE_CREDENTIALS_FILE` in `.env`

**Sheet Columns (auto-created):**

| Timestamp | Telegram User ID | Username | Full Name | Phone Number | Email | Industry | Struggle Without AI | Desired AI Solution | Status |
|---|---|---|---|---|---|---|---|---|---|

---

## 🔒 Security Notes

> [!CAUTION]
> The following files contain **real credentials and sensitive data** — they are gitignored and must **NEVER** be committed to git:
> - `.env` — Bot token, API keys
> - `*.json` — Google service account credentials
> - `*.session` — Telegram user session (equivalent to full account access)
> - `*.xlsx` / `*.xls` — May contain PII (phone numbers, names, emails)

**If you accidentally commit secrets**, rotate them immediately:
- Revoke and regenerate your Telegram Bot Token via [@BotFather](https://t.me/BotFather)
- Reset your Telegram API Hash at [my.telegram.org](https://my.telegram.org)
- Revoke and re-create your Google Service Account key

---

## 📄 License

Internal use — Infinite Tech © 2026. All rights reserved.
