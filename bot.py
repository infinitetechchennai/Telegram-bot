import os
import logging
import sys
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    ConversationHandler,
    MessageHandler,
    ContextTypes,
    filters,
)
from telegram.request import HTTPXRequest

import gspread
from google.oauth2.service_account import Credentials

# Ensure UTF-8 output in Windows console
if sys.platform.startswith("win"):
    sys.stdout.reconfigure(encoding="utf-8")

# Load configuration from .env file
load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKEN", "8782835122:AAFM1XS0uZ7zV6j01aEB2Fh8susIQVOxA3w")
GROUP_INVITE_LINK = os.getenv("GROUP_INVITE_LINK", "https://t.me/+VBu0wcHeRF8yMjJl")
LOG_FILE = "user_responses.xlsx"
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "1oRwivp9yLL6dOPnwJZqk28xhKWLvK7-564HGN0NX8gU")
GOOGLE_CREDENTIALS_FILE = os.getenv("GOOGLE_CREDENTIALS_FILE", "ai-content-publisher-475010-3c397fab6a58.json")

# Configure logging to both console and bot.log
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
    handlers=[
        logging.FileHandler("bot.log", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)

# Google Sheets helper
def get_google_sheet():
    """Initializes and returns the Google Sheet worksheet."""
    try:
        if os.path.exists(GOOGLE_CREDENTIALS_FILE) and GOOGLE_SHEET_ID:
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ]
            creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS_FILE, scopes=scopes)
            client = gspread.authorize(creds)
            return client.open_by_key(GOOGLE_SHEET_ID).sheet1
    except Exception as e:
        logger.error(f"Google Sheets connection error: {e}")
    return None

# Conversation states
(
    CHOOSING_INTEREST,
    ASK_NAME,
    ASK_PHONE,
    ASK_EMAIL,
    ASK_INDUSTRY,
    ASK_STRUGGLE,
    ASK_SOLUTION,
) = range(7)


def log_response(
    user_id: int,
    username: str,
    full_name: str,
    status: str,
    phone: str = "N/A",
    email: str = "N/A",
    industry: str = "N/A",
    struggle: str = "N/A",
    solution: str = "N/A",
):
    """Save user interaction, contact info, pain points, and requirements into Google Sheets & Excel."""
    timestamp_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted_username = f"@{username}" if username else "N/A"
    
    # 1. Save directly to Google Sheet
    try:
        sheet = get_google_sheet()
        if sheet:
            sheet_row = [
                timestamp_str,
                str(user_id),
                formatted_username,
                full_name or "N/A",
                phone or "N/A",
                email or "N/A",
                industry or "N/A",
                struggle or "N/A",
                solution or "N/A",
                status,
            ]
            sheet.append_row(sheet_row)
            logger.info(f"✅ Logged to Google Sheet for user {user_id} ({full_name}): {status}")
        else:
            logger.warning("Google Sheet not accessible, skipping Google Sheet append.")
    except Exception as e:
        logger.error(f"Error appending to Google Sheet: {e}")

    # 2. Local Excel backup
    try:
        columns = [
            "Timestamp",
            "Telegram User ID",
            "Username",
            "Full Name",
            "Phone Number",
            "Email",
            "Industry",
            "Struggle Without AI",
            "Desired AI Solution",
            "Status",
        ]
        try:
            df = pd.read_excel(LOG_FILE)
            for col in columns:
                if col not in df.columns:
                    df[col] = "N/A"
        except FileNotFoundError:
            df = pd.DataFrame(columns=columns)

        new_entry = {
            "Timestamp": timestamp_str,
            "Telegram User ID": user_id,
            "Username": formatted_username,
            "Full Name": full_name or "N/A",
            "Phone Number": phone or "N/A",
            "Email": email or "N/A",
            "Industry": industry or "N/A",
            "Struggle Without AI": struggle or "N/A",
            "Desired AI Solution": solution or "N/A",
            "Status": status,
        }
        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        df.to_excel(LOG_FILE, index=False)
        logger.info(f"✅ Logged to local Excel for user {user_id} ({full_name}): {status}")
    except Exception as e:
        logger.error(f"Error saving response to local Excel: {e}")


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles the /start command when a user taps the bot link."""
    user = update.effective_user
    name = user.first_name if user.first_name else "there"
    logger.info(f"/start command received from {user.id} ({name})")

    keyboard = [
        [
            InlineKeyboardButton("✅ Yes, I'm Interested", callback_data="interested_yes"),
            InlineKeyboardButton("❌ Not Now", callback_data="interested_no"),
        ]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)

    welcome_text = (
        f"👋 Hello <b>{name}</b>!\n\n"
        f"Welcome to <b>Infinite Tech</b>! We build custom AI and automation solutions for growing businesses.\n\n"
        f"<b>Are you interested in exploring AI solutions for your business and joining our community?</b>"
    )

    if update.message:
        await update.message.reply_text(
            welcome_text,
            reply_markup=reply_markup,
            parse_mode="HTML",
        )
    elif update.callback_query:
        await update.callback_query.answer()
        await update.callback_query.message.reply_text(
            welcome_text,
            reply_markup=reply_markup,
            parse_mode="HTML",
        )
    return CHOOSING_INTEREST


async def start_lead_collection(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Step 1: Ask for Full Name."""
    query = update.callback_query
    await query.answer()
    user = update.effective_user
    logger.info(f"User {user.id} clicked 'interested_yes'")

    await query.edit_message_text(
        "🎉 <b>Awesome! Let's get your details to get started.</b>\n\n"
        "👤 <b>Step 1/6:</b> What is your <b>Full Name</b>?",
        parse_mode="HTML",
    )
    return ASK_NAME


async def handle_not_interested(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles 'Not Now' button click."""
    query = update.callback_query
    await query.answer()

    user = update.effective_user
    full_name = f"{user.first_name or ''} {user.last_name or ''}".strip()
    logger.info(f"User {user.id} clicked 'interested_no'")
    log_response(user.id, user.username, full_name, "NO")

    await query.edit_message_text(
        "👍 <b>No problem at all!</b>\n\n"
        "Thank you for your time. If you ever change your mind, you can type /start anytime.\n\n"
        "Have a wonderful day! 😊",
        parse_mode="HTML",
    )
    return ConversationHandler.END


async def get_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Step 2: Store Name and ask for Mobile Number."""
    name = update.message.text.strip()
    user = update.effective_user
    logger.info(f"get_name from {user.id}: {name}")
    context.user_data["lead_name"] = name

    await update.message.reply_text(
        f"Nice to meet you, <b>{name}</b>!\n\n"
        "📱 <b>Step 2/6:</b> What is your <b>Mobile / WhatsApp Number</b>?",
        parse_mode="HTML",
    )
    return ASK_PHONE


async def get_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Step 3: Store Phone and ask for Email."""
    phone = update.message.text.strip()
    user = update.effective_user
    logger.info(f"get_phone from {user.id}: {phone}")
    context.user_data["phone"] = phone

    await update.message.reply_text(
        "✉️ <b>Step 3/6:</b> What is your <b>Work / Business Email Address</b>?",
        parse_mode="HTML",
    )
    return ASK_EMAIL


def get_industry_markup():
    keyboard = [
        [
            InlineKeyboardButton("🏭 Manufacturing / Logistics", callback_data="ind_Manufacturing"),
            InlineKeyboardButton("🛒 E-Commerce / Retail", callback_data="ind_Ecommerce"),
        ],
        [
            InlineKeyboardButton("💻 IT / Software / SaaS", callback_data="ind_IT_Software"),
            InlineKeyboardButton("🏡 Real Estate / Construction", callback_data="ind_RealEstate"),
        ],
        [
            InlineKeyboardButton("🏥 Healthcare / Services", callback_data="ind_Healthcare"),
            InlineKeyboardButton("✍️ Other (Type Below)", callback_data="ind_other"),
        ],
    ]
    return InlineKeyboardMarkup(keyboard)


async def get_email(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Step 4: Store Email and ask for Industry."""
    email = update.message.text.strip()
    user = update.effective_user
    logger.info(f"get_email from {user.id}: {email}")
    context.user_data["email"] = email

    await update.message.reply_text(
        "🏢 <b>Step 4/6: What industry is your business in?</b>\n"
        "<i>(Tap an option below or type your industry directly)</i>",
        reply_markup=get_industry_markup(),
        parse_mode="HTML",
    )
    return ASK_INDUSTRY


def get_struggle_markup():
    keyboard = [
        [
            InlineKeyboardButton("📉 Slow Lead Response / Lost Sales", callback_data="str_SlowLeads"),
        ],
        [
            InlineKeyboardButton("⏳ Repetitive Manual Work & High Staff Cost", callback_data="str_ManualWork"),
        ],
        [
            InlineKeyboardButton("💬 Customer Support Overload (Need 24/7)", callback_data="str_Support"),
        ],
        [
            InlineKeyboardButton("📊 Lack of Data Insights & Forecasting", callback_data="str_Data"),
        ],
        [
            InlineKeyboardButton("✍️ Other Challenge (Type Below)", callback_data="str_other"),
        ],
    ]
    return InlineKeyboardMarkup(keyboard)


async def proceed_to_struggle(send_func, industry_text: str, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["industry"] = industry_text
    await send_func(
        f"Industry noted: <b>{industry_text}</b>\n\n"
        "⚠️ <b>Step 5/6: What is your biggest sales or operational challenge right now without AI?</b>\n"
        "<i>(Tap a challenge below or type your own)</i>",
        reply_markup=get_struggle_markup(),
        parse_mode="HTML",
    )
    return ASK_STRUGGLE


async def handle_industry_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    labels = {
        "ind_Manufacturing": "Manufacturing / Logistics",
        "ind_Ecommerce": "E-Commerce / Retail",
        "ind_IT_Software": "IT / Software / SaaS",
        "ind_RealEstate": "Real Estate / Construction",
        "ind_Healthcare": "Healthcare / Services",
    }

    if data == "ind_other":
        await query.edit_message_text(
            "✍️ <b>Please type your industry or business domain below:</b>",
            parse_mode="HTML",
        )
        return ASK_INDUSTRY

    industry_name = labels.get(data, "Other")
    await query.edit_message_text(f"Selected Industry: <b>{industry_name}</b>", parse_mode="HTML")
    return await proceed_to_struggle(query.message.reply_text, industry_name, context)


async def handle_industry_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    industry_name = update.message.text.strip()
    return await proceed_to_struggle(update.message.reply_text, industry_name, context)


def get_solution_markup():
    keyboard = [
        [
            InlineKeyboardButton("🤖 24/7 AI Sales Agent & Chatbot", callback_data="sol_SalesBot"),
        ],
        [
            InlineKeyboardButton("⚡ Workflow & Process Automation", callback_data="sol_Automation"),
        ],
        [
            InlineKeyboardButton("🎯 Automated Lead Generation Engine", callback_data="sol_LeadGen"),
        ],
        [
            InlineKeyboardButton("🧠 Custom Enterprise AI / Knowledge Base", callback_data="sol_CustomAI"),
        ],
        [
            InlineKeyboardButton("✍️ Custom Requirement (Type Below)", callback_data="sol_other"),
        ],
    ]
    return InlineKeyboardMarkup(keyboard)


async def proceed_to_solution(send_func, struggle_text: str, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["struggle"] = struggle_text
    await send_func(
        "💡 <b>Step 6/6: What AI solution are you looking to implement?</b>\n"
        "<i>(Tap a solution below or type your specific requirement)</i>",
        reply_markup=get_solution_markup(),
        parse_mode="HTML",
    )
    return ASK_SOLUTION


async def handle_struggle_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    labels = {
        "str_SlowLeads": "Slow Lead Response / Lost Sales",
        "str_ManualWork": "Repetitive Manual Work & High Staff Cost",
        "str_Support": "Customer Support Overload (Need 24/7)",
        "str_Data": "Lack of Data Insights & Forecasting",
    }

    if data == "str_other":
        await query.edit_message_text(
            "✍️ <b>Please type your operational challenge or struggle below:</b>",
            parse_mode="HTML",
        )
        return ASK_STRUGGLE

    struggle_name = labels.get(data, "Other Challenge")
    await query.edit_message_text(f"Recorded Challenge: <b>{struggle_name}</b>", parse_mode="HTML")
    return await proceed_to_solution(query.message.reply_text, struggle_name, context)


async def handle_struggle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    struggle_name = update.message.text.strip()
    return await proceed_to_solution(update.message.reply_text, struggle_name, context)


async def finish_submission(send_func, solution_text: str, user, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["solution"] = solution_text

    lead_name = context.user_data.get("lead_name", user.first_name or "N/A")
    phone = context.user_data.get("phone", "N/A")
    email = context.user_data.get("email", "N/A")
    industry = context.user_data.get("industry", "N/A")
    struggle = context.user_data.get("struggle", "N/A")

    # Save to Excel
    log_response(
        user_id=user.id,
        username=user.username,
        full_name=lead_name,
        status="YES",
        phone=phone,
        email=email,
        industry=industry,
        struggle=struggle,
        solution=solution_text,
    )

    summary_text = (
        f"🎉 <b>Thank you, {lead_name}! Your details have been submitted.</b>\n\n"
        f"📋 <b>Your Submitted Profile:</b>\n"
        f"• <b>Name:</b> {lead_name}\n"
        f"• <b>Mobile:</b> {phone}\n"
        f"• <b>Email:</b> {email}\n"
        f"• <b>Industry:</b> {industry}\n"
        f"• <b>Challenge without AI:</b> {struggle}\n"
        f"• <b>Desired Solution:</b> {solution_text}\n\n"
        f"Our AI Solutions Team will review your requirements and get back to you shortly!\n\n"
        f"👉 <b>Click the button below to join our Telegram Community:</b>"
    )

    keyboard = [
        [InlineKeyboardButton("👥 Join Infinite Tech AI Community", url=GROUP_INVITE_LINK)]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)

    await send_func(
        summary_text,
        reply_markup=reply_markup,
        parse_mode="HTML",
        disable_web_page_preview=False,
    )
    return ConversationHandler.END


async def handle_solution_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data

    labels = {
        "sol_SalesBot": "24/7 AI Sales Agent & Chatbot",
        "sol_Automation": "Workflow & Process Automation",
        "sol_LeadGen": "Automated Lead Generation Engine",
        "sol_CustomAI": "Custom Enterprise AI / Knowledge Base",
    }

    if data == "sol_other":
        await query.edit_message_text(
            "✍️ <b>Please type your AI solution requirement below:</b>",
            parse_mode="HTML",
        )
        return ASK_SOLUTION

    solution_name = labels.get(data, "Custom Requirement")
    await query.edit_message_text(f"Selected Solution: <b>{solution_name}</b>", parse_mode="HTML")
    return await finish_submission(query.message.reply_text, solution_name, update.effective_user, context)


async def handle_solution_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    solution_name = update.message.text.strip()
    return await finish_submission(update.message.reply_text, solution_name, update.effective_user, context)


async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Cancels the conversation."""
    await update.message.reply_text(
        "❌ Process cancelled. You can type /start anytime to begin again.",
        parse_mode="HTML",
    )
    return ConversationHandler.END


async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Log the error."""
    logger.error("Exception while handling an update:", exc_info=context.error)


def main():
    """Main function to launch the Telegram bot."""
    print("==================================================")
    print("InfiniteTechLeadGenBot is starting...")
    print(f"Invite Link configured: {GROUP_INVITE_LINK}")
    print(f"Responses will be saved to: {LOG_FILE}")
    print("==================================================")

    request = HTTPXRequest(
        connection_pool_size=10,
        read_timeout=30.0,
        write_timeout=30.0,
        connect_timeout=30.0,
        pool_timeout=30.0,
    )

    app = ApplicationBuilder().token(BOT_TOKEN).request(request).build()

    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            CallbackQueryHandler(start_lead_collection, pattern="^interested_yes$"),
        ],
        states={
            CHOOSING_INTEREST: [
                CallbackQueryHandler(start_lead_collection, pattern="^interested_yes$"),
                CallbackQueryHandler(handle_not_interested, pattern="^interested_no$"),
            ],
            ASK_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_name)],
            ASK_PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_phone)],
            ASK_EMAIL: [MessageHandler(filters.TEXT & ~filters.COMMAND, get_email)],
            ASK_INDUSTRY: [
                CallbackQueryHandler(handle_industry_button, pattern="^ind_"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_industry_text),
            ],
            ASK_STRUGGLE: [
                CallbackQueryHandler(handle_struggle_button, pattern="^str_"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_struggle_text),
            ],
            ASK_SOLUTION: [
                CallbackQueryHandler(handle_solution_button, pattern="^sol_"),
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_solution_text),
            ],
        },
        fallbacks=[CommandHandler("cancel", cancel), CommandHandler("start", start)],
        per_message=False,
    )

    app.add_handler(conv_handler)
    app.add_handler(CallbackQueryHandler(handle_not_interested, pattern="^interested_no$"))
    app.add_error_handler(error_handler)

    app.run_polling(bootstrap_retries=-1, timeout=30)


if __name__ == "__main__":
    main()
