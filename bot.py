import os
import pandas as pd
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters

# Read token from environment variable
TOKEN = os.getenv("TOKEN")

# Path to your Excel file (can also be set via environment variable)
EXCEL_PATH = os.getenv("EXCEL_PATH", "PostShipment Master UpdatedFy25.xlsm")

# --- Handlers ---

async def start(update, context):
    await update.message.reply_text("✅ Bot is running... Waiting for Telegram commands.")

async def handle_message(update, context):
    user_text = update.message.text.strip()
    await update.message.reply_text(f"You said: {user_text}")

async def leo07(update, context):
    """Example command: /leo07 INDIA"""
    args = context.args
    if not args:
        await update.message.reply_text("Usage: /leo07 <Country>")
        return

    country = args[0].upper()

    try:
        # Load Excel file
        df = pd.read_excel(EXCEL_PATH)

        # Example: filter by country column
        if "Country" in df.columns:
            filtered = df[df["Country"].str.upper() == country]
            if filtered.empty:
                await update.message.reply_text(f"No data found for {country}")
            else:
                # Example: show totals
                total = filtered["Amount"].sum() if "Amount" in filtered.columns else len(filtered)
                await update.message.reply_text(f"Total for {country}: {total}")
        else:
            await update.message.reply_text("Excel file does not have a 'Country' column.")

    except Exception as e:
        await update.message.reply_text(f"Error reading Excel: {e}")

# --- Main ---

def main():
    application = ApplicationBuilder().token(TOKEN).build()

    # Command handlers
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("leo07", leo07))

    # Text handler
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    # Run bot
    application.run_polling()

if __name__ == "__main__":
    main()
