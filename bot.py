import os
import pandas as pd
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, filters

TOKEN = os.getenv("TOKEN")
EXCEL_PATH = os.getenv("EXCEL_PATH", "PostShipment Master UpdatedFy25.xlsm")

async def start(update, context):
    await update.message.reply_text("✅ Bot is running... Waiting for Telegram commands.")

async def handle_message(update, context):
    user_text = update.message.text.strip()
    await update.message.reply_text(f"You said: {user_text}")

async def leo07(update, context):
    args = context.args
    if not args:
        await update.message.reply_text("Usage: /leo07 <Country>")
        return

    country = args[0].upper()

    try:
        df = pd.read_excel(EXCEL_PATH)
        country_col = "COUNTRY"
        value_col = "VALUE"

        if country_col in df.columns:
            filtered = df[df[country_col].astype(str).str.upper() == country]
            if filtered.empty:
                await update.message.reply_text(f"No data found for {country}")
            else:
                total = filtered[value_col].sum() if value_col in df.columns else len(filtered)
                await update.message.reply_text(f"Total VALUE for {country}: {total}")
        else:
            await update.message.reply_text(f"Excel file does not have a '{country_col}' column.")
    except Exception as e:
        await update.message.reply_text(f"Error reading Excel: {e}")

def main():
    application = ApplicationBuilder().token(TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("leo07", leo07))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    application.run_polling()

if __name__ == "__main__":
    main()
