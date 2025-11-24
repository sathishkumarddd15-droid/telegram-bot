import os
import re
import logging
import pandas as pd
from flask import Flask, request
from telegram import Update
from telegram.ext import Application, MessageHandler, filters

# ---------------------------
# Logging setup
# ---------------------------
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

# ---------------------------
# Bot token and Excel config
# ---------------------------
TOKEN = os.getenv("TOKEN")
EXCEL_PATH = os.getenv("EXCEL_PATH", "PostShipment Master UpdatedFy25.xlsm")
SHEET_NAME = "Master Data"

# ---------------------------
# Currency conversion rates
# ---------------------------
CURRENCY_RATES = {
    "INR": 1.0,
    "USD": 83.0,
    "EUR": 90.0,
    "AED": 22.5,
    "GBP": 105.0
}

# ---------------------------
# Helpers
# ---------------------------
def load_df():
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    df.columns = df.columns.str.strip().str.upper()
    return df

def parse_leo_dates(df):
    col = 'LEO DATE'
    if col not in df.columns:
        raise KeyError("LEO DATE column not found")
    s = df[col]
    dt = pd.to_datetime(s, errors='coerce', dayfirst=True, infer_datetime_format=True)
    numeric_mask = s.apply(lambda x: isinstance(x, (int, float)))
    if numeric_mask.any():
        serials = s[numeric_mask].astype(float)
        serial_dt = pd.to_datetime('1899-12-30') + pd.to_timedelta(serials, unit='D')
        dt.loc[numeric_mask] = serial_dt
    df[col] = dt
    return df

def extract_month_and_country(text):
    parts = text.strip().split()
    cmd = parts[0].lower()
    m = re.match(r"^/leo(\d{1,2})$", cmd)
    if not m:
        return None, None
    month = int(m.group(1))
    if not (1 <= month <= 12):
        return None, None
    country = None
    if len(parts) > 1:
        country = " ".join(parts[1:]).upper()
    return month, country

def format_vehicle_table(summary, month, country=None):
    header = f"{'Sub Category':<22} {'Leo Qty':>10}"
    lines = [header, "-" * len(header)]
    total_qty = 0
    for _, row in summary.iterrows():
        sub = str(row['SUB INV TYPE']) if pd.notna(row['SUB INV TYPE']) else ""
        qty = int(row['QTY']) if pd.notna(row['QTY']) else 0
        total_qty += qty
        lines.append(f"{sub:<22} {qty:>10,}")
    lines.append("-" * len(header))
    lines.append(f"{'TOTAL':<22} {total_qty:>10,}")
    title = f"📊 LEO Month {month:02d} Vehicle Summary"
    if country: title += f" ({country})"
    return title + ":\n```\n" + "\n".join(lines[:60]) + "\n```"

def format_spares_table(summary, month, country=None):
    header = f"{'Sub Category':<22} {'Leo Value (Cr INR)':>20}"
    lines = [header, "-" * len(header)]
    total_val_in_cr = 0.0

    summary['CURRENCY'] = summary['CURRENCY'].fillna("INR").str.upper()
    summary['RATE'] = summary['CURRENCY'].map(CURRENCY_RATES).fillna(1.0)
    summary['VALUE_INR'] = summary['VALUE'] * summary['RATE']
    collapsed = summary.groupby('SUB INV TYPE')['VALUE_INR'].sum().reset_index()

    for _, row in collapsed.iterrows():
        sub = str(row['SUB INV TYPE']) if pd.notna(row['SUB INV TYPE']) else ""
        val_in_cr = row['VALUE_INR'] / 1e7
        total_val_in_cr += val_in_cr
        lines.append(f"{sub:<22} {val_in_cr:>15,.2f}")

    lines.append("-" * len(header))
    lines.append(f"{'TOTAL':<22} {total_val_in_cr:>15,.2f}")

    title = f"📊 LEO Month {month:02d} Spares Summary"
    if country: title += f" ({country})"
    return title + ":\n```\n" + "\n".join(lines[:60]) + "\n```"

# ---------------------------
# Command Handler
# ---------------------------
async def dynamic_leo(update: Update, context):
    try:
        command = update.message.text
        month, country = extract_month_and_country(command)
        if month is None:
            await update.message.reply_text("⚠️ Invalid format. Use /leoMM or /leoMM COUNTRY (e.g., /leo07 INDIA).")
            return

        df = load_df()
        df = parse_leo_dates(df)
        filtered = df[df['LEO DATE'].dt.month == month]

        if country:
            filtered = filtered[filtered['COUNTRY'].str.upper() == country]

        if filtered.empty:
            msg = f"ℹ️ No rows found for LEO month {month:02d}"
            if country: msg += f" in {country}"
            await update.message.reply_text(msg)
            return

        group_cols = ['INV TYPE', 'SUB INV TYPE']
        if 'CURRENCY' in filtered.columns:
            group_cols.append('CURRENCY')

        summary = (
            filtered.groupby(group_cols)
                    .agg({'QTY': 'sum', 'VALUE': 'sum'})
                    .reset_index()
        )

        vehicle_summary = summary[summary['INV TYPE'].str.contains("VEHICLE", case=False, na=False)]
        spares_summary = summary[summary['INV TYPE'].str.contains("SPARES", case=False, na=False)]

        if not vehicle_summary.empty:
            vehicle_summary = vehicle_summary.groupby('SUB INV TYPE').agg({'QTY': 'sum'}).reset_index()
            vehicle_text = format_vehicle_table(vehicle_summary, month, country)
            await update.message.reply_text(vehicle_text, parse_mode="Markdown")

        if not spares_summary.empty:
            spares_text = format_spares_table(spares_summary, month, country)
            await update.message.reply_text(spares_text, parse_mode="Markdown")

        if vehicle_summary.empty and spares_summary.empty:
            msg = f"ℹ️ No Vehicle or Spares data found for month {month:02d}"
            if country: msg += f" in {country}"
            await update.message.reply_text(msg)

    except Exception as e:
        logging.error(f"/leoMM error: {e}")
        await update.message.reply_text("⚠️ Error generating LEO summary.")

# ---------------------------
# Flask app for webhook
# ---------------------------
app = Flask(__name__)
application = Application.builder().token(TOKEN).build()
application.add_handler(MessageHandler(filters.Regex(r"^/leo\d{1,2}.*$"), dynamic_leo))

@app.route(f"/{TOKEN}", methods=["POST"])
def webhook():
    update = Update.de_json(request.get_json(force=True), application.bot)
    application.update_queue.put(update)
    return "OK", 200

@app.route("/")
def home():
    return "Bot is running!", 200

# ---------------------------
# Main entry
# ---------------------------
if __name__ == "__main__":
    # Set webhook URL (your Render service URL)
    url = os.getenv("RENDER_EXTERNAL_URL", "https://ib-bot-c33s.onrender.com")
    webhook_url = f"{url}/{TOKEN}"
    application.bot.set_webhook(webhook_url)
    app.run(host="0.0.0.0", port=10000)
