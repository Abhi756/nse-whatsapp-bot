# -*- coding: utf-8 -*-
import time
import pickle
from twilio.rest import Client

# === Twilio Config ===
ACCOUNT_SID = "your_account_sid_here"
AUTH_TOKEN = "your_auth_token_here"
FROM_WHATSAPP = "whatsapp:+14155238886"  # Twilio sandbox sender
TO_WHATSAPP = "whatsapp:+91XXXXXXXXXX"  # your WhatsApp number

client = Client(ACCOUNT_SID, AUTH_TOKEN)

def send_whatsapp_message(message):
    try:
        client.messages.create(
            body=message,
            from_=FROM_WHATSAPP,
            to=TO_WHATSAPP
        )
        print(f"WhatsApp message sent: {message}")
    except Exception as e:
        print(f"Failed to send WhatsApp message: {e}")

def safe_sum(df, col):
    return df[col].sum() if col is not None and col in df.columns else None

def get_interpretation(key, coi_sum, qty_sum):
    if coi_sum is None or qty_sum is None:
        return "⚪ No data"
    
    coi_pos = coi_sum > 0
    qty_pos = qty_sum > 0

    if key == "Max CE Buy":
        if coi_pos and qty_pos: return "🟢 Fresh Long Build-up → Bullish"
        if coi_pos and not qty_pos: return "🔴 Fresh Short Build-up → Bearish"
        if not coi_pos and qty_pos: return "🟢 Short Covering → Bullish short-term"
        if not coi_pos and not qty_pos: return "🔴 Long Unwinding → Bearish short-term"

    elif key == "Max CE Sell":
        if coi_pos and qty_pos: return "🔴 Fresh Short Build-up → Bearish"
        if coi_pos and not qty_pos: return "🟢 Fresh Long Build-up → Bullish"
        if not coi_pos and qty_pos: return "🔴 Long Unwinding → Bearish"
        if not coi_pos and not qty_pos: return "🟢 Short Covering → Bullish"

    elif key == "Max PE Buy":
        if coi_pos and qty_pos: return "🔴 Fresh Long Build-up → Bearish"
        if coi_pos and not qty_pos: return "🟢 Fresh Short Build-up → Bullish"
        if not coi_pos and qty_pos: return "🟢 Short Covering → Bullish"
        if not coi_pos and not qty_pos: return "🔴 Long Unwinding → Bearish"

    elif key == "Max PE Sell":
        if coi_pos and qty_pos: return "🟢 Fresh Short Build-up → Bullish"
        if coi_pos and not qty_pos: return "🔴 Fresh Long Build-up → Bearish"
        if not coi_pos and qty_pos: return "🟢 Long Unwinding → Bullish"
        if not coi_pos and not qty_pos: return "🔴 Short Covering → Bearish"

    return "⚪ Unknown"

while True:
    try:
        with open('prev_data.pkl', 'rb') as f:
            data = pickle.load(f)
    except Exception as e:
        print("Error loading pickle:", e)
        time.sleep(60)
        continue

    df_ce = data.get('CE')
    df_pe = data.get('PE')
    if df_ce is None or df_pe is None:
        print("'CE' or 'PE' not found.")
        time.sleep(60)
        continue

    # Quantity sums
    max_ce_buy = safe_sum(df_ce, 'Buy Change')
    max_pe_buy = safe_sum(df_pe, 'Buy Change')
    max_ce_sell = safe_sum(df_ce, 'Sell Change')
    max_pe_sell = safe_sum(df_pe, 'Sell Change')

    # Change in OI sums
    max_ce_coi = safe_sum(df_ce, 'changeinOpenInterest')
    max_pe_coi = safe_sum(df_pe, 'changeinOpenInterest')

    threshold = 100_000
    combo_triggered = False

    # ✅ Combo 1: CE Buy + PE Sell
    if max_ce_coi and max_pe_coi and max_ce_buy and max_pe_sell:
        if max_ce_coi > 0 and max_ce_buy > threshold and max_pe_coi > 0 and max_pe_sell > threshold:
            combo_msg = f"✅ Combo Triggered: CE Buy + PE Sell | CE Buy Qty: {max_ce_buy:,.0f}, OI: {max_ce_coi:,.0f} | PE Sell Qty: {max_pe_sell:,.0f}, OI: {max_pe_coi:,.0f}"
            send_whatsapp_message(combo_msg)
            combo_triggered = True

    # ✅ Combo 2: PE Buy + CE Sell
    if max_pe_coi and max_ce_coi and max_pe_buy and max_ce_sell:
        if max_pe_coi > 0 and max_pe_buy > threshold and max_ce_coi > 0 and max_ce_sell > threshold:
            combo_msg = f"✅ Combo Triggered: PE Buy + CE Sell | PE Buy Qty: {max_pe_buy:,.0f}, OI: {max_pe_coi:,.0f} | CE Sell Qty: {max_ce_sell:,.0f}, OI: {max_ce_coi:,.0f}"
            send_whatsapp_message(combo_msg)
            combo_triggered = True

    # 🚨 Only send top_key message if NO combo triggered
    if not combo_triggered:
        max_vals = {
            'Max CE Buy': max_ce_buy,
            'Max PE Buy': max_pe_buy,
            'Max CE Sell': max_ce_sell,
            'Max PE Sell': max_pe_sell
        }

        coi_sums = {
            'Max CE Buy': max_ce_coi,
            'Max PE Buy': max_pe_coi,
            'Max CE Sell': max_ce_coi,
            'Max PE Sell': max_pe_coi
        }

        valid_above_threshold = {k: v for k, v in max_vals.items() if v is not None and v > threshold}
        if valid_above_threshold:
            top_key = max(valid_above_threshold, key=valid_above_threshold.get)
            interpretation = get_interpretation(top_key, coi_sums[top_key], max_vals[top_key])

            message = f"{top_key}: {valid_above_threshold[top_key]:,.0f} | Change in OI: {coi_sums[top_key]:,.0f} | {interpretation}"
            send_whatsapp_message(message)
        else:
            print(f"No values above threshold of {threshold}.")

    time.sleep(60)

