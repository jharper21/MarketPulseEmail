print("🟢 SCRIPT LOADED. IMPORTING LIBRARIES...") 

import os
import json
import smtplib
from datetime import datetime
import pytz 
import yfinance as yf
import pandas as pd
import gspread
from google import genai
from google.genai import types
import matplotlib.pyplot as plt
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import math
import io
import re

# --- CONFIGURATION ---
SHEET_NAME = "Portfolio_Master_DB"
AI_MODEL_NAME = 'gemini-2.5-pro' 

def get_sheet_data():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = json.loads(os.environ['GCP_SERVICE_ACCOUNT'])
    client = gspread.authorize(ServiceAccountCredentials.from_json_keyfile_dict(creds, scope))
    return client.open(SHEET_NAME)

def fetch_market_data(tickers):
    if not tickers: return pd.DataFrame()
    print(f"📡 Fetching data for: {tickers}")
    try:
        # Download 3 months to ensure valid Monthly calc
        data = yf.download(tickers, period="3mo", group_by='ticker', progress=False)
    except Exception as e:
        print(f"❌ Yahoo API Error: {e}")
        return pd.DataFrame()

    results = []
    for ticker in tickers:
        try:
            hist = data[ticker] if len(tickers) > 1 else data
            hist = hist.dropna(subset=['Close'])
            if hist.empty: continue

            current = float(hist['Close'].iloc[-1])
            prev = float(hist['Close'].iloc[-2])
            day_pct = ((current - prev) / prev) * 100
            
            # Monthly Calc (Approx 21 trading days)
            if len(hist) >= 22:
                month_start = float(hist['Close'].iloc[-22])
                month_pct = ((current - month_start) / month_start) * 100
            else:
                month_pct = 0.0

            results.append({
                "Ticker": ticker,
                "Price": current,
                "Day_Chg_Pct": day_pct,
                "Month_Chg_Pct": month_pct
            })
        except:
            continue
    return pd.DataFrame(results)

def get_ai_insights(port_df, watch_df, total_val, day_gain_dollar):
    try:
        print(f"🧠 Asking {AI_MODEL_NAME}...")
        client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])
        
        port_sorted = port_df.reindex(port_df.Day_Chg_Pct.abs().sort_values(ascending=False).index)
        watch_sorted = watch_df.reindex(watch_df.Day_Chg_Pct.abs().sort_values(ascending=False).index)
        
        p_str = port_sorted[['Ticker', 'Day_Chg_Pct', 'Month_Chg_Pct', 'Total_Gain_Loss']].head(15).to_string(index=False)
        w_str = watch_sorted[['Ticker', 'Day_Chg_Pct', 'Month_Chg_Pct']].head(15).to_string(index=False)
        
        prompt = f"""
        You are a Hedge Fund CIO.
        **STATUS:** Net Worth: ${total_val:,.0f} | Today's Move: ${day_gain_dollar:,.0f}
        **HOLDINGS:**
        {p_str}
        **WATCHLIST:**
        {w_str}
        
        **TASK:** Write a 3-bullet executive summary in pure HTML (no markdown code blocks).
        
        1. <b>The Why:</b> Analyze drivers (Macro vs Company).
        2. <b>The Risk:</b> Identify overextended positions (>20% Monthly).
        3. <b>The Hunt:</b> Flag Reversal setups in Watchlist.
        
        **CRITICAL:** End your response with a section titled "<br><b>🔗 Sources:</b>" followed by an HTML unordered list (<ul>) containing 1-2 direct links (<a href='...'>Article Title</a>) to the news used.
        """
        
        response = client.models.generate_content(
            model=AI_MODEL_NAME,
            contents=prompt,
            config=types.GenerateContentConfig(
                tools=[types.Tool(google_search=types.GoogleSearch())]
            )
        )
        
        text = response.text
        # CLEANER: Remove markdown code blocks if present
        text = text.replace("```html", "").replace("```", "")
        return text
        
    except Exception as e:
        return f"<i>AI Analysis Unavailable: {e}</i>"

def generate_chart(history_data):
    if len(history_data) < 2: return None
    
    # Reset to Standard Light Theme
    plt.style.use('default')
    
    df = pd.DataFrame(history_data)
    df['Date'] = pd.to_datetime(df['Date'])
    df = df.sort_values('Date').tail(30)
    
    start_val = float(df['Total_Value'].iloc[0])
    end_val = float(df['Total_Value'].iloc[-1])
    
    if start_val > 0:
        pct_change = ((end_val - start_val) / start_val) * 100
    else:
        pct_change = 0.0
    
    # Calculate actual date range for accurate title
    date_range_days = (df['Date'].iloc[-1] - df['Date'].iloc[0]).days
    if date_range_days > 0:
        title_prefix = f"{date_range_days}-Day Trend"
    else:
        title_prefix = "Trend"
        
    sign = "+" if pct_change >= 0 else ""
    title_str = f"{title_prefix} ({sign}{pct_change:.2f}%)"
    
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(df['Date'], df['Total_Value'], marker='.', color='#0052cc', linewidth=2)
    ax.fill_between(df['Date'], df['Total_Value'], alpha=0.1, color='#0052cc')
    
    title_color = 'green' if pct_change >= 0 else 'red'
    ax.set_title(title_str, fontsize=14, fontweight='bold', color=title_color)
    
    # FIX: Set Y-axis to actual data range with 5% padding for better granularity
    y_min = df['Total_Value'].min()
    y_max = df['Total_Value'].max()
    y_padding = (y_max - y_min) * 0.1  # 10% padding
    
    # Handle edge case where all values are the same
    if y_padding == 0:
        y_padding = y_min * 0.02  # 2% of value if no variation
    
    ax.set_ylim(y_min - y_padding, y_max + y_padding)
    
    # Format Y-axis with dollar signs and commas
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
    
    ax.grid(True, linestyle='--', alpha=0.3)
    plt.xticks(rotation=45)
    plt.tight_layout()
    
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=100)
    buf.seek(0)
    plt.close()
    return buf

def send_email(subject, body, img_buf):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = os.environ["GMAIL_USER"]
    msg['To'] = os.environ["GMAIL_USER"]
    msg.attach(MIMEText(body, 'html'))
    if img_buf:
        img = MIMEImage(img_buf.read())
        img.add_header('Content-ID', '<chart>')
        msg.attach(img)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as s:
        s.login(os.environ["GMAIL_USER"], os.environ["GMAIL_PASS"])
        s.send_message(msg)

def make_rows(df, cols, is_watch=False):
    rows = ""
    for _, r in df.iterrows():
        rows += "<tr>"
        for c in cols:
            val = r[c]
            if c == "Ticker":
                rows += f"<td style='text-align:left; font-weight:bold; padding:6px; border-bottom:1px solid #f0f0f0;'>{val}</td>"
            elif isinstance(val, (float, int)):
                if "Pct" in c:
                    color = "#27ae60" if val >= 0 else "#c0392b"
                    rows += f"<td style='color:{color}; text-align:right; padding:6px; border-bottom:1px solid #f0f0f0;'>{val:+.2f}%</td>"
                elif "Gain" in c:
                    color = "#27ae60" if val >= 0 else "#c0392b"
                    rows += f"<td style='color:{color}; text-align:right; padding:6px; border-bottom:1px solid #f0f0f0;'>{val:,.0f}</td>"
                else: 
                    rows += f"<td style='text-align:right; padding:6px; border-bottom:1px solid #f0f0f0;'>{val:,.2f}</td>"
            else:
                rows += f"<td style='text-align:right; padding:6px; border-bottom:1px solid #f0f0f0;'>{val}</td>"
        rows += "</tr>"
    return rows

def make_watchlist_rows(left_df, right_df, cols):
    """Generate watchlist rows with two columns side by side, separated by a vertical divider."""
    rows = ""
    max_len = max(len(left_df), len(right_df))
    
    left_list = left_df.to_dict('records')
    right_list = right_df.to_dict('records')
    
    for i in range(max_len):
        rows += "<tr>"
        
        # Left side columns
        if i < len(left_list):
            r = left_list[i]
            for c in cols:
                val = r[c]
                if c == "Ticker":
                    rows += f"<td style='text-align:left; font-weight:bold; padding:6px; border-bottom:1px solid #f0f0f0;'>{val}</td>"
                elif isinstance(val, (float, int)):
                    if "Pct" in c:
                        color = "#27ae60" if val >= 0 else "#c0392b"
                        rows += f"<td style='color:{color}; text-align:right; padding:6px; border-bottom:1px solid #f0f0f0;'>{val:+.2f}%</td>"
                    else:
                        rows += f"<td style='text-align:right; padding:6px; border-bottom:1px solid #f0f0f0;'>{val:,.2f}</td>"
                else:
                    rows += f"<td style='text-align:right; padding:6px; border-bottom:1px solid #f0f0f0;'>{val}</td>"
        else:
            # Empty cells for left side
            for _ in cols:
                rows += "<td style='padding:6px; border-bottom:1px solid #f0f0f0;'>&nbsp;</td>"
        
        # Vertical divider cell
        rows += "<td style='width:2px; background-color:#ccc; padding:0;'></td>"
        
        # Right side columns
        if i < len(right_list):
            r = right_list[i]
            for c in cols:
                val = r[c]
                if c == "Ticker":
                    rows += f"<td style='text-align:left; font-weight:bold; padding:6px; border-bottom:1px solid #f0f0f0;'>{val}</td>"
                elif isinstance(val, (float, int)):
                    if "Pct" in c:
                        color = "#27ae60" if val >= 0 else "#c0392b"
                        rows += f"<td style='color:{color}; text-align:right; padding:6px; border-bottom:1px solid #f0f0f0;'>{val:+.2f}%</td>"
                    else:
                        rows += f"<td style='text-align:right; padding:6px; border-bottom:1px solid #f0f0f0;'>{val:,.2f}</td>"
                else:
                    rows += f"<td style='text-align:right; padding:6px; border-bottom:1px solid #f0f0f0;'>{val}</td>"
        else:
            # Empty cells for right side
            for _ in cols:
                rows += "<td style='padding:6px; border-bottom:1px solid #f0f0f0;'>&nbsp;</td>"
        
        rows += "</tr>"
    
    return rows

def main():
    print("🚀 TO THE MOON INITIATED.")
    sh = get_sheet_data()
    port_ws = sh.worksheet("Portfolio")
    watch_ws = sh.worksheet("Watchlist")
    hist_ws = sh.worksheet("History_Log")
    
    port_df = pd.DataFrame(port_ws.get_all_records())
    watch_df = pd.DataFrame(watch_ws.get_all_records())
    
    port_df.columns = [c.replace(' ', '_') for c in port_df.columns]
    
    for col in ['Shares', 'Cost_Basis']:
        if col in port_df.columns:
            port_df[col] = pd.to_numeric(port_df[col], errors='coerce').fillna(0)
    
    if 'Cost_Basis' not in port_df.columns:
        port_df['Cost_Basis'] = 0.0

    all_tickers = list(set(list(port_df['Ticker']) + list(watch_df['Ticker'])))
    all_tickers = ['^VIX' if t == '.VIX' else t for t in all_tickers]

    market_df = fetch_market_data(all_tickers)
    if market_df.empty: return
    market_df['Ticker'] = market_df['Ticker'].replace('^VIX', '.VIX')

    # --- PORTFOLIO CALCS ---
    port_merged = port_df.merge(market_df, on="Ticker")
    
    # 1. Clean Data Types
    port_merged['Cost_Basis'] = port_merged['Cost_Basis'].astype(float)
    port_merged['Shares'] = port_merged['Shares'].astype(float)
    port_merged['Price'] = port_merged['Price'].astype(float)
    
    # 2. Fix Zero Cost Basis
    port_merged.loc[port_merged['Cost_Basis'] <= 0, 'Cost_Basis'] = port_merged['Price']
    
    # 3. Row Calcs
    port_merged['Value'] = port_merged['Shares'] * port_merged['Price']
    port_merged['Total_Gain_Loss'] = port_merged['Value'] - (port_merged['Shares'] * port_merged['Cost_Basis'])
    port_merged = port_merged.sort_values(by='Day_Chg_Pct', ascending=False)
    
    # --- ROBUST TOTAL CALCS ---
    total_val = port_merged['Value'].sum()
    total_cost = (port_merged['Shares'] * port_merged['Cost_Basis']).sum()
    
    # FIX: Sum the individual row Total_Gain_Loss values to match the displayed column
    total_gain_loss = port_merged['Total_Gain_Loss'].sum()
    
    # Total Gain % (Overall)
    total_gain_pct = ((total_val - total_cost) / total_cost * 100) if total_cost > 0 else 0
    
    # RECONSTRUCTED YESTERDAY'S VALUE (For accurate daily %)
    # Yesterday Value = Current Value / (1 + Day%/100)
    port_merged['Prev_Value'] = port_merged['Value'] / (1 + (port_merged['Day_Chg_Pct'] / 100))
    total_prev_val = port_merged['Prev_Value'].sum()
    
    # Day Dollar Gain (Actual Difference)
    day_gain_dollar = total_val - total_prev_val
    
    # Weighted Day % 
    day_change_pct = ((total_val - total_prev_val) / total_prev_val * 100) if total_prev_val > 0 else 0

    # RECONSTRUCTED MONTH AGO VALUE
    port_merged['Prev_Month_Value'] = port_merged['Value'] / (1 + (port_merged['Month_Chg_Pct'] / 100))
    total_prev_month_val = port_merged['Prev_Month_Value'].sum()
    
    # Weighted Month %
    total_month_pct = ((total_val - total_prev_month_val) / total_prev_month_val * 100) if total_prev_month_val > 0 else 0

    # --- WATCHLIST ---
    watch_merged = watch_df.merge(market_df, on="Ticker").sort_values(by='Day_Chg_Pct', ascending=False)
    
    mid_idx = math.ceil(len(watch_merged) / 2)
    watch_left = watch_merged.iloc[:mid_idx]
    watch_right = watch_merged.iloc[mid_idx:]

    tz = pytz.timezone('America/Los_Angeles')
    today = datetime.now(tz).strftime("%Y-%m-%d")
    hist_records = hist_ws.get_all_records()
    existing_dates = [str(r['Date']) for r in hist_records]
    
    if today in existing_dates:
        # Find the row number (add 2: 1 for header, 1 for 0-index)
        row_idx = existing_dates.index(today) + 2
        hist_ws.update(values=[[today, total_val, total_gain_loss]], range_name=f'A{row_idx}:C{row_idx}')
        print(f"📝 Updated existing row for {today}")
    else:
        hist_ws.append_row([today, total_val, total_gain_loss])
        print(f"📝 Added new row for {today}")

    ai_text = get_ai_insights(port_merged, watch_merged, total_val, day_gain_dollar)
    
    # --- HTML (Email-Safe with Inline Styles and Table-Based Layout) ---
    html = f"""
    <html>
    <head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    </head>
    <body style="font-family: Helvetica, Arial, sans-serif; color: #333; background-color: #ffffff; margin: 0; padding: 20px;">
        
        <!-- Header -->
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
                <td style="padding-bottom: 20px;">
                    <h2 style="margin: 0;">🚀 Market Pulse: ${total_val:,.0f} 
                        <span style="color:{'#27ae60' if day_gain_dollar>=0 else '#c0392b'}">({day_gain_dollar:+,.0f})</span>
                        <span style="font-size: 16px; color:{'#27ae60' if day_change_pct>=0 else '#c0392b'}"> {day_change_pct:+.2f}%</span>
                    </h2>
                </td>
            </tr>
        </table>
        
        <!-- AI Section -->
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 25px;">
            <tr>
                <td style="background-color: #f4f6f8; padding: 15px; border-left: 5px solid #2c3e50; border-radius: 4px;">
                    <h3 style="margin-top: 0; margin-bottom: 10px;">🧠 CIO Executive Summary</h3>
                    {ai_text}
                </td>
            </tr>
        </table>
        
        <!-- Main Content Container -->
        <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
                <!-- Holdings Box -->
                <td width="48%" valign="top" style="border: 1px solid #eee; padding: 15px; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
                    <h3 style="margin-top: 0;">💼 Holdings</h3>
                    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="font-size: 11px;">
                        <tr>
                            <th style="background: #f8f9fa; text-align: left; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Ticker</th>
                            <th style="background: #f8f9fa; text-align: right; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Price</th>
                            <th style="background: #f8f9fa; text-align: right; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Day %</th>
                            <th style="background: #f8f9fa; text-align: right; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Mth %</th>
                            <th style="background: #f8f9fa; text-align: right; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Total G/L</th>
                        </tr>
                        {make_rows(port_merged, ['Ticker', 'Price', 'Day_Chg_Pct', 'Month_Chg_Pct', 'Total_Gain_Loss'])}
                        <tr>
                            <td style="text-align: left; font-weight: bold; padding: 8px; border-top: 2px solid #ccc; background-color: #fafafa;">TOTAL</td>
                            <td style="text-align: right; font-weight: bold; padding: 8px; border-top: 2px solid #ccc; background-color: #fafafa;">-</td>
                            <td style="text-align: right; font-weight: bold; padding: 8px; border-top: 2px solid #ccc; background-color: #fafafa; color: {'#27ae60' if day_change_pct>=0 else '#c0392b'};">{day_change_pct:+.2f}%</td>
                            <td style="text-align: right; font-weight: bold; padding: 8px; border-top: 2px solid #ccc; background-color: #fafafa; color: {'#27ae60' if total_month_pct>=0 else '#c0392b'};">{total_month_pct:+.2f}%</td>
                            <td style="text-align: right; font-weight: bold; padding: 8px; border-top: 2px solid #ccc; background-color: #fafafa; color: {'#27ae60' if total_gain_loss>=0 else '#c0392b'};">{total_gain_loss:+,.0f}</td>
                        </tr>
                    </table>
                </td>
                
                <!-- Spacer -->
                <td width="4%">&nbsp;</td>
                
                <!-- Watchlist Box -->
                <td width="48%" valign="top" style="border: 1px solid #eee; padding: 15px; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.05);">
                    <h3 style="margin-top: 0;">👀 Watchlist</h3>
                    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="font-size: 11px;">
                        <tr>
                            <!-- Left Column Headers -->
                            <th style="background: #f8f9fa; text-align: left; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Ticker</th>
                            <th style="background: #f8f9fa; text-align: right; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Price</th>
                            <th style="background: #f8f9fa; text-align: right; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Day %</th>
                            <th style="background: #f8f9fa; text-align: right; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Mth %</th>
                            <!-- Divider -->
                            <th style="width: 2px; background-color: #ccc; padding: 0; border-bottom: 2px solid #ccc;"></th>
                            <!-- Right Column Headers -->
                            <th style="background: #f8f9fa; text-align: left; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Ticker</th>
                            <th style="background: #f8f9fa; text-align: right; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Price</th>
                            <th style="background: #f8f9fa; text-align: right; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Day %</th>
                            <th style="background: #f8f9fa; text-align: right; padding: 8px; color: #666; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #eee;">Mth %</th>
                        </tr>
                        {make_watchlist_rows(watch_left, watch_right, ['Ticker', 'Price', 'Day_Chg_Pct', 'Month_Chg_Pct'])}
                    </table>
                </td>
            </tr>
        </table>
        
        <!-- Chart Section -->
        <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top: 25px;">
            <tr>
                <td>
                    <h3>📅 30-Day Trend</h3>
                    <img src="cid:chart" style="width: 100%; max-width: 800px; border: 1px solid #ddd; border-radius: 4px;">
                </td>
            </tr>
        </table>
        
    </body>
    </html>
    """
    
    print("📧 Sending email...")
    hist_data = hist_ws.get_all_records()
    send_email(f"📊 Market Pulse: ${total_val:,.0f}", html, generate_chart(hist_data))
    print("✅ Done.")

if __name__ == "__main__":
    main()