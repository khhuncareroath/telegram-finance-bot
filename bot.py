from telegram import BotCommand, InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.ext import CommandHandler, CallbackQueryHandler, ContextTypes
from telegram_bot_calendar import DetailedTelegramCalendar, LSTEP
import sqlite3
import json
import urllib.request
import random
import os
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from google import genai
from telegram import Update, ReplyKeyboardMarkup, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, CallbackQueryHandler, filters, ContextTypes

BOT_TOKEN = os.environ.get("BOT_TOKEN", "YOUR_TELEGRAM_TOKEN_HERE")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "YOUR_GEMINI_API_KEY_HERE")

def init_db():
    conn = sqlite3.connect('finance.db')
    cursor = conn.cursor()
    # Added user_id to all tables
    cursor.execute('''CREATE TABLE IF NOT EXISTS transactions (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, amount REAL, category TEXT, type TEXT, wallet TEXT DEFAULT 'Personal')''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS goals (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, name TEXT, target REAL, current REAL DEFAULT 0, UNIQUE(user_id, name))''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS portfolio (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, ticker TEXT, amount REAL, avg_price REAL, portfolio_type TEXT DEFAULT 'Main', UNIQUE(user_id, ticker))''')
    conn.commit()
    conn.close()

def get_progress_bar(current, target):
    if target <= 0: return "[██████████] 100%"
    percent = min(current / target, 1.0)
    filled = int(percent * 10)
    bar = "█" * filled + "░" * (10 - filled)
    return f"`{bar}` {percent*100:.1f}%"

def get_live_price(ticker):
    try:
        url = f"https://api.binance.com/api/v3/ticker/price?symbol={ticker.upper()}USDT"
        req = urllib.request.urlopen(url)
        data = json.loads(req.read())
        return float(data['price'])
    except Exception:
        return 0.0

async def export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    conn = sqlite3.connect('finance.db')
    df = pd.read_sql_query("SELECT id, amount, category, type, wallet FROM transactions WHERE user_id = ?", conn, params=(user_id,))
    conn.close()
    
    if df.empty:
        await update.message.reply_text("❌ No transaction data found to export.")
        return
        
    file_path = "Financial_Report.xlsx"
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Transactions', index=False)
        
    await update.message.reply_document(document=open(file_path, 'rb'), filename="Financial_Report.xlsx", caption="📊 Structured Excel financial report!")

async def export_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    conn = sqlite3.connect('finance.db')
    cursor = conn.cursor()
    cursor.execute("SELECT SUM(amount) FROM transactions WHERE user_id = ? AND type = 'expense'", (user_id,))
    total_expense = cursor.fetchone()[0] or 0.0
    cursor.execute("SELECT SUM(amount) FROM transactions WHERE user_id = ? AND type = 'income'", (user_id,))
    total_income = cursor.fetchone()[0] or 0.0
    cursor.execute("SELECT amount, category, type, wallet FROM transactions WHERE user_id = ?", (user_id,))
    rows = cursor.fetchall()
    conn.close()
    
    file_path = "Financial_Summary.pdf"
    c = canvas.Canvas(file_path, pagesize=letter)
    width, height = letter
    
    c.setFillColorRGB(0.1, 0.2, 0.4)
    c.rect(0, height - 80, width, 80, fill=1, stroke=0)
    c.setFillColorRGB(1, 1, 1)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(50, height - 45, "Financial Management Report")
    
    c.setFillColorRGB(0, 0, 0)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 120, "Executive Summary")
    c.setFont("Helvetica", 11)
    c.drawString(50, height - 150, f"Total Income: ${total_income:,.2f}")
    c.drawString(50, height - 170, f"Total Expenses: ${total_expense:,.2f}")
    c.drawString(50, height - 190, f"Net Balance: ${total_income - total_expense:,.2f}")
    
    c.setFont("Helvetica-Bold", 14)
    c.drawString(50, height - 230, "Recent Transactions Ledger")
    
    y = height - 260
    c.setFont("Helvetica-Bold", 10)
    c.drawString(50, y, "Wallet")
    c.drawString(140, y, "Type")
    c.drawString(230, y, "Category")
    c.drawString(380, y, "Amount")
    
    y -= 20
    c.setFont("Helvetica", 10)
    for amount, category, t_type, wallet in rows[-15:]:
        if y < 50: break
        c.drawString(50, y, str(wallet))
        c.drawString(140, y, str(t_type))
        c.drawString(230, y, str(category))
        c.drawString(380, y, f"${amount:,.2f}")
        y -= 18
        
    c.save()
    await update.message.reply_document(document=open(file_path, 'rb'), filename="Financial_Summary.pdf", caption="📄 Styled PDF report!")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        ["📊 Quick Summary", "📈 Crypto & Trading"],
        ["🎯 Savings Goals", "🧠 Consultant & Settings"]
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    await update.message.reply_text("Welcome to your private Financial Tracker! 🚀\nJust type a number (e.g., `20`) to start logging.", reply_markup=reply_markup, parse_mode="Markdown")

async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = query.from_user.id
    
    if data.startswith("exp_"):
        parts = data.split("_")
        amount = parts[1]
        category = parts[2]
        keyboard = [
            [InlineKeyboardButton("👤 Personal", callback_data=f"wal_exp_{amount}_{category}_Personal"),
             InlineKeyboardButton("💼 Business", callback_data=f"wal_exp_{amount}_{category}_Business")],
            [InlineKeyboardButton("🏦 Savings", callback_data=f"wal_exp_{amount}_{category}_Savings")]
        ]
        await query.edit_message_text(f"Select Wallet for ${float(amount):.2f} ({category}):", reply_markup=InlineKeyboardMarkup(keyboard))
        
    elif data.startswith("wal_exp_"):
        parts = data.split("_")
        amount = float(parts[2])
        category = parts[3]
        wallet = parts[4]
        
        conn = sqlite3.connect('finance.db')
        conn.cursor().execute("INSERT INTO transactions (user_id, amount, category, type, wallet) VALUES (?, ?, ?, 'expense', ?)", (user_id, amount, category, wallet))
        conn.commit()
        conn.close()
        await query.edit_message_text(f"✅ Saved! Spent ${amount:.2f} on {category} [{wallet} Wallet].")

    elif data.startswith("inc_"):
        amount_val = data.split("_")[1]
        keyboard = [
            [InlineKeyboardButton("👤 Personal", callback_data=f"wal_inc_{amount_val}_Personal"),
             InlineKeyboardButton("💼 Business", callback_data=f"wal_inc_{amount_val}_Business")],
            [InlineKeyboardButton("🏦 Savings", callback_data=f"wal_inc_{amount_val}_Savings")]
        ]
        await query.edit_message_text(f"Select Wallet for Income of ${float(amount_val):.2f}:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data.startswith("wal_inc_"):
        parts = data.split("_")
        amount = float(parts[2])
        wallet = parts[3]
        conn = sqlite3.connect('finance.db')
        conn.cursor().execute("INSERT INTO transactions (user_id, amount, category, type, wallet) VALUES (?, ?, 'Income', 'income', ?)", (user_id, amount, wallet))
        conn.commit()
        conn.close()
        await query.edit_message_text(f"💰 Added ${amount:.2f} Income to [{wallet} Wallet].")

    elif data.startswith("buy_") and data.endswith("_menu"):
        amount = float(data.split("_")[1])
        keyboard = [
            [InlineKeyboardButton("₿ BTC", callback_data=f"buy_{amount}_BTC"), InlineKeyboardButton("♦️ ETH", callback_data=f"buy_{amount}_ETH")],
            [InlineKeyboardButton("🪙 PAXG", callback_data=f"buy_{amount}_PAXG"), InlineKeyboardButton("☀️ SOL", callback_data=f"buy_{amount}_SOL")]
        ]
        await query.edit_message_text(f"🪙 Asset to BUY for ${amount:.2f}:", reply_markup=InlineKeyboardMarkup(keyboard))
        
    elif data.startswith("buy_"):
        parts = data.split("_")
        invested_dollars = float(parts[1])
        ticker = parts[2]
        live_price = get_live_price(ticker)
        
        if live_price == 0:
            await query.edit_message_text(f"❌ Market price fetch failed for {ticker}.")
            return
            
        coin_amount = invested_dollars / live_price
        conn = sqlite3.connect('finance.db')
        cursor = conn.cursor()
        cursor.execute("SELECT amount, avg_price FROM portfolio WHERE user_id = ? AND ticker = ?", (user_id, ticker))
        row = cursor.fetchone()
        
        if row:
            old_amount, old_avg_price = row
            new_amount = old_amount + coin_amount
            new_avg_price = ((old_amount * old_avg_price) + invested_dollars) / new_amount
            cursor.execute("UPDATE portfolio SET amount = ?, avg_price = ? WHERE user_id = ? AND ticker = ?", (new_amount, new_avg_price, user_id, ticker))
        else:
            cursor.execute("INSERT INTO portfolio (user_id, ticker, amount, avg_price, portfolio_type) VALUES (?, ?, ?, ?, 'Main')", (user_id, ticker, coin_amount, live_price))
            
        cursor.execute("INSERT INTO transactions (user_id, amount, category, type, wallet) VALUES (?, ?, ?, 'expense', 'Personal')", (user_id, invested_dollars, f"Bought {ticker}"))
        conn.commit()
        conn.close()
        await query.edit_message_text(f"✅ Bag secured! Bought `${invested_dollars:.2f}` of {ticker}.")

    elif data.startswith("sell_") and data.endswith("_menu"):
        amount = float(data.split("_")[1])
        conn = sqlite3.connect('finance.db')
        cursor = conn.cursor()
        cursor.execute("SELECT ticker FROM portfolio WHERE user_id = ?", (user_id,))
        holdings = cursor.fetchall()
        conn.close()
        
        if not holdings:
            await query.edit_message_text("❌ No assets to sell.")
            return
            
        keyboard = [[InlineKeyboardButton(f"Sell {hold[0]}", callback_data=f"sell_{amount}_{hold[0]}")] for hold in holdings]
        await query.edit_message_text(f"🔴 Asset to SELL ${amount:.2f} of:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data.startswith("sell_"):
        parts = data.split("_")
        sell_dollars = float(parts[1])
        ticker = parts[2]
        live_price = get_live_price(ticker)
        
        if live_price == 0:
            await query.edit_message_text("❌ Market price fetch failed.")
            return
            
        coins_to_sell = sell_dollars / live_price
        conn = sqlite3.connect('finance.db')
        cursor = conn.cursor()
        cursor.execute("SELECT amount FROM portfolio WHERE user_id = ? AND ticker = ?", (user_id, ticker))
        row = cursor.fetchone()
        
        if not row or (row[0] * live_price) < (sell_dollars * 0.99):
            await query.edit_message_text(f"❌ Insufficient holdings.")
            conn.close()
            return
            
        new_amount = row[0] - coins_to_sell
        if new_amount < 0.000001:
            cursor.execute("DELETE FROM portfolio WHERE user_id = ? AND ticker = ?", (user_id, ticker))
        else:
            cursor.execute("UPDATE portfolio SET amount = ? WHERE user_id = ? AND ticker = ?", (new_amount, user_id, ticker))
            
        cursor.execute("INSERT INTO transactions (user_id, amount, category, type, wallet) VALUES (?, ?, ?, 'income', 'Personal')", (user_id, sell_dollars, f"Sold {ticker}"))
        conn.commit()
        conn.close()
        await query.edit_message_text(f"✅ Sold `${sell_dollars:.2f}` of {ticker}. Added to Income.")

    elif data.startswith("goaladd_") and data.endswith("_menu"):
        amount = float(data.split("_")[1])
        conn = sqlite3.connect('finance.db')
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM goals WHERE user_id = ?", (user_id,))
        goals = cursor.fetchall()
        conn.close()
        
        if not goals:
            await query.edit_message_text("❌ No goals found. Create one with `goal [name] [amount]`.")
            return
            
        keyboard = [[InlineKeyboardButton(f"🎯 Add to {g[0].capitalize()}", callback_data=f"goaladd_{amount}_{g[0]}")] for g in goals]
        await query.edit_message_text(f"🎯 Target goal for ${amount:.2f}:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data.startswith("goaladd_"):
        parts = data.split("_")
        amount = float(parts[1])
        goal_name = parts[2]
        conn = sqlite3.connect('finance.db')
        cursor = conn.cursor()
        cursor.execute("UPDATE goals SET current = current + ? WHERE user_id = ? AND name = ?", (amount, user_id, goal_name))
        cursor.execute("INSERT INTO transactions (user_id, amount, category, type, wallet) VALUES (?, ?, ?, 'expense', 'Savings')", (user_id, amount, f"Saved to {goal_name}"))
        conn.commit()
        conn.close()
        await query.edit_message_text(f"✅ Deposited ${amount:.2f} into {goal_name.capitalize()} fund!")

    elif data == "reset_exp":
        conn = sqlite3.connect('finance.db')
        conn.cursor().execute("DELETE FROM transactions WHERE user_id = ? AND type = 'expense'", (user_id,))
        conn.commit()
        conn.close()
        await query.edit_message_text("✅ Expenses reset!")
    elif data == "reset_inc":
        conn = sqlite3.connect('finance.db')
        conn.cursor().execute("DELETE FROM transactions WHERE user_id = ? AND type = 'income'", (user_id,))
        conn.commit()
        conn.close()
        await query.edit_message_text("✅ Income reset!")
    elif data == "reset_goals":
        conn = sqlite3.connect('finance.db')
        conn.cursor().execute("DELETE FROM goals WHERE user_id = ?", (user_id,))
        conn.commit()
        conn.close()
        await query.edit_message_text("✅ Goals reset!")
    elif data == "reset_portfolio":
        conn = sqlite3.connect('finance.db')
        conn.cursor().execute("DELETE FROM portfolio WHERE user_id = ?", (user_id,))
        conn.commit()
        conn.close()
        await query.edit_message_text("✅ Portfolio reset!")

async def handle_clicks(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    text = update.message.text.strip()
    text_lower = text.lower()
    
    clean_text = text.lstrip("+-").strip()
    first_word = clean_text.split()[0] if clean_text else ""
    is_number = False
    try:
        float(first_word)
        is_number = True
    except ValueError:
        is_number = False

    if is_number and len(clean_text.split()) == 1:
        amount = float(first_word)
        keyboard = [
            [InlineKeyboardButton("🍔 Meals", callback_data=f"exp_{amount}_Meals"), InlineKeyboardButton("🧋 Coffee", callback_data=f"exp_{amount}_Coffee")],
            [InlineKeyboardButton("🚕 Transport", callback_data=f"exp_{amount}_Transport"), InlineKeyboardButton("🛒 Convenience", callback_data=f"exp_{amount}_Convenience")],
            [InlineKeyboardButton("💰 Log as Income", callback_data=f"inc_{amount}")],
            [InlineKeyboardButton("🪙 Buy Asset", callback_data=f"buy_{amount}_menu"), InlineKeyboardButton("🔴 Sell Asset", callback_data=f"sell_{amount}_menu")],
            [InlineKeyboardButton("🎯 Save to Goal", callback_data=f"goaladd_{amount}_menu")]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await update.message.reply_text(f"Select action for **${amount:.2f}**:", reply_markup=reply_markup, parse_mode="Markdown")
        return

    if text == "📊 Quick Summary":
        conn = sqlite3.connect('finance.db')
        cursor = conn.cursor()
        cursor.execute("SELECT wallet, type, SUM(amount) FROM transactions WHERE user_id = ? GROUP BY wallet, type", (user_id,))
        data = cursor.fetchall()
        conn.close()
        
        wallets = {}
        for wallet, t_type, amt in data:
            if wallet not in wallets: wallets[wallet] = {'income': 0, 'expense': 0}
            wallets[wallet][t_type] += amt
            
        response = "📊 *Multi-Wallet Overview*\n\n"
        for w, totals in wallets.items():
            net = totals['income'] - totals['expense']
            response += f"📂 *{w} Wallet*\n"
            response += f"🟢 In: `${totals['income']:,.2f}` | 🔴 Out: `${totals['expense']:,.2f}`\n"
            response += f"💵 Net: `${net:,.2f}`\n\n"
        await update.message.reply_text(response, parse_mode="Markdown")
        
    elif text == "📈 Crypto & Trading":
        conn = sqlite3.connect('finance.db')
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT ticker, amount, avg_price, portfolio_type FROM portfolio WHERE user_id = ?", (user_id,))
            holdings = cursor.fetchall()
        except sqlite3.OperationalError:
            holdings = []
        conn.close()
        
        if not holdings:
            await update.message.reply_text("You have no assets! Type a number like `20` and select 'Buy Asset'.")
            return
            
        await update.message.reply_text("📡 Fetching live prices...")
        total_invested = 0
        total_current_value = 0
        response = "📈 *Your Trading Hub*\n\n"
        
        for ticker, amount, avg_price, port_type in holdings:
            live_price = get_live_price(ticker)
            invested = amount * avg_price
            current_value = amount * live_price
            total_invested += invested
            total_current_value += current_value
            pnl = current_value - invested
            pnl_percent = (pnl / invested) * 100 if invested > 0 else 0
            icon = "🟢" if pnl >= 0 else "🔴"
            
            response += f"*{ticker}* (Avg: `${avg_price:,.2f}`)\n"
            response += f"Invested: `${invested:,.2f}` | Value: `${current_value:,.2f}`\n"
            response += f"{icon} PnL: `${pnl:,.2f}` ({pnl_percent:+.2f}%)\n\n"
            
        total_pnl = total_current_value - total_invested
        total_icon = "🟢" if total_pnl >= 0 else "🔴"
        response += "-----------------------\n"
        response += f"💼 *Total Invested:* `${total_invested:,.2f}`\n"
        response += f"🏆 *Total Value:* `${total_current_value:,.2f}`\n"
        response += f"{total_icon} *Net PnL:* `${total_pnl:,.2f}`"
        await update.message.reply_text(response, parse_mode="Markdown")
        
    elif text == "🎯 Savings Goals":
        conn = sqlite3.connect('finance.db')
        cursor = conn.cursor()
        cursor.execute("SELECT name, current, target FROM goals WHERE user_id = ?", (user_id,))
        goals = cursor.fetchall()
        conn.close()
        
        if not goals:
            await update.message.reply_text("No goals yet! Create one: `goal ipad 500`", parse_mode="Markdown")
            return
            
        response = "🎯 *Savings Goals (Savings Wallet)*\n\n"
        for name, current, target in goals:
            bar = get_progress_bar(current, target)
            response += f"*{name.capitalize()}*: ${current:,.0f} /${target:,.0f}\n{bar}\n\n"
        await update.message.reply_text(response, parse_mode="Markdown")

    elif text in ["⚙️ Settings", "🧠 Consultant & Settings"]:
        keyboard = [
            [InlineKeyboardButton("🗑 Reset Expenses", callback_data="reset_exp"), InlineKeyboardButton("🗑 Reset Income", callback_data="reset_inc")],
            [InlineKeyboardButton("🗑 Reset Goals", callback_data="reset_goals"), InlineKeyboardButton("🗑 Reset Portfolio", callback_data="reset_portfolio")]
        ]
        await update.message.reply_text("💡 _Commands: `/export`, `/excel`, `/pdf`_\n\n⚙️ **Settings Controls:**", reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")
        
    elif text_lower.startswith("goal "):
        try:
            parts = text.split(maxsplit=2)
            conn = sqlite3.connect('finance.db')
            conn.cursor().execute("INSERT OR IGNORE INTO goals (user_id, name, target) VALUES (?, ?, ?)", (user_id, parts[1], float(parts[2])))
            conn.commit()
            conn.close()
            await update.message.reply_text(f"🎯 Goal created: {parts[1].capitalize()} for ${float(parts[2]):,.0f}!")
        except Exception:
            pass

    elif not is_number:
        if "YOUR_GEMINI_API_KEY_HERE" in GEMINI_API_KEY:
            await update.message.reply_text("🧠 Add your Gemini API key first!")
            return
            
        conn = sqlite3.connect('finance.db')
        cursor = conn.cursor()
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE user_id = ? AND type = 'expense'", (user_id,))
        exp = cursor.fetchone()[0] or 0.0
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE user_id = ? AND type = 'income'", (user_id,))
        inc = cursor.fetchone()[0] or 0.0
        conn.close()
        
        await update.message.reply_text("🧠 Thinking...")
        
        try:
            gemini_client = genai.Client(api_key=GEMINI_API_KEY)
            prompt = f"Financial advisor bot. Income: ${inc}, Expenses:${exp}. Client query: '{text}'. Keep it brief and practical."
            response = gemini_client.models.generate_content(model='gemini-2.0-flash', contents=prompt)
            await update.message.reply_text(f"🧠 Consultant:\n\n{response.text}")
        except Exception as e:
            await update.message.reply_text(f"❌ AI Error: {e}")

async def setup_bot_menu(application):
    commands = [
        BotCommand("start", "Open main menu"),
        BotCommand("calendar", "View calendar"),
        BotCommand("guide", "How to use the bot")
    ]
    await application.bot.set_my_commands(commands)

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("📅 Open Calendar", callback_data='action_calendar'),
         InlineKeyboardButton("📊 Weekly/Monthly Stats", callback_data='action_stats')],
        [InlineKeyboardButton("📖 Read Quick Guide", callback_data='action_guide')]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    welcome_text = "🚀 *Welcome to Manager Bot!*\nWhat would you like to check today?"
    await update.message.reply_text(welcome_text, reply_markup=reply_markup, parse_mode='Markdown')

async def guide_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    guide_text = "📖 *Quick Start Guide*\nLogging your daily spending is fast. Just type the amount followed by the description.\n*Examples:*\n🍜 `3.50 Noodle House`\n🚕 `4 PassApp`"
    await update.message.reply_text(guide_text, parse_mode='Markdown')

async def calendar_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    calendar, step = DetailedTelegramCalendar().build()
    await update.message.reply_text(f"Select {LSTEP[step]}:", reply_markup=calendar)

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    if query.data == 'action_guide':
        await guide_command(update, context)
    elif query.data == 'action_calendar':
        calendar, step = DetailedTelegramCalendar().build()
        await query.message.reply_text("Select a date:", reply_markup=calendar)
    else:
        result, key, step = DetailedTelegramCalendar().process(query.data)
        if not result and key:
            await query.message.edit_text(f"Select {LSTEP[step]}:", reply_markup=key)
        elif result:
            await query.message.edit_text(f"📅 You selected {result}. You spent $0 on this day.")

import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
import os

class SimpleHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"Bot is alive!")

def run_web_server():
    port = int(os.environ.get("PORT", 10000))
    server = HTTPServer(('0.0.0.0', port), SimpleHandler)
    server.serve_forever()

if __name__ == '__main__':
    init_db()
    
    # Start dummy web server so Render web service is happy
    t = threading.Thread(target=run_web_server)
    t.daemon = True
    t.start()
    
    app = ApplicationBuilder().token(BOT_TOKEN).connect_timeout(30.0).read_timeout(30.0).build()
    # 1. The New Handlers
    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("guide", guide_command))
    app.add_handler(CommandHandler("calendar", calendar_command))
    
    # 2. Your Existing Handlers (Kept exactly the same)
    app.add_handler(CommandHandler("export", export_data) if 'export_data' in globals() else CommandHandler("export", start_command))
    app.add_handler(CommandHandler("excel", export_excel))
    app.add_handler(CommandHandler("pdf", export_pdf))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_clicks))
    
    # 3. The New Button Handler & Menu Builder
    app.add_handler(CallbackQueryHandler(button_handler))
    app.post_init = setup_bot_menu
    print("Multi-User Multi-Wallet Bot is running!")
    app.run_polling()