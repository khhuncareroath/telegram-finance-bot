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

import sqlite3
import matplotlib.pyplot as plt
import os
from datetime import datetime

def init_db():
    conn = sqlite3.connect('finance.db')
    c = conn.cursor()
    # Existing tables
    c.execute('''CREATE TABLE IF NOT EXISTS expenses (id INTEGER PRIMARY KEY, user_id INTEGER, amount REAL, category TEXT, date TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS income (id INTEGER PRIMARY KEY, user_id INTEGER, amount REAL, source TEXT, date TEXT)''')
    
    # NEW: Premium feature tables
    c.execute('''CREATE TABLE IF NOT EXISTS budgets (user_id INTEGER, category TEXT, monthly_limit REAL, PRIMARY KEY(user_id, category))''')
    c.execute('''CREATE TABLE IF NOT EXISTS portfolio (id INTEGER PRIMARY KEY, user_id INTEGER, asset TEXT, amount REAL, buy_price REAL)''')
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
        import requests 
        
        # Translate "GOLD" to the PAXG crypto ticker
        if ticker.upper() == "GOLD":
            ticker = "PAXG"
            
        # Use CryptoCompare which does not block cloud servers
        url = f"https://min-api.cryptocompare.com/data/price?fsym={ticker.upper()}&tsyms=USD"
        response = requests.get(url).json()
        
        return float(response['USD'])
    except Exception as e:
        print(f"Error fetching {ticker}: {e}")
        return 0

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
    # A clean, modern dashboard layout using Markdown formatting
    dashboard_text = (
        "🏦 *PORTFOLIO MANAGER* \n"
        "━━━━━━━━━━━━━━━━━━\n\n"
        "Welcome back. All systems operational.\n\n"
        "📊 *Quick Snapshot*\n"
        "• *Today:* `$0.00`\n"
        "• *This Week:* `$0.00`\n\n"
        "⚡ *Quick Action*\n"
        "Type an amount and description to log it instantly.\n"
        "_Example: `3.50 Coffee`_"
    )
    
    # Modern 2x2 grid layout for your main buttons
    keyboard = [
        ["📊 Quick Summary", "📈 Crypto Portfolio"],
        ["🎯 Savings Goals", "⚙️ Settings & Tools"]
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    
    await update.message.reply_text(
        text=dashboard_text, 
        reply_markup=reply_markup, 
        parse_mode='Markdown'
    )

async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = query.from_user.id
    
    if data.startswith("exp_"):
        parts = data.split("_")
        amount = float(parts[1])
        category = parts[2]
        keyboard = [
            [
                InlineKeyboardButton("👤 Personal", callback_data=f"wal_exp_{amount}_{category}_Personal"),
                InlineKeyboardButton("💼 Business", callback_data=f"wal_exp_{amount}_{category}_Business")
            ],
            [
                InlineKeyboardButton("🏦 Savings", callback_data=f"wal_exp_{amount}_{category}_Savings")
            ]
        ]
        
        receipt_text = (
            "🧾 *TRANSACTION LOGGED*\n"
            "━━━━━━━━━━━━━━━━━━\n"
            f"💵 *Amount:* `${amount:,.2f}`\n"
            f"🏷️ *Category:* `{category}`\n"
            "📅 *Date:* `Today`\n\n"
            "Select the funding source below:"
        )
        await query.edit_message_text(text=receipt_text, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode='Markdown')
        return

    elif data.startswith("wal_exp_"):
        try:
            user_id = query.from_user.id
            parts = data.split("_")
            amount = float(parts[2])
            category = parts[3]
            wallet = parts[4]
            
            import sqlite3
            conn = sqlite3.connect('finance.db')
            cursor = conn.cursor()
            
            # Notepad-safe table creation
            cursor.execute(
                "CREATE TABLE IF NOT EXISTS transactions ("
                "id INTEGER PRIMARY KEY AUTOINCREMENT, "
                "user_id INTEGER, amount REAL, category TEXT, "
                "type TEXT, wallet TEXT, "
                "timestamp DATETIME DEFAULT CURRENT_TIMESTAMP)"
            )
            
            # Notepad-safe insert
            cursor.execute(
                "INSERT INTO transactions "
                "(user_id, amount, category, type, wallet) "
                "VALUES (?, ?, ?, 'expense', ?)", 
                (user_id, amount, category, wallet)
            )
            conn.commit()
            conn.close()
            
            success_msg = f"✅ *Success:* Logged `${amount:,.2f}` for {category} [*{wallet} Wallet*]."
            await query.edit_message_text(success_msg, parse_mode='Markdown')
            
        except Exception as e:
            await query.message.reply_text(f"⚠️ Button Error: {e}")
        return

    elif data.startswith("inc_"):
        amount_val = data.split("_")[1]
        keyboard = [
            [InlineKeyboardButton("👤 Personal", callback_data=f"wal_inc_{amount_val}_Personal"),
             InlineKeyboardButton("💼 Business", callback_data=f"wal_inc_{amount_val}_Business")],
            [InlineKeyboardButton("🏦 Savings", callback_data=f"wal_inc_{amount_val}_Savings")]
        ]
        await query.edit_message_text(f"Select Wallet for Income of ${float(amount_val):.2f}:", reply_markup=InlineKeyboardMarkup(keyboard))

    elif data.startswith("wal_inc_"):
        try:
            user_id = query.from_user.id
            parts = data.split("_")
            amount = float(parts[2])
            wallet = parts[3]
            
            import sqlite3
            conn = sqlite3.connect('finance.db')
            cursor = conn.cursor()
            
            cursor.execute(
                "CREATE TABLE IF NOT EXISTS transactions ("
                "id INTEGER PRIMARY KEY AUTOINCREMENT, "
                "user_id INTEGER, amount REAL, category TEXT, "
                "type TEXT, wallet TEXT, "
                "timestamp DATETIME DEFAULT CURRENT_TIMESTAMP)"
            )
            
            cursor.execute(
                "INSERT INTO transactions "
                "(user_id, amount, category, type, wallet) "
                "VALUES (?, ?, 'Income', 'income', ?)", 
                (user_id, amount, wallet)
            )
            conn.commit()
            conn.close()
            
            msg = f"✅ *Income Saved:* `${amount:,.2f}` to [*{wallet} Wallet*]."
            await query.edit_message_text(msg, parse_mode='Markdown')
            
        except Exception as e:
            await query.message.reply_text(f"⚠️ Income Error: {e}")
        return

    elif data.startswith("reset_"):
        try:
            user_id = query.from_user.id
            import sqlite3
            conn = sqlite3.connect('finance.db')
            cursor = conn.cursor()
            
            # Ensure tables exist to prevent invisible crashes
            cursor.execute("CREATE TABLE IF NOT EXISTS transactions (id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER, amount REAL, category TEXT, type TEXT, wallet TEXT, timestamp DATETIME DEFAULT CURRENT_TIMESTAMP)")
            cursor.execute("CREATE TABLE IF NOT EXISTS goals (id INTEGER PRIMARY KEY, user_id INTEGER, amount REAL, name TEXT)")
            cursor.execute("CREATE TABLE IF NOT EXISTS portfolio (id INTEGER PRIMARY KEY, user_id INTEGER, asset TEXT, amount REAL)")
            
            val = data.lower()
            
            # Look for abbreviations so we don't have to guess the exact name!
            if "exp" in val:
                cursor.execute("DELETE FROM transactions WHERE user_id = ? AND type = 'expense'", (user_id,))
                msg = "🗑️ *All Expenses Reset!*"
            elif "inc" in val:
                cursor.execute("DELETE FROM transactions WHERE user_id = ? AND type = 'income'", (user_id,))
                msg = "🗑️ *All Income Reset!*"
            elif "goal" in val:
                cursor.execute("DELETE FROM goals WHERE user_id = ?", (user_id,))
                msg = "🗑️ *All Goals Reset!*"
            elif "port" in val or "crypto" in val:
                cursor.execute("DELETE FROM portfolio WHERE user_id = ?", (user_id,))
                msg = "🗑️ *Portfolio Reset!*"
            else:
                # If it still doesn't match, safely print the name without crashing Telegram
                safe_data = data.replace("_", "-")
                msg = f"⚠️ Unknown button: {safe_data}"
                
            conn.commit()
            conn.close()
            await query.edit_message_text(msg, parse_mode='Markdown')
            
        except Exception as e:
            await query.message.reply_text(f"⚠️ Reset Error: {e}")
            return
        if data.startswith("buy_") and "_menu" in data:
            amount = data.split("_")[1]
            msg = f"📈 *Buy Asset:*\nYou chose to invest `${float(amount):,.2f}`.\nTo log this, reply with the coin name like this:\n`Buy {amount} BTC`"
            await query.edit_message_text(msg, parse_mode='Markdown')
            return

        elif data.startswith("sell_") and "_menu" in data:
            amount = data.split("_")[1]
            msg = f"📉 *Sell Asset:*\nYou chose to sell `${float(amount):,.2f}`.\nTo log this, reply with the coin name like this:\n`Sell {amount} BTC`"
            await query.edit_message_text(msg, parse_mode='Markdown')
            return

        elif data.startswith("goaladd_") and "_menu" in data:
            amount = data.split("_")[1]
            msg = f"🎯 *Save to Goal:*\nYou chose to save `${float(amount):,.2f}`.\nTo log this, reply with the goal name like this:\n`Goal {amount} Laptop`"
            await query.edit_message_text(msg, parse_mode='Markdown')
            return    
async def handle_clicks(update: Update, context: ContextTypes.DEFAULT_TYPE):
        user_id = update.effective_user.id
        text = update.message.text.strip()
        # 1. Show Goal Progress
    if text == "🎯 Savings Goals" or text == "Savings Goals":
        try:
            user_id = update.message.from_user.id
            import sqlite3
            conn = sqlite3.connect('finance.db')
            cursor = conn.cursor()
            
            # Anti-Crash: Ensure tables exist before trying to read them
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS transactions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER, amount REAL, category TEXT,
                    type TEXT, wallet TEXT,
                    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP)
            """)
            cursor.execute("CREATE TABLE IF NOT EXISTS goals (id INTEGER PRIMARY KEY, user_id INTEGER, amount REAL, name TEXT)")
            
            # Fetch Goals
            cursor.execute("SELECT name, amount FROM goals WHERE user_id = ?", (user_id,))
            goals = cursor.fetchall()
            
            # Calculate what is actually in the Savings Wallet
            cursor.execute("SELECT SUM(amount) FROM transactions WHERE user_id = ? AND type = 'income' AND wallet = 'Savings'", (user_id,))
            sav_in = cursor.fetchone()[0] or 0.0
            cursor.execute("SELECT SUM(amount) FROM transactions WHERE user_id = ? AND type = 'expense' AND wallet = 'Savings'", (user_id,))
            sav_out = cursor.fetchone()[0] or 0.0
            curr_savings = sav_in - sav_out
            conn.close()
            
            	if not goals:
                	await update.message.reply_text(
                    	f"🏦 *Savings Balance:* `${curr_savings:,.2f}`\n\n"
                    	"⚠️ You have no active goals.\n"
                    	"To add one, just type a message like this:\n`Goal 500 New Laptop`",
                    	parse_mode='Markdown'
                	)
                	return
                
            msg = f"🏦 *Savings Balance:* `${curr_savings:,.2f}`\n\n🎯 *Your Goals:*\n"
            for name, amount in goals:
                progress = min(100, (curr_savings / amount) * 100) if amount > 0 else 0
                msg += f"• *{name}*: `${amount:,.2f}` ({progress:.1f}% funded)\n"
                
            await update.message.reply_text(msg, parse_mode='Markdown')
        
        # 2. Command to Save a New Goal
        if text.lower().startswith("goal "):
            try:
                user_id = update.message.from_user.id
                parts = text.split(" ", 2)
                amount = float(parts[1])
                name = parts[2]
            
                import sqlite3
                conn = sqlite3.connect('finance.db')
                cursor = conn.cursor()
                cursor.execute("CREATE TABLE IF NOT EXISTS goals (id INTEGER PRIMARY KEY, user_id INTEGER, amount REAL, name TEXT)")
                cursor.execute("INSERT INTO goals (user_id, amount, name) VALUES (?, ?, ?)", (user_id, amount, name))
                conn.commit()
                conn.close()
            
                await update.message.reply_text(f"🎯 *Goal Saved:* `${amount:,.2f}` for *{name}*", parse_mode='Markdown')
            except Exception as e:
                await update.message.reply_text(f"⚠️ Error: {e}")
            return
        # Intercept the Crypto Portfolio button
    if text == "📈 Crypto Portfolio":
        loading_message = await update.message.reply_text("🔄 *Syncing with market data...*", parse_mode='Markdown')
        # (Your existing crypto logic will go here later!)
        return
    
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

    # Intercept the Quick Summary button
    if text == "📊 Quick Summary":
        import sqlite3
        try:
            conn = sqlite3.connect('finance.db')
            cursor = conn.cursor()
            
            # Calculate Total Income for this user
            cursor.execute("SELECT SUM(amount) FROM transactions WHERE user_id = ? AND type = 'income'", (user_id,))
            total_income = cursor.fetchone()[0] or 0.0
            
            # Calculate Total Expenses for this user
            cursor.execute("SELECT SUM(amount) FROM transactions WHERE user_id = ? AND type = 'expense'", (user_id,))
            total_expense = cursor.fetchone()[0] or 0.0
            
            conn.close()
            
            net_balance = total_income - total_expense
            
            summary_text = (
                "📊 *YOUR FINANCIAL REPORT*\n"
                "━━━━━━━━━━━━━━━━━━\n"
                f"📈 *Total Income:* `${total_income:,.2f}`\n"
                f"📉 *Total Spending:* `${total_expense:,.2f}`\n"
                "━━━━━━━━━━━━━━━━━━\n"
                f"💰 *Net Balance:* `${net_balance:,.2f}`"
            )
            await update.message.reply_text(summary_text, parse_mode='Markdown')
            
        except Exception as e:
            await update.message.reply_text("⚠️ *No data found.* Have you logged any transactions yet?", parse_mode='Markdown')
            print(f"Report error: {e}")
        return
        
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
            response = gemini_client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
            await update.message.reply_text(f"🧠 Consultant:\n\n{response.text}")
        except Exception as e:
            fallback_message = (
                "Oops! The AI assistant is taking a quick nap right now. 😴\n\n"
                "Don't worry, you can still log your finances! Just type an amount to get started.\n"
                "*Example:* `3.50`"
            )
            await update.message.reply_text(fallback_message, parse_mode='Markdown')

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

async def chart_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.chat_id
    current_month = datetime.now().strftime('%Y-%m')
    
    conn = sqlite3.connect('finance.db')
    c = conn.cursor()
    c.execute("SELECT category, SUM(amount) FROM expenses WHERE user_id=? AND date LIKE ? GROUP BY category", (user_id, f"{current_month}%"))
    data = c.fetchall()
    conn.close()

    if not data:
        await update.message.reply_text("No expenses yet this month to chart!")
        return

    # Generate Chart
    categories = [row[0] for row in data]
    amounts = [row[1] for row in data]
    
    plt.figure(figsize=(6, 6))
    plt.pie(amounts, labels=categories, autopct='%1.1f%%', startangle=140)
    plt.title(f"Expenses: {current_month}")
    
    # Save temporarily and send
    chart_path = "temp_chart.png"
    plt.savefig(chart_path)
    plt.close()
    
    with open(chart_path, 'rb') as photo:
        await update.message.reply_photo(photo=photo, caption="📊 Here is your visual breakdown for this month!")
    os.remove(chart_path)

async def check_budget_alert(user_id, category, current_expense):
    conn = sqlite3.connect('finance.db')
    c = conn.cursor()
    c.execute("SELECT monthly_limit FROM budgets WHERE user_id=? AND category=?", (user_id, category))
    limit_row = c.fetchone()
    
    if limit_row:
        limit = limit_row[0]
        current_month = datetime.now().strftime('%Y-%m')
        c.execute("SELECT SUM(amount) FROM expenses WHERE user_id=? AND category=? AND date LIKE ?", (user_id, category, f"{current_month}%"))
        total_spent = c.fetchone()[0] or 0
        
        conn.close()
        
        if total_spent + current_expense >= (limit * 0.8):
            return f"\n⚠️ *Warning:* You are approaching your monthly limit for {category}!"
    else:
        conn.close()
    return ""
async def monthly_report_job(context: ContextTypes.DEFAULT_TYPE):
    job = context.job
    user_id = job.chat_id
    
    # We will trigger a simple message for now, but later we can attach your PDF export here!
    await context.bot.send_message(chat_id=user_id, text="📅 Here is your automated monthly financial summary!")

async def portfolio_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.message.chat_id
    
    # Connect to the portfolio database table we created
    conn = sqlite3.connect('finance.db')
    c = conn.cursor()
    c.execute("SELECT asset, SUM(amount), AVG(buy_price) FROM portfolio WHERE user_id=? GROUP BY asset", (user_id,))
    assets = c.fetchall()
    conn.close()

    if not assets:
        await update.message.reply_text("Your portfolio is currently empty! Buy an asset first to track your PnL.")
        return

    message = "📈 *Your Advanced Portfolio & PnL:*\n\n"
    total_profit = 0

    for asset in assets:
        ticker, amount, avg_buy_price = asset
        current_price = get_live_price(ticker) 
        
        if current_price > 0:
            current_value = amount * current_price
            invested_value = amount * avg_buy_price
            pnl = current_value - invested_value
            pnl_percent = (pnl / invested_value) * 100 if invested_value > 0 else 0
            total_profit += pnl
            
            icon = "🟢" if pnl >= 0 else "🔴"
            message += f"*{ticker}* (Amount: {amount})\n"
            message += f"Avg Buy Price: ${avg_buy_price:.2f} | Current Price: ${current_price:.2f}\n"
            message += f"{icon} PnL: ${pnl:.2f} ({pnl_percent:.2f}%)\n\n"

    message += f"*Total Net PnL:* ${total_profit:.2f}"
    
    await update.message.reply_text(message, parse_mode='Markdown')

async def button_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    
# This single line is mandatory to stop the buttons from infinitely loading!
    await query.answer()
    # Catch the 3 wallet buttons
    if query.data in ['wallet_personal', 'wallet_business', 'wallet_savings']:
        wallet_choice = query.data.split('_')[1].capitalize()
        
        # We will just send a confirmation message for now
        await query.edit_message_text(text=f"✅ Successfully saved to your {wallet_choice} wallet!")
        return

    if query.data == 'action_guide':
        await guide_command(update, context)
    elif query.data == 'action_calendar':
        calendar, step = DetailedTelegramCalendar().build()
        await query.message.reply_text("Select a date:", reply_markup=calendar)
    elif query.data == 'action_stats':
        await query.message.reply_text("📊 *Your Stats:*\n\n*This Week:* $0.00\n*This Month:* $0.00", parse_mode='Markdown')
    elif query.data.startswith('cbcal'):
        # This strictly catches calendar clicks only so it doesn't crash
        result, key, step = DetailedTelegramCalendar().process(query.data)
        if not result and key:
            await query.answer()
            await query.message.edit_text(f"Select {LSTEP[step]}:", reply_markup=key)
        elif result:
            await query.answer()
            await query.message.edit_text(f"📅 You selected {result}.")
    else:
        # This sends all your original buttons (Meals, Buy Asset, etc.) to your old code
        await button_callback(update, context)

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
 # Start the automated monthly report timer
    # Replace YOUR_TELEGRAM_ID with your actual number ID if you want it to start immediately
    # app.job_queue.run_repeating(monthly_report_job, interval=2592000, first=10, chat_id=YOUR_TELEGRAM_ID)
   # 1. The New Handlers
    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("guide", guide_command))
    app.add_handler(CommandHandler("calendar", calendar_command))
    app.add_handler(CommandHandler("chart", chart_command))
    app.add_handler(CommandHandler("portfolio", portfolio_command))

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