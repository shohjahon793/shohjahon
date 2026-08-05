"""
LogiCheck - Accounting & Logistics Telegram Bot
Powered by Google Gemini + Google Sheets
"""

import os
import io
import json
import logging
from datetime import datetime, timedelta
import google.generativeai as genai
import gspread
from google.oauth2.service_account import Credentials

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters, ContextTypes
from telegram.constants import ParseMode

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Gemini setup ──────────────────────────────────────────────────────────────
genai.configure(api_key=os.environ["GEMINI_API_KEY"])
model = genai.GenerativeModel(
    model_name="gemini-1.5-flash",
    system_instruction="""You are "LogiCheck," an expert AI Accounting & Logistics Assistant for a motor carrier company. Your primary responsibility is to prevent costly accounting errors, flag payment holds, and assist with financial reconciliations.
Your communication style is concise, clear, and direct. Use simple Markdown (bold with *, bullet points).
When user mentions payment holds, deductions, missing docs, rate conflicts — acknowledge and categorize:
Labels: [CARRIER HOLD], [DRIVER DEDUCTION], [VENDOR ALERT], [MISSING DOCS], [RATE CONFLICT]
PRE-PAYMENT AUDIT TRIGGER: "processing payments", "running payouts", "starting settlements", "run audit", "audit now"
Keep responses short, bullet points, mobile-friendly. Never process payments yourself."""
)

# ── Google Sheets setup ───────────────────────────────────────────────────────
SHEET_ID = "1c3DpHm5KSJex93CvhfZFDnqZz0D-0LruTEMBvKEuyVA"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def get_sheets_client():
    creds_json = json.loads(os.environ["GOOGLE_CREDENTIALS"])
    creds = Credentials.from_service_account_info(creds_json, scopes=SCOPES)
    return gspread.authorize(creds)

def get_week_number(date: datetime) -> str:
    return f"W{date.isocalendar()[1]}"

def append_to_sheet(tab_name: str, row: list):
    try:
        gc = get_sheets_client()
        sh = gc.open_by_key(SHEET_ID)
        ws = sh.worksheet(tab_name)
        ws.append_row(row, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        logger.error(f"Sheets error: {e}")
        return False

# ── Type lists ────────────────────────────────────────────────────────────────
DEDUCTION_TYPES = [
    "Registration", "Scale", "Parking", "Layover", "Flight",
    "Truck wash", "Escrow back", "Travel expense", "Toll", "Dispatch fee",
    "Cash advance", "DOT", "Fridge", "Oregon permit", "Truck repair",
    "Driver charge", "Fuel", "Other"
]

REIMBURSEMENT_TYPES = [
    "Fridge", "Trip expense", "Cleaning supplies", "Cash advance",
    "Fleet service", "Windshield wipers", "Truck repair", "Scale",
    "Detention", "T-handle", "Fax", "Coolant", "Fuel", "Straps",
    "Parking", "Tire gauge", "Maintenance", "Taxi", "Bonus",
    "Truck cleaning", "Other"
]

# ── In-memory state per user ──────────────────────────────────────────────────
user_sessions: dict[int, dict] = {}

FLAG_KEYWORDS = {
    "[CARRIER HOLD]": "CARRIER HOLD",
    "[DRIVER DEDUCTION]": "DRIVER DEDUCTION",
    "[VENDOR ALERT]": "VENDOR ALERT",
    "[MISSING DOCS]": "MISSING DOCS",
    "[RATE CONFLICT]": "RATE CONFLICT",
}

AUDIT_KEYWORDS = [
    "processing payments", "running payouts", "starting settlements",
    "pre-payment check", "run audit", "audit now", "payment run",
]

# ── Main keyboard ─────────────────────────────────────────────────────────────
MAIN_KEYBOARD = ReplyKeyboardMarkup(
    [["📋 Records", "➕ Add New"], ["⚠️ Audit", "📊 Export"]],
    resize_keyboard=True, is_persistent=True,
)

# ── Session helpers ───────────────────────────────────────────────────────────
def get_session(user_id: int) -> dict:
    if user_id not in user_sessions:
        user_sessions[user_id] = {
            "chat": model.start_chat(history=[]),
            "flags": [], "flag_counter": 0,
            "records": [], "record_counter": 0,
            "awaiting": None,
            "pending_record": {},
        }
    return user_sessions[user_id]

def next_id(session: dict) -> str:
    session["record_counter"] += 1
    return f"#{session['record_counter']:03d}"

def add_flag(session: dict, category: str, note: str) -> str:
    session["flag_counter"] += 1
    fid = f"F{session['flag_counter']:03d}"
    session["flags"].append({"id": fid, "category": category, "note": note,
                              "timestamp": datetime.now().strftime("%H:%M"), "resolved": False})
    return fid

def resolve_flag(session: dict, flag_id: str) -> bool:
    for f in session["flags"]:
        if f["id"].upper() == flag_id.upper():
            f["resolved"] = True
            return True
    return False

def build_flags_summary(session: dict) -> str:
    open_flags = [f for f in session["flags"] if not f.get("resolved")]
    resolved = [f for f in session["flags"] if f.get("resolved")]
    if not open_flags and not resolved:
        return "📋 *No flags logged this session.*"
    lines = ["📋 *SESSION FLAGS*\n"]
    if open_flags:
        lines.append("🔴 *OPEN FLAGS:*")
        for f in open_flags:
            lines.append(f"  • `[{f['id']}]` *{f['category']}* — {f['note']} _({f['timestamp']})_")
    else:
        lines.append("🟢 No open flags.")
    if resolved:
        lines.append("\n✅ *RESOLVED:*")
        for f in resolved:
            lines.append(f"  • `[{f['id']}]` {f['category']} — {f['note']}")
    return "\n".join(lines)

def call_gemini(session: dict, text: str) -> str:
    is_audit = any(kw in text.lower() for kw in AUDIT_KEYWORDS)
    open_flags = [f for f in session["flags"] if not f.get("resolved")]
    ctx = ""
    if open_flags:
        ctx = "\n\n[Active flags:]\n" + "\n".join(f"  - [{f['id']}] {f['category']}: {f['note']}" for f in open_flags)
    msg = text + ctx + ("\n\n[SYSTEM: PRE-PAYMENT AUDIT triggered.]" if is_audit else "")
    return session["chat"].send_message(msg).text

# ── Inline keyboards ──────────────────────────────────────────────────────────
def records_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("💸 Deductions", callback_data="rec_Deductions"),
         InlineKeyboardButton("💚 Reimbursements", callback_data="rec_Reimbursements")],
        [InlineKeyboardButton("🔙 Back", callback_data="main_menu")],
    ])

def add_new_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("💸 Deduction", callback_data="add_deduction"),
         InlineKeyboardButton("💚 Reimbursement", callback_data="add_reimbursement")],
        [InlineKeyboardButton("⭐ Bonus", callback_data="add_bonus")],
        [InlineKeyboardButton("🔙 Back", callback_data="main_menu")],
    ])

def audit_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🔍 Run Audit", callback_data="audit_run"),
         InlineKeyboardButton("📋 View Flags", callback_data="audit_flags")],
        [InlineKeyboardButton("🔙 Back", callback_data="main_menu")],
    ])

def export_menu():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("👤 By Name", callback_data="export_name"),
         InlineKeyboardButton("🚛 By Unit", callback_data="export_unit")],
        [InlineKeyboardButton("📋 By Type", callback_data="export_type"),
         InlineKeyboardButton("📅 By Period", callback_data="export_period")],
        [InlineKeyboardButton("🔙 Back", callback_data="main_menu")],
    ])

def type_buttons(types: list, prefix: str):
    rows = []
    row = []
    for i, t in enumerate(types):
        row.append(InlineKeyboardButton(t, callback_data=f"{prefix}_{t[:30]}"))
        if len(row) == 2:
            rows.append(row)
            row = []
    if row:
        rows.append(row)
    rows.append([InlineKeyboardButton("🔙 Back", callback_data="show_add")])
    return InlineKeyboardMarkup(rows)

def dynamic_filter_buttons(session: dict, field: str, prefix: str):
    values = list(set(str(r.get(field, "")) for r in session["records"] if r.get(field)))
    rows = [[InlineKeyboardButton("All", callback_data=f"{prefix}_ALL")]]
    row = []
    for v in sorted(values):
        row.append(InlineKeyboardButton(v, callback_data=f"{prefix}_{v[:30]}"))
        if len(row) == 2:
            rows.append(row)
            row = []
    if row:
        rows.append(row)
    rows.append([InlineKeyboardButton("🔙 Export Menu", callback_data="show_export")])
    return InlineKeyboardMarkup(rows)

def period_buttons(prefix: str):
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("This Week", callback_data=f"{prefix}_this_week"),
         InlineKeyboardButton("Last Week", callback_data=f"{prefix}_last_week")],
        [InlineKeyboardButton("This Month", callback_data=f"{prefix}_this_month"),
         InlineKeyboardButton("Last Month", callback_data=f"{prefix}_last_month")],
        [InlineKeyboardButton("All", callback_data=f"{prefix}_all")],
        [InlineKeyboardButton("🔙 Export Menu", callback_data="show_export")],
    ])

def filter_by_period(records: list, period: str) -> list:
    now = datetime.now()
    if period == "this_week":
        monday = (now - timedelta(days=now.weekday())).replace(hour=0, minute=0, second=0)
        return [r for r in records if r["datetime"] >= monday]
    elif period == "last_week":
        monday = (now - timedelta(days=now.weekday() + 7)).replace(hour=0, minute=0, second=0)
        sunday = monday + timedelta(days=6, hours=23, minutes=59)
        return [r for r in records if monday <= r["datetime"] <= sunday]
    elif period == "this_month":
        start = now.replace(day=1, hour=0, minute=0, second=0)
        return [r for r in records if r["datetime"] >= start]
    elif period == "last_month":
        first = now.replace(day=1, hour=0, minute=0, second=0)
        end = first - timedelta(seconds=1)
        start = end.replace(day=1, hour=0, minute=0, second=0)
        return [r for r in records if start <= r["datetime"] <= end]
    return records

def format_records_list(records: list, title: str) -> str:
    if not records:
        return f"*{title}*\n\nNo records found."
    lines = [f"*{title}* — {len(records)} record(s)\n"]
    for r in records:
        lines.append(
            f"`{r['id']}` *{r['name']}* | Unit {r['unit']} | "
            f"${r['amount']:,.2f} | {r['type']} | _{r.get('note', '—')}_"
        )
    lines.append(f"\n💰 *Total: ${sum(r['amount'] for r in records):,.2f}*")
    return "\n".join(lines)

# ── Command handlers ──────────────────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    get_session(update.effective_user.id)
    await update.message.reply_text(
        "👋 *LogiCheck Online*\nAccounting & Logistics Safety Net — Active\n\nUse the menu below:",
        parse_mode=ParseMode.MARKDOWN, reply_markup=MAIN_KEYBOARD,
    )

async def cmd_resolve(update: Update, context: ContextTypes.DEFAULT_TYPE):
    session = get_session(update.effective_user.id)
    if not context.args:
        await update.message.reply_text("Usage: `/resolve F001`", parse_mode=ParseMode.MARKDOWN)
        return
    fid = context.args[0].upper()
    if resolve_flag(session, fid):
        await update.message.reply_text(f"✅ Flag `{fid}` resolved.", parse_mode=ParseMode.MARKDOWN)
    else:
        await update.message.reply_text(f"❌ Flag `{fid}` not found.", parse_mode=ParseMode.MARKDOWN)

# ── Callback handler ──────────────────────────────────────────────────────────
async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    session = get_session(query.from_user.id)
    data = query.data

    # ── MAIN MENU ──
    if data == "main_menu":
        await query.message.reply_text("Main Menu:", reply_markup=MAIN_KEYBOARD)

    # ── RECORDS ──
    elif data == "show_records":
        await query.message.reply_text("📋 *Records*\nChoose a category:", parse_mode=ParseMode.MARKDOWN, reply_markup=records_menu())

    elif data.startswith("rec_"):
        tab = data.replace("rec_", "")
        filtered = [r for r in session["records"] if r.get("tab") == tab]
        msg = format_records_list(filtered, tab)
        await query.message.reply_text(msg, parse_mode=ParseMode.MARKDOWN, reply_markup=records_menu())

    # ── ADD NEW ──
    elif data == "show_add":
        await query.message.reply_text("➕ *Add New*\nChoose type:", parse_mode=ParseMode.MARKDOWN, reply_markup=add_new_menu())

    elif data == "add_deduction":
        session["pending_record"] = {"tab": "Deductions", "category": "deduction"}
        session["awaiting"] = "select_type"
        await query.message.reply_text(
            "💸 *New Deduction*\nSelect type:",
            parse_mode=ParseMode.MARKDOWN,
            reply_markup=type_buttons(DEDUCTION_TYPES, "dtype"),
        )

    elif data == "add_reimbursement":
        session["pending_record"] = {"tab": "Reimbursements", "category": "reimbursement"}
        session["awaiting"] = "select_type"
        await query.message.reply_text(
            "💚 *New Reimbursement*\nSelect type:",
            parse_mode=ParseMode.MARKDOWN,
            reply_markup=type_buttons(REIMBURSEMENT_TYPES, "rtype"),
        )

    elif data == "add_bonus":
        session["pending_record"] = {"tab": "Reimbursements", "category": "bonus", "type": "Bonus"}
        session["awaiting"] = "enter_details"
        await query.message.reply_text(
            "⭐ *New Bonus*\n\nSend details in this format:\n"
            "`Unit - Driver Name - Amount - Note(optional) - PaymentMethod(optional)`\n\n"
            "Example:\n`2154 - Jama Ahmed - 300 - Safety bonus - Zelle`",
            parse_mode=ParseMode.MARKDOWN,
            reply_markup=ReplyKeyboardRemove(),
        )

    elif data.startswith("dtype_") or data.startswith("rtype_"):
        ptype = data.split("_", 1)[1]
        session["pending_record"]["type"] = ptype
        session["awaiting"] = "enter_details"
        cat = session["pending_record"]["category"]
        emoji = "💸" if cat == "deduction" else "💚"
        await query.message.reply_text(
            f"{emoji} *Type: {ptype}*\n\nSend details:\n"
            "`Unit - Driver Name - Amount - Note(optional) - PaymentMethod(optional)`\n\n"
            "Example:\n`2154 - Jama Ahmed - 2500 - Trailer tire repair - EFS`",
            parse_mode=ParseMode.MARKDOWN,
            reply_markup=ReplyKeyboardRemove(),
        )

    # ── AUDIT ──
    elif data == "show_audit":
        await query.message.reply_text("⚠️ *Audit*", parse_mode=ParseMode.MARKDOWN, reply_markup=audit_menu())

    elif data == "audit_run":
        open_flags = [f for f in session["flags"] if not f.get("resolved")]
        records = session["records"]
        deductions = [r for r in records if r.get("tab") == "Deductions"]
        reimbursements = [r for r in records if r.get("tab") == "Reimbursements"]
        totals = (
            f"\n📊 *SESSION TOTALS:*\n"
            f"  💸 Deductions: ${sum(r['amount'] for r in deductions):,.2f} ({len(deductions)} records)\n"
            f"  💚 Reimbursements/Bonuses: ${sum(r['amount'] for r in reimbursements):,.2f} ({len(reimbursements)} records)"
        )
        if not open_flags:
            msg = "✅ *PRE-PAYMENT AUDIT*\n\nNo active flags.\n*Cleared to process.*\n" + totals
        else:
            items = "\n".join(f"  • `[{f['id']}]` *{f['category']}* — {f['note']}" for f in open_flags)
            msg = f"⚠️ *PRE-PAYMENT AUDIT*\n\n🔴 *{len(open_flags)} UNRESOLVED*\n\n{items}\n\n❗ Resolve all before processing.\n" + totals
        await query.message.reply_text(msg, parse_mode=ParseMode.MARKDOWN, reply_markup=audit_menu())

    elif data == "audit_flags":
        summary = build_flags_summary(session)
        open_count = len([f for f in session["flags"] if not f.get("resolved")])
        resolved_count = len([f for f in session["flags"] if f.get("resolved")])
        await query.message.reply_text(
            summary + f"\n\n📊 *{open_count} open* | ✅ *{resolved_count} resolved*",
            parse_mode=ParseMode.MARKDOWN, reply_markup=audit_menu(),
        )

    # ── EXPORT ──
    elif data == "show_export":
        await query.message.reply_text("📊 *Export*\nFilter by:", parse_mode=ParseMode.MARKDOWN, reply_markup=export_menu())

    elif data == "export_name":
        if not session["records"]:
            await query.message.reply_text("⚠️ No records yet.", reply_markup=export_menu())
        else:
            await query.message.reply_text("👤 Select driver:", reply_markup=dynamic_filter_buttons(session, "name", "fname"))

    elif data == "export_unit":
        if not session["records"]:
            await query.message.reply_text("⚠️ No records yet.", reply_markup=export_menu())
        else:
            await query.message.reply_text("🚛 Select unit:", reply_markup=dynamic_filter_buttons(session, "unit", "funit"))

    elif data == "export_type":
        if not session["records"]:
            await query.message.reply_text("⚠️ No records yet.", reply_markup=export_menu())
        else:
            await query.message.reply_text("📋 Select type:", reply_markup=dynamic_filter_buttons(session, "type", "ftype"))

    elif data == "export_period":
        await query.message.reply_text("📅 Select period:", reply_markup=period_buttons("ponly"))

    elif data.startswith("fname_"):
        value = data.replace("fname_", "")
        session["export_filter"] = {"field": "name", "value": value}
        await query.message.reply_text(f"📅 Select period for *{value}*:", parse_mode=ParseMode.MARKDOWN, reply_markup=period_buttons("pfname"))

    elif data.startswith("funit_"):
        value = data.replace("funit_", "")
        session["export_filter"] = {"field": "unit", "value": value}
        await query.message.reply_text(f"📅 Select period for unit *{value}*:", parse_mode=ParseMode.MARKDOWN, reply_markup=period_buttons("pfunit"))

    elif data.startswith("ftype_"):
        value = data.replace("ftype_", "")
        session["export_filter"] = {"field": "type", "value": value}
        await query.message.reply_text(f"📅 Select period for *{value}*:", parse_mode=ParseMode.MARKDOWN, reply_markup=period_buttons("pftype"))

    elif data.startswith("ponly_") or data.startswith("pfname_") or data.startswith("pfunit_") or data.startswith("pftype_"):
        parts = data.split("_", 1)
        prefix = parts[0]
        period = parts[1]

        records = session["records"]
        ef = session.get("export_filter", {})

        if ef.get("field") and ef.get("value") and ef["value"] != "ALL":
            records = [r for r in records if str(r.get(ef["field"], "")).lower() == ef["value"].lower()]

        if period != "all":
            records = filter_by_period(records, period)

        label = f"{ef.get('value', 'All')} | {period.replace('_', ' ').title()}"
        msg = format_records_list(records, f"Export: {label}")
        await query.message.reply_text(msg, parse_mode=ParseMode.MARKDOWN, reply_markup=export_menu())

# ── Main message handler ──────────────────────────────────────────────────────
async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    session = get_session(update.effective_user.id)
    text = update.message.text.strip()

    # Persistent keyboard buttons
    if text == "📋 Records":
        await update.message.reply_text("📋 *Records*\nChoose:", parse_mode=ParseMode.MARKDOWN, reply_markup=records_menu())
        return
    elif text == "➕ Add New":
        await update.message.reply_text("➕ *Add New*\nChoose type:", parse_mode=ParseMode.MARKDOWN, reply_markup=add_new_menu())
        return
    elif text == "⚠️ Audit":
        await update.message.reply_text("⚠️ *Audit*", parse_mode=ParseMode.MARKDOWN, reply_markup=audit_menu())
        return
    elif text == "📊 Export":
        await update.message.reply_text("📊 *Export*\nFilter by:", parse_mode=ParseMode.MARKDOWN, reply_markup=export_menu())
        return

    # Awaiting record details input
    if session.get("awaiting") == "enter_details":
        parts = [p.strip() for p in text.split("-")]
        if len(parts) < 3:
            await update.message.reply_text(
                "❌ Need at least: `Unit - Name - Amount`\n\nExample:\n`2154 - Jama Ahmed - 2500`",
                parse_mode=ParseMode.MARKDOWN,
            )
            return

        try:
            unit = parts[0]
            name = parts[1]
            amount = float(parts[2].replace("$", "").replace(",", ""))
            note = parts[3] if len(parts) > 3 else ""
            payment_method = parts[4] if len(parts) > 4 else ""
        except (ValueError, IndexError):
            await update.message.reply_text(
                "❌ Wrong format. Use:\n`Unit - Name - Amount - Note - PaymentMethod`",
                parse_mode=ParseMode.MARKDOWN,
            )
            return

        now = datetime.now()
        pending = session["pending_record"]
        tab = pending.get("tab", "Deductions")
        ptype = pending.get("type", "Other")
        category = pending.get("category", "deduction")

        record = {
            "id": next_id(session), "datetime": now,
            "unit": unit, "name": name, "amount": amount,
            "type": ptype, "note": note, "payment_method": payment_method,
            "tab": tab, "category": category,
        }
        session["records"].append(record)
        session["awaiting"] = None
        session["pending_record"] = {}

        # Write to Google Sheets
        week = get_week_number(now)
        date_str = now.strftime("%m/%d/%Y")

        if tab == "Deductions":
            # Columns: A=Week, B=Date, C=Driver, D=Unit, E=Type, F=Amount, G=Deducted, H=Left, I=Deducted period, J=Note, K=Payment method
            row = [week, date_str, name, unit, ptype, amount, "", "", "", note, payment_method]
        else:
            # Reimbursements: A=Week, B=Date, C=Driver, D=Unit, E=Type, F=Amount, J=Note, K=Payment method
            row = [week, date_str, name, unit, ptype, amount, "", "", "", note, payment_method]

        success = append_to_sheet(tab, row)

        emoji_map = {"deduction": "💸", "reimbursement": "💚", "bonus": "⭐"}
        emoji = emoji_map.get(category, "📄")
        sheet_status = "✅ Saved to Google Sheet" if success else "⚠️ Sheet save failed — check connection"

        await update.message.reply_text(
            f"{emoji} *{category.title()} Added*\n\n"
            f"🚛 *Unit:* {unit}\n"
            f"👤 *Driver:* {name}\n"
            f"💰 *Amount:* ${amount:,.2f}\n"
            f"📋 *Type:* {ptype}\n"
            f"📝 *Note:* {note or '—'}\n"
            f"💳 *Method:* {payment_method or '—'}\n"
            f"📅 *Week:* {week}\n\n"
            f"{sheet_status}",
            parse_mode=ParseMode.MARKDOWN,
            reply_markup=MAIN_KEYBOARD,
        )
        return

    # Gemini AI fallback
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    reply = call_gemini(session, text)
    for kw, category in FLAG_KEYWORDS.items():
        if kw in reply:
            fid = add_flag(session, category, text[:120])
            reply = reply.replace(kw, f"{kw} `[{fid}]`", 1)
    try:
        await update.message.reply_text(reply[:4000], parse_mode=ParseMode.MARKDOWN, reply_markup=MAIN_KEYBOARD)
    except Exception:
        await update.message.reply_text(reply[:4000], reply_markup=MAIN_KEYBOARD)

# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        raise ValueError("TELEGRAM_BOT_TOKEN not set.")
    if not os.environ.get("GEMINI_API_KEY"):
        raise ValueError("GEMINI_API_KEY not set.")
    if not os.environ.get("GOOGLE_CREDENTIALS"):
        raise ValueError("GOOGLE_CREDENTIALS not set.")

    app = Application.builder().token(token).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("resolve", cmd_resolve))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    logger.info("LogiCheck bot starting (Gemini + Google Sheets)...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)

if __name__ == "__main__":
    main()
