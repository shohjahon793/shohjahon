"""
LogiCheck - Accounting & Logistics Telegram Bot
Powered by Google Gemini (Free)
"""

import os
import io
import logging
from datetime import datetime, timedelta
import google.generativeai as genai
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    ConversationHandler,
    filters,
    ContextTypes,
)
from telegram.constants import ParseMode

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ── Gemini setup ──────────────────────────────────────────────────────────────
genai.configure(api_key=os.environ["GEMINI_API_KEY"])
model = genai.GenerativeModel(
    model_name="gemini-1.5-flash",
    system_instruction="""You are "LogiCheck," an expert AI Accounting & Logistics Assistant for a motor carrier company. Your primary responsibility is to prevent costly accounting errors, flag payment holds, and assist with financial reconciliations.

Your communication style is concise, clear, and direct—optimized for Telegram chat interface. Use simple Markdown (bold with *, bullet points).

### MODULE 1: NOTIFICATION & PAYMENT SAFETY NET
1. When user mentions payment holds, deductions, missing docs (PODs/BOL), rate conflicts — acknowledge and categorize:
   Labels: [CARRIER HOLD], [DRIVER DEDUCTION], [VENDOR ALERT], [MISSING DOCS], [RATE CONFLICT]

2. PRE-PAYMENT AUDIT TRIGGER phrases: "processing payments", "running payouts", "starting settlements", "pre-payment check", "run audit", "audit now"
   When triggered: generate full audit report of all active flags.

3. If details are ambiguous, ask for clarification immediately.
- Keep responses short, bullet points, mobile-friendly.
- Never process or approve payments yourself — only audit and flag.
"""
)

# ── Conversation states ───────────────────────────────────────────────────────
ASKING_PERIOD = 1

# ── In-memory state per user ──────────────────────────────────────────────────
user_sessions: dict[int, dict] = {}

AUDIT_TRIGGER_KEYWORDS = [
    "processing payments", "running payouts", "starting settlements",
    "pre-payment check", "run audit", "audit now", "payment run",
    "send payments", "batch payment", "process payroll", "run payroll",
]

FLAG_KEYWORDS = {
    "[CARRIER HOLD]": "CARRIER HOLD",
    "[DRIVER DEDUCTION]": "DRIVER DEDUCTION",
    "[VENDOR ALERT]": "VENDOR ALERT",
    "[MISSING DOCS]": "MISSING DOCS",
    "[RATE CONFLICT]": "RATE CONFLICT",
}

PAYMENT_TYPES = ["deduction", "reimbursement", "bonus", "wagehold"]

# ── Helpers ───────────────────────────────────────────────────────────────────

def get_session(user_id: int) -> dict:
    if user_id not in user_sessions:
        user_sessions[user_id] = {
            "chat": model.start_chat(history=[]),
            "flags": [],
            "flag_counter": 0,
            "records": [],       # unified records list
            "record_counter": 0,
        }
    return user_sessions[user_id]


def next_record_id(session: dict) -> str:
    session["record_counter"] += 1
    return f"#{session['record_counter']:03d}"


def get_week_label(date: datetime) -> str:
    """Return 'Week of MM/DD/YYYY' for the Monday of that week."""
    monday = date - timedelta(days=date.weekday())
    return f"Week of {monday.strftime('%m/%d/%Y')}"


def contains_audit_trigger(text: str) -> bool:
    low = text.lower()
    return any(kw in low for kw in AUDIT_TRIGGER_KEYWORDS)


def add_flag(session: dict, category: str, note: str) -> str:
    session["flag_counter"] += 1
    fid = f"F{session['flag_counter']:03d}"
    session["flags"].append({
        "id": fid, "category": category, "note": note,
        "timestamp": datetime.now().strftime("%H:%M"), "resolved": False,
    })
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


def filter_records_by_period(records: list, period: str) -> list:
    """Filter records by period string: 'this week', 'last week', 'this month', 'last month', or date range MM/DD/YYYY-MM/DD/YYYY"""
    now = datetime.now()
    period = period.strip().lower()

    if period == "this week":
        monday = now - timedelta(days=now.weekday())
        monday = monday.replace(hour=0, minute=0, second=0)
        return [r for r in records if r["datetime"] >= monday]

    elif period == "last week":
        monday = now - timedelta(days=now.weekday() + 7)
        monday = monday.replace(hour=0, minute=0, second=0)
        sunday = monday + timedelta(days=6, hours=23, minutes=59)
        return [r for r in records if monday <= r["datetime"] <= sunday]

    elif period == "this month":
        start = now.replace(day=1, hour=0, minute=0, second=0)
        return [r for r in records if r["datetime"] >= start]

    elif period == "last month":
        first_this = now.replace(day=1, hour=0, minute=0, second=0)
        last_month_end = first_this - timedelta(seconds=1)
        last_month_start = last_month_end.replace(day=1, hour=0, minute=0, second=0)
        return [r for r in records if last_month_start <= r["datetime"] <= last_month_end]

    elif period == "all":
        return records

    else:
        # Try date range MM/DD/YYYY-MM/DD/YYYY
        try:
            parts = period.split("-")
            if len(parts) == 6:  # MM/DD/YYYY-MM/DD/YYYY splits into 6
                start_str = "-".join(parts[:3])
                end_str = "-".join(parts[3:])
            else:
                start_str, end_str = period.split(" to ")
            start_dt = datetime.strptime(start_str.strip(), "%m/%d/%Y")
            end_dt = datetime.strptime(end_str.strip(), "%m/%d/%Y").replace(hour=23, minute=59)
            return [r for r in records if start_dt <= r["datetime"] <= end_dt]
        except Exception:
            return records  # fallback: return all


def format_records_list(records: list, rtype: str = None) -> str:
    filtered = [r for r in records if rtype is None or r["payment_type"].lower() == rtype] if rtype else records
    if not filtered:
        return f"No records found."

    emoji_map = {"deduction": "💸", "reimbursement": "💚", "bonus": "⭐", "wagehold": "🔴"}
    lines = []
    total = 0
    for r in filtered:
        e = emoji_map.get(r["payment_type"].lower(), "📄")
        lines.append(
            f"{e} `{r['id']}` *{r['name']}* | Unit {r['unit']} | "
            f"${r['amount']:,.2f} | {r['payment_type']} | {r['payment_method']} | _{r['note']}_"
        )
        total += r["amount"]
    lines.append(f"\n💰 *Total: ${total:,.2f}*")
    return "\n".join(lines)


def parse_new_record(text: str) -> dict | None:
    """
    Format: Unit - Name - Amount - Type - PaymentMethod - Note
    Required: Name, Amount, Type
    Optional: Unit, PaymentMethod, Note
    """
    parts = [p.strip() for p in text.split("-")]
    if len(parts) < 3:
        return None
    try:
        unit = parts[0] if parts[0] else "N/A"
        name = parts[1]
        amount = float(parts[2].replace("$", "").replace(",", ""))
        ptype = parts[3].lower() if len(parts) > 3 else "deduction"
        payment_method = parts[4] if len(parts) > 4 else "—"
        note = parts[5] if len(parts) > 5 else "—"
        if not name or not amount or not ptype:
            return None
        return {
            "unit": unit, "name": name, "amount": amount,
            "payment_type": ptype, "payment_method": payment_method, "note": note,
        }
    except (ValueError, IndexError):
        return None


def build_excel(records: list, period_label: str) -> bytes:
    """Build a styled Excel file from records and return as bytes."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "LogiCheck Records"

    # Column widths
    col_widths = {
        "A": 20, "B": 14, "C": 22, "D": 12,
        "E": 18, "F": 14, "G": 12, "H": 12,
        "I": 12, "J": 30, "K": 18,
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Header styles
    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Title row
    ws.merge_cells("A1:K1")
    title_cell = ws["A1"]
    title_cell.value = f"LogiCheck — Payment Records | {period_label}"
    title_cell.font = Font(name="Arial", bold=True, size=13, color="1F3864")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Headers row 2
    headers = {
        "A": "Week Of", "B": "Date", "C": "Driver Name", "D": "Unit #",
        "E": "Payment Type", "F": "Amount ($)", "G": "", "H": "",
        "I": "", "J": "Note", "K": "Payment Method",
    }
    for col_letter, header in headers.items():
        cell = ws[f"{col_letter}2"]
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = border
    ws.row_dimensions[2].height = 22

    # Data rows
    type_colors = {
        "deduction":     "FFF2CC",
        "reimbursement": "E2EFDA",
        "bonus":         "DDEEFF",
        "wagehold":      "FCE4D6",
    }
    alt_colors = {
        "deduction":     "FFF9E6",
        "reimbursement": "F0F9EC",
        "bonus":         "EEF6FF",
        "wagehold":      "FEF0EB",
    }

    data_font = Font(name="Arial", size=10)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for i, r in enumerate(records):
        row = i + 3
        ptype = r["payment_type"].lower()
        fill_color = type_colors.get(ptype, "FFFFFF") if i % 2 == 0 else alt_colors.get(ptype, "F9F9F9")
        row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

        dt = r["datetime"]
        week_label = get_week_label(dt)
        date_str = dt.strftime("%m/%d/%Y")

        values = {
            "A": week_label, "B": date_str, "C": r["name"], "D": r["unit"],
            "E": r["payment_type"].title(), "F": r["amount"],
            "G": "", "H": "", "I": "",
            "J": r["note"], "K": r["payment_method"],
        }

        for col_letter, val in values.items():
            cell = ws[f"{col_letter}{row}"]
            cell.value = val
            cell.font = data_font
            cell.fill = row_fill
            cell.border = border
            if col_letter == "F":
                cell.number_format = '$#,##0.00'
                cell.alignment = center
            elif col_letter in ("A", "B", "D", "E", "K"):
                cell.alignment = center
            else:
                cell.alignment = left

        ws.row_dimensions[row].height = 18

    # Totals row
    total_row = len(records) + 3
    ws[f"A{total_row}"].value = "TOTAL"
    ws[f"A{total_row}"].font = Font(name="Arial", bold=True, size=10)
    ws[f"A{total_row}"].alignment = center

    if records:
        ws[f"F{total_row}"].value = f"=SUM(F3:F{total_row - 1})"
    else:
        ws[f"F{total_row}"].value = 0

    ws[f"F{total_row}"].font = Font(name="Arial", bold=True, size=10)
    ws[f"F{total_row}"].number_format = '$#,##0.00'
    ws[f"F{total_row}"].alignment = center
    ws[f"F{total_row}"].fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    ws[f"F{total_row}"].font = Font(name="Arial", bold=True, size=10, color="FFFFFF")

    # Freeze header rows
    ws.freeze_panes = "A3"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── Gemini call ───────────────────────────────────────────────────────────────

def call_gemini(session: dict, user_message: str) -> str:
    is_audit = contains_audit_trigger(user_message)
    open_flags = [f for f in session["flags"] if not f.get("resolved")]
    flags_context = ""
    if open_flags:
        flag_list = "\n".join(f"  - [{f['id']}] {f['category']}: {f['note']}" for f in open_flags)
        flags_context = f"\n\n[SYSTEM CONTEXT — Active flags:]\n{flag_list}"
    if is_audit:
        augmented = f"{user_message}{flags_context}\n\n[SYSTEM: User triggered PRE-PAYMENT AUDIT. Generate full report now.]"
    else:
        augmented = user_message + flags_context
    response = session["chat"].send_message(augmented)
    return response.text


# ── /excel conversation ───────────────────────────────────────────────────────

async def cmd_excel_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("This Week", callback_data="period_this week"),
         InlineKeyboardButton("Last Week", callback_data="period_last week")],
        [InlineKeyboardButton("This Month", callback_data="period_this month"),
         InlineKeyboardButton("Last Month", callback_data="period_last month")],
        [InlineKeyboardButton("All Records", callback_data="period_all")],
    ]
    await update.message.reply_text(
        "📊 *Export to Excel*\n\nSelect the period for your report:",
        parse_mode=ParseMode.MARKDOWN,
        reply_markup=InlineKeyboardMarkup(keyboard),
    )


async def handle_period_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if not query.data.startswith("period_"):
        return

    period = query.data.replace("period_", "")
    session = get_session(query.from_user.id)
    filtered = filter_records_by_period(session["records"], period)

    await query.edit_message_text(
        f"⏳ Generating Excel for *{period}*... ({len(filtered)} records)",
        parse_mode=ParseMode.MARKDOWN,
    )

    period_labels = {
        "this week": "This Week",
        "last week": "Last Week",
        "this month": "This Month",
        "last month": "Last Month",
        "all": "All Records",
    }
    label = period_labels.get(period, period.title())

    excel_bytes = build_excel(filtered, label)
    filename = f"LogiCheck_{label.replace(' ', '_')}_{datetime.now().strftime('%m%d%Y')}.xlsx"

    await query.message.reply_document(
        document=io.BytesIO(excel_bytes),
        filename=filename,
        caption=(
            f"📊 *LogiCheck Export*\n"
            f"Period: *{label}*\n"
            f"Records: *{len(filtered)}*\n"
            f"Total: *${sum(r['amount'] for r in filtered):,.2f}*"
        ),
        parse_mode=ParseMode.MARKDOWN,
    )


# ── Record commands ───────────────────────────────────────────────────────────

USAGE_FORMAT = (
    "`/new[type] Unit - Name - Amount - Type - PaymentMethod - Note`\n\n"
    "*Required:* Name, Amount, Type\n"
    "*Optional:* Unit, PaymentMethod, Note\n\n"
    "*Types:* deduction, reimbursement, bonus, wagehold\n\n"
    "*Example:*\n"
    "`/newdeduction 2154 - Jama Ahmed - 2500 - deduction - Check - Cash advance`"
)


async def add_record(update: Update, context: ContextTypes.DEFAULT_TYPE, default_type: str):
    session = get_session(update.effective_user.id)
    text = " ".join(context.args)

    if not text:
        await update.message.reply_text(
            f"📝 *Add a {default_type.title()}*\n\n{USAGE_FORMAT}",
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    record = parse_new_record(text)
    if not record:
        await update.message.reply_text(
            f"❌ *Wrong format.* Use:\n\n{USAGE_FORMAT}",
            parse_mode=ParseMode.MARKDOWN,
        )
        return

    # Override type if using type-specific command
    if default_type != "any":
        record["payment_type"] = default_type

    rid = next_record_id(session)
    now = datetime.now()
    record.update({"id": rid, "datetime": now, "date": now.strftime("%m/%d/%Y %H:%M")})
    session["records"].append(record)

    # Auto-flag wage holds
    if record["payment_type"].lower() == "wagehold":
        add_flag(session, "CARRIER HOLD", f"Wage hold — {record['name']} ${record['amount']:,.2f}")

    emoji_map = {"deduction": "💸", "reimbursement": "💚", "bonus": "⭐", "wagehold": "🔴"}
    emoji = emoji_map.get(record["payment_type"].lower(), "📄")

    await update.message.reply_text(
        f"{emoji} *{record['payment_type'].title()} Added* `{rid}`\n\n"
        f"🚛 *Unit:* {record['unit']}\n"
        f"👤 *Driver:* {record['name']}\n"
        f"💰 *Amount:* ${record['amount']:,.2f}\n"
        f"📋 *Type:* {record['payment_type'].title()}\n"
        f"💳 *Payment Method:* {record['payment_method']}\n"
        f"📝 *Note:* {record['note']}\n"
        f"📅 *Week:* {get_week_label(now)}",
        parse_mode=ParseMode.MARKDOWN,
    )


async def cmd_newdeduction(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await add_record(update, context, "deduction")

async def cmd_newreimbursement(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await add_record(update, context, "reimbursement")

async def cmd_newbonus(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await add_record(update, context, "bonus")

async def cmd_newwagehold(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await add_record(update, context, "wagehold")


async def show_records(update: Update, context: ContextTypes.DEFAULT_TYPE, rtype: str, emoji: str, title: str):
    session = get_session(update.effective_user.id)
    filtered = [r for r in session["records"] if r["payment_type"].lower() == rtype]
    msg = f"{emoji} *{title}* — {len(filtered)} record(s)\n\n"
    if not filtered:
        msg += f"No entries yet.\nUse `/new{rtype}` to add one."
    else:
        for r in filtered:
            msg += (
                f"`{r['id']}` *{r['name']}* | Unit {r['unit']} | "
                f"${r['amount']:,.2f} | {r['payment_method']} | _{r['note']}_\n"
            )
        msg += f"\n💰 *Total: ${sum(r['amount'] for r in filtered):,.2f}*"
    msg += f"\n\n_Export all to Excel: /excel_"
    await update.message.reply_text(msg, parse_mode=ParseMode.MARKDOWN)


async def cmd_deductions(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_records(update, context, "deduction", "💸", "DEDUCTIONS")

async def cmd_reimbursements(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_records(update, context, "reimbursement", "💚", "REIMBURSEMENTS")

async def cmd_bonuses(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_records(update, context, "bonus", "⭐", "BONUSES")

async def cmd_wageholds(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_records(update, context, "wagehold", "🔴", "WAGE HOLDS")


# ── Flags & Audit ─────────────────────────────────────────────────────────────

async def cmd_flags(update: Update, context: ContextTypes.DEFAULT_TYPE):
    session = get_session(update.effective_user.id)
    summary = build_flags_summary(session)
    open_count = len([f for f in session["flags"] if not f.get("resolved")])
    resolved_count = len([f for f in session["flags"] if f.get("resolved")])
    await update.message.reply_text(
        summary + f"\n\n📊 *{open_count} open* | ✅ *{resolved_count} resolved*",
        parse_mode=ParseMode.MARKDOWN,
    )


async def cmd_resolve(update: Update, context: ContextTypes.DEFAULT_TYPE):
    session = get_session(update.effective_user.id)
    if not context.args:
        await update.message.reply_text("⚠️ Usage: `/resolve F001`", parse_mode=ParseMode.MARKDOWN)
        return
    flag_id = context.args[0].upper()
    if resolve_flag(session, flag_id):
        await update.message.reply_text(f"✅ Flag `{flag_id}` marked as *resolved*.", parse_mode=ParseMode.MARKDOWN)
    else:
        await update.message.reply_text(f"❌ Flag `{flag_id}` not found. Use /flags to see IDs.", parse_mode=ParseMode.MARKDOWN)


async def cmd_audit(update: Update, context: ContextTypes.DEFAULT_TYPE):
    session = get_session(update.effective_user.id)
    await update.message.reply_text("🔍 Running pre-payment audit...")
    open_flags = [f for f in session["flags"] if not f.get("resolved")]

    records = session["records"]
    totals = (
        f"\n📊 *SESSION TOTALS:*\n"
        f"  💸 Deductions: ${sum(r['amount'] for r in records if r['payment_type']=='deduction'):,.2f}\n"
        f"  💚 Reimbursements: ${sum(r['amount'] for r in records if r['payment_type']=='reimbursement'):,.2f}\n"
        f"  ⭐ Bonuses: ${sum(r['amount'] for r in records if r['payment_type']=='bonus'):,.2f}\n"
        f"  🔴 Wage Holds: ${sum(r['amount'] for r in records if r['payment_type']=='wagehold'):,.2f}\n"
        f"  📋 Total Records: {len(records)}"
    )

    if not open_flags:
        audit_msg = "✅ *PRE-PAYMENT AUDIT COMPLETE*\n\nNo active flags.\n*Cleared to process payments.*\n" + totals
    else:
        items = "\n".join(f"  • `[{f['id']}]` *{f['category']}* — {f['note']}" for f in open_flags)
        audit_msg = (
            f"⚠️ *PRE-PAYMENT AUDIT REPORT*\n\n"
            f"🔴 *{len(open_flags)} UNRESOLVED FLAG(S)*\n\n{items}\n\n"
            f"─────────────────\n"
            f"❗ *Do NOT send batch until all flags resolved.*\n"
            f"Use `/resolve [ID]` to clear.\n" + totals
        )

    keyboard = [[
        InlineKeyboardButton("📋 View Flags", callback_data="view_flags"),
        InlineKeyboardButton("📊 Export Excel", callback_data="period_all"),
    ]]
    await update.message.reply_text(audit_msg, parse_mode=ParseMode.MARKDOWN, reply_markup=InlineKeyboardMarkup(keyboard))


async def cmd_clearall(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [[
        InlineKeyboardButton("Yes, clear all", callback_data="confirm_clear"),
        InlineKeyboardButton("Cancel", callback_data="cancel_clear"),
    ]]
    await update.message.reply_text(
        "⚠️ *Clear ALL session data?*\nAll records, flags, deductions — everything. Cannot be undone.",
        parse_mode=ParseMode.MARKDOWN,
        reply_markup=InlineKeyboardMarkup(keyboard),
    )


# ── Start / Help ──────────────────────────────────────────────────────────────

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    get_session(update.effective_user.id)
    welcome = (
        "👋 *LogiCheck Online*\n"
        "Accounting & Logistics Safety Net — Active\n\n"
        "📋 *RECORDS:*\n"
        "/deductions — View deductions\n"
        "/reimbursements — View reimbursements\n"
        "/bonuses — View bonuses\n"
        "/wageholds — View wage holds\n\n"
        "➕ *ADD RECORDS:*\n"
        "`/newdeduction Unit \\- Name \\- Amount \\- Type \\- Method \\- Note`\n"
        "`/newreimbursement Unit \\- Name \\- Amount \\- ...`\n"
        "`/newbonus Unit \\- Name \\- Amount \\- ...`\n"
        "`/newwagehold Unit \\- Name \\- Amount \\- ...`\n\n"
        "📊 *EXPORT:*\n"
        "/excel — Export records to Excel file\n\n"
        "⚠️ *AUDIT:*\n"
        "/flags — View session flags\n"
        "/resolve F001 — Mark flag resolved\n"
        "/audit — Run pre\\-payment audit\n"
        "/clearall — Clear all session data\n\n"
        "💬 Or just *type naturally* to log issues\\!"
    )
    await update.message.reply_text(welcome, parse_mode=ParseMode.MARKDOWN_V2)


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await cmd_start(update, context)


# ── Callback handler ──────────────────────────────────────────────────────────

async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    session = get_session(query.from_user.id)

    if query.data.startswith("period_"):
        await handle_period_callback(update, context)

    elif query.data == "view_flags":
        summary = build_flags_summary(session)
        await query.message.reply_text(summary, parse_mode=ParseMode.MARKDOWN)

    elif query.data == "all_clear":
        await query.message.reply_text(
            "✅ Acknowledged. Proceed with payment batch.\n_Ensure all flags are truly resolved._",
            parse_mode=ParseMode.MARKDOWN,
        )

    elif query.data == "confirm_clear":
        user_sessions[query.from_user.id] = {
            "chat": model.start_chat(history=[]),
            "flags": [], "flag_counter": 0,
            "records": [], "record_counter": 0,
        }
        await query.edit_message_text("🗑️ Session cleared. Starting fresh.")

    elif query.data == "cancel_clear":
        await query.edit_message_text("❌ Clear cancelled. Session intact.")


# ── Main message handler ──────────────────────────────────────────────────────

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    session = get_session(update.effective_user.id)
    text = update.message.text.strip()
    await context.bot.send_chat_action(chat_id=update.effective_chat.id, action="typing")
    reply = call_gemini(session, text)
    for kw, category in FLAG_KEYWORDS.items():
        if kw in reply:
            note = text[:120] + ("..." if len(text) > 120 else "")
            fid = add_flag(session, category, note)
            reply = reply.replace(kw, f"{kw} `[{fid}]`", 1)
    try:
        await update.message.reply_text(reply[:4000], parse_mode=ParseMode.MARKDOWN)
    except Exception:
        await update.message.reply_text(reply[:4000])


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    if not token:
        raise ValueError("TELEGRAM_BOT_TOKEN not set.")
    if not os.environ.get("GEMINI_API_KEY"):
        raise ValueError("GEMINI_API_KEY not set.")

    app = Application.builder().token(token).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("flags", cmd_flags))
    app.add_handler(CommandHandler("resolve", cmd_resolve))
    app.add_handler(CommandHandler("audit", cmd_audit))
    app.add_handler(CommandHandler("clearall", cmd_clearall))
    app.add_handler(CommandHandler("excel", cmd_excel_start))
    app.add_handler(CommandHandler("deductions", cmd_deductions))
    app.add_handler(CommandHandler("newdeduction", cmd_newdeduction))
    app.add_handler(CommandHandler("reimbursements", cmd_reimbursements))
    app.add_handler(CommandHandler("newreimbursement", cmd_newreimbursement))
    app.add_handler(CommandHandler("bonuses", cmd_bonuses))
    app.add_handler(CommandHandler("newbonus", cmd_newbonus))
    app.add_handler(CommandHandler("wageholds", cmd_wageholds))
    app.add_handler(CommandHandler("newwagehold", cmd_newwagehold))
    app.add_handler(CallbackQueryHandler(button_callback))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    logger.info("LogiCheck bot starting (Gemini + Excel)...")
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
