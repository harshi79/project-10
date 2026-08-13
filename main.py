"""
Yori Cleaner Bot — Telegram file processing + fake-data filtering bot.

What it does now:
  • Parses .txt files of cards / email combos / phone combos (mixed files too)
  • VALIDATES every entry and classifies it as VALID or FAKE:
      - Cards   : Luhn checksum, BIN brand, length, expiry, CVV
      - Emails  : format + disposable (temp-mail) domain detection
      - Phones  : number format + length
      - Passwords: common-password list + strength scoring
  • Lets you filter output to Valid only / Fake only / All
  • Outputs TXT, CSV and a real Excel (.xlsx) report with Status + Reason
  • Per-user and global stats, broadcast, queue history

All data used is fake/test data — this is an educational project about how
filtering works, nothing else. No card or account is ever contacted.

Offline demo (no internet / no deps):   python validator.py sample.txt
Run the bot:                             set TELEGRAM_BOT_TOKEN, then python main.py
"""

import os
import sys
import csv as _csv
import sqlite3
import logging
import threading
from io import BytesIO, StringIO
from datetime import datetime, timezone
from collections import defaultdict
from http.server import HTTPServer, BaseHTTPRequestHandler

import aiohttp
from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    filters,
    ContextTypes,
)
from telegram.error import Forbidden, BadRequest
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

from validator import (
    analyse_line,
    analyse_file,
    validate_card,
    validate_email,
    validate_password,
    validate_phone,
    luhn_check,
    count_valid,
    filter_entries,
    reason_str,
)

# ── Config ──────────────────────────────────────────────────────────────────────

TOKEN       = os.environ.get("TELEGRAM_BOT_TOKEN", "")
OWNER_ID    = 7728424218
WATERMARK   = "\n\n— @yorifederation"
WEBHOOK_URL = os.environ.get("WEBHOOK_URL", "")
PORT        = int(os.environ.get("PORT", 8080))
DB_PATH     = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.db")

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ── Queue ───────────────────────────────────────────────────────────────────────

file_queue: list[dict] = []
queue_lock = threading.Lock()

# ── Health server ───────────────────────────────────────────────────────────────

class _Health(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path in ("/", "/health", "/healthz"):
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            self.end_headers()
            self.wfile.write(b"OK - Yori Bot is alive")
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, *a):
        pass


def _start_health():
    try:
        HTTPServer(("", PORT), _Health).serve_forever()
    except OSError:
        pass

# ── Database ────────────────────────────────────────────────────────────────────

def _conn() -> sqlite3.Connection:
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c


def init_db() -> None:
    with _conn() as c:
        c.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id        INTEGER PRIMARY KEY,
                name      TEXT    DEFAULT '',
                username  TEXT    DEFAULT '',
                files     INTEGER DEFAULT 0,
                lines     INTEGER DEFAULT 0,
                cards     INTEGER DEFAULT 0,
                combos    INTEGER DEFAULT 0,
                valid     INTEGER DEFAULT 0,
                fake      INTEGER DEFAULT 0,
                last_seen TEXT    DEFAULT ''
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS queue_history (
                id       INTEGER PRIMARY KEY AUTOINCREMENT,
                uid      INTEGER,
                filename TEXT,
                status   TEXT,
                ts       TEXT
            )
        """)
        # Migration for DBs created before the valid/fake columns existed
        cols = [r[1] for r in c.execute("PRAGMA table_info(users)").fetchall()]
        if "valid" not in cols:
            c.execute("ALTER TABLE users ADD COLUMN valid INTEGER DEFAULT 0")
        if "fake" not in cols:
            c.execute("ALTER TABLE users ADD COLUMN fake INTEGER DEFAULT 0")


def upsert_user(uid: int, name: str, username: str, n_cards: int, n_combos: int,
                n_valid: int, n_fake: int) -> None:
    total = n_cards + n_combos
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    with _conn() as c:
        c.execute("""
            INSERT INTO users (id, name, username, files, lines, cards, combos, valid, fake, last_seen)
            VALUES (?,?,?,1,?,?,?,?,?,?)
            ON CONFLICT(id) DO UPDATE SET
                name      = excluded.name,
                username  = excluded.username,
                files     = files  + 1,
                lines     = lines  + excluded.lines,
                cards     = cards  + excluded.cards,
                combos    = combos + excluded.combos,
                valid     = valid  + excluded.valid,
                fake      = fake   + excluded.fake,
                last_seen = excluded.last_seen
        """, (uid, name, username, total, n_cards, n_combos, n_valid, n_fake, ts))


def get_user(uid: int):
    with _conn() as c:
        return c.execute("SELECT * FROM users WHERE id=?", (uid,)).fetchone()


def get_all_user_ids() -> list[int]:
    with _conn() as c:
        return [r[0] for r in c.execute("SELECT id FROM users").fetchall()]


def get_global_stats():
    with _conn() as c:
        total_users = c.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        total_files = c.execute("SELECT COALESCE(SUM(files),0) FROM users").fetchone()[0]
        total_lines = c.execute("SELECT COALESCE(SUM(lines),0) FROM users").fetchone()[0]
        total_valid = c.execute("SELECT COALESCE(SUM(valid),0) FROM users").fetchone()[0]
        total_fake  = c.execute("SELECT COALESCE(SUM(fake),0) FROM users").fetchone()[0]
        top5 = c.execute(
            "SELECT name, username, lines FROM users ORDER BY lines DESC LIMIT 5"
        ).fetchall()
    return total_users, total_files, total_lines, total_valid, total_fake, top5


def log_queue(uid: int, filename: str, status: str) -> None:
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    with _conn() as c:
        c.execute("INSERT INTO queue_history (uid, filename, status, ts) VALUES (?,?,?,?)",
                  (uid, filename, status, ts))


def get_queue_history(uid: int, limit: int = 10):
    with _conn() as c:
        return c.execute(
            "SELECT filename, status, ts FROM queue_history WHERE uid=? ORDER BY id DESC LIMIT ?",
            (uid, limit)
        ).fetchall()


# ── Combo helpers ───────────────────────────────────────────────────────────────

def get_domains(combos: list[dict]) -> list[str]:
    domain_counts: dict[str, int] = defaultdict(int)
    for c in combos:
        domain_counts[c["domain"]] += 1
    return sorted(domain_counts.keys(), key=lambda d: (-domain_counts[d], d))


def sort_combos_by_domain(combos: list[dict]) -> list[dict]:
    return sorted(combos, key=lambda c: (c["domain"], c["email"].lower()))


def filter_combos_by_domain(combos: list[dict], domain: str) -> list[dict]:
    return [c for c in combos if c["domain"] == domain.lower()]


# ── Output builders ─────────────────────────────────────────────────────────────

def build_txt(cards: list, combos: list, phones: list) -> str:
    parts: list[str] = []
    sections = sum([bool(cards), bool(combos), bool(phones)])
    if sections > 1:
        if cards:
            parts += [f"━━━ CARDS ({len(cards)}) ━━━", *[c["value"] for c in cards], ""]
        if combos:
            parts += [f"━━━ COMBOS ({len(combos)}) ━━━",
                      *[f"{c['email']}   {c['pw']}" for c in combos], ""]
        if phones:
            parts += [f"━━━ PHONES ({len(phones)}) ━━━",
                      *[f"{c['phone']}   {c['pw']}" for c in phones]]
    else:
        parts += ([c["value"] for c in cards]
                  or [f"{c['email']}   {c['pw']}" for c in combos]
                  or [f"{c['phone']}   {c['pw']}" for c in phones])
    return "\n".join(parts) + WATERMARK


def build_csv(cards: list, combos: list, phones: list) -> str:
    rows: list[list] = []
    if cards:
        rows.append(["Type", "Number", "Month", "Year", "CVV", "Brand", "Status", "Reason"])
        for c in cards:
            rows.append(["card", c["num"], c["mm"], c["yy"], c["cvv"], c["brand"],
                         "VALID" if c["valid"] else "FAKE", reason_str(c["reasons"])])
    if combos:
        rows.append(["Type", "Email", "Password", "Domain", "Status", "Reason", "Strength"])
        for c in combos:
            rows.append(["combo", c["email"], c["pw"], c["domain"],
                         "VALID" if c["valid"] else "FAKE",
                         reason_str(c["reasons"]), c["pw_strength"]])
    if phones:
        rows.append(["Type", "Phone", "Password", "Status", "Reason", "Strength"])
        for c in phones:
            rows.append(["phone", c["phone"], c["pw"],
                         "VALID" if c["valid"] else "FAKE",
                         reason_str(c["reasons"]), c["pw_strength"]])
    buf = StringIO()
    _csv.writer(buf).writerows(rows)
    return buf.getvalue()


def _style_header(cell) -> None:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="4472C4")
    cell.alignment = Alignment(horizontal="center")


def _write_table(ws, headers: list, rows: list) -> None:
    ws.append(headers)
    for cell in ws[1]:
        _style_header(cell)
    for row in rows:
        ws.append(row)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            v = str(cell.value or "")
            if v == "FAKE":
                cell.fill = PatternFill("solid", fgColor="FFC7CE")
                cell.font = Font(color="9C0006")
            elif v == "VALID":
                cell.fill = PatternFill("solid", fgColor="C6EFCE")
                cell.font = Font(color="006100")
    for col in ws.columns:
        letter = col[0].column_letter
        m = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[letter].width = min(max(m + 2, 8), 42)


def _write_summary_sheet(ws, cards: list, combos: list, phones: list) -> None:
    ws.append(["Yori Cleaner — Fake-Data Filtering Report"])
    ws.append(["Generated", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")])
    ws.append([])
    ws.append(["Type", "Total", "Valid", "Fake", "Fake %"])
    for cell in ws[4]:
        _style_header(cell)
    for name, entries in (("Cards", cards), ("Emails", combos), ("Phones", phones)):
        total = len(entries)
        v = count_valid(entries)
        f = total - v
        pct = f"{f / total * 100:.1f}%" if total else "0%"
        ws.append([name, total, v, f, pct])
    ws.append([])
    ws.append(["Note",
               "Validation is structural (Luhn checksum, format, disposable-domain "
               "& common-password heuristics). Educational use on test data only."])
    for letter, width in (("A", 12), ("B", 10), ("C", 10), ("D", 10), ("E", 10)):
        ws.column_dimensions[letter].width = width


def build_xlsx(cards: list, combos: list, phones: list) -> bytes:
    wb = Workbook()

    ws = wb.active
    ws.title = "Summary"
    _write_summary_sheet(ws, cards, combos, phones)

    if cards:
        ws_c = wb.create_sheet("Cards")
        _write_table(ws_c,
                     ["Number", "Month", "Year", "CVV", "Brand", "Status", "Reason"],
                     [[c["num"], c["mm"], c["yy"], c["cvv"], c["brand"],
                       "VALID" if c["valid"] else "FAKE",
                       reason_str(c["reasons"])] for c in cards])
    if combos:
        ws_e = wb.create_sheet("Emails")
        _write_table(ws_e,
                     ["Email", "Password", "Domain", "Status", "Reason", "Strength"],
                     [[c["email"], c["pw"], c["domain"],
                       "VALID" if c["valid"] else "FAKE",
                       reason_str(c["reasons"]), c["pw_strength"]] for c in combos])
    if phones:
        ws_p = wb.create_sheet("Phones")
        _write_table(ws_p,
                     ["Phone", "Password", "Status", "Reason", "Strength"],
                     [[c["phone"], c["pw"],
                       "VALID" if c["valid"] else "FAKE",
                       reason_str(c["reasons"]), c["pw_strength"]] for c in phones])

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Keyboards ───────────────────────────────────────────────────────────────────

MAIN_KB = ReplyKeyboardMarkup(
    [["📊 My Stats", "ℹ️ Help"], ["🏷️ About"]],
    resize_keyboard=True, is_persistent=True,
)


def type_ikb(cards: list, combos: list, phones: list, uid: int, fmt: str = "txt",
             selected_types: set | None = None, domain: str | None = None,
             sort: bool = False, validity: str = "all") -> InlineKeyboardMarkup:
    """Inline keyboard: format, types, valid/fake filter, domains, generate."""
    rows = []
    sel = selected_types or set()

    # Row 1: Format
    rows.append([
        InlineKeyboardButton(f"📄 TXT {'✅' if fmt == 'txt' else ''}", callback_data=f"fmt:txt:{uid}"),
        InlineKeyboardButton(f"📊 CSV {'✅' if fmt == 'csv' else ''}", callback_data=f"fmt:csv:{uid}"),
        InlineKeyboardButton(f"📈 Excel {'✅' if fmt == 'xlsx' else ''}", callback_data=f"fmt:xlsx:{uid}"),
    ])

    # Row 2: Type selection
    type_row = []
    if cards:
        type_row.append(InlineKeyboardButton(
            f"💳 Cards {'✅' if 'cards' in sel else ''}", callback_data=f"sel:cards:{uid}"))
    if combos:
        type_row.append(InlineKeyboardButton(
            f"🔑 Emails {'✅' if 'combos' in sel else ''}", callback_data=f"sel:combos:{uid}"))
    if phones:
        type_row.append(InlineKeyboardButton(
            f"📱 Phones {'✅' if 'phones' in sel else ''}", callback_data=f"sel:phones:{uid}"))
    if len(type_row) > 1:
        type_row.append(InlineKeyboardButton(
            f"🔀 All {'✅' if not sel else ''}", callback_data=f"sel:all:{uid}"))
    if type_row:
        rows.append(type_row)

    # Row 3: Valid / Fake filter
    rows.append([
        InlineKeyboardButton(f"🟢 Valid {'✅' if validity == 'valid' else ''}",
                             callback_data=f"val:valid:{uid}"),
        InlineKeyboardButton(f"🔴 Fake {'✅' if validity == 'fake' else ''}",
                             callback_data=f"val:fake:{uid}"),
        InlineKeyboardButton(f"🔀 All {'✅' if validity == 'all' else ''}",
                             callback_data=f"val:all:{uid}"),
    ])

    # Domain buttons (if combos exist)
    if combos:
        domains = get_domains(combos)[:5]
        if domains:
            rows.append([InlineKeyboardButton(
                f"📧 {d} {'✅' if domain == d else ''}",
                callback_data=f"seld:{d}:{uid}") for d in domains])
        rows.append([InlineKeyboardButton(
            f"🔤 Sort by domain {'✅' if sort else ''}",
            callback_data=f"sels:domain:{uid}")])

    rows.append([InlineKeyboardButton("🚀 GENERATE", callback_data=f"gen:{fmt}:{uid}")])

    return InlineKeyboardMarkup(rows)


def result_ikb(uid: int) -> InlineKeyboardMarkup:
    rows = [[
        InlineKeyboardButton("📊 My Stats", callback_data=f"stats:{uid}"),
        InlineKeyboardButton("ℹ️ Help", callback_data="help"),
    ]]
    if uid == OWNER_ID:
        rows.append([InlineKeyboardButton("👑 Global Stats", callback_data="owner:stats")])
    return InlineKeyboardMarkup(rows)


# ── Text helpers ────────────────────────────────────────────────────────────────

HELP_TEXT = (
    "<b>ℹ️ How to use</b>\n"
    "──────────────────\n"
    "Send any <b>.txt</b> file. Every line is analysed and classified as\n"
    "🟢 <b>VALID</b> (passes checks) or 🔴 <b>FAKE</b> (filtered out).\n\n"
    "💳 <b>Cards</b> — any separator:\n"
    "  <code>4111111111111111|05|33|496 — 🇦🇪 AE</code>\n"
    "  <code>4111111111111111 05 33 496</code>\n\n"
    "🔑 <b>Email combos</b>:\n"
    "  <code>user@gmail.com:Password1</code>\n"
    "  <code>user@gmail.com|Password1</code>\n\n"
    "📱 <b>Phone combos</b>:\n"
    "  <code>+12345678901:Password1</code>\n\n"
    "🔍 <b>What is checked:</b>\n"
    "  💳 Luhn checksum, brand, length, expiry, CVV\n"
    "  📧 format + disposable (temp-mail) domains\n"
    "  📱 number format + length\n"
    "  🔑 common / weak passwords\n\n"
    "🧪 Quick single-line check:\n"
    "  <code>/check 4111111111111111|05|33|496</code>\n\n"
    "📊 Choose TXT / CSV / Excel and filter Valid / Fake / All.\n\n"
    "<i>— @yorifederation</i>"
)

ABOUT_TEXT = (
    "<b>🏷️ @yorifederation Cleaner + Filter Bot</b>\n"
    "──────────────────\n"
    "⚡ Instant .txt file analysis\n"
    "🧠 Fuzzy per-line pattern detection\n"
    "💳 Card validation (Luhn, BIN, expiry, CVV)\n"
    "📧 Disposable / temp-mail domain detection\n"
    "📱 Phone number format validation\n"
    "🔑 Common-password + strength analysis\n"
    "🟢🔴 Valid / Fake classification & filtering\n"
    "📊 TXT / CSV / real Excel (.xlsx) report\n"
    "🔑 Deduplication built in\n"
    "📊 Per-user stats (SQLite)\n\n"
    "<i>Educational use on test data only.\n"
    "— @yorifederation</i>"
)


def user_stats_text(row) -> str:
    return (
        f"<b>📊 Your Stats</b>\n"
        f"──────────────────\n"
        f"👤 {row['name'] or 'Unknown'}  <i>{row['username']}</i>\n\n"
        f"📁 Files processed  <b>{row['files']:,}</b>\n"
        f"📝 Total lines      <b>{row['lines']:,}</b>\n"
        f"💳 Card lines       <b>{row['cards']:,}</b>\n"
        f"🔑 Combo lines      <b>{row['combos']:,}</b>\n"
        f"🟢 Valid entries    <b>{row['valid']:,}</b>\n"
        f"🔴 Fake entries     <b>{row['fake']:,}</b>\n\n"
        f"🕒 <i>{row['last_seen']}</i>\n\n"
        f"<i>— @yorifederation</i>"
    )


def global_stats_text() -> str:
    total_users, total_files, total_lines, total_valid, total_fake, top5 = get_global_stats()
    top_str = "\n".join(
        f"  {i + 1}. {r['name'] or r['username']} — {r['lines']:,} lines"
        for i, r in enumerate(top5)
    ) or "  No data yet."
    return (
        f"<b>👑 Global Stats</b>\n"
        f"──────────────────\n"
        f"👥 Total users   <b>{total_users:,}</b>\n"
        f"📁 Total files   <b>{total_files:,}</b>\n"
        f"📝 Total lines   <b>{total_lines:,}</b>\n"
        f"🟢 Valid entries <b>{total_valid:,}</b>\n"
        f"🔴 Fake entries  <b>{total_fake:,}</b>\n\n"
        f"🏆 <b>Top 5</b>\n{top_str}\n\n"
        f"<i>— @yorifederation</i>"
    )


def check_text(raw: str) -> str:
    """HTML verdict for a single line (used by /check)."""
    r = analyse_line(raw)
    if r is None:
        return (
            "❓ <b>Not recognised.</b>\n\n"
            "Send a card, email:pass or phone:pass line, e.g.:\n"
            "<code>4111111111111111|05|33|496</code>\n"
            "<code>user@gmail.com:Password1</code>\n"
            "<code>+12345678901:Password1</code>"
        )

    if r[0] == "card":
        _, num, mm, yy, cvv = r
        v = validate_card(num, mm, yy, cvv)
        verdict = "🟢 <b>VALID</b>" if v["valid"] else "🔴 <b>FAKE</b>"
        out = (f"{verdict}\n"
               f"━━━━━━━━━━━━━━━━\n"
               f"💳 {num}\n"
               f"🏷️ Brand: <b>{v['brand']}</b>\n"
               f"📅 Expiry: {mm}/{yy}\n"
               f"🔒 CVV: {cvv}\n"
               f"🧪 Luhn: {'pass 🟢' if luhn_check(num) else 'fail 🔴'}\n")
        out += (f"⚠️ Reasons: {reason_str(v['reasons'])}" if v["reasons"]
                else "✅ No issues found")
        return out

    if r[0] == "combo":
        _, email, pw = r
        e = validate_email(email)
        p = validate_password(pw)
        valid = e["valid"] and p["valid"]
        verdict = "🟢 <b>VALID</b>" if valid else "🔴 <b>FAKE</b>"
        lines = [
            f"{verdict}",
            "━━━━━━━━━━━━━━━━",
            f"📧 {email}",
            f"🌐 Domain: {e['domain']}",
            f"🔑 Password strength: <b>{p['strength']}</b>",
        ]
        reasons = list(e["reasons"]) + [f"password: {x}" for x in p["reasons"]]
        lines.append(f"⚠️ Reasons: {reason_str(reasons)}" if reasons else "✅ No issues found")
        return "\n".join(lines)

    _, phone, pw = r
    ph = validate_phone(phone)
    p = validate_password(pw)
    valid = ph["valid"] and p["valid"]
    verdict = "🟢 <b>VALID</b>" if valid else "🔴 <b>FAKE</b>"
    reasons = list(ph["reasons"]) + [f"password: {x}" for x in p["reasons"]]
    return (f"{verdict}\n"
            f"━━━━━━━━━━━━━━━━\n"
            f"📱 {phone}\n"
            f"🔑 Password strength: <b>{p['strength']}</b>\n"
            + (f"⚠️ Reasons: {reason_str(reasons)}" if reasons else "✅ No issues found"))


# ── Handlers ────────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    name = update.effective_user.first_name or "there"
    await update.message.reply_text(
        f"⚡ <b>Yori Cleaner</b>  <i>by @yorifederation</i>\n"
        f"──────────────────────\n\n"
        f"Welcome, <b>{name}</b>.\n\n"
        f"Drop a <b>.txt</b> file and I will:\n"
        f"  🧠 Analyse every line automatically\n"
        f"  🟢🔴 Classify each entry as VALID or FAKE\n"
        f"  💳 Check Luhn / BIN / expiry / CVV\n"
        f"  📧 Detect disposable emails & weak passwords\n"
        f"  📊 Let you export TXT / CSV / Excel\n"
        f"  🚫 Filter out the fake ones (Valid / Fake / All)\n\n"
        f"No commands needed — just send the file.",
        parse_mode="HTML", reply_markup=MAIN_KB,
    )


async def cmd_check(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if not ctx.args:
        await update.message.reply_text(
            "🧪 <b>Check a single line</b>\n"
            "Usage:\n"
            "<code>/check 4111111111111111|05|33|496</code>\n"
            "<code>/check user@gmail.com:Password1</code>\n"
            "<code>/check +12345678901:Password1</code>",
            parse_mode="HTML",
        )
        return
    raw = " ".join(ctx.args)
    await update.message.reply_text(check_text(raw), parse_mode="HTML")


async def cmd_stats(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if update.effective_user.id != OWNER_ID:
        await update.message.reply_text("⛔ Owner only.")
        return
    await update.message.reply_text(global_stats_text(), parse_mode="HTML")


async def cmd_broadcast(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    if update.effective_user.id != OWNER_ID:
        await update.message.reply_text("⛔ Owner only.")
        return
    if not ctx.args:
        await update.message.reply_text(
            "Usage: /broadcast <message>\n\nBlocked users are automatically skipped."
        )
        return

    text = " ".join(ctx.args)
    users = get_all_user_ids()
    sent = blocked = failed = 0

    status_msg = await update.message.reply_text(
        f"📣 Broadcasting to {len(users):,} users…"
    )

    for uid in users:
        try:
            await ctx.bot.send_message(uid, text, parse_mode="HTML")
            sent += 1
        except (Forbidden, BadRequest):
            blocked += 1
        except Exception as e:
            log.warning("Broadcast error uid=%s: %s", uid, e)
            failed += 1

    await status_msg.edit_text(
        f"📣 <b>Broadcast complete</b>\n"
        f"──────────────────\n"
        f"✅ Sent:            <b>{sent:,}</b>\n"
        f"🚫 Blocked/skipped: <b>{blocked:,}</b>\n"
        f"❌ Other errors:    <b>{failed:,}</b>",
        parse_mode="HTML",
    )


async def cmd_queue(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    uid = update.effective_user.id
    history = get_queue_history(uid)
    if not history:
        await update.message.reply_text("📭 No queue history yet.")
        return
    lines = ["<b>📋 Recent Queue</b>\n──────────────────"]
    for fn, status, ts in history:
        icon = "✅" if status == "done" else "⏳" if status == "queued" else "❌"
        lines.append(f"{icon} <code>{fn[:30]}</code> — {status} — {ts}")
    await update.message.reply_text("\n".join(lines), parse_mode="HTML")


# ── File processing ─────────────────────────────────────────────────────────────

async def handle_document(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    doc = update.message.document
    uid = update.effective_user.id
    user = update.effective_user

    if not doc.file_name or not doc.file_name.lower().endswith(".txt"):
        await update.message.reply_text(
            "❌ <b>Only .txt files accepted.</b>\n\n"
            "Send a plain text file — I'll take care of the rest.",
            parse_mode="HTML", reply_markup=MAIN_KB,
        )
        return

    with queue_lock:
        user_queued = sum(1 for q in file_queue if q["uid"] == uid)
    if user_queued >= 3:
        await update.message.reply_text(
            "⏳ <b>Queue full.</b> You have 3 files pending.\n"
            "Wait for one to finish before sending more.",
            parse_mode="HTML",
        )
        return

    thinking = await update.message.reply_text(
        "🧠 <b>Analysing & validating file…</b>", parse_mode="HTML"
    )

    try:
        tg_file = await ctx.bot.get_file(doc.file_id)
        async with aiohttp.ClientSession() as session:
            async with session.get(tg_file.file_path) as resp:
                content = await resp.text(encoding="utf-8", errors="replace")

        result = analyse_file(content)
        cards, combos, phones = result["cards"], result["combos"], result["phones"]
        skipped, total = result["skipped"], result["total"]

        await thinking.delete()

        if not cards and not combos and not phones:
            await update.message.reply_text(
                f"⚠️ <b>Nothing recognised</b> in <code>{doc.file_name}</code>.\n\n"
                f"Supported: cards, email combos, phone combos, or any mix.",
                parse_mode="HTML", reply_markup=MAIN_KB,
            )
            return

        base = doc.file_name.rsplit(".txt", 1)[0].rsplit(".TXT", 1)[0]

        full_name = " ".join(filter(None, [user.first_name, user.last_name]))
        uname = f"@{user.username}" if user.username else "—"
        with queue_lock:
            file_queue.append({
                "uid": uid,
                "filename": doc.file_name,
                "base": base,
                "cards": cards,
                "combos": combos,
                "phones": phones,
                "skipped": skipped,
                "total": total,
                "selected_types": set(),
                "selected_format": "txt",
                "selected_domain": None,
                "sort_by_domain": False,
                "selected_validity": "all",
                "status": "analysed",
                "user_name": full_name,
                "user_username": uname,
            })
            log_queue(uid, doc.file_name, "queued")

        type_parts = []
        if cards:
            v = count_valid(cards)
            type_parts.append(f"💳 <b>Cards:</b> {len(cards)}  <i>(🟢 {v} · 🔴 {len(cards) - v})</i>")
        if combos:
            v = count_valid(combos)
            type_parts.append(f"🔑 <b>Emails:</b> {len(combos)}  <i>(🟢 {v} · 🔴 {len(combos) - v})</i>")
        if phones:
            v = count_valid(phones)
            type_parts.append(f"📱 <b>Phones:</b> {len(phones)}  <i>(🟢 {v} · 🔴 {len(phones) - v})</i>")
        if skipped:
            type_parts.append(f"🗑️ <b>Skipped:</b> {skipped}")

        await update.message.reply_text(
            f"✅ <b>File analysed!</b>\n"
            f"──────────────────\n"
            f"📁 <code>{doc.file_name}</code>\n\n"
            + "\n".join(type_parts) + "\n\n"
            f"🟢 = valid (passes checks) · 🔴 = fake\n\n"
            f"👇 <b>Filter & choose format:</b>",
            parse_mode="HTML",
            reply_markup=type_ikb(cards, combos, phones, uid, "txt"),
        )

    except Exception:
        log.exception("handle_document error")
        try:
            await thinking.delete()
        except Exception:
            pass
        await update.message.reply_text(
            "❌ <b>Something went wrong.</b> Please try again.",
            parse_mode="HTML", reply_markup=MAIN_KB,
        )


async def process_queue_item(ctx: ContextTypes.DEFAULT_TYPE, queue_item: dict) -> None:
    uid = queue_item["uid"]
    cards = queue_item["cards"]
    combos = queue_item["combos"]
    phones = queue_item["phones"]
    base = queue_item["base"]
    fmt = queue_item.get("selected_format", "txt")
    selected_types = queue_item.get("selected_types", set())
    domain = queue_item.get("selected_domain", None)
    sort_by_domain = queue_item.get("sort_by_domain", False)
    validity = queue_item.get("selected_validity", "all")
    user_name = queue_item.get("user_name", "")
    user_username = queue_item.get("user_username", "")

    # Type filter
    out_cards = cards if (not selected_types or "cards" in selected_types) else []
    out_combos = combos if (not selected_types or "combos" in selected_types) else []
    out_phones = phones if (not selected_types or "phones" in selected_types) else []

    # Domain + sorting
    if domain and out_combos:
        out_combos = filter_combos_by_domain(out_combos, domain)
    if sort_by_domain and out_combos:
        out_combos = sort_combos_by_domain(out_combos)

    # Valid / Fake filter
    out_cards = filter_entries(out_cards, validity)
    out_combos = filter_entries(out_combos, validity)
    out_phones = filter_entries(out_phones, validity)

    if not out_cards and not out_combos and not out_phones:
        await ctx.bot.send_message(
            uid,
            "⚠️ <b>No data left after filtering.</b>\nTry a different type, domain or valid/fake filter.",
            parse_mode="HTML",
        )
        return

    if fmt == "csv":
        output = build_csv(out_cards, out_combos, out_phones)
        ext = "csv"
    elif fmt == "xlsx":
        output_bytes = build_xlsx(out_cards, out_combos, out_phones)
        ext = "xlsx"
    else:
        output = build_txt(out_cards, out_combos, out_phones)
        ext = "txt"

    if fmt == "xlsx":
        buf = BytesIO(output_bytes)
    else:
        buf = BytesIO(output.encode("utf-8"))
    buf.name = f"{base}_filtered.{ext}"

    parts = []
    if out_cards:
        parts.append(f"💳 Cards: {len(out_cards)}")
    if out_combos:
        parts.append(f"🔑 Emails: {len(out_combos)}")
    if out_phones:
        parts.append(f"📱 Phones: {len(out_phones)}")
    if validity == "valid":
        parts.append("🧪 Filter: 🟢 Valid only")
    elif validity == "fake":
        parts.append("🧪 Filter: 🔴 Fake only")
    else:
        parts.append("🧪 Filter: All")
    if domain:
        parts.append(f"📧 Domain: {domain}")
    if sort_by_domain:
        parts.append("🔤 Sorted by domain")
    parts.append(f"📄 Format: {ext.upper()}")

    caption = (
        f"✅ <b>Done!</b>\n"
        f"──────────────────\n"
        + "\n".join(parts) + "\n\n"
        + "<i>— @yorifederation</i>"
    )

    try:
        await ctx.bot.send_document(
            uid,
            buf,
            filename=f"{base}_filtered.{ext}",
            caption=caption,
            parse_mode="HTML",
            reply_markup=result_ikb(uid),
        )

        n_valid = count_valid(out_cards) + count_valid(out_combos) + count_valid(out_phones)
        n_fake = (len(out_cards) + len(out_combos) + len(out_phones)) - n_valid
        upsert_user(uid, user_name, user_username,
                    len(out_cards), len(out_combos) + len(out_phones), n_valid, n_fake)

        log_queue(uid, queue_item["filename"], "done")
        log.info("Processed %s uid=%s cards=%d combos=%d phones=%d valid=%d fake=%d fmt=%s",
                 queue_item["filename"], uid, len(out_cards), len(out_combos),
                 len(out_phones), n_valid, n_fake, fmt)

    except Exception as e:
        log.error("Error sending file to uid=%s: %s", uid, e)
        log_queue(uid, queue_item["filename"], "error")
        await ctx.bot.send_message(
            uid,
            "❌ <b>Failed to send file.</b> Please try again.",
            parse_mode="HTML",
        )


def _get_user_queue_item(uid: int):
    with queue_lock:
        user_items = [q for q in file_queue if q["uid"] == uid and q["status"] == "analysed"]
        return user_items[-1] if user_items else None


async def _refresh_keyboard(query, item: dict) -> None:
    try:
        await query.message.edit_reply_markup(
            reply_markup=type_ikb(
                item["cards"], item["combos"], item["phones"], item["uid"],
                item["selected_format"], item["selected_types"],
                item["selected_domain"], item["sort_by_domain"],
                item["selected_validity"],
            )
        )
    except BadRequest:
        pass


async def handle_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    uid = query.from_user.id
    data = query.data or ""
    await query.answer()

    if data.startswith("stats:"):
        row = get_user(uid)
        if not row:
            await query.message.reply_text("📊 No stats yet — drop a .txt file to get started!")
        else:
            await query.message.reply_text(user_stats_text(row), parse_mode="HTML")

    elif data == "help":
        await query.message.reply_text(HELP_TEXT, parse_mode="HTML")

    elif data == "owner:stats":
        if uid != OWNER_ID:
            await query.message.reply_text("⛔ Owner only.")
            return
        await query.message.reply_text(global_stats_text(), parse_mode="HTML")

    elif data.startswith("fmt:"):
        parts = data.split(":")
        if len(parts) >= 3:
            fmt = parts[1]
            item = _get_user_queue_item(uid)
            if not item:
                await query.answer("⏳ No pending file. Send a new file.", show_alert=True)
                return
            item["selected_format"] = fmt
            await _refresh_keyboard(query, item)
            await query.answer(f"✅ Format: {fmt.upper()}")

    elif data.startswith("val:"):
        parts = data.split(":")
        if len(parts) >= 3:
            val = parts[1]  # valid | fake | all
            item = _get_user_queue_item(uid)
            if not item:
                await query.answer("⏳ No pending file. Send a new file.", show_alert=True)
                return
            item["selected_validity"] = val
            await _refresh_keyboard(query, item)
            label = {"valid": "Valid only", "fake": "Fake only", "all": "All entries"}.get(val, val)
            await query.answer(f"✅ Filter: {label}")

    elif data.startswith("sel:"):
        parts = data.split(":")
        if len(parts) >= 3:
            type_name = parts[1]
            item = _get_user_queue_item(uid)
            if not item:
                await query.answer("⏳ No pending file. Send a new file.", show_alert=True)
                return

            if type_name == "all":
                item["selected_types"] = set()
            elif type_name in item["selected_types"]:
                item["selected_types"].discard(type_name)
            else:
                item["selected_types"].add(type_name)

            await _refresh_keyboard(query, item)
            await query.answer(f"✅ Types: {', '.join(item['selected_types']) or 'All'}")

    elif data.startswith("seld:"):
        parts = data.split(":")
        if len(parts) >= 3:
            domain = parts[1]
            item = _get_user_queue_item(uid)
            if not item:
                await query.answer("⏳ No pending file. Send a new file.", show_alert=True)
                return
            item["selected_domain"] = domain if item.get("selected_domain") != domain else None
            await _refresh_keyboard(query, item)
            await query.answer(f"✅ Domain: {item['selected_domain'] or 'All'}")

    elif data.startswith("sels:domain:"):
        item = _get_user_queue_item(uid)
        if not item:
            await query.answer("⏳ No pending file. Send a new file.", show_alert=True)
            return
        item["sort_by_domain"] = not item.get("sort_by_domain", False)
        await _refresh_keyboard(query, item)
        await query.answer(f"✅ Sort: {'ON' if item['sort_by_domain'] else 'OFF'}")

    elif data.startswith("gen:"):
        parts = data.split(":")
        if len(parts) >= 3:
            item = _get_user_queue_item(uid)
            if not item:
                await query.answer("⏳ No pending file. Send a new file.", show_alert=True)
                return

            item["status"] = "processing"
            await process_queue_item(ctx, item)

            with queue_lock:
                if item in file_queue:
                    file_queue.remove(item)


async def handle_message(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    text = update.message.text or ""
    uid = update.effective_user.id

    if text == "📊 My Stats":
        row = get_user(uid)
        if not row:
            await update.message.reply_text(
                "📊 No stats yet — drop a .txt file to get started!",
                reply_markup=MAIN_KB,
            )
        else:
            await update.message.reply_text(
                user_stats_text(row), parse_mode="HTML", reply_markup=MAIN_KB
            )

    elif text == "ℹ️ Help":
        await update.message.reply_text(HELP_TEXT, parse_mode="HTML", reply_markup=MAIN_KB)

    elif text == "🏷️ About":
        await update.message.reply_text(ABOUT_TEXT, parse_mode="HTML", reply_markup=MAIN_KB)

    elif not text.startswith("/"):
        await update.message.reply_text(
            "👋 <b>Drop a .txt file</b> and I'll analyse & filter it instantly.\n\n"
            "🟢 valid · 🔴 fake — with reasons for every entry.",
            parse_mode="HTML", reply_markup=MAIN_KB,
        )


# ── Entry point ─────────────────────────────────────────────────────────────────

def main() -> None:
    init_db()

    if not TOKEN:
        log.error("TELEGRAM_BOT_TOKEN is not set. Export it and restart.")
        raise SystemExit("TELEGRAM_BOT_TOKEN environment variable is required. "
                         "For an offline demo run: python validator.py sample.txt")

    threading.Thread(target=_start_health, daemon=True).start()

    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("stats", cmd_stats))
    app.add_handler(CommandHandler("broadcast", cmd_broadcast))
    app.add_handler(CommandHandler("queue", cmd_queue))
    app.add_handler(CommandHandler("check", cmd_check))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(CallbackQueryHandler(handle_callback))

    log.info("Bot starting on port %d", PORT)

    if WEBHOOK_URL:
        app.run_webhook(
            listen="0.0.0.0",
            port=PORT,
            webhook_url=WEBHOOK_URL,
            drop_pending_updates=True,
        )
    else:
        app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--check":
        from validator import run_cli
        run_cli(sys.argv[2])
    else:
        main()
