"""
Yori Cleaner Bot — Telegram file processing bot
Supports: cards, email combos, phone combos, mixed files
Output: TXT, CSV, Excel
"""

import os
import re
import sqlite3
import logging
import threading
import asyncio
from io import BytesIO
from datetime import datetime
from collections import defaultdict
from http.server import HTTPServer, BaseHTTPRequestHandler

import aiohttp
from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    Bot,
)
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    filters,
    ContextTypes,
    ApplicationHandlerStop,
)
from telegram.error import Forbidden, BadRequest

# ── Config ──────────────────────────────────────────────────────────────────────

TOKEN       = os.environ["TELEGRAM_BOT_TOKEN"]
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

# ── Health server ─────────────────────────────────────────────────────────────

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
    def log_message(self, *a): pass

def _start_health():
    try:
        HTTPServer(("", PORT), _Health).serve_forever()
    except OSError:
        pass

threading.Thread(target=_start_health, daemon=True).start()

# ── Database ──────────────────────────────────────────────────────────────────

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

def upsert_user(uid: int, name: str, username: str, n_cards: int, n_combos: int) -> None:
    total = n_cards + n_combos
    ts    = datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
    with _conn() as c:
        c.execute("""
            INSERT INTO users (id, name, username, files, lines, cards, combos, last_seen)
            VALUES (?,?,?,1,?,?,?,?)
            ON CONFLICT(id) DO UPDATE SET
                name      = excluded.name,
                username  = excluded.username,
                files     = files  + 1,
                lines     = lines  + excluded.lines,
                cards     = cards  + excluded.cards,
                combos    = combos + excluded.combos,
                last_seen = excluded.last_seen
        """, (uid, name, username, total, n_cards, n_combos, ts))

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
        top5 = c.execute(
            "SELECT name, username, lines FROM users ORDER BY lines DESC LIMIT 5"
        ).fetchall()
    return total_users, total_files, total_lines, top5

def log_queue(uid: int, filename: str, status: str) -> None:
    ts = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
    with _conn() as c:
        c.execute("INSERT INTO queue_history (uid, filename, status, ts) VALUES (?,?,?,?)",
                  (uid, filename, status, ts))

def get_queue_history(uid: int, limit: int = 10):
    with _conn() as c:
        return c.execute(
            "SELECT filename, status, ts FROM queue_history WHERE uid=? ORDER BY id DESC LIMIT ?",
            (uid, limit)
        ).fetchall()

# ── Smart Line Analyser ─────────────────────────────────────────────────────────

CARD_RE = re.compile(
    r"^(\d{13,19})[\s|:;]+(\d{1,2})[\s|:;]+(\d{2,4})[\s|:;]+(\d{3,4})"
    r"(?:[\s]*[\u2014\u2013\-]+.*)?$"
)
EMAIL_RE       = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]{2,}$")
EMAIL_SEP_RE   = re.compile(r"^([^\s@:;|]+@[^\s@:;|]+\.[^\s@:;|]{2,})[:;|](.+)$")
EMAIL_SPACE_RE = re.compile(r"^([^\s@]+@[^\s@]+\.[^\s@]{2,})\s+(\S+)$")
PHONE_SEP_RE   = re.compile(r"^(\+\d{7,15}|\d{7,12})[:;|](.+)$")
PHONE_SPACE_RE = re.compile(r"^(\+\d{7,15})\s+(\S+)$")
JUNK_RE        = re.compile(r"^(https?://|tg://|t\.me/)", re.I)
TG_HEAD_RE     = re.compile(r"^.{1,80},\s*\[\d{1,2}/\d{1,2}/\d{4}")

def analyse_line(raw: str) -> tuple | None:
    t = raw.strip()
    if not t:
        return None
    if JUNK_RE.match(t) or TG_HEAD_RE.match(t):
        return None

    m = CARD_RE.match(t)
    if m:
        num, mm, yy, cvv = m.groups()
        return ("card", f"{num}|{mm}|{yy}|{cvv}")

    m = EMAIL_SEP_RE.match(t)
    if m:
        email, pw = m.group(1).strip(), m.group(2).strip()
        if EMAIL_RE.match(email) and pw:
            return ("combo", email, pw)

    if "@" in t:
        m = EMAIL_SPACE_RE.match(t)
        if m:
            email, pw = m.group(1).strip(), m.group(2).strip()
            if EMAIL_RE.match(email) and pw:
                return ("combo", email, pw)

    m = PHONE_SEP_RE.match(t)
    if m:
        phone, pw = m.group(1).strip(), m.group(2).strip()
        if pw:
            return ("phone", phone, pw)

    m = PHONE_SPACE_RE.match(t)
    if m:
        phone, pw = m.group(1).strip(), m.group(2).strip()
        if pw:
            return ("phone", phone, pw)

    return None

def analyse_file(content: str):
    cards:   list[str] = []
    combos:  list[tuple] = []   # (email, pass)
    phones:  list[tuple] = []   # (phone, pass)
    seen_cards  = set()
    seen_combos = set()
    seen_phones = set()
    skipped = 0
    total_nonempty = sum(1 for l in content.splitlines() if l.strip())

    for raw in content.splitlines():
        r = analyse_line(raw)
        if r is None:
            if raw.strip():
                skipped += 1
            continue

        if r[0] == "card":
            val = r[1]
            if val not in seen_cards:
                seen_cards.add(val)
                cards.append(val)
            else:
                skipped += 1

        elif r[0] == "combo":
            _, email, pw = r
            key = f"{email.lower()}:::{pw}"
            if key not in seen_combos:
                seen_combos.add(key)
                combos.append((email, pw))
            else:
                skipped += 1

        elif r[0] == "phone":
            _, phone, pw = r
            key = f"{phone}:::{pw}"
            if key not in seen_phones:
                seen_phones.add(key)
                phones.append((phone, pw))
            else:
                skipped += 1

    return cards, combos, phones, skipped, total_nonempty

# ── Output Builders ─────────────────────────────────────────────────────────────

def build_txt(cards: list, combos: list, phones: list) -> str:
    parts: list[str] = []
    sections = sum([bool(cards), bool(combos), bool(phones)])
    if sections > 1:
        if cards:
            parts += [f"━━━ CARDS ({len(cards)}) ━━━", *cards, ""]
        if combos:
            parts += [f"━━━ COMBOS ({len(combos)}) ━━━", *[f"{e}   {p}" for e, p in combos], ""]
        if phones:
            parts += [f"━━━ PHONES ({len(phones)}) ━━━", *[f"{p}   {pw}" for p, pw in phones]]
    else:
        parts += cards or [f"{e}   {p}" for e, p in combos] or [f"{p}   {pw}" for p, pw in phones]
    return "\n".join(parts) + WATERMARK

def build_csv(cards: list, combos: list, phones: list) -> str:
    lines = []
    if cards:
        lines.append("Type,Number,Month,Year,CVV")
        for c in cards:
            parts = c.split("|")
            if len(parts) == 4:
                lines.append(f"card,{','.join(parts)}")
    if combos:
        lines.append("Type,Email,Password")
        for e, p in combos:
            lines.append(f"combo,{e},{p}")
    if phones:
        lines.append("Type,Phone,Password")
        for p, pw in phones:
            lines.append(f"phone,{p},{pw}")
    return "\n".join(lines)

def build_xlsx(cards: list, combos: list, phones: list) -> bytes:
    """Build a simple Excel-compatible HTML table (works as .xlsx for most apps)"""
    html = """<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/1999/xhtml">
<head><meta charset="UTF-8"/><style>table{border-collapse:collapse;font-family:Arial;}th,td{border:1px solid #ccc;padding:6px;text-align:left;}th{background:#f0f0f0;}</style></head>
<body><table>"""
    rows = []
    if cards:
        rows.append("<tr><th>Type</th><th>Number</th><th>Month</th><th>Year</th><th>CVV</th></tr>")
        for c in cards:
            parts = c.split("|")
            if len(parts) == 4:
                rows.append(f"<tr><td>card</td><td>{parts[0]}</td><td>{parts[1]}</td><td>{parts[2]}</td><td>{parts[3]}</td></tr>")
    if combos:
        rows.append("<tr><th>Type</th><th>Email</th><th>Password</th></tr>")
        for e, p in combos:
            rows.append(f"<tr><td>combo</td><td>{e}</td><td>{p}</td></tr>")
    if phones:
        rows.append("<tr><th>Type</th><th>Phone</th><th>Password</th></tr>")
        for p, pw in phones:
            rows.append(f"<tr><td>phone</td><td>{p}</td><td>{pw}</td></tr>")
    html += "\n".join(rows)
    html += "</table></body></html>"
    return html.encode("utf-8")

def get_domains(combos: list) -> list[str]:
    """Extract unique domains from combos, sorted by count desc"""
    domain_counts: dict[str, int] = defaultdict(int)
    for email, _ in combos:
        domain = email.split("@")[-1].lower()
        domain_counts[domain] += 1
    return sorted(domain_counts.keys(), key=lambda d: (-domain_counts[d], d))

def sort_combos_by_domain(combos: list) -> list[tuple]:
    """Sort combos by domain then by email"""
    return sorted(combos, key=lambda x: (x[0].split("@")[-1].lower(), x[0].lower()))

def filter_combos_by_domain(combos: list, domain: str) -> list[tuple]:
    return [(e, p) for e, p in combos if e.split("@")[-1].lower() == domain.lower()]

# ── Keyboards ──────────────────────────────────────────────────────────────────

MAIN_KB = ReplyKeyboardMarkup(
    [["📊 My Stats", "ℹ️ Help"], ["🏷️ About"]],
    resize_keyboard=True, is_persistent=True,
)

def type_ikb(cards: list, combos: list, phones: list, uid: int, fmt: str = "txt",
             selected_types: set | None = None, domain: str | None = None,
             sort: bool = False) -> InlineKeyboardMarkup:
    """Inline keyboard: select format, select types, pick domains, then generate"""
    rows = []
    sel = selected_types or set()

    # Row 1: Format buttons
    rows.append([
        InlineKeyboardButton(f"📄 TXT {'✅' if fmt == 'txt' else ''}", callback_data=f"fmt:txt:{uid}"),
        InlineKeyboardButton(f"📊 CSV {'✅' if fmt == 'csv' else ''}", callback_data=f"fmt:csv:{uid}"),
        InlineKeyboardButton(f"📈 Excel {'✅' if fmt == 'xlsx' else ''}", callback_data=f"fmt:xlsx:{uid}"),
    ])

    # Row 2: Type selection (checkbox style)
    type_row = []
    if cards:
        type_row.append(InlineKeyboardButton(
            f"💳 Cards {'✅' if 'cards' in sel else ''}",
            callback_data=f"sel:cards:{uid}"))
    if combos:
        type_row.append(InlineKeyboardButton(
            f"🔑 Emails {'✅' if 'combos' in sel else ''}",
            callback_data=f"sel:combos:{uid}"))
    if phones:
        type_row.append(InlineKeyboardButton(
            f"📱 Phones {'✅' if 'phones' in sel else ''}",
            callback_data=f"sel:phones:{uid}"))
    if len(type_row) > 1:
        type_row.append(InlineKeyboardButton(
            f"🔀 All {'✅' if not sel else ''}",
            callback_data=f"sel:all:{uid}"))
    rows.append(type_row)

    # Domain buttons (if combos exist)
    if combos:
        domains = get_domains(combos)[:5]
        if domains:
            domain_row = [InlineKeyboardButton(
                f"📧 {d} {'✅' if domain == d else ''}",
                callback_data=f"seld:{d}:{uid}") for d in domains]
            rows.append(domain_row)
        rows.append([InlineKeyboardButton(
            f"🔤 Sort by domain {'✅' if sort else ''}",
            callback_data=f"sels:domain:{uid}")])

    # Generate button at bottom
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
    "Send any <b>.txt</b> file. Every line is analysed:\n\n"
    "💳 <b>Cards</b> — any separator:\n"
    "  <code>4111111111111111|05|33|496 — 🇦🇪 AE</code>\n"
    "  <code>4111111111111111 05 33 496</code>\n\n"
    "🔑 <b>Email combos</b>:\n"
    "  <code>user@gmail.com:Password1</code>\n"
    "  <code>user@gmail.com|Password1</code>\n\n"
    "📱 <b>Phone combos</b> (Telegram logins):\n"
    "  <code>+12345678901:Password1</code>\n"
    "  <code>12345678901:Password1</code>\n\n"
    "🔀 Mixed files handled in one pass.\n"
    "📊 Choose output format: TXT / CSV / Excel\n\n"
    "<i>— @yorifederation</i>"
)

ABOUT_TEXT = (
    "<b>🏷️ @yorifederation Cleaner Bot</b>\n"
    "──────────────────\n"
    "⚡ Instant .txt file analysis\n"
    "🧠 Fuzzy per-line pattern detection\n"
    "📱 Supports phone + email + card combos\n"
    "🔀 Mixed files in one pass\n"
    "📊 TXT / CSV / Excel output\n"
    "🔑 Deduplication built in\n"
    "💧 Auto-watermark on every output\n"
    "📊 Per-user stats (SQLite)\n\n"
    "<i>— @yorifederation</i>"
)

def user_stats_text(row) -> str:
    return (
        f"<b>📊 Your Stats</b>\n"
        f"──────────────────\n"
        f"👤 {row['name'] or 'Unknown'}  <i>{row['username']}</i>\n\n"
        f"📁 Files cleaned  <b>{row['files']:,}</b>\n"
        f"📝 Total lines    <b>{row['lines']:,}</b>\n"
        f"💳 Card lines     <b>{row['cards']:,}</b>\n"
        f"🔑 Combo lines    <b>{row['combos']:,}</b>\n\n"
        f"🕒 <i>{row['last_seen']}</i>\n\n"
        f"<i>— @yorifederation</i>"
    )

def global_stats_text() -> str:
    total_users, total_files, total_lines, top5 = get_global_stats()
    top_str = "\n".join(
        f"  {i+1}. {r['name'] or r['username']} — {r['lines']:,} lines"
        for i, r in enumerate(top5)
    ) or "  No data yet."
    return (
        f"<b>👑 Global Stats</b>\n"
        f"──────────────────\n"
        f"👥 Total users   <b>{total_users:,}</b>\n"
        f"📁 Total files   <b>{total_files:,}</b>\n"
        f"📝 Total lines   <b>{total_lines:,}</b>\n\n"
        f"🏆 <b>Top 5</b>\n{top_str}\n\n"
        f"<i>— @yorifederation</i>"
    )

# ── Handlers ───────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    name = update.effective_user.first_name or "there"
    await update.message.reply_text(
        f"⚡ <b>Yori Cleaner</b>  <i>by @yorifederation</i>\n"
        f"──────────────────────\n\n"
        f"Welcome, <b>{name}</b>.\n\n"
        f"Drop a <b>.txt</b> file and I will:\n"
        f"  🧠 Analyse every line automatically\n"
        f"  💳 Clean card data\n"
        f"  🔑 Format email combos\n"
        f"  📱 Format phone / Telegram combos\n"
        f"  📊 Let you choose TXT / CSV / Excel output\n"
        f"  🔀 Handle mixed files in one pass\n\n"
        f"No commands needed — just send the file.",
        parse_mode="HTML", reply_markup=MAIN_KB,
    )

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
    """Show user's queue status"""
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

    # Check queue
    with queue_lock:
        user_queued = sum(1 for q in file_queue if q["uid"] == uid)
    if user_queued >= 3:
        await update.message.reply_text(
            "⏳ <b>Queue full.</b> You have 3 files pending.\n"
            "Wait for one to finish before sending more.",
            parse_mode="HTML",
        )
        return

    # Download and analyse
    thinking = await update.message.reply_text(
        "🧠 <b>Analysing file…</b>", parse_mode="HTML"
    )

    try:
        tg_file = await ctx.bot.get_file(doc.file_id)
        async with aiohttp.ClientSession() as session:
            async with session.get(tg_file.file_path) as resp:
                content = await resp.text(encoding="utf-8", errors="replace")

        cards, combos, phones, skipped, total = analyse_file(content)
        await thinking.delete()

        if not cards and not combos and not phones:
            await update.message.reply_text(
                f"⚠️ <b>Nothing recognised</b> in <code>{doc.file_name}</code>.\n\n"
                f"Supported: cards, email combos, phone combos, or any mix.",
                parse_mode="HTML", reply_markup=MAIN_KB,
            )
            return

        base = doc.file_name.rsplit(".txt", 1)[0].rsplit(".TXT", 1)[0]

        # Add to queue
        full_name = " ".join(filter(None, [user.first_name, user.last_name]))
        uname = f"@{user.username}" if user.username else "—"
        with queue_lock:
            queue_item = {
                "uid": uid,
                "filename": doc.file_name,
                "base": base,
                "cards": cards,
                "combos": combos,
                "phones": phones,
                "skipped": skipped,
                "total": total,
                "content": content,
                "selected_types": set(),
                "selected_format": "txt",
                "selected_domain": None,
                "sort_by_domain": False,
                "status": "analysed",
                "user_name": full_name,
                "user_username": uname,
            }
            file_queue.append(queue_item)
            log_queue(uid, doc.file_name, "queued")

        # Show analysis with inline buttons
        type_parts = []
        if cards:
            type_parts.append(f"💳 <b>Cards:</b> {len(cards)}")
        if combos:
            type_parts.append(f"🔑 <b>Emails:</b> {len(combos)}")
        if phones:
            type_parts.append(f"📱 <b>Phones:</b> {len(phones)}")
        if skipped:
            type_parts.append(f"🗑️ <b>Skipped:</b> {skipped}")

        type_text = "\n".join(type_parts)

        await update.message.reply_text(
            f"✅ <b>File analysed!</b>\n"
            f"──────────────────\n"
            f"📁 <code>{doc.file_name}</code>\n\n"
            f"{type_text}\n\n"
            f"👇 <b>Choose what to include and format:</b>",
            parse_mode="HTML",
            reply_markup=type_ikb(cards, combos, phones, uid, "txt"),
        )

    except Exception:
        log.exception("handle_document error")
        await thinking.delete()
        await update.message.reply_text(
            "❌ <b>Something went wrong.</b> Please try again.",
            parse_mode="HTML", reply_markup=MAIN_KB,
        )

async def process_queue_item(ctx: ContextTypes.DEFAULT_TYPE, queue_item: dict) -> None:
    """Process a queued item and send output"""
    uid = queue_item["uid"]
    cards = queue_item["cards"]
    combos = queue_item["combos"]
    phones = queue_item["phones"]
    base = queue_item["base"]
    fmt = queue_item.get("selected_format", "txt")
    selected_types = queue_item.get("selected_types", set())
    domain = queue_item.get("selected_domain", None)
    sort_by_domain = queue_item.get("sort_by_domain", False)
    user_name = queue_item.get("user_name", "")
    user_username = queue_item.get("user_username", "")

    # Apply type filters
    out_cards = cards if (not selected_types or "cards" in selected_types) else []
    out_combos = combos if (not selected_types or "combos" in selected_types) else []
    out_phones = phones if (not selected_types or "phones" in selected_types) else []

    # Apply domain filter
    if domain and out_combos:
        out_combos = filter_combos_by_domain(out_combos, domain)

    # Apply sorting
    if sort_by_domain and out_combos:
        out_combos = sort_combos_by_domain(out_combos)

    if not out_cards and not out_combos and not out_phones:
        await ctx.bot.send_message(
            uid,
            "⚠️ <b>No data left after filtering.</b>\nTry selecting a different type or domain.",
            parse_mode="HTML",
        )
        return

    # Build output
    if fmt == "csv":
        output = build_csv(out_cards, out_combos, out_phones)
        ext = "csv"
    elif fmt == "xlsx":
        output_bytes = build_xlsx(out_cards, out_combos, out_phones)
        ext = "xlsx"
    else:
        output = build_txt(out_cards, out_combos, out_phones)
        ext = "txt"

    # Create buffer
    if fmt == "xlsx":
        buf = BytesIO(output_bytes)
    else:
        buf = BytesIO(output.encode("utf-8"))
    buf.name = f"{base}_cleaned.{ext}"

    # Caption
    parts = []
    if out_cards:
        parts.append(f"💳 Cards: {len(out_cards)}")
    if out_combos:
        parts.append(f"🔑 Combos: {len(out_combos)}")
    if out_phones:
        parts.append(f"📱 Phones: {len(out_phones)}")
    if domain:
        parts.append(f"📧 Domain: {domain}")
    if sort_by_domain:
        parts.append(f"🔤 Sorted by domain")
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
            filename=f"{base}_cleaned.{ext}",
            caption=caption,
            parse_mode="HTML",
            reply_markup=result_ikb(uid),
        )

        # Update stats
        upsert_user(uid, user_name, user_username, len(out_cards), len(out_combos) + len(out_phones))

        log_queue(uid, queue_item["filename"], "done")
        log.info("Processed %s for uid=%s cards=%d combos=%d phones=%d fmt=%s",
                 queue_item["filename"], uid, len(out_cards), len(out_combos), len(out_phones), fmt)

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
        # fmt:FMT:UID
        parts = data.split(":")
        if len(parts) >= 3:
            fmt = parts[1]
            item = _get_user_queue_item(uid)
            if not item:
                await query.answer("⏳ No pending file. Send a new file.", show_alert=True)
                return
            item["selected_format"] = fmt
            try:
                await query.message.edit_reply_markup(
                    reply_markup=type_ikb(
                        item["cards"], item["combos"], item["phones"], uid, fmt,
                        item["selected_types"], item["selected_domain"], item["sort_by_domain"]
                    )
                )
            except BadRequest:
                pass
            await query.answer(f"✅ Format: {fmt.upper()}")

    elif data.startswith("sel:"):
        # sel:TYPE:UID  — toggle type checkbox
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

            try:
                await query.message.edit_reply_markup(
                    reply_markup=type_ikb(
                        item["cards"], item["combos"], item["phones"], uid,
                        item["selected_format"], item["selected_types"],
                        item["selected_domain"], item["sort_by_domain"]
                    )
                )
            except BadRequest:
                pass
            await query.answer(f"✅ Types: {', '.join(item['selected_types']) or 'All'}")

    elif data.startswith("seld:"):
        # seld:DOMAIN:UID — select domain
        parts = data.split(":")
        if len(parts) >= 3:
            domain = parts[1]
            item = _get_user_queue_item(uid)
            if not item:
                await query.answer("⏳ No pending file. Send a new file.", show_alert=True)
                return
            item["selected_domain"] = domain if item.get("selected_domain") != domain else None
            try:
                await query.message.edit_reply_markup(
                    reply_markup=type_ikb(
                        item["cards"], item["combos"], item["phones"], uid,
                        item["selected_format"], item["selected_types"],
                        item["selected_domain"], item["sort_by_domain"]
                    )
                )
            except BadRequest:
                pass
            await query.answer(f"✅ Domain: {item['selected_domain'] or 'All'}")

    elif data.startswith("sels:domain:"):
        # sels:domain:UID — toggle sort by domain
        parts = data.split(":")
        if len(parts) >= 3:
            item = _get_user_queue_item(uid)
            if not item:
                await query.answer("⏳ No pending file. Send a new file.", show_alert=True)
                return
            item["sort_by_domain"] = not item.get("sort_by_domain", False)
            try:
                await query.message.edit_reply_markup(
                    reply_markup=type_ikb(
                        item["cards"], item["combos"], item["phones"], uid,
                        item["selected_format"], item["selected_types"],
                        item["selected_domain"], item["sort_by_domain"]
                    )
                )
            except BadRequest:
                pass
            await query.answer(f"✅ Sort: {'ON' if item['sort_by_domain'] else 'OFF'}")

    elif data.startswith("gen:"):
        # gen:FMT:UID — generate output
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
            "👋 <b>Drop a .txt file</b> and I'll clean it instantly.\n\n"
            "Use the buttons below to navigate.",
            parse_mode="HTML", reply_markup=MAIN_KB,
        )

# ── Entry point ─────────────────────────────────────────────────────────────────

def main() -> None:
    init_db()
    app = Application.builder().token(TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("stats", cmd_stats))
    app.add_handler(CommandHandler("broadcast", cmd_broadcast))
    app.add_handler(CommandHandler("queue", cmd_queue))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    app.add_handler(CallbackQueryHandler(handle_callback))

    log.info("Bot starting on port %d", PORT)

    if WEBHOOK_URL:
        # Webhook mode
        app.run_webhook(
            listen="0.0.0.0",
            port=PORT,
            webhook_url=WEBHOOK_URL,
            drop_pending_updates=True,
        )
    else:
        # Polling mode
        app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
