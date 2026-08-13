"""
Yori Cleaner Bot — Telegram advanced filtering bot.

Parses .txt files of cards / email combos / phone combos (mixed files too),
cleans them, deduplicates, and lets you FILTER the output:

  • Type filter      — Cards / Emails / Phones / All
  • Brand filter     — card brand from BIN (Visa, Mastercard, Amex, …)
  • Country filter   — country tag at end of line (e.g. "… — 🇦🇪 AE")
  • Phone code filter— country code (e.g. +91, +1, +44)
  • Domain filter    — email domain + sort by domain
  • Junk removal     — URLs, telegram headers, comments, blank lines
  • Deduplication    — exact + case-insensitive duplicates

Output: TXT, CSV, Excel (.xlsx)

All data used is fake/test data — this is an educational project about how
filtering works, nothing else. No account or card is ever contacted.
"""

import os
import re
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
    ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
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

# Country tag at the end of a line: "… — 🇦🇪 AE"  (flag optional)
COUNTRY_RE = re.compile(
    r"[\u2014\u2013\-]\s*(?:[\U0001F1E6-\U0001F1FF]{2}\s*)?([A-Za-z]{2})\s*$"
)

# Common ITU country calling codes (E.164). Used for longest-prefix matching so
# "+14155552671" → "+1" (US), "+919876543210" → "+91" (India), etc.
COUNTRY_CODES = {
    "1", "7", "20", "27", "30", "31", "32", "33", "34", "36", "39", "40",
    "41", "43", "44", "45", "46", "47", "48", "49", "51", "52", "53", "54",
    "55", "56", "57", "58", "60", "61", "62", "63", "64", "65", "66", "81",
    "82", "84", "86", "90", "91", "92", "93", "94", "95", "98", "211", "212",
    "213", "216", "218", "220", "221", "222", "223", "224", "225", "226",
    "227", "228", "229", "230", "231", "232", "233", "234", "235", "236",
    "237", "238", "239", "240", "241", "242", "243", "244", "245", "246",
    "247", "248", "249", "250", "251", "252", "253", "254", "255", "256",
    "257", "258", "260", "261", "262", "263", "264", "265", "266", "267",
    "268", "269", "290", "291", "297", "298", "299", "350", "351", "352",
    "353", "354", "355", "356", "357", "358", "359", "370", "371", "372",
    "373", "374", "375", "376", "377", "378", "380", "381", "382", "383",
    "385", "386", "387", "389", "420", "421", "423", "500", "501", "502",
    "503", "504", "505", "506", "507", "508", "509", "590", "591", "592",
    "593", "594", "595", "596", "597", "598", "599", "670", "672", "673",
    "674", "675", "676", "677", "678", "679", "680", "681", "682", "683",
    "685", "686", "687", "688", "689", "690", "691", "692", "850", "852",
    "853", "855", "856", "880", "886", "960", "961", "962", "963", "964",
    "965", "966", "967", "968", "970", "971", "972", "973", "974", "975",
    "976", "977", "992", "993", "994", "995", "996", "998",
}
CODE_KEYS = sorted(COUNTRY_CODES, key=lambda c: (-len(c), -int(c)))

# Brand detection from BIN (first digits) — used for FILTERING, not validation.
BRAND_RULES = [
    ("Visa",             lambda n: n[0] == "4"),
    ("Mastercard",       lambda n: len(n) >= 2 and 51 <= int(n[:2]) <= 55),
    ("Mastercard",       lambda n: len(n) >= 4 and 2221 <= int(n[:4]) <= 2720),
    ("American Express", lambda n: n.startswith(("34", "37"))),
    ("Discover",         lambda n: len(n) >= 6 and 622126 <= int(n[:6]) <= 622925),
    ("Discover",         lambda n: n.startswith(("6011", "65"))),
    ("JCB",              lambda n: len(n) >= 4 and 3528 <= int(n[:4]) <= 3589),
    ("UnionPay",         lambda n: n.startswith("62")),
    ("Diners Club",      lambda n: n.startswith(("300", "301", "302", "303", "304", "305", "36", "38"))),
    ("Maestro",          lambda n: n.startswith(("50", "56", "57", "58", "63"))),
]


def card_brand(number: str) -> str:
    n = "".join(c for c in number if c.isdigit())
    for brand, rule in BRAND_RULES:
        if n and rule(n):
            return brand
    return "Unknown"


def extract_country(raw: str):
    """Return (country_code_or_None, line_without_tag)."""
    m = COUNTRY_RE.search(raw)
    if m:
        return m.group(1).upper(), raw[:m.start()].rstrip()
    return None, raw


def phone_code(phone: str):
    m = re.match(r"^\+(\d+)", phone.strip())
    if not m:
        return None
    digits = m.group(1)
    for code in CODE_KEYS:
        if digits.startswith(code):
            return f"+{code}"
    return None


def analyse_line(raw: str):
    """Classify a single line.

    Returns one of:
      ("card",  num, mm, yy, cvv, country)
      ("combo", email, pw, country)
      ("phone", phone, pw, country, code)
      None
    """
    t = raw.strip()
    if not t or t.startswith("#"):
        return None
    if JUNK_RE.match(t) or TG_HEAD_RE.match(t):
        return None

    country, t = extract_country(t)
    t = t.strip()
    if not t:
        return None

    m = CARD_RE.match(t)
    if m:
        num, mm, yy, cvv = m.groups()
        return ("card", num, mm, yy, cvv, country)

    m = EMAIL_SEP_RE.match(t)
    if m:
        email, pw = m.group(1).strip(), m.group(2).strip()
        if EMAIL_RE.match(email) and pw:
            return ("combo", email, pw, country)

    if "@" in t:
        m = EMAIL_SPACE_RE.match(t)
        if m:
            email, pw = m.group(1).strip(), m.group(2).strip()
            if EMAIL_RE.match(email) and pw:
                return ("combo", email, pw, country)

    m = PHONE_SEP_RE.match(t)
    if m:
        phone, pw = m.group(1).strip(), m.group(2).strip()
        if pw:
            return ("phone", phone, pw, country, phone_code(phone))

    m = PHONE_SPACE_RE.match(t)
    if m:
        phone, pw = m.group(1).strip(), m.group(2).strip()
        if pw:
            return ("phone", phone, pw, country, phone_code(phone))

    return None


def analyse_file(content: str) -> dict:
    cards:  list[dict] = []
    combos: list[dict] = []
    phones: list[dict] = []
    seen_cards, seen_combos, seen_phones = set(), set(), set()
    skipped = 0
    total_nonempty = sum(
        1 for l in content.splitlines()
        if l.strip() and not l.lstrip().startswith("#")
    )

    for raw in content.splitlines():
        r = analyse_line(raw)
        if r is None:
            if raw.strip() and not raw.lstrip().startswith("#"):
                skipped += 1
            continue

        if r[0] == "card":
            _, num, mm, yy, cvv, country = r
            value = f"{num}|{mm}|{yy}|{cvv}"
            if value in seen_cards:
                skipped += 1
                continue
            seen_cards.add(value)
            cards.append({
                "value": value, "num": num, "mm": mm, "yy": yy, "cvv": cvv,
                "brand": card_brand(num), "country": country,
            })

        elif r[0] == "combo":
            _, email, pw, country = r
            key = f"{email.lower()}:::{pw}"
            if key in seen_combos:
                skipped += 1
                continue
            seen_combos.add(key)
            combos.append({
                "email": email, "pw": pw,
                "domain": email.rsplit("@", 1)[-1].lower(), "country": country,
            })

        elif r[0] == "phone":
            _, phone, pw, country, code = r
            key = f"{phone}:::{pw}"
            if key in seen_phones:
                skipped += 1
                continue
            seen_phones.add(key)
            phones.append({
                "phone": phone, "pw": pw, "code": code, "country": country,
            })

    return {
        "cards": cards, "combos": combos, "phones": phones,
        "skipped": skipped, "total": total_nonempty,
    }


# ── Filter helpers ──────────────────────────────────────────────────────────────

def get_domains(combos: list[dict]) -> list[str]:
    counts: dict[str, int] = defaultdict(int)
    for c in combos:
        counts[c["domain"]] += 1
    return sorted(counts.keys(), key=lambda d: (-counts[d], d))


def sort_combos_by_domain(combos: list[dict]) -> list[dict]:
    return sorted(combos, key=lambda c: (c["domain"], c["email"].lower()))


def filter_combos_by_domain(combos: list[dict], domain: str) -> list[dict]:
    return [c for c in combos if c["domain"] == domain.lower()]


def get_available_brands(cards: list[dict]) -> list[str]:
    return sorted({c["brand"] for c in cards})


def get_available_countries(cards: list[dict], combos: list[dict], phones: list[dict]) -> list[str]:
    codes = set()
    for c in cards:
        if c.get("country"):
            codes.add(c["country"])
    for c in combos:
        if c.get("country"):
            codes.add(c["country"])
    for c in phones:
        if c.get("country"):
            codes.add(c["country"])
    return sorted(codes)


def get_available_codes(phones: list[dict]) -> list[str]:
    return sorted({c["code"] for c in phones if c.get("code")})


def next_choice(current, options: list[str]):
    """Cycle through [None(All), *options]."""
    seq = [None] + list(options)
    try:
        i = seq.index(current)
    except ValueError:
        i = 0
    return seq[(i + 1) % len(seq)]


# ── Output Builders ─────────────────────────────────────────────────────────────

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
        rows.append(["Type", "Number", "Month", "Year", "CVV", "Brand", "Country"])
        for c in cards:
            rows.append(["card", c["num"], c["mm"], c["yy"], c["cvv"],
                         c["brand"], c.get("country") or ""])
    if combos:
        rows.append(["Type", "Email", "Password", "Domain", "Country"])
        for c in combos:
            rows.append(["combo", c["email"], c["pw"], c["domain"], c.get("country") or ""])
    if phones:
        rows.append(["Type", "Phone", "Password", "Code", "Country"])
        for c in phones:
            rows.append(["phone", c["phone"], c["pw"], c.get("code") or "", c.get("country") or ""])
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
    for col in ws.columns:
        letter = col[0].column_letter
        m = max((len(str(c.value)) for c in col if c.value is not None), default=8)
        ws.column_dimensions[letter].width = min(max(m + 2, 8), 42)


def _write_summary_sheet(ws, cards: list, combos: list, phones: list) -> None:
    ws.append(["Yori Cleaner — Filtered Output"])
    ws.append(["Generated", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")])
    ws.append([])
    ws.append(["Type", "Total"])
    for cell in ws[4]:
        _style_header(cell)
    for name, entries in (("Cards", cards), ("Emails", combos), ("Phones", phones)):
        ws.append([name, len(entries)])
    for letter, width in (("A", 12), ("B", 10)):
        ws.column_dimensions[letter].width = width


def build_xlsx(cards: list, combos: list, phones: list) -> bytes:
    wb = Workbook()

    ws = wb.active
    ws.title = "Summary"
    _write_summary_sheet(ws, cards, combos, phones)

    if cards:
        ws_c = wb.create_sheet("Cards")
        _write_table(ws_c,
                     ["Number", "Month", "Year", "CVV", "Brand", "Country"],
                     [[c["num"], c["mm"], c["yy"], c["cvv"], c["brand"],
                       c.get("country") or ""] for c in cards])
    if combos:
        ws_e = wb.create_sheet("Emails")
        _write_table(ws_e,
                     ["Email", "Password", "Domain", "Country"],
                     [[c["email"], c["pw"], c["domain"], c.get("country") or ""] for c in combos])
    if phones:
        ws_p = wb.create_sheet("Phones")
        _write_table(ws_p,
                     ["Phone", "Password", "Code", "Country"],
                     [[c["phone"], c["pw"], c.get("code") or "", c.get("country") or ""] for c in phones])

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
             sort: bool = False, brand: str | None = None,
             country: str | None = None, code: str | None = None) -> InlineKeyboardMarkup:
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

    # Row 3: Advanced filters (cycle buttons)
    filter_row = []
    if cards:
        filter_row.append(InlineKeyboardButton(
            f"💳 Brand: {brand or 'All'}", callback_data=f"brand:{uid}"))
    countries = get_available_countries(cards, combos, phones)
    if countries:
        filter_row.append(InlineKeyboardButton(
            f"🌍 Country: {country or 'All'}", callback_data=f"cty:{uid}"))
    if phones:
        filter_row.append(InlineKeyboardButton(
            f"📞 Code: {code or 'All'}", callback_data=f"code:{uid}"))
    if filter_row:
        rows.append(filter_row)

    # Row 4: Domain filter + sort
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
    "Send any <b>.txt</b> file. Every line is analysed and cleaned:\n\n"
    "💳 <b>Cards</b> — any separator:\n"
    "  <code>4111111111111111|05|33|496 — 🇦🇪 AE</code>\n"
    "  <code>4111111111111111 05 33 496</code>\n\n"
    "🔑 <b>Email combos</b>:\n"
    "  <code>user@gmail.com:Password1</code>\n"
    "  <code>user@gmail.com|Password1</code>\n\n"
    "📱 <b>Phone combos</b>:\n"
    "  <code>+12345678901:Password1</code>\n"
    "  <code>12345678901:Password1</code>\n\n"
    "🧰 <b>Filters (tap to cycle)</b>:\n"
    "  💳 Brand — Visa / Mastercard / Amex / …\n"
    "  🌍 Country — from the country tag at line end\n"
    "  📞 Code — phone country code (+91, +1, +44 …)\n"
    "  📧 Domain — email domain + sort by domain\n\n"
    "🗑️ Junk (URLs, telegram headers, comments) is removed automatically.\n"
    "📊 Choose output format: TXT / CSV / Excel\n\n"
    "<i>— @yorifederation</i>"
)

ABOUT_TEXT = (
    "<b>🏷️ @yorifederation Cleaner + Filter Bot</b>\n"
    "──────────────────\n"
    "⚡ Instant .txt file analysis\n"
    "🧠 Fuzzy per-line pattern detection\n"
    "💳 Filter by card brand (BIN)\n"
    "🌍 Filter by country tag\n"
    "📞 Filter by phone country code\n"
    "📧 Filter by email domain\n"
    "🔀 Mixed files in one pass\n"
    "🗑️ Automatic junk removal\n"
    "🔑 Deduplication built in\n"
    "📊 TXT / CSV / Excel output\n"
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
        f"  {i + 1}. {r['name'] or r['username']} — {r['lines']:,} lines"
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


# ── Handlers ────────────────────────────────────────────────────────────────────

async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE) -> None:
    name = update.effective_user.first_name or "there"
    await update.message.reply_text(
        f"⚡ <b>Yori Cleaner</b>  <i>by @yorifederation</i>\n"
        f"──────────────────────\n\n"
        f"Welcome, <b>{name}</b>.\n\n"
        f"Drop a <b>.txt</b> file and I will:\n"
        f"  🧠 Analyse every line automatically\n"
        f"  🗑️ Remove junk (URLs, headers, comments)\n"
        f"  🔑 Deduplicate entries\n"
        f"  💳 Filter by card brand\n"
        f"  🌍 Filter by country\n"
        f"  📞 Filter by phone country code\n"
        f"  📧 Filter by email domain\n"
        f"  📊 Export TXT / CSV / Excel\n\n"
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
        "🧠 <b>Analysing file…</b>", parse_mode="HTML"
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
                "selected_brand": None,
                "selected_country": None,
                "selected_code": None,
                "sort_by_domain": False,
                "status": "analysed",
                "user_name": full_name,
                "user_username": uname,
            })
            log_queue(uid, doc.file_name, "queued")

        type_parts = []
        if cards:
            type_parts.append(f"💳 <b>Cards:</b> {len(cards)}")
        if combos:
            type_parts.append(f"🔑 <b>Emails:</b> {len(combos)}")
        if phones:
            type_parts.append(f"📱 <b>Phones:</b> {len(phones)}")
        if skipped:
            type_parts.append(f"🗑️ <b>Skipped:</b> {skipped}")

        await update.message.reply_text(
            f"✅ <b>File analysed!</b>\n"
            f"──────────────────\n"
            f"📁 <code>{doc.file_name}</code>\n\n"
            + "\n".join(type_parts) + "\n\n"
            f"👇 <b>Apply filters & choose format:</b>",
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
    brand = queue_item.get("selected_brand", None)
    country = queue_item.get("selected_country", None)
    code = queue_item.get("selected_code", None)
    sort_by_domain = queue_item.get("sort_by_domain", False)
    user_name = queue_item.get("user_name", "")
    user_username = queue_item.get("user_username", "")

    # Type filter
    out_cards = cards if (not selected_types or "cards" in selected_types) else []
    out_combos = combos if (not selected_types or "combos" in selected_types) else []
    out_phones = phones if (not selected_types or "phones" in selected_types) else []

    # Brand filter
    if brand and out_cards:
        out_cards = [c for c in out_cards if c["brand"] == brand]

    # Country filter (applies to all types)
    if country:
        out_cards = [c for c in out_cards if c.get("country") == country]
        out_combos = [c for c in out_combos if c.get("country") == country]
        out_phones = [c for c in out_phones if c.get("country") == country]

    # Phone code filter
    if code and out_phones:
        out_phones = [c for c in out_phones if c.get("code") == code]

    # Domain filter + sort
    if domain and out_combos:
        out_combos = filter_combos_by_domain(out_combos, domain)
    if sort_by_domain and out_combos:
        out_combos = sort_combos_by_domain(out_combos)

    if not out_cards and not out_combos and not out_phones:
        await ctx.bot.send_message(
            uid,
            "⚠️ <b>No data left after filtering.</b>\nTry a different type, brand, country, code or domain.",
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
    filters_applied = []
    if brand:
        filters_applied.append(f"💳 {brand}")
    if country:
        filters_applied.append(f"🌍 {country}")
    if code:
        filters_applied.append(f"📞 {code}")
    if domain:
        filters_applied.append(f"📧 {domain}")
    if sort_by_domain:
        filters_applied.append("🔤 sorted")
    if filters_applied:
        parts.append("🧰 " + " · ".join(filters_applied))
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

        upsert_user(uid, user_name, user_username,
                    len(out_cards), len(out_combos) + len(out_phones))

        log_queue(uid, queue_item["filename"], "done")
        log.info("Processed %s uid=%s cards=%d combos=%d phones=%d fmt=%s",
                 queue_item["filename"], uid, len(out_cards), len(out_combos),
                 len(out_phones), fmt)

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
                item["selected_brand"], item["selected_country"], item["selected_code"],
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

    elif data.startswith("brand:"):
        item = _get_user_queue_item(uid)
        if not item:
            await query.answer("⏳ No pending file. Send a new file.", show_alert=True)
            return
        options = get_available_brands(item["cards"])
        item["selected_brand"] = next_choice(item.get("selected_brand"), options)
        await _refresh_keyboard(query, item)
        await query.answer(f"✅ Brand: {item['selected_brand'] or 'All'}")

    elif data.startswith("cty:"):
        item = _get_user_queue_item(uid)
        if not item:
            await query.answer("⏳ No pending file. Send a new file.", show_alert=True)
            return
        options = get_available_countries(item["cards"], item["combos"], item["phones"])
        item["selected_country"] = next_choice(item.get("selected_country"), options)
        await _refresh_keyboard(query, item)
        await query.answer(f"✅ Country: {item['selected_country'] or 'All'}")

    elif data.startswith("code:"):
        item = _get_user_queue_item(uid)
        if not item:
            await query.answer("⏳ No pending file. Send a new file.", show_alert=True)
            return
        options = get_available_codes(item["phones"])
        item["selected_code"] = next_choice(item.get("selected_code"), options)
        await _refresh_keyboard(query, item)
        await query.answer(f"✅ Code: {item['selected_code'] or 'All'}")

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
            "👋 <b>Drop a .txt file</b> and I'll clean & filter it instantly.\n\n"
            "Use the buttons below to navigate.",
            parse_mode="HTML", reply_markup=MAIN_KB,
        )


# ── Entry point ─────────────────────────────────────────────────────────────────

def main() -> None:
    init_db()

    if not TOKEN:
        log.error("TELEGRAM_BOT_TOKEN is not set. Export it and restart.")
        raise SystemExit("TELEGRAM_BOT_TOKEN environment variable is required.")

    threading.Thread(target=_start_health, daemon=True).start()

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
        app.run_webhook(
            listen="0.0.0.0",
            port=PORT,
            webhook_url=WEBHOOK_URL,
            drop_pending_updates=True,
        )
    else:
        app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
