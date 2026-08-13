"""
validator.py — parsing + validation engine for Yori Cleaner Bot.

Everything here works on LOCAL structure / checksum / heuristic rules only:

  • Cards   — Luhn checksum, BIN brand detection, length, expiry, CVV
  • Emails  — format check + disposable (temp-mail) domain detection
  • Phones  — number format + length
  • Passwords — common-password list + strength scoring

It NEVER contacts any bank, payment network or live service, and it cannot
tell you whether a card is actually "live" — only whether the data looks
well-formed or fake. This is for education / security-team use on test data.

Run it standalone (no dependencies, no internet):

    python validator.py sample.txt
"""

from __future__ import annotations

import re
import csv
import sys
import os
from datetime import datetime, timezone

# ── Line parsing ────────────────────────────────────────────────────────────────

CARD_RE = re.compile(
    r"^(\d{13,19})[\s|:;]+(\d{1,2})[\s|:;]+(\d{2,4})[\s|:;]+(\d{1,4})"
    r"(?:[\s]*[\u2014\u2013\-]+.*)?$"
)
EMAIL_RE       = re.compile(r"^[^\s@]+@[^\s@]+\.[^\s@]{2,}$")
EMAIL_SEP_RE   = re.compile(r"^([^\s@:;|]+@[^\s@:;|]+\.[^\s@:;|]{2,})[:;|](.+)$")
EMAIL_SPACE_RE = re.compile(r"^([^\s@]+@[^\s@]+\.[^\s@]{2,})\s+(\S+)$")
PHONE_SEP_RE   = re.compile(r"^(\+\d{7,15}|\d{7,12})[:;|](.+)$")
PHONE_SPACE_RE = re.compile(r"^(\+\d{7,15})\s+(\S+)$")
JUNK_RE        = re.compile(r"^(https?://|tg://|t\.me/)", re.I)
TG_HEAD_RE     = re.compile(r"^.{1,80},\s*\[\d{1,2}/\d{1,2}/\d{4}")
# Loose fallback: any "something@something : password" line, so malformed
# emails get caught and flagged FAKE instead of being silently skipped.
LOOSE_COMBO_RE = re.compile(r"^([^\s:;|]+@[^\s:;|]+)[:;|](\S+)$")


def analyse_line(raw: str):
    """Classify a single raw line.

    Returns one of:
      ("card",  num, mm, yy, cvv)
      ("combo", email, pw)
      ("phone", phone, pw)
      None        (empty / junk / unrecognised)
    """
    t = raw.strip()
    if not t:
        return None
    if JUNK_RE.match(t) or TG_HEAD_RE.match(t):
        return None

    m = CARD_RE.match(t)
    if m:
        num, mm, yy, cvv = m.groups()
        return ("card", num, mm, yy, cvv)

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

    # Malformed emails (e.g. bad@@domain.com, user@gmail) still count as combos
    # so they get validated and flagged as FAKE.
    if "@" in t and (":" in t or "|" in t or ";" in t):
        m = LOOSE_COMBO_RE.match(t)
        if m:
            email, pw = m.group(1).strip(), m.group(2).strip()
            if email and pw:
                return ("combo", email, pw)

    return None


# ── Card validation ─────────────────────────────────────────────────────────────

BRAND_LENGTHS = {
    "Visa":             {13, 16, 19},
    "Mastercard":       {16},
    "American Express": {15},
    "Discover":         {16, 19},
    "JCB":              {16, 17, 18, 19},
    "UnionPay":         {16, 17, 18, 19},
    "Diners Club":      {14},
    "Maestro":          {12, 13, 14, 15, 16, 17, 18, 19},
    "Unknown":          set(),
}


def card_brand(number: str) -> str:
    """Guess the card network from the BIN (first digits)."""
    n = "".join(c for c in number if c.isdigit())
    if not n:
        return "Unknown"
    if n[0] == "4":
        return "Visa"
    if len(n) >= 2 and 51 <= int(n[:2]) <= 55:
        return "Mastercard"
    if len(n) >= 4 and 2221 <= int(n[:4]) <= 2720:
        return "Mastercard"
    if n.startswith(("34", "37")):
        return "American Express"
    if len(n) >= 6 and 622126 <= int(n[:6]) <= 622925:
        return "Discover"
    if n.startswith(("6011", "65")):
        return "Discover"
    if len(n) >= 4 and 3528 <= int(n[:4]) <= 3589:
        return "JCB"
    if n.startswith("62"):
        return "UnionPay"
    if n.startswith(("300", "301", "302", "303", "304", "305", "36", "38")):
        return "Diners Club"
    if n.startswith("50") or n.startswith(("56", "57", "58", "63")):
        return "Maestro"
    return "Unknown"


def luhn_check(number: str) -> bool:
    """Standard Luhn (mod-10) checksum validation."""
    digits = [int(c) for c in number if c.isdigit()]
    if not digits:
        return False
    total = 0
    parity = len(digits) % 2
    for i, d in enumerate(digits):
        if i % 2 == parity:
            d *= 2
            if d > 9:
                d -= 9
        total += d
    return total % 10 == 0


def validate_card(num: str, mm: str, yy: str, cvv: str) -> dict:
    reasons: list[str] = []
    n = "".join(c for c in num if c.isdigit())
    brand = card_brand(num)

    if not n or len(n) != len(num):
        reasons.append("non-numeric characters")
    if not luhn_check(n):
        reasons.append("fails Luhn checksum")

    lengths = BRAND_LENGTHS.get(brand)
    if lengths and len(n) not in lengths:
        reasons.append(f"length {len(n)} not valid for {brand}")

    if not cvv.isdigit() or len(cvv) not in (3, 4):
        reasons.append("invalid CVV")

    month_ok = mm.isdigit() and 1 <= int(mm) <= 12
    if not month_ok:
        reasons.append("invalid month")

    year_ok = yy.isdigit() and len(yy) in (2, 4)
    if not year_ok:
        reasons.append("invalid year")
    elif month_ok:
        yr = int(yy) if len(yy) == 4 else 2000 + int(yy)
        now = datetime.now(timezone.utc)
        if yr < now.year or (yr == now.year and int(mm) < now.month):
            reasons.append("expired")

    return {"valid": not reasons, "reasons": reasons, "brand": brand}


# ── Email validation ────────────────────────────────────────────────────────────

DISPOSABLE_DOMAINS = {
    "mailinator.com", "tempmail.com", "10minutemail.com", "guerrillamail.com",
    "yopmail.com", "throwawaymail.com", "getnada.com", "temp-mail.org",
    "sharklasers.com", "maildrop.cc", "trashmail.com", "fakeinbox.com",
    "dispostable.com", "mintemail.com", "spamgourmet.com", "mailexpire.com",
    "mailnesia.com", "anonymbox.com", "tempinbox.com", "burnermail.io",
    "mailcatch.com", "mohmal.com", "emailondeck.com", "getairmail.com",
    "33mail.com", "dropmail.me", "mail-temp.com", "tempmail.ninja",
    "tmpmail.org", "onetempmail.com", "emailfake.com", "mailinator.net",
}


def validate_email(email: str) -> dict:
    reasons: list[str] = []
    e = email.strip().lower()
    kind = "valid"
    domain = e.rsplit("@", 1)[-1] if "@" in e else ""

    if not EMAIL_RE.match(e):
        reasons.append("malformed email")
        kind = "malformed"
    elif domain in DISPOSABLE_DOMAINS:
        reasons.append("disposable / temp-mail domain")
        kind = "disposable"

    return {"valid": not reasons, "reasons": reasons, "kind": kind, "domain": domain}


# ── Password validation ─────────────────────────────────────────────────────────

COMMON_PASSWORDS = {
    "123456", "password", "123456789", "12345678", "12345", "qwerty",
    "abc123", "111111", "123123", "admin", "letmein", "welcome", "monkey",
    "password1", "1234567890", "iloveyou", "000000", "1234", "123",
    "1q2w3e4r", "qwerty123", "dragon", "football", "passw0rd", "master",
    "hello", "freedom", "whatever", "qazwsx", "trustno1", "password123",
    "admin123", "welcome123", "login", "princess", "sunshine", "superman",
}


def password_strength(pw: str) -> str:
    score = 0
    if len(pw) >= 8:
        score += 1
    if len(pw) >= 12:
        score += 1
    if re.search(r"[a-z]", pw) and re.search(r"[A-Z]", pw):
        score += 1
    if re.search(r"\d", pw):
        score += 1
    if re.search(r"[^A-Za-z0-9]", pw):
        score += 1
    if score >= 4:
        return "strong"
    if score >= 2:
        return "medium"
    return "weak"


def validate_password(pw: str) -> dict:
    reasons: list[str] = []
    if pw.lower() in COMMON_PASSWORDS:
        reasons.append("common password")
    if len(pw) < 6:
        reasons.append("too short")
    return {"valid": not reasons, "reasons": reasons, "strength": password_strength(pw)}


# ── Phone validation ────────────────────────────────────────────────────────────

PHONE_RE = re.compile(r"^\+?\d{7,15}$")


def validate_phone(phone: str) -> dict:
    reasons: list[str] = []
    p = phone.strip()
    digits = re.sub(r"\D", "", p)
    if not PHONE_RE.match(p):
        reasons.append("bad phone format")
    if len(digits) < 7 or len(digits) > 15:
        reasons.append(f"unusual length ({len(digits)} digits)")
    return {"valid": not reasons, "reasons": reasons}


# ── Combined validators ─────────────────────────────────────────────────────────

def validate_combo(email: str, pw: str) -> dict:
    e = validate_email(email)
    p = validate_password(pw)
    reasons = list(e["reasons"]) + [f"password: {r}" for r in p["reasons"]]
    return {
        "valid": e["valid"] and p["valid"],
        "reasons": reasons,
        "domain": e["domain"],
        "pw_strength": p["strength"],
    }


def validate_phone_combo(phone: str, pw: str) -> dict:
    ph = validate_phone(phone)
    p = validate_password(pw)
    reasons = list(ph["reasons"]) + [f"password: {r}" for r in p["reasons"]]
    return {
        "valid": ph["valid"] and p["valid"],
        "reasons": reasons,
        "pw_strength": p["strength"],
    }


# ── File analysis ───────────────────────────────────────────────────────────────

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
        if not raw.strip() or raw.lstrip().startswith("#"):
            continue
        r = analyse_line(raw)
        if r is None:
            if raw.strip():
                skipped += 1
            continue

        if r[0] == "card":
            _, num, mm, yy, cvv = r
            key = f"{num}|{mm}|{yy}|{cvv}"
            if key in seen_cards:
                skipped += 1
                continue
            seen_cards.add(key)
            v = validate_card(num, mm, yy, cvv)
            cards.append({
                "num": num, "mm": mm, "yy": yy, "cvv": cvv,
                "value": key,
                "brand": v["brand"],
                "valid": v["valid"],
                "reasons": v["reasons"],
            })

        elif r[0] == "combo":
            _, email, pw = r
            key = f"{email.lower()}:::{pw}"
            if key in seen_combos:
                skipped += 1
                continue
            seen_combos.add(key)
            v = validate_combo(email, pw)
            combos.append({
                "email": email, "pw": pw,
                "domain": v["domain"],
                "valid": v["valid"],
                "reasons": v["reasons"],
                "pw_strength": v["pw_strength"],
            })

        elif r[0] == "phone":
            _, phone, pw = r
            key = f"{phone}:::{pw}"
            if key in seen_phones:
                skipped += 1
                continue
            seen_phones.add(key)
            v = validate_phone_combo(phone, pw)
            phones.append({
                "phone": phone, "pw": pw,
                "valid": v["valid"],
                "reasons": v["reasons"],
                "pw_strength": v["pw_strength"],
            })

    return {
        "cards": cards,
        "combos": combos,
        "phones": phones,
        "skipped": skipped,
        "total": total_nonempty,
    }


# ── Shared helpers ──────────────────────────────────────────────────────────────

def count_valid(entries: list[dict]) -> int:
    return sum(1 for e in entries if e["valid"])


def filter_entries(entries: list[dict], validity: str) -> list[dict]:
    if validity == "valid":
        return [e for e in entries if e["valid"]]
    if validity == "fake":
        return [e for e in entries if not e["valid"]]
    return list(entries)


def reason_str(reasons: list[str]) -> str:
    return "; ".join(reasons) if reasons else "—"


# ── Standalone CLI (offline demo — no dependencies, no internet) ────────────────

def run_cli(path: str) -> None:
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            content = fh.read()
    except OSError as e:
        print(f"Could not read {path}: {e}")
        raise SystemExit(1)

    res = analyse_file(content)
    cards, combos, phones = res["cards"], res["combos"], res["phones"]

    print("=" * 64)
    print("YORI CLEANER — fake-data filtering report (offline)")
    print("=" * 64)
    print(f"File   : {path}")
    print(f"Lines  : {res['total']} non-empty | skipped/unrecognised: {res['skipped']}")
    print("-" * 64)

    for name, entries in (("CARDS", cards), ("EMAILS", combos), ("PHONES", phones)):
        if not entries:
            continue
        v = count_valid(entries)
        f = len(entries) - v
        print(f"{name}: {len(entries)} total | {v} VALID | {f} FAKE")
        for e in entries[:25]:
            if name == "CARDS":
                label = f"{e['brand']:<14} {e['value']}"
            elif name == "EMAILS":
                label = f"{e['email']}  (pw: {e['pw_strength']})"
            else:
                label = f"{e['phone']}  (pw: {e['pw_strength']})"
            mark = "VALID" if e["valid"] else "FAKE "
            print(f"  [{mark}] {label}")
            if e["reasons"]:
                print(f"          -> {reason_str(e['reasons'])}")
        if len(entries) > 25:
            print(f"  ... and {len(entries) - 25} more")
        print("-" * 64)

    out_csv = os.path.splitext(path)[0] + "_report.csv"
    with open(out_csv, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["Type", "Value", "Extra", "Status", "Reason"])
        for e in cards:
            w.writerow(["card", e["value"], e["brand"],
                        "VALID" if e["valid"] else "FAKE", reason_str(e["reasons"])])
        for e in combos:
            w.writerow(["combo", f"{e['email']}:{e['pw']}", e["pw_strength"],
                        "VALID" if e["valid"] else "FAKE", reason_str(e["reasons"])])
        for e in phones:
            w.writerow(["phone", f"{e['phone']}:{e['pw']}", e["pw_strength"],
                        "VALID" if e["valid"] else "FAKE", reason_str(e["reasons"])])

    print(f"CSV report saved: {out_csv}")
    print("=" * 64)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python validator.py <file.txt>")
        print("Classifies every line as VALID or FAKE, prints a report")
        print("and writes <file>_report.csv. No internet or extra libs needed.")
        sys.exit(1)
    run_cli(sys.argv[1])
