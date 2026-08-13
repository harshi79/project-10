# Yori Cleaner + Filter Bot

A Telegram bot that reads a `.txt` file of **cards**, **email combos** and
**phone combos**, cleans them up, and lets you **filter** the output in many
ways. Built as a cybersecurity-team school project.

> ⚠️ **Educational use only.** All data is fake/test data. The bot only
> *parses, cleans and filters* text — it never contacts any bank, payment
> network or live service.

---

## Features

### 🧹 Cleaning
- Auto-detects card / email-combo / phone-combo lines in **mixed files** in one pass
- Removes junk automatically (URLs, telegram headers, `#` comments, blank lines)
- **Deduplication** (exact + case-insensitive email dedupe)
- Normalises card format to `number|month|year|cvv`

### 🧰 Advanced filters
| Filter | What it does |
|--------|--------------|
| 💳 Brand | Filter cards by brand detected from the BIN (Visa, Mastercard, Amex, Discover, JCB, UnionPay, Diners, Maestro) |
| 🌍 Country | Filter by the country tag at the end of a line (`… — 🇦🇪 AE`) |
| 📞 Code | Filter phones by country code (`+91`, `+1`, `+44` …) |
| 📧 Domain | Filter emails by domain + optional sort-by-domain |
| 💳🔑📱 Type | Include only Cards / Emails / Phones / All |

### 📊 Output
- **TXT** — cleaned, optionally sectioned by type
- **CSV** — with Brand / Country / Code / Domain columns
- **Excel (.xlsx)** — real workbook with a Summary sheet and one sheet per type

### 👥 Extras
- Per-user stats (files, lines, cards, combos) in SQLite
- Global stats + Top 5 (owner only)
- `/broadcast` to all users (owner only)
- `/queue` history
- Health endpoint (`/health`) for hosting
- Auto-watermark on every TXT output

---

## Running the bot

1. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

2. Get a token from [@BotFather](https://t.me/BotFather) and run:

   ```bash
   export TELEGRAM_BOT_TOKEN="123456:ABC..."
   python main.py
   ```

   For a hosted setup you can also set `WEBHOOK_URL` and `PORT`.

3. Send the bot any `.txt` file (see `sample.txt` for the format), then tap the
   buttons to filter and export.

### Commands

| Command | Description |
|---------|-------------|
| `/start` | Welcome + instructions |
| `📊 My Stats` | Your totals (files, lines, cards, combos) |
| `/stats` | Global stats (owner only) |
| `/broadcast <msg>` | Message all users (owner only) |
| `/queue` | Recent processing history |

---

## Input format examples

```
4111111111111111|05|33|496 — 🇦🇪 AE      (card + country tag)
5555555555554444 11 30 123              (card, space separated)
alice@gmail.com:Tr0ub4dor&3             (email combo)
+919876543210:India#2026                (phone combo)
```

The country tag at the end is optional — if present it powers the 🌍 Country
filter, and the `+`-prefixed number powers the 📞 Code filter.

---

## Project layout

```
main.py          Telegram bot (parse → clean → filter → export)
sample.txt       Fake/test data for demos
requirements.txt
```

---

*— @yorifederation*
