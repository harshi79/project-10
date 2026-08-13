# Yori Cleaner + Fake-Data Filter Bot

A Telegram bot that reads a `.txt` file of **cards**, **email combos** and
**phone combos**, then **validates every line** and filters out the fake ones.
Built as a cybersecurity-team school project.

> ⚠️ **Educational use only.** All data is fake/test data. The bot checks the
> *structure* of data (checksums, formats, heuristics). It never contacts any
> bank, payment network or live service, and it cannot tell you if a card is
> actually "live" — only whether the data looks well-formed or fake.

---

## What it detects

| Type | Checks |
|------|--------|
| 💳 Cards | Luhn (mod-10) checksum, BIN brand (Visa/Mastercard/Amex/Discover/JCB/…), length-per-brand, expiry date, CVV length |
| 📧 Emails | Format (RFC-ish), disposable / temp-mail domains (mailinator, tempmail, …) |
| 📱 Phones | Number format + length (7–15 digits) |
| 🔑 Passwords | Common-password list + strength scoring (weak / medium / strong) |

Every entry is labelled **🟢 VALID** (passes checks) or **🔴 FAKE** (filtered
out), with a human-readable **reason** for each fake.

---

## Features

- Auto-detects card / email-combo / phone-combo lines in **mixed files** in one pass
- **Valid / Fake / All** filtering, plus type, domain and sort-by-domain filters
- Output as **TXT**, **CSV**, or a real **Excel (.xlsx)** report with a summary
  sheet and colour-coded VALID/FAKE rows
- Deduplication, per-user stats, global stats (owner only), broadcast, queue history
- Quick single-line check: `/check 4111111111111111|05|33|496`
- Offline CLI demo (no internet, no dependencies)

---

## Offline demo (no setup needed)

The validation engine is dependency-free. Just run:

```bash
python validator.py sample.txt
```

It prints a VALID/FAKE report for every line and writes `sample_report.csv`.

---

## Running the Telegram bot

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

3. In Telegram, send the bot any `.txt` file (see `sample.txt` for the format),
   then tap the buttons to filter and export.

### Commands

| Command | Description |
|---------|-------------|
| `/start` | Welcome + instructions |
| `/check <line>` | Validate a single card / combo / phone line |
| `📊 My Stats` | Your totals (files, lines, valid, fake) |
| `/stats` | Global stats (owner only) |
| `/broadcast <msg>` | Message all users (owner only) |
| `/queue` | Recent processing history |

---

## Project layout

```
main.py        Telegram bot (parsing results → validation → export)
validator.py   Parsing + validation engine (offline demo lives here)
sample.txt     Fake/test data for demos
requirements.txt
```

---

## How the Luhn check works (for your viva/exam)

The Luhn algorithm is a checksum, not a "is this card live" check:

1. Starting from the **right**, double every second digit.
2. If doubling gives a number > 9, subtract 9.
3. Sum all digits.
4. The number is valid if the total is a multiple of 10.

Example: `4111 1111 1111 1111` (a public Visa **test** number) passes; change
one digit and it fails — that's how fakes with a random number get caught.

---

*— @yorifederation*
