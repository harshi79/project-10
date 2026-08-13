# Yori Prime Filter Bot

A Telegram bot that reads almost any `.txt` file, auto-detects each line's type,
cleans it, deduplicates, and lets you **filter** the output in many ways. Built
as a cybersecurity-team school project.

> ⚠️ **Educational use only.** All data is fake/test data. The bot only
> *parses, cleans and filters* text — it never contacts any bank, payment
> network or live service.

---

## Supported types (auto-detected)

| Type | Example |
|------|---------|
| 💳 Cards | `4111111111111111\|05\|33\|496 — 🇦🇪 AE` |
| 🔑 Email combos | `user@gmail.com:Password1` |
| 📱 Phone combos | `+919876543210:Password1` |
| 🌐 Proxies | `1.2.3.4:8080` · `socks5://u:p@1.2.3.4:1080` · `host:port:user:pass` |
| 🔗 URLs | `https://example.com/page` |
| 💰 Crypto | BTC (`bc1…`, `1…`, `3…`) and ETH (`0x…`) addresses |

---

## Filters

| Filter | Applies to |
|--------|------------|
| 💳 Brand | Cards — brand from BIN (Visa, Mastercard, Amex, Discover, JCB, UnionPay, Diners, Maestro) |
| 🌍 Country | All types — from the country tag at line end (`… — 🇦🇪 AE`) |
| 📞 Code | Phones — country calling code (`+91`, `+1`, `+44` …) |
| 🌐 Protocol | Proxies — `http` / `socks4` / `socks5` |
| 🔌 Port | Proxies — by port number |
| 💰 Network | Crypto — `BTC` / `ETH` |
| 📧 Domain | Emails + URLs — by domain, plus optional sort-by-domain |
| Type | Include only Cards / Emails / Phones / Proxies / URLs / Crypto / All |

Filter buttons are **cycle buttons** — tap to move through the options and back
to "All".

## Cleaning

- Auto-removes junk (bare `t.me/` links, `tg://` links, telegram headers,
  `#` comments, blank lines)
- **Deduplication** across every type (case-insensitive for emails)

## Output

- **TXT** — cleaned, sectioned by type
- **CSV** — per-type columns (Brand, Country, Code, Protocol, Port, Network, Domain…)
- **Excel (.xlsx)** — real workbook with a Summary sheet + one sheet per type

## Upload limit

**20 MB** maximum (Telegram's bot download limit). Bigger files are rejected
with a clear message.

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

3. Send the bot any `.txt` file, then tap the buttons to filter and export.

### Commands

| Command | Description |
|---------|-------------|
| `/start` | Welcome + instructions |
| `📊 My Stats` | Your totals (files, lines, cards, combos, proxies, urls, crypto) |
| `/stats` | Global stats (owner only) |
| `/broadcast <msg>` | Message all users (owner only) |
| `/queue` | Recent processing history |

---

## Owner

👑 **Owner:** [https://t.me/yorichiiprime](https://t.me/yorichiiprime)

---

## Project layout

```
main.py          Telegram bot (parse → clean → filter → export)
sample.txt       Fake/test data covering every supported type
requirements.txt
```

*— Yori Prime*
