# Pullback-in-Uptrend Stock Screener

Scans a ticker list for stocks that are:
- In a confirmed uptrend (50-day SMA above 200-day SMA, price above 200-day SMA)
- Pulling back in a **healthy** way (RSI < 45), while trying to filter out
  stocks that are actually **rolling over** into a downtrend.

## What makes a candidate PASS

All of the following must be true:

1. Golden cross intact: SMA50 > SMA200, price still above SMA200
2. RSI(14) < 45 -- some pullback, not extended/overbought
3. SMA50 is still rising (20-day lookback) -- confirms the uptrend itself
   hasn't started flattening or reversing
4. Volume during the pullback is not elevated vs. the 20-day average --
   avoids "distribution" pullbacks where big holders are dumping
5. Price is within 15% of its 3-month high -- avoids deep breakdowns that
   look like a "pullback" on RSI alone but are really a trend change
6. MACD line is still above zero -- confirms the medium-term trend hasn't
   genuinely turned negative (MACD dipping below its *signal* line during
   a dip is normal and shown as an informational flag only, not a fail)

Every ticker's row shows exactly which checks passed/failed so you can see
*why* something was excluded, not just a binary yes/no.

All strictness thresholds live in the `CONFIG` dict at the top of
`screener.py`. The ticker list is separate, in `config/tickers.txt` --
one ticker per line, `#` for comments. Add or remove tickers there
without touching the script at all (easy to edit directly from GitHub's
web UI on your phone: open the file → pencil icon → edit → commit).

## One-time setup (GitHub)

1. Create a new GitHub repo (public or private -- private works fine with
   the free Actions minutes for personal use).
2. Upload this whole folder's contents to the repo root, keeping the
   `.github/workflows/screener.yml` and `config/tickers.txt` paths
   exactly as-is.
3. Go to the repo's **Settings → Actions → General → Workflow permissions**
   and set it to **"Read and write permissions"** (needed so the workflow
   can commit the updated Excel file back to the repo).
4. (Optional, for email/SMS -- see below) add the secrets.
5. Go to the **Actions** tab, select "Daily Stock Screener," and click
   **"Run workflow"** once manually to confirm it works.
6. After that it runs automatically on the schedule in the workflow file
   (weekdays, 13:00 UTC by default), pushes an updated
   `Output/screener_results.xlsx` each time results change, and emails/
   texts you if configured.

## Email / SMS setup (optional)

Nothing is hardcoded in the script -- all credentials come from GitHub
repo secrets, so they're never visible in the code.

Go to **Settings → Secrets and variables → Actions → New repository
secret** and add:

| Secret name     | Example / notes |
|---|---|
| `SMTP_SERVER`   | `smtp.gmail.com` |
| `SMTP_PORT`     | `587` |
| `SMTP_USERNAME` | the Gmail (or other) address you're sending *from* |
| `SMTP_PASSWORD` | an **app password**, not your real password -- for Gmail: Google Account → Security → 2-Step Verification → App passwords |
| `EMAIL_TO`      | where the full results table lands, e.g. `you@example.com` (comma-separate for more than one) |
| `SMS_TO`        | *optional* -- a carrier email-to-SMS gateway address, e.g. `5551234567@vtext.com` (Verizon), `5551234567@txt.att.net` (AT&T), `5551234567@tmomail.net` (T-Mobile) |

**What each channel gets:**
- **Email** (`EMAIL_TO`): the full plain-text table (readable in any mail
  app, since most render plain text in a monospace-ish font) plus the
  `.xlsx` file attached.
- **SMS** (`SMS_TO`, optional): a short one-line summary of tickers that
  *passed* only -- e.g. `Screener 2026-08-30 PASS: AAPL, NVDA`. A real
  table doesn't render on SMS, so this is intentionally condensed.
  Carrier email-to-SMS gateways can be a bit unreliable/delayed and some
  carriers are phasing them out -- if it stops working, that's a carrier-side
  issue, not the script. A paid service like Twilio is the reliable
  alternative if you want guaranteed SMS delivery.

Leave `SMS_TO` (or all the SMTP secrets) unset if you don't want that
channel -- the script skips it silently rather than failing.

## Getting the file onto your phone / desktop

- Easiest: open the repo in the GitHub app or mobile browser, navigate to
  `Output/screener_results.xlsx`, and download it -- or "Open with" your
  Excel/Sheets app directly from GitHub's file view.
- You can also set up a free Zapier/IFTTT rule to email you the file on
  each push, if you don't want to check GitHub manually.

## Running it locally instead (no GitHub)

```bash
pip install -r requirements.txt
python screener.py
```

Output lands in `Output/screener_results.xlsx`.

## Known limitation

Yahoo Finance occasionally rate-limits or hiccups on `yfinance` calls.
If a run fails, it's almost always transient -- just re-run it (or wait
for the next scheduled run). This is a known quirk of the free data
source, not a bug in the logic.
