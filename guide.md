# Deploying the Webjet Scraper as a Cron Job on Render.com

This guide walks you through hosting `webjet_brightdata.py` as a daily scheduled
cron job on [Render.com](https://render.com). The job scrapes Webjet flight prices
for PER→MJK and MJK→PER, writes results to a persistent Excel file, and emails
a summary with the file attached to your client after every run.

---

## Why `build.sh` fails on Render

The original `build.sh` runs:
```
playwright install --with-deps chromium
```
The `--with-deps` flag tries to call `sudo` / `su` to install OS libraries, which
fails on Render's container environment (`su: Authentication failure`).

**Fix (already applied in `render.yaml`):**  
Render's Ubuntu image already ships the required system libraries. Drop `--with-deps`:
```
pip install -r requirements.txt && playwright install chromium
```
This is now the `buildCommand` in `render.yaml` — no `build.sh` needed.

---

## How it works on Render

```
GitHub push
    └─► Render build
            pip install -r requirements.txt
            playwright install chromium   (no sudo, works on Render)
                ↓
        Daily cron fires  (0 22 * * * UTC = 6 AM Perth)
            python webjet_brightdata.py
                ↓
            /data  (persistent disk — survives between runs)
            ├── webjet_results.xlsx  ← appended every day
            ├── webjet_logs/
            └── webjet_debug/
                ↓
            Email sent to client
            (Excel attached + run summary)
```

---

## Prerequisites

| Requirement | Notes |
|---|---|
| GitHub account | Render deploys from a Git repository |
| Render account | Free tier works; persistent disk requires a paid plan |
| Bright Data credentials | `BD_BROWSER_USER` and `BD_BROWSER_PASS` from your Bright Data zone |
| Gmail + App Password | See email setup section below |

> **Cost note:** Persistent disks cost ~$0.25/GB/month and require at least the
> Starter plan on Render. Without a disk the Excel file is lost after every run.

---

## Step 1 — Push the repo to GitHub

```bash
git add .
git commit -m "render deployment config"
git remote add origin https://github.com/YOUR_USER/YOUR_REPO.git
git push -u origin main
```

---

## Step 2 — Create a Cron Job on Render

1. Log in to [dashboard.render.com](https://dashboard.render.com).
2. Click **New +** → **Cron Job**.
3. Connect your GitHub account if prompted, then select your repository.
4. Render will detect `render.yaml` automatically. Click **Apply from render.yaml**
   when prompted, then click **Create Cron Job**.

If Render does NOT detect `render.yaml`, fill in manually:

| Field | Value |
|---|---|
| Runtime | Python |
| Build Command | `pip install -r requirements.txt && playwright install chromium` |
| Command | `python webjet_brightdata.py` |
| Schedule | `0 22 * * *` |

---

## Step 3 — Set Python version

The script requires Python 3.10+ (uses `X | Y` union type hints).  
`render.yaml` already sets `pythonVersion: "3.11.0"` AND the env var
`PYTHON_VERSION=3.11.0` so Render pins 3.11 regardless of its default.

If you configured the service manually, go to **Settings** → **Python Version**
and enter `3.11.0`, or add `PYTHON_VERSION = 3.11.0` under **Environment**.

---

## Step 4 — Set secret environment variables

The following variables are marked `sync: false` in `render.yaml`, meaning you
**must** set them in the Render dashboard — they are intentionally not committed
to Git.

Go to your cron job → **Environment** tab → **Add Environment Variable**.

### Bright Data (required)

| Key | Value | Secret? |
|---|---|---|
| `BD_BROWSER_USER` | Your Bright Data zone username | Yes |
| `BD_BROWSER_PASS` | Your Bright Data zone password | Yes |

### Email report (required for email to work)

| Key | Value | Secret? |
|---|---|---|
| `EMAIL_SENDER` | Gmail address that sends the report | Yes |
| `EMAIL_PASSWORD` | 16-char Google App Password | Yes |
| `EMAIL_RECIPIENT` | Client's email (comma-separate for multiple) | Yes |

All other variables (`WJ_OUTPUT_EXCEL`, `PYTHON_VERSION`, `EMAIL_SMTP_HOST`,
etc.) are already set in `render.yaml` with safe defaults.

---

## Step 5 — Set up Gmail App Password

You need a **Google App Password**, not your regular Gmail password.

1. Go to [myaccount.google.com/security](https://myaccount.google.com/security).
2. Enable **2-Step Verification** if not already on.
3. Search for **App Passwords** in the search bar at the top, or go to  
   **Security → 2-Step Verification → App Passwords**.
4. Select app: **Mail** / device: **Other** → type `Webjet Scraper` → click **Generate**.
5. Copy the **16-character password** (shown once, no spaces needed).
6. Paste it as `EMAIL_PASSWORD` in the Render dashboard.

> **Important:** Use the 16-character App Password, not your Gmail login password.
> App Passwords work even if you have 2FA enabled.

---

## Step 6 — Add the persistent disk

1. In the Render dashboard, go to your cron job → **Disks** tab.
2. Click **Add Disk**:
   - **Name:** `webjet-data`
   - **Mount Path:** `/data`
   - **Size:** 1 GB
3. Save. Render re-deploys with the disk attached.

This makes `webjet_results.xlsx` survive between daily runs and enables the
resume feature (`WJ_RESUME=1`).

---

## Step 7 — Trigger a manual test run

Before waiting for the 6 AM schedule, verify everything works:

1. Render dashboard → your cron job → **Trigger Run**.
2. Open **Logs** to watch in real time.

A successful run ends with:
```
  ✅  PER → MJK
  ✅  MJK → PER
══════════════════════════════════════════════════════════════
📧 Sending run report to client@example.com...
✅ Email report sent.
```

If email env vars are not set, you will see:
```
ℹ️  Email report skipped — EMAIL_SENDER / EMAIL_PASSWORD / EMAIL_RECIPIENT not set.
```

---

## What the email contains

**Subject:** `Webjet Scraper — Run 20250524 complete`

**Body (plain text):**
```
Webjet Scraper — Daily Run Summary
========================================
Run ID  : 20250524
Time    : 24 May 2025 06:12 AWST
Output  : /data/webjet_results.xlsx

Route Results
----------------------------------------
  PER -> MJK : OK
  MJK -> PER : OK

Row Counts (this run)
----------------------------------------
  NO_FARE_AVAILABLE: 14
  SUCCESS: 154
  TOTAL: 168
```

**Attachment:** `webjet_results.xlsx` (the full Excel file with all historical data).

---

## Schedule reference

```
"0 22 * * *"
 │  │  │ │ └── every day of week
 │  │  │ └──── every month
 │  │  └─────── every day of month
 │  └────────── 22:00 UTC  =  06:00 AWST (Perth, UTC+8, no daylight saving)
 └───────────── minute 0
```

Use [crontab.guru](https://crontab.guru) to check any expression. Edit
`schedule` in `render.yaml` and push to change the run time.

---

## Full environment variable reference

| Variable | Default in render.yaml | Description |
|---|---|---|
| `PYTHON_VERSION` | `3.11.0` | Python version Render uses |
| `BD_BROWSER_HOST` | `brd.superproxy.io` | Bright Data proxy host |
| `BD_BROWSER_PORT` | `9222` | Bright Data CDP port |
| `BD_BROWSER_USER` | *(secret)* | Bright Data zone username |
| `BD_BROWSER_PASS` | *(secret)* | Bright Data zone password |
| `WJ_OUTPUT_EXCEL` | `/data/webjet_results.xlsx` | Output Excel path |
| `WJ_LOG_DIR` | `/data/webjet_logs` | Log directory |
| `WJ_DEBUG_DIR` | `/data/webjet_debug` | Debug screenshots/HTML |
| `REX_TIMEZONE` | `Australia/Perth` | Timezone for date logic |
| `WJ_TOTAL_DAYS` | `84` | Days ahead to scrape per route |
| `WJ_START_OFFSET_DAYS` | `1` | Start from N days ahead (1 = tomorrow) |
| `WJ_MAX_ATTEMPTS` | `3` | Retry attempts per date |
| `WJ_FINAL_RETRY_ROUNDS` | `1` | End-of-route retry rounds |
| `WJ_RETRY_BACKOFF_SECONDS` | `8` | Base backoff between retries |
| `WJ_MAX_RETRY_BACKOFF_SECONDS` | `90` | Maximum backoff cap |
| `WJ_JOB_TIMEOUT_SECONDS` | `240` | Hard timeout per date attempt |
| `WJ_PAGE_LOAD_TIMEOUT` | `90` | Max seconds waiting for Webjet page |
| `WJ_RESUME` | `1` | Skip already-completed dates (1=yes) |
| `WJ_CHECKPOINT_EVERY` | `7` | Save Excel every N rows |
| `WJ_INTER_ROUTE_DELAY_SECONDS` | `30` | Pause between routes (IP rotation) |
| `EMAIL_SENDER` | *(secret)* | Gmail address sending the report |
| `EMAIL_PASSWORD` | *(secret)* | 16-char Google App Password |
| `EMAIL_RECIPIENT` | *(secret)* | Client email(s), comma-separated |
| `EMAIL_SMTP_HOST` | `smtp.gmail.com` | SMTP server |
| `EMAIL_SMTP_PORT` | `587` | SMTP port (STARTTLS) |
| `EMAIL_SUBJECT_PREFIX` | `Webjet Scraper` | Email subject prefix |

---

## Troubleshooting

**`su: Authentication failure` / `Failed to install browsers`**  
This was caused by `playwright install --with-deps` requiring root.
The fix is already applied: `buildCommand` now runs
`pip install -r requirements.txt && playwright install chromium` (no `--with-deps`).

**Render defaults to Python 3.14 (or another unexpected version)**  
Two guards are set: `pythonVersion: "3.11.0"` in `render.yaml` and the env var
`PYTHON_VERSION=3.11.0`. If Render still picks the wrong version, open the
service **Settings** tab and set the Python version explicitly to `3.11.0`.

**`ModuleNotFoundError: No module named 'zoneinfo'`**  
Only happens on Python < 3.9. Ensure Python 3.11 is selected (see above).

**Email not sending — `SMTPAuthenticationError`**  
- Check that `EMAIL_PASSWORD` is the 16-char App Password, not your Gmail login.
- Ensure 2-Step Verification is enabled on the Gmail account.
- Check that `EMAIL_SENDER` matches the Gmail account used to generate the App Password.

**Email not sending — `Connection refused` / timeout**  
- `EMAIL_SMTP_HOST` should be `smtp.gmail.com` and `EMAIL_SMTP_PORT` should be `587`.
- Port 465 (SSL) is not used here; the code uses STARTTLS on 587.

**Excel file missing after run**  
Confirm the disk is mounted: Render dashboard → **Disks** tab → mount path `/data`.
Check `WJ_OUTPUT_EXCEL` is set to `/data/webjet_results.xlsx`.

**Run takes too long / Render kills it**  
Render cron jobs have a max runtime per plan. With 84 dates × 2 routes the job
runs roughly 3–6 hours. If your plan's limit is lower, reduce `WJ_TOTAL_DAYS`
or split into two cron jobs (one per route) with staggered schedules.
