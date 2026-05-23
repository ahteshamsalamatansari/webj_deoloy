# Deploying the Webjet Scraper as a Cron Job on Render.com

This guide walks you through hosting `webjet_brightdata.py` as a daily scheduled
cron job on [Render.com](https://render.com). The job scrapes Webjet flight prices
for PER→MJK and MJK→PER, writes results to a persistent Excel file, and resumes
safely if interrupted.

---

## Prerequisites

| Requirement | Notes |
|---|---|
| GitHub account | The repo must be pushed to GitHub — Render deploys from Git |
| Render account | Free tier works; a paid plan is needed for persistent disks |
| Bright Data credentials | `BD_BROWSER_USER` and `BD_BROWSER_PASS` from your Bright Data zone |

> **Cost note:** Persistent disks start at ~$0.25/GB/month and require at least
> the Starter plan. Free-tier cron jobs are possible without a disk, but then
> the output Excel file is lost after every run.

---

## How it works on Render

```
GitHub push  →  Render builds (build.sh)  →  Daily cron fires (render.yaml schedule)
                  pip install                   python webjet_brightdata.py
                  playwright install
                                           ↕
                                    /data  (persistent disk)
                                    └── webjet_results.xlsx   ← survives daily runs
                                    └── webjet_logs/
                                    └── webjet_debug/
```

The `render.yaml` at the repo root tells Render everything it needs: schedule,
build command, run command, disk, and environment variables.

---

## Step 1 — Push the repo to GitHub

```bash
git init          # if not already a git repo
git add .
git commit -m "initial commit"
git remote add origin https://github.com/YOUR_USER/YOUR_REPO.git
git push -u origin main
```

Make sure `build.sh` has the executable bit set:

```bash
git update-index --chmod=+x build.sh
git commit -m "mark build.sh executable"
git push
```

---

## Step 2 — Create a new Render service

1. Go to [dashboard.render.com](https://dashboard.render.com) and click **New +**.
2. Choose **Cron Job**.
3. Connect your GitHub account if prompted, then select your repository.
4. Render will detect `render.yaml` automatically and pre-fill the settings.
   Click **Apply from render.yaml** if shown, then click **Create Cron Job**.

> If Render does NOT detect `render.yaml`, fill in manually:
> - **Runtime:** Python
> - **Build Command:** `bash build.sh`
> - **Command:** `python webjet_brightdata.py`
> - **Schedule:** `0 22 * * *` (10 PM UTC = 6 AM Perth)

---

## Step 3 — Set secret environment variables

`BD_BROWSER_USER` and `BD_BROWSER_PASS` are marked `sync: false` in
`render.yaml`, meaning Render will NOT store them from the YAML file — you must
set them manually so they stay out of version control.

1. Open your cron job in the Render dashboard.
2. Go to **Environment** tab → **Add Environment Variable**.
3. Add both secrets:

| Key | Value | Secret? |
|---|---|---|
| `BD_BROWSER_USER` | `brd-customer-hl_xxxx-zone-cont_rex` | Yes |
| `BD_BROWSER_PASS` | `your_bright_data_password` | Yes |

All other variables (`WJ_OUTPUT_EXCEL`, `REX_TIMEZONE`, etc.) are already set
in `render.yaml` and will be applied automatically.

---

## Step 4 — Add the persistent disk

1. In the Render dashboard, go to your cron job → **Disks** tab.
2. Click **Add Disk** and fill in:
   - **Name:** `webjet-data`
   - **Mount Path:** `/data`
   - **Size:** 1 GB
3. Save. Render will re-deploy with the disk attached.

This ensures `webjet_results.xlsx` persists across daily runs and the resume
feature (`WJ_RESUME=1`) works correctly.

---

## Step 5 — Trigger a manual test run

Before waiting for the scheduled time, verify the setup works:

1. Go to your cron job in the Render dashboard.
2. Click **Trigger Run** (top-right button).
3. Open **Logs** to watch the scraper run in real time.

A successful run looks like:

```
✅ Bright Data Browser API connected.
📅 [1/84]  Monday, 26 May 2025
   🌐 Navigating to Webjet URL...
   ✅ Matrix table loaded (2 flight row(s))
   ✅ 2 flight row(s) saved.
...
📊 Output saved: /data/webjet_results.xlsx
```

---

## Step 6 — Download the output file

Render's persistent disk is not directly browsable from the dashboard. To
retrieve `webjet_results.xlsx` you have two options:

**Option A — Upload to Google Drive / S3 at the end of the run**
Add a post-scrape upload step. This does not require changing the scraper logic
— run it as a separate script or a shell wrapper in `startCommand`.

**Option B — Use Render Shell**
Render provides a shell for services. If your plan supports it:
1. Go to **Shell** tab in the dashboard.
2. Run: `cp /data/webjet_results.xlsx /tmp/download.xlsx`
3. Use `render cp` CLI to pull the file locally.

**Option C — Mount a cloud bucket**
Set `WJ_OUTPUT_EXCEL` to a path in a mounted S3-compatible bucket (requires
extra setup outside this project).

---

## Schedule reference

The cron schedule in `render.yaml` uses UTC:

```
"0 22 * * *"
 │  │  │ │ └── every day of week
 │  │  │ └──── every month
 │  │  └─────── every day of month
 │  └────────── 22:00 UTC  =  06:00 AWST (Perth, UTC+8, no DST)
 └───────────── minute 0
```

To change the run time, edit the `schedule` field in `render.yaml` and push.
Use [crontab.guru](https://crontab.guru) to verify your expression.

---

## Environment variables quick reference

| Variable | Default | Description |
|---|---|---|
| `BD_BROWSER_USER` | *(secret)* | Bright Data zone username |
| `BD_BROWSER_PASS` | *(secret)* | Bright Data zone password |
| `BD_BROWSER_HOST` | `brd.superproxy.io` | Bright Data proxy host |
| `BD_BROWSER_PORT` | `9222` | Bright Data CDP port |
| `WJ_OUTPUT_EXCEL` | `/data/webjet_results.xlsx` | Output Excel path |
| `WJ_LOG_DIR` | `/data/webjet_logs` | Log directory |
| `WJ_DEBUG_DIR` | `/data/webjet_debug` | Debug screenshots/HTML |
| `REX_TIMEZONE` | `Australia/Perth` | Timezone for date calculations |
| `WJ_TOTAL_DAYS` | `84` | Days ahead to scrape per route |
| `WJ_START_OFFSET_DAYS` | `1` | Start from N days ahead (1 = tomorrow) |
| `WJ_MAX_ATTEMPTS` | `3` | Retry attempts per date |
| `WJ_RESUME` | `1` | Skip already-completed dates (1=yes) |
| `WJ_JOB_TIMEOUT_SECONDS` | `240` | Hard timeout per date attempt |
| `WJ_INTER_ROUTE_DELAY_SECONDS` | `30` | Pause between routes (IP rotation) |

---

## Troubleshooting

**Build fails — `playwright install` error**
Render's build environment has all required OS libraries for Chromium.
If you see a missing-lib error, ensure `build.sh` runs
`playwright install --with-deps chromium` (the `--with-deps` flag installs OS
packages too). The current `build.sh` already includes this flag.

**`ModuleNotFoundError: No module named 'zoneinfo'`**
This requires Python 3.9+. The `render.yaml` pins `pythonVersion: "3.11.0"`.
If you configured the service manually, set the Python version in the dashboard
under **Settings → Python Version**.

**Bright Data connection refused**
Check that `BD_BROWSER_USER` and `BD_BROWSER_PASS` are set correctly in the
Render dashboard (Environment tab). The scraper prints a redacted WSS URL on
startup — verify the username prefix `brd-customer-` is present.

**Excel file missing after run**
The disk must be mounted before the first run. Check **Disks** tab in the
dashboard and confirm the mount path is `/data`. The env var
`WJ_OUTPUT_EXCEL` must point to `/data/webjet_results.xlsx`.

**Run exceeds Render's timeout**
Render cron jobs have a maximum runtime. With 84 dates × 2 routes the job
typically runs 3–6 hours. If the plan's timeout is lower, reduce `WJ_TOTAL_DAYS`
or split into two separate cron jobs (one per route) with staggered schedules.
