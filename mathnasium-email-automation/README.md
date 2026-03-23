# Mathnasium Email Automation

Nightly pipeline that detects completed assessments in Radius, generates parent emails via Claude, and presents them in a review dashboard for approval before sending via Gmail.

---

## How It Works

```
GitHub Actions (6am daily)
  → Playwright logs into Radius
  → Downloads Pre Assessment report (yesterday)
  → Downloads Student Report (all enrolled)
  → Downloads Learning Plan PDFs for triggered students
  → Pushes everything + manifest.json to Google Drive

Google Apps Script (hourly trigger, picks up new Drive files)
  → Parses both reports
  → Identifies: Level Up emails (score ≥ 90%) + Progress Check emails
  → Calls Claude API to generate two email versions per student
  → Writes to Google Sheet queue
  → Emails you: "X emails ready to review — [link]"

Review Dashboard (Apps Script Web App)
  → You open the link
  → See each pending email with both draft versions
  → Edit inline if needed
  → Click Send → goes via Gmail to all guardian addresses
  → Or Skip with optional reason
```

---

## Folder Structure

```
mathnasium-email-automation/
├── .github/
│   └── workflows/
│       └── nightly.yml          # GitHub Actions scheduler
├── scraper/
│   ├── scraper.js               # Playwright scraper
│   └── package.json
├── apps-script/
│   ├── Code.gs                  # Google Apps Script pipeline + web app
│   └── dashboard.html           # Review dashboard UI
└── README.md
```

---

## Setup — Step by Step

### Step 1: Google Drive Folder

1. Go to [drive.google.com](https://drive.google.com)
2. Create a new folder called `mathnasium-email-queue`
3. Copy the folder ID from the URL:
   `https://drive.google.com/drive/folders/THIS_IS_YOUR_FOLDER_ID`

---

### Step 2: Google Sheet (Email Queue)

1. Go to [sheets.google.com](https://sheets.google.com) and create a new blank spreadsheet
2. Name it `Mathnasium Email Queue`
3. Copy the Sheet ID from the URL:
   `https://docs.google.com/spreadsheets/d/THIS_IS_YOUR_SHEET_ID/edit`

---

### Step 3: Google Service Account (for GitHub → Drive upload)

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create a new project (or use existing)
3. Enable the **Google Drive API**
4. Go to **IAM & Admin → Service Accounts → Create Service Account**
5. Name it `mathnasium-scraper`
6. Click **Create and Continue**, skip role assignment, click **Done**
7. Click the service account → **Keys → Add Key → Create new key → JSON**
8. Download the JSON file — this is your `GOOGLE_SERVICE_ACCOUNT_KEY`
9. **Share your Drive folder** with the service account email
   (looks like `mathnasium-scraper@your-project.iam.gserviceaccount.com`)
   and give it **Editor** access

---

### Step 4: GitHub Repository Secrets

1. Go to your GitHub repo → **Settings → Secrets and variables → Actions → New repository secret**
2. Add these secrets:

| Secret Name | Value |
|---|---|
| `RADIUS_EMAIL` | Your Radius login email |
| `RADIUS_PASSWORD` | Your Radius login password |
| `GOOGLE_DRIVE_FOLDER_ID` | The folder ID from Step 1 |
| `GOOGLE_SERVICE_ACCOUNT_KEY` | The full contents of the JSON file from Step 3 |

---

### Step 5: Update Radius Selectors in the Scraper

The scraper has `⚠️ UPDATE` comments everywhere a selector needs to match your actual Radius UI.

**To find the correct selectors:**
1. Log into Radius in Chrome
2. Right-click any input field or button → **Inspect**
3. Look for `id`, `name`, `data-testid`, or unique class attributes
4. Replace the placeholder selectors in `scraper/scraper.js`

**Key places to update:**
- `loginToRadius()` — email field, password field, submit button, post-login element
- `downloadPreAssessmentReport()` — report page URL, date fields, export button
- `downloadStudentReport()` — report page URL, export button
- `downloadSingleLearningPlan()` — student page URL pattern, download button

**Tip:** Run the scraper locally first with `--dry-run` to test selectors without uploading:
```bash
cd scraper
RADIUS_EMAIL=you@email.com RADIUS_PASSWORD=yourpass \
  GOOGLE_DRIVE_FOLDER_ID=xxx GOOGLE_SERVICE_ACCOUNT_KEY='{}' \
  node scraper.js --dry-run
```

---

### Step 6: Google Apps Script Setup

1. Go to [script.google.com](https://script.google.com)
2. Click **New project**
3. Rename it to `Mathnasium Email Automation`
4. Replace the default `Code.gs` content with the contents of `apps-script/Code.gs`
5. Click **+** next to Files → **HTML** → name it `dashboard`
6. Paste the contents of `apps-script/dashboard.html`
7. Go to **Project Settings** (gear icon) → **Script Properties** → **Add script property**

Add these properties:

| Property | Value |
|---|---|
| `ANTHROPIC_API_KEY` | Your Claude API key from console.anthropic.com |
| `DRIVE_FOLDER_ID` | Same folder ID from Step 1 |
| `QUEUE_SHEET_ID` | Sheet ID from Step 2 |
| `NOTIFY_EMAIL` | Your email address for daily summaries |

8. Run `setupTrigger()` once:
   - In the editor, select `setupTrigger` from the function dropdown
   - Click **Run**
   - Authorize the script when prompted

9. Deploy as Web App:
   - Click **Deploy → New deployment**
   - Type: **Web app**
   - Execute as: **Me**
   - Who has access: **Anyone with the link** (or restrict to your organization)
   - Click **Deploy**
   - Copy the Web App URL — this is your dashboard link

---

### Step 7: Test the Full Pipeline

1. **Manual scraper test:**
   ```bash
   cd scraper
   npm install
   node scraper.js --dry-run
   ```

2. **Manual trigger test in Apps Script:**
   - In the script editor, run `processNewDriveFiles()` manually
   - Check the execution log for errors

3. **Trigger the GitHub Action manually:**
   - Go to your repo → **Actions → Nightly Email Queue Builder → Run workflow**

---

## Detection Logic

| Condition | Email Type |
|---|---|
| Row in Pre Assessment report with Date Taken = yesterday AND score ≥ 90% | Level Up email |
| Row in Pre Assessment report with Date Taken = yesterday AND score < 90% | Progress Check email |
| Student Report: Last Progress Check = yesterday AND no Pre Assessment taken same day | Progress Check email |

**Score threshold:** 90% (edit `LEVEL_UP_SCORE_THRESHOLD` in `scraper.js` to change)

---

## Adding a New Center

Edit the `GMAIL_SENDER_NAME` logic in `Code.gs` — it currently uses `item['Center']` to automatically route the sender name to `Mathnasium of Teaneck` or `Mathnasium of Englewood`.

---

## Troubleshooting

**Scraper fails to log in:**
- Check `debug/login-failed.png` screenshot
- Verify RADIUS_EMAIL and RADIUS_PASSWORD secrets are correct
- Update selectors in `loginToRadius()`

**Scraper can't find export button:**
- Check `debug/` folder for screenshots taken before the error
- Radius may have updated their UI — re-inspect the button and update the selector

**Learning plan download fails for some students:**
- The LP may not have been created yet in Radius
- These students will be flagged as "LP pending" in the dashboard
- You can attach the LP manually before sending, or send without it

**Apps Script quota exceeded:**
- Google Apps Script has a 6-minute execution limit per run
- If you have many students in one day, the Claude API calls may time out
- Solution: the pipeline will process what it can; unprocessed items appear next run

**Claude API returns wrong level name:**
- This is a known edge case when the assessment PDF rendering is ambiguous
- The dashboard lets you edit the email body before sending — correct it there
- Consider adding the student's name + correct level to the "Admin Note" field

---

## Costs

| Service | Cost |
|---|---|
| GitHub Actions | Free (2000 min/month free tier, this uses ~5 min/day) |
| Google Apps Script | Free |
| Google Drive / Sheets | Free (within storage limits) |
| Claude API | ~$0.01–0.03 per student email (Sonnet pricing) |
| Gmail API | Free |

Estimated monthly cost for 20 emails/month: **< $1**
