/**
 * Mathnasium Nightly Radius Scraper
 *
 * What this does each night:
 * 1. Logs into Radius
 * 2. Downloads the Student Report (all enrolled students, both centers)
 * 3. Parses it to identify who needs an email draft today:
 *    - Last Assessment = yesterday → Level Up draft
 *    - Last Progress Check = yesterday AND Last Assessment is NOT within
 *      the last 7 days → Progress Check draft (admin reviews before sending)
 *    - Both triggered → Level Up wins, no duplicate email
 * 4. Navigates to each triggered student's page and downloads their Learning Plan PDF
 * 5. Pushes Student Report + Learning Plan PDFs + manifest.json to Google Drive
 *
 * All drafts require manual admin approval in the dashboard before sending.
 */

const { chromium } = require('playwright');
const { google }   = require('googleapis');
const XLSX         = require('xlsx');
const fs           = require('fs');
const path         = require('path');
const { format, subDays, differenceInDays } = require('date-fns');

// ─────────────────────────────────────────────
// CONFIG
// ─────────────────────────────────────────────
const RADIUS_BASE_URL = 'https://radius.mathnasium.com';
const DOWNLOAD_DIR    = path.join(__dirname, 'downloads');
const DEBUG_DIR       = path.join(__dirname, 'debug');
const DRY_RUN         = process.argv.includes('--dry-run');

const yesterday     = subDays(new Date(), 1);
const yesterdayStr  = format(yesterday, 'M/d/yyyy');     // e.g. 3/22/2026
const folderDateStr = format(new Date(), 'yyyy-MM-dd');  // e.g. 2026-03-23

// If a student's Last Assessment is within this many days of their
// Last Progress Check, we treat it as a Level Up (not a standalone Progress Check)
const LEVEL_UP_WINDOW_DAYS = 7;

// ─────────────────────────────────────────────
// MAIN
// ─────────────────────────────────────────────
async function main() {
  console.log(`\n🏫 Mathnasium Email Queue Builder`);
  console.log(`📅 Running for: ${yesterdayStr}`);
  if (DRY_RUN) console.log(`🧪 DRY RUN MODE — no files will be uploaded\n`);

  [DOWNLOAD_DIR, DEBUG_DIR].forEach(d => {
    if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
  });

  const browser = await chromium.launch({
    headless: true,
    args: ['--no-sandbox', '--disable-setuid-sandbox']
  });

  const context = await browser.newContext({
    acceptDownloads: true,
    viewport: { width: 1280, height: 900 }
  });

  const page = await context.newPage();

  try {
    // ── 1. Login ──────────────────────────────────────────────────────────
    await loginToRadius(page);

    // ── 2. Download Student Report ────────────────────────────────────────
    const studentReportPath = await downloadStudentReport(page);

    // ── 3. Parse report — identify triggered students ─────────────────────
    const triggeredStudents = parseStudentReport(studentReportPath);

    if (triggeredStudents.length === 0) {
      console.log('\n✅ No students triggered today — nothing to do');
      await browser.close();
      return;
    }

    console.log(`\n📋 ${triggeredStudents.length} student(s) need drafts today:`);
    triggeredStudents.forEach(s =>
      console.log(`  • ${s.studentName} [${s.emailType}] — ${s.emails.join(', ')}`)
    );

    // ── 4. Download Learning Plans ────────────────────────────────────────
    await downloadLearningPlans(page, triggeredStudents);

    // ── 5. Upload to Google Drive ─────────────────────────────────────────
    if (!DRY_RUN) {
      await uploadToGoogleDrive(studentReportPath, triggeredStudents);
    } else {
      console.log('\n🧪 DRY RUN: skipping Drive upload');
      console.log('Local files ready in scraper/downloads/');
    }

    console.log('\n✅ Done');

  } catch (err) {
    console.error('\n❌ Scraper failed:', err.message);
    await page.screenshot({ path: path.join(DEBUG_DIR, `error-${Date.now()}.png`) });
    throw err;
  } finally {
    await browser.close();
  }
}

// ─────────────────────────────────────────────
// 1. LOGIN
// ─────────────────────────────────────────────
async function loginToRadius(page) {
  console.log('🔐 Logging in...');

  await page.goto(`${RADIUS_BASE_URL}/Account/Login`, { waitUntil: 'networkidle' });
  await page.fill('#UserName', process.env.RADIUS_USERNAME);
  await page.fill('#Password', process.env.RADIUS_PASSWORD);
  await page.click('#login');

  await page.waitForSelector('.icon-bar', { timeout: 15000 }).catch(async () => {
    await page.screenshot({ path: path.join(DEBUG_DIR, 'login-failed.png') });
    throw new Error('Login failed — see debug/login-failed.png');
  });

  console.log('  ✓ Logged in');
}

// ─────────────────────────────────────────────
// 2. DOWNLOAD STUDENT REPORT
// ─────────────────────────────────────────────
async function downloadStudentReport(page) {
  console.log('\n📋 Downloading Student Report...');

  await page.goto(`${RADIUS_BASE_URL}/StudentReport`, { waitUntil: 'networkidle' });

  // Set Enrollment Filter to "Enrolled" using Kendo UI dropdown
  // Step 1: click the visible dropdown widget to open it
  await page.click('[aria-owns="enrollmentFiltersDropDownList_listbox"]');

  // Step 2: wait for the dropdown list to appear and click "Enrolled"
  await page.waitForSelector('[id="enrollmentFiltersDropDownList_listbox"]', { timeout: 5000 });
  await page.click('[id="enrollmentFiltersDropDownList_listbox"] li:has-text("Enrolled")');

  // Leave dates blank — date filters don't apply to the columns we care about
  // Leave Centers as "All" — we want both Teaneck and Englewood

  // Click Search
  await page.click('#btnsearch');
  await page.waitForLoadState('networkidle');

  // Wait for data to load — the table shows "0" items while loading
  await page.waitForFunction(() => {
    const el = document.querySelector('.stdntRpt');
    return el && !el.textContent.includes('0 items');
  }, { timeout: 30000 }).catch(() => {
    console.log('  ⚠️  Table may still be loading — proceeding anyway');
  });

  // Click Export to Excel
  const downloadPromise = page.waitForEvent('download', { timeout: 30000 });
  await page.click('#btnExport');
  const download = await downloadPromise;

  const filePath = path.join(DOWNLOAD_DIR, `student-report-${folderDateStr}.xlsx`);
  await download.saveAs(filePath);
  console.log(`  ✓ Saved: ${path.basename(filePath)}`);
  return filePath;
}

// ─────────────────────────────────────────────
// 3. PARSE STUDENT REPORT
// ─────────────────────────────────────────────
function parseStudentReport(filePath) {
  console.log('\n🔍 Parsing Student Report...');

  const wb    = XLSX.readFile(filePath);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows  = XLSX.utils.sheet_to_json(sheet);

  const triggeredStudents = [];
  const levelUpNames = new Set();

  // First pass — find Level Ups (Last Assessment = yesterday)
  for (const row of rows) {
    if (row['Enrollment Status'] !== 'Enrolled') continue;

    const lastAssessment = normalizeDate(row['Last Assessment']);
    if (lastAssessment !== yesterdayStr) continue;

    const name = (row['Student Name'] || '').trim();
    if (!name || name === 'x x') continue; // skip test records

    const emails = parseEmails(row['Guardian Emails'] || row['Guardian Email List']);
    if (emails.length === 0) {
      console.warn(`  ⚠️  ${name} has no guardian email — skipping`);
      continue;
    }

    triggeredStudents.push({
      studentName:  name,
      firstName:    name.split(' ')[0],
      emailType:    'level-up',
      emails,
      grade:        String(row['Grade'] || ''),
      center:       row['Center'] || '',
      leadId:       row['Lead Id'] || row['LeadId'] || '',
      accountId:    row['Account Id'] || row['AccountId'] || '',
      lastAssessment,
      lastProgressCheck: normalizeDate(row['Last Progress Check']),
      learningPlanStatus: 'pending',
      learningPlanPath:   null,
    });

    levelUpNames.add(name.toLowerCase());
    console.log(`  ✓ Level Up: ${name}`);
  }

  // Second pass — find standalone Progress Checks
  // (Last Progress Check = yesterday AND not already a Level Up)
  for (const row of rows) {
    if (row['Enrollment Status'] !== 'Enrolled') continue;

    const lastPC = normalizeDate(row['Last Progress Check']);
    if (lastPC !== yesterdayStr) continue;

    const name = (row['Student Name'] || '').trim();
    if (!name || name === 'x x') continue;

    // Skip if already captured as a Level Up
    if (levelUpNames.has(name.toLowerCase())) continue;

    // Skip if Last Assessment is within the level-up window
    // (means a PC was graded late but the Pre Assessment already happened)
    const lastAssessment = normalizeDate(row['Last Assessment']);
    if (lastAssessment) {
      const assessDate = parseDate(lastAssessment);
      const pcDate     = parseDate(lastPC);
      if (assessDate && pcDate) {
        const daysDiff = Math.abs(differenceInDays(assessDate, pcDate));
        if (daysDiff <= LEVEL_UP_WINDOW_DAYS) {
          // Like Leo's situation — PC graded after Assessment, treat as Level Up
          console.log(`  ✓ Level Up (late-graded PC): ${name} — Assessment was ${daysDiff} day(s) from PC`);
          const emails = parseEmails(row['Guardian Emails'] || row['Guardian Email List']);
          if (emails.length === 0) {
            console.warn(`  ⚠️  ${name} has no guardian email — skipping`);
            continue;
          }
          triggeredStudents.push({
            studentName:  name,
            firstName:    name.split(' ')[0],
            emailType:    'level-up',
            emails,
            grade:        String(row['Grade'] || ''),
            center:       row['Center'] || '',
            leadId:       row['Lead Id'] || row['LeadId'] || '',
            accountId:    row['Account Id'] || row['AccountId'] || '',
            lastAssessment,
            lastProgressCheck: lastPC,
            learningPlanStatus: 'pending',
            learningPlanPath:   null,
          });
          continue;
        }
      }
    }

    const emails = parseEmails(row['Guardian Emails'] || row['Guardian Email List']);
    if (emails.length === 0) {
      console.warn(`  ⚠️  ${name} has no guardian email — skipping`);
      continue;
    }

    triggeredStudents.push({
      studentName:  name,
      firstName:    name.split(' ')[0],
      emailType:    'progress-check',
      emails,
      grade:        String(row['Grade'] || ''),
      center:       row['Center'] || '',
      leadId:       row['Lead Id'] || row['LeadId'] || '',
      accountId:    row['Account Id'] || row['AccountId'] || '',
      lastAssessment,
      lastProgressCheck: lastPC,
      learningPlanStatus: 'pending',
      learningPlanPath:   null,
    });

    console.log(`  ✓ Progress Check: ${name}`);
  }

  console.log(`\n  Total: ${triggeredStudents.length} (${triggeredStudents.filter(s=>s.emailType==='level-up').length} level-up, ${triggeredStudents.filter(s=>s.emailType==='progress-check').length} progress-check)`);
  return triggeredStudents;
}

// ─────────────────────────────────────────────
// 4. DOWNLOAD LEARNING PLANS
// ─────────────────────────────────────────────
async function downloadLearningPlans(page, triggeredStudents) {
  console.log('\n📚 Downloading Learning Plans...');

  for (const student of triggeredStudents) {
    try {
      const filePath = await downloadSingleLearningPlan(page, student);
      student.learningPlanStatus = 'ready';
      student.learningPlanPath   = filePath;
      console.log(`  ✓ ${student.studentName}`);
    } catch (err) {
      console.warn(`  ⚠️  ${student.studentName} — LP download failed: ${err.message}`);
      student.learningPlanStatus = 'error';
      await page.screenshot({
        path: path.join(DEBUG_DIR, `lp-error-${student.studentName.replace(/\s+/g, '-')}.png`)
      });
    }
  }
}

async function downloadSingleLearningPlan(page, student) {
  const studentPageUrl = `${RADIUS_BASE_URL}/Student/Details/${student.leadId}`;
  await page.goto(studentPageUrl, { waitUntil: 'networkidle' });

  // Selector confirmed from Radius student page inspection
  const lpButton = page.locator('a.k-grid-LPReport').first();

  const count = await lpButton.count();
  if (!count) {
    throw new Error('Learning Plan download button not found — need to update selector');
  }

  const downloadPromise = page.waitForEvent('download', { timeout: 15000 });
  await lpButton.click();
  const download = await downloadPromise;

  const safeName = student.studentName.replace(/[^a-z0-9]/gi, '-').toLowerCase();
  const filePath = path.join(DOWNLOAD_DIR, `lp-${safeName}-${folderDateStr}.pdf`);
  await download.saveAs(filePath);
  return filePath;
}

// ─────────────────────────────────────────────
// 5. UPLOAD TO GOOGLE DRIVE
// ─────────────────────────────────────────────
async function uploadToGoogleDrive(studentReportPath, triggeredStudents) {
  console.log('\n☁️  Uploading to Google Drive...');

  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY),
    scopes: ['https://www.googleapis.com/auth/drive.file'],
  });
  const drive = google.drive({ version: 'v3', auth });

  // Create dated subfolder
  const folder = await drive.files.create({
    requestBody: {
      name: folderDateStr,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [process.env.GOOGLE_DRIVE_FOLDER_ID],
    },
    fields: 'id',
  });
  const folderId = folder.data.id;
  console.log(`  ✓ Created folder: ${folderDateStr}`);

  // Upload Student Report
  await uploadFile(drive, studentReportPath, folderId,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

  // Upload Learning Plan PDFs
  for (const student of triggeredStudents) {
    if (student.learningPlanPath) {
      await uploadFile(drive, student.learningPlanPath, folderId, 'application/pdf');
    }
  }

  // Write and upload manifest so Apps Script knows what to process
  const manifest = {
    date: folderDateStr,
    generatedAt: new Date().toISOString(),
    studentReportFile: path.basename(studentReportPath),
    triggeredStudents: triggeredStudents.map(s => ({
      studentName:        s.studentName,
      firstName:          s.firstName,
      emailType:          s.emailType,
      emails:             s.emails,
      grade:              s.grade,
      center:             s.center,
      leadId:             s.leadId,
      lastAssessment:     s.lastAssessment,
      lastProgressCheck:  s.lastProgressCheck,
      learningPlanStatus: s.learningPlanStatus,
      learningPlanFile:   s.learningPlanPath ? path.basename(s.learningPlanPath) : null,
    }))
  };

  const manifestPath = path.join(DOWNLOAD_DIR, `manifest-${folderDateStr}.json`);
  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2));
  await uploadFile(drive, manifestPath, folderId, 'application/json');

  console.log(`  ✓ Uploaded manifest with ${triggeredStudents.length} students`);
}

async function uploadFile(drive, filePath, folderId, mimeType) {
  await drive.files.create({
    requestBody: { name: path.basename(filePath), parents: [folderId] },
    media: { mimeType, body: fs.createReadStream(filePath) },
  });
  console.log(`  ✓ Uploaded: ${path.basename(filePath)}`);
}

// ─────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────

function parseEmails(emailStr) {
  if (!emailStr) return [];
  return String(emailStr)
    .split(',')
    .map(e => e.trim().toLowerCase())
    .filter(e => e.includes('@') && !e.includes('mathnasium.com'));
}

function normalizeDate(val) {
  if (!val) return null;
  if (typeof val === 'number') {
    // Excel serial date number
    const d = XLSX.SSF.parse_date_code(val);
    return `${d.m}/${d.d}/${d.y}`;
  }
  const s = String(val).trim().split(' ')[0]; // strip any time component
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return `${parseInt(m[1])}/${parseInt(m[2])}/${m[3]}`;
  return s;
}

function parseDate(dateStr) {
  if (!dateStr) return null;
  const [m, d, y] = dateStr.split('/').map(Number);
  if (!m || !d || !y) return null;
  return new Date(y, m - 1, d);
}

// ─────────────────────────────────────────────
// RUN
// ─────────────────────────────────────────────
main().catch(err => {
  console.error(err);
  process.exit(1);
});
