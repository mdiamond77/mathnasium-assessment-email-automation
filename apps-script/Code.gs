/**
 * Mathnasium Email Queue Builder — Google Apps Script
 *
 * SETUP:
 * 1. Create a new Google Apps Script project at script.google.com
 * 2. Paste this entire file as Code.gs
 * 3. Create a second file called dashboard.html and paste the dashboard code
 * 4. Set Script Properties (Project Settings → Script Properties):
 *    - ANTHROPIC_API_KEY    → your Claude API key
 *    - DRIVE_FOLDER_ID      → the Google Drive folder ID the scraper uploads to
 *    - QUEUE_SHEET_ID       → the Google Sheet ID to use as the email queue
 *    - NOTIFY_EMAIL         → your email address for daily summaries
 *    - GMAIL_SENDER_NAME    → "Mathnasium of Teaneck" or "Mathnasium of Englewood"
 * 5. Run setupTrigger() once to register the nightly Drive-watch trigger
 * 6. Deploy as Web App (Deploy → New Deployment → Web App → Anyone with link)
 */

// ─────────────────────────────────────────────
// ENTRY POINTS
// ─────────────────────────────────────────────

/**
 * Run once to set up the time-based trigger.
 * Checks for new Drive files every hour between 6-8am.
 */
function setupTrigger() {
  // Delete existing triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Check hourly — the scraper runs at 6am, this catches it promptly
  ScriptApp.newTrigger('processNewDriveFiles')
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log('✓ Trigger set up: processNewDriveFiles runs every hour');
}

/**
 * Main pipeline — called by the hourly trigger.
 * Finds new dated folders in Drive, processes any unprocessed ones.
 */
function processNewDriveFiles() {
  const props = PropertiesService.getScriptProperties();
  const driveFolderId = props.getProperty('DRIVE_FOLDER_ID');

  const rootFolder = DriveApp.getFolderById(driveFolderId);
  const subFolders = rootFolder.getFolders();

  while (subFolders.hasNext()) {
    const folder = subFolders.next();
    const folderName = folder.getName(); // e.g. "2026-03-23"

    // Skip if already processed
    if (wasAlreadyProcessed(folderName)) continue;

    // Skip if folder is older than 3 days (avoid reprocessing old data)
    if (isTooOld(folderName)) continue;

    // Look for the manifest file
    const files = folder.getFilesByName(`manifest-${folderName}.json`);
    if (!files.hasNext()) continue;

    const manifestFile = files.next();
    const manifest = JSON.parse(manifestFile.getBlob().getDataAsString());

    Logger.log(`Processing folder: ${folderName} — ${manifest.triggeredStudents.length} students`);

    processManifest(manifest, folder, folderName);
    markAsProcessed(folderName);
  }
}

// ─────────────────────────────────────────────
// CORE PROCESSING
// ─────────────────────────────────────────────

function processManifest(manifest, driveFolder, dateStr) {
  const queue = [];

  for (const student of manifest.triggeredStudents) {
    Logger.log(`  Processing: ${student.studentName} (${student.emailType})`);

    // Get learning plan PDF blob if available
    let learningPlanBlob = null;
    if (student.learningPlanFile) {
      const lpFiles = driveFolder.getFilesByName(student.learningPlanFile);
      if (lpFiles.hasNext()) {
        learningPlanBlob = lpFiles.next().getBlob();
        Logger.log(`    ✓ Learning plan found: ${student.learningPlanFile}`);
      }
    }

    // Get assessment PDF/XLSX blob if available
    let assessmentBlob = null;
    if (student.preAssessmentFile) {
      const aFiles = driveFolder.getFilesByName(student.preAssessmentFile);
      if (aFiles.hasNext()) {
        assessmentBlob = aFiles.next().getBlob();
      }
    }

    // Generate emails via Claude
    let emailDrafts = null;
    try {
      emailDrafts = generateEmailsViaClaude(student, learningPlanBlob, assessmentBlob);
      Logger.log(`    ✓ Emails generated`);
    } catch (err) {
      Logger.log(`    ✗ Claude API failed: ${err.message}`);
      emailDrafts = {
        error: err.message,
        warm: { subject: '', body: '' },
        pro: { subject: '', body: '' }
      };
    }

    queue.push({
      date: dateStr,
      studentName: student.studentName,
      firstName: student.firstName,
      emailType: student.emailType,
      assessmentLevel: student.assessmentLevel || '—',
      scorePercent: student.scorePercent !== null ? `${student.scorePercent}%` : '—',
      grade: student.grade,
      center: student.center,
      emails: student.emails.join(', '),
      learningPlanStatus: student.learningPlanStatus,
      warmSubject: emailDrafts.warm.subject,
      warmBody: emailDrafts.warm.body,
      proSubject: emailDrafts.pro.subject,
      proBody: emailDrafts.pro.body,
      status: 'pending',
      error: emailDrafts.error || '',
    });
  }

  // Write to Google Sheet queue
  writeToQueue(queue);

  // Send summary notification email
  sendSummaryEmail(queue, dateStr);
}

// ─────────────────────────────────────────────
// CLAUDE API
// ─────────────────────────────────────────────

function generateEmailsViaClaude(student, learningPlanBlob, assessmentBlob) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('ANTHROPIC_API_KEY');

  const location = student.center === 'Teaneck' ? 'Teaneck' : 'Englewood';
  const isLevelUp = student.emailType === 'level-up';
  const schoolGrade = student.grade ? `${student.grade}th grade` : null;

  const systemPrompt = `You are writing parent update emails for Mathnasium of ${location}, a math learning center.

Each document is clearly labeled with === START === and === END === markers. Read each document only for what its label says.

════════════════════════════
LEVEL IDENTIFICATION — CRITICAL
════════════════════════════
• The Mathnasium curriculum has distinct levels: Level 1, Level 2, Level 3, Level 4, Level 5, Level 6, Level 7, Algebra Readiness, Algebra 1, and beyond.
• Level 7 and Algebra Readiness are COMPLETELY DIFFERENT levels.
• Read the level name ONLY from the large bold title at the top of each document.
• NEVER borrow a level name from one document and apply it to another.

════════════════════════════
READING ASSESSMENTS
════════════════════════════
• BLACK printed answers = CORRECT.
• RED answers or red marks = INCORRECT.
• Identify 1–2 topic areas of strength from correct answers.
• Frame developing areas only as "what we're building toward" — never as failures.

════════════════════════════
NAMING CONVENTION — ALWAYS FOLLOW
════════════════════════════
Always use the specific level name when referring to documents:
• "his Level 7 Assessment" — NOT "the assessment"
• "his Level 7 Learning Plan" — NOT "the learning plan"
• "her Algebra Readiness Learning Plan" — NOT "her new learning plan"

════════════════════════════
EMAIL RULES
════════════════════════════
✗ NEVER mention session counts or attendance numbers.
✗ NEVER invent scores, percentages, or statistics not visible on the documents.
✗ NEVER suggest at-home practice or ask parents to work on problems with their child.
✗ NEVER use: "struggled," "missed," "got wrong," "needs work," "incorrect," "challenging."
✗ NEVER use generic document references — always use the specific level name.
✓ Keep each email to 3 short paragraphs.
✓ Sign off: "The Team at Mathnasium of ${location}"
${schoolGrade ? `✓ Student is in ${schoolGrade}. If their Mathnasium level aligns with or exceeds typical ${schoolGrade} content, note it with genuine pride. If building toward grade level, frame it as strong meaningful progress.` : ''}

════════════════════════════
TWO VERSIONS — SAME CONTENT, DIFFERENT ENERGY
════════════════════════════

VERSION 1 — CELEBRATORY
Tone: Enthusiastic, warm, coach energy. 3 short paragraphs.
P1: Celebrate the ${isLevelUp ? 'Level Up milestone' : 'Progress Check milestone'} — name the specific level.
P2: One strength from the assessment + what their Learning Plan covers next (if available), named specifically.
P3: Warm, forward-looking close.

VERSION 2 — WARM & STRAIGHTFORWARD
Tone: Friendly, genuine, matter-of-fact. Same 3 paragraphs, calmer delivery, no exclamation points.

Both versions should be similar in length. Difference is energy, not quantity.

════════════════════════════
RESPOND IN EXACT JSON ONLY — no markdown:
════════════════════════════
{
  "student_name": "first name only",
  "assessed_level": "exact level from assessment header",
  "plan_level": "exact level from learning plan header, or same as assessed_level if no plan",
  "topics_strong": ["topic 1", "topic 2"],
  "topics_developing": ["topic"],
  "warm": { "subject": "...", "body": "..." },
  "pro":  { "subject": "...", "body": "..." }
}`;

  // Build message content — labeled documents
  const messageContent = [];

  if (assessmentBlob) {
    const assessmentB64 = Utilities.base64Encode(assessmentBlob.getBytes());
    const assessmentMime = assessmentBlob.getContentType() || 'application/pdf';
    messageContent.push({
      type: 'text',
      text: `\n=== START: ${student.emailType.toUpperCase().replace('-', ' ')} ASSESSMENT ===\nThis is the student's assessment. Read the level name ONLY from the large bold title at the top.`
    });
    if (assessmentMime === 'application/pdf') {
      messageContent.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: assessmentB64 }});
    } else {
      messageContent.push({ type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: assessmentB64 }});
    }
    messageContent.push({ type: 'text', text: `=== END: ASSESSMENT ===\n` });
  }

  if (learningPlanBlob) {
    const lpB64 = Utilities.base64Encode(learningPlanBlob.getBytes());
    messageContent.push({
      type: 'text',
      text: `\n=== START: LEARNING PLAN ===\nThis is the Learning Plan. Read the level name ONLY from its header. This is what the student works on going forward.`
    });
    messageContent.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: lpB64 }});
    messageContent.push({ type: 'text', text: `=== END: LEARNING PLAN ===\n` });
  }

  // Always include student context as text even if no documents
  messageContent.push({
    type: 'text',
    text: `Student: ${student.firstName} (${student.emailType}, ${student.assessmentLevel || 'level unknown'}, score: ${student.scorePercent !== null ? student.scorePercent + '%' : 'not available'}).
${!assessmentBlob && !learningPlanBlob ? 'No documents are available for this student — write a general warm update based on the student context provided.' : ''}
Generate both email versions following all instructions.`
  });

  const payload = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 1500,
    system: systemPrompt,
    messages: [{ role: 'user', content: messageContent }]
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error(`Claude API returned ${responseCode}: ${responseText}`);
  }

  const data = JSON.parse(responseText);
  const rawText = data.content.map(b => b.text || '').join('').trim();

  let parsed;
  try {
    parsed = JSON.parse(rawText.replace(/```json|```/g, '').trim());
  } catch {
    const match = rawText.match(/\{[\s\S]*\}/);
    if (match) parsed = JSON.parse(match[0]);
    else throw new Error('Could not parse Claude response as JSON');
  }

  return parsed;
}

// ─────────────────────────────────────────────
// GOOGLE SHEET QUEUE
// ─────────────────────────────────────────────

const QUEUE_HEADERS = [
  'ID', 'Date', 'Status', 'Student Name', 'First Name', 'Email Type',
  'Assessment Level', 'Score', 'Grade', 'Center', 'Guardian Emails',
  'LP Status', 'Warm Subject', 'Warm Body', 'Pro Subject', 'Pro Body',
  'Sent At', 'Sent Version', 'Error', 'Skip Reason'
];

function getQueueSheet() {
  const props = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('QUEUE_SHEET_ID');
  const ss = SpreadsheetApp.openById(sheetId);
  let sheet = ss.getSheetByName('Queue');

  if (!sheet) {
    sheet = ss.insertSheet('Queue');
    sheet.appendRow(QUEUE_HEADERS);
    sheet.setFrozenRows(1);
    // Format header row
    sheet.getRange(1, 1, 1, QUEUE_HEADERS.length)
      .setBackground('#E8352A')
      .setFontColor('white')
      .setFontWeight('bold');
  }

  return sheet;
}

function writeToQueue(queue) {
  const sheet = getQueueSheet();

  for (const item of queue) {
    const id = `${item.date}-${item.studentName.replace(/\s/g, '-').toLowerCase()}`;
    const row = [
      id,
      item.date,
      item.error ? 'error' : 'pending',
      item.studentName,
      item.firstName,
      item.emailType,
      item.assessmentLevel,
      item.scorePercent,
      item.grade,
      item.center,
      item.emails,
      item.learningPlanStatus,
      item.warmSubject,
      item.warmBody,
      item.proSubject,
      item.proBody,
      '', // Sent At
      '', // Sent Version
      item.error,
      '', // Skip Reason
    ];
    sheet.appendRow(row);
  }

  Logger.log(`✓ Wrote ${queue.length} rows to queue sheet`);
}

function markQueueItemSent(rowIndex, version) {
  const sheet = getQueueSheet();
  sheet.getRange(rowIndex, 3).setValue('sent');              // Status
  sheet.getRange(rowIndex, 17).setValue(new Date().toISOString()); // Sent At
  sheet.getRange(rowIndex, 18).setValue(version);             // Sent Version
}

function markQueueItemSkipped(rowIndex, reason) {
  const sheet = getQueueSheet();
  sheet.getRange(rowIndex, 3).setValue('skipped');           // Status
  sheet.getRange(rowIndex, 20).setValue(reason);              // Skip Reason
}

// ─────────────────────────────────────────────
// SUMMARY EMAIL
// ─────────────────────────────────────────────

function sendSummaryEmail(queue, dateStr) {
  const props = PropertiesService.getScriptProperties();
  const notifyEmail = props.getProperty('NOTIFY_EMAIL');
  const dashboardUrl = ScriptApp.getService().getUrl();

  const pending = queue.filter(q => q.status === 'pending' || q.status === '');
  const errors = queue.filter(q => q.error);

  if (pending.length === 0) {
    Logger.log('No pending emails — skipping summary notification');
    return;
  }

  const teaneckCount = pending.filter(q => q.center === 'Teaneck').length;
  const englewoodCount = pending.filter(q => q.center === 'Englewood').length;

  const studentList = pending.map(q => {
    const lpFlag = q.learningPlanStatus !== 'ready' ? ' ⚠️ LP pending' : '';
    return `• ${q.studentName} — ${q.emailType} — ${q.assessmentLevel} — ${q.center}${lpFlag}`;
  }).join('\n');

  const subject = `📧 ${pending.length} parent email${pending.length > 1 ? 's' : ''} ready to review — ${dateStr}`;

  const body = `
Hi,

${pending.length} parent email${pending.length > 1 ? 's are' : ' is'} ready for your review from ${dateStr}.

Teaneck: ${teaneckCount}
Englewood: ${englewoodCount}

Students:
${studentList}

${errors.length > 0 ? `⚠️ ${errors.length} email(s) had errors generating — review in the dashboard.\n` : ''}
👉 Review and send here:
${dashboardUrl}

— Mathnasium Email Automation
  `.trim();

  GmailApp.sendEmail(notifyEmail, subject, body);
  Logger.log(`✓ Summary email sent to ${notifyEmail}`);
}

// ─────────────────────────────────────────────
// WEB APP — SEND EMAIL ACTION
// ─────────────────────────────────────────────

/**
 * Serves the dashboard HTML
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('dashboard')
    .setTitle('Mathnasium Email Queue')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Called by the dashboard to get all pending queue items
 */
function getQueueItems() {
  const sheet = getQueueSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  return rows
    .map((row, index) => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      obj._rowIndex = index + 2; // 1-indexed, +1 for header
      return obj;
    })
    .filter(row => row['Status'] === 'pending' || row['Status'] === 'error');
}

/**
 * Called by the dashboard to send an approved email
 */
function sendApprovedEmail(rowIndex, version, editedSubject, editedBody) {
  const sheet = getQueueSheet();
  const row = sheet.getRange(rowIndex, 1, 1, QUEUE_HEADERS.length).getValues()[0];

  const emailsRaw = row[QUEUE_HEADERS.indexOf('Guardian Emails')];
  const studentName = row[QUEUE_HEADERS.indexOf('Student Name')];
  const center = row[QUEUE_HEADERS.indexOf('Center')];

  const emails = emailsRaw.split(',').map(e => e.trim()).filter(Boolean);
  if (emails.length === 0) {
    throw new Error(`No email addresses found for ${studentName}`);
  }

  const senderName = `Mathnasium of ${center}`;

  // Send to all guardian emails
  emails.forEach(email => {
    GmailApp.sendEmail(email, editedSubject, editedBody, {
      name: senderName,
      replyTo: Session.getActiveUser().getEmail(),
    });
  });

  markQueueItemSent(rowIndex, version);
  Logger.log(`✓ Email sent to ${emails.join(', ')} for ${studentName}`);

  return { success: true, sentTo: emails };
}

/**
 * Called by the dashboard to skip an email
 */
function skipEmail(rowIndex, reason) {
  markQueueItemSkipped(rowIndex, reason || 'Skipped by admin');
  return { success: true };
}

/**
 * Called by the dashboard to update email body inline
 */
function saveEdits(rowIndex, warmSubject, warmBody, proSubject, proBody) {
  const sheet = getQueueSheet();
  sheet.getRange(rowIndex, QUEUE_HEADERS.indexOf('Warm Subject') + 1).setValue(warmSubject);
  sheet.getRange(rowIndex, QUEUE_HEADERS.indexOf('Warm Body') + 1).setValue(warmBody);
  sheet.getRange(rowIndex, QUEUE_HEADERS.indexOf('Pro Subject') + 1).setValue(proSubject);
  sheet.getRange(rowIndex, QUEUE_HEADERS.indexOf('Pro Body') + 1).setValue(proBody);
  return { success: true };
}

// ─────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────

function wasAlreadyProcessed(folderName) {
  const props = PropertiesService.getScriptProperties();
  const processed = JSON.parse(props.getProperty('processed_folders') || '[]');
  return processed.includes(folderName);
}

function markAsProcessed(folderName) {
  const props = PropertiesService.getScriptProperties();
  const processed = JSON.parse(props.getProperty('processed_folders') || '[]');
  processed.push(folderName);
  // Keep only last 30 days
  const recent = processed.slice(-30);
  props.setProperty('processed_folders', JSON.stringify(recent));
}

function isTooOld(folderName) {
  try {
    const folderDate = new Date(folderName);
    const threeDaysAgo = new Date();
    threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
    return folderDate < threeDaysAgo;
  } catch {
    return false;
  }
}
