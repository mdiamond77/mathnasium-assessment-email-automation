"""
Mathnasium Nightly Radius Scraper (Python)

1. Logs into Radius
2. Downloads the Student Report (all enrolled students)
3. Parses it to identify who needs an email draft:
   - Last Assessment = yesterday  → Level Up draft
   - Last Progress Check = yesterday AND no Assessment within 7 days → Progress Check draft
4. Downloads Learning Plan PDF for each triggered student
5. Pushes everything to Google Drive
"""

import os
import json
import re
from datetime import datetime, timedelta
from pathlib import Path

from playwright.sync_api import sync_playwright
import openpyxl
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
RADIUS_BASE_URL     = 'https://radius.mathnasium.com'
DOWNLOAD_DIR        = Path(__file__).parent / 'downloads'
DEBUG_DIR           = Path(__file__).parent / 'debug'
DRY_RUN             = '--dry-run' in os.sys.argv
LEVEL_UP_WINDOW_DAYS = 7

yesterday       = datetime.now() - timedelta(days=1)
yesterday_str   = f"{yesterday.month}/{yesterday.day}/{yesterday.year}"
folder_date_str = datetime.now().strftime('%Y-%m-%d')

DOWNLOAD_DIR.mkdir(exist_ok=True)
DEBUG_DIR.mkdir(exist_ok=True)

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    print(f"\n🏫 Mathnasium Email Queue Builder")
    print(f"📅 Running for: {yesterday_str}")
    if DRY_RUN:
        print("🧪 DRY RUN MODE — no files will be uploaded\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=True,
            args=['--no-sandbox', '--disable-setuid-sandbox']
        )
        context = browser.new_context(
            accept_downloads=True,
            viewport={'width': 1280, 'height': 900}
        )
        page = context.new_page()

        try:
            # 1. Login
            login(page)

            # 2. Download Student Report
            student_report_path = download_student_report(page)

            # 3. Parse report
            triggered_students = parse_student_report(student_report_path)

            if not triggered_students:
                print('\n✅ No students triggered today — nothing to do')
                return

            print(f"\n📋 {len(triggered_students)} student(s) need drafts today:")
            for s in triggered_students:
                print(f"  • {s['studentName']} [{s['emailType']}] — {', '.join(s['emails'])}")

            # 4. Download Learning Plans
            download_learning_plans(page, triggered_students)

            # 5. Upload to Google Drive
            if not DRY_RUN:
                upload_to_google_drive(student_report_path, triggered_students)
            else:
                print('\n🧪 DRY RUN: skipping Drive upload')
                print('Local files ready in scraper/downloads/')

            print('\n✅ Done')

        except Exception as e:
            print(f'\n❌ Scraper failed: {e}')
            page.screenshot(path=str(DEBUG_DIR / f'error-{int(datetime.now().timestamp())}.png'))
            raise
        finally:
            browser.close()

# ─────────────────────────────────────────────
# 1. LOGIN
# ─────────────────────────────────────────────
def login(page):
    print('🔐 Logging in...')
    page.goto(RADIUS_BASE_URL)
    page.wait_for_load_state('networkidle')
    page.fill('#UserName', os.environ['RADIUS_USERNAME'])
    page.fill('#Password', os.environ['RADIUS_PASSWORD'])
    page.click('#login')
    page.wait_for_load_state('networkidle')

    try:
        # "Sign Out" link only appears when successfully logged in
        page.wait_for_selector('a:has-text("Sign Out")', timeout=15000)
        print('  ✓ Logged in')
    except Exception:
        page.screenshot(path=str(DEBUG_DIR / 'login-failed.png'))
        raise Exception('Login failed — see debug/login-failed.png')

# ─────────────────────────────────────────────
# 2. DOWNLOAD STUDENT REPORT
# ─────────────────────────────────────────────
def download_student_report(page):
    print(chr(10) + chr(128203) + chr(32) + 'Downloading Student Report...')

    import time

    page.goto(f"{RADIUS_BASE_URL}/StudentReport")
    page.wait_for_load_state('networkidle')
    time.sleep(2)

    # Open the Enrollment Filter Kendo dropdown and select Enrolled
    page.wait_for_selector('[aria-owns="enrollmentFiltersDropDownList_listbox"]', timeout=10000)
    page.click('[aria-owns="enrollmentFiltersDropDownList_listbox"]')
    page.wait_for_selector('#enrollmentFiltersDropDownList_listbox', timeout=10000)
    time.sleep(1)
    page.click("#enrollmentFiltersDropDownList_listbox li:text-is(\"Enrolled\")")
    time.sleep(1)

    # Click Search and wait generously for all students to load
    page.click('#btnsearch')
    page.wait_for_load_state('networkidle')
    time.sleep(10)

    # Save screenshot to downloads so it appears in artifacts
    page.screenshot(path=str(DOWNLOAD_DIR / 'before-export.png'))
    print('  Screenshot saved to downloads/before-export.png')

    # Export to Excel
    with page.expect_download(timeout=30000) as download_info:
        page.click('#btnExport')
    download = download_info.value

    file_path = DOWNLOAD_DIR / f'student-report-{folder_date_str}.xlsx'
    download.save_as(str(file_path))
    print(f'  Saved: {file_path.name}')
    return file_path

# 3. PARSE STUDENT REPORT
# ─────────────────────────────────────────────
def parse_student_report(file_path):
    print('\n🔍 Parsing Student Report...')

    wb = openpyxl.load_workbook(str(file_path))
    ws = wb.active
    headers = [cell.value for cell in ws[1]]

    def get(row, col_name):
        if col_name in headers:
            val = row[headers.index(col_name)]
            return val.value if hasattr(val, 'value') else val
        return None

    triggered_students = []
    level_up_names = set()

    rows = list(ws.iter_rows(min_row=2))

    # First pass — Level Ups (Last Assessment = yesterday)
    for row in rows:
        if get(row, 'Enrollment Status') != 'Enrolled':
            continue

        last_assessment = normalize_date(get(row, 'Last Assessment'))
        if last_assessment != yesterday_str:
            continue

        name = str(get(row, 'Student Name') or '').strip()
        if not name or name == 'x x':
            continue

        emails = parse_emails(get(row, 'Guardian Emails') or get(row, 'Guardian Email List'))
        if not emails:
            print(f'  ⚠️  {name} has no guardian email — skipping')
            continue

        triggered_students.append({
            'studentName':       name,
            'firstName':         name.split(' ')[0],
            'emailType':         'level-up',
            'emails':            emails,
            'grade':             str(get(row, 'Grade') or ''),
            'center':            str(get(row, 'Center') or ''),
            'leadId':            str(get(row, 'Lead Id') or get(row, 'LeadId') or ''),
            'lastAssessment':    last_assessment,
            'lastProgressCheck': normalize_date(get(row, 'Last Progress Check')),
            'learningPlanStatus': 'pending',
            'learningPlanPath':   None,
        })
        level_up_names.add(name.lower())
        print(f'  ✓ Level Up: {name}')

    # Second pass — standalone Progress Checks
    for row in rows:
        if get(row, 'Enrollment Status') != 'Enrolled':
            continue

        last_pc = normalize_date(get(row, 'Last Progress Check'))
        if last_pc != yesterday_str:
            continue

        name = str(get(row, 'Student Name') or '').strip()
        if not name or name == 'x x':
            continue
        if name.lower() in level_up_names:
            continue

        # Check if Assessment is within the level-up window
        last_assessment = normalize_date(get(row, 'Last Assessment'))
        if last_assessment:
            assess_date = parse_date(last_assessment)
            pc_date     = parse_date(last_pc)
            if assess_date and pc_date:
                diff = abs((assess_date - pc_date).days)
                if diff <= LEVEL_UP_WINDOW_DAYS:
                    emails = parse_emails(get(row, 'Guardian Emails') or get(row, 'Guardian Email List'))
                    if not emails:
                        print(f'  ⚠️  {name} has no guardian email — skipping')
                        continue
                    print(f'  ✓ Level Up (late-graded PC): {name} — {diff} day(s) gap')
                    triggered_students.append({
                        'studentName':       name,
                        'firstName':         name.split(' ')[0],
                        'emailType':         'level-up',
                        'emails':            emails,
                        'grade':             str(get(row, 'Grade') or ''),
                        'center':            str(get(row, 'Center') or ''),
                        'leadId':            str(get(row, 'Lead Id') or get(row, 'LeadId') or ''),
                        'lastAssessment':    last_assessment,
                        'lastProgressCheck': last_pc,
                        'learningPlanStatus': 'pending',
                        'learningPlanPath':   None,
                    })
                    continue

        emails = parse_emails(get(row, 'Guardian Emails') or get(row, 'Guardian Email List'))
        if not emails:
            print(f'  ⚠️  {name} has no guardian email — skipping')
            continue

        triggered_students.append({
            'studentName':       name,
            'firstName':         name.split(' ')[0],
            'emailType':         'progress-check',
            'emails':            emails,
            'grade':             str(get(row, 'Grade') or ''),
            'center':            str(get(row, 'Center') or ''),
            'leadId':            str(get(row, 'Lead Id') or get(row, 'LeadId') or ''),
            'lastAssessment':    last_assessment,
            'lastProgressCheck': last_pc,
            'learningPlanStatus': 'pending',
            'learningPlanPath':   None,
        })
        print(f'  ✓ Progress Check: {name}')

    level_ups = sum(1 for s in triggered_students if s['emailType'] == 'level-up')
    pcs       = sum(1 for s in triggered_students if s['emailType'] == 'progress-check')
    print(f'\n  Total: {len(triggered_students)} ({level_ups} level-up, {pcs} progress-check)')
    return triggered_students

# ─────────────────────────────────────────────
# 4. DOWNLOAD LEARNING PLANS
# ─────────────────────────────────────────────
def download_learning_plans(page, triggered_students):
    print('\n📚 Downloading Learning Plans...')
    for student in triggered_students:
        try:
            file_path = download_single_learning_plan(page, student)
            student['learningPlanStatus'] = 'ready'
            student['learningPlanPath']   = str(file_path)
            print(f"  ✓ {student['studentName']}")
        except Exception as e:
            print(f"  ⚠️  {student['studentName']} — LP failed: {e}")
            student['learningPlanStatus'] = 'error'
            safe = student['studentName'].replace(' ', '-')
            page.screenshot(path=str(DEBUG_DIR / f'lp-error-{safe}.png'))

def download_single_learning_plan(page, student):
    url = f"{RADIUS_BASE_URL}/Student/Details/{student['leadId']}"
    page.goto(url)
    page.wait_for_load_state('networkidle')

    lp_button = page.locator('a.k-grid-LPReport').first()
    if lp_button.count() == 0:
        raise Exception('LP Report button not found')

    with page.expect_download(timeout=15000) as dl_info:
        lp_button.click()
    download = dl_info.value

    safe_name = re.sub(r'[^a-z0-9]', '-', student['studentName'].lower())
    file_path = DOWNLOAD_DIR / f"lp-{safe_name}-{folder_date_str}.pdf"
    download.save_as(str(file_path))
    return file_path

# ─────────────────────────────────────────────
# 5. UPLOAD TO GOOGLE DRIVE
# ─────────────────────────────────────────────
def upload_to_google_drive(student_report_path, triggered_students):
    print('\n☁️  Uploading to Google Drive...')

    key_data = json.loads(os.environ['GOOGLE_SERVICE_ACCOUNT_KEY'])
    creds = service_account.Credentials.from_service_account_info(
        key_data,
        scopes=['https://www.googleapis.com/auth/drive.file']
    )
    drive = build('drive', 'v3', credentials=creds)

    # Create dated subfolder
    folder_meta = {
        'name': folder_date_str,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [os.environ['GOOGLE_DRIVE_FOLDER_ID']]
    }
    folder = drive.files().create(body=folder_meta, fields='id').execute()
    folder_id = folder['id']
    print(f'  ✓ Created folder: {folder_date_str}')

    # Upload Student Report
    upload_file(drive, student_report_path, folder_id,
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    # Upload Learning Plan PDFs
    for student in triggered_students:
        if student['learningPlanPath']:
            upload_file(drive, Path(student['learningPlanPath']), folder_id, 'application/pdf')

    # Write and upload manifest
    manifest = {
        'date': folder_date_str,
        'generatedAt': datetime.now().isoformat(),
        'studentReportFile': student_report_path.name,
        'triggeredStudents': [{
            'studentName':        s['studentName'],
            'firstName':          s['firstName'],
            'emailType':          s['emailType'],
            'emails':             s['emails'],
            'grade':              s['grade'],
            'center':             s['center'],
            'leadId':             s['leadId'],
            'lastAssessment':     s['lastAssessment'],
            'lastProgressCheck':  s['lastProgressCheck'],
            'learningPlanStatus': s['learningPlanStatus'],
            'learningPlanFile':   Path(s['learningPlanPath']).name if s['learningPlanPath'] else None,
        } for s in triggered_students]
    }

    manifest_path = DOWNLOAD_DIR / f'manifest-{folder_date_str}.json'
    manifest_path.write_text(json.dumps(manifest, indent=2))
    upload_file(drive, manifest_path, folder_id, 'application/json')

    print(f"  ✓ Uploaded manifest with {len(triggered_students)} students")

def upload_file(drive, file_path, folder_id, mime_type):
    file_path = Path(file_path)
    media = MediaFileUpload(str(file_path), mimetype=mime_type)
    drive.files().create(
        body={'name': file_path.name, 'parents': [folder_id]},
        media_body=media
    ).execute()
    print(f'  ✓ Uploaded: {file_path.name}')

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
def parse_emails(email_str):
    if not email_str:
        return []
    return [
        e.strip().lower()
        for e in str(email_str).split(',')
        if '@' in e and 'mathnasium.com' not in e
    ]

def normalize_date(val):
    if val is None:
        return None
    if isinstance(val, (datetime,)):
        return f"{val.month}/{val.day}/{val.year}"
    s = str(val).strip().split(' ')[0]
    m = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})', s)
    if m:
        return f"{int(m.group(1))}/{int(m.group(2))}/{m.group(3)}"
    return s

def parse_date(date_str):
    try:
        parts = date_str.split('/')
        return datetime(int(parts[2]), int(parts[0]), int(parts[1]))
    except Exception:
        return None

if __name__ == '__main__':
    main()
