/* ═══════════════════════════════════════════════════════════════
   Armee ICT Support Scraper
   ───────────────────────────────────────────────────────────────
   Logs into ssgujarat.org/CAL (captcha solved with Tesseract OCR),
   downloads the Complaint List "Export To Excel" file, and POSTs
   every Pending / InProgress row to the Armee Apps Script backend
   (action: import_department_complaints). The backend dedupes by
   TicketId and emails branch admins about genuinely new tickets.

   READ-ONLY on the government portal — this never writes to it.

   Usage:
     node run.js                      full scrape + push
     node run.js --test-captcha       OCR the login captcha 5x and report (no login)
     node run.js --file export.xlsx   skip the portal, push an already-downloaded export
     node run.js --dry-run            scrape + parse but do NOT push

   Config: .env file next to this script (see .env.example).
═══════════════════════════════════════════════════════════════ */
require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const { chromium } = require('playwright');
const { createWorker } = require('tesseract.js');
const { execSync } = require('child_process');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const PORTAL_URL = 'https://www.ssgujarat.org/CAL/CALLogin.aspx';
const SCRIPT_URL = process.env.SCRIPT_URL;
const IMPORT_KEY = process.env.IMPORT_KEY;
const PORTAL_USER = process.env.PORTAL_USER;
const PORTAL_PASS = process.env.PORTAL_PASS;
const SEND_EMAILS = (process.env.SEND_EMAILS || 'true').toLowerCase() !== 'false';
const HEADLESS = (process.env.HEADLESS || 'true').toLowerCase() !== 'false';
const CHUNK_SIZE = 150;
const LOGIN_ATTEMPTS = 3;

const EXPORT_COLUMNS = [
  'District', 'BlockId', 'Block', 'ClusterId', 'Cluster', 'VillageId', 'Village',
  'SchoolId', 'School', 'TicketId', 'Agency', 'Asset_Type', 'Device_Type', 'Issue_Type',
  'Issue_Details', 'Issue_Photo', 'Contact_Name', 'Phone_Number', 'Time_Preference_Call',
  'Ticket_Status', 'CreatedBy', 'CreatedDate', 'UpdatedBy', 'UpdatedDate',
  'Diagnosis_Notes_Agency', 'Schedule_Visit_Date', 'Technician_Name', 'Technician_Number',
  'Issue_Resolved_By', 'TotalDaysOfTicket'
];

function log(msg) {
  console.log(new Date().toISOString().replace('T', ' ').substring(0, 19) + '  ' + msg);
}

async function postJson(payload) {
  const res = await fetch(SCRIPT_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
    redirect: 'follow',
  });
  return res.json();
}

/** Failure alert → backend emails the configured ALERT_EMAIL (or the sheet owner). */
async function sendAlert(message) {
  try {
    await postJson({ action: 'scraper_alert', importKey: IMPORT_KEY, message });
    log('Alert email requested.');
  } catch (e) {
    log('Could not send alert: ' + e.message);
  }
}

/* ── Captcha OCR ── */
const Jimp = require('jimp');
let ocrWorker = null;

/**
 * The portal captcha is red glyphs on a light noisy background. Keeping only
 * red-dominant pixels (black on white) and upscaling 4x makes OCR far more
 * reliable than the raw image.
 */
async function preprocessCaptcha(pngBuffer) {
  const img = await Jimp.read(pngBuffer);
  img.scan(0, 0, img.bitmap.width, img.bitmap.height, function (x, y, idx) {
    const r = this.bitmap.data[idx], g = this.bitmap.data[idx + 1], b = this.bitmap.data[idx + 2];
    const isRed = r > 100 && r > g * 1.4 && r > b * 1.4;
    const v = isRed ? 0 : 255;
    this.bitmap.data[idx] = v; this.bitmap.data[idx + 1] = v; this.bitmap.data[idx + 2] = v;
  });
  img.scale(4, Jimp.RESIZE_NEAREST_NEIGHBOR);
  return img.getBufferAsync(Jimp.MIME_PNG);
}

async function ocrCaptcha(pngBuffer) {
  // 1. Try solving via ddddocr Python helper first
  const tempPath = path.join(__dirname, 'downloads', 'temp_captcha_' + Date.now() + '.png');
  try {
    fs.mkdirSync(path.dirname(tempPath), { recursive: true });
    fs.writeFileSync(tempPath, pngBuffer);
  } catch (err) {
    log('Failed to write temp captcha file: ' + err.message);
  }

  let ddddocrResult = null;
  if (fs.existsSync(tempPath)) {
    for (const pyCmd of ['python3', 'python']) {
      try {
        const solverPath = path.join(__dirname, 'solve_captcha.py');
        const stdout = execSync(`"${pyCmd}" "${solverPath}" "${tempPath}"`, { encoding: 'utf8', stdio: ['pipe', 'pipe', 'ignore'] }).trim();
        if (stdout && stdout !== 'ddddocr_not_installed' && !stdout.startsWith('error_') && !stdout.includes('missing_image_path')) {
          ddddocrResult = stdout.replace(/[^A-Z0-9]/gi, '').toUpperCase();
          log(`ddddocr guessed: "${ddddocrResult}"`);
          break;
        }
      } catch (e) {
        // Ignore and try next interpreter or fallback
      }
    }
    try { fs.unlinkSync(tempPath); } catch (e) {}
  }

  if (ddddocrResult && ddddocrResult.length >= 4) {
    return ddddocrResult;
  }

  log('ddddocr not available or failed. Falling back to Tesseract OCR...');

  // 2. Fallback to Tesseract OCR
  if (!ocrWorker) {
    ocrWorker = await createWorker('eng');
    await ocrWorker.setParameters({
      tessedit_char_whitelist: 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789',
      tessedit_pageseg_mode: '7', // single text line
    });
  }
  const processed = await preprocessCaptcha(pngBuffer);
  const { data } = await ocrWorker.recognize(processed);
  return (data.text || '').replace(/[^A-Z0-9]/gi, '').toUpperCase();
}

async function grabCaptcha(page) {
  const img = page.locator('img[src*="CaptchaImage.axd"]');
  await img.waitFor({ state: 'visible', timeout: 15000 });
  return img.screenshot();
}

/* ── Login (retries with fresh captcha) ── */
async function login(page) {
  for (let attempt = 1; attempt <= LOGIN_ATTEMPTS; attempt++) {
    await page.goto(PORTAL_URL, { waitUntil: 'domcontentloaded' });
    const guess = await ocrCaptcha(await grabCaptcha(page));
    log(`Login attempt ${attempt}: captcha OCR guess "${guess}"`);
    if (guess.length < 4) { log('Guess too short, refreshing captcha.'); continue; }

    await page.fill('#TxtUName', PORTAL_USER);
    await page.fill('#TxtUPass', PORTAL_PASS);
    await page.fill('#txtCaptcha', guess);
    await Promise.all([
      page.waitForLoadState('domcontentloaded'),
      page.click('#ImgSubmit'),
    ]);
    await page.waitForTimeout(2000);

    if (!page.url().includes('CALLogin')) {
      log('Login successful → ' + page.url());
      return true;
    }
    log('Login attempt ' + attempt + ' failed (wrong captcha or credentials).');
  }
  return false;
}

/* ── Download the Excel export ── */
async function downloadExport(page) {
  await page.click('text=Complaint List');
  await page.waitForLoadState('domcontentloaded');
  await page.waitForTimeout(3000); // grid loads via postback

  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 120000 }),
    page.click('text=Export To Excel'),
  ]);
  const filePath = path.join(__dirname, 'downloads', 'export_' + Date.now() + '.xlsx');
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  await download.saveAs(filePath);
  log('Export downloaded: ' + filePath);
  return filePath;
}

/* ── Parse + filter ── */
function parseExport(filePath) {
  const wb = XLSX.readFile(filePath);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { raw: false, defval: '' });
  if (!rows.length) throw new Error('Export file has no data rows');
  for (const col of ['TicketId', 'Ticket_Status', 'District', 'SchoolId']) {
    if (!(col in rows[0])) throw new Error('Export format changed — missing column: ' + col);
  }
  const open = [];
  let skipped = 0;
  for (const r of rows) {
    const norm = String(r.Ticket_Status || '').trim().toLowerCase().replace(/[\s-]/g, '');
    if (norm !== 'pending' && norm !== 'inprogress') { skipped++; continue; }
    const clean = {};
    for (const col of EXPORT_COLUMNS) clean[col] = r[col] !== undefined ? String(r[col]) : '';
    open.push(clean);
  }
  log(`Parsed ${rows.length} rows → ${open.length} open (Pending/InProgress), ${skipped} skipped.`);
  return open;
}

/* ── Push to backend in chunks ── */
async function pushRows(rows) {
  let inserted = 0, updated = 0;
  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const chunk = rows.slice(i, i + CHUNK_SIZE);
    const out = await postJson({
      action: 'import_department_complaints',
      importKey: IMPORT_KEY,
      importSource: 'AUTO_SCRAPE',
      sendEmails: SEND_EMAILS,
      rows: chunk,
    });
    if (out.status !== 'ok') throw new Error('Backend rejected chunk: ' + (out.message || JSON.stringify(out)));
    inserted += out.inserted; updated += out.updated;
    log(`Pushed ${Math.min(i + chunk.length, rows.length)}/${rows.length} — running total: ${inserted} new, ${updated} updated.`);
  }
  return { inserted, updated };
}

/* ── Modes ── */
async function testCaptchaMode() {
  log('CAPTCHA TEST MODE: OCR-ing 5 fresh captchas from the live login page (no login).');
  const browser = await chromium.launch({ headless: HEADLESS });
  const page = await browser.newPage();
  for (let i = 1; i <= 5; i++) {
    await page.goto(PORTAL_URL, { waitUntil: 'domcontentloaded' }); // fresh captcha each load
    const buf = await grabCaptcha(page);
    fs.mkdirSync(path.join(__dirname, 'downloads'), { recursive: true });
    fs.writeFileSync(path.join(__dirname, 'downloads', `captcha_${i}.png`), buf);
    fs.writeFileSync(path.join(__dirname, 'downloads', `captcha_${i}_processed.png`), await preprocessCaptcha(buf));
    const guess = await ocrCaptcha(buf);
    log(`Captcha ${i}: OCR guess "${guess}" (raw + processed images saved in downloads/)`);
  }
  await browser.close();
  if (ocrWorker) await ocrWorker.terminate();
}

async function main() {
  const args = process.argv.slice(2);

  if (args.includes('--test-captcha')) return testCaptchaMode();

  if (!SCRIPT_URL || !IMPORT_KEY) throw new Error('SCRIPT_URL and IMPORT_KEY must be set in .env');

  const fileArgIdx = args.indexOf('--file');
  if (fileArgIdx !== -1) {
    const filePath = args[fileArgIdx + 1];
    if (!filePath || !fs.existsSync(filePath)) throw new Error('--file: file not found: ' + filePath);
    const rows = parseExport(filePath);
    if (args.includes('--dry-run')) { log('Dry run — not pushing.'); return; }
    const result = await pushRows(rows);
    log(`DONE (file mode): ${result.inserted} new ticket(s), ${result.updated} updated.`);
    return;
  }

  if (!PORTAL_USER || !PORTAL_PASS) throw new Error('PORTAL_USER and PORTAL_PASS must be set in .env');

  const browser = await chromium.launch({ headless: HEADLESS });
  try {
    const page = await browser.newPage({ acceptDownloads: true });
    const ok = await login(page);
    if (!ok) throw new Error('Login failed after ' + LOGIN_ATTEMPTS + ' captcha attempts');

    const filePath = await downloadExport(page);
    const rows = parseExport(filePath);
    if (args.includes('--dry-run')) { log('Dry run — not pushing.'); return; }
    const result = await pushRows(rows);
    log(`DONE: ${result.inserted} new ticket(s), ${result.updated} updated.`);

    // keep only the 10 newest downloads
    const dir = path.join(__dirname, 'downloads');
    const files = fs.readdirSync(dir).filter(f => f.startsWith('export_')).sort();
    files.slice(0, -10).forEach(f => fs.unlinkSync(path.join(dir, f)));
  } finally {
    await browser.close();
    if (ocrWorker) await ocrWorker.terminate();
  }
}

main().catch(async (err) => {
  log('FAILED: ' + err.message);
  if (!process.argv.includes('--test-captcha') && !process.argv.includes('--dry-run')) {
    await sendAlert('ICT scraper run failed at ' + new Date().toISOString() + '\n\n' + err.message);
  }
  process.exit(1);
});
