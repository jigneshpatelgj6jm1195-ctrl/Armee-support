# Armee ICT Support Scraper

Automatically pulls open tickets (Pending / InProgress) from the government ICT
Support portal (ssgujarat.org/CAL) and pushes them into the Armee complaint
system every 30 minutes. The backend deduplicates by TicketId and emails each
branch admin a digest of genuinely new tickets. **Read-only on the portal.**

## One-time setup (Ubuntu VM)

```bash
# 1. Node.js 20+ (if not installed)
curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash - && sudo apt-get install -y nodejs

# 2. Copy this scraper folder to the VM, e.g. /opt/armee-ict/scraper, then:
cd /opt/armee-ict/scraper
npm install
npx playwright install --with-deps chromium

# 3. Credentials
cp .env.example .env
nano .env          # fill PORTAL_PASS and IMPORT_KEY
chmod 600 .env

# 4. Test by hand first
node run.js --test-captcha     # OCR accuracy check, no login
node run.js --dry-run          # full run, but nothing pushed
node run.js                    # real run

# 5. Schedule (VM must be on Asia/Kolkata time: sudo timedatectl set-timezone Asia/Kolkata)
crontab -e
# every 30 min, 07:00–19:30, Mon–Sat:
*/30 7-19 * * 1-6 cd /opt/armee-ict/scraper && /usr/bin/node run.js >> /var/log/armee-scraper.log 2>&1
```

## Windows (interim option — run on the office PC)

```powershell
cd F:\Office\Automatization\form\scraper
npm install
npx playwright install chromium
copy .env.example .env    # then edit .env
node run.js
```
Schedule with Task Scheduler: action = `node run.js`, start in the scraper folder,
trigger every 30 minutes between 07:00 and 19:30, Mon–Sat.

## Modes

| Command | What it does |
|---|---|
| `node run.js` | Login → download export → push open tickets |
| `node run.js --test-captcha` | Grab 5 live captchas, print OCR guesses, save images |
| `node run.js --dry-run` | Everything except the push |
| `node run.js --file export.xlsx` | Skip the portal — push an export downloaded by hand |

## Failure behaviour
- The captcha is retried up to 3 times per run (fresh image each time).
- Any failed run sends an **alert email** (backend `scraper_alert` action → the
  address in the Apps Script `ALERT_EMAIL` script property, or the sheet owner).
- A missed run is harmless: the next run picks everything up (imports are
  idempotent — existing TicketIds are updated, never duplicated).
