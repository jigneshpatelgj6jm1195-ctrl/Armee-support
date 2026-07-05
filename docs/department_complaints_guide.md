# Department Complaints — User Guide

*(ssgujarat.org government portal tickets inside the Armee system — July 2026)*

## For field engineers (mobile app — support.armee.online)

1. Open the app. After your name/phone setup you now see **two choices**:
   - **📝 Advance Complaint** — the normal complaint form, unchanged.
   - **🏛️ Department Complaint** — government portal tickets.
2. In Department Complaint, type the school's **DISE code** and press Search.
   Every open portal ticket for that school appears with its details.
3. After fixing the issue, pick one action on the ticket:
   - **✅ Close with OTP** — the school contact tells you the OTP; type it in, add notes, Confirm. Done.
   - **📩 Close without OTP** — couldn't get the OTP on site? Use this; the office will finish it later.
   - **🔧 Part Request** — the repair needs a spare part; the ticket waits in the Part Request queue.
4. Teachers never see this menu — only engineers.

## For the back office (admin panel → 🏛️ Dept Complaints tab, Super Admin)

- **Pendency boxes** — click any number (Pending / In Progress / Part Request / Pending OTP / Closed) to see those tickets.
- **Branch table** — every branch's open tickets split by age: 0–2 / 3–5 / 6+ business days (Sundays not counted). Red numbers = overdue. Click a row for that branch's tickets, oldest first.
- **📩 Enter OTP button** — appears on tickets an engineer closed without OTP. Click it, type the OTP collected from the school, and the ticket closes properly.
- **Bulk Upload** — drag in the portal's "Export To Excel" file any time; duplicates are impossible (same ticket number = update, not copy).

## For the back office (🏢 Branch Management tab, Super Admin)

- Add/edit District Offices and Branches; set each **Branch Manager's name** so it shows on the dashboard.
- **Unmapped Values** — if a new district spelling ever appears in imports, it shows here; pick the branch and click Map. All its tickets (past and future) move instantly.
- Deleting anything is blocked until nothing references it, and is always recoverable (Status becomes Inactive).

## The automatic scraper (scraper folder)

- Every 30 minutes (7:00–19:30, Mon–Sat) it logs into ssgujarat.org, solves the captcha,
  downloads the complaint export, and imports new open tickets automatically.
  Branch admins get one digest email per run listing their new tickets.
- If a run fails, **jignesh.patel@armeeinfotech.com receives an alert email**. A missed run
  is harmless — the next one catches up.
- Logs: `scraper\logs\scraper.log`. To pause/resume: Task Scheduler → "Armee ICT Scraper" → Disable/Enable.

## Emergency rollback

- Website: previous versions in the Vercel dashboard (project *armee-support*).
- Backend: Apps Script → Deploy → Manage deployments → pencil → pick an older version → Deploy.
