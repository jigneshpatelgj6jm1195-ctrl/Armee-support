# ArMee Technology Services Pvt. Ltd.
## PWA Installation, Server Deployment & Backup Guide

This document details PWA installation instructions for mobile and desktop users, local server launching, data backups, and restoration procedures for the School Complaint Management System.

---

## 1. PWA Installation Guide

### Android Devices (Chrome)
1. Launch **Google Chrome** on your mobile device.
2. Navigate to your deployed application link: e.g. `https://complaint.armee.in` (or local test address `http://localhost:8765/complaint_form.html`).
3. If prompted by a bottom banner, tap **"Add School Complaint Form to Home screen"**.
4. If the banner doesn't appear:
   - Tap the **three dots** in the top-right corner.
   - Select **"Add to Home screen"** or **"Install App"**.
5. Once installed, launch the app directly from your home screen icon.

### iOS Devices (Safari)
1. Open **Safari** on your iPhone/iPad.
2. Go to the application URL.
3. Tap the **Share button** (square with an up arrow) at the bottom.
4. Scroll down and tap **"Add to Home Screen"**.
5. Tap **Add** in the top-right corner.

### Windows Laptop / Desktop (Edge / Chrome)
1. Open **Microsoft Edge** or **Google Chrome** on your laptop.
2. Go to the application URL.
3. In the address bar (right side), click the **App Available / Install Icon** (Edge shows a three-square icon with a plus; Chrome shows an installation icon).
4. Click **Install**.
5. Pin to your desktop and Taskbar for offline field usage.

---

## 2. Server Deployment & Launch Guide
To run the system locally on an admin machine or local network server:
1. Ensure **Python 3.x** is installed on the machine.
2. Place all project files in a single folder (e.g. `F:\Office\Automatization\form`).
3. Double-click the file named **`OPEN FORM.bat`**. This file runs `export_and_launch.py`, which:
   - Verifies libraries are installed (auto-installs `openpyxl` if missing).
   - Reads the school master sheet spreadsheet.
   - Updates `school_data.json`.
   - Starts a local HTTP server on port `8765`.
   - Opens the PWA form in your default web browser.

---

## 3. Backup and Restore Guide

All local complaints, user authorizations, and equipment masters are saved on the server inside the folder. It is critical to perform regular backups of these database files to prevent data loss.

### Databases to Backup
There are three main database files:
1. **`school_data.json`** — The exported school master records database.
2. **`master_data.json`** — The equipment masters, complaint natures list, and authorized users list.
3. **`complaints.json`** — The central repository of all field engineer ticket submissions.

### Backup Procedure
1. Create a folder named `backups` on an external storage drive, server, or cloud storage.
2. Copy the three database files (`school_data.json`, `master_data.json`, `complaints.json`) and the Excel spreadsheet (`School complaint format 23.6.25.xlsx`).
3. Save them in a dated folder, for example: `backups/ArMee_Backup_2026-06-04/`.
4. *Recommendation:* Perform this backup weekly or after modifying the school database.

### Restoration Procedure
If database files become corrupted or you migrate to a new server:
1. Stop the python server (close the command prompt window running `OPEN FORM.bat`).
2. Navigate to the project directory.
3. Copy the backup files from your backup directory and paste them into the project directory, overwriting the corrupted files.
4. Double-click `OPEN FORM.bat` to launch the server and resume operations.
