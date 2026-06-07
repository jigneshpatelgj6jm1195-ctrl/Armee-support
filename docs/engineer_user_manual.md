# ArMee Technology Services Pvt. Ltd.
## Field Engineer User Guide — School Complaint Management App

Welcome to the **ArMee School Complaint Management App**. This guide provides step-by-step instructions for using the application in the field to log equipment complaints for Gyan Kunj (GK) and ICT projects in Gujarat.

---

## Table of Contents
1. [App Installation (PWA)](#1-app-installation-pwa)
2. [Step-by-Step Logging Flow](#2-step-by-step-logging-flow)
3. [Troubleshooting & FAQs](#3-troubleshooting--faqs)

---

## 1. App Installation (PWA)
The application is a Progressive Web App (PWA) which means it can be installed directly onto your mobile home screen or laptop desktop without downloading from an app store.

### For Android Mobile (Chrome)
1. Open the application URL provided by your coordinator (e.g., `https://complaint.armee.in` or `http://localhost:8765/complaint_form.html`).
2. Tap the **three dots** in the top-right corner of Chrome.
3. Tap **"Add to Home screen"** or **"Install App"**.
4. Confirm installation. The ArMee icon will appear on your home screen.

### For Windows Laptop (Edge/Chrome)
1. Open the application URL in Edge or Chrome.
2. In the address bar, look for the **Install App icon** (looks like a monitor with an arrow or three squares with a plus).
3. Click **Install**.
4. Right-click the app on your desktop and select **"Pin to Taskbar"** for quick access.

---

## 2. Step-by-Step Logging Flow

### Step 1: User Log In (First-time Setup)
When you open the app for the first time, you will see a profile setup screen.
- Enter your **Full Name**.
- Enter your **10-Digit Mobile Number** (Must match your registered number in the User Master).
- Select your **Role** (e.g. Field Engineer).
- Tap **Continue**.

### Step 2: Select Project
On the home screen, select the project category:
- **GK (Gyan Kunj)**
- **ICT**

### Step 3: School Lookup
- Type the 11-digit **DISE Code** of the school.
- The app automatically queries the database and populates:
  - School Name
  - School Code
  - District & Taluka
  - Principal Name & Contact Number
  - School Address & Pin Code

### Step 4: Verify & Edit Contact Info
- Look at the **Principal** and **Contact** chips.
- Verify details with the school staff.
- If incorrect, tap the **✏️ Edit** button on the chip, enter the new name/number, and tap **Save**. This updates the master records instantly.

### Step 5: Complaint Details & Quantity
- Select the **Equipment** (e.g., Desktop Computer, Projector).
- Select the **Nature of Complaint** (e.g., No Display, slow performance).
- Enter the **Quantity** of faulty items:
  - If Quantity = 1: Proceed normally.
  - If Quantity > 1: The app sets up a **Group Flow** to help you register each item sequentially.

### Step 6: Barcode Scanning & Serial Numbers
- Type the **Serial Number** of the equipment.
- Alternatively, tap the **📷 Scan** button.
- Align the barcode or QR code on the back of the device inside the highlighted viewfinder box to auto-capture the Serial Number.

### Step 7: Capture Photo (Mandatory 📸)
- Tap the **📷 Camera** button to take a live photo of the faulty equipment.
- Tap **Gallery** to attach files or existing images.
- *Note:* At least **one photo is mandatory** to submit a complaint.

### Step 8: GPS Location Capture (Mandatory 📍)
- When submitting, the app will automatically request your GPS location.
- You must allow Chrome/Edge to access your **Location / GPS**.
- If GPS access is denied, submission will be blocked.

### Step 9: Submit & Ticket Generation
- Review the details in the **Complaint Summary**.
- Click **Submit Complaint**.
- The app will generate a unique **Case ID (Ticket Number)** for tracking.
- If quantity was > 1, tap **Next Item** on the success pop-up to log details for the next item.

### Step 10: Post-Submission Options
After completing all submissions, you can choose to:
1. **Register Another Complaint**: Start fresh for a new school.
2. **Share Summary**: Copy/Share a text report of the ticket details via WhatsApp or Email.
3. **Exit**: Close the app.

---

## 3. Troubleshooting & FAQs

### How does offline logging work?
If you lose mobile network connectivity at a school, you can still fill out the entire form, scan barcodes, and capture photos. When you tap **Submit**, the app securely queues your complaint. The next time you are online (e.g. at home or on the road), the app will automatically upload all queued tickets to the server.

### The Camera/GPS is blocked. How do I fix it?
1. Go to your browser settings.
2. Under **Site Settings**, locate the app URL.
3. Set **Camera** and **Location** permissions to **"Allow"**.
