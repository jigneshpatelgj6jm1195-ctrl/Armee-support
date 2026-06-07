# ArMee Technology Services Pvt. Ltd.
## Administrator User Guide — Complaint Management System

This guide outlines how to use the central **ArMee Admin Panel** to manage the school database, equipment lists, authorized engineers, monitor live complaints, and export reports for Power BI / Excel analysis.

---

## Accessing the Admin Panel
The Admin Panel is accessible at the following URL when running the local server:
`http://localhost:8765/admin.html`

---

## 1. Dashboard Tab
The Dashboard provides a real-time summary of logging activity across Gujarat.
- **KPI Metrics:**
  - **Total Tickets:** Count of all complaints registered.
  - **Open / In Progress:** Active tickets currently pending technical resolution.
  - **Resolved Tickets:** Count of tickets completed.
  - **Total Schools:** Number of schools in the database.
- **Visual Analytics:**
  - **Complaints by Equipment:** Horizontal bar chart showing which items fail most frequently.
  - **Complaints by Project:** Percentage breakdown of Gyan Kunj (GK) vs. ICT complaints.
- **Activity Feed:** Live log of the 5 most recent complaints showing Case ID, School, Equipment, and Status.

---

## 2. School Master Tab
Allows managing the database of all schools in Gujarat.
- **School Master Excel Upload:**
  - Drag and drop or select the primary database spreadsheet (`School complaint format 23.6.25.xlsx`).
  - Clicking upload processes the sheet and automatically updates `school_data.json` without stopping the web app.
- **Search & Update School Records:**
  - Enter any school's **DISE Code** and click **Search**.
  - Review or edit details (Taluka, Principal Name, Principal Mobile, Pin Code, Address).
  - Click **Update School Record** to save.

---

## 3. Equipment Master Tab
Enables full control over the Equipment and Nature of Complaint options displayed in the field engineer form.
- **Add Equipment:** Click **＋ Add Equipment**, enter the name, and click save.
- **Toggle Equipment Status:** Click **Disable** to hide an equipment category from the form, or **Enable** to make it active.
- **Complaint Natures:** Under each equipment card:
  - Add complaint descriptions (e.g. "Not starting").
  - Toggle individual complaint descriptions active/inactive.
  - Delete unused complaint categories.
- *Note:* Remember to click **Save Changes** in the bottom save bar to persist updates!

---

## 4. User Master Tab
Restricts access to authorized ArMee Field Engineers.
- **Add User:** Click **＋ Add Engineer**, enter the engineer's Name, Mobile Number, and Role.
- **Active/Inactive Toggle:** Click **Disable** on a user row to block that engineer from submitting complaints or logging in.
- **Delete User:** Remove the engineer's authorization record.
- *Security:* If this list is empty, anyone can access the app. Once a user is added, the login page is locked down and validates against active phone numbers.

---

## 5. Complaints & Reports Tab
Central log of all submissions with advanced query and data export features.
- **Data Filtering:** Filter complaints by Project (GK/ICT), Equipment type, Ticket Status (Open, In Progress, Resolved), or search by school name/serial number.
- **Download Excel Report:** Click to export the filtered dataset into a branded, styled Excel-compatible file (`.xls`) containing custom ArMee headers.
- **Download CSV Report:** Click to download standard CSV files for import into database systems or Power BI dashboards.
- **Manage Ticket Status:** Update a ticket's status (Open ➔ In Progress ➔ Resolved ➔ Closed) using the status dropdown menu on each row. Click **Save Changes** to write updates to the database file.
