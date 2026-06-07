# ArMee Technology Services Pvt. Ltd.
## Database Schema, JSON Structures & API Documentation

This document describes the data architectures, file structures, relationships, and HTTP server endpoints used by the School Complaint Management Application.

---

## 1. File-Based Database Structures
The application uses flat-file JSON databases for local speed, simplicity, and offline compatibility.

### 1.1 Master Data (`master_data.json`)
Stores the equipment lists, active complaint types, and authorized engineers.
```json
{
  "equipment": [
    {
      "id": "EQ01",
      "name": "Desktop Computer",
      "status": "active",
      "complaints": [
        { "id": "NC001", "name": "Not Starting / Dead", "status": "active" }
      ]
    }
  ],
  "users": [
    {
      "id": "USR100",
      "name": "Rajesh Patel",
      "phone": "9876543210",
      "role": "Field Engineer",
      "status": "active"
    }
  ]
}
```

### 1.2 Complaints Log (`complaints.json`)
Centrally logs all complaints logged by engineers.
```json
[
  {
    "submittedAt": "2026-06-04T12:00:00.000Z",
    "complainantName": "Rajesh Patel",
    "complainantPhone": "9876543210",
    "complainantRole": "Field Engineer",
    "project": "GK",
    "dise": "24010100101",
    "schoolCode": "SCH001",
    "district": "Ahmedabad",
    "taluka": "City",
    "school": "Model School Ahmedabad",
    "principal": "A. K. Sharma",
    "contact": "9825098250",
    "address": "Ahmedabad, Gujarat",
    "pincode": "380001",
    "equipment": "Desktop Computer",
    "natureOfComplaint": "Not Starting / Dead",
    "serialNumber": "SN2026889",
    "quantity": 1,
    "itemNumber": 1,
    "groupId": "",
    "complaintDate": "2026-06-04",
    "medium": "Visit",
    "description": "System dead after power surge.",
    "photoCount": 1,
    "latitude": 23.0225,
    "longitude": 72.5714,
    "status": "Open",
    "caseId": "CASE-1464627889"
  }
]
```

### 1.3 School Database (`school_data.json`)
Indexed by DISE code (11 characters) for O(1) instant school details lookup.
```json
{
  "24010100101": [
    {
      "project": "GK",
      "dise": "24010100101",
      "schoolCode": "SCH001",
      "district": "Ahmedabad",
      "block": "City",
      "school": "Model School Ahmedabad",
      "principal": "A. K. Sharma",
      "mobile": "9825098250",
      "address": "Ahmedabad, Gujarat",
      "pincode": "380001"
    }
  ]
}
```

---

## 2. API Endpoints
Served by `export_and_launch.py` (Local HTTP Server running on `http://localhost:8765`).

### 2.1 Submit Complaint
- **Endpoint:** `POST /submit_complaint`
- **Body:** JSON payload matching the Complaints Log schema.
- **Description:** Centrally stores complaint data by appending the submission object to `complaints.json`.
- **Response:** `200 OK` `{ "ok": true }`

### 2.2 Rebuild School Master
- **Endpoint:** `POST /upload_school_excel`
- **Body:** Raw binary bytes of Excel workbook file (.xlsx).
- **Description:** Overwrites `School complaint format 23.6.25.xlsx` on the server and parses sheet data to rebuild `school_data.json`.
- **Response:** `200 OK` `{ "ok": true, "count": 2540 }`

### 2.3 Update School Info
- **Endpoint:** `POST /update_school`
- **Body:** `{ "dise": "...", "field": "...", "newValue": "...", "oldValue": "..." }`
- **Description:** Modifies a school's principal name or contact number in `school_data.json`.
- **Response:** `200 OK` `{ "ok": true, "updated": 1 }`

### 2.4 Update Master Config
- **Endpoint:** `POST /update_master`
- **Body:** JSON object matching the Master Config schema.
- **Description:** Overwrites `master_data.json` to save equipment definitions or user listings.
- **Response:** `200 OK` `{ "ok": true }`

### 2.5 Re-Write Complaints Log
- **Endpoint:** `POST /update_complaints`
- **Body:** Complete updated list of complaints.
- **Description:** Replaces `complaints.json` with the updated list (used for admin status changes).
- **Response:** `200 OK` `{ "ok": true }`
