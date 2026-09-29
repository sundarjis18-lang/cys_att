# CYS Attendance System — Setup Guide

## What this does
- Mobile-first web app for marking attendance in your Google Sheet
- Supports **Full Day Present (`/`)**, **Full Day Absent (`a`)**, **Forenoon Half Day Absent (`a/`)**, and **Afternoon Half Day Absent (`/a`)**
- Enter absentee roll numbers or tap students in the roster
- Sheet updates instantly with complete attendance marks and summary records

---

## Step 1 — Add the Apps Script to your Google Sheet

1. Open your Google Sheet:  
   https://docs.google.com/spreadsheets/d/1SMZ48JP46DdSzpa0Krk9VRV6T1WSCGwMME41SuA6vMU/edit?usp=sharing

2. Click **Extensions → Apps Script**

3. Delete any existing code in the editor and paste the entire contents of **`Code.gs`**

4. Click **Save** (💾 icon or Ctrl+S)

---

## Step 2 — Deploy as a Web App

1. Click **Deploy → New deployment**
2. Click the ⚙️ gear icon → choose **Web app**
3. Set:
   - **Description**: `CYS Attendance v2 (with Half Day)`
   - **Execute as**: `Me (your-email@gmail.com)` *(IMPORTANT! Do not choose "User accessing the web app")*
   - **Who has access**: `Anyone` *(so your web app can communicate with it)*
4. Click **Deploy**
5. Click **Authorize access** and approve the permissions
6. Your active Web App URL is already configured by default:
   ```
   https://script.google.com/macros/s/AKfycbykag8-HpZ0TnINVJpGUvKOPCkBvpaPv1ad_ANCa8tGrNkkTDBSuLJbbG6Y1likvcS0YQ/exec
   ```

*(Note: Every time you edit `Code.gs` in Apps Script, go to **Deploy → Manage deployments → Edit ✏️ → Version: New version → Deploy** so the changes take effect!)*

---

## Step 3 — Open the Web App

Simply open **`index.html`** in your browser!
- Zero manual setup required — the new Google Sheet and Apps Script URL are pre-configured by default.
- Auto-fetches all 66 students and month sheets automatically.
3. Select **Month / Year / Week / Day Order**.
4. Choose absence type (**Full Day `a`**, **FN Half Day `a/`**, or **AN Half Day `/a`**), then type the roll suffix (e.g., `05`) and tap **+**.
   - You can also tap the badge on any added absentee chip to cycle between `Full`, `FN`, and `AN`.
   - Or open the **Student Roster** to tap-and-mark directly!
5. Tap **Mark Attendance** — all students in the sheet are recorded automatically!

---

## Attendance Marks Legend

| Symbol | Status | Meaning |
|--------|--------|---------|
| `/` | Present | Full Day Present |
| `a` | Full Day Absent | Absent both Forenoon & Afternoon |
| `a/` | FN Absent | Absent in Forenoon, Present in Afternoon |
| `/a` | AN Absent | Present in Forenoon, Absent in Afternoon |

---

## Sheet Column Mapping

| Column | Week | Day Order |
|--------|------|-----------|
| D | I | 1 |
| E | I | 2 |
| F | I | 3 |
| G | I | 4 |
| H | I | 5 |
| I | I | 6 |
| J | II | 1 |
| K | II | 2 |
| … | … | … |
| AA | IV | 6 |

---

## Notes
- The Apps Script scans ALL rows where column B has an 8-digit roll number.
- Works across multiple sections (Sem II, IV etc.) in one pass.
- After marking, a preview table displays each student's exact mark.
