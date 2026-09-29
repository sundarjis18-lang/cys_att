# CYS Attendance System

Mobile-first web application for recording attendance into Google Sheets for Computer Science with Cyber Security.

## Features
- **Attendance Markings**:
  - Full Day Present: `/`
  - Full Day Absent: `a`
  - Forenoon Half-Day Absent: `a/`
  - Afternoon Half-Day Absent: `/a`
- **Interactive Student Roster**: Tap any student to cycle presence / half-day status.
- **Roll Number Keypad Entry**: Fast numeric roll suffix entry with interactive absentee chips.
- **Real-Time Google Sheets Sync**: Directly updates the bound Google Sheet via Apps Script Web App.
- **Zero Config**: Connected by default to the CYS attendance spreadsheet.

---

## File Structure

```
├── index.html        # Main web application (GitHub Pages entry point)
├── Code.gs           # Google Apps Script code for the Google Sheet
├── SETUP.md          # Setup and configuration guide
├── README.md         # Repository documentation
└── .gitignore        # Git ignore rules
```

---

## How to Enable GitHub Pages

1. Push this repository to GitHub:
   ```bash
   git init
   git add .
   git commit -m "Initial commit - CYS Attendance v2 with Half-day support"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<repo-name>.git
   git push -u origin main
   ```
2. Go to your repository on GitHub.
3. Click **Settings** → **Pages** (under the "Code and automation" sidebar).
4. Under **Build and deployment**:
   - **Source**: `Deploy from a branch`
   - **Branch**: `main` / `/ (root)`
5. Click **Save**.
6. In ~1 minute, GitHub will give you your live URL:
   `https://<your-username>.github.io/<repo-name>/`
