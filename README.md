# ⚡ Codeforces Daily Challenge & Penalty Tracker

An automated competitive programming accountability tracker built with **Google Sheets**, **Google Apps Script**, and **Discord**. 

Participants must solve **5 valid problems every day** on Codeforces. Every missed problem adds a **10 BDT penalty** to their running balance at midnight.

## 🧠 Challenge Rules & Logic
- **Daily Target:** 5 unique `OK` (Accepted) problems per day (12:00 AM – 11:59 PM BDT).
- **Dynamic Rating Threshold:** A problem only counts if its rating meets the user's dynamic floor:
  `Min Rating = max(800, Rating - 100 - (Rating % 100))`
  *(Unrated live contest problems also count as valid).*
- **Anti-Cheese Protection:** Resubmitting a problem that was already solved prior to yesterday/today is ignored.
- **Silent Settlement:** At 12:00 AM, the ledger updates automatically. Discord is only pinged if someone misses the goal and incurs a penalty.

---

## 📂 Project Structure
- [`apps-script/Code.gs`](./apps-script/Code.gs) — Google Apps Script backend, Codeforces API integration, automated triggers, and live HTML Leaderboard Web App.
- `discord-bot/` *(Coming Next)* — Serverless Cloudflare Worker handling Discord Slash Commands (`/status`, `/penalty`, `/adduser`).

---

## 🛠️ Manual Cloud Setup Guide

### 1. Google Sheet Setup
Create a Google Sheet with two tabs:
1. **`Dashboard`** (Row 1 Headers):
   `Profile Link or Handle` | `Start-of-Day Rating` | `Solved Yesterday` | `Solved Today` | `Total Penalty (BDT)`
2. **`History`** (Row 1 Headers):
   `Date` | `Handle` | `Evaluated Rating` | `Min Problem Rating` | `Solved` | `Penalty Added (BDT)` | `New Total Penalty (BDT)`

### 2. Google Apps Script & Triggers
1. In the Google Sheet, go to **Extensions → Apps Script** and paste the contents of [`apps-script/Code.gs`](./apps-script/Code.gs).
2. Open **Project Settings (⚙️)** and set the **Time zone** to `(GMT+06:00) Bangladesh Standard Time - Dhaka`.
3. Open **Triggers (⏰)** and add three triggers (all set to `Deployment: Head`):
   - **Auto-Sync on Sheet Open:** `syncAllExistingRowsNow` → `From spreadsheet` → `On open`.
   - **Auto-Fill New Handle:** `onSheetEdit` → `From spreadsheet` → `On edit`.
   - **Midnight Settlement:** `runDailyCodeforcesCheck` → `Time-driven` → `Day timer` → `Midnight to 1am`.
4. Click **Deploy → New deployment → Web app** (`Execute as: Me`, `Who has access: Anyone`) to generate the live Leaderboard Web App URL.
