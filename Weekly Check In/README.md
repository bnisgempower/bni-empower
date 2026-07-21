# BNI Empower — Weekly Check-In System

Self-contained check-in form for a 27-member BNI chapter.  
Members submit every Sunday by 8 pm; data flows automatically into Google Sheets.

---

## Files

| File | Purpose |
|------|---------|
| `index.html` | Complete mobile-first wizard — deploy this to Netlify |
| `appsscript.js` | Google Apps Script backend — paste into Apps Script editor |
| `README.md` | This file |

---

## Setup (one-time, ~10 minutes)

### 1 — Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a **new blank spreadsheet**.
2. Name it something like **BNI Empower Attendance**.
3. Copy the **Sheet ID** from the URL:
   ```
   https://docs.google.com/spreadsheets/d/THIS_IS_YOUR_SHEET_ID/edit
   ```

---

### 2 — Set up Apps Script

1. Inside the spreadsheet: **Extensions → Apps Script**.
2. Delete the default `Code.gs` content.
3. Paste the entire contents of **`appsscript.js`**.
4. On line 14, replace the placeholder:
   ```js
   const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';
   // → becomes →
   const SHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms';  // example
   ```
5. Click **Save** (Ctrl+S / Cmd+S).

---

### 3 — Bootstrap the sheets

1. In the Apps Script editor, select the function **`setupSheets`** from the dropdown (top-left, next to the Run ▶ button).
2. Click **Run ▶**.
3. Approve any permission prompts Google shows.
4. Check your spreadsheet — it should now have **7 tabs**:
   - `Members` (pre-seeded with all 27 member names)
   - `Weekly_Submissions`
   - `Attendance_View`
   - `Visitor_Log`
   - `Referrals`
   - `Core_Value_Sharing`
   - `Actual_Attendance`

---

### 4 — Deploy the Apps Script as a Web App

1. In Apps Script: click **Deploy → New deployment**.
2. Click the ⚙️ gear icon next to "Select type" and choose **Web app**.
3. Fill in:
   - **Description**: `BNI Empower Check-In v1`
   - **Execute as**: `Me`
   - **Who has access**: `Anyone`
4. Click **Deploy**.
5. Copy the **Web app URL** — it looks like:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```

> **Redeploy after code changes**: Every time you update `appsscript.js`, go to  
> **Deploy → Manage deployments → Edit → New version → Deploy**.  
> The URL stays the same.

---

### 5 — Wire the URL into index.html

Open `index.html` and find line ~210:

```js
const APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_URL_HERE';
```

Replace the placeholder with your Web app URL:

```js
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycb.../exec';
```

Save the file.

---

### 6 — Deploy to Netlify

**Option A — Drag & drop (fastest)**

1. Go to [app.netlify.com/drop](https://app.netlify.com/drop).
2. Drag your `index.html` file (or the whole folder) onto the page.
3. Netlify gives you a live URL in seconds — share it with your chapter.

**Option B — Netlify CLI**

```bash
npm install -g netlify-cli
netlify deploy --prod --dir . --filter index.html
```

---

## Google Sheets — Tab Reference

| Tab | Written by | Notes |
|-----|-----------|-------|
| `Members` | `setupSheets()` one-time | Admin can edit trade/company/phone/email columns |
| `Weekly_Submissions` | Apps Script on every submit | Raw form data; do not manually edit |
| `Attendance_View` | Formula-driven (manual setup) | See formula hint below |
| `Visitor_Log` | Apps Script when visitor = yes | MC marks "Attended Day-Of" column |
| `Referrals` | Manual by MC | Date, From, To, Type, Detail, Value SGD |
| `Core_Value_Sharing` | Manual by MC | Week Of, Member, Topic |
| `Actual_Attendance` | Manual by MC day-of | Y / Late / Absent |

### Attendance_View — formula setup

After a few submissions exist, build the grid:

1. Column A: paste the 27 member names (copy from `Members` tab, column A).
2. Row 1, columns B onward: enter the week dates (e.g., `14 Apr 2026`, `21 Apr 2026`, …).
3. In **B2** enter this formula, then drag right and down:

```
=IFERROR(
  LET(
    wk, B$1,
    mem, $A2,
    row, FILTER(Weekly_Submissions!A:P,
                (Weekly_Submissions!B:B=wk)*
                (Weekly_Submissions!C:C=mem)),
    att, INDEX(row,1,4),
    sub, INDEX(row,1,6),
    IF(att="yes","Present",IF(sub<>"","Sub","Missing"))
  ),
  "Missing"
)
```

4. Apply conditional formatting:
   - `Present` → green fill
   - `Sub`     → yellow fill
   - `Missing` → red fill

---

## Customisation

| Change | Where |
|--------|-------|
| Add / remove members | Edit `MEMBERS` array in `index.html` **and** `ROSTER` in `appsscript.js` |
| Change meeting day (week-of calculation) | Edit `calcWeekOf()` in `index.html` |
| Tweak colours | Edit CSS `:root` variables at the top of `index.html` |
| Add form fields | Add inputs in `index.html`, expand `payload` object, add columns to `HEADERS.SUBMISSIONS` in `appsscript.js`, then redeploy both |

---

## How the no-CORS submit works

The form POSTs JSON as `Content-Type: text/plain` with `mode: 'no-cors'`.  
This avoids a CORS preflight. The browser sends the request and receives an **opaque** response (unreadable), so the app optimistically shows a success screen.  

If you need server-side error feedback (e.g., sheet quota exceeded), switch to a regular `fetch` without `no-cors` — Apps Script already returns JSON `{ status: "ok" }` on success.

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| Data not appearing in sheet | Re-check `SHEET_ID` and redeploy Apps Script after any change |
| "Script function not found: doPost" | Make sure you saved **and redeployed** after pasting the code |
| Form submits but shows network error | Check that the URL in `index.html` is the deployed `/exec` URL, not the `/dev` URL |
| Members tab empty | Run `setupSheets()` manually from the Apps Script editor |
| Need to test the endpoint | Open the Web app URL in a browser — `doGet` returns a JSON health check |
