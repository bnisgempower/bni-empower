# BNI Empower — Member Hub

The chapter's home-base app. Every member opens this to:

- **Submit their weekly check-in** (links to the existing Check-In app)
- See **who's presenting next week** — Core Value, Network Education, Featured Presentation (pulled live from the **Roster** sheet)
- Open **chapter resources** — Google Drive folders, Zoom recordings, docs
- Reach **quick links** — power teams, bio sheets, WhatsApp, etc.

It's a single-file PWA (`index.html`), same dark red/gold design as the Check-In app, installable to the home screen.

---

## 1 — Fill in the config (required)

Open `index.html` and edit the **CONFIG** block near the bottom (`<script>` section). Everything you need to change lives between the two `CONFIG` banners.

| Setting | What to put |
|---|---|
| `APPS_SCRIPT_URL` | Already set — same Web App URL as the Check-In app. Change only if you redeploy to a new URL. |
| `CHECKIN_URL` | The live URL of your **Weekly Check-In** app (its Netlify address). |
| `ONBOARD_URL` | Live URL of `onboarding.html` (new-member form). Set to `''` to hide the footer link. |
| `RESOURCES[]` | Your Google Drive folder/doc links. Replace each `url`. Add/delete rows freely. |
| `QUICK_LINKS[]` | Internal tools & handy pages. Any row whose `url` is still a placeholder is hidden automatically. |

> **How to get a Drive folder link:** open the folder in Google Drive → **Share** → set to *Anyone with the link → Viewer* → **Copy link**. Paste that as the `url`.

**Placeholder behaviour:** any URL still starting with `REPLACE` (or ending `_LEAVE`) is treated as unset —
resources show greyed-out with a reminder, quick links are hidden, and the check-in button warns you. So the app never shows a broken link.

Icons available for `RESOURCES`/`QUICK_LINKS` `icon` field:
`folder`, `video`, `slides`, `sheet`, `team`, `id`, `checkin`, `chat`, `star`, `mic`, `heart`, `link`.

---

## 2 — Deploy the hub

Same as the Check-In app — this is a static folder.

**Drag & drop (fastest)**
1. Go to [app.netlify.com/drop](https://app.netlify.com/drop)
2. Drag this **Empower Hub** folder onto the page
3. Netlify gives you a live URL — share it with the chapter as their main app

**Netlify CLI**
```bash
netlify deploy --prod --dir .
```

Deploy this as a **separate** site from the Check-In app (the hub links *to* the check-in app via `CHECKIN_URL`).

---

## 3 — Backend changes (one-time, in Apps Script)

The hub needs two additions that are **already written** into `BNI Empower Appsscript.js` (in the *Weekly Check In* folder). You just need to push them live:

1. Open your Google Sheet → **Extensions → Apps Script**.
2. Paste the updated `BNI Empower Appsscript.js` (overwrite the old code) and **Save**.
3. **Redeploy:** Deploy → Manage deployments → Edit (✏️) → New version → **Deploy**. (The `/exec` URL stays the same.)
4. Reload the sheet, then from the **BNI Empower** menu run **🎞️ Setup Slide Map Sheet** once (see below).

### What was added

**a) `getNextPresenters` endpoint**
Reads the **Roster** sheet (`Meeting Date | Meeting Type | Network Education | Core Value | Featured Presentation`), finds the next meeting date that is today or later, and returns that meeting's three presenters + a few upcoming weeks. The hub calls `?action=getNextPresenters`.
→ **To change who's presenting, just edit the Roster sheet.** The hub updates automatically.

**b) Sheet-driven slide map** (the "spreadsheet mapped to the boxes in the slides")
The Support-Team box positions used by the auto-slide feature used to be hardcoded. They're now editable in two sheet tabs:

- **`Slide_Map`** — one row per Support-Team member: which slide boxes (shape + photo) they occupy and the exact position/scale.
- **`Intro_Slides`** — each member's 30-second intro slide object ID.

Run **BNI Empower → 🎞️ Setup Slide Map Sheet** to create these tabs, pre-filled with the current values. Edit them to move a box or re-point a member — no code needed. If a tab is empty or missing, the script safely falls back to the built-in defaults, so the slide automation keeps working either way.

---

## How the automatic slides work (already built)

Members submit attendance → every **Sunday 9pm** the `weeklySundayTrigger` runs:

1. `restoreAllSlides()` — resets the Support Team slide to its full layout
2. `updateSlidesFromAttendance()` —
   - **hides absent** Support-Team members from Slide 15 and **re-arranges** the remaining photos into a clean grid (≤4 → one row, >4 → two rows)
   - moves a **truly-absent** member's 30-sec intro slide to the end (kept in place if they sent a substitute)

You can also run these manually from the **BNI Empower** menu:
🖼️ *Update Slides from Attendance* · ↩️ *Restore All Slides*.

---

## Files

| File | Purpose |
|---|---|
| `index.html` | The hub — deploy this |
| `manifest.json` | PWA install metadata |
| `icon.svg` | Home-screen / tab icon |
| `sw.js` | Service worker (offline shell + always-fresh roster) |
| `README.md` | This file |

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| Next Meeting shows "No upcoming meeting" | Add a future-dated row to the **Roster** sheet (date today or later). |
| Presenter names show "To be confirmed" | Those Roster cells are blank — fill in the NE / CV / FP names. |
| "Couldn't load the roster" | Confirm the Apps Script was **redeployed** after pasting, and `APPS_SCRIPT_URL` matches the `/exec` URL. |
| Check-in button greyed out | `CHECKIN_URL` is still a placeholder — set it to your live check-in URL. |
| A resource tile is greyed out | Its `url` is still `REPLACE_…` — paste the real Drive link. |
| Slide automation stopped after edits | Check the `Slide_Map` / `Intro_Slides` tabs for typos; clear a bad row and it falls back to defaults. |
