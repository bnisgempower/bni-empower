# BNI Empower — Flyer Generator

Generates member one-pager flyers (like the Xool Aircon sample) from a Google
Slides template. Quick input via `flyer.html` → ready-to-send PNG + PDF, with
full edit history in the chapter Sheet.

## Architecture
- **Template deck** (design master, never modified): one slide, the flyer
  layout, each dynamic text in its own text box containing its `{{TOKEN}}`.
- **`Flyers` tab** in the existing chapter Sheet: one row per member = the
  history. Re-editing = pick member, tweak, regenerate.
- **Per-member deck**: first generate copies the template to
  `Flyer — <Name>` in the output folder. Photos / logo / QR are placed
  **manually once** in that copy — regenerating only rewrites the text boxes,
  so images are never touched.
- **Exports**: every generate saves `Flyer — <Name>.png` (WhatsApp) and
  `.pdf` (print) to the output folder and writes the links into the sheet.

## Tokens (each in its own text box on the template)
`{{NAME}}` `{{COMPANY}}` `{{HP}}` `{{BRAND}}` `{{TAGLINE}}` `{{HEADING}}`
`{{ABOUT}}` `{{WHY_US}}` `{{SERVICES}}`

Multi-line boxes: `{{ABOUT}}` (paragraphs), `{{WHY_US}}` / `{{SERVICES}}`
(pre-styled as bulleted lists — one bullet per line of input).
Set text boxes to **shrink text on overflow** so long content never breaks
the layout.

## Setup (one-time)
1. Create the template deck in Google Slides (single slide, page size set to
   the flyer ratio, e.g. A4 portrait). Put the tokens above in their own text
   boxes, styled exactly as the final flyer. Placeholder images go where the
   headshot / photos / QR sit (purely visual — script ignores them).
2. Create a Drive folder for outputs.
3. Paste both IDs into `FLYER_TEMPLATE_ID` / `FLYER_FOLDER_ID` in
   `Flyer Generator Appsscript.js`.
4. New Apps Script project (separate from Weekly Check-In) → paste the file →
   Deploy as Web App (Execute as: Me / Access: Anyone).
5. Paste the `/exec` URL into `APPS_SCRIPT_URL` in `flyer.html`.
6. Verify the template: open
   `<exec-url>?action=checkTemplate&pin=<PIN>` — it lists found/missing tokens.

## Weekly use
1. Open `flyer.html` → pick a member (✓ = existing flyer loads for editing;
   new members prefill from Active_Members).
2. Edit fields → PIN → **Save & Generate**.
3. First time only: open the member's deck via the "Edit deck" link and drop
   in their headshot / photos / QR manually, then hit Save & Generate again to
   re-export.
