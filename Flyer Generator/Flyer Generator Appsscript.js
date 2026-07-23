/**
 * BNI Empower — Member Flyer Generator
 * Google Apps Script backend (SEPARATE project from the Weekly Check-In script)
 *
 * Deploy: script.google.com → New project → paste this file → Deploy as Web App
 *   Execute as:  Me
 *   Who has access: Anyone
 * Then paste the /exec URL into flyer.html (APPS_SCRIPT_URL).
 *
 * HOW IT WORKS
 *   • FLYER_TEMPLATE_ID is a one-slide deck: the flyer design with each dynamic
 *     text as its OWN text box, containing its {{TOKEN}} (see FIELDS below).
 *     The template is never modified — it's the permanent design master.
 *   • Each member gets a private copy of the template ("Flyer — <Name>") in
 *     FLYER_FOLDER_ID. A Drive copy keeps the template's element objectIds, so
 *     we find each box by locating its {{TOKEN}} in the template, then write
 *     the member's text into the same objectId in their copy (insert-then-
 *     delete, so the template styling is preserved).
 *   • Regenerating rewrites ONLY those text boxes — photos/logo/QR placed
 *     manually in a member's copy are never touched.
 *   • Every generate exports the slide as PDF + PNG into the folder and stores
 *     the links in the Flyers sheet (the history: re-edit = pick member, tweak,
 *     regenerate).
 */
// ── CONFIG ────────────────────────────────────────────────────────────────────
const SHEET_ID          = '1cmBnBAiOWPB71oXmJEUMzCIam7RffTb0INPVM4TNhGw'; // same chapter sheet
const FLYER_TEMPLATE_ID = '1Q5-nTuMyRfAmK-RJ__VYnyK_bV17I7lMHgaG6Uw0vig'; // one-slide flyer template deck
const FLYER_FOLDER_ID   = '1tonRzwjUEtc8eNy8klZClURa9Ykfpquv';           // "One Pager Intro" Drive folder
const ADMIN_PIN         = '0303';

const SN_FLYERS         = 'Flyers';
const SN_ACTIVE_MEMBERS = 'Active_Members';

// Field key ↔ template token ↔ Flyers sheet column (order = sheet order,
// after the leading 'Member' key column). Each token must live in its OWN
// text box on the template slide.
const FIELDS = [
  { key: 'displayName', token: '{{NAME}}',    col: 'Display Name' },  // RAYMOND YEOH
  { key: 'company',    token: '{{COMPANY}}',  col: 'Company' },       // XOOL INDUSTRIES PTE LTD
  { key: 'hp',         token: '{{HP}}',       col: 'HP' },            // HP: 9685 1011
  { key: 'brand',      token: '{{BRAND}}',    col: 'Brand' },         // Xool Aircon Experts
  { key: 'tagline',    token: '{{TAGLINE}}',  col: 'Tagline' },       // No. 1 Aircon Service Provider…
  { key: 'heading',    token: '{{HEADING}}',  col: 'Heading' },       // SERVICE YOU CAN TRUST
  { key: 'about',      token: '{{ABOUT}}',    col: 'About' },         // paragraphs (blank line between)
  { key: 'whyUs',      token: '{{WHY_US}}',   col: 'Why Us' },        // one bullet per line
  { key: 'services',   token: '{{SERVICES}}', col: 'Services' },      // one bullet per line
];
// Non-field columns appended after the FIELDS columns:
const EXTRA_COLS = ['Deck ID', 'Deck URL', 'PDF URL', 'PNG URL', 'Updated'];

// ── doGet ─────────────────────────────────────────────────────────────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  try {
    if (action === 'getMembers')    return jsonOk_({ members: getMembers_() });
    if (action === 'getFlyer')      return jsonOk_({ flyer: getFlyerRow_(e.parameter.member || '') });
    if (action === 'checkTemplate') return (e.parameter.pin === ADMIN_PIN)
                                      ? jsonOk_(checkTemplate_())
                                      : jsonErr_('Invalid PIN');
    if (action === 'setupFlyers')   return (e.parameter.pin === ADMIN_PIN)
                                      ? jsonOk_({ ready: !!getFlyersSheet_() })
                                      : jsonErr_('Invalid PIN');
  } catch (err) {
    logError_('doGet:' + action, err);
    return jsonErr_(err.message);
  }
  return jsonOk_({ app: 'BNI Empower Flyer Generator', timestamp: new Date().toISOString() });
}

// ── doPost — save + generate ──────────────────────────────────────────────────
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('No POST data received');
    const data = JSON.parse(e.postData.contents);

    if (data.formType === 'saveFlyer') {
      if (data.pin !== ADMIN_PIN) return jsonErr_('Invalid PIN');
      if (!String(data.memberName || '').trim()) return jsonErr_('Member name is required');
      const row    = upsertFlyerRow_(data);
      const result = generateFlyer_(row);
      return jsonOk_(result);
    }

    return jsonErr_('Unknown formType');
  } catch (err) {
    logError_('doPost', err);
    return jsonErr_(err.message);
  }
}

// ── Members + prefill (from Active_Members) ───────────────────────────────────
// Returns every active member with prefill data pulled from Active_Members,
// plus whether they already have a saved flyer row.
function getMembers_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const am = ss.getSheetByName(SN_ACTIVE_MEMBERS);
  if (!am || am.getLastRow() < 2) return [];

  const haveFlyer = new Set(getFlyerNames_().map(n => n.toLowerCase()));

  // A..AL (38 cols): B name, C first, D surname, J trade, K business name,
  // L tag line, N paragraph, O products, AG mobile
  const rows = am.getRange(2, 1, am.getLastRow() - 1, 38).getValues();
  return rows.map(r => {
    const first   = String(r[2]).trim();
    const surname = String(r[3]).trim();
    const name    = (first + (surname ? ' ' + surname : '')).trim() || String(r[1]).trim();
    if (!name) return null;
    return {
      name,
      hasFlyer: haveFlyer.has(name.toLowerCase()),
      prefill: {
        company:  String(r[10]).trim(),
        hp:       String(r[32]).trim() ? 'HP: ' + String(r[32]).trim() : '',
        brand:    String(r[10]).trim(),
        tagline:  String(r[11]).trim(),
        about:    String(r[13]).trim(),
        services: String(r[14]).trim(),
        trade:    String(r[9]).trim(),
      },
    };
  }).filter(Boolean).sort((a, b) => a.name.localeCompare(b.name));
}

// ── Flyers sheet ──────────────────────────────────────────────────────────────
// Col A = 'Member' (roster key, row identity). Then the FIELDS cols, then EXTRA_COLS.
function flyerHeaders_() { return ['Member'].concat(FIELDS.map(f => f.col)).concat(EXTRA_COLS); }

function getFlyersSheet_() {
  const ss  = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SN_FLYERS);
  if (!sheet) {
    sheet = ss.insertSheet(SN_FLYERS);
    const headers = flyerHeaders_();
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground('#1a5276').setFontColor('#ffffff')
      .setFontWeight('bold').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getFlyerNames_() {
  const sheet = getFlyersSheet_();
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(r => String(r[0]).trim()).filter(Boolean);
}

// Row index (1-based) for a member, or -1.
function findFlyerRow_(member) {
  const names = getFlyerNames_();
  const idx   = names.findIndex(n => n.toLowerCase() === member.trim().toLowerCase());
  return idx === -1 ? -1 : idx + 2;
}

function getFlyerRow_(member) {
  const rowIdx = findFlyerRow_(member);
  if (rowIdx === -1) return null;
  const sheet  = getFlyersSheet_();
  const vals   = sheet.getRange(rowIdx, 1, 1, flyerHeaders_().length).getValues()[0];
  const out    = { memberName: String(vals[0] || '') };
  FIELDS.forEach((f, i) => { out[f.key] = String(vals[i + 1] == null ? '' : vals[i + 1]); });
  const base = 1 + FIELDS.length;
  out.deckId  = String(vals[base]     || '');
  out.deckUrl = String(vals[base + 1] || '');
  out.pdfUrl  = String(vals[base + 2] || '');
  out.pngUrl  = String(vals[base + 3] || '');
  out.updated = String(vals[base + 4] || '');
  return out;
}

// Writes/updates the member's field values; returns the full row object.
function upsertFlyerRow_(data) {
  const sheet  = getFlyersSheet_();
  const clean  = v => String(v == null ? '' : v).replace(/\r/g, '').trim();
  const member = clean(data.memberName);

  let rowIdx = findFlyerRow_(member);
  if (rowIdx === -1) {
    rowIdx = sheet.getLastRow() + 1;
    sheet.getRange(rowIdx, 1, 1, flyerHeaders_().length).setValues([new Array(flyerHeaders_().length).fill('')]);
  }
  const fieldVals = FIELDS.map(f =>
    f.key === 'displayName' ? (clean(data.displayName) || member.toUpperCase()) : clean(data[f.key]));
  sheet.getRange(rowIdx, 1, 1, 1 + FIELDS.length).setValues([[member].concat(fieldVals)]);

  const tz = Session.getScriptTimeZone();
  sheet.getRange(rowIdx, flyerHeaders_().length).setValue(
    Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'));

  const row = getFlyerRow_(member);
  row._rowIdx = rowIdx;
  return row;
}

// ── Template slot discovery ───────────────────────────────────────────────────
// Reads the TEMPLATE deck once per run and finds which text box holds each
// {{TOKEN}}. Copies keep the same objectIds, so this map works on every
// member's deck too.
function templateSlotMap_() {
  const boxes = readBoxTexts_(FLYER_TEMPLATE_ID);
  const map   = {};
  FIELDS.forEach(f => {
    for (const [id, text] of Object.entries(boxes)) {
      if (text.indexOf(f.token) !== -1) { map[f.key] = id; break; }
    }
  });
  return map;
}

// Health check for the template: which tokens were found / are missing.
function checkTemplate_() {
  const map     = templateSlotMap_();
  const found   = FIELDS.filter(f => map[f.key]).map(f => f.token);
  const missing = FIELDS.filter(f => !map[f.key]).map(f => f.token);
  return { found, missing, ok: missing.length === 0 };
}

// { objectId: text } for every shape text box on a deck (one API call).
function readBoxTexts_(presId) {
  const url  = 'https://slides.googleapis.com/v1/presentations/' + presId +
               '?fields=' + encodeURIComponent('slides(objectId,pageElements(objectId,shape(text(textElements(textRun(content))))))');
  const resp = UrlFetchApp.fetch(url, { headers: authHeader_(), muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('Slides read failed: ' + resp.getContentText().slice(0, 200));
  const map = {};
  (JSON.parse(resp.getContentText()).slides || []).forEach(s => {
    (s.pageElements || []).forEach(pe => {
      if (pe.shape && pe.shape.text && pe.shape.text.textElements) {
        map[pe.objectId] = pe.shape.text.textElements.map(te => (te.textRun ? te.textRun.content : '')).join('');
      }
    });
  });
  return map;
}

// ── Generate ──────────────────────────────────────────────────────────────────
function generateFlyer_(row) {
  const member = row.memberName;
  const folder = DriveApp.getFolderById(FLYER_FOLDER_ID);
  const sheet  = getFlyersSheet_();
  const rowIdx = row._rowIdx || findFlyerRow_(member);
  const base   = 1 + FIELDS.length; // 0-based col offset of EXTRA_COLS ('Member' + fields)

  // 1. Ensure the member has their own copy of the template
  let deckId = row.deckId;
  let deckOk = false;
  if (deckId) { try { DriveApp.getFileById(deckId); deckOk = true; } catch (_) {} }
  if (!deckOk) {
    const copy = DriveApp.getFileById(FLYER_TEMPLATE_ID).makeCopy('Flyer — ' + member, folder);
    deckId = copy.getId();
    sheet.getRange(rowIdx, base + 1).setValue(deckId);
    sheet.getRange(rowIdx, base + 2).setValue('https://docs.google.com/presentation/d/' + deckId + '/edit');
  }

  // 2. Rewrite each field's text box in place (styling preserved; images untouched)
  const slots    = templateSlotMap_();
  const current  = readBoxTexts_(deckId);
  const requests = [];
  FIELDS.forEach(f => {
    const box = slots[f.key];
    if (!box) return;                                   // token missing from template — skip
    const neu = String(row[f.key] || '').trim() || ' '; // never write an empty string
    const old = String(current[box] || '').replace(/\n+$/, '');
    if (old === neu) return;
    requests.push({ insertText: { objectId: box, insertionIndex: 0, text: neu } });
    if (old.length) {
      requests.push({ deleteText: { objectId: box, textRange: {
        type: 'FIXED_RANGE', startIndex: neu.length, endIndex: neu.length + old.length,
      } } });
    }
  });
  if (requests.length) {
    const resp = UrlFetchApp.fetch(
      'https://slides.googleapis.com/v1/presentations/' + deckId + ':batchUpdate', {
        method: 'post', contentType: 'application/json', headers: authHeader_(),
        payload: JSON.stringify({ requests }), muteHttpExceptions: true,
      });
    if (resp.getResponseCode() !== 200) throw new Error('Slides update failed: ' + resp.getContentText().slice(0, 300));
  }

  // 3. Export PDF + PNG (replacing last week's files of the same name)
  const pdfUrl = exportPdf_(deckId, 'Flyer — ' + member + '.pdf', folder);
  const pngUrl = exportPng_(deckId, 'Flyer — ' + member + '.png', folder);
  sheet.getRange(rowIdx, base + 3).setValue(pdfUrl);
  sheet.getRange(rowIdx, base + 4).setValue(pngUrl);

  return {
    generated: true,
    member,
    boxesChanged: requests.length ? Math.ceil(requests.length / 2) : 0,
    deckUrl: 'https://docs.google.com/presentation/d/' + deckId + '/edit',
    pdfUrl, pngUrl,
  };
}

function exportPdf_(deckId, fileName, folder) {
  const resp = UrlFetchApp.fetch(
    'https://docs.google.com/presentation/d/' + deckId + '/export/pdf',
    { headers: authHeader_(), muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error('PDF export failed (' + resp.getResponseCode() + ')');
  return saveExport_(resp.getBlob().setName(fileName), fileName, folder);
}

function exportPng_(deckId, fileName, folder) {
  // First (only) slide's pageId, then the 1600px thumbnail
  const meta = UrlFetchApp.fetch(
    'https://slides.googleapis.com/v1/presentations/' + deckId + '?fields=slides.objectId',
    { headers: authHeader_(), muteHttpExceptions: true });
  if (meta.getResponseCode() !== 200) throw new Error('PNG export failed reading deck');
  const pageId = (JSON.parse(meta.getContentText()).slides || [])[0].objectId;

  const thumb = UrlFetchApp.fetch(
    'https://slides.googleapis.com/v1/presentations/' + deckId + '/pages/' + pageId +
    '/thumbnail?thumbnailProperties.thumbnailSize=LARGE',
    { headers: authHeader_(), muteHttpExceptions: true });
  if (thumb.getResponseCode() !== 200) throw new Error('PNG export failed (' + thumb.getResponseCode() + ')');

  const img = UrlFetchApp.fetch(JSON.parse(thumb.getContentText()).contentUrl);
  return saveExport_(img.getBlob().setName(fileName), fileName, folder);
}

// Trash any previous export with this name, save the new one, return its link.
function saveExport_(blob, fileName, folder) {
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?id=' + file.getId();
}

// ── Utility ───────────────────────────────────────────────────────────────────
function authHeader_() { return { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }; }

function jsonOk_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', ...payload }))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr_(message) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message }))
    .setMimeType(ContentService.MimeType.JSON);
}

function logError_(ctx, err) {
  console.error('[Flyer %s] %s\n%s', ctx, err.message, err.stack || '');
}
