/**
 * BNI Empower — Weekly Check-In
 * Google Apps Script backend
 *
 * Deploy: Extensions → Apps Script → paste this file → Deploy as Web App
 *   Execute as:  Me
 *   Who has access: Anyone
 *
 * Set SHEET_ID below to your Google Sheet's ID (from its URL).
 */
// ── CONFIG ────────────────────────────────────────────────────────────────────
const SHEET_ID          = '1cmBnBAiOWPB71oXmJEUMzCIam7RffTb0INPVM4TNhGw';
const PRESENTATION_ID   = '1inhHToZYqXes1zjUT6T8czAVOqU4jXQW7HghDBadpX0';
const HEADSHOTS_FOLDER_ID = '1Cb2JJpoTOttma0iKKX-1-TY5hyU_N9sW';
const LOGOS_FOLDER_ID     = '1R2jzxXnvlfcYUnYVw_TzmbrPhgS9NZbq';


// VP who receives the missing submissions report every Sunday 8pm
const VP_EMAIL = 'weekhai.pang@sg-alliance.com';
const VP_NAME  = 'Wee Khai';

// President — cc'd on the weekly missing submissions email
const PRESIDENT_EMAIL = 'contact.iskons@gmail.com';

// ── SHEET NAMES ───────────────────────────────────────────────────────────────
const SN = {
  MEMBERS:           'Members',
  ACTIVE_MEMBERS:    'Active_Members',
  SUBMISSIONS:       'Weekly_Submissions',
  VISITOR_LOG:       'Visitor_Log',
  CORE_VALUE:        'Core_Value_Sharing',
  ACTUAL_ATTENDANCE: 'Actual_Attendance',
  NETWORK_EDU:       'Network_Education',
  FEATURED_PRES:     'Featured_Presentation',
  DASHBOARD:         'Dashboard',
  ROSTER:            'Roster',
  INTRO_SLIDES:      'Intro_Slides',   // 30-sec intro slide objectIds per member
  COMMITTEE_REPORT:  'Committee_Report', // VP weekly numbers + running totals
  ROLES:             'Roles',          // President-managed: Team | Role | Member | Trade
};

// ROSTER is read live from Active_Members cols C+D (First Name + Surname)
// To add/remove members, edit the Active_Members sheet directly.
const EXCLUDED_MEMBERS = [
  'Cheryl Lee',
  'Timmy San',
  'Sean Wee',
];

let _rosterCache = null;
function getRoster_() {
  if (_rosterCache) return _rosterCache;
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.ACTIVE_MEMBERS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const excluded = new Set(EXCLUDED_MEMBERS.map(n => n.toLowerCase()));
  _rosterCache = sheet.getRange(2, 3, sheet.getLastRow() - 1, 2) // cols C & D
    .getValues()
    .map(r => {
      const first   = String(r[0]).trim();
      const surname = String(r[1]).trim();
      return surname ? first + ' ' + surname : first;
    })
    .filter(n => n && !excluded.has(n.toLowerCase()))
    .sort();
  return _rosterCache;
}

// ── doGet — health check / attendance status fetch / admin actions ────────────
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'getAttendance')        return getAttendanceStatus_(e.parameter.weekOf || '');
  if (action === 'getMissing')           return adminGetMissing_();
  if (action === 'getAttendanceList')    return adminGetAttendanceList_();
  if (action === 'clearSubmissions')     return adminClearSubmissions_(e.parameter.pin || '');
  if (action === 'highlightDuplicates')  return adminHighlightDuplicates_(e.parameter.pin || '');
  if (action === 'getPowerTeams')        return getPowerTeams_();
  if (action === 'getRoster')            return jsonOk_({ members: getRoster_() });
  if (action === 'getNextPresenters')    return getNextPresenters_();
  if (action === 'getCommitteeTotals')   return getCommitteeTotalsAction_();
  if (action === 'resetNextPresenters')  return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_(resetAllPresent_())
                                            : jsonErr_('Invalid PIN');
  if (action === 'dumpSlides')           return (e.parameter.pin === ADMIN_PIN)
                                            ? dumpSlides_(e.parameter.find || '', e.parameter.slide || '', e.parameter.raw || '')
                                            : jsonErr_('Invalid PIN');
  if (action === 'testCommittee')        return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_({ pairs: reflowGridTeam_(COMMITTEE, parseAbsent_(e.parameter.absent || '')) })
                                            : jsonErr_('Invalid PIN');
  if (action === 'setupRoles')           return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_(setupRolesSheet())
                                            : jsonErr_('Invalid PIN');
  if (action === 'getRoles')             return jsonOk_({ roles: getRoles_() });
  if (action === 'getCoreValueOptions')  return getCoreValueOptions_();
  if (action === 'testCoreValue')        return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_(saveCoreValue_({ member: e.parameter.member || '', coreValue: e.parameter.value || '' }))
                                            : jsonErr_('Invalid PIN');
  if (action === 'renderCommittee')      return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_(renderCommitteeText_())
                                            : jsonErr_('Invalid PIN');
  if (action === 'renderVisitorHost')    return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_(renderVisitorHostText_())
                                            : jsonErr_('Invalid PIN');
  if (action === 'renderTeams')          return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_(renderAllGridText_())
                                            : jsonErr_('Invalid PIN');
  if (action === 'checkHeadshots')       return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_(checkHeadshots_(e.parameter.team || ''))
                                            : jsonErr_('Invalid PIN');
  if (action === 'updateTitleDate')      return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_(updateTitleDate_())
                                            : jsonErr_('Invalid PIN');
  if (action === 'rebuildSupportPhotos') return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_(rebuildSupportPhotos_())
                                            : jsonErr_('Invalid PIN');
  if (action === 'relayoutSupport')      return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_(relayoutSupport_(e.parameter.recrop === '1'))
                                            : jsonErr_('Invalid PIN');
  if (action === 'rebuildSupportSlides') return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_(rebuildSupportSlides_())
                                            : jsonErr_('Invalid PIN');
  if (action === 'testAbsent')           return (e.parameter.pin === ADMIN_PIN)
                                            ? jsonOk_(applySupportAttendance_(
                                                String(e.parameter.absent || '').split(',').map(s => s.trim()).filter(Boolean)))
                                            : jsonErr_('Invalid PIN');

  const response = {
    status:    'ok',
    app:       'BNI Empower Weekly Check-In',
    timestamp: new Date().toISOString(),
    sheets:    Object.values(SN),
  };
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Admin GET actions ─────────────────────────────────────────────────────────
const ADMIN_PIN = '0303';

function adminGetMissing_() {
  const missing        = getMissingSubmitters_();
  const missingSet     = new Set(missing.map(n => n.toLowerCase()));
  const submittedNames = getRoster_().filter(n => !missingSet.has(n.toLowerCase()));
  return jsonOk_({ missing, submittedNames, total: getRoster_().length, submitted: submittedNames.length });
}

function adminGetAttendanceList_() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.ACTUAL_ATTENDANCE);
  const tz    = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy');

  const lastCol = sheet.getLastColumn();
  let dateCol = -1, dateLabel = '';
  if (lastCol >= 4) {
    const headers = sheet.getRange(5, 4, 1, lastCol - 3).getValues()[0];
    headers.forEach((v, i) => {
      const lbl = (v instanceof Date) ? Utilities.formatDate(v, tz, 'dd MMM yyyy') : String(v).trim();
      if (lbl === today) { dateCol = 4 + i; dateLabel = lbl; }
    });
  }

  const lastRow = sheet.getLastRow();
  const members = [];

  if (dateCol === -1) {
    getRoster_().forEach(name => members.push({ name, status: null }));
    return jsonOk_({ dateLabel: today, members });
  }

  if (lastRow >= 6) {
    const names    = sheet.getRange(6, 2, lastRow - 5, 2).getValues();
    const statuses = sheet.getRange(6, dateCol, lastRow - 5, 1).getValues();
    names.forEach((row, i) => {
      const first   = String(row[0]).trim();
      const surname = String(row[1]).trim();
      if (first) members.push({ name: first + (surname ? ' ' + surname : ''), status: String(statuses[i][0]).trim() || null });
    });
  }
  return jsonOk_({ dateLabel, members });
}

function adminClearSubmissions_(pin) {
  if (pin !== ADMIN_PIN) return jsonErr_('Invalid PIN');
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.SUBMISSIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  return jsonOk_({ cleared: true });
}

function adminHighlightDuplicates_(pin) {
  if (pin !== ADMIN_PIN) return jsonErr_('Invalid PIN');
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.SUBMISSIONS);

  const existing = sheet.getConditionalFormatRules().filter(r => {
    const crit = r.getBooleanCondition();
    return !(crit && crit.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA &&
             crit.getCriteriaValues()[0].toString().includes('COUNTIF'));
  });
  const range = sheet.getRange('A2:N1000');
  const rule  = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=COUNTIF($C:$C,$C2)>1')
    .setBackground('#FFE0B2')
    .setFontColor('#7B3F00')
    .setRanges([range])
    .build();
  existing.push(rule);
  sheet.setConditionalFormatRules(existing);

  const lastRow = sheet.getLastRow();
  let dupeCount = 0;
  if (lastRow > 1) {
    const names = sheet.getRange(2, 3, lastRow - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase()).filter(Boolean);
    const counts = {};
    names.forEach(n => { counts[n] = (counts[n] || 0) + 1; });
    dupeCount = Object.values(counts).filter(c => c > 1).length;
  }
  return jsonOk_({ highlighted: true, duplicateMembers: dupeCount });
}

function getAttendanceStatus_(weekOf) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.ACTUAL_ATTENDANCE);
  const tz    = Session.getScriptTimeZone();

  let dateLabel;
  try {
    const d = new Date(weekOf);
    dateLabel = isNaN(d.getTime()) ? String(weekOf).trim()
      : Utilities.formatDate(d, tz, 'dd MMM yyyy');
  } catch(_) { dateLabel = String(weekOf).trim(); }

  const lastCol = sheet.getLastColumn();
  let dateCol = -1;
  if (lastCol >= 4) {
    sheet.getRange(5, 4, 1, lastCol - 3).getValues()[0].forEach((v, i) => {
      const lbl = (v instanceof Date)
        ? Utilities.formatDate(v, tz, 'dd MMM yyyy')
        : String(v).trim();
      if (lbl === dateLabel) dateCol = 4 + i;
    });
  }

  const members = [];

  if (dateCol === -1) {
    getRoster_().forEach(name => members.push({ name, status: null }));
    return jsonOk_({ dateLabel, members });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow >= 6) {
    const names    = sheet.getRange(6, 2, lastRow - 5, 2).getValues();
    const statuses = sheet.getRange(6, dateCol, lastRow - 5, 1).getValues();
    names.forEach((row, i) => {
      const first   = String(row[0]).trim();
      const surname = String(row[1]).trim();
      const full    = first + (surname ? ' ' + surname : '');
      const status  = String(statuses[i][0]).trim() || null;
      if (first) members.push({ name: full, status });
    });
  }

  return jsonOk_({ dateLabel, members });
}

// ── doPost — receives form submission from index.html ─────────────────────────
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No POST data received');
    }

    const data = JSON.parse(e.postData.contents);

    if (data.formType === 'attendance') {
      writeAttendance_(data);
      return jsonOk_({ written: true, member: data.memberName, status: data.status });
    }

    if (data.formType === 'newMember') {
      writeNewMember_(data);
      return jsonOk_({ written: true, member: data.fullName });
    }

    if (data.formType === 'committeeReport') {
      if (data.pin !== ADMIN_PIN) return jsonErr_('Invalid PIN');
      const totals = writeCommitteeReport_(data);
      return jsonOk_({ written: true, totals, display: committeeDisplay_(totals) });
    }

    if (data.formType === 'saveRoles') {
      if (data.pin !== ADMIN_PIN) return jsonErr_('Invalid PIN');
      return jsonOk_(saveRoles_(data));
    }

    if (data.formType === 'saveCoreValue') {
      if (data.pin !== ADMIN_PIN) return jsonErr_('Invalid PIN');
      return jsonOk_(saveCoreValue_(data));
    }

    writeSubmission_(data);

    if (data.attending === 'no' && !data.subName) {
      const weekOfClean = (data.weekOf || '').replace(/^[A-Za-z]+,\s*/, '');
      writeAttendance_({
        memberName: data.memberName,
        weekOf:     weekOfClean,
        status:     'Absent',
        lateMins:   0,
      });
    }

    if (data.attending === 'no' && data.subName) {
      writeVisitorLog_(data, {
        name:      data.subName,
        prefix:    data.subDialCode || '',
        contact:   data.subPhone,
        invitedBy: data.memberName + ' (sub)',
        isSub:     true,
      });
    }

    if (data.hasVisitor === 'yes' && Array.isArray(data.visitors)) {
      data.visitors.forEach(v => writeVisitorLog_(data, v));
    }

    return jsonOk_({ written: true });
  } catch (err) {
    logError_('doPost', err);
    return jsonErr_(err.message);
  }
}

// ── Write helpers ─────────────────────────────────────────────────────────────
function writeSubmission_(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.SUBMISSIONS);
  const tz    = Session.getScriptTimeZone();
  const ts    = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");

  const visitors     = Array.isArray(data.visitors) ? data.visitors : [];
  const visitorNames = visitors.map(v => v.name).join(', ');
  const visitorDials = visitors.map(v => v.prefix || '').join(', ');
  const visitorPhones= visitors.map(v => v.contact || '').join(', ');
  const invitedBy    = visitors.map(v => data.memberName).join(', ');

  sheet.appendRow([
    ts,
    data.weekOf        || '',
    data.memberName    || '',
    data.attending     || '',
    data.paymentStatus || '',
    data.subName       || '',
    data.subDialCode   || '',
    data.subPhone      || '',
    data.subStatus     || '',
    data.hasVisitor    || 'no',
    visitors.length    || 0,
    visitorNames,
    visitorDials,
    visitorPhones,
    invitedBy,
    data.parkingCoupons || '',
  ]);
}

function writeAttendance_(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.ACTUAL_ATTENDANCE);
  const tz    = Session.getScriptTimeZone();

  const rawDate  = (data.weekOf || '').trim();
  const dateObj  = new Date(rawDate);
  const dateLabel = isNaN(dateObj.getTime())
    ? rawDate
    : Utilities.formatDate(dateObj, tz, 'dd MMM yyyy');

  const lastCol = sheet.getLastColumn();
  let dateCol = -1;
  if (lastCol >= 4) {
    sheet.getRange(5, 4, 1, lastCol - 3).getValues()[0].forEach((v, i) => {
      const lbl = (v instanceof Date)
        ? Utilities.formatDate(v, tz, 'dd MMM yyyy')
        : String(v).trim();
      if (lbl === dateLabel) dateCol = 4 + i;
    });
  }

  if (dateCol === -1) {
    dateCol = lastCol + 1;
    sheet.getRange(5, dateCol).setValue(dateLabel).setFontWeight('bold').setHorizontalAlignment('center');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) return;

  const names = sheet.getRange(6, 2, lastRow - 5, 2).getValues();
  const normTarget = data.memberName.trim().toLowerCase();

  names.forEach((row, i) => {
    const first   = String(row[0]).trim();
    const surname = String(row[1]).trim();
    const full    = (first + ' ' + surname).trim().toLowerCase();
    const firstLower = first.toLowerCase();

    if (full === normTarget || firstLower === normTarget) {
      let statusVal = '';
      if (data.status === 'Absent') {
        statusVal = 'Absent (1)';
      } else if (data.status === 'Late') {
        const mins = parseInt(data.lateMins, 10) || 0;
        statusVal = mins > 0 ? `Late (${mins} min) (1)` : 'Late (1)';
      } else if (data.status === 'Y') {
        statusVal = 'Y';
      } else {
        statusVal = data.status || '';
      }
      sheet.getRange(6 + i, dateCol).setValue(statusVal);
    }
  });
}

function writeVisitorLog_(data, visitor) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.VISITOR_LOG);
  const tz    = Session.getScriptTimeZone();
  const ts    = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");

  sheet.appendRow([
    ts,
    data.weekOf     || '',
    visitor.name    || '',
    visitor.prefix  || '',
    visitor.contact || '',
    data.memberName || '',
    visitor.isSub ? 'Sub' : 'Visitor',
  ]);
}

function writeNewMember_(data) {
  const ss            = SpreadsheetApp.openById(SHEET_ID);
  const tz            = Session.getScriptTimeZone();
  const ts            = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");
  const preferredName = (data.firstName || '').trim();
  const surname       = (data.surname   || '').trim();
  const memberName    = (preferredName + (surname ? ' ' + surname : '')).trim() || data.fullName || 'Unknown';

  // Save headshot to Drive — named "<Preferred Name> <Last Name>"
  if (data.headshotBase64 && data.headshotMime) {
    try {
      const ext = data.headshotMime === 'image/png' ? '.png' : '.jpg';
      saveFileToDrive_(data.headshotBase64, data.headshotMime, memberName + ext, HEADSHOTS_FOLDER_ID);
    } catch(e) { Logger.log('Headshot upload error: ' + e.message); }
  }

  // Save logo to Drive, make link-accessible, store URL in Active_Members col AB
  let logoUrl = '';
  if (data.logoBase64 && data.logoMime) {
    try {
      const ext  = data.logoMime === 'image/png' ? '.png' : '.jpg';
      const file = saveFileToDrive_(data.logoBase64, data.logoMime, memberName + ' Logo' + ext, LOGOS_FOLDER_ID);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      logoUrl = 'https://drive.google.com/uc?id=' + file.getId();
    } catch(e) { Logger.log('Logo upload error: ' + e.message); }
  }

  // Write to Members sheet
  const membersSheet = ss.getSheetByName(SN.MEMBERS);
  membersSheet.appendRow([
    data.fullName       || '',  // A: Name
    data.tradeCategory  || '',  // B: Trade / Industry
    data.businessName   || '',  // C: Company
    data.mobile         || '',  // D: Phone
    data.businessEmail  || '',  // E: Email
    'Y',                        // F: Active (Y/N)
    data.birthday       || '',  // G: Birthday
    data.industry       || '',  // H: Industry
    data.sponsoredBy    || '',  // I: Sponsored By
    data.inTradeSince   || '',  // J: In Trade Since
    data.officeAddress  || '',  // K: Office Address
    data.companyWebsite || '',  // L: Website
    data.fbPage         || '',  // M: Facebook
    data.businessIG     || '',  // N: Instagram
    data.linkedin       || '',  // O: LinkedIn
    data.otherSocial    || '',  // P: Other Social
    data.driveEmail     || '',  // Q: Drive Email
    data.calendarEmail  || '',  // R: Calendar Email
    data.nextOfKin      || '',  // S: Next of Kin
    ts,                         // T: Submission timestamp
  ]);

  // Add to Active_Members — write all fields to matching columns
  const activeSheet = ss.getSheetByName(SN.ACTIVE_MEMBERS);
  if (activeSheet) {
    const nextRow = activeSheet.getLastRow() + 1;

    // Auto-number: find highest No in col A and increment
    let nextNo = 1;
    if (nextRow > 2) {
      activeSheet.getRange(2, 1, nextRow - 2, 1).getValues().forEach(r => {
        const n = parseInt(r[0], 10);
        if (!isNaN(n) && n >= nextNo) nextNo = n + 1;
      });
    }

    // Build row array for cols A–AL (38 cols), blanks for non-form columns
    const row = new Array(38).fill('');
    row[0]  = nextNo;                              // A:  No
    row[1]  = data.fullName        || memberName;  // B:  Member Full Name
    row[2]  = preferredName        || '';          // C:  First Name
    row[3]  = surname              || '';          // D:  Surname
    row[4]  = data.birthday        || '';          // E:  Birthday
    row[5]  = data.sponsoredBy     || '';          // F:  Sponsored by
    // G: Notes — blank
    // H: Empower Power Team — blank (set manually)
    row[8]  = data.industry        || '';          // I:  Industry
    row[9]  = data.tradeCategory   || '';          // J:  Trade Category
    row[10] = data.businessName    || '';          // K:  Business Name
    // L–O: Tag line, one-liner, paragraph, products — blank (filled separately)
    row[15] = data.inTradeSince    || '';          // P:  In Trade Since
    // Q–Z: referrals, hobbies, etc. — blank
    row[26] = data.companyWebsite  || '';          // AA: Company Website
    row[27] = logoUrl              || '';          // AB: Company Logo URL
    row[28] = data.fbPage          || '';          // AC: FB Page
    row[29] = data.businessIG      || '';          // AD: Business IG
    row[30] = data.linkedin        || '';          // AE: LinkedIn
    row[31] = data.officeAddress   || '';          // AF: Office Address
    row[32] = data.mobile          || '';          // AG: Mobile
    row[33] = data.nextOfKin       || '';          // AH: Next of KIN Contact
    row[34] = data.businessEmail   || '';          // AI: Business Email Address
    row[35] = data.driveEmail      || '';          // AJ: Google Drive Email
    row[36] = data.calendarEmail   || '';          // AK: Google Calendar Email
    row[37] = data.otherSocial     || '';          // AL: Other Social Media

    activeSheet.getRange(nextRow, 1, 1, 38).setValues([row]);
  }
}

function saveFileToDrive_(base64, mimeType, fileName, folderId) {
  const blob   = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, fileName);
  const folder = DriveApp.getFolderById(folderId);
  return folder.createFile(blob);
}

// ── Power Teams ───────────────────────────────────────────────────────────────
function getPowerTeams_() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('POWER TEAM LIST');
  if (!sheet) return jsonErr_('POWER TEAM LIST sheet not found');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonOk_({ teams: [] });

  const data = sheet.getRange(1, 1, lastRow, 10).getValues();

  function buildTeam(rowIdx, nameCol, memberCol) {
    const name    = String(data[rowIdx][nameCol - 1]).trim();
    const trades  = [];
    for (let r = rowIdx; r < Math.min(rowIdx + 14, data.length); r++) {
      const trade = String(data[r][memberCol - 1]).trim();
      if (trade) trades.push(trade);
    }
    return { name, trades };
  }

  const teams = [
    buildTeam(1,  2, 3),
    buildTeam(1,  5, 6),
    buildTeam(1,  8, 9),
    buildTeam(16, 2, 3),
    buildTeam(16, 5, 6),
    buildTeam(16, 8, 9),
  ];

  return jsonOk_({ teams });
}

// ── Upcoming presenters (read live by the Member Hub) ─────────────────────────
/**
 * Reads the Roster sheet and returns the upcoming meetings (today or later),
 * soonest first, so the hub can show this week + the week after.
 *
 * Roster layout (live "Official BNI Empower Attendance 2026"):
 *   Row 1 = headers.  Data from row 2.
 *   A: Meeting Date
 *   C: Network Education
 *   D: Core Value Sharing
 *   E: Featured Presentation (presenter 1)
 *   F: Featured Presentation (presenter 2 — optional; blank if only one)
 */
function getNextPresenters_() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.ROSTER);
  if (!sheet || sheet.getLastRow() < 2) {
    return jsonOk_({ upcoming: [] });
  }

  const tz    = Session.getScriptTimeZone();
  const now   = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate()); // midnight today

  const rows  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues(); // A..F
  const clean = v => { const s = String(v).trim(); return (s && s !== '-') ? s : ''; };

  const parsed = rows.map(r => {
    let dateObj = (r[0] instanceof Date) ? r[0] : new Date(String(r[0]).trim());
    if (isNaN(dateObj.getTime())) dateObj = null;
    return {
      dateObj,
      date:       dateObj ? Utilities.formatDate(dateObj, tz, 'EEEE, dd MMM yyyy') : clean(r[0]),
      networkEdu: clean(r[2]),                         // C
      coreValue:  clean(r[3]),                         // D
      featured:   [clean(r[4]), clean(r[5])].filter(Boolean), // E + F (0, 1 or 2 names)
    };
  }).filter(p => p.dateObj); // keep only rows with a real date

  parsed.sort((a, b) => a.dateObj - b.dateObj);

  const upcoming = parsed
    .filter(p => p.dateObj >= today)
    .slice(0, 4)
    .map(p => ({ date: p.date, networkEdu: p.networkEdu, coreValue: p.coreValue, featured: p.featured }));

  return jsonOk_({ upcoming });
}

// ── Slide structure dump (read-only — for mapping the deck) ───────────────────
/**
 * Returns the presentation's structure so the box object IDs can be mapped.
 *   ?action=dumpSlides&pin=####                 → compact index of every slide
 *   ?action=dumpSlides&pin=####&find=Referral   → only slides whose text matches
 *   ?action=dumpSlides&pin=####&slide=42        → full detail for one slide (1-based)
 */
function dumpSlides_(find, slideParam, rawMode) {
  const fields = 'slides(objectId,pageElements(objectId,size,transform,' +
                 'shape(shapeType,text(textElements(textRun(content)))),image(contentUrl)))';
  const url  = 'https://slides.googleapis.com/v1/presentations/' + PRESENTATION_ID + '?fields=' + encodeURIComponent(fields);
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) {
    return jsonErr_('Slides API ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 300));
  }

  const raw    = JSON.parse(resp.getContentText()).slides || [];
  const findLC = String(find || '').trim().toLowerCase();
  const only   = slideParam ? parseInt(slideParam, 10) : 0;

  const slides = raw.map((s, i) => {
    const elements = (s.pageElements || []).map(pe => {
      let text = '';
      if (pe.shape && pe.shape.text && pe.shape.text.textElements) {
        text = pe.shape.text.textElements.map(te => (te.textRun ? te.textRun.content : '')).join('');
        // raw mode preserves line breaks (shown as \n) so cell structure is visible.
        text = rawMode ? text.replace(/\n/g, '\\n').replace(/[ \t]+/g, ' ').trim() : text.replace(/\s+/g, ' ').trim();
      }
      const t  = pe.transform || {};
      const sz = pe.size || {};
      return {
        id:   pe.objectId,
        kind: pe.image ? 'image' : (pe.shape ? (pe.shape.shapeType || 'shape') : 'other'),
        text: text.slice(0, rawMode ? 240 : 90),
        x:    Math.round(t.translateX || 0),
        y:    Math.round(t.translateY || 0),
        sx:   +Number(t.scaleX || 1).toFixed(4),
        sy:   +Number(t.scaleY || 1).toFixed(4),
        w:    sz.width  ? Math.round(sz.width.magnitude)  : 0,
        h:    sz.height ? Math.round(sz.height.magnitude) : 0,
      };
    });
    const allText = elements.map(el => el.text).filter(Boolean).join(' | ');
    return { n: i + 1, slideId: s.objectId, textPreview: allText.slice(0, 160), elements };
  });

  // Single slide detail
  if (only) {
    const one = slides[only - 1];
    return jsonOk_({ slideCount: slides.length, slide: one || null });
  }

  // Filtered detail
  if (findLC) {
    const matched = slides.filter(sl => sl.elements.some(el => el.text.toLowerCase().includes(findLC)));
    return jsonOk_({ slideCount: slides.length, matches: matched.length, slides: matched });
  }

  // Compact index (no per-element detail, to keep it small)
  return jsonOk_({
    presentationId: PRESENTATION_ID,
    slideCount: slides.length,
    index: slides.map(sl => ({ n: sl.n, slideId: sl.slideId, elements: sl.elements.length, textPreview: sl.textPreview })),
  });
}

// ── VP Committee Report → Slide 98 running totals ─────────────────────────────
// The VP submits THIS WEEK's numbers from a phone form; they auto-add to the
// "Since Launch" totals on Slide 98 and are logged in the Committee_Report sheet.
const STAT_SLIDE_ID = 'g3d1b184308e_0_0';   // Slide 98 — "BNI Empower Statistics (Since Launch)"
const STAT_BOX = {
  referrals: 'g3d1b184308e_0_10',   // "332 Referrals Passed"
  visitors:  'g3d1b184308e_0_11',   // "89 Visitors Invited"
  business:  'g3d1b184308e_0_12',   // "$209,288 Business Done"
};
// Starting totals currently on the slide — used only when Committee_Report is empty.
const STAT_SEED = { referrals: 332, visitors: 89, business: 209288 };

function fmtInt_(n) { return Number(n || 0).toLocaleString('en-US'); }

function committeeDisplay_(t) {
  return {
    referrals: fmtInt_(t.referrals),
    visitors:  fmtInt_(t.visitors),
    business:  '$' + fmtInt_(t.business),
  };
}

// Current running totals (last row of Committee_Report, else the seed).
function getCommitteeTotals_() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.COMMITTEE_REPORT);
  if (sheet && sheet.getLastRow() >= 2) {
    const last = sheet.getRange(sheet.getLastRow(), 6, 1, 3).getValues()[0]; // F,G,H
    return {
      referrals: Number(last[0]) || STAT_SEED.referrals,
      visitors:  Number(last[1]) || STAT_SEED.visitors,
      business:  Number(last[2]) || STAT_SEED.business,
    };
  }
  return Object.assign({}, STAT_SEED);
}

function getCommitteeTotalsAction_() {
  const t = getCommitteeTotals_();
  return jsonOk_({ totals: t, display: committeeDisplay_(t) });
}

// Append this week's numbers, bump the totals, update the slide.
function writeCommitteeReport_(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SN.COMMITTEE_REPORT);
  if (!sheet) {
    sheet = ss.insertSheet(SN.COMMITTEE_REPORT);
    sheet.getRange(1, 1, 1, 9)
      .setValues([['Timestamp', 'Week Of', 'Referrals (wk)', 'Visitors (wk)', 'Business (wk)',
                   'Referrals Total', 'Visitors Total', 'Business Total', 'Submitted By']])
      .setFontWeight('bold').setBackground('#1a5276').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    [150, 150, 110, 110, 120, 120, 110, 130, 140].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  }

  const tz = Session.getScriptTimeZone();
  const ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");

  const num = v => { const n = parseFloat(String(v).replace(/[^0-9.]/g, '')); return isNaN(n) ? 0 : n; };
  const wk = {
    referrals: Math.round(num(data.referrals)),
    visitors:  Math.round(num(data.visitors)),
    business:  Math.round(num(data.business)),
  };

  const prev = getCommitteeTotals_();
  const next = {
    referrals: prev.referrals + wk.referrals,
    visitors:  prev.visitors  + wk.visitors,
    business:  prev.business  + wk.business,
  };

  sheet.appendRow([
    ts, data.weekOf || '',
    wk.referrals, wk.visitors, wk.business,
    next.referrals, next.visitors, next.business,
    data.submittedBy || '',
  ]);

  updateStatSlideNumbers_(next);
  return next;
}

// Reads the current text of the 3 stat boxes on the stats slide → { boxId: text }.
function readStatBoxes_() {
  const url  = 'https://slides.googleapis.com/v1/presentations/' + PRESENTATION_ID +
               '/pages/' + STAT_SLIDE_ID +
               '?fields=' + encodeURIComponent('pageElements(objectId,shape(text(textElements(textRun(content)))))');
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) {
    logError_('readStatBoxes_', new Error(resp.getContentText().slice(0, 200)));
    return {};
  }
  const els  = JSON.parse(resp.getContentText()).pageElements || [];
  const byId = {};
  els.forEach(pe => {
    let t = '';
    if (pe.shape && pe.shape.text && pe.shape.text.textElements) {
      t = pe.shape.text.textElements.map(te => (te.textRun ? te.textRun.content : '')).join('');
    }
    byId[pe.objectId] = t;
  });
  return byId;
}

// Self-healing: reads whatever number is currently in each box and overwrites
// just that number with the new total. Targets each box by its own objectId,
// so a manual slide edit is corrected on the next submit — and boxes never
// interfere with each other. Number styling is preserved (we replace the
// number in place, leaving the label untouched).
function updateStatSlideNumbers_(next) {
  const boxText = readStatBoxes_();
  const plan = [
    { id: STAT_BOX.referrals, neu: fmtInt_(next.referrals) },
    { id: STAT_BOX.visitors,  neu: fmtInt_(next.visitors) },
    { id: STAT_BOX.business,  neu: '$' + fmtInt_(next.business) },
  ];

  const requests = [];
  plan.forEach(p => {
    const text = String(boxText[p.id] || '');
    const m = text.match(/\$?\d[\d,]*/);   // first number token in the box
    if (!m) { logError_('updateStatSlideNumbers_', new Error('no number in box ' + p.id)); return; }
    const oldTok = m[0];
    const start  = m.index;
    if (oldTok === p.neu) return;           // already correct — nothing to do
    // Insert the new number, then delete the old one (keeps the number's styling).
    requests.push({ insertText: { objectId: p.id, insertionIndex: start, text: p.neu } });
    requests.push({ deleteText: { objectId: p.id, textRange: {
      type: 'FIXED_RANGE',
      startIndex: start + p.neu.length,
      endIndex:   start + p.neu.length + oldTok.length,
    } } });
  });
  if (!requests.length) return;

  const url  = 'https://slides.googleapis.com/v1/presentations/' + PRESENTATION_ID + ':batchUpdate';
  const resp = UrlFetchApp.fetch(url, {
    method:      'post',
    contentType: 'application/json',
    headers:     { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload:     JSON.stringify({ requests }),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) {
    logError_('updateStatSlideNumbers_', new Error(resp.getContentText().slice(0, 300)));
  }
}

// ── Weekly "Next Presenter" chain (slides 62–91) ──────────────────────────────
// Ordered weekly presenters + the single on-slide "Next Presenter" box each one
// carries. Format: "Next Presenter: {next} ⇒ {the one after}". The last presenter
// has no box; the second-to-last shows only the final name.
const NEXT_PRES_ORDER = [
  { name: 'Ariel Chong',       box: 'g3d498035fa2_1_62'  },
  { name: 'Arun Prasad',       box: 'g3f56d718e25_0_31'  },
  { name: 'Ben Wong',          box: 'g3d037e0c29a_0_266' },
  { name: 'Ben Tee',           box: 'g3f56d718e25_0_32'  },
  { name: 'Benjamin Ng',       box: 'g3d037e0c29a_0_268' },
  { name: 'Daniel Yen',        box: 'g3d037e0c29a_0_269' },
  { name: 'Deborah Chueh',     box: 'g3ec253ce018_0_13'  },
  { name: 'Delia Tan',         box: 'g3f56d718e25_0_30'  },
  { name: 'Ismail Khamis',     box: 'g3e31e8fdc1f_1_0'   },
  { name: 'Ivan Ang',          box: 'g3d037e0c29a_0_273' },
  { name: 'Jaron Chan',        box: 'g3efe084945a_1_1'   },
  { name: 'Jay Tan',           box: 'g3f22d44700a_3_1'   },
  { name: 'Lee Jia Zheng',     box: 'g3f7ba902fea_0_10'  },
  { name: 'Zhang Junxian',     box: 'g3dcdba80422_1_9'   },
  { name: 'Joanne Sooi',       box: 'g3d037e0c29a_0_277' },
  { name: 'Jonathan Tan',      box: 'g3f7b7626906_0_5'   },
  { name: 'Kay Tan',           box: 'g3d037e0c29a_0_278' },
  { name: 'Kevin Phua',        box: 'g3d037e0c29a_0_279' },
  { name: 'Lawrence Ku',       box: 'g3d037e0c29a_0_280' },
  { name: 'Lee E Mae',         box: 'g3e6b5710eea_0_5'   },
  { name: 'Mark Duma',         box: 'g3d037e0c29a_0_281' },
  { name: 'Pamela Lin',        box: 'g3d037e0c29a_0_282' },
  { name: 'Rajiv',             box: 'g3dcdba80422_1_20'  },
  { name: 'Sandy Au',          box: 'g3da1b38dbc6_0_39'  },
  { name: 'Zhao Shu Hui',      box: 'g3d037e0c29a_0_286' },
  { name: 'Kuek Yu Xi',        box: 'g3d037e0c29a_0_288' },
  { name: 'Zefirelli Noordin', box: 'g3da1b38dbc6_0_41'  },
  { name: 'Rachel Teo',        box: 'g3da1b38dbc6_0_42'  },
  { name: 'Pang Wee Khai',     box: 'g3d037e0c29a_0_287' },
  { name: 'Iskons',            box: null                 }, // last — no box
];
// Off-screen leftover boxes from manual week-to-week editing (to remove).
const NEXT_PRES_ORPHANS = ['g6cba5c9273fffe28_7', 'g3ea5c59300a_0_15', 'g3f56d718e25_0_29'];

// Each member's 30-sec weekly presentation slide (slides 62–91), in order.
// Used to hide/unhide (skip/unskip) slides based on attendance.
const WEEKLY_SLIDES = [
  { name: 'Ariel Chong',       id: 'g3d498035fa2_1_55'  },
  { name: 'Arun Prasad',       id: 'g6cba5c9273fffe28_1' },
  { name: 'Ben Wong',          id: 'g3d037e0c29a_0_62'   },
  { name: 'Ben Tee',           id: 'g3d037e0c29a_0_255'  },
  { name: 'Benjamin Ng',       id: 'g3d037e0c29a_0_78'   },
  { name: 'Daniel Yen',        id: 'g3d037e0c29a_0_86'   },
  { name: 'Deborah Chueh',     id: 'g3d037e0c29a_0_94'   },
  { name: 'Delia Tan',         id: 'g3e95dbb8fd0_0_0'    },
  { name: 'Ismail Khamis',     id: 'g3d037e0c29a_0_110'  },
  { name: 'Ivan Ang',          id: 'g3d037e0c29a_0_118'  },
  { name: 'Jaron Chan',        id: 'g3d037e0c29a_0_126'  },
  { name: 'Jay Tan',           id: 'g3d037e0c29a_0_134'  },
  { name: 'Lee Jia Zheng',     id: 'g3d037e0c29a_0_142'  },
  { name: 'Zhang Junxian',     id: 'g3dcdba80422_1_3'    },
  { name: 'Joanne Sooi',       id: 'g3d037e0c29a_0_150'  },
  { name: 'Jonathan Tan',      id: 'g3f7b7626906_0_0'    },
  { name: 'Kay Tan',           id: 'g3d037e0c29a_0_158'  },
  { name: 'Kevin Phua',        id: 'g3d037e0c29a_0_166'  },
  { name: 'Lawrence Ku',       id: 'g3d037e0c29a_0_174'  },
  { name: 'Lee E Mae',         id: 'g3e6b5710eea_0_0'    },
  { name: 'Mark Duma',         id: 'g3d037e0c29a_0_182'  },
  { name: 'Pamela Lin',        id: 'g3d037e0c29a_0_190'  },
  { name: 'Rajiv',             id: 'g3dcdba80422_1_14'   },
  { name: 'Sandy Au',          id: 'g3d037e0c29a_0_206'  },
  { name: 'Zhao Shu Hui',      id: 'g3d037e0c29a_0_222'  },
  { name: 'Kuek Yu Xi',        id: 'g3d037e0c29a_0_238'  },
  { name: 'Zefirelli Noordin', id: 'g3d037e0c29a_0_246'  },
  { name: 'Rachel Teo',        id: 'g3d037e0c29a_0_198'  },
  { name: 'Pang Wee Khai',     id: 'g3d037e0c29a_0_230'  },
  { name: 'Iskons',            id: 'g3d037e0c29a_0_214'  },
];

// Unhide (unskip) every weekly presentation slide; returns names that were hidden.
function unhideAllWeeklySlides_() {
  const want = {};
  WEEKLY_SLIDES.forEach(w => { want[w.id] = w.name; });
  const pres = SlidesApp.openById(PRESENTATION_ID);
  const unhidden = [];
  pres.getSlides().forEach(s => {
    const id = s.getObjectId();
    if (want[id] && s.isSkipped()) { s.setSkipped(false); unhidden.push(want[id]); }
  });
  return unhidden;
}

// Reset to the all-present baseline: correct Next Presenter chain + unhide everyone.
function resetAllPresent_() {
  const chain    = applyNextPresenterChain_(NEXT_PRES_ORDER);
  const unhidden = unhideAllWeeklySlides_();
  return { changed: chain.changed, deleted: chain.deleted, unhidden: unhidden, unhiddenCount: unhidden.length };
}

const NP_ARROW = '⇒'; // ⇒

function nextPresenterText_(order, i) {
  const a = order[i + 1] ? order[i + 1].name : '';
  const b = order[i + 2] ? order[i + 2].name : '';
  // Label on its own line, presenters on the next line (line break after the colon).
  if (a && b) return 'Next Presenter:\n' + a + ' ' + NP_ARROW + ' ' + b;
  if (a)      return 'Next Presenter:\n' + a;
  return '';
}

// Normalise arrow glyph + spacing but PRESERVE the line break, so a box that's
// missing the break after "Next Presenter:" is detected and rewritten.
function npNorm_(s) {
  return String(s || '')
    .replace(/\r/g, '')
    .replace(/[→⇒⟹➔⮕>]+/g, '@')
    .replace(/[ \t]+/g, ' ')
    .replace(/ *\n+ */g, '\n')
    .trim();
}

// { objectId: text } for every shape text box in the deck (one API call).
function readDeckBoxText_() {
  const url  = 'https://slides.googleapis.com/v1/presentations/' + PRESENTATION_ID +
               '?fields=' + encodeURIComponent('slides(pageElements(objectId,shape(text(textElements(textRun(content))))))');
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  const map = {};
  if (resp.getResponseCode() !== 200) { logError_('readDeckBoxText_', new Error(resp.getContentText().slice(0, 200))); return map; }
  (JSON.parse(resp.getContentText()).slides || []).forEach(s => {
    (s.pageElements || []).forEach(pe => {
      if (pe.shape && pe.shape.text && pe.shape.text.textElements) {
        map[pe.objectId] = pe.shape.text.textElements.map(te => (te.textRun ? te.textRun.content : '')).join('');
      }
    });
  });
  return map;
}

function slidesBatchUpdate_(requests) {
  const url  = 'https://slides.googleapis.com/v1/presentations/' + PRESENTATION_ID + ':batchUpdate';
  const resp = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ requests }), muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) logError_('slidesBatchUpdate_', new Error(resp.getContentText().slice(0, 400)));
  return resp.getResponseCode();
}

// Writes the Next Presenter chain for the given order. Only rewrites boxes whose
// text differs (styling preserved via insert-then-delete) and removes orphans.
function applyNextPresenterChain_(order) {
  const cur = readDeckBoxText_();
  const requests = [];
  let changed = 0, deleted = 0;

  NEXT_PRES_ORPHANS.forEach(id => {
    if (cur[id] !== undefined) { requests.push({ deleteObject: { objectId: id } }); deleted++; }
  });

  order.forEach((m, i) => {
    if (!m.box) return;
    const target  = nextPresenterText_(order, i);
    if (!target) return;
    const old     = String(cur[m.box] || '');
    const oldCore = old.replace(/\n+$/, '');            // keep the box's trailing newline
    if (npNorm_(oldCore) === npNorm_(target)) return;   // already correct
    requests.push({ insertText: { objectId: m.box, insertionIndex: 0, text: target } });
    if (oldCore.length) {
      requests.push({ deleteText: { objectId: m.box, textRange: {
        type: 'FIXED_RANGE', startIndex: target.length, endIndex: target.length + oldCore.length,
      } } });
    }
    changed++;
  });

  if (requests.length) slidesBatchUpdate_(requests);
  return { changed, deleted };
}

function resetNextPresentersAllPresent() {
  const r = resetAllPresent_();
  try {
    SpreadsheetApp.getUi().alert('✅ Reset to all present\n\n' +
      'Next Presenter boxes corrected: ' + r.changed + '\n' +
      'Off-screen leftovers removed: ' + r.deleted + '\n' +
      'Hidden slides unhidden: ' + r.unhiddenCount +
      (r.unhidden.length ? '\n(' + r.unhidden.join(', ') + ')' : ''));
  } catch (_) {}
  return r;
}

// ── Grid team: Membership Committee (slide 11 / p10) ──────────────────────────
// Each member has a photo (image) + a name box. On absence the pair moves
// off-screen and the row's remaining members re-centre. Vertical position and
// scale never change — only horizontal position (and hide/show).
const COMMITTEE = {
  slideId: 'p10',
  rows: {
    1: ['benjamin wong', 'pamela lin', 'joanne sooi'],
    2: ['jay tan', 'deborah chueh'],
  },
  el: {
    'benjamin wong': { img: { id: 'p10_i238', tx: 2989363, ty: 1459827, sx: 82.7942, sy: 91.0733, w: 18900 },  name: { id: 'p10_i239', tx: 2713964, ty: 3240212, sx: 0.7052, sy: 0.1263, w: 3000000 } },
    'pamela lin':    { img: { id: 'p10_i236', tx: 5247835, ty: 1451502, sx: 58.5594, sy: 64.5887, w: 26650 },  name: { id: 'p10_i251', tx: 5038194, ty: 3259522, sx: 0.7052, sy: 0.1865, w: 3000000 } },
    'joanne sooi':   { img: { id: 'p10_i245', tx: 7636840, ty: 1518926, sx: 52.2278, sy: 57.3763, w: 30000 },  name: { id: 'p10_i248', tx: 7362419, ty: 3259509, sx: 0.7052, sy: 0.1865, w: 3000000 } },
    'jay tan':       { img: { id: 'p10_i249', tx: 4064715, ty: 4085988, sx: 30.6023, sy: 33.6189, w: 51200 },  name: { id: 'p10_i250', tx: 3790294, ty: 5826571, sx: 0.7052, sy: 0.1865, w: 3000000 } },
    'deborah chueh': { img: { id: 'g3e4e817fee7_0_0', tx: 6322638, ty: 4069292, sx: 58.56, sy: 43.8676, w: 26650 }, name: { id: 'p10_i241', tx: 5801988, ty: 5895767, sx: 0.8673, sy: 0.1619, w: 3000000 } },
  },
};

const GRID_OFFSCREEN = 99000000;

function elemTransform_(el, translateX) {
  return { updatePageElementTransform: {
    objectId:  el.id,
    transform: { scaleX: el.sx, scaleY: el.sy, translateX: Math.round(translateX), translateY: el.ty, unit: 'EMU' },
    applyMode: 'ABSOLUTE',
  } };
}
function elemCentre_(el) { return el.tx + (el.w * el.sx) / 2; }         // current visual centre-X
function centreToTx_(el, cx) { return cx - (el.w * el.sx) / 2; }        // translateX to put centre at cx

// Reflow one grid team. absentSet = lowercased member names to hide.
function reflowGridTeam_(team, absentSet) {
  const requests = [];
  Object.values(team.rows).forEach(members => {
    const centres = members.map(m => elemCentre_(team.el[m].img));
    const spacing = centres.length > 1 ? (centres[centres.length - 1] - centres[0]) / (centres.length - 1) : 2300000;
    const mid     = centres.reduce((a, b) => a + b, 0) / centres.length;
    const present = members.filter(m => !absentSet.has(m));
    const k       = present.length;
    const allHere = (k === members.length);

    members.forEach(m => {
      const e = team.el[m];
      if (absentSet.has(m)) {
        requests.push(elemTransform_(e.img,  GRID_OFFSCREEN));
        requests.push(elemTransform_(e.name, GRID_OFFSCREEN));
        return;
      }
      const idx = present.indexOf(m);
      const cx  = allHere ? elemCentre_(e.img) : Math.round(mid + (idx - (k - 1) / 2) * spacing);
      requests.push(elemTransform_(e.img,  centreToTx_(e.img,  cx)));
      requests.push(elemTransform_(e.name, centreToTx_(e.name, cx)));
    });
  });
  if (requests.length) slidesBatchUpdate_(requests);
  return requests.length / 2; // element-pairs touched
}

// Test/apply helper: parse "a, b" → lowercased set.
function parseAbsent_(s) {
  return new Set(String(s || '').toLowerCase().split(',').map(x => x.trim()).filter(Boolean));
}

// ── Roles roster (President-managed source of truth for the grid slides) ───────
// Team | Role | Member | Trade. Seeded once from the current deck; the President
// edits this going forward (term swaps, new members). The slide automation
// reads THIS to fill each team's grid.
const ROLES_TEAMS = ['Leadership Team', 'Membership Committee', 'Visitor Host', 'Support Leadership'];

// Seed: [Team, Role, Member] read off the current slides. Trade is looked up
// from Active_Members automatically. The President verifies/adjusts after.
const ROLES_SEED = [
  ['Leadership Team',      'President',               'Iskons'],
  ['Leadership Team',      'Vice-President',          'Pang Wee Khai'],
  ['Leadership Team',      'Secretary Treasurer',     'Rachel Teo'],
  ['Membership Committee', 'Committee Member',        'Benjamin Wong'],
  ['Membership Committee', 'Committee Member',        'Pamela Lin'],
  ['Membership Committee', 'Committee Member',        'Joanne Sooi'],
  ['Membership Committee', 'Committee Member',        'Jay Tan'],
  ['Membership Committee', 'Committee Member',        'Deborah Chueh'],
  ['Visitor Host',         'Visitor Host',            'Sandy Au'],
  ['Visitor Host',         'Visitor Host',            'Lee Jia Zheng'],
  ['Visitor Host',         'Visitor Host',            'Ivan Ang'],
  ['Visitor Host',         'Visitor Host',            'Daniel Yen'],
  ['Visitor Host',         'Visitor Orientation',     'Kay Tan'],
  ['Visitor Host',         'Visitor Host',            'Kuek Yu Xi'],
  ['Support Leadership',   'Network Education',       'Kevin Phua'],
  ['Support Leadership',   'Presentation Taskforce',  'Benjamin Ng'],
  ['Support Leadership',   'Event Coordinator',       'Delia Tan'],
  ['Support Leadership',   'Growth Coordinator',      'Zefirelli Noordin'],
  ['Support Leadership',   'Visitor Registration',    'Zhao Shu Hui'],
  ['Support Leadership',   'Photography & Videography','Mark Duma'],
  ['Support Leadership',   'Logistics',               'Jaron Chan'],
  ['Support Leadership',   'Marketing Lead',          'Arun Prasad'],
  ['Support Leadership',   'Tech Coordinator',        'Lee Jia Zheng'],
];

// Build member name → trade from Active_Members (col B name, col J trade category).
function tradeByMember_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const am = ss.getSheetByName(SN.ACTIVE_MEMBERS);
  const map = {};
  if (am && am.getLastRow() > 1) {
    am.getRange(2, 2, am.getLastRow() - 1, 9).getValues().forEach(r => { // B..J
      const name  = String(r[0]).trim().toLowerCase();
      const trade = String(r[8]).trim();
      if (name) map[name] = trade;
    });
  }
  return map;
}

// Reads the Roles sheet → [{team, role, member, trade}]. Empty if not set up.
function getRoles_() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.ROLES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .map(r => ({ team: String(r[0]).trim(), role: String(r[1]).trim(), member: String(r[2]).trim(), trade: String(r[3]).trim() }))
    .filter(r => r.member);
}

// Creates/refreshes the Roles sheet with dropdowns; seeds once if empty.
function setupRolesSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SN.ROLES);
  if (!sheet) sheet = ss.insertSheet(SN.ROLES);

  sheet.getRange(1, 1, 1, 4)
    .setValues([['Team', 'Role', 'Member', 'Trade']])
    .setBackground('#1a5276').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  [180, 210, 180, 240].forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  // Dropdowns: Team (col A) and Member (col C, from roster)
  const teamRule = SpreadsheetApp.newDataValidation().requireValueInList(ROLES_TEAMS, true).setAllowInvalid(true).build();
  const nameRule = SpreadsheetApp.newDataValidation().requireValueInList(getRoster_(), true).setAllowInvalid(true).build();
  sheet.getRange(2, 1, 300, 1).setDataValidation(teamRule);
  sheet.getRange(2, 3, 300, 1).setDataValidation(nameRule);

  let seeded = 0;
  if (sheet.getLastRow() < 2) {
    const trades = tradeByMember_();
    const rows = ROLES_SEED.map(([team, role, member]) => [team, role, member, trades[member.toLowerCase()] || '']);
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    seeded = rows.length;
  }

  try {
    SpreadsheetApp.getUi().alert('✅ Roles sheet ready.\n\n' +
      (seeded ? 'Seeded ' + seeded + ' roles from the current slides.\nPlease verify the roles/trades, then we wire the slides to it.'
              : 'Roles sheet already has data — left untouched.'));
  } catch (_) {}
  return { seeded, total: getRoles_().length };
}

// Overwrite the Roles sheet from the editor UI (validated PIN in doPost).
function saveRoles_(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SN.ROLES);
  if (!sheet) { setupRolesSheet(); sheet = ss.getSheetByName(SN.ROLES); }

  const roles = Array.isArray(data.roles) ? data.roles : [];
  const clean = v => String(v == null ? '' : v).trim();
  const rows  = roles
    .map(r => [clean(r.team), clean(r.role), clean(r.member), clean(r.trade)])
    .filter(r => r[2]); // keep rows that have a member

  const last = sheet.getLastRow();
  if (last > 1) sheet.getRange(2, 1, last - 1, 4).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, 4).setValues(rows);

  return { saved: rows.length };
}

// ── Phase B: render grid TEXT from the Roles sheet ────────────────────────────
// Each cell's text box holds "Name\nTrade". We replace the name-run and
// trade-run in place (via replaceAllText scoped to the slide), which keeps each
// run's own styling (name bold, trade red). Only changed values are touched.
// textSlots = the cell text-box IDs in the SAME order as the team's Roles rows.
const COMMITTEE_TEXT_SLOTS = ['p10_i239', 'p10_i251', 'p10_i248', 'p10_i250', 'p10_i241'];

// A slot is either a box-id string (cell holds Name\nTrade) or
// { box, role:true } (cell holds Role\nName\nTrade). Slots are listed in the
// SAME order as the team's Roles rows. Defensive: a cell that can't be split
// safely (too few lines) is SKIPPED and reported — never mangled.
function renderTeamText_(teamName, slideId, slots) {
  const roles = getRoles_().filter(r => r.team.toLowerCase() === teamName.toLowerCase());
  const cur   = readDeckBoxText_();
  const norm  = s => (typeof s === 'string') ? { box: s, role: false } : { box: s.box, role: !!s.role };
  const requests = [];
  let changed = 0;
  const skipped = [];

  const swap = (oldT, newT) => {
    oldT = String(oldT || '').trim(); newT = String(newT || '').trim();
    if (oldT && newT && oldT !== newT) {
      requests.push({ replaceAllText: { containsText: { text: oldT, matchCase: true }, replaceText: newT, pageObjectIds: [slideId] } });
      changed++;
    }
  };

  slots.forEach((raw, i) => {
    const slot = norm(raw);
    const role = roles[i];
    if (!role) return;                                          // fewer members than cells
    // Cells break lines with \n OR a vertical-tab soft return (Shift+Enter in
    // Slides). Split on every line-break type, drop blank lines.
    const curText = String(cur[slot.box] || '');
    const parts   = curText.split(/[\n\r\v\f\u0085\u2028\u2029]+/).map(p => p.trim()).filter(p => p.length);
    const need    = slot.role ? 3 : 2;                          // role cells need Role/Name/Trade
    if (parts.length < need) { skipped.push(slot.box); return; } // unsafe to split — leave it

    let oldRole = '', oldName = '', oldTrade = '';
    if (slot.role) { oldRole = parts[0]; oldName = parts[1]; oldTrade = parts.slice(2).join(' '); }
    else           { oldName = parts[0]; oldTrade = parts.slice(1).join(' '); }

    if (slot.role) swap(oldRole, role.role);
    swap(oldName,  role.member);
    swap(oldTrade, role.trade);
  });

  if (requests.length) slidesBatchUpdate_(requests);
  return { changed, skipped, team: teamName };
}

function renderCommitteeText_() {
  return renderTeamText_('Membership Committee', COMMITTEE.slideId, COMMITTEE_TEXT_SLOTS);
}

// Visitor Host grid (slide 13 / p21). 6 cells in Roles-row order; Kay Tan's
// cell carries a "Visitor Orientation" role label (role:true).
const VISITOR_HOST_SLOTS = [
  'p21_i415',                              // Sandy Au
  'g3da1b38dbc6_0_2',                      // Jia Zheng Lee  (slide: "Lee Jia Zheng")
  'g3da1b38dbc6_0_9',                      // Ivan Ang
  'g3da1b38dbc6_0_6',                      // Daniel Yen
  { box: 'g3d3588e3aa5_0_17', role: true },// Kay Tan — Visitor Orientation
  'g3da1b38dbc6_0_11',                     // Yu Xi Kuek  (slide: "Kuek Yu Xi")
];
function renderVisitorHostText_() {
  return renderTeamText_('Visitor Host', 'p21', VISITOR_HOST_SLOTS);
}

// Render every wired grid team's text from the Roles sheet.
function renderAllGridText_() {
  return {
    committee:    renderCommitteeText_(),
    visitorHost:  renderVisitorHostText_(),
  };
}

// ── Phase C prep: verify each member has a usable headshot in Drive ───────────
// Matches Roles members to files in the Headshots folder by name-token overlap,
// so we know BEFORE building whether any photo is missing/misnamed.
function checkHeadshots_(teamName) {
  const roles  = getRoles_().filter(r => teamName ? r.team.toLowerCase() === teamName.toLowerCase() : true);
  const folder = DriveApp.getFolderById(HEADSHOTS_FOLDER_ID);
  const files  = [];
  const it     = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); files.push({ id: f.getId(), name: f.getName() }); }

  const tokens  = s => String(s).toLowerCase().replace(/[^a-z ]/g, ' ').split(/\s+/).filter(Boolean);
  const letters = s => String(s).toLowerCase().replace(/\.[a-z0-9]+$/, '').replace(/[^a-z]/g, '');
  const sortL   = s => letters(s).split('').sort().join('');

  const matches = roles.map(r => {
    const mtok = tokens(r.member), mL = letters(r.member), mS = sortL(r.member);
    let best = null, bestScore = 0, how = '';
    files.forEach(f => {
      const ftok = tokens(f.name), fL = letters(f.name), fS = sortL(f.name);
      let score = 0, h = '';
      const overlap = mtok.filter(t => ftok.includes(t)).length;
      if (overlap >= Math.min(2, mtok.length))       { score = 4; h = 'tokens'; }   // e.g. "Kevin Phua"
      else if (fL.includes(mL) || mL.includes(fL))   { score = 3; h = 'substr'; }    // "Jia Zheng" ⊂ "Jia Zheng Lee"
      else if (fS === mS)                            { score = 2; h = 'anagram'; }   // "ShuHui Zhao" ↔ "Zhao Shu Hui"
      else if (overlap === 1 && mtok.length >= 2)    { score = 1; h = 'onetoken'; }  // surname-only (weak)
      if (score > bestScore) { bestScore = score; best = f; how = h; }
    });
    const found = bestScore >= 2;   // accept tokens/substr/anagram; ignore weak one-token
    return { member: r.member, found, file: found ? best.name : null, fileId: found ? best.id : null, how: found ? how : (best ? 'weak:' + how : 'none') };
  });

  return {
    team:      teamName || 'ALL',
    total:     matches.length,
    withPhoto: matches.filter(m => m.found).length,
    missing:   matches.filter(m => !m.found).map(m => m.member),
    folderFileCount: files.length,
    allFiles:  files.map(f => f.name).sort(),
    matches,
  };
}

// ── Phase C: photo insertion (de-risk test) ──────────────────────────────────
// Share a Drive file link-viewable and return a Slides-fetchable image URL.
function shareAndThumbUrl_(fileId) {
  const f = DriveApp.getFileById(fileId);
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { logError_('share', e); }
  return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1600';
}

// ── Support Leadership pages ─────────────────────────────────────────────────
// Two generated pages (slide 14 + a duplicate of it). Photos come from Drive,
// text cells from the Roles sheet, and every position is COMPUTED — nothing on
// these pages is hand-placed, so the layout survives roster changes.
const SUPPORT_SLIDE_ID = 'g3da1b38dbc6_0_19';  // slide 14 — page 1

const SUPPORT_GRID = {
  // Cell width matches slide 11 (2115600 = 0.7052 × 3000000), so trade names
  // wrap the same way they do on the Membership Committee slide.
  textSx: 0.7052, textSy: 0.32, textW: 3000000,
};

function suppKey_(n) { return String(n).trim().toLowerCase(); }

// Resolve the members + their Drive photo IDs, in Roles order.
function suppRoster_() {
  const roles = getRoles_().filter(r => r.team.toLowerCase() === 'support leadership');
  const chk   = checkHeadshots_('Support Leadership');
  const photo = {};
  chk.matches.forEach(m => { if (m.fileId) photo[suppKey_(m.member)] = m.fileId; });
  const missing = roles.filter(r => !photo[suppKey_(r.member)]).map(r => r.member);
  return { roles, photo, missing, allFiles: chk.allFiles };
}

// { objectId: {w,h} } native sizes of every element on a page.
function readSlideElementSizes_(slideId) {
  const url  = 'https://slides.googleapis.com/v1/presentations/' + PRESENTATION_ID +
               '/pages/' + slideId + '?fields=' + encodeURIComponent('pageElements(objectId,size)');
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  const map = {};
  if (resp.getResponseCode() !== 200) return map;
  (JSON.parse(resp.getContentText()).pageElements || []).forEach(pe => {
    if (pe.size && pe.size.width && pe.size.height) map[pe.objectId] = { w: pe.size.width.magnitude, h: pe.size.height.magnitude };
  });
  return map;
}

// Replace a cell's Role/Name/Trade lines in place, preserving each line's style.
// rawTmpl = the cell's current text (with line breaks). Offsets come from the
// actual string, so any line-break type works. Edits run last→first so indices
// stay valid. Returns the batchUpdate requests (empty if <3 lines).
function styledCellReplace_(cellId, rawTmpl, newR, newN, newT) {
  const raw  = String(rawTmpl).replace(/[\n\r\v\f]+$/, '');
  const segs = [];
  let last = 0, m;
  const re = /[\n\r\v\f]+/g;
  while ((m = re.exec(raw)) !== null) { segs.push({ text: raw.slice(last, m.index), start: last }); last = m.index + m[0].length; }
  segs.push({ text: raw.slice(last), start: last });
  if (segs.length < 3) return [];

  const reqs = [];
  const rep = (seg, neu) => {
    neu = String(neu || '').trim();
    if (!seg.text || !neu || seg.text === neu) return;
    reqs.push({ insertText: { objectId: cellId, insertionIndex: seg.start, text: neu } });
    reqs.push({ deleteText: { objectId: cellId, textRange: {
      type: 'FIXED_RANGE', startIndex: seg.start + neu.length, endIndex: seg.start + neu.length + seg.text.length } } });
  };
  rep(segs[2], newT);   // trade  (last first)
  rep(segs[1], newN);   // name
  rep(segs[0], newR);   // role
  return reqs;
}

// ── Support Leadership: max 5 per slide ──────────────────────────────────────
// 9 members on one slide reads too dense. Members 1–5 (Roles-sheet order) stay
// on slide 14; the rest go on a DUPLICATE of it, so page 2 inherits slide 14's
// exact styling instead of being hand-positioned. Reorder rows in the Roles
// editor to change who lands on which page.
const SUPP_PER_SLIDE = 5;
const SUPP_SLIDE2    = 'supp2';
const SUPP_TITLE_1   = 'g3da1b38dbc6_0_29';
// Slide 14's 5 text cells, in Roles-row order.
const SUPP_TEXT_1 = ['g3da1b38dbc6_0_20', 'g3da1b38dbc6_0_33', 'g3da1b38dbc6_0_21',
                     'g3d3588e3aa5_0_19', 'g3d3588e3aa5_0_20'];
// ── Layout engine ────────────────────────────────────────────────────────────
// Rows: 5 members read better as 3 over 2 than as one long row of 5.
// A row of 4 still reads fine unbroken, so it stays on one line.
function suppRowSplit_(n) {
  if (n <= 4) return [n];
  return [Math.ceil(n / 2), Math.floor(n / 2)];   // 5 → [3, 2]
}

// ── Sizing, matched to slide 11 (Membership Committee) ──
// Measured off that slide: photo box 1563000×1721285 (aspect 0.908), column
// pitch 2324245 = 1.487 × photo width, cell width 2115600. Spreading columns
// evenly across the full slide width instead (the old rule) gave a pitch ratio
// of 3.07 — hence the gaps that read as too far apart.
const SUPP_ASPECT       = 0.908;     // photo box width : height
const SUPP_PITCH_RATIO  = 1.487;     // column pitch as a multiple of photo width
const SUPP_SIDE_MARGIN  =  700000;   // min slide edge → outer photo edge
// Photo height by row count. One row gets a noticeably bigger box, otherwise a
// 4-member slide is mostly empty space with small photos.
const SUPP_PHOTO_H_1ROW = 2150000;
const SUPP_PHOTO_H_2ROW = 1720000;

function suppPhotoBox_(rowCount) {
  const h = rowCount > 1 ? SUPP_PHOTO_H_2ROW : SUPP_PHOTO_H_1ROW;
  return { w: Math.round(h * SUPP_ASPECT), h: h };
}

// Column centres per row. Every row uses the SAME pitch (set by the widest row)
// and is centred on the slide, so a short row nests symmetrically under a long
// one. The preferred pitch is tightened only if the widest row would overflow.
function suppColumns_(split, photoW) {
  const maxK = Math.max.apply(null, split);
  let pitch  = photoW * SUPP_PITCH_RATIO;
  if (maxK > 1) {
    const cap = (SLIDE_W - 2 * SUPP_SIDE_MARGIN - photoW) / (maxK - 1);
    if (cap < pitch) pitch = cap;
  }
  const midX = SLIDE_W / 2;
  return split.map(k => Array.from({ length: k },
    (_, i) => Math.round(midX + (i - (k - 1) / 2) * pitch)));
}

// ── Vertical framing ──
// The inherited layout anchored the row near the top, leaving a large dead band
// below. Instead the whole photo+text block is centred in the space under the
// title bar. The cell's BOX is 960000 EMU tall but only ~4 lines actually
// render, so the block is measured by rendered height, not box height.
const SUPP_SLIDE_H        = 6858000;   // slide canvas height
const SUPP_TITLE_BOTTOM   = 1032725;   // bottom edge of the title box
const SUPP_TEXT_RENDER_H  =  750000;   // ~3-4 rendered lines of Role/Name/Trade
const SUPP_PHOTO_TEXT_GAP =   90000;   // photo bottom → cell top
const SUPP_ROW_GAP        =  280000;   // breathing room between the two rows

// Top-y for each row's photo band and text band.
function suppBands_(rowCount, photoH) {
  const rowBlock = photoH + SUPP_PHOTO_TEXT_GAP + SUPP_TEXT_RENDER_H;
  const avail    = SUPP_SLIDE_H - SUPP_TITLE_BOTTOM;
  const total    = rowCount * rowBlock + (rowCount - 1) * SUPP_ROW_GAP;
  const top      = SUPP_TITLE_BOTTOM + (avail - total) / 2;
  return Array.from({ length: rowCount }, (_, r) => {
    const py = Math.round(top + r * (rowBlock + SUPP_ROW_GAP));
    return { photoY: py, textY: py + photoH + SUPP_PHOTO_TEXT_GAP };
  });
}

// Position a list of photo+cell pairs on one slide, laid out for pairs.length.
// Taking an explicit LIST (rather than an index range) is what lets absence
// handling drop a member and re-flow the remainder into a tighter row.
// Photo scale comes from each image's NATIVE size so every box ends up identical.
//
// urlFor(j) (optional) → a Drive image URL for the j-th pair. When supplied the
// photo is re-cropped AFTER its transform, so CENTER_CROP measures the NEW box.
// Needed whenever the box ASPECT changes: a crop baked for one aspect gets
// stretched if the box is later reshaped, which distorts faces.
function suppPlaceAll_(pairs, sizes, urlFor) {
  const G     = SUPPORT_GRID;
  const split = suppRowSplit_(pairs.length);
  const box   = suppPhotoBox_(split.length);
  const cols  = suppColumns_(split, box.w);
  const bands = suppBands_(split.length, box.h);
  const reqs  = [];
  let k = 0;
  split.forEach((rowN, r) => {
    for (let c = 0; c < rowN; c++, k++) {
      const ids = pairs[k];
      if (!ids || !ids.img || !ids.txt) continue;
      const cx = cols[r][c], band = bands[r];
      const nz = sizes[ids.img] || box;
      reqs.push({ updatePageElementTransform: { objectId: ids.img,
        transform: { scaleX: box.w / nz.w, scaleY: box.h / nz.h,
                     translateX: Math.round(cx - box.w / 2), translateY: band.photoY, unit: 'EMU' },
        applyMode: 'ABSOLUTE' } });
      if (urlFor) {
        const u = urlFor(k);
        if (u) reqs.push({ replaceImage: { imageObjectId: ids.img, url: u, imageReplaceMethod: 'CENTER_CROP' } });
      }
      const tw = (sizes[ids.txt] && sizes[ids.txt].w) ? sizes[ids.txt].w : G.textW;
      reqs.push({ updatePageElementTransform: { objectId: ids.txt,
        transform: { scaleX: G.textSx, scaleY: G.textSy,
                     translateX: Math.round(cx - (tw * G.textSx) / 2), translateY: band.textY, unit: 'EMU' },
        applyMode: 'ABSOLUTE' } });
    }
  });
  return reqs;
}

// ── The two generated pages, in Roles-sheet order ────────────────────────────
// Each member's photo + cell live on a fixed page; only the columns within a
// page reflow. Elements are never moved between pages, so a member always
// appears where the deck expects them.
function suppPages_() {
  const roles = getRoles_().filter(r => r.team.toLowerCase() === 'support leadership');
  return [
    { slideId: SUPPORT_SLIDE_ID, members: roles.slice(0, SUPP_PER_SLIDE),
      img: i => 'suppimg' + i,         txt: i => SUPP_TEXT_1[i] },
    { slideId: SUPP_SLIDE2,      members: roles.slice(SUPP_PER_SLIDE),
      img: i => SUPP_SLIDE2 + 'i' + i, txt: i => SUPP_SLIDE2 + 't' + i },
  ].filter(p => p.members.length);
}

function suppOffscreen_(id) {
  return { updatePageElementTransform: { objectId: id,
    transform: { scaleX: 1, scaleY: 1, translateX: GRID_OFFSCREEN, translateY: GRID_OFFSCREEN, unit: 'EMU' },
    applyMode: 'ABSOLUTE' } };
}

// Names reach us from Weekly_Submissions and may be ordered differently to the
// Roles sheet ("Zhao Shu Hui" vs "ShuHui Zhao"), so compare on letters alone
// AND on sorted letters — the latter makes the comparison order-independent.
function suppNameKeys_(n) {
  const flat = String(n).toLowerCase().replace(/[^a-z]/g, '');
  return [flat, flat.split('').sort().join('')];
}
function suppAbsentKeys_(names) {
  const set = new Set();
  (names || []).forEach(n => suppNameKeys_(n).forEach(k => set.add(k)));
  return set;
}

// Skip/unskip a whole page (used when every member on it is away).
function suppSkipPage_(slideId, skip) {
  try {
    SlidesApp.openById(PRESENTATION_ID).getSlides().forEach(s => {
      if (s.getObjectId() === slideId && s.isSkipped() !== skip) s.setSkipped(skip);
    });
  } catch (e) { logError_('suppSkipPage_', e); }
}

// ── The one layout entry point ───────────────────────────────────────────────
// Everything — weekly attendance, manual relayout, restore — goes through here,
// so the automated path and the manual path can never drift apart.
//   absentKeys : names to hide (empty Set = everyone shown)
//   recrop     : re-CENTER_CROP visible photos (only needed if the box ASPECT changed)
function layoutSupportPages_(absentKeys, recrop) {
  let photo = null;
  if (recrop) {
    const r = suppRoster_();
    if (r.missing.length) return { error: 'missing photos', missing: r.missing, allFiles: r.allFiles };
    photo = r.photo;
  }

  const out = { hidden: [], pages: [] };
  suppPages_().forEach(p => {
    const sizes = readSlideElementSizes_(p.slideId);
    const present = [], away = [];
    p.members.forEach((m, i) => {
      const isAway = suppNameKeys_(m.member).some(k => absentKeys.has(k));
      (isAway ? away : present).push(i);
      if (isAway) out.hidden.push(m.member);
    });

    const reqs = [];
    away.forEach(i => { reqs.push(suppOffscreen_(p.img(i))); reqs.push(suppOffscreen_(p.txt(i))); });
    if (present.length) {
      const pairs  = present.map(i => ({ img: p.img(i), txt: p.txt(i) }));
      const urlFor = photo
        ? (j => shareAndThumbUrl_(photo[suppKey_(p.members[present[j]].member)]))
        : null;
      reqs.push.apply(reqs, suppPlaceAll_(pairs, sizes, urlFor));
    }

    const code = reqs.length ? slidesBatchUpdate_(reqs) : 200;
    // A page with nobody left on it is skipped rather than shown empty.
    suppSkipPage_(p.slideId, present.length === 0);
    out.pages.push({ slide: p.slideId, shown: present.length, away: away.length,
                     split: present.length ? suppRowSplit_(present.length) : [], code });
  });
  return out;
}

// Hide the given absent members and re-flow the rest. Called by the weekly run.
function applySupportAttendance_(absentNames) {
  return layoutSupportPages_(suppAbsentKeys_(absentNames), false);
}

// Show everyone again, freshly laid out. Also the manual "fix the layout" call.
function relayoutSupport_(recrop) {
  return layoutSupportPages_(new Set(), !!recrop);
}

// Recreate every Support photo from Drive, then place it. Use after replacing or
// re-cropping headshots in Drive. Images are RECREATED (not replaced) so each
// element reports its true native size, which the scale maths depends on.
function rebuildSupportPhotos_() {
  const { photo, missing, allFiles } = suppRoster_();
  if (missing.length) return { error: 'missing photos', missing, allFiles };

  const out = { pages: [] };
  suppPages_().forEach(p => {
    p.members.forEach((r, i) => {
      try { slidesBatchUpdate_([{ deleteObject: { objectId: p.img(i) } }]); } catch (e) {}
    });

    const createCode = slidesBatchUpdate_(p.members.map((r, i) => ({ createImage: {
      objectId: p.img(i),
      url: shareAndThumbUrl_(photo[suppKey_(r.member)]),
      elementProperties: { pageObjectId: p.slideId,
        transform: { scaleX: 1, scaleY: 1, translateX: 100000, translateY: 100000, unit: 'EMU' } },
    } })));

    const sizes  = readSlideElementSizes_(p.slideId);
    const pairs  = p.members.map((r, i) => ({ img: p.img(i), txt: p.txt(i) }));
    const urlFor = j => shareAndThumbUrl_(photo[suppKey_(p.members[j].member)]);
    const placeCode = slidesBatchUpdate_(suppPlaceAll_(pairs, sizes, urlFor));

    out.pages.push({ slide: p.slideId, members: p.members.length, createCode, placeCode });
  });
  return out;
}

function rebuildSupportSlides_() {
  const roles = getRoles_().filter(r => r.team.toLowerCase() === 'support leadership');
  if (!roles.length) return { error: 'no Support Leadership roles' };
  if (roles.length > 2 * SUPP_PER_SLIDE) return { error: 'more than 10 members — a 3rd slide is needed', count: roles.length };

  const { photo, missing, allFiles } = suppRoster_();
  if (missing.length) return { error: 'missing photos', missing, allFiles };

  const report = { page1: [], page2: [], removed: [] };

  // ── Page 1 — drop the 2nd-row overflow built earlier, keep the first 5 ──
  const drop = [];
  for (let i = 5; i < 9; i++) drop.push({ deleteObject: { objectId: 'suppimg' + i } });
  for (let i = 6; i <= 9; i++) drop.push({ deleteObject: { objectId: 'supptext' + i } });
  drop.forEach(req => { try { slidesBatchUpdate_([req]); } catch (e) {} });

  const first = roles.slice(0, SUPP_PER_SLIDE);
  const cur   = readDeckBoxText_();
  first.forEach((r, i) => {
    const reqs = styledCellReplace_(SUPP_TEXT_1[i], String(cur[SUPP_TEXT_1[i]] || ''), r.role, r.member, r.trade);
    if (reqs.length) slidesBatchUpdate_(reqs);
    report.page1.push(r.member);
  });
  const pairs1 = first.map((r, i) => ({ img: 'suppimg' + i, txt: SUPP_TEXT_1[i] }));
  const url1   = j => shareAndThumbUrl_(photo[suppKey_(first[j].member)]);
  const sizes1 = readSlideElementSizes_(SUPPORT_SLIDE_ID);
  report.split1   = suppRowSplit_(first.length);
  report.pos1Code = slidesBatchUpdate_(suppPlaceAll_(pairs1, sizes1, url1));

  // ── Page 2 — a fresh duplicate of page 1, trimmed to the overflow members ──
  const rest = roles.slice(SUPP_PER_SLIDE);
  if (rest.length) {
    try { slidesBatchUpdate_([{ deleteObject: { objectId: SUPP_SLIDE2 } }]); } catch (e) {}

    const idMap = {};
    idMap[SUPPORT_SLIDE_ID] = SUPP_SLIDE2;
    idMap[SUPP_TITLE_1]     = SUPP_SLIDE2 + 'title';
    for (let i = 0; i < SUPP_PER_SLIDE; i++) {
      idMap['suppimg' + i]  = SUPP_SLIDE2 + 'i' + i;
      idMap[SUPP_TEXT_1[i]] = SUPP_SLIDE2 + 't' + i;
    }
    report.dupCode = slidesBatchUpdate_([{ duplicateObject: { objectId: SUPPORT_SLIDE_ID, objectIds: idMap } }]);

    // Remove the copied cells page 2 doesn't need.
    for (let i = rest.length; i < SUPP_PER_SLIDE; i++) {
      [SUPP_SLIDE2 + 'i' + i, SUPP_SLIDE2 + 't' + i].forEach(id => {
        try { slidesBatchUpdate_([{ deleteObject: { objectId: id } }]); } catch (e) {}
      });
    }

    // Swap in the overflow members' text. Photos are re-cropped inside the
    // placement pass below, so CENTER_CROP measures the FINAL box, not the
    // inherited one.
    const cur2 = readDeckBoxText_();
    rest.forEach((r, i) => {
      const t = SUPP_SLIDE2 + 't' + i;
      const reqs = styledCellReplace_(t, String(cur2[t] || ''), r.role, r.member, r.trade);
      if (reqs.length) slidesBatchUpdate_(reqs);
      report.page2.push(r.member);
    });

    const pairs2 = rest.map((r, i) => ({ img: SUPP_SLIDE2 + 'i' + i, txt: SUPP_SLIDE2 + 't' + i }));
    const url2   = j => shareAndThumbUrl_(photo[suppKey_(rest[j].member)]);
    const sizes2 = readSlideElementSizes_(SUPP_SLIDE2);
    report.split2   = suppRowSplit_(rest.length);
    report.pos2Code = slidesBatchUpdate_(suppPlaceAll_(pairs2, sizes2, url2));
  }

  return report;
}

// ── Title slide date — auto-set to the upcoming Tuesday meeting ───────────────
// Meetings are Tuesday mornings. This writes the coming Tuesday's date into the
// title-slide date box (self-healing: replaces the old date, keeps its styling).
const TITLE_DATE_BOX = 'p2_i158';   // Slide 1 date line — currently "21st July 2026"

function ordinal_(n) {
  const s = ['th', 'st', 'nd', 'rd'], v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}

// The coming meeting Tuesday: today if today is Tuesday, else the next Tuesday.
function upcomingTuesday_() {
  const now = new Date();
  const add = (2 - now.getDay() + 7) % 7;   // 0=Sun,1=Mon,2=Tue → days until Tuesday
  return new Date(now.getFullYear(), now.getMonth(), now.getDate() + add);
}

function meetingDateText_(d) {
  const tz = Session.getScriptTimeZone();
  return ordinal_(d.getDate()) + ' ' + Utilities.formatDate(d, tz, 'MMMM yyyy');  // e.g. "28th July 2026"
}

// Update the title-slide date to the upcoming Tuesday. Only writes if it changed.
function updateTitleDate_() {
  const newText = meetingDateText_(upcomingTuesday_());
  const cur     = String(readDeckBoxText_()[TITLE_DATE_BOX] || '');
  const oldCore = cur.replace(/\n+$/, '');
  if (oldCore.trim() === newText.trim()) return { changed: false, date: newText };

  const requests = [{ insertText: { objectId: TITLE_DATE_BOX, insertionIndex: 0, text: newText } }];
  if (oldCore.length) {
    requests.push({ deleteText: { objectId: TITLE_DATE_BOX, textRange: {
      type: 'FIXED_RANGE', startIndex: newText.length, endIndex: newText.length + oldCore.length,
    } } });
  }
  slidesBatchUpdate_(requests);
  return { changed: true, date: newText, was: oldCore };
}

// Menu-callable wrapper.
function updateMeetingDate() {
  const r = updateTitleDate_();
  try { SpreadsheetApp.getUi().alert('📅 Title date set to: ' + r.date + (r.changed ? '' : ' (already current)')); } catch (_) {}
  return r;
}

// ── Slide 19 — Core Value Sharing presenter (Hub-driven) ─────────────────────
// The Hub "Core Value Sharing" page lets the MC pick the presenter + which of
// the 7 BNI core values they're sharing. This writes name+trade, the core value
// in the centre box, and swaps the photo — all on slide 19. Self-healing: it
// replaces whatever is there now, so it can be re-run any week.
const CORE_VALUE_SLIDE     = 'g3e794ddfa92_1_0';   // slide 19
const CORE_VALUE_NAME_BOX  = 'g3e794ddfa92_1_4';   // "Name\nTrade"
const CORE_VALUE_QUOTE_BOX = 'g3e794ddfa92_1_5';   // centre box → the core value
const CORE_VALUE_PHOTO     = 'g3f56d718e25_0_28';  // presenter photo
const CORE_VALUES = ['Givers Gain', 'Lifelong Learning', 'Traditions + Innovation',
                     'Positive Attitude', 'Building Relationships', 'Accountability', 'Recognition'];

// Replace a box's lines in place, keeping each line's own styling. newLines[i]
// replaces the i-th line; a blank/identical new line is skipped (old kept).
// Edits run last→first so earlier offsets stay valid.
function setCellLinesInPlace_(cellId, rawTmpl, newLines) {
  const raw  = String(rawTmpl).replace(/[\n\r\v\f]+$/, '');
  const segs = [];
  let last = 0, m;
  const re = /[\n\r\v\f]+/g;
  while ((m = re.exec(raw)) !== null) { segs.push({ text: raw.slice(last, m.index), start: last }); last = m.index + m[0].length; }
  segs.push({ text: raw.slice(last), start: last });

  const reqs = [];
  for (let i = Math.min(segs.length, newLines.length) - 1; i >= 0; i--) {
    const seg = segs[i], neu = String(newLines[i] == null ? '' : newLines[i]).trim();
    if (!seg.text || !neu || seg.text === neu) continue;
    reqs.push({ insertText: { objectId: cellId, insertionIndex: seg.start, text: neu } });
    reqs.push({ deleteText: { objectId: cellId, textRange: {
      type: 'FIXED_RANGE', startIndex: seg.start + neu.length, endIndex: seg.start + neu.length + seg.text.length } } });
  }
  return reqs;
}

// Find a member's headshot fileId in the Headshots folder (name-token match).
// Works for ANY member, not only those in the Roles sheet.
function findHeadshotFileId_(name) {
  const folder = DriveApp.getFolderById(HEADSHOTS_FOLDER_ID);
  const files  = [];
  const it     = folder.getFiles();
  while (it.hasNext()) { const f = it.next(); files.push({ id: f.getId(), name: f.getName() }); }

  const tokens  = s => String(s).toLowerCase().replace(/[^a-z ]/g, ' ').split(/\s+/).filter(Boolean);
  const letters = s => String(s).toLowerCase().replace(/\.[a-z0-9]+$/, '').replace(/[^a-z]/g, '');
  const sortL   = s => letters(s).split('').sort().join('');
  const mtok = tokens(name), mL = letters(name);

  let best = null, bestScore = 0;
  files.forEach(f => {
    const ftok = tokens(f.name), fL = letters(f.name);
    let score = 0;
    const overlap = mtok.filter(t => ftok.includes(t)).length;
    if (overlap >= Math.min(2, mtok.length))     score = 4;   // both name tokens present
    else if (fL.includes(mL) || mL.includes(fL)) score = 3;   // substring
    else if (sortL(f.name) === sortL(name))      score = 2;   // same letters, any order
    if (score > bestScore) { bestScore = score; best = f; }
  });
  return bestScore >= 2 ? best.id : null;
}

// Serve the 7 core values + roster to the Hub page.
function getCoreValueOptions_() {
  return jsonOk_({ coreValues: CORE_VALUES, members: getRoster_() });
}

// Write the Core Value Sharing presenter onto slide 19.
function saveCoreValue_(data) {
  const member    = String(data.member || '').trim();
  const coreValue = String(data.coreValue || '').trim();
  if (!member)    return { error: 'no member selected' };
  if (!coreValue) return { error: 'no core value selected' };

  const trade = tradeByMember_()[member.toLowerCase()] || '';
  const cur   = readDeckBoxText_();

  const reqs = []
    .concat(setCellLinesInPlace_(CORE_VALUE_NAME_BOX,  String(cur[CORE_VALUE_NAME_BOX]  || ''), [member, trade]))
    .concat(setCellLinesInPlace_(CORE_VALUE_QUOTE_BOX, String(cur[CORE_VALUE_QUOTE_BOX] || ''), [coreValue]));
  const textCode = reqs.length ? slidesBatchUpdate_(reqs) : 200;

  let photoCode = null, photoNote = '';
  const fileId = findHeadshotFileId_(member);
  if (fileId) {
    photoCode = slidesBatchUpdate_([{ replaceImage: {
      imageObjectId: CORE_VALUE_PHOTO, url: shareAndThumbUrl_(fileId), imageReplaceMethod: 'CENTER_CROP' } }]);
  } else {
    photoNote = 'no headshot in Drive for "' + member + '" — photo left unchanged';
  }

  return { ok: true, member, trade, coreValue, textCode, photoCode, photoNote };
}

// ── Utility ───────────────────────────────────────────────────────────────────
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
  console.error('[BNI %s] %s\n%s', ctx, err.message, err.stack || '');
}

// ── Sheets UI menu ────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BNI Empower')
    .addItem('📋 Check missing submissions',        'checkMissingNow')
    .addItem('🗑️  Clear Weekly Submissions',         'clearWeeklySubmissions')
    .addSeparator()
    .addItem('🙋 Check Attendance',                  'checkMissingAttendance')
    .addSeparator()
    .addItem('🎂 Test: send birthday reminders now', 'birthdayDailyCheck')
    .addSeparator()
    .addItem('📅 Add Meeting Column — Core Value',          'addMeetingColumn')
    .addItem('📅 Add Meeting Column — Network Education',   'addNetworkEducationColumn')
    .addItem('📅 Add Meeting Column — Featured Pres.',      'addFeaturedPresColumn')
    .addSeparator()
    .addItem('🏗️  Setup New Presentation Sheets',           'setupNewPresentationSheets')
    .addItem('🔄 Sync Member Rows',                         'syncMemberRows')
    .addItem('📊 Refresh Dashboard',                        'refreshDashboard')
    .addSeparator()
    .addItem('🎞️  Setup Intro Slides Sheet',                'setupSlideMapSheet')
    .addItem('🖼️  Update Slides from Attendance',            'updateSlidesFromAttendance')
    .addItem('↩️  Restore All Slides',                       'restoreAllSlides')
    .addItem('🔁 Reset Next Presenters (all present)',      'resetNextPresentersAllPresent')
    .addItem('📅 Update Title Date (next Tuesday)',         'updateMeetingDate')
    .addSeparator()
    .addItem('🗓️  Setup Roster Sheet',                      'setupRosterSheet')
    .addItem('⬆️  Push Roster to Sheets',                   'pushRosterToSheets')
    .addSeparator()
    .addItem('👥 Setup Roles Sheet (teams/roles)',          'setupRolesSheet')
    .addToUi();
}

// ── Clear Weekly Submissions ──────────────────────────────────────────────────
function clearWeeklySubmissions() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.alert(
    'Clear Weekly Submissions?',
    'This will delete all submission rows (header row is kept). This cannot be undone.',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.SUBMISSIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  ui.alert('✅ Weekly Submissions cleared.');
}

// ── Missing submissions check ─────────────────────────────────────────────────
function getMissingSubmitters_() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.SUBMISSIONS);

  const submitted = new Set();
  const lastRow   = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 3, lastRow - 1, 1).getValues().forEach(row => {
      const name = String(row[0]).trim().toLowerCase();
      if (name) submitted.add(name);
    });
  }

  return getRoster_().filter(name => !submitted.has(name.toLowerCase()));
}

function checkMissingNow() {
  const missing = getMissingSubmitters_();
  const ui      = SpreadsheetApp.getUi();

  if (missing.length === 0) {
    ui.alert('✅ All members have submitted their check-in for this week!');
    return;
  }

  ui.alert(
    '⚠️  Missing submissions (' + missing.length + ' members)',
    missing.join('\n'),
    ui.ButtonSet.OK
  );
}

function checkMissingAttendance() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.ACTUAL_ATTENDANCE);
  const ui    = SpreadsheetApp.getUi();
  const tz    = Session.getScriptTimeZone();

  if (!sheet) { ui.alert('Actual_Attendance sheet not found.'); return; }

  const lastCol = sheet.getLastColumn();
  if (lastCol < 4) { ui.alert('No meeting date columns found in Actual_Attendance.'); return; }

  const today = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy');
  const colHeaders = sheet.getRange(5, 4, 1, lastCol - 3).getValues()[0];
  let dateCol = -1;
  colHeaders.forEach((v, i) => { if (String(v).trim() === today) dateCol = 4 + i; });

  if (dateCol === -1) dateCol = lastCol;

  const lastRow = sheet.getLastRow();
  if (lastRow < 6) { ui.alert('No member rows found.'); return; }

  const names    = sheet.getRange(6, 2, lastRow - 5, 2).getValues();
  const statuses = sheet.getRange(6, dateCol, lastRow - 5, 1).getValues();

  const missing = [];
  names.forEach((row, i) => {
    const first   = String(row[0]).trim();
    const surname = String(row[1]).trim();
    const status  = String(statuses[i][0]).trim();
    if (first && !status) missing.push(first + (surname ? ' ' + surname : ''));
  });

  if (missing.length === 0) {
    ui.alert('✅ All members have signed in!');
    return;
  }

  ui.alert(
    missing.length + ' member' + (missing.length === 1 ? '' : 's') + ' not here yet',
    missing.join('\n'),
    ui.ButtonSet.OK
  );
}

// ── Auto email — called by Sunday 8pm trigger ─────────────────────────────────
function weeklyMissingReminder() {
  const missing = getMissingSubmitters_();
  if (missing.length === 0) {
    Logger.log('✅  All members submitted — no reminder sent.');
    return;
  }
  sendMissingEmail_(missing, false);
  Logger.log('✉️  Reminder sent to ' + VP_EMAIL + ' — ' + missing.length + ' missing.');
}

function sendMissingEmail_(missing) {
  const tz       = Session.getScriptTimeZone();
  const now      = new Date();
  const day      = now.getDay();
  const add      = day === 2 ? 0 : (9 - day) % 7;
  const tue      = new Date(now);
  tue.setDate(now.getDate() + add);
  const meetingDate = Utilities.formatDate(tue, tz, 'EEEE, dd MMM yyyy');

  const subject = 'BNI Empower — Missing check-ins for ' + meetingDate + ' (' + missing.length + ' members)';
  const nameList = missing.map((n, i) => (i + 1) + '. ' + n).join('\n');

  const body =
    'Hi ' + VP_NAME + ',\n\n' +
    'The following ' + missing.length + ' member' + (missing.length > 1 ? 's have' : ' has') +
    ' not submitted their weekly check-in for ' + meetingDate + ':\n\n' +
    nameList + '\n\n' +
    'Deadline was Sunday 8pm. Please follow up or update the slides accordingly.\n\n' +
    '— BNI Empower Check-In System';

  const htmlBody =
    '<p>Hi ' + VP_NAME + ',</p>' +
    '<p>The following <strong>' + missing.length + ' member' + (missing.length > 1 ? 's</strong> have' : '</strong> has') +
    ' not submitted their weekly check-in for <strong>' + meetingDate + '</strong>:</p>' +
    '<ol>' + missing.map(n => '<li>' + n + '</li>').join('') + '</ol>' +
    '<p>Deadline was Sunday 8pm. Please follow up or update the slides accordingly.</p>' +
    '<p style="color:#888;font-size:12px">— BNI Empower Check-In System</p>';

  MailApp.sendEmail({
    to:       VP_EMAIL,
    cc:       PRESIDENT_EMAIL,
    subject:  subject,
    body:     body,
    htmlBody: htmlBody,
  });
}

// ═════════════════════════════════════════════════════════════════════════════
// ── SLIDE AUTOMATION MODULE ───────────────────────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════
//
// RULES:
//   • Leadership Team slides (President, VP, Sec-Treas) — NEVER modified.
//     They always show regardless of attendance.
//   • Support Leadership pages (slide 14 + its generated page 2) — absent
//     members (even with a sub) are moved off-screen and the remaining members
//     re-flow into a tighter row:
//         ≤ 4 present → single centred row, larger photos
//         5 present   → 3 over 2
//     Members never move between pages, and a page with nobody left is skipped.
//     Membership comes from the Roles sheet, so no per-member config lives here.
//   • 30-sec intro slides:
//       ABSENT (no sub)  → slide moved to end of intro section
//       ABSENT with SUB  → slide kept in place (sub presents)
//       PRESENT          → no change
//
// ── Support Team ─────────────────────────────────────────────────────────────
// The Support Leadership pages are fully generated (see suppPages_ /
// layoutSupportPages_ above): membership comes from the Roles sheet, positions
// are computed, and absence is handled by re-flowing the remaining members.
// The old hand-maintained element registry that used to live here is gone — it
// pinned photo objectIds that the rebuild deleted, so every batch containing it
// failed silently and no attendance change reached the slide.

// Slide canvas (EMU)
const SLIDE_W = 12192000;

// Member 30-sec intro slide objectIds
const MEMBER_INTRO_SLIDES_DEFAULT = {
  'arun prasad':           'g3d498035fa2_1_55',
  'ben wong':              'g3d037e0c29a_0_62',
  'ben tee':               'g3d037e0c29a_0_255',
  'benjamin ng':           'g3d037e0c29a_0_78',
  'daniel yen':            'g3d037e0c29a_0_86',
  'deborah chueh':         'g3d037e0c29a_0_94',
  'ismail khamis':         'g3d037e0c29a_0_110',
  'ivan ang':              'g3d037e0c29a_0_118',
  'jaron chan':            'g3d037e0c29a_0_126',
  'jay tan':               'g3d037e0c29a_0_134',
  'jia zheng lee':         'g3d037e0c29a_0_142',
  'lee jia zheng':         'g3d037e0c29a_0_142',
  'junxian zhang':         'g3dcdba80422_1_3',
  'zhang junxian':         'g3dcdba80422_1_3',
  'joanne sooi':           'g3d037e0c29a_0_150',
  'kay tan':               'g3d037e0c29a_0_158',
  'kevin phua':            'g3d037e0c29a_0_166',
  'lawrence ku':           'g3d037e0c29a_0_174',
  'mark duma':             'g3d037e0c29a_0_182',
  'pamela lin':            'g3d037e0c29a_0_190',
  'rajivgandhi ponnusamy': 'g3dcdba80422_1_14',
  'sandy au':              'g3d037e0c29a_0_206',
  'zhao shuhui':           'g3d037e0c29a_0_222',
  'shuhui zhao':           'g3d037e0c29a_0_222',
  'yu xi kuek':            'g3d037e0c29a_0_238',
  'kuek yu xi':            'g3d037e0c29a_0_238',
  'zefirelli noordin':     'g3d037e0c29a_0_246',
  'zef':                   'g3d037e0c29a_0_246',
  'rachel teo':            'g3d037e0c29a_0_198',
  'wee khai pang':         'g3d037e0c29a_0_230',
  'pang wee khai':         'g3d037e0c29a_0_230',
  'wee khai':              'g3d037e0c29a_0_230',
  'iskons':                'g3d037e0c29a_0_214',
};

// ── Sheet-driven intro-slide map (falls back to the constants above) ─────────
// Intro_Slides lets officers correct a member's 30-sec slide ID without code.
// (The old Slide_Map tab is gone — Support Team box positions are computed now,
// so there is nothing left for it to override.)
let _slideMapCache = null;

function readSlideMap_() {
  if (_slideMapCache) return _slideMapCache;

  const result = { intros: MEMBER_INTRO_SLIDES_DEFAULT };

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const introSheet = ss.getSheetByName(SN.INTRO_SLIDES);
    if (introSheet && introSheet.getLastRow() >= 2) {
      const rows   = introSheet.getRange(2, 1, introSheet.getLastRow() - 1, 2).getValues();
      const intros = {};
      rows.forEach(r => {
        const member  = String(r[0]).trim().toLowerCase();
        const slideId = String(r[1]).trim();
        if (member && slideId) intros[member] = slideId;
      });
      if (Object.keys(intros).length) result.intros = intros;
    }
  } catch (err) {
    logError_('readSlideMap_', err);   // any failure → keep the safe defaults
  }

  _slideMapCache = result;
  return result;
}

function getMemberIntroSlides_() { return readSlideMap_().intros; }

/**
 * Creates (or refreshes) the Intro_Slides tab, seeded with the current
 * defaults. Existing edits are preserved: it is only seeded when empty.
 * Menu: BNI Empower → 🎞️ Setup Intro Slides Sheet
 */
function setupSlideMapSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let ui; try { ui = SpreadsheetApp.getUi(); } catch (_) { ui = null; }
  const notes = [];

  // ── Intro_Slides (30-sec intro slide IDs) ──
  let introSheet = ss.getSheetByName(SN.INTRO_SLIDES);
  if (!introSheet) introSheet = ss.insertSheet(SN.INTRO_SLIDES);

  introSheet.getRange(1, 1, 1, 2)
    .setValues([['Member (lowercase key / alias)', 'Intro Slide Object ID']])
    .setBackground('#1a5276').setFontColor('#ffffff')
    .setFontWeight('bold').setHorizontalAlignment('center');
  introSheet.setFrozenRows(1);
  introSheet.setColumnWidth(1, 230);
  introSheet.setColumnWidth(2, 220);

  if (introSheet.getLastRow() < 2) {
    const introRows = Object.entries(MEMBER_INTRO_SLIDES_DEFAULT).map(([m, id]) => [m, id]);
    if (introRows.length) introSheet.getRange(2, 1, introRows.length, 2).setValues(introRows);
    notes.push('Intro_Slides seeded with ' + introRows.length + ' entries.');
  } else {
    notes.push('Intro_Slides already has data — left untouched.');
  }

  _slideMapCache = null;  // force re-read on next slide update
  if (ui) ui.alert('✅ Intro slide map ready.\n\n' + notes.join('\n') +
    '\n\nEdit this tab any time; changes apply on the next slide update.');
}

// ── Read attendance from Weekly_Submissions ───────────────────────────────────
function getWeeklyAttendance_() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.SUBMISSIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const rows    = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const result  = {};

  rows.forEach(row => {
    const memberName = String(row[2]).trim();
    const attending  = String(row[3]).trim().toLowerCase();
    const subName    = String(row[5]).trim();
    if (!memberName) return;
    const key = memberName.toLowerCase();
    result[key] = { originalName: memberName, attending, hasSub: !!subName, subName };
  });

  return result;
}

// ── Main slide update function ────────────────────────────────────────────────
/**
 * Called automatically by the Sunday 9pm trigger, or manually from the
 * BNI Empower menu → "Update Slides from Attendance".
 */
function updateSlidesFromAttendance() {
  const attendance = getWeeklyAttendance_();

  if (Object.keys(attendance).length === 0) {
    const ui = SpreadsheetApp.getUi();
    if (ui) ui.alert('⚠️ No attendance submissions found in Weekly_Submissions. Slides not updated.');
    Logger.log('updateSlidesFromAttendance: no submissions found.');
    return;
  }

  // Determine absent members
  const absentNames = [];                // every absent member, any team
  const trulyAbsent = new Set();         // absent with NO sub (30-sec slide moves)
  const hasSubSet   = new Set();         // absent but has sub (30-sec slide stays)

  Object.entries(attendance).forEach(([normName, info]) => {
    if (info.attending !== 'no') return;
    absentNames.push(info.originalName || normName);
    if (info.hasSub) hasSubSet.add(normName); else trulyAbsent.add(normName);
  });

  Logger.log('Absent: ' + absentNames.join(', '));
  Logger.log('Truly absent (no sub): ' + [...trulyAbsent].join(', '));
  Logger.log('Absent with sub: ' + [...hasSubSet].join(', '));

  // Support Leadership pages — hide whoever is away, re-flow the rest.
  // Names that aren't on that team are simply ignored by the matcher.
  const supp = applySupportAttendance_(absentNames);
  Logger.log('Support pages: ' + JSON.stringify(supp));

  // Update 30-sec intro slides
  updateIntroSlides_(trulyAbsent);

  // Show confirmation if run manually from menu
  try {
    const ui = SpreadsheetApp.getUi();
    const movedNames = [...trulyAbsent].map(n => n.split(' ').map(w => w[0].toUpperCase() + w.slice(1)).join(' '));
    ui.alert(
      '✅ Slides updated!',
      (supp.hidden && supp.hidden.length
        ? 'Hidden from Support Leadership: ' + supp.hidden.join(', ') + '\n'
        : 'No Support Leadership changes.\n') +
      (movedNames.length ? 'Intro slides moved to end: ' + movedNames.join(', ') : 'No intro slides moved.'),
      ui.ButtonSet.OK
    );
  } catch(_) {
    // Running from trigger — no UI available, that's fine
  }
}

// ── Update 30-sec intro slides ────────────────────────────────────────────────
function updateIntroSlides_(trulyAbsent) {
  if (trulyAbsent.size === 0) {
    Logger.log('updateIntroSlides_: no slides to move.');
    return;
  }

  // Fetch current slide order
  const url  = 'https://slides.googleapis.com/v1/presentations/' + PRESENTATION_ID + '?fields=slides.objectId';
  const resp = UrlFetchApp.fetch(url, {
    method:  'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200) {
    Logger.log('updateIntroSlides_ ERROR fetching slides: ' + resp.getContentText());
    return;
  }

  const slideIds    = JSON.parse(resp.getContentText()).slides.map(s => s.objectId);
  const anchorSlide = 'p53';  // "Did We Miss Anyone" slide
  let   anchorPos   = slideIds.indexOf(anchorSlide);
  if (anchorPos === -1) anchorPos = 79;  // fallback

  const INTRO_SLIDES = getMemberIntroSlides_();
  trulyAbsent.forEach(normName => {
    // Find the slide ID for this member (try all alias keys)
    let slideId = INTRO_SLIDES[normName];
    if (!slideId) {
      // Try partial match
      for (const [key, sid] of Object.entries(INTRO_SLIDES)) {
        if (normName.includes(key) || key.includes(normName)) { slideId = sid; break; }
      }
    }
    if (!slideId) {
      Logger.log('updateIntroSlides_: no intro slide found for ' + normName);
      return;
    }

    const currentPos = slideIds.indexOf(slideId);
    if (currentPos === -1) {
      Logger.log('updateIntroSlides_: slide ' + slideId + ' not found in presentation');
      return;
    }

    const targetPos = anchorPos - 1;
    if (currentPos === targetPos) {
      Logger.log('updateIntroSlides_: ' + normName + ' already at correct position');
      return;
    }

    const batchUrl  = 'https://slides.googleapis.com/v1/presentations/' + PRESENTATION_ID + ':batchUpdate';
    const batchResp = UrlFetchApp.fetch(batchUrl, {
      method:      'post',
      contentType: 'application/json',
      headers:     { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload:     JSON.stringify({ requests: [{ updateSlidesPosition: { slideObjectIds: [slideId], insertionIndex: targetPos } }] }),
      muteHttpExceptions: true,
    });

    if (batchResp.getResponseCode() !== 200) {
      Logger.log('updateIntroSlides_ ERROR for ' + normName + ': ' + batchResp.getContentText());
    } else {
      Logger.log('updateIntroSlides_: moved ' + normName + ' intro slide to position ' + targetPos);
      // Update local index
      slideIds.splice(currentPos, 1);
      const adjustedTarget = currentPos < targetPos ? targetPos - 1 : targetPos;
      slideIds.splice(adjustedTarget, 0, slideId);
      anchorPos = slideIds.indexOf(anchorSlide);
      if (anchorPos === -1) anchorPos = 79;
    }
  });
}

// ── Restore all slides to the everyone-present state ─────────────────────────
/**
 * Puts every Support Leadership member back on their page and re-flows the
 * layout. Run at the start of a new week before applying fresh attendance.
 * Menu: BNI Empower → "Restore All Slides".
 */
function restoreAllSlides() {
  const r  = relayoutSupport_(false);
  const ok = !r.error && (r.pages || []).every(p => p.code === 200);
  const summary = r.error
    ? '❌ ' + r.error + (r.missing ? ': ' + r.missing.join(', ') : '')
    : (r.pages || []).map(p => p.slide + ': ' + p.shown + ' shown').join('\n');

  try {
    SpreadsheetApp.getUi().alert(
      (ok ? '✅ Support Leadership restored.\n\n' : '⚠️ Restore had problems.\n\n') + summary);
  } catch (_) {
    Logger.log('restoreAllSlides: ' + summary);
  }
  return r;
}

// ── Weekly trigger function (Sunday 9pm) ──────────────────────────────────────
/**
 * This is the master Sunday trigger function.
 * Assign this to a time-based trigger: Every Sunday at 9pm.
 * It runs AFTER the 8pm missing-submissions email, so attendance is finalised.
 *
 * Order of operations:
 *   1. Restore all slides to original state (clean slate)
 *   2. Apply this week's attendance to update slides
 *   3. Send missing submissions email (if not already sent by weeklyMissingReminder)
 */
function weeklySundayTrigger() {
  Logger.log('=== weeklySundayTrigger START ===');
  restoreAllSlides();
  Utilities.sleep(2000);  // brief pause to ensure restore completes
  updateSlidesFromAttendance();
  try { updateTitleDate_(); } catch (e) { Logger.log('updateTitleDate_ error: ' + e.message); }
  Logger.log('=== weeklySundayTrigger END ===');
}

// ═════════════════════════════════════════════════════════════════════════════
// ── BIRTHDAY MODULE (unchanged from original) ─────────────────────────────────
// ═════════════════════════════════════════════════════════════════════════════

function birthdayDailyCheck() {
  const tz      = Session.getScriptTimeZone();
  const today   = new Date();
  const members = getActiveMembersWithBirthdays_();

  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  if (tomorrow.getDate() === 1) {
    const nextMonth     = tomorrow.getMonth() + 1;
    const nextMonthName = Utilities.formatDate(tomorrow, tz, 'MMMM');
    const nextYear      = tomorrow.getFullYear();
    const babies        = members.filter(m => m.month === nextMonth);
    if (babies.length > 0) {
      sendMonthlyBirthdayEmail_(babies, nextMonthName, nextYear);
      Logger.log('🎂 Monthly birthday email sent for ' + nextMonthName + ' — ' + babies.length + ' member(s).');
    } else {
      Logger.log('No birthdays in ' + nextMonthName + ' — no monthly email sent.');
    }
  }

  const todayMonth = today.getMonth() + 1;
  const todayDay   = today.getDate();
  members
    .filter(m => m.month === todayMonth && m.day === todayDay)
    .forEach(m => {
      sendDayOfBirthdayEmail_(m);
      Logger.log('🎂 Day-of birthday email sent for ' + m.name);
    });
}

function getActiveMembersWithBirthdays_() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SN.ACTIVE_MEMBERS);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data   = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const result = [];

  data.forEach(row => {
    const fullName    = String(row[1]).trim();
    const birthdayRaw = row[4];

    if (!fullName || !birthdayRaw) return;

    const MONTH_NAMES = {
      jan:1, feb:2, mar:3, apr:4,  may:5,  jun:6,
      jul:7, aug:8, sep:9, oct:10, nov:11, dec:12,
    };

    let month, day, year;
    if (birthdayRaw instanceof Date) {
      month = birthdayRaw.getMonth() + 1;
      day   = birthdayRaw.getDate();
      year  = birthdayRaw.getFullYear();
    } else {
      const parts = String(birthdayRaw).trim().split(/\s+/);
      if (parts.length < 2) return;
      day   = parseInt(parts[0], 10);
      month = MONTH_NAMES[parts[1].toLowerCase().slice(0, 3)];
      year  = parts[2] ? parseInt(parts[2], 10) : null;
      if (!month) return;
    }

    if (!month || !day || month < 1 || month > 12 || day < 1 || day > 31) return;

    result.push({ name: fullName, month, day, year });
  });

  result.sort((a, b) => a.month !== b.month ? a.month - b.month : a.day - b.day);
  return result;
}

function getMCTeamEmails_() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('LT / ST');
  if (!sheet) { Logger.log('getMCTeamEmails_: "LT / ST" sheet not found.'); return ''; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return '';

  const emails = sheet.getRange(2, 3, lastRow - 1, 1).getValues()
    .map(row => String(row[0]).trim())
    .filter(e => e.includes('@'));

  return emails.join(',');
}

function sendMonthlyBirthdayEmail_(babies, monthName, year) {
  const to = getMCTeamEmails_();
  if (!to) { Logger.log('sendMonthlyBirthdayEmail_: no MC team emails found — email not sent.'); return; }

  const subject = 'BNI Empower — ' + monthName + ' ' + year + ' Birthday Members';
  const lines   = babies.map(m => m.day + ' ' + monthName + ' — ' + m.name);

  const body =
    'Hi!\n\n' +
    'Heads up — the following member' + (babies.length > 1 ? 's have' : ' has') +
    ' a birthday in ' + monthName + ' ' + year + ':\n\n' +
    lines.map((l, i) => (i + 1) + '. ' + l).join('\n') + '\n\n' +
    'Please prepare the birthday presentation and cake for the ' + monthName + ' physical meeting!\n\n' +
    '— BNI Empower Check-In System';

  const htmlLines = babies.map(m =>
    '<li><strong>' + m.day + ' ' + monthName + '</strong> — ' + m.name + '</li>'
  );

  const htmlBody =
    '<p>Hi!</p>' +
    '<p>Heads up — the following <strong>' + babies.length + ' member' +
    (babies.length > 1 ? 's</strong> have' : '</strong> has') +
    ' a birthday in <strong>' + monthName + ' ' + year + '</strong>:</p>' +
    '<ol>' + htmlLines.join('') + '</ol>' +
    '<p>Please prepare the birthday presentation and cake for the ' +
    monthName + ' physical meeting! 🎂🎉</p>' +
    '<p style="color:#888;font-size:12px">— BNI Empower Check-In System</p>';

  MailApp.sendEmail({ to, subject, body, htmlBody });
}

function sendDayOfBirthdayEmail_(member) {
  const to = getMCTeamEmails_();
  if (!to) { Logger.log('sendDayOfBirthdayEmail_: no MC team emails found — email not sent.'); return; }

  const subject = member.name + "'s Birthday Today!";

  const body =
    'Hi!\n\n' +
    "Today is " + member.name + "'s birthday! 🎂\n\n" +
    'Remember to wish ' + member.name + ' in the WhatsApp group today. 🎉\n\n' +
    '— BNI Empower Check-In System';

  const htmlBody =
    '<p>Hi!</p>' +
    '<p>Today is <strong>' + member.name + "'s</strong> birthday! 🎂</p>" +
    '<p>Remember to wish <strong>' + member.name + '</strong> in the WhatsApp group today. 🎉</p>' +
    '<p style="color:#888;font-size:12px">— BNI Empower Check-In System</p>';

  MailApp.sendEmail({ to, subject, body, htmlBody });
}

// ── Presentation Tracking Sheets (Core Value / Network Education / Featured Pres.) ──
function addMeetingColumnHelper_(sheetName, presLabel, finalizeFn) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  const ui    = SpreadsheetApp.getUi();

  if (!sheet) { ui.alert('Sheet not found: ' + sheetName); return; }

  const tz  = Session.getScriptTimeZone();
  const now = new Date();
  const daysUntilTue = (2 - now.getDay() + 7) % 7;
  const meetingDate  = new Date(now);
  meetingDate.setDate(now.getDate() + daysUntilTue);

  const type      = meetingDate.getDate() <= 7 ? 'Physical' : 'Online';
  const dateLabel = Utilities.formatDate(meetingDate, tz, 'dd MMM yyyy');

  const lastCol   = sheet.getLastColumn();
  const headerRow = lastCol >= 3
    ? sheet.getRange(2, 3, 1, lastCol - 2).getValues()[0]
    : [];

  if (headerRow.some(v => String(v).trim() === dateLabel)) {
    ui.alert('A column for ' + dateLabel + ' already exists.');
    return;
  }

  const newCol = lastCol + 1;

  sheet.getRange(1, newCol)
    .setValue(type)
    .setBackground('#c0392b')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setHorizontalAlignment('center');

  sheet.getRange(2, newCol)
    .setValue(dateLabel)
    .setFontWeight('bold')
    .setBackground('#962d22')
    .setFontColor('#ecf0f1')
    .setHorizontalAlignment('center');

  sheet.setColumnWidth(newCol, 100);

  const options = getRoster_().map(n => '<option value="' + n + '">').join('');
  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;padding:14px;margin:0">' +
    '<p style="margin:0 0 6px"><b>' + type + ' &mdash; ' + dateLabel + '</b></p>' +
    '<label style="font-size:13px">' + presLabel + ':</label><br>' +
    '<input id="n" list="m" placeholder="Start typing a name…" ' +
    '  style="width:100%;box-sizing:border-box;margin-top:5px;padding:5px;font-size:13px">' +
    '<datalist id="m">' + options + '</datalist>' +
    '<p style="font-size:11px;color:#888;margin:4px 0 10px">Leave blank if not confirmed yet</p>' +
    '<div style="text-align:right">' +
    '  <button onclick="submit()" style="padding:5px 18px">OK</button>' +
    '</div>' +
    '<script>' +
    'document.getElementById("n").addEventListener("keydown",function(e){if(e.key==="Enter")submit();});' +
    'function submit(){' +
    '  google.script.run' +
    '    .withSuccessHandler(function(){google.script.host.close()})' +
    '    .withFailureHandler(function(e){alert(e.message);google.script.host.close()})' +
    '    .' + finalizeFn + '(document.getElementById("n").value,' + newCol + ');' +
    '}' +
    '<\/script></body></html>'
  ).setWidth(360).setHeight(170);

  ui.showModalDialog(html, 'Who is presenting ' + presLabel + '?');
}

function finalizeMeetingColumn_(sheetName, presenterInput, newCol) {
  const ss       = SpreadsheetApp.openById(SHEET_ID);
  const sheet    = ss.getSheetByName(sheetName);
  const presName = (presenterInput || '').trim();
  if (!presName) return;

  const lastRow  = sheet.getLastRow();
  const nameData = sheet.getRange(3, 1, lastRow - 2, 2).getValues();
  let matched = false;
  nameData.forEach((row, i) => {
    const full  = (row[0] + ' ' + row[1]).trim().toLowerCase();
    const first = row[0].toString().toLowerCase();
    if (full === presName.toLowerCase() || first === presName.toLowerCase()) {
      sheet.getRange(3 + i, newCol).setValue('Y');
      matched = true;
    }
  });

  if (!matched) {
    SpreadsheetApp.getUi().alert('⚠️  "' + presName + '" not found in member list — Y not marked. Check the spelling.');
  }
}

// Core Value
function addMeetingColumn() {
  addMeetingColumnHelper_(SN.CORE_VALUE, 'Core Value Presenter', 'finalizeAddMeetingColumn');
}
function finalizeAddMeetingColumn(presenterInput, newCol) {
  finalizeMeetingColumn_(SN.CORE_VALUE, presenterInput, newCol);
}

// Network Education
function addNetworkEducationColumn() {
  addMeetingColumnHelper_(SN.NETWORK_EDU, 'Network Educator', 'finalizeAddNetworkEduColumn');
}
function finalizeAddNetworkEduColumn(presenterInput, newCol) {
  finalizeMeetingColumn_(SN.NETWORK_EDU, presenterInput, newCol);
}

// Featured Presentation
function addFeaturedPresColumn() {
  addMeetingColumnHelper_(SN.FEATURED_PRES, 'Featured Presenter', 'finalizeAddFeaturedPresColumn');
}
function finalizeAddFeaturedPresColumn(presenterInput, newCol) {
  finalizeMeetingColumn_(SN.FEATURED_PRES, presenterInput, newCol);
}

// ── Sync Member Rows ──────────────────────────────────────────────────────────
function syncMemberRows() {
  const ss     = SpreadsheetApp.openById(SHEET_ID);
  const ui     = SpreadsheetApp.getUi();
  const roster = getRoster_();
  const report = [];

  // nameCol: 1-based column where first name sits; startRow: first member row
  const targets = [
    { name: SN.CORE_VALUE,        nameCol: 1, startRow: 3 },
    { name: SN.NETWORK_EDU,       nameCol: 1, startRow: 3 },
    { name: SN.FEATURED_PRES,     nameCol: 1, startRow: 3 },
    { name: SN.ACTUAL_ATTENDANCE, nameCol: 2, startRow: 6 },
  ];

  targets.forEach(({ name: sheetName, nameCol, startRow }) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) { report.push(sheetName + ': not found'); return; }

    const lastRow  = sheet.getLastRow();
    const existing = new Set();
    if (lastRow >= startRow) {
      sheet.getRange(startRow, nameCol, lastRow - startRow + 1, 2).getValues().forEach(row => {
        const full = (String(row[0]).trim() + ' ' + String(row[1]).trim()).trim().toLowerCase();
        if (full) existing.add(full);
      });
    }

    let added = 0;
    roster.forEach(fullName => {
      if (existing.has(fullName.toLowerCase())) return;
      const spaceIdx = fullName.lastIndexOf(' ');
      const first    = spaceIdx > 0 ? fullName.slice(0, spaceIdx) : fullName;
      const surname  = spaceIdx > 0 ? fullName.slice(spaceIdx + 1) : '';
      const newRow   = sheet.getLastRow() + 1;
      sheet.getRange(newRow, nameCol, 1, 2).setValues([[first, surname]]);
      added++;
    });

    report.push(sheetName + ': ' + (added > 0 ? '+' + added + ' new member(s) added' : 'already in sync'));
  });

  ui.alert('✅ Sync complete\n\n' + report.join('\n'));
}

// ── Setup New Presentation Sheets ─────────────────────────────────────────────
function setupNewPresentationSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ui = SpreadsheetApp.getUi();
  const created = [];

  [SN.NETWORK_EDU, SN.FEATURED_PRES].forEach(sheetName => {
    if (ss.getSheetByName(sheetName)) return;
    const sheet = ss.insertSheet(sheetName);
    created.push(sheetName);

    sheet.getRange(1, 1).setValue('Meeting Type').setFontWeight('bold');
    sheet.getRange(2, 1).setValue('Date').setFontWeight('bold');
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 140);

    getRoster_().forEach((fullName, i) => {
      const spaceIdx = fullName.lastIndexOf(' ');
      const first   = spaceIdx > 0 ? fullName.slice(0, spaceIdx) : fullName;
      const surname = spaceIdx > 0 ? fullName.slice(spaceIdx + 1) : '';
      sheet.getRange(3 + i, 1).setValue(first);
      sheet.getRange(3 + i, 2).setValue(surname);
    });
  });

  if (!ss.getSheetByName(SN.DASHBOARD)) {
    ss.insertSheet(SN.DASHBOARD);
    created.push(SN.DASHBOARD);
    refreshDashboard(true);
  }

  if (created.length > 0) {
    ui.alert('✅ Created: ' + created.join(', ') + '\n\nUse "Refresh Dashboard" any time to update the summary.');
  } else {
    ui.alert('All presentation sheets already exist — no changes made.');
  }
}

// ── Dashboard ─────────────────────────────────────────────────────────────────
function refreshDashboard(suppressAlert) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let dashboard = ss.getSheetByName(SN.DASHBOARD);
  if (!dashboard) dashboard = ss.insertSheet(SN.DASHBOARD);

  const memberMap = {};
  getRoster_().forEach(fullName => {
    memberMap[fullName.toLowerCase()] = {
      name: fullName, cv: 0, ne: 0, fp: 0,
      cvLast: '', neLast: '', fpLast: '',
    };
  });

  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const todayCutoff = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1); // end of today

  function tally(sheetName, countKey, lastKey) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 3 || sheet.getLastColumn() < 3) return;

    const lastRow     = sheet.getLastRow();
    const numDateCols = sheet.getLastColumn() - 2;
    const names       = sheet.getRange(3, 1, lastRow - 2, 2).getValues();
    const dates       = sheet.getRange(2, 3, 1, numDateCols).getValues()[0];
    const marks       = sheet.getRange(3, 3, lastRow - 2, numDateCols).getValues();

    names.forEach((row, i) => {
      const full = (String(row[0]).trim() + ' ' + String(row[1]).trim()).trim().toLowerCase();
      if (!full || !memberMap[full]) return;
      marks[i].forEach((val, j) => {
        if (String(val).trim().toUpperCase() !== 'Y') return;
        const d = dates[j];
        // Skip future dates — planned slots don't count as completed
        const dateObj = (d instanceof Date) ? d : new Date(String(d).trim());
        if (!isNaN(dateObj.getTime()) && dateObj >= todayCutoff) return;
        memberMap[full][countKey]++;
        memberMap[full][lastKey] = (d instanceof Date)
          ? Utilities.formatDate(d, tz, 'ddMMMyyyy')
          : String(d).trim();
      });
    });
  }

  tally(SN.CORE_VALUE,    'cv', 'cvLast');
  tally(SN.NETWORK_EDU,   'ne', 'neLast');
  tally(SN.FEATURED_PRES, 'fp', 'fpLast');

  const refreshed = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy HH:mm');

  dashboard.clearContents();
  dashboard.clearFormats();

  // Title
  dashboard.getRange(1, 1, 1, 8).merge()
    .setValue('BNI Empower — Presentation Dashboard   (refreshed: ' + refreshed + ')')
    .setBackground('#1a5276')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12)
    .setHorizontalAlignment('center');

  // Column headers
  dashboard.getRange(2, 1, 1, 8)
    .setValues([['Member', 'Core Value', 'Last CV', 'Network Edu', 'Last NE', 'Featured Pres', 'Last FP', 'Total']])
    .setBackground('#154360')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const sortedMembers = Object.values(memberMap).sort((a, b) => a.name.localeCompare(b.name));
  const rows = sortedMembers.map(m => [
    m.name, m.cv, m.cvLast, m.ne, m.neLast, m.fp, m.fpLast, m.cv + m.ne + m.fp,
  ]);

  if (rows.length > 0) {
    dashboard.getRange(3, 1, rows.length, 8).setValues(rows);

    rows.forEach((row, i) => {
      dashboard.getRange(3 + i, 1, 1, 8).setBackground(i % 2 === 0 ? '#d6eaf8' : '#ffffff');
      // Highlight in red any presentation type the member has never done
      if (row[1] === 0) dashboard.getRange(3 + i, 2).setBackground('#fadbd8').setFontColor('#c0392b');
      if (row[3] === 0) dashboard.getRange(3 + i, 4).setBackground('#fadbd8').setFontColor('#c0392b');
      if (row[5] === 0) dashboard.getRange(3 + i, 6).setBackground('#fadbd8').setFontColor('#c0392b');
    });

    // Centre count columns
    [2, 4, 6, 8].forEach(col =>
      dashboard.getRange(3, col, rows.length, 1).setHorizontalAlignment('center')
    );
  }

  [180, 80, 120, 100, 120, 110, 120, 60].forEach((w, i) => dashboard.setColumnWidth(i + 1, w));
  dashboard.setFrozenRows(2);

  if (!suppressAlert) {
    try { SpreadsheetApp.getUi().alert('✅ Dashboard refreshed.'); } catch(_) {}
  }
}

// ── Roster Sheet — plan upcoming presentations in advance ─────────────────────
// Layout: Meeting Date | Meeting Type | Network Education | Core Value | Featured Presentation
// Name columns use dropdowns sourced from Active_Members so names always match exactly.
function setupRosterSheet() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const ui    = SpreadsheetApp.getUi();
  const isNew = !ss.getSheetByName(SN.ROSTER);
  const sheet = isNew ? ss.insertSheet(SN.ROSTER) : ss.getSheetByName(SN.ROSTER);

  sheet.getRange(1, 1, 1, 5)
    .setValues([['Meeting Date', 'Meeting Type', 'Network Education', 'Core Value', 'Featured Presentation']])
    .setBackground('#1a5276')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  [130, 150, 165, 165, 175].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.setFrozenRows(1);

  // Meeting Type dropdown
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Physical Meeting', 'Virtual Meeting'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 2, 100, 1).setDataValidation(typeRule);

  // Name dropdowns for NE / CV / FP columns — sourced from Active_Members roster
  const roster = getRoster_();
  const nameRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(roster, true)
    .setAllowInvalid(true)   // allow free text for guest/TBC entries
    .build();
  sheet.getRange(2, 3, 100, 3).setDataValidation(nameRule);

  if (isNew) {
    ui.alert(
      '✅ Roster sheet created.\n\n' +
      'Columns C–E have name dropdowns from Active_Members.\n' +
      'Copy your existing roster data into this sheet, then use\n' +
      '"Push Roster to Sheets" to stamp the Ys automatically.'
    );
  } else {
    ui.alert('✅ Roster sheet dropdowns refreshed from Active_Members.');
  }
}

function pushRosterToSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ui = SpreadsheetApp.getUi();
  const tz = Session.getScriptTimeZone();

  const rosterSheet = ss.getSheetByName(SN.ROSTER);
  if (!rosterSheet || rosterSheet.getLastRow() < 2) {
    ui.alert('Roster sheet not found or empty. Run "Setup Roster Sheet" first.');
    return;
  }

  const lastRow = rosterSheet.getLastRow();

  // Row layout: [0] Date | [1] Meeting Type | [2] Network Edu | [3] Core Value | [4] Featured Pres
  const data = rosterSheet.getRange(2, 1, lastRow - 1, 5).getValues();

  const mappings = [
    { colIdx: 2, sheetName: SN.NETWORK_EDU },
    { colIdx: 3, sheetName: SN.CORE_VALUE },
    { colIdx: 4, sheetName: SN.FEATURED_PRES },
  ];

  let totalMarked = 0;
  const errors = [];

  data.forEach(row => {
    const dateVal = row[0];
    if (!dateVal) return;

    let dateObj, dateLabel;
    if (dateVal instanceof Date) {
      dateObj   = dateVal;
      dateLabel = Utilities.formatDate(dateVal, tz, 'dd MMM yyyy');
    } else {
      const str = String(dateVal).trim();
      if (!str) return;
      dateObj   = new Date(str);
      dateLabel = isNaN(dateObj.getTime()) ? str : Utilities.formatDate(dateObj, tz, 'dd MMM yyyy');
    }

    // "Physical Meeting" → "Physical", anything else → "Online"
    const type = String(row[1]).trim().toLowerCase().includes('physical') ? 'Physical' : 'Online';

    mappings.forEach(({ colIdx, sheetName }) => {
      const presenterName = String(row[colIdx]).trim();
      if (!presenterName || presenterName === '-') return;

      const presSheet = ss.getSheetByName(sheetName);
      if (!presSheet) { errors.push('Sheet not found: ' + sheetName); return; }

      // Find or create the date column
      let dateCol   = -1;
      const lastCol = presSheet.getLastColumn();
      if (lastCol >= 3) {
        presSheet.getRange(2, 3, 1, lastCol - 2).getValues()[0].forEach((v, i) => {
          const lbl = (v instanceof Date)
            ? Utilities.formatDate(v, tz, 'dd MMM yyyy')
            : String(v).trim();
          if (lbl === dateLabel) dateCol = 3 + i;
        });
      }

      if (dateCol === -1) {
        dateCol = lastCol + 1;
        presSheet.getRange(1, dateCol)
          .setValue(type)
          .setBackground('#c0392b').setFontColor('#ffffff')
          .setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
        presSheet.getRange(2, dateCol)
          .setValue(dateLabel)
          .setFontWeight('bold').setBackground('#962d22')
          .setFontColor('#ecf0f1').setHorizontalAlignment('center');
        presSheet.setColumnWidth(dateCol, 100);
      }

      // Find member row and stamp Y
      const presLastRow = presSheet.getLastRow();
      if (presLastRow < 3) { errors.push(sheetName + ': no member rows'); return; }

      const nameData   = presSheet.getRange(3, 1, presLastRow - 2, 2).getValues();
      const normTarget = presenterName.toLowerCase();
      let matched      = false;

      nameData.forEach((nameRow, i) => {
        const full  = (String(nameRow[0]).trim() + ' ' + String(nameRow[1]).trim()).trim().toLowerCase();
        const first = String(nameRow[0]).trim().toLowerCase();
        if (full === normTarget || first === normTarget) {
          presSheet.getRange(3 + i, dateCol).setValue('Y');
          matched = true;
          totalMarked++;
        }
      });

      if (!matched) errors.push('"' + presenterName + '" not found in ' + sheetName + ' (' + dateLabel + ')');
    });
  });

  let msg = '✅ Push complete — ' + totalMarked + ' Y-mark(s) written.';
  if (errors.length > 0) msg += '\n\n⚠️ Warnings (check spellings):\n' + errors.join('\n');
  ui.alert(msg);
}
