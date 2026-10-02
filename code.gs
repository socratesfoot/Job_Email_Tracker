/**
 * Job Email Tracker
 * Scans Gmail for job offers, logs them to a "Jobs" sheet, ages them
 * (NEW -> OLD -> OUTDATED), and tracks replies (interview / rejection).
 *
 * An admin console (a "Settings" sheet tab, created automatically) controls
 * what gets kept when a new posting is found: require a contact email,
 * require a location, require a phone number, a distance cap for
 * hybrid/onsite jobs, and a blacklist of phrases/keywords. These filters
 * only affect NEW postings as they're found - they never delete or hide
 * rows you've already captured.
 *
 * SETUP: paste into Extensions > Apps Script in a Google Sheet, then
 *   1. Run setup()    (grants permissions, builds the Jobs and Settings
 *                      sheets, installs the 15-min trigger)
 *   2. Open the "Settings" tab and set your filters (all off by default)
 *   3. Run backfill() (one-time scan of the last 90 days)
 *
 * Reopening the spreadsheet also adds a "Job Tracker" menu with shortcuts
 * to the Settings tab, a manual scan, and a manual backfill.
 */

const SHEET_NAME = 'Jobs';
const HEADERS = [
  'Status', 'Priority', 'Stage', 'Date Received', 'Company', 'Position',
  'Type', 'Work Mode', 'Location', 'Contact Name', 'Contact Email',
  'Contact Phone', 'Subject', 'Thread ID', 'Message IDs', 'Last Activity'
];
const C = {};
HEADERS.forEach((h, i) => (C[h] = i));

const OLD_AFTER_DAYS = 7;
const OUTDATED_AFTER_DAYS = 90;

const NOREPLY = /no-?reply|do-?not-?reply|notifications?@|mailer/i;
const GENERIC = ['gmail', 'yahoo', 'outlook', 'hotmail', 'icloud', 'aol',
  'greenhouse', 'lever', 'workday', 'indeed', 'linkedin', 'ziprecruiter',
  'smartrecruiters', 'icims', 'jobvite', 'ashby', 'taleo'];

const RE = {
  ack: /thank(s| you) for (your )?(interest|applying|application|submitting)|received your (application|resume)|application (has been |was )?received/i,
  reject: /unfortunately,? (we|after|at this|the position)|not (be )?(moving|proceeding) forward|decided to (move|proceed|go) (forward )?with (other|another)|position (has been|was) filled|no longer (under consideration|being considered)|will not be (moving|proceeding)|not selected/i,
  interview: /(schedule|set up|arrange)\s+(an?|your|the)?\s*(\w+\s+)?(interview|call|screen|chat)|invite you to (an? )?interview|like to (interview|speak with|meet with) you|phone screen|your availability|next round/i,
  phone: /(?<!\d)(?:\+?1[\s.\-]?)?\(?\d{3}\)?[\s.\-]?\d{3}[\s.\-]?\d{4}(?!\d)/
};

/* ---------- admin console (Settings sheet) ---------- */

const SETTINGS_SHEET_NAME = 'Settings';

// [key, row label, value type]. Order = row order on the Settings sheet.
// A label starting with two spaces is a sub-field of the checkbox above it.
const SETTINGS_FIELDS = [
  ['requireEmail',    'Omit rows that lack a contact email',             'checkbox'],
  ['requireLocation', 'Omit rows that lack a location',                  'checkbox'],
  ['requirePhone',    'Omit rows that lack a contact phone number',      'checkbox'],
  ['limitDistance',   'Omit hybrid/onsite jobs over a certain distance', 'checkbox'],
  ['homeZip',         '  ZIP code',                                     'text'],
  ['maxMiles',        '  Distance in miles',                            'number'],
  ['useBlacklist',    'Omit rows with blacklisted phrases/keywords',     'checkbox'],
  ['blacklist',       '  Blacklisted phrases (comma separated)',        'text']
];

function getSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (sh) return sh;

  sh = ss.insertSheet(SETTINGS_SHEET_NAME);
  sh.getRange(1, 1, 1, 2).setValues([['Setting', 'Value']]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 360);
  sh.setColumnWidth(2, 280);

  const rows = SETTINGS_FIELDS.map(f => [f[1], f[2] === 'checkbox' ? false : '']);
  sh.getRange(2, 1, rows.length, 2).setValues(rows);

  SETTINGS_FIELDS.forEach((f, i) => {
    if (f[2] === 'checkbox') sh.getRange(i + 2, 2).insertCheckboxes();
  });

  const noteFor = (key, note) => {
    const i = SETTINGS_FIELDS.findIndex(f => f[0] === key);
    if (i !== -1) sh.getRange(i + 2, 2).setNote(note);
  };
  noteFor('homeZip', 'Only used when the distance checkbox above is on.');
  noteFor('maxMiles', 'Only used when the distance checkbox above is on. Leave blank or 0 to disable the distance check even if the box is checked.');
  noteFor('blacklist',
    'Comma separated. A new job is skipped if its email contains any of these ' +
    '(case-insensitive, matched anywhere in the subject or body). Example: ' +
    'Sign up for, weekly, Save up to, Buy, unsubscribe, points');

  return sh;
}

function getSettings_() {
  const defaults = {
    requireEmail: false,
    requireLocation: false,
    requirePhone: false,
    limitDistance: false,
    homeZip: '',
    maxMiles: 0,
    useBlacklist: false,
    blacklist: []
  };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (!sh) return defaults;

  const values = sh.getRange(2, 1, SETTINGS_FIELDS.length, 2).getValues();
  const byLabel = {};
  values.forEach(r => { byLabel[String(r[0]).trim()] = r[1]; });

  const s = Object.assign({}, defaults);
  SETTINGS_FIELDS.forEach(f => {
    const key = f[0], label = f[1].trim(), type = f[2];
    const raw = byLabel[label];
    if (raw === undefined || raw === '') return;
    if (type === 'checkbox') {
      s[key] = raw === true || String(raw).toUpperCase() === 'TRUE';
    } else if (type === 'number') {
      const n = parseFloat(raw);
      s[key] = isNaN(n) ? 0 : n;
    } else if (key === 'blacklist') {
      s[key] = String(raw).split(',').map(x => x.trim()).filter(Boolean);
    } else {
      s[key] = String(raw).trim();
    }
  });
  return s;
}

function blank_(v) {
  return !String(v || '').trim();
}

function normText_(s) {
  return String(s).toLowerCase().replace(/\s+/g, ' ').trim();
}

// Geocodes are cached for 6 hours (CacheService's max) so repeated ZIP
// codes and repeated job locations don't re-hit the Maps service.
function geocodeCached_(address) {
  if (!address) return null;
  const cache = CacheService.getScriptCache();
  const key = 'geo_' + normText_(address).replace(/[^a-z0-9]/g, '_').slice(0, 200);
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  try {
    const res = Maps.newGeocoder().geocode(address);
    if (!res || !res.results || !res.results.length) return null;
    const loc = res.results[0].geometry.location;
    cache.put(key, JSON.stringify(loc), 21600);
    return loc;
  } catch (e) {
    return null;
  }
}

function haversineMiles_(lat1, lng1, lat2, lng2) {
  const R = 3958.8;
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) * Math.sin(dLng / 2) ** 2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// Returns a short reason string if the row should be skipped, or null to keep it.
// Only called for brand new postings - it never touches rows already on the sheet.
function shouldOmit_(row, text, settings, homeLoc) {
  if (settings.requireEmail && blank_(row[C['Contact Email']])) return 'missing contact email';
  if (settings.requireLocation && blank_(row[C['Location']])) return 'missing location';
  if (settings.requirePhone && blank_(row[C['Contact Phone']])) return 'missing contact phone';

  if (settings.useBlacklist && settings.blacklist.length) {
    const hay = normText_(text);
    const hit = settings.blacklist.find(k => k && hay.indexOf(normText_(k)) !== -1);
    if (hit) return 'blacklisted phrase: ' + hit;
  }

  if (settings.limitDistance && settings.maxMiles > 0 && homeLoc) {
    const mode = row[C['Work Mode']];
    if ((mode === 'Hybrid' || mode === 'Onsite') && !blank_(row[C['Location']])) {
      const jobLoc = geocodeCached_(row[C['Location']]);
      if (jobLoc) {
        const miles = haversineMiles_(homeLoc.lat, homeLoc.lng, jobLoc.lat, jobLoc.lng);
        if (miles > settings.maxMiles) return 'too far: ' + Math.round(miles) + ' mi (limit ' + settings.maxMiles + ')';
      }
    }
  }
  return null;
}

/* ---------- entry points ---------- */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(2, C['Date Received'] + 1, 1000, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, C['Last Activity'] + 1, 1000, 1).setNumberFormat('yyyy-mm-dd');

  const rng = sh.getRange(2, 1, 1000, HEADERS.length);
  const rule = (formula, bg, font, bold) => {
    let b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setRanges([rng]);
    if (bg) b = b.setBackground(bg);
    if (font) b = b.setFontColor(font);
    if (bold) b = b.setBold(true);
    return b.build();
  };
  sh.setConditionalFormatRules([
    rule('=$B2="IMPORTANT"', '#fff2a8', null, true),
    rule('=$A2="CANCELLED"', null, '#999999', false),
    rule('=$A2="OUTDATED"', '#eeeeee', '#999999', false),
    rule('=$A2="OLD"', '#fce8b2', null, false),
    rule('=$A2="NEW"', '#d9ead3', null, false)
  ]);

  getSettingsSheet_();

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'run')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('run').timeBased().everyMinutes(15).create();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Job Tracker')
    .addItem('Open settings', 'openSettings')
    .addItem('Scan now', 'run')
    .addItem('Backfill last 90 days', 'backfill')
    .addToUi();
}

function openSettings() {
  const sh = getSettingsSheet_();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sh);
}

function backfill() {
  processMailbox_(90);
  refreshStatuses();
}

function run() {
  processMailbox_(3);
  refreshStatuses();
}

/* ---------- mailbox scan ---------- */

function processMailbox_(days) {
  const sh = getSheet_();
  const rows = readRows_(sh);
  const settings = getSettings_();
  const homeLoc = (settings.limitDistance && settings.homeZip) ? geocodeCached_(settings.homeZip) : null;

  const seen = new Set();
  rows.forEach(r => String(r[C['Message IDs']]).split(',').forEach(id => id && seen.add(id)));

  const me = Session.getActiveUser().getEmail().toLowerCase();
  const query = 'newer_than:' + days + 'd -in:sent -in:drafts -in:spam -in:trash ' +
    '(job OR position OR opportunity OR role OR hiring OR recruiter OR application OR interview OR "your interest")';

  const msgs = [];
  for (let start = 0; start < 500; start += 100) {
    const threads = GmailApp.search(query, start, 100);
    if (!threads.length) break;
    threads.forEach(t => t.getMessages().forEach(m => {
      if (!seen.has(m.getId()) && m.getFrom().toLowerCase().indexOf(me) === -1) msgs.push(m);
    }));
  }

  msgs.sort((a, b) => a.getDate() - b.getDate()); // oldest first so later replies win
  msgs.forEach(m => handleMessage_(rows, m, settings, homeLoc));
  writeRows_(sh, rows);
}

function handleMessage_(rows, m, settings, homeLoc) {
  settings = settings || getSettings_();
  const subject = m.getSubject() || '';
  const body = m.getPlainBody() || '';
  const text = subject + '\n' + body;
  const sender = parseSender_(m.getFrom());
  const threadId = m.getThread().getId();

  const sig = RE.reject.test(text) ? 'REJECTED'
    : RE.interview.test(text) ? 'INTERVIEW'
    : RE.ack.test(text) ? 'ACK' : null;

  // 1) Same thread as a tracked job: update it
  const byThread = rows.find(r => r[C['Thread ID']] === threadId);
  if (byThread) {
    sig ? applySignal_(byThread, sig, m) : touch_(byThread, m);
    return;
  }

  // 2) Reply in a new thread (e.g. from an ATS): match by company
  if (sig) {
    const byCompany = findRowByCompany_(rows, sender.domainRoot, text);
    if (byCompany) {
      applySignal_(byCompany, sig, m);
      return;
    }
  }

  // 3) New posting (or an untracked response worth logging)
  if (isPosting_(text) || sig) {
    const row = buildRow_(m, subject, body, text, sender, threadId);
    const reason = shouldOmit_(row, text, settings, homeLoc);
    if (reason) {
      Logger.log('Skipped (' + reason + '): ' + subject);
      return;
    }
    rows.push(row);
    if (sig) applySignal_(row, sig, m);
  }
}

/* ---------- row logic ---------- */

function buildRow_(m, subject, body, text, sender, threadId) {
  const r = new Array(HEADERS.length).fill('');
  const bodyEmail = (body.match(/[\w.+\-]+@[\w\-]+(?:\.[\w\-]+)+/g) || []).find(e => !NOREPLY.test(e));
  const phone = (body.match(RE.phone) || [''])[0].trim();

  const typeSrc = field_(text, ['Employment Type', 'Job Type', 'Type', 'Duration']) || text;
  const type = /\b(c2h|contract[- ]to[- ]hire|contract|corp[- ]to[- ]corp|c2c|1099)\b/i.test(typeSrc) ? 'Contract'
    : /\b(full[- ]time|fte|permanent|direct[- ]hire|perm)\b/i.test(typeSrc) ? 'FTE' : '';

  const mode = /hybrid/i.test(text) ? 'Hybrid'
    : /on-?site|in[- ]office|in[- ]person/i.test(text) ? 'Onsite'
    : /remote|work from home|wfh/i.test(text) ? 'Remote' : '';

  const cityState = (text.match(/\b([A-Z][a-z]+(?:\s[A-Z][a-z]+)*,\s?[A-Z]{2})\b/) || [])[1] || '';

  r[C['Stage']] = 'Open';
  r[C['Date Received']] = m.getDate();
  r[C['Company']] = field_(body, ['Company', 'Client', 'Employer', 'Organization']) || companyFromDomain_(sender.domainRoot);
  r[C['Position']] = field_(text, ['Job Title', 'Position', 'Role', 'Title', 'Opening']) || cleanSubject_(subject);
  r[C['Type']] = type;
  r[C['Work Mode']] = mode;
  r[C['Location']] = field_(text, ['Job Location', 'Location', 'City']) || cityState;
  r[C['Contact Name']] = NOREPLY.test(sender.email) ? field_(body, ['Recruiter', 'Contact']) : sender.name;
  r[C['Contact Email']] = NOREPLY.test(sender.email) ? (bodyEmail || sender.email) : sender.email;
  r[C['Contact Phone']] = phone;
  r[C['Subject']] = subject;
  r[C['Thread ID']] = threadId;
  r[C['Message IDs']] = m.getId();
  r[C['Last Activity']] = m.getDate();
  return r;
}

function applySignal_(row, sig, m) {
  touch_(row, m);
  if (sig === 'REJECTED') {
    row[C['Stage']] = 'REJECTED';
    row[C['Priority']] = '';
  } else if (sig === 'INTERVIEW') {
    row[C['Stage']] = 'INTERVIEW REQUEST';
    row[C['Priority']] = 'IMPORTANT';
  } else if (sig === 'ACK') {
    const s = row[C['Stage']];
    if (!s || s === 'Open') row[C['Stage']] = 'Applied';
  }
}

function touch_(row, m) {
  const ids = String(row[C['Message IDs']]).split(',').filter(Boolean);
  if (ids.indexOf(m.getId()) === -1) ids.push(m.getId());
  row[C['Message IDs']] = ids.join(',');
  row[C['Last Activity']] = m.getDate();
}

function findRowByCompany_(rows, root, text) {
  const hay = norm_(text);
  const generic = !root || GENERIC.some(g => root.indexOf(g) !== -1);
  for (let i = rows.length - 1; i >= 0; i--) {
    const co = norm_(rows[i][C['Company']]);
    if (co.length < 4 || co === 'unknown') continue;
    if ((!generic && (co.indexOf(root) !== -1 || root.indexOf(co) !== -1)) || hay.indexOf(co) !== -1) {
      return rows[i];
    }
  }
  return null;
}

/* ---------- aging ---------- */

function refreshStatuses() {
  const sh = getSheet_();
  const rows = readRows_(sh);
  if (!rows.length) return;
  const now = Date.now();
  const out = rows.map(r => {
    if (r[C['Stage']] === 'REJECTED') return ['CANCELLED'];
    const age = (now - new Date(r[C['Date Received']]).getTime()) / 86400000;
    return [age > OUTDATED_AFTER_DAYS ? 'OUTDATED' : age > OLD_AFTER_DAYS ? 'OLD' : 'NEW'];
  });
  sh.getRange(2, 1, out.length, 1).setValues(out);
}

/* ---------- helpers ---------- */

function getSheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('Run setup() first.');
  return sh;
}

function readRows_(sh) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, HEADERS.length).getValues() : [];
}

function writeRows_(sh, rows) {
  if (rows.length) sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
}

function isPosting_(text) {
  const t = text.toLowerCase();
  const keys = ['job opening', 'position', 'role', 'opportunity', 'hiring',
    'responsibilities', 'qualifications', 'requirements', 'apply', 'contract',
    'full-time', 'w2', 'rate'];
  return keys.filter(k => t.indexOf(k) !== -1).length >= 3;
}

function parseSender_(from) {
  const m = from.match(/^"?([^"<]*?)"?\s*<([^>]+)>$/);
  const email = (m ? m[2] : from).trim().toLowerCase();
  const name = m ? m[1].trim() : '';
  const parts = (email.split('@')[1] || '').split('.');
  const root = parts.length >= 2 ? parts[parts.length - 2] : (parts[0] || '');
  return { name: name, email: email, domainRoot: root };
}

function companyFromDomain_(root) {
  if (!root || GENERIC.some(g => root.indexOf(g) !== -1)) return 'Unknown';
  return root.charAt(0).toUpperCase() + root.slice(1);
}

function cleanSubject_(s) {
  return s.replace(/^((re|fwd?):\s*)+/i, '')
    .replace(/^(job|new|urgent|hot)?\s*(opportunity|opening|position|role)\s*[:\-\u2013]\s*/i, '')
    .trim().slice(0, 120);
}

function field_(text, labels) {
  const m = text.match(new RegExp('(?:^|\\n)\\s*(?:' + labels.join('|') + ')\\s*[:\\-\\u2013]\\s*(.+)', 'i'));
  return m ? m[1].trim().slice(0, 120) : '';
}

function norm_(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
}
