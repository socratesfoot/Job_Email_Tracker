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
 * Extra behavior:
 *  - Extracts compensation into a "Salary" column ($50/hr, $120k, ...).
 *  - Adds a "Gmail Link" column that jumps straight to the email thread.
 *  - Scores every posting 0-100 ("Score" column) against your targets and
 *    marks high scorers IMPORTANT in the Priority column:
 *      * full-time employment preferred; contracts need 12+ months
 *      * remote preferred; $200k+/yr target
 *      * specialties: key management, information security, AI,
 *        TMS, project management
 *  - Filters out non-serious leads: educational/training pitches, part-time
 *    roles, and contracts under 12 months (when a duration is stated).
 *  - Merges duplicate postings for the same company + position into one row
 *    instead of creating a new row each time a recruiter re-sends a role.
 *  - Rows with an active INTERVIEW REQUEST stage are never auto-aged to
 *    OLD / OUTDATED, and any new message on one keeps it IMPORTANT so
 *    replies you requested don't get missed.
 *  - Only rows that actually changed are written back on each run (the old
 *    version rewrote the entire sheet every 15 minutes).
 *  - Warns in the log when distance filtering is on but the Maps service
 *    isn't enabled, instead of silently skipping the check.
 *
 * Tune the targets below (TARGET_ANNUAL_SALARY, MIN_CONTRACT_MONTHS,
 * SPECIALTIES, IMPORTANT_AT) to taste.
 *
 * SETUP: paste into Extensions > Apps Script in a Google Sheet, then
 *   1. If you use distance filtering, enable the Maps service:
 *      Services (+) > Maps. Without it, distance checks are skipped and a
 *      warning is logged each run.
 *   2. Run setup()    (grants permissions, builds the Jobs and Settings
 *                      sheets, installs the 15-min trigger)
 *   3. Open the "Settings" tab and set your filters (all off by default)
 *   4. Run backfill() (one-time scan of the last 90 days)
 *
 * Reopening the spreadsheet also adds a "Job Tracker" menu with shortcuts
 * to the Settings tab, a manual scan, and a manual backfill.
 */

const SHEET_NAME = 'Jobs';
const HEADERS = [
  'Status', 'Priority', 'Stage', 'Date Received', 'Company', 'Position',
  'Type', 'Work Mode', 'Location', 'Contact Name', 'Contact Email',
  'Contact Phone', 'Subject', 'Thread ID', 'Message IDs', 'Last Activity',
  'Salary', 'Gmail Link', 'Score'
];
const C = {};
HEADERS.forEach((h, i) => (C[h] = i));

const OLD_AFTER_DAYS = 7;
const OUTDATED_AFTER_DAYS = 90;

// Stages that represent a live opportunity: never auto-aged to OLD/OUTDATED.
const ACTIVE_STAGES = ['INTERVIEW REQUEST'];

// ---- prioritization targets (tune to taste) ----
const TARGET_ANNUAL_SALARY = 200000; // $/yr full-time-equivalent target
const MIN_CONTRACT_MONTHS = 12;      // contracts must run at least this long
const BASE_SCORE = 40;               // every posting starts here (0-100 scale)
const IMPORTANT_AT = 75;             // score at/above this -> Priority IMPORTANT

// Specialty areas that raise a posting's score (regex per area).
const SPECIALTIES = [
  /key management/i,
  /information security|infosec|cyber\s?security/i,
  /\bAI\b|artificial intelligence|machine learning/i,
  /\bTMS\b/i,
  /project management/i
];

// Patterns that mark a lead as not serious (filtered in shouldOmit_).
const NOT_SERIOUS = [
  [/educat|bootcamp|training course|certification program|learn to code/i, 'educational/training pitch'],
  [/part[- ]time/i, 'part-time role']
];

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

// True when the Maps advanced service is enabled (Services (+) > Maps).
// Referencing Maps when the service is off throws, so guard with typeof.
function mapsAvailable_() {
  try {
    return typeof Maps !== 'undefined' && !!Maps.newGeocoder;
  } catch (e) {
    return false;
  }
}

// Geocodes are cached for 6 hours (CacheService's max) so repeated ZIP
// codes and repeated job locations don't re-hit the Maps service.
function geocodeCached_(address) {
  if (!address || !mapsAvailable_()) return null;
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

  // Not-serious leads: training pitches, part-time roles.
  for (const [re, label] of NOT_SERIOUS) {
    if (re.test(text)) return label;
  }

  // Contracts must run at least MIN_CONTRACT_MONTHS (only when a duration
  // is actually stated - unknown durations are kept and scored lower).
  if (row[C['Type']] === 'Contract') {
    const months = extractContractMonths_(text);
    if (months > 0 && months < MIN_CONTRACT_MONTHS) {
      return 'contract too short: ' + months + ' mo (minimum ' + MIN_CONTRACT_MONTHS + ')';
    }
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

/* ---------- compensation ---------- */

// Pulls a compensation snippet like "$65/hr", "$120k-$150k" or
// "$95,000 per year" out of the email text. Returns '' when nothing
// compensation-like is found (a bare "$" or "k" is required, so things
// like "5 years experience" don't match). Scans all candidates and takes
// the first that looks like compensation, so an early bare number (e.g.
// the "6" in "6-month contract") doesn't shadow a real "$70/hr" later on.
function extractSalary_(text) {
  const re = /\$?\s*\d[\d,]*(?:\.\d{1,2})?\s*k?\s*(?:[-–]\s*\$?\s*\d[\d,]*(?:\.\d{1,2})?\s*k?)?\s*(\/\s*(?:hr|hour|year|yr|annum)|per\s+(?:hour|year|annum))?/gi;
  const src = String(text);
  let m;
  while ((m = re.exec(src)) !== null) {
    const s = m[0].replace(/\s+/g, ' ').trim();
    if (!s) continue;
    const looksPaid = /\$/.test(s)
      || /k(\s*([-–]|\/|$)|$)/i.test(s)
      || /(\/\s*(hr|hour|year|yr|annum)|per\s+(hour|year|annum))/i.test(s);
    if (looksPaid) return s.slice(0, 60);
  }
  return '';
}

// Best-effort annual USD equivalent of any compensation mentioned, or 0.
// Hourly rates are annualized at 2080 hrs; "k" suffixes expand; bare "$"
// amounts are assumed annual. Returns the largest figure found.
function salaryAnnualMax_(text) {
  const re = /\$?\s*(\d[\d,]*(?:\.\d{1,2})?)\s*(k)?\s*(?:[-–]\s*\$?\s*(\d[\d,]*(?:\.\d{1,2})?)\s*(k)?)?\s*(\/\s*(hr|hour|year|yr|annum)|per\s+(hour|year|annum))?/gi;
  const src = String(text);
  let max = 0, m;
  while ((m = re.exec(src)) !== null) {
    const hasMoney = /\$/.test(m[0]) || /k/i.test(m[0]) || m[5];
    if (!hasMoney) continue;
    const unit = (m[2] || m[4]) ? 1000 : 1;
    const lo = parseFloat(m[1].replace(/,/g, '')) * unit;
    const hi = m[3] ? parseFloat(m[3].replace(/,/g, '')) * unit : lo;
    let annual = Math.max(lo, hi);
    if (/hr|hour/i.test(m[5] || '')) annual *= 2080;
    if (annual > max) max = annual;
  }
  return Math.round(max);
}

// Contract length in months, or 0 when none is stated. Only trusts a number
// when it appears near contract language, so "5 years experience" doesn't
// count as a 60-month contract.
function extractContractMonths_(text) {
  const t = String(text).toLowerCase();
  let m = t.match(/(\d+)\s*(?:-|–)?\s*months?[^.\n]{0,40}(contract|term|assignment|engagement)|(contract|term|duration|assignment|engagement)[^.\n]{0,60}?(\d+)\s*(?:-|–)?\s*months?/);
  if (m) return parseInt(m[1] || m[4], 10);
  m = t.match(/(\d+)\s*years?[^.\n]{0,40}(contract|term|assignment)|(contract|term|duration|assignment)[^.\n]{0,60}?(\d+)\s*years?/);
  if (m) return parseInt(m[1] || m[4], 10) * 12;
  return 0;
}

// 0-100 fit score for a posting against your targets. Higher = closer to
// full-time, remote, $200k+, and your specialty areas.
function scorePosting_(row, text) {
  let score = BASE_SCORE;
  const type = row[C['Type']];
  const mode = row[C['Work Mode']];

  if (type === 'FTE') {
    score += 30;
  } else if (type === 'Contract') {
    const months = extractContractMonths_(text);
    if (months >= MIN_CONTRACT_MONTHS) score += 15;
    else if (months > 0) score -= 30; // shouldn't happen (filtered), belt & braces
  }

  if (mode === 'Remote') score += 20;
  else if (mode === 'Hybrid') score += 5;

  const annual = salaryAnnualMax_(text);
  if (annual >= TARGET_ANNUAL_SALARY) score += 25;
  else if (annual >= 150000) score += 10;
  else if (annual > 0) score -= 5;

  let hits = 0;
  SPECIALTIES.forEach(re => { if (re.test(text)) hits++; });
  score += Math.min(hits, 5) * 5;

  return Math.max(0, Math.min(100, Math.round(score)));
}

/* ---------- entry points ---------- */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.getRange(2, C['Date Received'] + 1, 1000, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, C['Last Activity'] + 1, 1000, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, C['Score'] + 1, 1000, 1).setNumberFormat('0');

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
  ensureHeaders_(sh);
  const rows = readRows_(sh);
  const settings = getSettings_();

  const mapsOk = mapsAvailable_();
  if (settings.limitDistance && !mapsOk) {
    Logger.log('WARNING: "Omit hybrid/onsite jobs over a certain distance" is on, ' +
      'but the Maps service is not enabled (Apps Script > Services (+) > Maps). ' +
      'Distance checks will be skipped until it is enabled.');
  }
  const homeLoc = (settings.limitDistance && settings.homeZip && mapsOk) ? geocodeCached_(settings.homeZip) : null;
  if (settings.limitDistance && settings.homeZip && !homeLoc) {
    Logger.log('WARNING: home ZIP "' + settings.homeZip + '" could not be geocoded; distance checks skipped this run.');
  }

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
  const dirty = new Set(); // row indexes (into rows) that changed this run
  msgs.forEach(m => handleMessage_(rows, m, settings, homeLoc, dirty));
  writeDirtyRows_(sh, rows, dirty);
}

function handleMessage_(rows, m, settings, homeLoc, dirty) {
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
  const ti = rows.findIndex(r => r[C['Thread ID']] === threadId);
  if (ti !== -1) {
    const row = rows[ti];
    sig ? applySignal_(row, sig, m) : touch_(row, m);
    // Never miss a reply on a live interview thread: any new message there
    // keeps the row IMPORTANT even when it isn't itself an interview invite.
    if (!sig && row[C['Stage']] === 'INTERVIEW REQUEST') row[C['Priority']] = 'IMPORTANT';
    dirty.add(ti);
    return;
  }

  // 2) Reply in a new thread (e.g. from an ATS): match by company
  if (sig) {
    const byCompany = findRowByCompany_(rows, sender.domainRoot, text);
    if (byCompany) {
      applySignal_(byCompany, sig, m);
      dirty.add(rows.indexOf(byCompany));
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
    // 3b) Same company + position already tracked? Merge instead of duplicating.
    const di = findDuplicate_(rows, row);
    if (di !== -1) {
      const dup = rows[di];
      touch_(dup, m);
      if (!dup[C['Salary']] && row[C['Salary']]) dup[C['Salary']] = row[C['Salary']];
      dup[C['Score']] = Math.max(dup[C['Score']] || 0, row[C['Score']]);
      if (dup[C['Score']] >= IMPORTANT_AT) dup[C['Priority']] = 'IMPORTANT';
      if (sig) applySignal_(dup, sig, m);
      dirty.add(di);
      return;
    }
    rows.push(row);
    if (sig) applySignal_(row, sig, m);
    dirty.add(rows.length - 1);
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
  r[C['Salary']] = extractSalary_(text);
  r[C['Score']] = scorePosting_(r, text);
  if (r[C['Score']] >= IMPORTANT_AT) r[C['Priority']] = 'IMPORTANT';
  // Gmail Link is built at write time from the Thread ID (see writeDirtyRows_).
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

// Same normalized company + position already on the sheet? Returns its index,
// or -1. Never matches rows with an unknown/generic company name.
function findDuplicate_(rows, row) {
  const co = norm_(row[C['Company']]);
  const pos = norm_(row[C['Position']]);
  if (co.length < 4 || co === 'unknown' || pos.length < 4) return -1;
  return rows.findIndex(r => norm_(r[C['Company']]) === co && norm_(r[C['Position']]) === pos);
}

/* ---------- aging ---------- */

function refreshStatuses() {
  const sh = getSheet_();
  const rows = readRows_(sh);
  if (!rows.length) return;
  const now = Date.now();
  const out = rows.map(r => {
    if (r[C['Stage']] === 'REJECTED') return ['CANCELLED'];
    // Active opportunities keep their current status instead of aging out.
    if (ACTIVE_STAGES.indexOf(r[C['Stage']]) !== -1) return [r[C['Status']] || 'NEW'];
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

// Adds any missing headers (e.g. Salary / Gmail Link on sheets created by an
// older version). New columns are always appended, so existing data is never
// shifted.
function ensureHeaders_(sh) {
  const have = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  let changed = false;
  const out = have.slice();
  for (let i = 0; i < HEADERS.length; i++) {
    if (String(out[i] || '').trim() !== HEADERS[i]) {
      out[i] = HEADERS[i];
      changed = true;
    }
  }
  if (changed) sh.getRange(1, 1, 1, HEADERS.length).setValues([out]);
}

function readRows_(sh) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, HEADERS.length).getValues() : [];
}

// Writes back only the rows that changed this run. The Gmail Link formula is
// rebuilt from the Thread ID here because getValues() returns a link's display
// text ("Open"), not its formula - writing that back would destroy the link.
function writeDirtyRows_(sh, rows, dirty) {
  dirty.forEach(i => {
    const r = rows[i].slice();
    const tid = r[C['Thread ID']];
    r[C['Gmail Link']] = tid
      ? '=HYPERLINK("https://mail.google.com/mail/u/0/#all/' + tid + '","Open")'
      : '';
    sh.getRange(i + 2, 1, 1, HEADERS.length).setValues([r]);
  });
}

// Score-based posting detection: plain keywords count 1, compensation counts
// 2 and an explicit work mode counts 1. Catches short recruiter emails like
// "6-month contract, remote, $70/hr" that the old 3-keyword bar missed.
function isPosting_(text) {
  const t = text.toLowerCase();
  const keys = ['job opening', 'position', 'role', 'opportunity', 'hiring',
    'responsibilities', 'qualifications', 'requirements', 'apply', 'contract',
    'full-time', 'w2'];
  let score = 0;
  keys.forEach(k => { if (t.indexOf(k) !== -1) score++; });
  if (extractSalary_(text)) score += 2;
  if (/\b(remote|hybrid|onsite|on-site|wfh)\b/i.test(text)) score += 1;
  return score >= 3;
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
  const alt = labels.join('|');
  let m = text.match(new RegExp('(?:^|\\n)\\s*(?:' + alt + ')\\s*[:\\-\\u2013]\\s*(.+)', 'i'));
  if (!m) {
    // Fallback: label appearing inline, e.g. "Remote | Location: Austin, TX".
    m = text.match(new RegExp('(?:^|[\\n\\s(\\[])(?:' + alt + ')\\s*[:\\-\\u2013]\\s*([^\\n;|]+)', 'i'));
  }
  return m ? m[1].trim().slice(0, 120) : '';
}

function norm_(s) {
  return String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
}
