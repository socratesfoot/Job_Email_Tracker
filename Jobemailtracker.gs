/**
 * Job Email Tracker
 * Scans Gmail for job offers, logs them to a "Jobs" sheet, ages them
 * (NEW -> OLD -> OUTDATED), and tracks replies (interview / rejection).
 *
 * SETUP: paste into Extensions > Apps Script in a Google Sheet, then
 *   1. Run setup()    (grants permissions, builds sheet, installs 15-min trigger)
 *   2. Run backfill() (one-time scan of the last 90 days)
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

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'run')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('run').timeBased().everyMinutes(15).create();
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
  msgs.forEach(m => handleMessage_(rows, m));
  writeRows_(sh, rows);
}

function handleMessage_(rows, m) {
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
