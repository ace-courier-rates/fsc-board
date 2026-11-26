/**
 * ACE expense capture.
 *
 * Every 15 minutes:
 *   1. Gmail threads labelled "ACE-Receipt" -> receipt PDF/images saved into the
 *      Ace Expenses Drive folder, thread relabelled "ACE-Receipt/Filed".
 *   2. New photos/PDFs in Ace Expenses (or any trip subfolder) -> read, renamed
 *         YYYY-MM-DD - <Trip> - <Vendor> - <Total>.<ext>
 *      moved into the trip folder, and logged as a row on "All Expenses".
 *
 * Trip is decided by, in order: the trip folder you uploaded into; the Trips tab
 * (receipt date within a trip's Start..End, +/- 1 day); otherwise "No Trip".
 */

const CFG = {
  ROOT_FOLDER_ID: '1EzM_1KjpT7EWKwApJC7Lkavu1b0xCNWz',      // Drive: Ace Expenses
  SPREADSHEET_ID: '18pMykonQD52QiCV4MevlY86T0wjX8w6I5gCa3Ogtqn8', // Sheet: Ace Expenses
  EXPENSE_SHEET: 'All Expenses',
  TRIPS_SHEET: 'Trips',
  HEADER_ROW: 5,
  GMAIL_LABEL: 'ACE-Receipt',
  GMAIL_DONE_LABEL: 'ACE-Receipt/Filed',
  NO_TRIP: 'No Trip',
  TRIP_SLACK_DAYS: 1,
  MAX_FILES_PER_RUN: 15,          // keeps each run well under the 6-minute limit
  MODEL: 'claude-opus-5-5',
  MARK: 'ace-expense:filed',      // written to a file's Drive description once handled
  CATEGORIES: ['Meals', 'Accommodation', 'Transport', 'Fuel', 'Other'],
};

// ---------------------------------------------------------------- entry points

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Expenses')
    .addItem('Process new receipts now', 'run')
    .addItem('Set up / reinstall', 'setup')
    .addToUi();
}

/** Run once. Creates the Trips tab, headers, Gmail labels and the 15-minute trigger. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  // Only files added after setup are touched, so existing receipts keep their names.
  if (!props.getProperty('START_AT')) props.setProperty('START_AT', new Date().toISOString());
  ensureTripsSheet_();
  ensureHeaders_();
  label_(CFG.GMAIL_LABEL);
  label_(CFG.GMAIL_DONE_LABEL);
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'run')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('run').timeBased().everyMinutes(15).create();
}

function run() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    importEmail_();
    processDrive_();
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- email

function importEmail_() {
  const root = DriveApp.getFolderById(CFG.ROOT_FOLDER_ID);
  const done = label_(CFG.GMAIL_DONE_LABEL);
  const q = 'label:' + slug_(CFG.GMAIL_LABEL) + ' -label:' + slug_(CFG.GMAIL_DONE_LABEL);
  GmailApp.search(q, 0, 20).forEach(thread => {
    thread.getMessages().forEach(msg => {
      const hint = 'email-date=' + iso_(msg.getDate()) + '|subject=' + msg.getSubject();
      const atts = msg.getAttachments({ includeInlineImages: false })
        .filter(a => /^image\/|application\/pdf/.test(a.getContentType()));
      if (atts.length) {
        atts.forEach(a => root.createFile(a.copyBlob()).setDescription(hint));
      } else {
        // No attachment: the email itself is the receipt.
        const html = '<p><b>From:</b> ' + esc_(msg.getFrom()) + '<br><b>Date:</b> ' + msg.getDate() +
          '<br><b>Subject:</b> ' + esc_(msg.getSubject()) + '</p><hr>' + msg.getBody();
        const pdf = Utilities.newBlob(html, 'text/html', 'email.html').getAs('application/pdf')
          .setName('email ' + iso_(msg.getDate()) + '.pdf');
        root.createFile(pdf).setDescription(hint);
      }
    });
    thread.addLabel(done);
  });
}

// ---------------------------------------------------------------- drive

function processDrive_() {
  const start = new Date(PropertiesService.getScriptProperties().getProperty('START_AT') || 0);
  const root = DriveApp.getFolderById(CFG.ROOT_FOLDER_ID);
  const trips = loadTrips_();
  const places = [{ folder: root, trip: null }];
  const it = root.getFolders();
  while (it.hasNext()) {
    const f = it.next();
    if (!/^batch/i.test(f.getName())) places.push({ folder: f, trip: f.getName() });
  }
  let n = 0;
  for (const p of places) {
    const files = p.folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (!isCandidate_(file, start)) continue;
      fileReceipt_(file, p.folder, p.trip, trips, root);
      if (++n >= CFG.MAX_FILES_PER_RUN) return;
    }
  }
}

function isCandidate_(file, start) {
  const mime = file.getMimeType();
  if (!(/^image\//.test(mime) || mime === MimeType.PDF)) return false;
  if (file.getDateCreated() < start) return false;
  return (file.getDescription() || '').indexOf(CFG.MARK) !== 0;
}

function fileReceipt_(file, folder, folderTrip, trips, root) {
  const hint = file.getDescription() || '';
  let r;
  try {
    r = extract_(file);
  } catch (e) {
    console.error('Extraction failed for ' + file.getName() + ': ' + e);
    file.setDescription(CFG.MARK + ' error: ' + e);
    appendRow_({ description: 'COULD NOT READ: ' + file.getName() }, '', file, 'ERROR');
    return;
  }
  const emailDate = (hint.match(/email-date=(\d{4}-\d{2}-\d{2})/) || [])[1];
  if (!r.date && emailDate) r.date = emailDate;

  const trip = folderTrip || matchTrip_(r.service_date || r.date, trips) ||
               matchTrip_(r.date, trips) || CFG.NO_TRIP;
  let dest = folder;
  if (!folderTrip && trip !== CFG.NO_TRIP) dest = subfolder_(root, trip);

  file.setName(uniqueName_(dest, buildName_(r, trip, file), file.getId()));
  if (dest.getId() !== folder.getId()) file.moveTo(dest);
  file.setDescription(CFG.MARK + ' ' + JSON.stringify(r));

  const ok = r.confident && r.date && r.vendor && r.total > 0 && trip !== CFG.NO_TRIP;
  appendRow_(r, trip === CFG.NO_TRIP ? '' : trip, file, ok ? 'OK' : 'CHECK');
}

function buildName_(r, trip, file) {
  const ext = (file.getName().match(/\.[A-Za-z0-9]{2,5}$/) || [''])[0].toLowerCase() ||
              (file.getMimeType() === MimeType.PDF ? '.pdf' : '.jpg');
  const total = r.total > 0 ? r.total.toFixed(2) : '0.00';
  const parts = [r.date || '0000-00-00', trip, r.vendor || 'Unknown vendor', total];
  return parts.map(s => String(s).replace(/[\\/:*?"<>|]/g, '').trim()).join(' - ') + ext;
}

function uniqueName_(folder, name, selfId) {
  const base = name.replace(/(\.[^.]+)$/, '');
  const ext = (name.match(/(\.[^.]+)$/) || [''])[0];
  let candidate = name;
  for (let i = 2; ; i++) {
    const hits = folder.getFilesByName(candidate);
    let clash = false;
    while (hits.hasNext()) if (hits.next().getId() !== selfId) clash = true;
    if (!clash) return candidate;
    candidate = base + ' (' + i + ')' + ext;
  }
}

function subfolder_(root, name) {
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

// ---------------------------------------------------------------- extraction

function extract_(file) {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (key) {
    try {
      return extractWithClaude_(file, key);
    } catch (e) {
      console.warn('Claude extraction failed, falling back to OCR: ' + e);
    }
  }
  return extractWithOcr_(file);
}

const RECEIPT_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['date', 'service_date', 'vendor', 'category', 'description',
             'total', 'gst', 'pst', 'tip', 'confident'],
  properties: {
    date: { type: 'string', description: 'Charge/purchase date, YYYY-MM-DD; empty string if unreadable' },
    service_date: { type: 'string', description: 'First date of the stay/sailing/travel if different from the charge date (prepaid hotel, ferry booking), YYYY-MM-DD; else empty string' },
    vendor: { type: 'string', description: 'Short everyday business name, e.g. "Tim Hortons", "BC Ferries", "Kwa\'lilas Hotel"' },
    category: { type: 'string', enum: CFG.CATEGORIES },
    description: { type: 'string', description: 'A few words, e.g. "Lunch", "Ferry reso B267467068 (Comox-Powell River)", "Hotel, 2 nights Oct 7-9"' },
    total: { type: 'number', description: 'Final amount charged, including tax and tip' },
    gst: { type: 'number', description: 'GST, or the GST part of HST; 0 if none' },
    pst: { type: 'number', description: 'PST (include MRDT for hotels); 0 if none' },
    tip: { type: 'number', description: 'Tip/gratuity; 0 if none' },
    confident: { type: 'boolean', description: 'false if any field was guessed or the image is hard to read' },
  },
};

function extractWithClaude_(file, key) {
  let block;
  if (file.getMimeType() === MimeType.PDF) {
    block = { type: 'document', source: { type: 'base64', media_type: 'application/pdf',
              data: Utilities.base64Encode(file.getBlob().getBytes()) } };
  } else {
    const img = imageForApi_(file);
    block = { type: 'image', source: { type: 'base64', media_type: img.getContentType(),
              data: Utilities.base64Encode(img.getBytes()) } };
  }
  const body = {
    model: CFG.MODEL,
    max_tokens: 4000,
    output_config: { effort: 'low', format: { type: 'json_schema', schema: RECEIPT_SCHEMA } },
    fallbacks: 'default',
    messages: [{ role: 'user', content: [block, { type: 'text', text:
      'This is a business expense receipt from a sales rep travelling in British Columbia, Canada. ' +
      'Extract the fields. Use the date the card was charged as "date". ' +
      'Category: food/coffee/snacks = Meals; hotels = Accommodation; ferries, parking, ' +
      'taxis, flights, car rental = Transport; gas/diesel = Fuel; anything else = Other.' }] }],
  };
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': key,
      'anthropic-version': '2026-10-21',
      'anthropic-beta': 'server-side-fallback-2026-07-01',
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('API ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 300));
  }
  const msg = JSON.parse(res.getContentText());
  if (msg.stop_reason === 'refusal') throw new Error('refused');
  const text = msg.content.filter(b => b.type === 'text').map(b => b.text).join('');
  const r = JSON.parse(text);
  r.date = validDate_(r.date);
  r.service_date = validDate_(r.service_date);
  return r;
}

/** Claude takes JPEG/PNG/GIF/WebP up to ~5 MB. iPhone HEIC or big photos go via Drive's thumbnail. */
function imageForApi_(file) {
  const blob = file.getBlob();
  const ok = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'].indexOf(blob.getContentType()) >= 0;
  if (ok && blob.getBytes().length < 4.5 * 1024 * 1024) return blob;
  const meta = Drive.Files.get(file.getId(), { fields: 'thumbnailLink' });
  if (!meta.thumbnailLink) throw new Error('Drive has no preview for this image yet');
  return UrlFetchApp.fetch(meta.thumbnailLink.replace(/=s\d+$/, '=s1600'), {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
  }).getBlob();
}

/** No API key: Google's free OCR plus pattern matching. Rough; every row is marked CHECK. */
function extractWithOcr_(file) {
  const doc = Drive.Files.create({ name: 'tmp-ocr-' + file.getId(), mimeType: MimeType.GOOGLE_DOCS },
                                 file.getBlob(), { ocrLanguage: 'en' });
  let text;
  try {
    text = DocumentApp.openById(doc.id).getBody().getText();
  } finally {
    Drive.Files.remove(doc.id);
  }
  const lines = text.split(/\n/).map(s => s.trim()).filter(Boolean);
  const amt = re => {
    const vals = [];
    lines.forEach(l => { if (re.test(l)) (l.match(/\d{1,5}[.,]\d{2}/g) || []).forEach(v => vals.push(+v.replace(',', '.'))); });
    return vals;
  };
  const totals = amt(/total|amount|balance|paid/i).concat([]);
  const subtotalFree = amt(/^(?!.*sub).*total/i);
  const total = Math.max.apply(null, (subtotalFree.length ? subtotalFree : totals).concat([0]));
  const vendor = (lines.find(l => /[A-Za-z]{3,}/.test(l) && !/receipt|invoice|welcome|tel|www|http/i.test(l)) || '')
    .replace(/[^A-Za-z0-9&' .-]/g, '').replace(/\s+/g, ' ').trim().slice(0, 40);
  return {
    date: findDate_(text),
    service_date: '',
    vendor: titleCase_(vendor),
    category: 'Other',
    description: '',
    total: total,
    gst: Math.max.apply(null, amt(/\bGST|HST/i).concat([0])),
    pst: Math.max.apply(null, amt(/\bPST/i).concat([0])),
    tip: Math.max.apply(null, amt(/tip|gratuity/i).concat([0])),
    confident: false,
  };
}

function findDate_(text) {
  const M = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  let m = text.match(/\b(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})\b/);
  if (m) return validDate_(ymd_(+m[1], +m[2], +m[3]));
  m = text.match(/\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2}),?\s+(20\d{2})\b/i);
  if (m) return validDate_(ymd_(+m[3], M[m[1].toLowerCase()], +m[2]));
  m = text.match(/\b(\d{1,2})[\s-](jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*[\s-,]+(20\d{2}|\d{2})\b/i);
  if (m) return validDate_(ymd_(yr_(m[3]), M[m[2].toLowerCase()], +m[1]));
  m = text.match(/\b(\d{1,2})[\/-](\d{1,2})[\/-](20\d{2}|\d{2})\b/);
  if (m) {
    const a = +m[1], b = +m[2];
    return a > 12 ? validDate_(ymd_(yr_(m[3]), b, a)) : validDate_(ymd_(yr_(m[3]), a, b));
  }
  return '';
}

// ---------------------------------------------------------------- trips

function ensureTripsSheet_() {
  const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  if (ss.getSheetByName(CFG.TRIPS_SHEET)) return;
  const sh = ss.insertSheet(CFG.TRIPS_SHEET);
  const rows = [
    ['Trip', 'Start', 'End'],
    ['Port Hardy Aug 12-14 2026', '2026-08-12', '2026-08-14'],
    ['Powell River Aug 27', '2026-08-27', '2026-08-27'],
    ['Powell River Sep 2', '2026-09-02', '2026-09-02'],
    ['Campbell River Sep 11', '2026-09-11', '2026-09-11'],
    ['Powell River Sep 22', '2026-09-21', '2026-09-22'],
    ['Courtenay Sep 28', '2026-09-28', '2026-09-28'],
    ['Port Hardy Oct 7-9', '2026-10-07', '2026-10-09'],
    ['Powell River Nov 2', '2026-11-02', '2026-11-02'],
    ['Gold River Nov 5', '2026-11-05', '2026-11-05'],
    ['Campbell River Nov 17', '2026-11-17', '2026-11-17'],
  ];
  sh.getRange(1, 1, rows.length, 3).setValues(rows);
  sh.getRange('B:C').setNumberFormat('yyyy-mm-dd');
  sh.getRange('A1:C1').setFontWeight('bold');
  sh.setFrozenRows(1);
}

function loadTrips_() {
  const sh = SpreadsheetApp.openById(CFG.SPREADSHEET_ID).getSheetByName(CFG.TRIPS_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(r => r[0] && r[1])
    .map(r => ({ name: String(r[0]).trim(), start: toDate_(r[1]), end: toDate_(r[2] || r[1]) }));
}

function matchTrip_(dateStr, trips) {
  if (!dateStr) return null;
  const d = toDate_(dateStr).getTime();
  const slack = CFG.TRIP_SLACK_DAYS * 864e5;
  let best = null, bestGap = Infinity;
  trips.forEach(t => {
    const s = t.start.getTime(), e = t.end.getTime();
    if (d < s - slack || d > e + slack) return;
    const gap = d < s ? s - d : d > e ? d - e : 0;
    if (gap < bestGap) { best = t.name; bestGap = gap; }
  });
  return best;
}

// ---------------------------------------------------------------- sheet

// Columns on All Expenses: B Description, C Charge Date, D Vendor, E Category,
// F Tot. Amt., G Base Amt., H GST, I PST, J Tip, K Total, L Submitted, M Batch Tot.,
// N Submission #, O Submission Description, P Paid out, Q (batch paid), R Receipt, S Status
function ensureHeaders_() {
  const sh = SpreadsheetApp.openById(CFG.SPREADSHEET_ID).getSheetByName(CFG.EXPENSE_SHEET);
  const h = sh.getRange(CFG.HEADER_ROW, 18, 1, 2);
  if (!h.getValues()[0].join('')) h.setValues([['Receipt', 'Status']]);
}

function appendRow_(r, trip, file, status) {
  const sh = SpreadsheetApp.openById(CFG.SPREADSHEET_ID).getSheetByName(CFG.EXPENSE_SHEET);
  const colC = sh.getRange(CFG.HEADER_ROW + 1, 3, Math.max(sh.getLastRow() - CFG.HEADER_ROW, 1), 1).getValues();
  let last = CFG.HEADER_ROW;
  colC.forEach((v, i) => { if (v[0] !== '') last = CFG.HEADER_ROW + 1 + i; });
  const row = last + 1;

  const n = x => (typeof x === 'number' && x > 0 ? Math.round(x * 100) / 100 : '');
  const total = n(r.total), tip = n(r.tip), gst = n(r.gst), pst = n(r.pst);
  const preTip = total !== '' ? total - (tip || 0) : '';
  const base = preTip !== '' ? Math.round((preTip - (gst || 0) - (pst || 0)) * 100) / 100 : '';
  const link = '=HYPERLINK("' + file.getUrl() + '","' + file.getName().replace(/"/g, "'") + '")';

  sh.getRange(row, 2, 1, 18).setValues([[
    r.description || '', r.date ? toDate_(r.date) : '', r.vendor || '', r.category || '',
    preTip, base, gst, pst, tip, total,
    '', '', '', trip, '', '', link, status,
  ]]);
  sh.getRange(row, 3).setNumberFormat('yyyy-mm-dd');
  sh.getRange(row, 6, 1, 6).setNumberFormat('$#,##0.00');
  if (status !== 'OK') sh.getRange(row, 19).setBackground('#fde68a');
}

// ---------------------------------------------------------------- helpers

function label_(name) { return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name); }
function slug_(label) { return label.toLowerCase().replace(/[\s\/]+/g, '-'); }
function esc_(s) { return String(s).replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c])); }
function iso_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function ymd_(y, m, d) { return y + '-' + ('0' + m).slice(-2) + '-' + ('0' + d).slice(-2); }
function yr_(s) { return s.length === 2 ? 2000 + +s : +s; }
function titleCase_(s) { return s.toLowerCase().replace(/\b[a-z]/g, c => c.toUpperCase()); }
function validDate_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s || '')) return '';
  const d = toDate_(s);
  return isNaN(d) || d.getFullYear() < 2020 || d.getFullYear() > 2100 ? '' : s;
}
function toDate_(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const p = String(v).split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}
