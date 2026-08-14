/********************************************************************
 * RDCO SAMPLE TRACKER — V3.1 ("GitHub Pages" model)
 * --------------------------------------------------------------
 * Backend is a pure JSON API. The web page lives on GitHub Pages,
 * which removes the "created by a Google Apps Script user" banner.
 *
 * WHAT CHANGED IN V3.1
 *  - SampleLines gains an "Exclude" column. Any text in that cell
 *    removes the sample from the supplier's speed and accuracy
 *    scoring, and the text itself is the reason. Blank = counted.
 *    The sample still belongs to its project and still shows as
 *    on-order / late — only the SCORING ignores it.
 *
 *    !!! upsert_() writes by POSITION using the headers array
 *    below. 'Exclude' must be the LAST entry in TABS.LINES.headers
 *    and the LAST column (N) in the SampleLines tab. If the array
 *    order and the sheet column order ever disagree, every field
 *    gets overwritten with its neighbour's value. Never insert a
 *    column in the middle — always append.
 *
 * WHAT'S IN V3
 *  - doGet / doPost serve JSON (no more HtmlService)
 *  - LockService on every write (no collision risk on simultaneous saves)
 *  - apiReceive: batch "mark arrived" in ONE server call
 *  - Simple shared token (speed bump, not a vault)
 *
 * SETUP:
 *  1. Paste this over Code.gs and save
 *  2. In the SampleLines tab, put  Exclude  in cell N1
 *  3. Deploy > Manage deployments > edit (pencil) > New version
 *     - Execute as: Me
 *     - Who has access: Anyone
 *     (editing the EXISTING deployment keeps your same /exec URL)
 *  4. Push index.html to the GitHub Pages repo — done
 ********************************************************************/

const CONFIG = {
  API_TOKEN: 'rdco-sam-7k2x9v',                   // <-- mirror this in index.html
  DIGEST_EMAIL: 'mbalson@carpetcolourcentre.com', // Matt — sole recipient
  DIGEST_HOUR: 9,          // 9 AM daily digest
  AT_RISK_DAYS: 3,         // project "at risk" window before Needed By
  SLOW_MEDIAN_DAYS: 12     // supplier flagged slow at/above this median
};

/* First names to KEEP from the old Staff tab (migration only). */
const TEAM_KEEP = ['doug', 'matt', 'felicia', 'mark', 'kevin', 'rory'];

const TABS = {
  TEAM: { name: 'Team', headers: ['Name', 'Email'] },
  SUPPLIERS: { name: 'SuppliersV2', headers: ['ID', 'Name', 'Website', 'Username', 'Sample Dept Email', 'Notes', 'Created'] },
  REPS: { name: 'Reps', headers: ['ID', 'Supplier ID', 'Name', 'Email', 'Phone', 'Product Line', 'Notes'] },
  PROJECTS: { name: 'ProjectsV2', headers: ['ID', 'Project Name', 'Client', 'Salesperson', 'Needed By', 'Status', 'Notes', 'Created', 'Completed'] },
  LINES: { name: 'SampleLines', headers: ['ID', 'Project ID', 'Supplier ID', 'Rep ID', 'Product', 'SKU', 'Date Ordered', 'Expected Arrival', 'Status', 'Accuracy', 'Date Arrived', 'Date Delivered', 'Notes', 'Exclude'] }
};

/* ------------------------------------------------ core helpers --- */

function TZ_() { return Session.getScriptTimeZone() || 'America/Edmonton'; }
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function todayStr_() { return Utilities.formatDate(new Date(), TZ_(), 'yyyy-MM-dd'); }

function addDays_(iso, n) {
  const d = new Date(iso + 'T12:00:00');
  d.setDate(d.getDate() + n);
  return Utilities.formatDate(d, TZ_(), 'yyyy-MM-dd');
}

function fmt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ_(), 'yyyy-MM-dd');
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function tab_(def) {
  let sh = ss_().getSheetByName(def.name);
  if (!sh) {
    sh = ss_().insertSheet(def.name);
    sh.getRange(1, 1, 1, def.headers.length).setValues([def.headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* One-time helper: adds any header in the TABS definition that is
   missing from an EXISTING sheet, appended on the right. Safe to run
   more than once. Run this instead of typing N1 by hand if you'd
   rather the script do it. */
function ensureHeaders() {
  Object.keys(TABS).forEach(function (k) {
    const def = TABS[k];
    const sh = tab_(def);
    const width = Math.max(sh.getLastColumn(), 1);
    const have = sh.getRange(1, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); });
    def.headers.forEach(function (h, i) {
      if (have.indexOf(h) === -1) {
        sh.getRange(1, have.length + 1).setValue(h).setFontWeight('bold');
        have.push(h);
        Logger.log('Added header "%s" to %s', h, def.name);
      } else if (have.indexOf(h) !== i) {
        Logger.log('WARNING: "%s" is column %s in %s but position %s in the headers array. ' +
          'Writes are positional — fix this before saving anything.', h, have.indexOf(h) + 1, def.name, i + 1);
      }
    });
  });
  Logger.log('ensureHeaders complete.');
}

function read_(def) {
  const sh = tab_(def);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const head = vals[0].map(String);
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    if (vals[i].join('') === '') continue;
    const o = {};
    head.forEach(function (h, c) { o[h] = fmt_(vals[i][c]); });
    out.push(o);
  }
  return out;
}

function upsert_(def, obj) {
  const sh = tab_(def);
  const row = def.headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(obj.ID)) {
      sh.getRange(i + 1, 1, 1, def.headers.length).setValues([row]);
      return obj;
    }
  }
  sh.appendRow(row);
  return obj;
}

function deleteById_(def, id) {
  const sh = tab_(def);
  const vals = sh.getDataRange().getValues();
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][0]) === String(id)) sh.deleteRow(i + 1);
  }
}

function newId_(prefix) {
  return prefix + Date.now() + '_' + Math.floor(Math.random() * 10000);
}

/* Every write goes through here — prevents two simultaneous saves
   from stepping on each other. */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('The sheet is busy — try again in a moment.');
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ------------------------------------------------------ web API --- */
/* The frontend on GitHub Pages talks to these two entry points.
 *  GET  ?action=getAll&token=...          -> full dataset
 *  POST body: {token, action, payload}    -> writes
 * POST bodies are sent as text/plain from the browser (a "simple
 * request"), which avoids the CORS preflight Apps Script can't
 * answer. Never change that.                                       */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'getAll') {
    if (p.token !== CONFIG.API_TOKEN) return json_({ ok: false, error: 'Bad token' });
    return json_({ ok: true, data: getAllData_() });
  }
  return json_({ ok: true, service: 'RDCO Sample Tracker API', hint: 'The app lives on GitHub Pages.' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); }
  catch (err) { return json_({ ok: false, error: 'Bad request body' }); }
  if (req.token !== CONFIG.API_TOKEN) return json_({ ok: false, error: 'Bad token' });

  try {
    const out = withLock_(function () { return route_(req.action, req.payload || {}); });
    return json_({ ok: true, result: out });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function route_(action, o) {
  switch (action) {
    case 'saveProject':
      if (!o.ID) o.ID = newId_('P');
      if (!o.Created) o.Created = todayStr_();
      if (!o.Status) o.Status = 'Active';
      upsert_(TABS.PROJECTS, o); return o.ID;

    case 'saveLine':
      if (!o.ID) o.ID = newId_('L');
      if (!o.Status) o.Status = 'Ordered';
      if (o.Exclude === undefined) o.Exclude = '';
      upsert_(TABS.LINES, o); return o.ID;

    case 'saveSupplier':
      if (!o.ID) o.ID = newId_('S');
      if (!o.Created) o.Created = todayStr_();
      upsert_(TABS.SUPPLIERS, o); return o.ID;

    case 'saveRep':
      if (!o.ID) o.ID = newId_('R');
      upsert_(TABS.REPS, o); return o.ID;

    /* Warehouse batch receive: payload = { items: [{id, accuracy}] }
       Marks every line Arrived + stamps today, in ONE call.
       Reads the whole row first, so Exclude and Notes survive. */
    case 'receive': {
      const byId = {};
      read_(TABS.LINES).forEach(function (l) { byId[l.ID] = l; });
      const t = todayStr_();
      (o.items || []).forEach(function (it) {
        const l = byId[it.id];
        if (!l) return;
        l.Status = 'Arrived';
        l['Date Arrived'] = t;
        l.Accuracy = it.accuracy || '';
        upsert_(TABS.LINES, l);
      });
      return (o.items || []).length;
    }

    case 'delete':
      if (o.kind === 'line') deleteById_(TABS.LINES, o.id);
      else if (o.kind === 'rep') deleteById_(TABS.REPS, o.id);
      else if (o.kind === 'project') {
        read_(TABS.LINES)
          .filter(function (l) { return l['Project ID'] === o.id; })
          .forEach(function (l) { deleteById_(TABS.LINES, l.ID); });
        deleteById_(TABS.PROJECTS, o.id);
      }
      return 'ok';

    default:
      throw new Error('Unknown action: ' + action);
  }
}

function getAllData_() {
  return {
    team: read_(TABS.TEAM),
    suppliers: read_(TABS.SUPPLIERS),
    reps: read_(TABS.REPS),
    projects: read_(TABS.PROJECTS),
    lines: read_(TABS.LINES),
    today: todayStr_()
  };
}

/* ------------------------------------------------- daily digest --- */

/** Run ONCE. Wipes every old trigger and installs the single
 *  9 AM digest. */
function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyDigest').timeBased().atHour(CONFIG.DIGEST_HOUR).everyDays(1).create();
  Logger.log('Done — all old triggers removed. Daily digest set for %s:00, sent only to %s.',
    CONFIG.DIGEST_HOUR, CONFIG.DIGEST_EMAIL);
}

/** Sends the digest right now so you can preview it. */
function testDigest() {
  const sent = dailyDigest();
  Logger.log(sent ? 'Digest sent to ' + CONFIG.DIGEST_EMAIL : 'Nothing overdue or at risk — no email sent.');
}

function dailyDigest() {
  const suppliers = read_(TABS.SUPPLIERS);
  const reps = read_(TABS.REPS);
  const projects = read_(TABS.PROJECTS);
  const lines = read_(TABS.LINES);
  const today = todayStr_();

  const supById = {}; suppliers.forEach(function (s) { supById[s.ID] = s; });
  const projById = {}; projects.forEach(function (p) { projById[p.ID] = p; });
  const repsBySup = {}; reps.forEach(function (r) {
    (repsBySup[r['Supplier ID']] = repsBySup[r['Supplier ID']] || []).push(r);
  });

  function contactFor(supId) {
    const s = supById[supId];
    if (!s) return '';
    if (s['Sample Dept Email']) return s['Sample Dept Email'];
    const rr = repsBySup[supId] || [];
    return rr.length && rr[0].Email ? rr[0].Email : '';
  }
  function daysLate(exp) {
    return Math.round((new Date(today + 'T12:00') - new Date(exp + 'T12:00')) / 86400000);
  }

  const overdue = lines.filter(function (l) {
    return l.Status === 'Ordered' && l['Expected Arrival'] && l['Expected Arrival'] < today;
  }).sort(function (a, b) { return a['Expected Arrival'] < b['Expected Arrival'] ? -1 : 1; });

  const riskLimit = addDays_(today, CONFIG.AT_RISK_DAYS);
  const atRisk = projects.filter(function (p) {
    if (p.Status !== 'Active' || !p['Needed By'] || p['Needed By'] > riskLimit) return false;
    return lines.some(function (l) { return l['Project ID'] === p.ID && l.Status === 'Ordered'; });
  });

  if (!overdue.length && !atRisk.length) return false;

  let html = '<div style="font-family:Georgia,serif;max-width:640px;margin:0 auto;color:#463340;">' +
    '<div style="background:#F8E8EC;border-radius:14px;padding:18px 24px;margin-bottom:18px;">' +
    '<h2 style="margin:0;color:#A84A5E;">Sample Tracker — Morning Digest</h2>' +
    '<p style="margin:6px 0 0;font-size:14px;">' + today + ' · ' + overdue.length +
    ' overdue · ' + atRisk.length + ' project(s) at risk</p></div>';

  if (overdue.length) {
    html += '<h3 style="color:#A84A5E;">⏰ Overdue samples</h3>' +
      '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
      '<tr style="background:#FBF6F2;text-align:left;">' +
      '<th style="padding:8px;">Project</th><th style="padding:8px;">Product</th>' +
      '<th style="padding:8px;">Supplier</th><th style="padding:8px;">Expected</th>' +
      '<th style="padding:8px;">Late</th><th style="padding:8px;">Chase</th></tr>';
    overdue.forEach(function (l) {
      const p = projById[l['Project ID']] || {};
      const s = supById[l['Supplier ID']] || {};
      const c = contactFor(l['Supplier ID']);
      html += '<tr style="border-bottom:1px solid #F0E2E6;">' +
        '<td style="padding:8px;">' + (p['Project Name'] || '—') + '</td>' +
        '<td style="padding:8px;">' + (l.Product || '—') + (l.SKU ? ' (' + l.SKU + ')' : '') + '</td>' +
        '<td style="padding:8px;">' + (s.Name || '—') + '</td>' +
        '<td style="padding:8px;">' + l['Expected Arrival'] + '</td>' +
        '<td style="padding:8px;color:#C24A4A;font-weight:bold;">' + daysLate(l['Expected Arrival']) + 'd</td>' +
        '<td style="padding:8px;">' + (c ? '<a href="mailto:' + c + '">' + c + '</a>' : '—') + '</td></tr>';
    });
    html += '</table>';
  }

  if (atRisk.length) {
    html += '<h3 style="color:#C39B5E;margin-top:22px;">⚠️ Projects at risk (needed-by date close)</h3><ul style="font-size:13px;">';
    atRisk.forEach(function (p) {
      const waiting = lines.filter(function (l) { return l['Project ID'] === p.ID && l.Status === 'Ordered'; }).length;
      html += '<li style="margin-bottom:6px;"><b>' + p['Project Name'] + '</b> — needed by ' +
        p['Needed By'] + ', still waiting on ' + waiting + ' sample(s)</li>';
    });
    html += '</ul>';
  }

  html += '<p style="font-size:12px;color:#93798A;margin-top:24px;">Sent automatically by the RDCO Sample Tracker.</p></div>';

  MailApp.sendEmail({
    to: CONFIG.DIGEST_EMAIL,
    subject: 'Samples: ' + overdue.length + ' overdue, ' + atRisk.length + ' project(s) at risk — ' + today,
    htmlBody: html
  });
  return true;
}

/* ---------------------------------------------------- migration --- */
/* Kept so resetV2Tabs / migrateData still work if you ever need to
   re-run against the old V1 tabs. */

function readLoose_(name, keyHeader) {
  const sh = ss_().getSheetByName(name);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (!vals.length) return [];
  const key = keyHeader.toLowerCase();
  let hr = -1;
  for (let i = 0; i < Math.min(6, vals.length); i++) {
    if (vals[i].some(function (c) { return clean_(c).indexOf(key) === 0; })) { hr = i; break; }
  }
  if (hr === -1) return [];
  const head = vals[hr].map(clean_);
  const out = [];
  for (let i = hr + 1; i < vals.length; i++) {
    if (vals[i].join('') === '') continue;
    const o = {};
    head.forEach(function (h, c) { if (h) o[h] = fmt_(vals[i][c]); });
    out.push(o);
  }
  return out;
}

function clean_(h) {
  return String(h === null || h === undefined ? '' : h)
    .replace(/\n/g, ' ').replace(/\(.*?\)/g, '').replace(/\s+/g, ' ')
    .trim().toLowerCase();
}

function pick_(row, candidates) {
  const keys = Object.keys(row);
  for (let c = 0; c < candidates.length; c++) {
    for (let k = 0; k < keys.length; k++) {
      if (keys[k].indexOf(candidates[c]) === 0 && row[keys[k]] !== '') return row[keys[k]];
    }
  }
  return '';
}

function normStatus_(s) {
  const v = String(s || '').toLowerCase();
  if (v.indexOf('deliver') !== -1) return 'Delivered';
  if (v.indexOf('arriv') !== -1) return 'Arrived';
  return 'Ordered';
}

function isJunk_(v) {
  return !v || v.charAt(0) === '←' || v.indexOf('Add new') === 0;
}

function resetV2Tabs() {
  Object.keys(TABS).forEach(function (k) {
    const sh = ss_().getSheetByName(TABS[k].name);
    if (sh) ss_().deleteSheet(sh);
  });
  Logger.log('V2 tabs removed. Old tabs untouched. Run migrateData next.');
}

function migrateData() {
  if (read_(TABS.PROJECTS).length > 0 || read_(TABS.LINES).length > 0) {
    throw new Error('Migration aborted — V2 tabs already contain data. Run resetV2Tabs first to re-run.');
  }
  Object.keys(TABS).forEach(function (k) { tab_(TABS[k]); });
  const summary = [];

  const oldStaff = readLoose_('Staff', 'name');
  const teamRows = [];
  oldStaff.forEach(function (r) {
    const full = pick_(r, ['name']);
    if (!full) return;
    const first = full.trim().split(' ')[0].toLowerCase();
    if (TEAM_KEEP.indexOf(first) === -1) return;
    let email = pick_(r, ['email']);
    if (/\.boz$/i.test(email)) email = email.replace(/\.boz$/i, '.biz');
    if (first === 'rory') email = CONFIG.DIGEST_EMAIL;
    teamRows.push([full.trim(), email]);
  });
  if (teamRows.length) tab_(TABS.TEAM).getRange(2, 1, teamRows.length, 2).setValues(teamRows);
  summary.push('Team: ' + teamRows.length + ' members');

  const oldSup = readLoose_('Suppliers', 'supplier name');
  const supMap = {};
  const firstRepBySup = {};
  const supRows = [], repRows = [];
  let si = 0, ri = 0;

  oldSup.forEach(function (r) {
    const name = pick_(r, ['supplier name']);
    if (isJunk_(name) || supMap[name.toLowerCase()]) return;
    const id = 'S_mig_' + (si++);
    supMap[name.toLowerCase()] = id;

    let sampleEmail = pick_(r, ['sample dept email', 'sample email']);
    let notes = pick_(r, ['notes']);
    if (sampleEmail && sampleEmail.indexOf('@') === -1) {
      notes = (notes ? notes + ' · ' : '') + sampleEmail;
      sampleEmail = '';
    }
    const prodLine = pick_(r, ['product line']);
    if (prodLine) notes = (notes ? notes + ' · ' : '') + 'Lines: ' + prodLine;

    supRows.push([id, name, '', '', sampleEmail, notes, todayStr_()]);

    const repName = pick_(r, ['rep name']);
    const repEmail = pick_(r, ['rep email']);
    if (repName || repEmail) {
      const rid = 'R_mig_' + (ri++);
      firstRepBySup[id] = rid;
      repRows.push([rid, id, repName, repEmail, pick_(r, ['rep phone']), prodLine, '']);
    }
  });
  if (supRows.length) tab_(TABS.SUPPLIERS).getRange(2, 1, supRows.length, TABS.SUPPLIERS.headers.length).setValues(supRows);
  if (repRows.length) tab_(TABS.REPS).getRange(2, 1, repRows.length, TABS.REPS.headers.length).setValues(repRows);
  summary.push('Suppliers: ' + supRows.length + ', Reps: ' + repRows.length);

  const oldProj = readLoose_('Projects', 'project id');
  const projRows = [];
  const projSeen = {};
  oldProj.forEach(function (r) {
    const id = pick_(r, ['project id']);
    const name = pick_(r, ['project name']);
    if (!id || !name || projSeen[id]) return;
    projSeen[id] = true;
    const status = String(pick_(r, ['project status', 'status'])).toLowerCase().indexOf('comp') !== -1 ? 'Complete' : 'Active';
    projRows.push([id, name, pick_(r, ['customer', 'client']),
      pick_(r, ['salesperson']), '', status,
      pick_(r, ['notes']), pick_(r, ['date created']) || todayStr_(), '']);
  });

  const oldOrders = readLoose_('Orders', 'order id');
  const newSupSheet = tab_(TABS.SUPPLIERS);
  const lineRows = [];
  oldOrders.forEach(function (r) {
    const oid = pick_(r, ['order id']);
    if (!oid) return;
    const pid = pick_(r, ['project id']);
    if (!pid || !projSeen[pid]) return;

    let supName = pick_(r, ['supplier']);
    if (isJunk_(supName)) supName = '';
    let supId = supName ? supMap[supName.toLowerCase()] : '';
    if (supName && !supId) {
      supId = 'S_mig_' + (si++);
      supMap[supName.toLowerCase()] = supId;
      newSupSheet.appendRow([supId, supName, '', '', '', 'Added during migration (found on an order)', todayStr_()]);
    }

    let sku = pick_(r, ['sku']);
    const colour = pick_(r, ['colour', 'color']);
    if (colour) sku = sku ? sku + ' · ' + colour : colour;

    let notes = pick_(r, ['notes']);
    const qty = parseFloat(pick_(r, ['qty'])) || 1;
    if (qty > 1) notes = 'Qty ' + qty + (notes ? ' — ' + notes : '');

    /* trailing '' is the Exclude column — the row width must match
       TABS.LINES.headers.length or setValues() below throws. */
    lineRows.push([oid, pid, supId || '', firstRepBySup[supId] || '',
      pick_(r, ['product name', 'product']), sku,
      pick_(r, ['date ordered']), pick_(r, ['expected arrival', 'expected']),
      normStatus_(pick_(r, ['status'])), '',
      pick_(r, ['date arrived']), pick_(r, ['date delivered']), notes, '']);
  });

  projRows.forEach(function (p) {
    const mine = lineRows.filter(function (l) { return l[1] === p[0]; });
    if (mine.length && mine.every(function (l) { return l[8] === 'Delivered'; })) {
      p[5] = 'Complete';
      p[8] = mine.reduce(function (m, l) { return l[11] > m ? l[11] : m; }, '');
    }
  });

  if (projRows.length) tab_(TABS.PROJECTS).getRange(2, 1, projRows.length, TABS.PROJECTS.headers.length).setValues(projRows);
  if (lineRows.length) tab_(TABS.LINES).getRange(2, 1, lineRows.length, TABS.LINES.headers.length).setValues(lineRows);
  summary.push('Projects: ' + projRows.length + ' (' + oldOrders.length + ' old orders -> ' + lineRows.length + ' sample lines)');

  const msg = 'MIGRATION COMPLETE\n' + summary.join('\n') +
    '\nOld tabs were not touched — hide them once you have verified the new data.';
  Logger.log(msg);
  return msg;
}
