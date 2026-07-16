/**
 * BOLLIN CLINIC INVENTORY — Apps Script backend  · v0.2
 * ------------------------------------------------------
 * Attach this to the Bollin_Inventory_Backend Google Sheet:
 *   Extensions → Apps Script → paste this file → Deploy → New deployment
 *   → type: Web app → Execute as: Me → Who has access: Anyone → Deploy.
 * Copy the /exec URL into CONFIG.API_URL in index.html.
 *
 * Alerts: run setupTriggers() once from the editor (Run ▶) to install
 * the daily 08:00 stock check. Fill alert_email / escalation_email in
 * the Settings tab first.
 */

const ITEM_TABS = ['Consumables', 'Meds', 'Garments', 'Instruments', 'Linen'];

/* ============================ HTTP entry points ============================ */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.ack) return ackFromEmail_(p.ack);          // one-click ack link in emails
  return json_(getAll_());
}

// action -> minimum role. Roles: common < staff < admin
const ACTION_ROLE = {
  getAll:'common', move:'common', link:'common', useLog:'common', request:'common',
  changePassword:'common',
  stocktake:'staff', respond:'staff', ack:'staff',
  implantAdd:'staff', implantUpdate:'staff',
  dispatchAssign:'staff', dispatchReceive:'staff', dispatchAdd:'staff', dispatchReturn:'staff',
  addItem:'admin', updateItem:'admin', setBarcodes:'admin'
};
const ROLE_RANK = { common: 0, staff: 1, admin: 2 };

function doPost(e) {
  let req = {};
  try { req = JSON.parse(e.postData.contents); } catch (err) {
    return json_({ ok: false, error: 'Bad JSON' });
  }
  const lock = LockService.getScriptLock();        // serialise writes
  lock.waitLock(20000);
  try {
    if (req.action === 'login') return json_(login_(req));
    const session = auth_(req.token);
    if (!session) return json_({ ok: false, error: 'AUTH', message: 'Please sign in' });
    const need = ACTION_ROLE[req.action] || 'admin';
    if (ROLE_RANK[session.role] < ROLE_RANK[need])
      return json_({ ok: false, error: 'Your access level (' + session.role + ') cannot do this' });
    req._session = session;
    switch (req.action) {
      case 'getAll':    return json_(getAll_(req._session));
      case 'changePassword': return json_(changePassword_(req));
      case 'move':      return json_(move_(req));
      case 'link':      return json_(link_(req));
      case 'addItem':   return json_(addItem_(req));
      case 'updateItem':return json_(updateItem_(req));
      case 'setBarcodes':return json_(setBarcodes_(req));
      case 'request':   return json_(addRequest_(req));
      case 'respond':   return json_(respondRequest_(req));
      case 'implantAdd':   return json_(implantAdd_(req));
      case 'implantUpdate':return json_(implantUpdate_(req));
      case 'dispatchAdd':   return json_(dispatchAdd_(req));
      case 'dispatchReturn':return json_(dispatchReturn_(req));
      case 'useLog':         return json_(useLog_(req));
      case 'dispatchAssign': return json_(dispatchAssign_(req));
      case 'dispatchReceive':return json_(dispatchReceive_(req));
      case 'ack':       return json_(ackAlert_(req));
      case 'stocktake': return json_(stocktake_(req));
      default:          return json_({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================================ helpers ================================= */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function readTab_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const head = vals[0].map(String);
  return vals.slice(1)
    .map((row, i) => {
      const o = { _row: i + 2 };                    // sheet row number
      head.forEach((h, c) => o[h] = row[c] === '' ? null : row[c]);
      return o;
    })
    .filter(o => Object.keys(o).some(k => k !== '_row' && o[k] !== null && o[k] !== undefined));
}

function findItem_(tracker, code) {
  const sh = ss_().getSheetByName(tracker);
  if (!sh) throw 'No tab: ' + tracker;
  const vals = sh.getDataRange().getValues();
  const head = vals[0].map(String);
  const cCode = head.indexOf('Code'), cBar = head.indexOf('Barcode'), cName = head.indexOf('Name');
  const target = String(code == null ? '' : code).trim();
  if (!target) throw 'No identifier given';
  let byName = null;
  for (let r = 1; r < vals.length; r++) {
    const vc = String(vals[r][cCode] == null ? '' : vals[r][cCode]).trim();
    const vb = cBar >= 0 ? String(vals[r][cBar] == null ? '' : vals[r][cBar]).trim() : '';
    if ((vc && vc === target) || (vb && vb === target)) {
      return { sh: sh, row: r + 1, head: head, vals: vals[r] };
    }
    if (cName >= 0 && !byName) {
      const vn = String(vals[r][cName] == null ? '' : vals[r][cName]).trim();
      if (vn && vn === target) byName = { sh: sh, row: r + 1, head: head, vals: vals[r] };
    }
  }
  if (byName) return byName;                       // fallback: exact name match
  throw 'Item not found: ' + code + ' in ' + tracker;
}

function addItem_(q) {
  // q: {tracker, fields:{Code,Barcode,Name,...}} — mapped by header name
  const sh = ss_().getSheetByName(q.tracker);
  if (!sh) return { ok: false, error: 'No tab: ' + q.tracker };
  if (!q.fields || !q.fields.Name) return { ok: false, error: 'Name required' };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const row = head.map(h => {
    let v = q.fields[h];
    if (v === undefined || v === null) return '';
    if (h === 'Expiry' && v) return new Date(v);
    return v;
  });
  sh.appendRow(row);
  if (q.fields.Barcode) {
    ss_().getSheetByName('BarcodeLinks')
      .appendRow([q.fields.Barcode, q.tracker, q.fields.Code || q.fields.Name]);
  }
  return { ok: true, row: sh.getLastRow() };
}

function setBarcodes_(q) {
  // q: {entries:[{tracker, code, barcode}]} — batch assign generated codes
  const links = ss_().getSheetByName('BarcodeLinks');
  let done = 0, errors = [];
  (q.entries || []).forEach(en => {
    try {
      const it = findItem_(en.tracker, en.code);
      it.sh.getRange(it.row, col_(it.head, 'Barcode')).setValue(en.barcode);
      links.appendRow([en.barcode, en.tracker, en.code]);
      done++;
    } catch (err) { errors.push(en.code + ': ' + err); }
  });
  return { ok: true, done: done, errors: errors.slice(0, 10) };
}

function itemAtRow_(tracker, row) {
  const sh = ss_().getSheetByName(tracker);
  if (!sh) throw 'No tab: ' + tracker;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  return { sh: sh, row: row, head: head, vals: vals };
}

function col_(head, name) {
  const i = head.indexOf(name);
  if (i < 0) throw 'Missing column ' + name;
  return i + 1;
}

function setting_(key, fallback) {
  const row = readTab_('Settings').find(r => r.Key === key);
  return row && row.Value != null && row.Value !== '' ? row.Value : fallback;
}

/* ================================ actions ================================= */

function getAll_(session) {
  const items = {};
  const hideCost = session && session.role === 'common';
  ITEM_TABS.forEach(t => {
    items[t] = readTab_(t);
    if (hideCost) items[t].forEach(r => { delete r.UnitCost; });
  });
  const tx = readTab_('Transactions');
  return {
    ok: true,
    role: session ? session.role : 'admin',
    displayName: session ? session.name : '',
    items: items,
    transactions: tx.slice(-200).reverse(),
    requests: readTab_('Requests').reverse(),
    alerts: readTab_('Alerts').slice(-100).reverse(),
    implants: readTab_('Implants').reverse(),
    dispatch: readTab_('DispatchLog').reverse(),
    barcodeLinks: readTab_('BarcodeLinks'),
    serverTime: new Date().toISOString()
  };
}

function move_(q) {
  // q: {tracker, code, row?, dir:'in'|'out', qty, by, batch, expiry, note}
  const it = q.row ? itemAtRow_(q.tracker, Number(q.row)) : findItem_(q.tracker, q.code);
  const qty = Math.max(1, Number(q.qty) || 1);
  let newQty = null;

  if (q.tracker === 'Instruments') {
    const c = col_(it.head, 'Status');
    it.sh.getRange(it.row, c).setValue(q.dir === 'out' ? 'At Countess' : 'At Bollin');
    if (q.dir === 'in') {                          // a return = one sterilisation cycle
      const cc = it.head.indexOf('CyclesToDate');
      if (cc >= 0) {
        const cur = Number(it.vals[cc]) || 0;
        it.sh.getRange(it.row, cc + 1).setValue(cur + 1);
      }
    }
  } else {
    const cQty = col_(it.head, 'Qty');
    const cur = Number(it.vals[cQty - 1]) || 0;
    newQty = q.dir === 'in' ? cur + qty : cur - qty;
    if (newQty < 0) return { ok: false, error: 'Only ' + cur + ' in stock' };
    it.sh.getRange(it.row, cQty).setValue(newQty);
    if (q.dir === 'in') {
      if (q.batch)  it.sh.getRange(it.row, col_(it.head, 'Batch')).setValue(q.batch);
      if (q.expiry && it.head.indexOf('Expiry') >= 0)
        it.sh.getRange(it.row, col_(it.head, 'Expiry')).setValue(new Date(q.expiry));
    }
  }

  const name = it.vals[it.head.indexOf('Name')];
  ss_().getSheetByName('Transactions').appendRow([
    new Date(), q.dir, q.tracker, q.code, name, qty,
    q.by || 'Unknown', q.batch || '', q.expiry || '', q.note || ''
  ]);
  return { ok: true, newQty: newQty, name: name };
}

function link_(q) {
  // q: {barcode, tracker, code}
  ss_().getSheetByName('BarcodeLinks').appendRow([q.barcode, q.tracker, q.code]);
  const it = findItem_(q.tracker, q.code);
  const cBar = col_(it.head, 'Barcode');
  if (!it.vals[cBar - 1]) it.sh.getRange(it.row, cBar).setValue(q.barcode);
  return { ok: true };
}

function addRequest_(q) {
  // Requests schema: Timestamp | By | Item | Qty | Size | Status | HandledBy | DateResponded | Remarks
  ss_().getSheetByName('Requests').appendRow([
    new Date(), q.by || 'Unknown', q.item,
    Math.max(1, Number(q.qty) || 1), q.size || '', 'Requested', '', '', ''
  ]);
  return { ok: true };
}

function respondRequest_(q) {
  // q: {row, status, by, remarks}
  const sh = ss_().getSheetByName('Requests');
  sh.getRange(q.row, 6).setValue(q.status);
  sh.getRange(q.row, 7).setValue(q.by || '');
  sh.getRange(q.row, 8).setValue(new Date());
  sh.getRange(q.row, 9).setValue(q.remarks || '');
  return { ok: true };
}

function updateItem_(q) {
  // q: {tracker, row, fields:{header:value,...}} — updates one item row in place
  const sh = ss_().getSheetByName(q.tracker);
  if (!sh) return { ok: false, error: 'No tab: ' + q.tracker };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  Object.keys(q.fields || {}).forEach(h => {
    const c = head.indexOf(h);
    if (c < 0) return;
    let v = q.fields[h];
    if (v === null || v === undefined) v = '';
    if (h === 'Expiry' && v) v = new Date(v);
    sh.getRange(q.row, c + 1).setValue(v);
  });
  return { ok: true };
}

/* ============================== implants ================================= */

const IMPLANT_HEADERS = ['Timestamp','PatientInitials','PATNumber','Surgeon','SurgeryDate',
  'ImplantDetails','Qty','Remarks','DateOrdered','OrderedBy','Status','ReceivedBy','ReceivedDate',
  'ReceivedQty','StorageLocation','ScannedRef','ReturnDetails','ReturnedQty','ReturnedBy','ReturnedDate','Notes'];

function implantAdd_(q) {
  const sh = ss_().getSheetByName('Implants');
  if (!sh) return { ok: false, error: 'Implants tab missing — run migrate() in the script editor' };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const f = q.fields || {};
  f.Timestamp = new Date();
  const row = head.map(h => {
    let v = f[h];
    if (v === undefined || v === null) return '';
    if (h.indexOf('Date') >= 0 && v) return new Date(v);
    return v;
  });
  sh.appendRow(row);
  return { ok: true, row: sh.getLastRow() };
}

function implantUpdate_(q) {
  // q: {row, fields:{Status, ReceivedBy, ReceivedDate, ScannedRef, ReturnDetails, ReturnedBy, ReturnedDate,...}}
  const sh = ss_().getSheetByName('Implants');
  if (!sh) return { ok: false, error: 'Implants tab missing — run migrate()' };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  Object.keys(q.fields || {}).forEach(h => {
    const c = head.indexOf(h);
    if (c < 0) return;
    let v = q.fields[h];
    if (v === null || v === undefined) v = '';
    if ((h.indexOf('Date') >= 0) && v) v = new Date(v);
    sh.getRange(q.row, c + 1).setValue(v);
  });
  return { ok: true };
}

/* ============================ dispatch log =============================== */

const DISPATCH_HEADERS = ['DispatchRef','DateSent','SentBy','ItemName','Code','CountOut',
  'BatchLot','Instructions','CollectedBy','DateCollected','CountessRef',
  'DateReturned','ReturnedBy','CountIn','Status','DateUsed','UsedBy'];

function dlSheet_() {
  const sh = ss_().getSheetByName('DispatchLog');
  if (!sh) throw 'DispatchLog tab missing — run migrate()';
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  return { sh: sh, head: head, col: n => head.indexOf(n) + 1 };
}

function useLog_(q) {
  // q: {by, date, lines:[{code, name, qty}]} -> rows with Status 'Pending Collection'
  const d = dlSheet_();
  const rows = (q.lines || []).map(l => {
    const r = new Array(d.head.length).fill('');
    r[d.col('ItemName') - 1] = l.name || '';
    r[d.col('Code') - 1] = l.code || '';
    r[d.col('CountOut') - 1] = Math.max(1, Number(l.qty) || 1);
    r[d.col('Status') - 1] = 'Pending Collection';
    r[d.col('DateUsed') - 1] = q.date ? new Date(q.date) : new Date();
    r[d.col('UsedBy') - 1] = q.by || '';
    return r;
  });
  if (!rows.length) return { ok: false, error: 'No lines' };
  d.sh.getRange(d.sh.getLastRow() + 1, 1, rows.length, d.head.length).setValues(rows);
  return { ok: true, added: rows.length };
}

function dispatchAssign_(q) {
  // q: {rows:[sheetRow,...], ref, date, by, batch, instructions}
  const d = dlSheet_();
  (q.rows || []).forEach(r => {
    d.sh.getRange(r, d.col('DispatchRef')).setValue(q.ref || '');
    d.sh.getRange(r, d.col('DateSent')).setValue(q.date ? new Date(q.date) : new Date());
    d.sh.getRange(r, d.col('SentBy')).setValue(q.by || '');
    if (q.batch) d.sh.getRange(r, d.col('BatchLot')).setValue(q.batch);
    if (q.instructions) d.sh.getRange(r, d.col('Instructions')).setValue(q.instructions);
    d.sh.getRange(r, d.col('Status')).setValue('Out');
  });
  return { ok: true, dispatched: (q.rows || []).length };
}

function dispatchReceive_(q) {
  // q: {rows:[sheetRow,...], by}
  const d = dlSheet_();
  (q.rows || []).forEach(r => {
    const out = Number(d.sh.getRange(r, d.col('CountOut')).getValue()) || 0;
    d.sh.getRange(r, d.col('DateReturned')).setValue(new Date());
    d.sh.getRange(r, d.col('ReturnedBy')).setValue(q.by || '');
    d.sh.getRange(r, d.col('CountIn')).setValue(out);
    d.sh.getRange(r, d.col('Status')).setValue('Returned');
  });
  return { ok: true, received: (q.rows || []).length };
}

function dispatchAdd_(q) {
  // q: {ref, date, by, batch, lines:[{code, name, qtyOut, instructions}]}
  const sh = ss_().getSheetByName('DispatchLog');
  if (!sh) return { ok: false, error: 'DispatchLog tab missing — run migrate()' };
  const rows = (q.lines || []).map(l => [
    q.ref || '', q.date ? new Date(q.date) : new Date(), q.by || '',
    l.name || '', l.code || '', Math.max(1, Number(l.qtyOut) || 1),
    q.batch || '', l.instructions || '', '', '', '', '', '', '', 'Out'
  ]);
  if (!rows.length) return { ok: false, error: 'No lines' };
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, DISPATCH_HEADERS.length).setValues(rows);
  return { ok: true, added: rows.length };
}

function dispatchReturn_(q) {
  // q: {row, qtyIn, by}
  const sh = ss_().getSheetByName('DispatchLog');
  if (!sh) return { ok: false, error: 'DispatchLog tab missing — run migrate()' };
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const get = n => head.indexOf(n) + 1;
  const qtyOut = Number(sh.getRange(q.row, get('CountOut')).getValue()) || 0;
  const qtyIn = Number(q.qtyIn) || qtyOut;
  sh.getRange(q.row, get('DateReturned')).setValue(new Date());
  sh.getRange(q.row, get('ReturnedBy')).setValue(q.by || '');
  sh.getRange(q.row, get('CountIn')).setValue(qtyIn);
  sh.getRange(q.row, get('Status')).setValue(qtyIn >= qtyOut ? 'Returned' : 'Partial');
  return { ok: true };
}

/* ============================== migration ================================ */
/** Run once from the editor after pasting this version (Run ▶ migrate). */
function migrate() {
  const ss = ss_();
  // 1. Implants tab (create, or append any missing columns)
  let imp = ss.getSheetByName('Implants');
  if (!imp) {
    imp = ss.insertSheet('Implants');
    imp.getRange(1, 1, 1, IMPLANT_HEADERS.length).setValues([IMPLANT_HEADERS])
      .setFontWeight('bold').setBackground('#2B6168').setFontColor('#FFFFFF');
    imp.setFrozenRows(1);
  } else {
    const have = imp.getRange(1, 1, 1, Math.max(imp.getLastColumn(), 1)).getValues()[0].map(String);
    IMPLANT_HEADERS.forEach(hd => {
      if (have.indexOf(hd) < 0) {
        const c2 = imp.getLastColumn() + 1;
        imp.getRange(1, c2).setValue(hd)
          .setFontWeight('bold').setBackground('#2B6168').setFontColor('#FFFFFF');
        have.push(hd);
      }
    });
  }
  // 1b. DispatchLog tab (sterilisation) — create, or append any missing columns
  let dl = ss.getSheetByName('DispatchLog');
  if (!dl) {
    dl = ss.insertSheet('DispatchLog');
    dl.getRange(1, 1, 1, DISPATCH_HEADERS.length).setValues([DISPATCH_HEADERS])
      .setFontWeight('bold').setBackground('#2B6168').setFontColor('#FFFFFF');
    dl.setFrozenRows(1);
  } else {
    const haveD = dl.getRange(1, 1, 1, Math.max(dl.getLastColumn(), 1)).getValues()[0].map(String);
    DISPATCH_HEADERS.forEach(hd => {
      if (haveD.indexOf(hd) < 0) {
        const c3 = dl.getLastColumn() + 1;
        dl.getRange(1, c3).setValue(hd)
          .setFontWeight('bold').setBackground('#2B6168').setFontColor('#FFFFFF');
        haveD.push(hd);
      }
    });
  }
  // 1c. Users tab
  if (!ss.getSheetByName('Users')) {
    const us = ss.insertSheet('Users');
    us.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS])
      .setFontWeight('bold').setBackground('#2B6168').setFontColor('#FFFFFF');
    us.setFrozenRows(1);
    us.getRange(2, 1, 3, USER_HEADERS.length).setValues([
      ['yasar', 'ChangeMe123', 'admin', 'Yasar', 'yes'],
      ['stockteam', 'ChangeMe123', 'staff', 'Stock Team', 'yes'],
      ['scrubs', 'ChangeMe123', 'common', 'Scrub Team', 'yes']
    ]);
    us.getRange(6, 1).setValue(
      'Roles: admin = everything · staff = all except adding/editing items · common = scan in/out + log used only, no values. ' +
      'Type an initial password in the Password column — it is replaced by a secure hash the first time the person signs in. Active: yes/no.')
      .setFontStyle('italic').setFontSize(9);
  }
  // 2. Requests tab: old schema (Tracker/Code, HandledAt) -> new (Size, DateResponded, Remarks)
  const sh = ss.getSheetByName('Requests');
  const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  const NEW = ['Timestamp','By','Item','Qty','Size','Status','HandledBy','DateResponded','Remarks'];
  if (head.join() !== NEW.join()) {
    const old = sh.getDataRange().getValues();
    const idx = {}; head.forEach((h, i) => idx[h] = i);
    const rows = old.slice(1).filter(r => r.join('') !== '').map(r => [
      r[idx['Timestamp']] || '', r[idx['By']] || '',
      r[idx['Item']] || '', r[idx['Qty']] || '',
      '',                                            // Size (new)
      r[idx['Status']] === 'Issued' ? 'Received' : (r[idx['Status']] || 'Requested'),
      r[idx['HandledBy']] || '',
      r[idx['HandledAt']] || '', ''                  // DateResponded, Remarks
    ]);
    sh.clear();
    sh.getRange(1, 1, 1, NEW.length).setValues([NEW])
      .setFontWeight('bold').setBackground('#2B6168').setFontColor('#FFFFFF');
    if (rows.length) sh.getRange(2, 1, rows.length, NEW.length).setValues(rows);
    sh.setFrozenRows(1);
  }
  return 'Migration complete';
}

function ackAlert_(q) {
  const sh = ss_().getSheetByName('Alerts');
  sh.getRange(q.row, 7).setValue('Acknowledged');
  sh.getRange(q.row, 8).setValue(q.by || '');
  sh.getRange(q.row, 9).setValue(new Date());
  return { ok: true };
}

function ackFromEmail_(rowStr) {
  const row = Number(rowStr);
  if (row >= 2) {
    const sh = ss_().getSheetByName('Alerts');
    if (sh.getRange(row, 7).getValue() === 'Sent') {
      sh.getRange(row, 7).setValue('Acknowledged');
      sh.getRange(row, 8).setValue('Via email link');
      sh.getRange(row, 9).setValue(new Date());
    }
  }
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:40px;text-align:center">' +
    '<h2>✅ Alert acknowledged</h2><p>Thank you — no escalation will be sent.</p></div>');
}

function stocktake_(q) {
  // q: {checkType:'Month-end'|'Annual', by, entries:[{tracker,code,item,expected,counted,store}], apply:true|false}
  const sh = ss_().getSheetByName('Stocktakes');
  const ts = new Date();
  (q.entries || []).forEach(en => {
    const variance = Number(en.counted) - Number(en.expected);
    sh.appendRow([ts, q.checkType || 'Month-end', en.store || '', en.tracker,
                  en.code, en.item, en.expected, en.counted, variance, q.by || 'Unknown']);
    if (q.apply && variance !== 0 && en.tracker !== 'Instruments') {
      try {
        const it = en.row ? itemAtRow_(en.tracker, Number(en.row)) : findItem_(en.tracker, en.code);
        it.sh.getRange(it.row, col_(it.head, 'Qty')).setValue(Number(en.counted));
      } catch (err) { /* item deleted mid-count — recorded but not applied */ }
    }
  });
  return { ok: true, saved: (q.entries || []).length };
}

/* ============================ users & auth =============================== */

const USER_HEADERS = ['Username','Password','Role','DisplayName','Active'];

function hash_(username, password) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(username).toLowerCase() + ':' + String(password), Utilities.Charset.UTF_8);
  return raw.map(b => ((b & 0xFF) + 0x100).toString(16).slice(1)).join('');
}

function usersSheet_() {
  const sh = ss_().getSheetByName('Users');
  if (!sh) throw 'Users tab missing — run migrate()';
  return sh;
}

function login_(q) {
  const u = String(q.username || '').trim().toLowerCase();
  const p = String(q.password || '');
  if (!u || !p) return { ok: false, error: 'Enter username and password' };
  const sh = usersSheet_();
  const vals = sh.getDataRange().getValues();
  for (let r = 1; r < vals.length; r++) {
    const [un, stored, role, name, active] = vals[r];
    if (String(un).trim().toLowerCase() !== u) continue;
    if (String(active).toLowerCase() === 'no') return { ok: false, error: 'Account disabled' };
    const s = String(stored);
    const matchHash = /^[0-9a-f]{64}$/.test(s) && s === hash_(u, p);
    const matchPlain = !/^[0-9a-f]{64}$/.test(s) && s === p;
    if (matchHash || matchPlain) {
      if (matchPlain) sh.getRange(r + 1, 2).setValue(hash_(u, p));  // upgrade to hash
      const token = Utilities.getUuid();
      saveToken_(token, { u: u, role: String(role || 'common').toLowerCase(),
                          name: String(name || u), exp: Date.now() + 7 * 86400000 });
      return { ok: true, token: token, role: String(role || 'common').toLowerCase(),
               displayName: String(name || u), username: u };
    }
    return { ok: false, error: 'Wrong password' };
  }
  return { ok: false, error: 'Unknown username' };
}

function changePassword_(q) {
  const s = q._session;
  const sh = usersSheet_();
  const vals = sh.getDataRange().getValues();
  for (let r = 1; r < vals.length; r++) {
    if (String(vals[r][0]).trim().toLowerCase() !== s.u) continue;
    const stored = String(vals[r][1]);
    const okOld = (/^[0-9a-f]{64}$/.test(stored) && stored === hash_(s.u, q.oldPassword)) ||
                  (!/^[0-9a-f]{64}$/.test(stored) && stored === String(q.oldPassword));
    if (!okOld) return { ok: false, error: 'Current password is wrong' };
    if (String(q.newPassword || '').length < 6)
      return { ok: false, error: 'New password must be at least 6 characters' };
    sh.getRange(r + 1, 2).setValue(hash_(s.u, q.newPassword));
    return { ok: true };
  }
  return { ok: false, error: 'User not found' };
}

function saveToken_(token, data) {
  const props = PropertiesService.getScriptProperties();
  let all = {};
  try { all = JSON.parse(props.getProperty('tokens') || '{}'); } catch (e) {}
  const now = Date.now();
  Object.keys(all).forEach(t => { if (all[t].exp < now) delete all[t]; });
  all[token] = data;
  props.setProperty('tokens', JSON.stringify(all));
}

function auth_(token) {
  if (!token) return null;
  let all = {};
  try { all = JSON.parse(
    PropertiesService.getScriptProperties().getProperty('tokens') || '{}'); } catch (e) {}
  const s = all[token];
  if (!s || s.exp < Date.now()) return null;
  return s;
}

/* ======================= daily alert + escalation ========================= */

function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('dailyStockCheck').timeBased().everyDays(1).atHour(8).create();
  ScriptApp.newTrigger('escalationCheck').timeBased().everyHours(4).create();
}

function dailyStockCheck() {
  const email = setting_('alert_email', '');
  if (!email) return;                              // not configured yet
  const expDays = Number(setting_('expiry_days', 90));
  const appName = setting_('app_name', 'Bollin Clinic Inventory');
  const alertsSh = ss_().getSheetByName('Alerts');
  const existing = readTab_('Alerts')
    .filter(a => a.Status !== 'Acknowledged')
    .map(a => a.Type + '|' + a.Tracker + '|' + a.Code);

  const low = [], out = [], exp = [];
  const today = new Date(); today.setHours(0, 0, 0, 0);

  ITEM_TABS.filter(t => t !== 'Instruments').forEach(t => {
    readTab_(t).forEach(it => {
      const qty = Number(it.Qty) || 0, ro = Number(it.ReorderLevel) || 0;
      if (qty === 0) out.push({ t: t, it: it });
      else if (ro > 0 && qty <= ro) low.push({ t: t, it: it });
      if (it.Expiry) {
        const d = Math.round((new Date(it.Expiry) - today) / 86400000);
        if (d <= expDays) exp.push({ t: t, it: it, d: d });
      }
    });
  });

  const fresh = [];
  function record(type, t, it, detail) {
    const key = type + '|' + t + '|' + (it.Code || it.Name);
    if (existing.indexOf(key) >= 0) return;        // already alerted, not yet acked
    alertsSh.appendRow([new Date(), type, t, it.Code || '', it.Name, detail, 'Sent', '', '']);
    fresh.push({ row: alertsSh.getLastRow(), type: type, name: it.Name, detail: detail });
  }
  out.forEach(x => record('Out of stock', x.t, x.it, 'Qty 0, reorder level ' + x.it.ReorderLevel));
  low.forEach(x => record('Low stock', x.t, x.it, 'Qty ' + x.it.Qty + ', reorder level ' + x.it.ReorderLevel));
  exp.forEach(x => record('Expiry', x.t, x.it,
    x.d < 0 ? 'EXPIRED ' + (-x.d) + ' days ago' : 'Expires in ' + x.d + ' days'));

  if (!fresh.length) return;
  const url = ScriptApp.getService().getUrl();
  let body = '<h3>' + appName + ' — daily stock alerts</h3><ul>';
  fresh.forEach(f => {
    body += '<li><b>' + f.type + ':</b> ' + f.name + ' — ' + f.detail +
            ' &nbsp;<a href="' + url + '?ack=' + f.row + '">Acknowledge</a></li>';
  });
  body += '</ul><p>Unacknowledged alerts escalate after ' +
          setting_('ack_hours', 48) + ' hours.</p>';
  MailApp.sendEmail({ to: email, subject: appName + ': ' + fresh.length + ' stock alert(s)',
                      htmlBody: body });
}

function escalationCheck() {
  const escEmail = setting_('escalation_email', '');
  if (!escEmail) return;
  const hours = Number(setting_('ack_hours', 48));
  const appName = setting_('app_name', 'Bollin Clinic Inventory');
  const sh = ss_().getSheetByName('Alerts');
  const rows = readTab_('Alerts');
  const overdue = rows.filter(a =>
    a.Status === 'Sent' &&
    (new Date() - new Date(a.Timestamp)) / 3600000 >= hours);
  if (!overdue.length) return;

  let body = '<h3>' + appName + ' — ESCALATION</h3>' +
    '<p>The following alerts were not acknowledged within ' + hours + ' hours:</p><ul>';
  overdue.forEach(a => {
    body += '<li><b>' + a.Type + ':</b> ' + a.Item + ' — ' + a.Detail +
            ' (raised ' + Utilities.formatDate(new Date(a.Timestamp),
              Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + ')</li>';
    sh.getRange(a._row, 7).setValue('Escalated');
  });
  body += '</ul>';
  MailApp.sendEmail({ to: escEmail, subject: appName + ': ESCALATION — ' +
                      overdue.length + ' unacknowledged alert(s)', htmlBody: body });
}
