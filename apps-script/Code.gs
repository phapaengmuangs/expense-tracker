/**
 * PayNote API for its own Google Sheet.
 *
 * Setup (once):
 *   1. In the sheet: Extensions > Apps Script, paste this file as Code.gs, save.
 *   2. Run setupToken() once, allow access, copy the TOKEN from the execution log.
 *      (It also creates any missing tab and header, so a blank sheet works too.)
 *   3. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 *   4. Put the web app URL and the TOKEN into the app (Settings tab).
 *
 * The app sends one POST per sync:
 *   {token, settings:{key:value}, wallets:[...], categories:[...], goals:[...],
 *    upserts:[transaction...], deletes:[id...]}
 * and gets back {ok, saved:{wallets,categories,goals,settings}, upserted, deleted, rejected, config}.
 * Setup changes are written before transactions, so an entry that uses a wallet,
 * category or goal created offline is accepted in the same sync.
 * Every write is idempotent by id, so a repeated sync never duplicates rows.
 *
 * Amounts: an expense or a payment from a goal stores the total paid in `amount`,
 * with the fee part of it in `fee` for reference. A transfer moves `amount`, and its
 * `fee` is paid on top by the source wallet and counted under the category chosen in
 * Settings (transfer_fee_category).
 */

const SHEETS = {
  Transactions: ['id', 'วันเวลา', 'ประเภท', 'category_id', 'ยอด (¥)', 'ค่าธรรมเนียม (¥)', 'จากกระเป๋า', 'ไปกระเป๋า', 'goal_id', 'หมายเหตุ', 'ที่มา', 'บันทึกเมื่อ'],
  Wallets: ['wallet_id', 'ชื่อ', 'ยอดเปิด (¥)', 'ยอด ณ วันที่', 'ลำดับ', 'สถานะ'],
  Categories: ['category_id', 'หมวดหลัก (parent_id)', 'ชื่อ', 'ลำดับ', 'สถานะ'],
  Goals: ['goal_id', 'ชื่อ', 'เป้าหมาย (¥)', 'วันที่เป้าหมาย', 'สถานะ', 'ลำดับ', 'สร้างเมื่อ'],
  Settings: ['key', 'value', 'คำอธิบาย']
};
const TX_COLS = 12;            // A..L are data; M.. are formula columns, never written here
const TYPES = ['expense', 'income', 'transfer', 'goal_pay', 'goal_in', 'goal_out', 'adjust'];
const SETTING_KEYS = {
  cycle_start_day: 'วันเริ่มรอบ 1–28 (1 = นับตามเดือนปฏิทิน) · รอบตั้งชื่อตามเดือนที่รอบสิ้นสุด',
  goal_wallet: 'กระเป๋าที่เก็บเงินของกระปุก (wallet_id)',
  transfer_fee_category: 'หมวดที่นับค่าธรรมเนียมตอนย้ายเงิน (category_id)'
};
const MAX_AMOUNT = 10000000;
const ID_RE = /^[A-Za-z0-9_-]{2,64}$/;

function setupToken() {
  ensureSheets_(SpreadsheetApp.getActive());
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  PropertiesService.getScriptProperties().setProperty('TOKEN', token);
  Logger.log('TOKEN: ' + token);
}

function ensureSheets_(ss) {
  Object.keys(SHEETS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]);
      sh.setFrozenRows(1);
    }
  });
  const st = ss.getSheetByName('Settings');
  const have = st.getLastRow() > 1 ? st.getRange(2, 1, st.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
  Object.keys(SETTING_KEYS).forEach(function (k) {
    if (have.indexOf(k) < 0) st.appendRow([k, k === 'cycle_start_day' ? 1 : '', SETTING_KEYS[k]]);
  });
}

function doGet() {
  return json_({ ok: true, service: 'paynote', hint: 'use POST' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_json' });
  }
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!token || typeof body.token !== 'string' || body.token !== token) {
    return json_({ ok: false, error: 'unauthorized' });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return json_({ ok: false, error: 'busy' });
  try {
    const ss = SpreadsheetApp.getActive();
    const tz = ss.getSpreadsheetTimeZone();
    const res = { ok: true, saved: { wallets: [], categories: [], goals: [], settings: [] }, upserted: [], deleted: [], rejected: [] };
    const list = function (x) { return Array.isArray(x) ? x.slice(0, 200) : []; };

    // 1. setup: settings, wallets, categories, goals
    if (body.settings && typeof body.settings === 'object') writeSettings_(ss, body.settings, res);
    const used = usage_(readTx_(ss).values);
    if (list(body.wallets).length) writeEntities_(ss, 'Wallets', list(body.wallets), validateWallet_, rowWallet_, used.wallets, 'hidden', res, 'wallets', tz);
    if (list(body.categories).length) writeEntities_(ss, 'Categories', list(body.categories), validateCategory_, rowCategory_, used.categories, 'hidden', res, 'categories', tz);
    if (list(body.goals).length) writeEntities_(ss, 'Goals', list(body.goals), validateGoal_, rowGoal_, used.goals, 'closed', res, 'goals', tz);

    // 2. transactions
    const ref = loadRefs_(ss);
    const upserts = list(body.upserts), deletes = list(body.deletes);
    if (upserts.length) {
      const tx = readTx_(ss);
      const append = [], appendAt = {};
      upserts.forEach(function (t) {
        const err = validateTx_(t, ref);
        if (err) { res.rejected.push({ id: String(t && t.id), error: err }); return; }
        const row = toRow_(t, ref.tz);
        const at = tx.rowById[t.id];
        if (at) {
          row[11] = tx.values[at - 2][11] || new Date();      // keep original created time
          tx.sheet.getRange(at, 1, 1, TX_COLS).setValues([row]);
        } else if (appendAt[t.id] !== undefined) {
          append[appendAt[t.id]] = row;                         // same new id twice in one batch: last wins
        } else {
          appendAt[t.id] = append.length;
          append.push(row);
        }
        if (res.upserted.indexOf(t.id) < 0) res.upserted.push(t.id);
      });
      if (append.length) {
        const start = tx.lastRow + 1;
        const need = start + append.length - 1 - tx.sheet.getMaxRows();
        if (need > 0) tx.sheet.insertRowsAfter(tx.sheet.getMaxRows(), need + 200);
        tx.sheet.getRange(start, 1, append.length, TX_COLS).setValues(append);
      }
    }
    if (deletes.length) {
      const tx = readTx_(ss);
      const rows = [];
      deletes.forEach(function (id) {
        id = String(id);
        if (tx.rowById[id]) rows.push(tx.rowById[id]);
        res.deleted.push(id);                                   // already absent counts as deleted
      });
      rows.sort(function (a, b) { return b - a; }).forEach(function (r) { tx.sheet.deleteRow(r); });
    }

    SpreadsheetApp.flush();
    res.config = buildConfig_(ss, loadRefs_(ss));
    return json_(res);
  } catch (err) {
    return json_({ ok: false, error: 'server', detail: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- reading ---------- */

function table_(ss, name) {
  const sh = ss.getSheetByName(name);
  const n = SHEETS[name].length;
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, n).getValues().filter(function (r) { return r[0] !== '' && r[0] !== null; });
}

function readSettings_(ss) {
  const s = {};
  const sh = ss.getSheetByName('Settings');
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) { if (r[0]) s[String(r[0])] = r[1]; });
  return s;
}

function loadRefs_(ss) {
  const s = readSettings_(ss);
  const sd = Math.min(28, Math.max(1, Math.floor(Number(s.cycle_start_day)) || 1));
  return {
    tz: ss.getSpreadsheetTimeZone(), startDay: sd,
    goalWallet: String(s.goal_wallet || ''), feeCategory: String(s.transfer_fee_category || ''),
    wallets: table_(ss, 'Wallets'), categories: table_(ss, 'Categories'), goals: table_(ss, 'Goals')
  };
}

function readTx_(ss) {
  const sheet = ss.getSheetByName('Transactions');
  const max = Math.max(sheet.getLastRow(), 1);
  const values = max > 1 ? sheet.getRange(2, 1, max - 1, TX_COLS).getValues() : [];
  const rowById = {};
  let lastRow = 1;
  values.forEach(function (r, i) {
    if (r[0] !== '' && r[0] !== null) { rowById[String(r[0])] = i + 2; lastRow = i + 2; }
  });
  return { sheet: sheet, values: values, rowById: rowById, lastRow: lastRow };
}

/* Which wallets, categories and goals have entries: those are hidden instead of deleted */
function usage_(values) {
  const u = { wallets: {}, categories: {}, goals: {} };
  values.forEach(function (r) {
    if (r[0] === '' || r[0] === null) return;
    if (r[3]) u.categories[r[3]] = true;
    if (r[6]) u.wallets[r[6]] = true;
    if (r[7]) u.wallets[r[7]] = true;
    if (r[8]) u.goals[r[8]] = true;
  });
  return u;
}

/* ---------- writing setup ---------- */

function writeSettings_(ss, settings, res) {
  const sh = ss.getSheetByName('Settings');
  const last = sh.getLastRow();
  const keys = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  Object.keys(settings).forEach(function (k) {
    if (!SETTING_KEYS[k]) return;
    let v = settings[k];
    if (k === 'cycle_start_day') {
      v = Number(v);
      if (!Number.isInteger(v) || v < 1 || v > 28) { res.rejected.push({ id: k, error: 'bad_setting', kind: 'settings' }); return; }
    } else {
      v = String(v || '');
      if (v && !ID_RE.test(v)) { res.rejected.push({ id: k, error: 'bad_setting', kind: 'settings' }); return; }
    }
    const i = keys.indexOf(k);
    if (i >= 0) sh.getRange(i + 2, 2).setValue(v);
    else { sh.appendRow([k, v, SETTING_KEYS[k]]); keys.push(k); }
    res.saved.settings.push(k);
  });
}

/* Upsert rows by id. {id, delete:true} removes a row nobody uses, or hides one in use. */
function writeEntities_(ss, name, items, validate, toRow, used, hideStatus, res, kind, tz) {
  const sh = ss.getSheetByName(name);
  const n = SHEETS[name].length;
  const read = function () {
    const last = sh.getLastRow();
    const v = last >= 2 ? sh.getRange(2, 1, last - 1, n).getValues() : [];
    const at = {};
    v.forEach(function (r, i) { if (r[0] !== '' && r[0] !== null) at[String(r[0])] = i + 2; });
    return { v: v, at: at, last: Math.max(last, 1) };
  };
  let t = read();
  const toDelete = [];
  items.forEach(function (it) {
    if (!it || typeof it.id !== 'string' || !ID_RE.test(it.id)) { res.rejected.push({ id: String(it && it.id), error: 'bad_id', kind: kind }); return; }
    const row = t.at[it.id];
    if (it.delete) {
      if (row && used[it.id]) {
        const r = t.v[row - 2].slice(); r[n === 7 ? 4 : n - 1] = hideStatus;     // status column
        sh.getRange(row, 1, 1, n).setValues([r]);
      } else if (row) toDelete.push(row);
      res.saved[kind].push(it.id);
      return;
    }
    const err = validate(it);
    if (err) { res.rejected.push({ id: it.id, error: err, kind: kind }); return; }
    const values = toRow(it, row ? t.v[row - 2] : null, tz);
    if (row) sh.getRange(row, 1, 1, n).setValues([values]);
    else {
      t.last += 1;
      sh.getRange(t.last, 1, 1, n).setValues([values]);
      t.at[it.id] = t.last; t.v[t.last - 2] = values;
    }
    res.saved[kind].push(it.id);
  });
  toDelete.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
}

const isInt_ = function (x, min) { return Number.isInteger(x) && x >= min && x <= MAX_AMOUNT; };
const okName_ = function (s) { return typeof s === 'string' && s.trim() !== '' && s.length <= 60; };
const okDate_ = function (s) { return !s || (typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s)); };
const day_ = function (s, tz) { return s ? Utilities.parseDate(s + ' 00:00', tz, 'yyyy-MM-dd HH:mm') : ''; };

function validateWallet_(w) {
  if (!okName_(w.name)) return 'bad_name';
  if (!Number.isInteger(w.opening) || Math.abs(w.opening) > MAX_AMOUNT) return 'bad_amount';
  if (!w.openDate || !okDate_(w.openDate)) return 'bad_date';
  if (!Number.isInteger(w.order)) return 'bad_order';
  if (['active', 'hidden'].indexOf(w.status) < 0) return 'bad_status';
  return '';
}
function rowWallet_(w, old, tz) { return [w.id, safe_(w.name), w.opening, day_(w.openDate, tz), w.order, w.status]; }

function validateCategory_(c) {
  if (!okName_(c.name)) return 'bad_name';
  if (c.parent && !ID_RE.test(c.parent)) return 'bad_parent';
  if (c.parent === c.id) return 'bad_parent';
  if (!Number.isInteger(c.order)) return 'bad_order';
  if (['active', 'hidden'].indexOf(c.status) < 0) return 'bad_status';
  return '';
}
function rowCategory_(c) { return [c.id, c.parent || '', safe_(c.name), c.order, c.status]; }

function validateGoal_(g) {
  if (!okName_(g.name)) return 'bad_name';
  if (!isInt_(g.target, 0)) return 'bad_target';
  if (!okDate_(g.due)) return 'bad_date';
  if (!Number.isInteger(g.order)) return 'bad_order';
  if (['active', 'closed'].indexOf(g.status) < 0) return 'bad_status';
  return '';
}
function rowGoal_(g, old, tz) { return [g.id, safe_(g.name), g.target, day_(g.due, tz), g.status, g.order, old ? old[6] : new Date()]; }

/* ---------- transactions ---------- */

function has_(list, id) {
  return list.some(function (r) { return String(r[0]) === String(id); });
}

function parseTs_(s, tz) {
  if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(s)) return null;
  const d = Utilities.parseDate(s.replace('T', ' '), tz, 'yyyy-MM-dd HH:mm');
  return isNaN(d.getTime()) ? null : d;
}

function validateTx_(t, ref) {
  if (!t || typeof t !== 'object') return 'not_object';
  if (typeof t.id !== 'string' || !ID_RE.test(t.id) || t.id.length < 4) return 'bad_id';
  if (TYPES.indexOf(t.type) < 0) return 'bad_type';
  if (!isInt_(t.amount, 1)) return 'bad_amount';
  const fee = t.fee || 0;
  if (!isInt_(fee, 0)) return 'bad_fee';
  if (!parseTs_(t.ts, ref.tz)) return 'bad_ts';
  const w = function (x) { return x && has_(ref.wallets, x); };
  const g = function (x) { return x && has_(ref.goals, x); };
  switch (t.type) {
    case 'expense':
      if (!t.category || !has_(ref.categories, t.category)) return 'bad_category';
      if (!w(t.from)) return 'bad_from';
      if (fee >= t.amount) return 'bad_fee';
      break;
    case 'income':
      if (!w(t.to)) return 'bad_to';
      if (fee) return 'bad_fee';
      break;
    case 'transfer':
      if (!w(t.from) || !w(t.to)) return 'bad_wallet';
      if (t.from === t.to) return 'same_wallet';
      if (fee && !has_(ref.categories, ref.feeCategory)) return 'no_fee_category';
      break;
    case 'goal_pay':
      if (!w(t.from)) return 'bad_from';
      if (!g(t.goal)) return 'bad_goal';
      if (fee >= t.amount) return 'bad_fee';
      break;
    case 'goal_in':
    case 'goal_out':
      if (!g(t.goal)) return 'bad_goal';
      if (fee) return 'bad_fee';
      break;
    case 'adjust':                    // balance correction after checking the real wallet: exactly one side
      if (!!t.from === !!t.to) return 'bad_adjust';
      if (t.from && !w(t.from)) return 'bad_from';
      if (t.to && !w(t.to)) return 'bad_to';
      if (fee) return 'bad_fee';
      break;
  }
  return '';
}

/* Text starting with = + - @ would be read as a formula; store it as plain text */
function safe_(s) {
  s = String(s == null ? '' : s).slice(0, 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function toRow_(t, tz) {
  const cat = t.type === 'expense' ? t.category : '';
  const goal = /^goal_/.test(t.type) ? t.goal : '';
  const from = ['expense', 'transfer', 'goal_pay', 'adjust'].indexOf(t.type) >= 0 ? t.from || '' : '';
  const to = ['income', 'transfer', 'adjust'].indexOf(t.type) >= 0 ? t.to || '' : '';
  return [t.id, parseTs_(t.ts, tz), t.type, safe_(cat), t.amount, t.fee || 0,
          safe_(from), safe_(to), safe_(goal), safe_(t.note), 'app', new Date()];
}

/* ---------- config sent back to the app ---------- */

/* A cycle is named after the month it ends in. With start day 1 it is the calendar month;
   with start day 25 the cycle 25 Nov – 24 Dec is 202612. */
function cycleOf_(date, tz, sd) {
  const p = Utilities.formatDate(date, tz, 'yyyy-M-d').split('-').map(Number);
  let y = p[0], m = p[1] + (sd > 1 && p[2] >= sd ? 1 : 0);
  if (m > 12) { m -= 12; y += 1; }
  return y * 100 + m;
}
function prevKey_(k) { return k % 100 === 1 ? k - 89 : k - 1; }
function range_(k, sd) {
  const y = Math.floor(k / 100), m = k % 100;
  const stz = Session.getScriptTimeZone();
  const start = sd > 1 ? new Date(y, m - 2, sd) : new Date(y, m - 1, 1);
  const end = sd > 1 ? new Date(y, m - 1, sd - 1) : new Date(y, m, 0);
  return { key: k, start: Utilities.formatDate(start, stz, 'yyyy-MM-dd'), end: Utilities.formatDate(end, stz, 'yyyy-MM-dd') };
}

function buildConfig_(ss, ref) {
  const tz = ref.tz, sd = ref.startDay;
  const key = cycleOf_(new Date(), tz, sd), pk = prevKey_(key);
  const fmt = function (d) { return d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : ''; };
  const txv = readTx_(ss).values.filter(function (r) { return r[0] !== '' && r[0] !== null && r[1] instanceof Date; });
  const used = usage_(txv);

  const spend = {}, totals = { income: [0, 0], goalPay: [0, 0] };
  const add = function (cat, c, v) {
    if (!cat || !v) return;
    if (!spend[cat]) spend[cat] = [0, 0];
    if (c === key) spend[cat][0] += v; else if (c === pk) spend[cat][1] += v;
  };
  txv.forEach(function (r) {
    const c = cycleOf_(r[1], tz, sd), amt = Number(r[4]) || 0, fee = Number(r[5]) || 0;
    const slot = c === key ? 0 : c === pk ? 1 : -1;
    if (r[2] === 'expense') add(r[3], c, amt);
    if (r[2] === 'transfer' && fee) add(ref.feeCategory, c, fee);
    if (slot >= 0 && r[2] === 'income') totals.income[slot] += amt;
    if (slot >= 0 && r[2] === 'goal_pay') totals.goalPay[slot] += amt;
  });

  const wallets = ref.wallets.map(function (r) {
    const id = String(r[0]), open = r[3] instanceof Date ? r[3] : new Date(0);
    let bal = Number(r[2]) || 0;
    txv.forEach(function (t) {
      if (t[1] < open) return;
      const amt = Number(t[4]) || 0, fee = Number(t[5]) || 0;
      if (t[7] === id) bal += amt;
      if (t[6] === id) bal -= amt + (t[2] === 'transfer' ? fee : 0);
    });
    return { id: id, name: String(r[1]), opening: Number(r[2]) || 0, openDate: fmt(r[3]), order: Number(r[4]) || 0,
             status: r[5] === 'hidden' ? 'hidden' : 'active', balance: bal, used: !!used.wallets[id] };
  });

  const categories = ref.categories.map(function (r) {
    const id = String(r[0]), s = spend[id] || [0, 0];
    return { id: id, parent: String(r[1] || ''), name: String(r[2]), order: Number(r[3]) || 0,
             status: r[4] === 'hidden' ? 'hidden' : 'active', cur: s[0], prev: s[1], used: !!used.categories[id] };
  });

  const goals = ref.goals.map(function (r) {
    const id = String(r[0]);
    let bal = 0;
    txv.forEach(function (t) {
      if (t[8] !== id) return;
      bal += t[2] === 'goal_in' ? Number(t[4]) : -Number(t[4]);
    });
    return { id: id, name: String(r[1]), target: Number(r[2]) || 0, due: fmt(r[3]),
             status: r[4] === 'closed' ? 'closed' : 'active', order: Number(r[5]) || 0, balance: bal, used: !!used.goals[id] };
  });

  return {
    settings: { cycle_start_day: sd, goal_wallet: ref.goalWallet, transfer_fee_category: ref.feeCategory },
    cycle: range_(key, sd), prevCycle: range_(pk, sd),
    wallets: wallets, categories: categories, goals: goals,
    income: { cur: totals.income[0], prev: totals.income[1] }, goalPay: { cur: totals.goalPay[0], prev: totals.goalPay[1] },
    serverTime: Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm")
  };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
