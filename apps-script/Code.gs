/**
 * PayNote API for the sheet "การเงินนาโกย่า - ระบบบันทึก".
 *
 * Setup (once):
 *   1. In the sheet: Extensions > Apps Script, paste this file as Code.gs, save.
 *   2. Run setupToken() once, allow access, copy the TOKEN from the execution log.
 *   3. Deploy > New deployment > Web app: Execute as "Me", Who has access "Anyone".
 *   4. Put the web app URL and the TOKEN into the app (Settings tab).
 *
 * The app sends one POST per sync:
 *   {token, funds:[goal...], upserts:[transaction...], deletes:[id...]}
 * and gets back {ok, funds, upserted, deleted, rejected, config}.
 * Goals (savings pockets) are written to the Funds tab before transactions, so a deposit
 * into a goal created offline is accepted in the same sync.
 * Writes are idempotent by id, so a repeated sync never duplicates rows.
 */

const TX_SHEET = 'Transactions';
const TX_COLS = 13;            // A..M are data; N is the cycle ARRAYFORMULA, never written here
const TYPES = ['expense', 'transfer', 'income', 'one_time', 'fund_move', 'adjust'];
const MAX_AMOUNT = 10000000;
const FUND_COLS = 8;           // id, name, target, opening, opening date, due date, note, status
const FUND_STATUS = ['active', 'closed'];

function setupToken() {
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
  PropertiesService.getScriptProperties().setProperty('TOKEN', token);
  Logger.log('TOKEN: ' + token);
}

function doGet() {
  return json_({ ok: true, service: 'expense-tracker', hint: 'use POST' });
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
    const res = { ok: true, funds: [], upserted: [], deleted: [], rejected: [] };
    const goals = Array.isArray(body.funds) ? body.funds.slice(0, 50) : [];
    if (goals.length) writeFunds_(ss, goals, res);
    const ref = loadRefs_(ss);

    const upserts = Array.isArray(body.upserts) ? body.upserts.slice(0, 500) : [];
    const deletes = Array.isArray(body.deletes) ? body.deletes.slice(0, 500) : [];

    if (upserts.length) {
      const tx = readTx_(ss);
      const append = [], appendAt = {};
      upserts.forEach(function (t) {
        const err = validate_(t, ref);
        if (err) { res.rejected.push({ id: String(t && t.id), error: err }); return; }
        const row = toRow_(t, ref.tz);
        const at = tx.rowById[t.id];
        if (at) {
          row[12] = tx.values[at - 2][12] || new Date();      // keep original created time
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
  const v = sh.getDataRange().getValues();
  v.shift();                                                    // header row
  return v.filter(function (r) { return r[0] !== '' && r[0] !== null; });
}

function loadRefs_(ss) {
  const settings = {};
  ss.getSheetByName('Settings').getRange('A2:B').getValues().forEach(function (r) {
    if (r[0]) settings[r[0]] = r[1];
  });
  return {
    tz: ss.getSpreadsheetTimeZone(),
    startDay: Number(settings.cycle_start_day) || 1,
    income: Number(settings.monthly_income) || 0,
    lines: table_(ss, 'Categories'),
    accounts: table_(ss, 'Accounts'),
    funds: table_(ss, 'Funds')
  };
}

/* ---------- goals (Funds tab) ---------- */

function writeFunds_(ss, goals, res) {
  const sh = ss.getSheetByName('Funds');
  const tz = ss.getSpreadsheetTimeZone();
  const max = Math.max(sh.getLastRow(), 1);
  const values = max > 1 ? sh.getRange(2, 1, max - 1, FUND_COLS).getValues() : [];
  const rowById = {};
  let lastRow = 1;
  values.forEach(function (r, i) {
    if (r[0] !== '' && r[0] !== null) { rowById[String(r[0])] = i + 2; lastRow = i + 2; }
  });
  goals.forEach(function (g) {
    const err = validateFund_(g, tz);
    if (err) { res.rejected.push({ id: String(g && g.id), error: err, kind: 'fund' }); return; }
    const due = g.due ? Utilities.parseDate(g.due + ' 00:00', tz, 'yyyy-MM-dd HH:mm') : '';
    const at = rowById[g.id];
    if (at) {
      const old = values[at - 2];       // opening balance, its date and the note stay as they are
      sh.getRange(at, 1, 1, FUND_COLS).setValues([[g.id, safe_(g.name), g.target, old[3], old[4], due, old[6], g.status]]);
    } else {
      lastRow += 1;
      sh.getRange(lastRow, 1, 1, FUND_COLS).setValues([[g.id, safe_(g.name), g.target, 0, new Date(), due, 'สร้างจากแอป', g.status]]);
      rowById[g.id] = lastRow;
      values[lastRow - 2] = [g.id, g.name, g.target, 0, new Date(), due, 'สร้างจากแอป', g.status];
    }
    res.funds.push(g.id);
  });
}

function validateFund_(g, tz) {
  if (!g || typeof g !== 'object') return 'not_object';
  if (typeof g.id !== 'string' || !/^[A-Za-z0-9_-]{2,40}$/.test(g.id)) return 'bad_id';
  if (typeof g.name !== 'string' || !g.name.trim() || g.name.length > 60) return 'bad_name';
  if (!Number.isInteger(g.target) || g.target < 0 || g.target > MAX_AMOUNT) return 'bad_target';
  if (g.due && !(typeof g.due === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(g.due))) return 'bad_due';
  if (FUND_STATUS.indexOf(g.status) < 0) return 'bad_status';
  return '';
}

function readTx_(ss) {
  const sheet = ss.getSheetByName(TX_SHEET);
  const max = Math.max(sheet.getLastRow(), 1);
  const values = max > 1 ? sheet.getRange(2, 1, max - 1, TX_COLS).getValues() : [];
  const rowById = {};
  let lastRow = 1;
  values.forEach(function (r, i) {
    if (r[0] !== '' && r[0] !== null) { rowById[String(r[0])] = i + 2; lastRow = i + 2; }
  });
  return { sheet: sheet, values: values, rowById: rowById, lastRow: lastRow };
}

/* ---------- validation and row building ---------- */

function has_(list, id) {
  return list.some(function (r) { return String(r[0]) === String(id); });
}

function parseTs_(s, tz) {
  if (typeof s !== 'string' || !/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/.test(s)) return null;
  const d = Utilities.parseDate(s.replace('T', ' '), tz, 'yyyy-MM-dd HH:mm');
  return isNaN(d.getTime()) ? null : d;
}

function validate_(t, ref) {
  if (!t || typeof t !== 'object') return 'not_object';
  if (typeof t.id !== 'string' || !/^[A-Za-z0-9_-]{4,64}$/.test(t.id)) return 'bad_id';
  if (TYPES.indexOf(t.type) < 0) return 'bad_type';
  if (!Number.isInteger(t.amount) || t.amount <= 0 || t.amount > MAX_AMOUNT) return 'bad_amount';
  if (!parseTs_(t.ts, ref.tz)) return 'bad_ts';
  const acc = function (a) { return a && has_(ref.accounts, a); };
  const fund = function (f) { return f && has_(ref.funds, f); };
  switch (t.type) {
    case 'expense':
      if (!t.line_id || !has_(ref.lines, t.line_id)) return 'bad_line';
      if (!acc(t.from_account)) return 'bad_from_account';
      break;
    case 'transfer':
      if (!acc(t.from_account) || !acc(t.to_account)) return 'bad_account';
      if (t.from_account === t.to_account) return 'same_account';
      break;
    case 'income':
      if (!acc(t.to_account)) return 'bad_to_account';
      if (t.to_fund && !fund(t.to_fund)) return 'bad_fund';
      break;
    case 'one_time':
      if (!acc(t.from_account)) return 'bad_from_account';
      if (!fund(t.from_fund)) return 'bad_fund';
      break;
    case 'fund_move':
      if (!fund(t.from_fund) && !fund(t.to_fund)) return 'bad_fund';
      break;
    case 'adjust':                    // balance correction after checking the real wallet: exactly one side
      if (!!t.from_account === !!t.to_account) return 'bad_adjust';
      if (t.from_account && !acc(t.from_account)) return 'bad_from_account';
      if (t.to_account && !acc(t.to_account)) return 'bad_to_account';
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
  return [
    t.id, parseTs_(t.ts, tz), t.type, safe_(t.line_id), t.amount,
    safe_(t.from_account), safe_(t.to_account), safe_(t.from_fund), safe_(t.to_fund),
    safe_(t.event_id), safe_(t.note), 'app', new Date()
  ];
}

/* ---------- config sent back to the app ---------- */

function cycleOf_(date, tz, startDay) {
  const p = Utilities.formatDate(date, tz, 'yyyy-M-d').split('-').map(Number);
  let y = p[0], m = p[1] + (p[2] >= startDay ? 1 : 0);
  if (m > 12) { m -= 12; y += 1; }
  return y * 100 + m;
}

function buildConfig_(ss, ref) {
  const tz = ref.tz, sd = ref.startDay;
  const key = cycleOf_(new Date(), tz, sd);
  const prevKey = key % 100 === 1 ? key - 89 : key - 1;      // 202701 -> 202612
  // built in the script's own time zone, so format them in that zone too
  const stz = Session.getScriptTimeZone();
  const range = function (k) {
    const y = Math.floor(k / 100), m = k % 100;
    return { key: k, start: Utilities.formatDate(new Date(y, m - 2, sd), stz, 'yyyy-MM-dd'),
             end: Utilities.formatDate(new Date(y, m - 1, sd - 1), stz, 'yyyy-MM-dd'), startDay: sd };
  };
  const fmt = function (d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); };

  const txv = readTx_(ss).values.filter(function (r) { return r[0] !== '' && r[0] !== null; });
  const spent = {}, prev = {};
  txv.forEach(function (r) {
    if (r[2] !== 'expense' || !(r[1] instanceof Date)) return;
    const c = cycleOf_(r[1], tz, sd);
    if (c === key) spent[r[3]] = (spent[r[3]] || 0) + Number(r[4]);
    else if (c === prevKey) prev[r[3]] = (prev[r[3]] || 0) + Number(r[4]);
  });

  const lines = ref.lines.map(function (r) {
    return { id: String(r[0]), category: String(r[1]), item: String(r[2]), account: String(r[3]),
             spent: spent[r[0]] || 0, prev: prev[r[0]] || 0 };
  });

  const accounts = ref.accounts.map(function (r) {
    const id = String(r[0]), open = r[3] instanceof Date ? r[3] : new Date(0);
    let bal = Number(r[2]) || 0;
    txv.forEach(function (t) {
      if (!(t[1] instanceof Date) || t[1] < open) return;
      if (t[6] === id) bal += Number(t[4]);
      if (t[5] === id) bal -= Number(t[4]);
    });
    return { id: id, name: String(r[1]), balance: bal };
  });

  const funds = ref.funds.map(function (r) {
    const id = String(r[0]);
    let bal = Number(r[3]) || 0;
    txv.forEach(function (t) {
      if (t[8] === id) bal += Number(t[4]);
      if (t[7] === id) bal -= Number(t[4]);
    });
    return { id: id, name: String(r[1]), target: Number(r[2]) || 0, balance: bal,
             due: r[5] instanceof Date ? fmt(r[5]) : '', status: FUND_STATUS.indexOf(r[7]) >= 0 ? r[7] : 'active' };
  });

  return {
    cycle: range(key), prevCycle: range(prevKey),
    income: ref.income, lines: lines, accounts: accounts, funds: funds,
    serverTime: Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm")
  };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
