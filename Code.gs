// ============================================================
// GAVAKSHI BADDIES — Attendance & Expense Tracker
// Google Apps Script Backend — V6
// Adds: count multiplier (sessions), expense paid_by, richer report
// ============================================================

const SHEET_PLAYERS  = "Players";
const SHEET_SESSIONS = "Sessions";
const SHEET_EXPENSES = "Expenses";
const SHEET_PLAY_LOG = "PlayLog";
const SHEET_META     = "Meta";

// ── Setup: create sheets if missing, patch missing columns ──
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet(ss, SHEET_PLAYERS,  ["id","name","status","added_date","deactivated_date"]);
  ensureSheet(ss, SHEET_SESSIONS, ["session_id","date","num_players","cost_type","cost_desc","amount","per_head","paid_by","count_mult"]);
  ensureSheet(ss, SHEET_EXPENSES, ["expense_id","date","amount","description","player_ids","expense_scope","paid_by"]);
  ensureSheet(ss, SHEET_PLAY_LOG, ["log_id","session_id","player_id","amount_owed","play_count"]);
  ensureSheet(ss, SHEET_META,     ["key","value"]);
  return ok("Setup complete");
}

// Ensures a sheet exists AND has at least the given headers.
// New columns are appended to existing sheets without touching data.
function ensureSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    return sh;
  }
  // Patch missing trailing columns
  const lastCol = sh.getLastColumn();
  const existing = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  headers.forEach((h, i) => {
    if (existing[i] !== h) {
      sh.getRange(1, i + 1).setValue(h);
    }
  });
  return sh;
}

// ── HTTP entry ──────────────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, payload } = body;
    let result;
    switch (action) {
      case "setup":           result = setup(); break;
      case "addPlayer":       result = addPlayer(payload); break;
      case "getPlayers":      result = getPlayers(); break;
      case "setPlayerStatus": result = setPlayerStatus(payload); break;
      case "saveSession":     result = saveSession(payload); break;
      case "getSession":      result = getSession(payload); break;
      case "saveExpense":     result = saveExpense(payload); break;
      case "getExpenses":     result = getExpenses(payload); break;
      case "getReport":       result = getReport(payload); break;
      default:                result = err("Unknown action: " + action);
    }
    return json(result);
  } catch(ex) {
    return json(err(ex.message));
  }
}

function doGet() { return json(ok("Gavakshi Baddies API V6 live")); }

// ── Players ─────────────────────────────────────────────────
function addPlayer({ name }) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_PLAYERS);
    const id = nextPlayerId(ss);
    const now = getISTDate();
    sh.appendRow([id, name.trim(), "active", now, ""]);
    SpreadsheetApp.flush();
    return ok({ id: String(id), name: name.trim(), status: "active", added_date: now, deactivated_date: "" });
  } finally {
    lock.releaseLock();
  }
}

function nextPlayerId(ss) {
  const meta = ss.getSheetByName(SHEET_META);
  const data = meta.getDataRange().getValues();
  let row = -1, current = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === "player_counter") { row = i + 1; current = parseInt(data[i][1]) || 0; break; }
  }
  const next = current + 1;
  if (row === -1) meta.appendRow(["player_counter", next]);
  else meta.getRange(row, 2).setValue(next);
  return next;
}

function getPlayers() {
  const rows = sheetRows(SHEET_PLAYERS);
  return ok(rows.map(r => ({
    id: r.id, name: r.name, status: r.status,
    added_date: r.added_date, deactivated_date: r.deactivated_date || ""
  })));
}

function setPlayerStatus({ id, status }) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_PLAYERS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sh.getRange(i+1, 3).setValue(status);
      sh.getRange(i+1, 5).setValue(status === "inactive" ? getISTDate() : "");
      SpreadsheetApp.flush();
      return ok({ id, status });
    }
  }
  return err("Player not found");
}

// ── Sessions ────────────────────────────────────────────────
// count_mult (1-4): multiplies the play COUNT per player, never the money.
function saveSession({ date, num_players, cost_type, cost_desc, amount, paid_by, players, count_mult }) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    const shS = ss.getSheetByName(SHEET_SESSIONS);
    const shL = ss.getSheetByName(SHEET_PLAY_LOG);
    const per = players.length > 0 ? (parseFloat(amount) / players.length) : 0;
    let mult = parseInt(count_mult) || 1;
    if (mult < 1) mult = 1;
    if (mult > 4) mult = 4;

    // Idempotent: remove any existing rows for this date first
    deleteSessionByDate(ss, date);

    const sid = "S" + Date.now();
    shS.appendRow([sid, date, num_players, cost_type, cost_desc, parseFloat(amount), round2(per), String(paid_by||""), mult]);
    players.forEach(pid => {
      shL.appendRow(["L" + Date.now() + Math.random().toString(36).slice(2), sid, String(pid), round2(per), mult]);
    });

    SpreadsheetApp.flush();
    return ok({ session_id: sid, date, per_head: round2(per), count_mult: mult, saved: true });
  } finally {
    lock.releaseLock();
  }
}

function deleteSessionByDate(ss, date) {
  const shS = ss.getSheetByName(SHEET_SESSIONS);
  const shL = ss.getSheetByName(SHEET_PLAY_LOG);
  const sData = shS.getDataRange().getValues();
  const sidsToDelete = [];
  for (let i = sData.length - 1; i >= 1; i--) {
    if (fmtCell(sData[i][1]) === date) {
      sidsToDelete.push(String(sData[i][0]));
      shS.deleteRow(i + 1);
    }
  }
  if (sidsToDelete.length > 0) {
    const lData = shL.getDataRange().getValues();
    for (let i = lData.length - 1; i >= 1; i--) {
      if (sidsToDelete.indexOf(String(lData[i][1])) !== -1) shL.deleteRow(i + 1);
    }
  }
}

function getSession({ date }) {
  const sessions = sheetRows(SHEET_SESSIONS);
  const session  = sessions.find(r => r.date === date);
  if (!session) return ok(null);
  const logs = sheetRows(SHEET_PLAY_LOG).filter(r => r.session_id === session.session_id);
  return ok({
    session_id: session.session_id,
    date: session.date,
    num_players: session.num_players,
    cost_type: session.cost_type,
    cost_desc: session.cost_desc,
    amount: session.amount,
    per_head: session.per_head,
    paid_by: session.paid_by || "",
    count_mult: parseInt(session.count_mult) || 1,
    players: logs.map(r => r.player_id)
  });
}

// ── Expenses ────────────────────────────────────────────────
function saveExpense({ date, amount, description, player_ids, expense_scope, paid_by }) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_EXPENSES);
    const eid = "E" + Date.now();
    sh.appendRow([eid, date, parseFloat(amount), description, (player_ids||[]).map(String).join(","), expense_scope||"player", String(paid_by||"")]);
    SpreadsheetApp.flush();
    return ok({ expense_id: eid, saved: true });
  } finally {
    lock.releaseLock();
  }
}

function getExpenses({ from, to }) {
  const rows = sheetRows(SHEET_EXPENSES).filter(r => r.date >= from && r.date <= to);
  return ok(rows.map(r => ({
    expense_id: r.expense_id, date: r.date, amount: r.amount,
    description: r.description,
    player_ids: r.player_ids ? r.player_ids.split(",").filter(Boolean) : [],
    expense_scope: r.expense_scope || "player",
    paid_by: r.paid_by || ""
  })));
}

// ── Reports ─────────────────────────────────────────────────
function getReport({ from, to }) {
  const players  = sheetRows(SHEET_PLAYERS);
  const sessions = sheetRows(SHEET_SESSIONS).filter(r => r.date >= from && r.date <= to);
  const logs     = sheetRows(SHEET_PLAY_LOG);
  const expenses = sheetRows(SHEET_EXPENSES).filter(r => r.date >= from && r.date <= to);

  const sidSet = {};
  sessions.forEach(s => { sidSet[s.session_id] = s; });

  const nameById = {};
  players.forEach(p => { nameById[p.id] = p.name; });

  // ── Section 1: player summary (count incorporates multiplier) ──
  const report = {};
  players.forEach(p => {
    report[p.id] = {
      id: p.id, name: p.name,
      weekday_sessions: 0, weekend_sessions: 0,
      weekday_cost: 0, weekend_cost: 0
    };
  });

  logs.forEach(l => {
    const s = sidSet[l.session_id];
    if (!s || !report[l.player_id]) return;
    const dow = new Date(s.date + "T00:00:00").getDay();
    const isWeekend = (dow === 0 || dow === 6);
    const amt = parseFloat(l.amount_owed) || 0;          // money: NOT multiplied
    const cnt = parseInt(l.play_count) || 1;             // count: multiplied
    if (isWeekend) { report[l.player_id].weekend_sessions += cnt; report[l.player_id].weekend_cost += amt; }
    else           { report[l.player_id].weekday_sessions += cnt; report[l.player_id].weekday_cost += amt; }
  });

  // Player-split incidental expenses fold into cost (not session count)
  expenses.filter(e => (e.expense_scope||"player") === "player").forEach(exp => {
    const pids = exp.player_ids ? exp.player_ids.split(",").filter(Boolean) : [];
    if (pids.length === 0) return;
    const share = parseFloat(exp.amount) / pids.length;
    const dow = new Date(exp.date + "T00:00:00").getDay();
    const isWeekend = (dow === 0 || dow === 6);
    pids.forEach(pid => {
      if (!report[pid]) return;
      if (isWeekend) report[pid].weekend_cost += share;
      else report[pid].weekday_cost += share;
    });
  });

  // ── Section 2: who paid for the court (weekday / weekend / total) ──
  const courtPaid = {};
  sessions.forEach(s => {
    const payer = String(s.paid_by || "");
    if (!payer) return;
    if (!courtPaid[payer]) courtPaid[payer] = { id: payer, name: nameById[payer] || ("#" + payer), weekday_paid: 0, weekend_paid: 0 };
    const dow = new Date(s.date + "T00:00:00").getDay();
    const isWeekend = (dow === 0 || dow === 6);
    const amt = parseFloat(s.amount) || 0;
    if (isWeekend) courtPaid[payer].weekend_paid += amt;
    else           courtPaid[payer].weekday_paid += amt;
  });
  const courtPaidArr = Object.values(courtPaid).map(c => ({
    id: c.id, name: c.name,
    weekday_paid: round2(c.weekday_paid),
    weekend_paid: round2(c.weekend_paid),
    total_paid:   round2(c.weekday_paid + c.weekend_paid)
  }));

  // ── Section 3 & 4: expense detail (date, paid_by, amount, desc, split) ──
  const expenseDetail = expenses.map(e => {
    const pids = e.player_ids ? e.player_ids.split(",").filter(Boolean) : [];
    return {
      date: e.date,
      amount: parseFloat(e.amount) || 0,
      description: e.description,
      paid_by: e.paid_by || "",
      paid_by_name: e.paid_by ? (nameById[e.paid_by] || ("#" + e.paid_by)) : "—",
      expense_scope: e.expense_scope || "player",
      split_ids: pids,
      split_names: pids.map(pid => nameById[pid] || ("#" + pid))
    };
  });

  // Kept for backward-compat with any older frontend
  const groupExpenses = expenses
    .filter(e => (e.expense_scope||"player") === "group")
    .map(e => ({ date: e.date, amount: parseFloat(e.amount), description: e.description }));

  return ok({
    players: Object.values(report),
    court_paid_by: courtPaidArr,
    expenses: expenseDetail,
    group_expenses: groupExpenses
  });
}

// ── Helpers ─────────────────────────────────────────────────
function getISTDate() {
  const ist = new Date(Date.now() + (5.5 * 3600 * 1000));
  return ist.toISOString().slice(0,10);
}

function fmtCell(v) {
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth()+1).padStart(2,"0");
    const d = String(v.getDate()).padStart(2,"0");
    return `${y}-${m}-${d}`;
  }
  return v !== undefined && v !== null ? String(v) : "";
}

function round2(n) { return Math.round(n * 100) / 100; }

function sheetRows(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0];
  return vals.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = fmtCell(r[i]); });
    return obj;
  });
}

function ok(data) { return { status: "ok", data }; }
function err(msg) { return { status: "error", message: msg }; }
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
