// =============================================
// 北海道機位管理系統 - Code.gs
// =============================================

const FLIGHT_INFO = {
  'BR116': { airline: 'BR', dest: '新千歲(CTS)', name: '長榮BR116' },
  'BR136': { airline: 'BR', dest: '旭川(AKJ)',   name: '長榮BR136旭川包機' },
  'BR166': { airline: 'BR', dest: '新千歲(CTS)', name: '長榮BR166' },
  'CI130': { airline: 'CI', dest: '新千歲(CTS)', name: '華航CI130'  },
  'JX850': { airline: 'JX', dest: '新千歲(CTS)', name: '星宇JX850' },
  'JX860': { airline: 'JX', dest: '新千歲(CTS)', name: '星宇JX860' },
};

// API 密鑰（LINE BOT 呼叫用）——部署前請改成自訂亂數字串
const API_KEY = 'hokkaido-seat-2026';

// ── Web App 進入點（POST：批次同步）─────────────
function doPost(e) {
  try {
    const key = e.parameter && e.parameter.key || '';
    if (key !== API_KEY) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const payload = JSON.parse(e.postData.contents);
    const result  = _apiSyncBatch(payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Web App 進入點 ────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';

  // LINE BOT API 端點
  if (action) {
    const key = e.parameter.key || '';
    if (key !== API_KEY) {
      return ContentService.createTextOutput(JSON.stringify({ error: 'Unauthorized' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    return _handleApiAction(action, e.parameter);
  }

  // 一般網頁介面
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('北海道機位管理系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── LINE BOT API 路由 ─────────────────────────
function _handleApiAction(action, params) {
  let result;
  try {
    switch (action) {
      case 'getSeats':
        result = _apiGetSeats(params);
        break;
      case 'getMonthSeats':
        result = _apiGetMonthSeats(params);
        break;
      case 'updateSeats':
        result = _apiUpdateSeats(params);
        break;
      case 'syncFlight':
        result = _apiSyncFlight(params);
        break;
      case 'getPaidGroups':
        result = _apiGetPaidGroups(params);
        break;
      case 'getJxpGroups':
        result = _apiGetJxpGroups(params);
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function _apiGetSeats(params) {
  const date     = params.date     || '';
  const airline  = params.airline  || '';
  const flightNo = params.flightNo || '';

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('flights');
  if (!sh || sh.getLastRow() < 2) return { flights: [] };

  const data = sh.getDataRange().getValues();
  const h = data[0];
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const obj = _rowToObj(h, data[i]);
    if (date    && obj.date     !== date)       continue;
    if (airline && obj.airline  !== airline)    continue;
    if (flightNo && obj.flight_no !== flightNo) continue;
    rows.push(obj);
  }
  return { flights: rows };
}

function _apiGetMonthSeats(params) {
  const year  = parseInt(params.year)  || new Date().getFullYear();
  const month = parseInt(params.month) || new Date().getMonth() + 1;
  return { flights: getFlightsForMonth(year, month) };
}

// 單筆 upsert（GET 用）
function _apiSyncFlight(params) {
  const items = [{
    date:           params.date       || '',
    flight_no:      params.flightNo   || '',
    original_seats: parseInt(params.originalSeats) || 0,
    current_seats:  parseInt(params.currentSeats)  || 0,
    used_seats:     parseInt(params.usedSeats)     || 0,
    type:           params.type       || 'SRS',
    note:           params.note       || 'ERP同步',
  }];
  return _upsertFlights(items);
}

// 批次 upsert（POST 用，接收 { items: [...], paid_groups: [...], jxp_groups: [...] }）
function _apiSyncBatch(payload) {
  const items = (payload && payload.items) || [];
  if (items.length === 0) return { error: 'items 為空' };
  const flightResult = _upsertFlights(items);

  const paidGroups = (payload && payload.paid_groups) || [];
  if (paidGroups.length > 0) _upsertPaidGroups(paidGroups);

  const jxpGroups = (payload && payload.jxp_groups) || [];
  if (jxpGroups.length > 0) _upsertJxpGroups(jxpGroups);

  return flightResult;
}

// ── 核心 upsert 邏輯（批次讀寫，大幅減少 Sheets API 呼叫次數）────
function _upsertFlights(items) {
  const sh = _ensureSheet('flights', [
    'date', 'airline', 'flight_no', 'dest',
    'type', 'original_seats', 'current_seats', 'used_seats',
    'note', 'created_at', 'created_by'
  ]);
  const data = sh.getDataRange().getValues();
  const h    = data[0];
  const dateIdx   = h.indexOf('date');
  const flightIdx = h.indexOf('flight_no');
  const noteIdx   = h.indexOf('note');

  const results  = { updated: 0, created: 0, errors: [] };
  const toUpdate = []; // { sheetRow, values }
  const toInsert = []; // 待新增的列

  // 建立查找 Map：key = "date|flight_no|daysPrefix"，value = data 陣列索引
  // 一次掃描全表，後續查找 O(1)
  const rowMap = {};
  for (let i = 1; i < data.length; i++) {
    if (!data[i][dateIdx]) continue;
    const rowDate   = _fmtDate(new Date(data[i][dateIdx]));
    const rowFlight = String(data[i][flightIdx]);
    const rowNote   = String(data[i][noteIdx] || '');
    const dpMatch   = rowNote.match(/^(\d+天\|)/);
    const dp        = dpMatch ? dpMatch[1] : '';
    rowMap[rowDate + '|' + rowFlight + '|' + dp] = i;
  }

  items.forEach(item => {
    if (!item.date || !item.flight_no) {
      results.errors.push('缺 date/flight_no: ' + JSON.stringify(item));
      return;
    }
    const p        = parseFlightNo(item.flight_no);
    const dp       = item.days ? item.days + '天|' : '';
    const existIdx = rowMap[item.date + '|' + item.flight_no + '|' + dp];

    if (existIdx !== undefined) {
      // 在記憶體修改該列，稍後一次整行寫回
      const newRow = data[existIdx].slice();
      newRow[h.indexOf('current_seats')] = item.current_seats;
      newRow[h.indexOf('used_seats')]    = item.used_seats;
      if (item.type) newRow[h.indexOf('type')] = item.type;
      if (item.note) newRow[h.indexOf('note')] = item.note;
      toUpdate.push({ sheetRow: existIdx + 1, values: newRow });
      results.updated++;
    } else {
      toInsert.push([
        new Date(item.date), p.airline, item.flight_no, p.dest,
        item.type || 'SRS', item.original_seats, item.current_seats, item.used_seats,
        item.note || 'ERP同步', new Date(), 'Chrome擴充'
      ]);
      results.created++;
    }
  });

  // 每列整行寫回（1 次 setValues 取代原本 5 次 setValue）
  toUpdate.forEach(u => {
    sh.getRange(u.sheetRow, 1, 1, u.values.length).setValues([u.values]);
  });

  // 所有新列一次批次寫入
  if (toInsert.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, toInsert.length, h.length).setValues(toInsert);
  }

  return results;
}

function _apiUpdateSeats(params) {
  const date     = params.date     || '';
  const flightNo = params.flightNo || '';
  const delta    = params.delta    || '';  // "+20", "-5", "=50"

  if (!date || !flightNo || !delta) {
    return { error: 'date, flightNo, delta 均必填' };
  }

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('flights');
  if (!sh || sh.getLastRow() < 2) return { error: 'flights sheet 不存在' };

  const data = sh.getDataRange().getValues();
  const h = data[0];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const obj = _rowToObj(h, data[i]);
    if (obj.date !== date || obj.flight_no !== flightNo) continue;

    const row     = i + 1;
    const currIdx = h.indexOf('current_seats') + 1;
    const curr    = Number(data[i][h.indexOf('current_seats')]) || 0;

    let newVal;
    if (delta.startsWith('=')) {
      newVal = parseInt(delta.slice(1));
    } else {
      newVal = curr + parseInt(delta);
    }
    if (isNaN(newVal)) return { error: 'delta 格式錯誤，請用 +20 / -5 / =50' };

    sh.getRange(row, currIdx).setValue(newVal);
    _logChange({
      flight_date: date,
      airline: obj.airline,
      flight_no: flightNo,
      change_type: delta.startsWith('=') ? '直接設值' : (parseInt(delta) >= 0 ? '增加' : '減少'),
      seats_delta: newVal - curr,
      changed_by: 'LINE BOT',
      note: `LINE 指令：${delta}`
    });
    return { success: true, flight_no: flightNo, date, prev: curr, new: newVal };
  }
  return { error: `找不到 ${flightNo} ${date}` };
}

// ── 初始化資料表 ──────────────────────────────
function initSheets() {
  _ensureSheet('flights', [
    'date', 'airline', 'flight_no', 'dest',
    'type', 'original_seats', 'current_seats', 'used_seats',
    'note', 'created_at', 'created_by'
  ]);
  _ensureSheet('seat_changes', [
    'change_id', 'flight_date', 'airline', 'flight_no', 'change_date',
    'change_type', 'seats_delta', 'changed_by', 'note'
  ]);
  _ensureSheet('deposit_cases', [
    'case_id', 'case_name', 'airline', 'deposit_type',
    'date_from', 'date_to', 'flight_no', 'ptn_entries',
    'deposit_per_seat', 'conversion_rate', 'calc_mode',
    'merge_with_group', 'deadline', 'status', 'note',
    'created_at', 'updated_at', 'created_by', 'locked_seats'
  ]);
  _ensureSheet('deposit_changes', [
    'change_id', 'case_id', 'changed_at', 'changed_by', 'description'
  ]);
  _ensureSheet('feedback', [
    'feedback_id', 'content', 'created_at', 'status', 'reply', 'replied_at'
  ]);
  return 'OK';
}

function _ensureSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// ── 航班解析（用航班號查 FLIGHT_INFO）─────────
function parseFlightNo(flightNo) {
  flightNo = (flightNo || '').toUpperCase().trim();
  const info = FLIGHT_INFO[flightNo];
  if (info) return { flightNo, airline: info.airline, dest: info.dest };
  const airline = flightNo.match(/^[A-Z]{2}/)?.[0] || '';
  return { flightNo, airline, dest: '新千歲(CTS)' };
}

// ── 取得 FLIGHT_INFO（供前端呼叫）────────────
function getFlightInfo() {
  return FLIGHT_INFO;
}

// ── 航班 CRUD ─────────────────────────────────
function getFlightsForMonth(year, month) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('flights');
  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const from = new Date(year, month - 1, 1);
  const to   = new Date(year, month, 0, 23, 59, 59);
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const d = new Date(data[i][0]);
    if (d >= from && d <= to) rows.push(_rowToObj(headers, data[i]));
  }
  return rows;
}

function getFlightsForDate(dateStr) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('flights');
  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (_fmtDate(new Date(data[i][0])) === dateStr) rows.push(_rowToObj(headers, data[i]));
  }
  return rows;
}

function addFlight(d) {
  const sh = _ensureSheet('flights', [
    'date', 'airline', 'flight_no', 'dest',
    'type', 'original_seats', 'current_seats', 'used_seats',
    'note', 'created_at', 'created_by'
  ]);
  const p = parseFlightNo(d.flight_no);
  sh.appendRow([
    new Date(d.date), p.airline, p.flightNo, p.dest,
    d.type, +d.original_seats, +d.original_seats, 0,
    d.note || '', new Date(),
    Session.getActiveUser().getEmail() || 'system'
  ]);
  return { success: true };
}

function updateFlight(d) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('flights');
  const data = sh.getDataRange().getValues();
  const h = data[0];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const rowDate   = _fmtDate(new Date(data[i][h.indexOf('date')]));
    const rowFlight = data[i][h.indexOf('flight_no')];

    if (rowDate === d.date && rowFlight === d.original_flight_no) {
      if (d.change_type) {
        _logChange({
          flight_date: d.date, airline: d.airline,
          flight_no: d.original_flight_no,
          change_type: d.change_type,
          seats_delta: d.seats_delta || 0,
          changed_by: d.changed_by || Session.getActiveUser().getEmail() || 'system',
          note: d.change_note || ''
        });
      }

      const row = i + 1;
      const set = (col, val) => sh.getRange(row, h.indexOf(col) + 1).setValue(val);

      if (d.flight_no && d.flight_no !== d.original_flight_no) {
        const p = parseFlightNo(d.flight_no);
        set('flight_no', p.flightNo);
        set('airline',   p.airline);
        set('dest',      p.dest);
      }
      if (d.current_seats !== undefined) set('current_seats', +d.current_seats);
      if (d.used_seats    !== undefined) set('used_seats',    +d.used_seats);
      if (d.type          !== undefined) set('type',           d.type);
      if (d.note          !== undefined) set('note',           d.note);

      return { success: true };
    }
  }
  return { success: false, error: 'Flight not found' };
}

function deleteFlight(date, flightNo) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('flights');
  const data = sh.getDataRange().getValues();
  const h = data[0];

  for (let i = data.length - 1; i >= 1; i--) {
    if (!data[i][0]) continue;
    if (_fmtDate(new Date(data[i][h.indexOf('date')])) === date &&
        data[i][h.indexOf('flight_no')] === flightNo) {
      sh.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

// ── 異動紀錄 ──────────────────────────────────
function _logChange(d) {
  const sh = _ensureSheet('seat_changes', [
    'change_id', 'flight_date', 'airline', 'flight_no', 'change_date',
    'change_type', 'seats_delta', 'changed_by', 'note'
  ]);
  sh.appendRow([
    'C' + Date.now(), new Date(d.flight_date), d.airline, d.flight_no,
    new Date(), d.change_type, d.seats_delta,
    d.changed_by, d.note || ''
  ]);
}

function getChangeHistory(date, flightNo) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('seat_changes');
  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getDataRange().getValues();
  const h = data[0];
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const matchDate   = _fmtDate(new Date(data[i][h.indexOf('flight_date')])) === date;
    const matchFlight = data[i][h.indexOf('flight_no')] === flightNo;
    if (matchDate && matchFlight) {
      const obj = _rowToObj(h, data[i]);
      obj.change_date = data[i][h.indexOf('change_date')]
        ? Utilities.formatDate(new Date(data[i][h.indexOf('change_date')]), 'Asia/Taipei', 'yyyy-MM-dd HH:mm')
        : '';
      rows.push(obj);
    }
  }
  return rows;
}

// ── 工具函式 ──────────────────────────────────
function _fmtDate(d) {
  return Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd');
}

function _rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    let v = row[i];
    if (v instanceof Date) v = _fmtDate(v);
    obj[h] = v;
  });
  return obj;
}

function _logDepositChange(caseId, user, description) {
  const sh = _ensureSheet('deposit_changes', [
    'change_id', 'case_id', 'changed_at', 'changed_by', 'description'
  ]);
  sh.appendRow(['DC' + Date.now(), caseId, new Date(), user, description]);
}

// ── 付訂案件 ──────────────────────────────────
function getDepositCases() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('deposit_cases');
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const h = data[0];
  return data.slice(1).filter(r => r[0]).map(r => {
    const obj = _rowToObj(h, r);
    try { obj.ptn_entries = obj.ptn_entries ? JSON.parse(obj.ptn_entries) : []; }
    catch(e) { obj.ptn_entries = []; }
    try { obj.locked_seats = obj.locked_seats ? JSON.parse(obj.locked_seats) : null; }
    catch(e) { obj.locked_seats = null; }
    return obj;
  });
}

function saveDepositCase(d) {
  const sh = _ensureSheet('deposit_cases', [
    'case_id', 'case_name', 'airline', 'deposit_type',
    'date_from', 'date_to', 'flight_no', 'ptn_entries',
    'deposit_per_seat', 'conversion_rate', 'calc_mode',
    'merge_with_group', 'deadline', 'status', 'note',
    'created_at', 'updated_at', 'created_by', 'locked_seats'
  ]);
  const user = Session.getActiveUser().getEmail() || 'system';
  const now  = new Date();

  if (d.case_id) {
    const data = sh.getDataRange().getValues();
    const h = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === d.case_id) {
        const row = i + 1;
        const set = (col, val) => sh.getRange(row, h.indexOf(col) + 1).setValue(val);
        set('case_name',        d.case_name || '');
        set('airline',          d.airline || '');
        set('deposit_type',     d.deposit_type || '');
        set('date_from',        d.date_from || '');
        set('date_to',          d.date_to || '');
        set('flight_no',        d.flight_no || '');
        set('ptn_entries',      JSON.stringify(d.ptn_entries || []));
        set('deposit_per_seat', d.deposit_per_seat || 0);
        set('conversion_rate',  d.conversion_rate || 1);
        set('calc_mode',        d.calc_mode || '');
        set('merge_with_group', d.merge_with_group || '');
        set('deadline',         d.deadline || '');
        set('status',           d.status || '進行中');
        set('note',             d.note || '');
        set('updated_at',       now);
        set('locked_seats',     JSON.stringify(d.locked_seats || null));
        _logDepositChange(d.case_id, user, '更新案件');
        return { success: true, case_id: d.case_id };
      }
    }
  }

  const caseId = 'D' + Date.now();
  sh.appendRow([
    caseId, d.case_name || '', d.airline || '', d.deposit_type || '',
    d.date_from || '', d.date_to || '', d.flight_no || '',
    JSON.stringify(d.ptn_entries || []),
    d.deposit_per_seat || 0, d.conversion_rate || 1, d.calc_mode || '',
    d.merge_with_group || '', d.deadline || '', d.status || '進行中', d.note || '',
    now, now, user, JSON.stringify(d.locked_seats || null)
  ]);
  _logDepositChange(caseId, user, '新增案件');
  return { success: true, case_id: caseId };
}

function deleteDepositCase(caseId) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('deposit_cases');
  if (!sh) return { success: false };
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === caseId) {
      sh.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false };
}

// ── 付訂成行率 ────────────────────────────────
function _upsertPaidGroups(groups) {
  const sh = _ensureSheet('paid_groups', [
    'date', 'group_no', 'airline', 'flight_no',
    'remark', 'total_seats', 'hk', 'kk', 'synced_at'
  ]);
  const data = sh.getDataRange().getValues();
  const h    = data[0];
  const gIdx = h.indexOf('group_no');

  // 建 Map：group_no → 列索引
  const rowMap = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][gIdx]) rowMap[String(data[i][gIdx])] = i;
  }

  const toUpdate = [], toInsert = [];
  groups.forEach(item => {
    const existIdx = rowMap[item.group_no];
    if (existIdx !== undefined) {
      const newRow = data[existIdx].slice();
      newRow[h.indexOf('hk')]          = item.hk || 0;
      newRow[h.indexOf('kk')]          = item.kk || 0;
      newRow[h.indexOf('total_seats')] = item.total_seats || 0;
      newRow[h.indexOf('remark')]      = item.remark || '';
      newRow[h.indexOf('synced_at')]   = new Date();
      toUpdate.push({ sheetRow: existIdx + 1, values: newRow });
    } else {
      toInsert.push([
        item.date, item.group_no, item.airline, item.flight_no,
        item.remark || '', item.total_seats || 0, item.hk || 0, item.kk || 0,
        new Date()
      ]);
    }
  });

  toUpdate.forEach(u => sh.getRange(u.sheetRow, 1, 1, u.values.length).setValues([u.values]));
  if (toInsert.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, toInsert.length, h.length).setValues(toInsert);
  }
}

function _apiGetPaidGroups(params) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('paid_groups');
  if (!sh || sh.getLastRow() < 2) return { groups: [] };
  const data = sh.getDataRange().getValues();
  const h = data[0];
  return {
    groups: data.slice(1).filter(r => r[0]).map(r => _rowToObj(h, r))
  };
}

// ── JX 暑假 P 階段團資料 ─────────────────────
function _upsertJxpGroups(groups) {
  const sh = _ensureSheet('jxp_groups', [
    'date', 'group_no', 'airline', 'flight_no',
    'total_seats', 'hk', 'kk', 'synced_at'
  ]);
  // 每次全量覆寫，確保已排除的團（如 * 共賣）不殘留
  if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  if (groups.length === 0) return;
  const rows = groups.filter(item => item.group_no).map(item => [
    item.date, item.group_no, item.airline, item.flight_no,
    item.total_seats || 0, item.hk || 0, item.kk || 0,
    new Date()
  ]);
  if (rows.length > 0) sh.getRange(2, 1, rows.length, 8).setValues(rows);
}

function _apiGetJxpGroups(params) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('jxp_groups');
  if (!sh || sh.getLastRow() < 2) return { groups: [] };
  const data = sh.getDataRange().getValues();
  const h = data[0];
  return {
    groups: data.slice(1).filter(r => r[0]).map(r => _rowToObj(h, r))
  };
}

// ── 意見回饋 ──────────────────────────────────
function submitFeedback(content) {
  const sh = _ensureSheet('feedback', [
    'feedback_id', 'content', 'created_at', 'status', 'reply', 'replied_at'
  ]);
  sh.appendRow(['F' + Date.now(), content, new Date(), '待處理', '', '']);
  return { success: true };
}
