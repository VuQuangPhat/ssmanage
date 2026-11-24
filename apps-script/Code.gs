/**
 * Quản lý thu mua – xuất khẩu thủy hải sản
 * Mã Apps Script gắn với Google Sheet (Tiện ích mở rộng > Apps Script).
 * Dán file này vào Code.gs và giao diện vào một file HTML tên Index.
 */

// Để trống nếu mã này mở từ chính Google Sheet (Extensions > Apps Script).
// Nếu mã nằm riêng, dán ID của Google Sheet vào đây (đoạn giữa /d/ và /edit trong link Sheet).
var SHEET_ID = '13D_F8RS0WjWfkBCuu3OGyI8NlgYNdhZnvI5HNET-jY4'; // Sheet QUẢN LÝ SUPER SEAFOODS; để trống nếu script gắn trực tiếp trong Sheet

var TABS = ['NHANH', 'DOI_TAC', 'HOP_DONG', 'LO_HANG', 'HD_LIEN_KET', 'CONG_NO', 'UNG_TRUOC', 'NCC_THI_TRUONG', 'THI_TRUONG', 'YEU_CAU_THI_TRUONG', 'CONG_TY', 'KHAU', 'BUOC_DON', 'LIEN_HE', 'CAI_DAT'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Quản lý thủy hải sản')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Chỉ cho phép trang ghi thêm dòng vào các tab này (nút Thêm hợp đồng, Ghi lần làm việc).
var WRITABLE = ['HOP_DONG', 'LIEN_HE', 'CONG_NO', 'UNG_TRUOC'];

// Chuyển giá trị gửi từ trang thành giá trị ô: chuỗi ngày thành ngày, chặn công thức do người dùng nhập.
function cellValue_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'string') {
    if (/^\d{4}-\d{2}-\d{2}$/.test(v)) { var p = v.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
    if (/^[=+\-@]/.test(v)) return "'" + v;
  }
  return v;
}

// Tab nào được sửa dòng có sẵn, và cột nào là khóa để tìm dòng.
var KEYCOL = { HOP_DONG: 'id', CONG_TY: 'khoa' };

/** Sửa các dòng có sẵn. items = [{key: 'HDB-001', obj: {tien_do_pct: 80}}]. Cần quyền chỉnh sửa Sheet. */
function updateRows(tab, items) {
  var kc = KEYCOL[tab];
  if (!kc) throw new Error('Không được sửa tab ' + tab + '.');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(tab);
    if (!sh) throw new Error('Không thấy tab ' + tab + '.');
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    var ki = head.indexOf(kc);
    if (ki < 0) throw new Error('Tab ' + tab + ' thiếu cột ' + kc + '.');
    var n = Math.max(sh.getLastRow() - 1, 0);
    var keys = n ? sh.getRange(2, ki + 1, n, 1).getValues() : [];
    (items || []).forEach(function (it) {
      var r = -1;
      for (var i = 0; i < keys.length; i++) { if (String(keys[i][0]).trim() === String(it.key).trim()) { r = i + 2; break; } }
      if (r < 0) throw new Error('Không thấy dòng ' + it.key + ' trong tab ' + tab + '.');
      Object.keys(it.obj || {}).forEach(function (f) {
        var ci = head.indexOf(f);
        if (ci < 0 || f === kc) return;
        sh.getRange(r, ci + 1).setValue(cellValue_(it.obj[f]));
      });
    });
  } finally {
    lock.releaseLock();
  }
  return true;
}

/** Thêm một dòng vào tab. Cần quyền chỉnh sửa Sheet; người chỉ có quyền xem sẽ nhận thông báo lỗi. */
function addRow(tab, obj) {
  if (WRITABLE.indexOf(tab) < 0) throw new Error('Không được ghi vào tab ' + tab + '.');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(tab);
    if (!sh) throw new Error('Không thấy tab ' + tab + '.');
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
    var row = head.map(function (h) { return cellValue_(obj[h]); });
    sh.appendRow(row);
  } finally {
    lock.releaseLock();
  }
  return true;
}

/** Trả dữ liệu các tab cho trang: mỗi tab là danh sách dòng, khóa là tên cột ở dòng đầu. */
function getData() {
  var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var out = {};
  TABS.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return; // trang sẽ báo thiếu tab nào
    var values = sh.getDataRange().getValues();
    if (values.length < 1) { out[name] = []; return; }
    var head = values[0].map(function (h) { return String(h).trim(); });
    var rows = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var empty = row.every(function (c) { return c === '' || c === null; });
      if (empty) continue;
      var obj = {};
      for (var j = 0; j < head.length; j++) {
        if (!head[j]) continue;
        var v = row[j];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        obj[head[j]] = v;
      }
      rows.push(obj);
    }
    out[name] = rows;
  });
  return out;
}
