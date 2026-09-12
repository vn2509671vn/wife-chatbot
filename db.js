const { DatabaseSync } = require('node:sqlite');
const path = require('path');
const fs = require('fs');

// DATA_DIR: thu muc luu du lieu lau dai. Khi deploy len Render, dat bien moi truong
// DATA_DIR tro vao persistent disk (vd: /var/data) de du lieu khong bi mat khi deploy lai.
// Chay local thi mac dinh luu ngay trong thu muc du an.
const DATA_DIR = process.env.DATA_DIR || __dirname;
fs.mkdirSync(DATA_DIR, { recursive: true });

const db = new DatabaseSync(path.join(DATA_DIR, 'data.sqlite'));
db.exec('PRAGMA journal_mode = WAL;');

// ==== SCHEMA ====
db.exec(`
CREATE TABLE IF NOT EXISTS messages (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  session_id TEXT NOT NULL,
  role TEXT NOT NULL,
  content TEXT NOT NULL,
  created_at INTEGER NOT NULL
);

-- File tam thoi: da upload/da xu ly nhung CHUA chac chan luu lai
CREATE TABLE IF NOT EXISTS temp_files (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  session_id TEXT NOT NULL,
  customer_name TEXT NOT NULL,
  original_filename TEXT NOT NULL,
  stored_path TEXT NOT NULL,
  processed_path TEXT,
  created_at INTEGER NOT NULL
);

-- File CHINH THUC da duoc luu vao ho so khach hang
CREATE TABLE IF NOT EXISTS customer_files (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  session_id TEXT NOT NULL,
  customer_name TEXT NOT NULL,
  original_filename TEXT NOT NULL,
  stored_path TEXT,
  processed_path TEXT,
  note TEXT,
  created_at INTEGER NOT NULL
);

-- Tung dong du lieu cua khach hang, co the gan rieng cho tung nhan vien
CREATE TABLE IF NOT EXISTS customer_rows (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  customer_file_id INTEGER NOT NULL,
  row_index INTEGER NOT NULL,
  data_json TEXT NOT NULL,
  assigned_employee TEXT,
  assigned_at INTEGER,
  FOREIGN KEY (customer_file_id) REFERENCES customer_files(id)
);

-- Danh muc khach hang (quan ly rieng, doc lap voi customer_files)
CREATE TABLE IF NOT EXISTS customers (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE,
  created_at INTEGER NOT NULL
);

-- Danh muc nhan vien
CREATE TABLE IF NOT EXISTS employees (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE,
  created_at INTEGER NOT NULL
);

-- File dinh kem tam trong Chat (khong lien quan khach hang, dung 1 lan roi xoa)
CREATE TABLE IF NOT EXISTS chat_attachments (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  session_id TEXT NOT NULL,
  original_filename TEXT NOT NULL,
  stored_path TEXT NOT NULL,
  has_header INTEGER DEFAULT 1,
  created_at INTEGER NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_messages_session ON messages(session_id);
CREATE INDEX IF NOT EXISTS idx_messages_time ON messages(created_at);
CREATE INDEX IF NOT EXISTS idx_rows_file ON customer_rows(customer_file_id);
CREATE INDEX IF NOT EXISTS idx_rows_assigned ON customer_rows(assigned_employee, assigned_at);
`);

// Migration an toan: them cot has_header neu database cu chua co (bo qua loi neu da ton tai)
try { db.exec(`ALTER TABLE temp_files ADD COLUMN has_header INTEGER DEFAULT 1`); } catch (e) {}
try { db.exec(`ALTER TABLE customer_files ADD COLUMN has_header INTEGER DEFAULT 1`); } catch (e) {}

const TWO_DAYS_MS = 2 * 24 * 60 * 60 * 1000;

function purgeOldMessages() {
  const cutoff = Date.now() - TWO_DAYS_MS;
  db.prepare(`DELETE FROM messages WHERE created_at < ?`).run(cutoff);
}

// Tra ve danh sach temp_files da qua han (2 ngay) de server.js xoa file vat ly tren dia
function getExpiredTempFiles() {
  const cutoff = Date.now() - TWO_DAYS_MS;
  return db.prepare(`SELECT * FROM temp_files WHERE created_at < ?`).all(cutoff);
}
function deleteTempFile(id) {
  db.prepare(`DELETE FROM temp_files WHERE id = ?`).run(id);
}

// ---- chat_attachments ----
function addChatAttachment({ sessionId, originalFilename, storedPath, hasHeader = true }) {
  const info = db.prepare(
    `INSERT INTO chat_attachments (session_id, original_filename, stored_path, has_header, created_at)
     VALUES (?, ?, ?, ?, ?)`
  ).run(sessionId, originalFilename, storedPath, hasHeader ? 1 : 0, Date.now());
  return Number(info.lastInsertRowid);
}
function getChatAttachment(id) {
  return db.prepare(`SELECT * FROM chat_attachments WHERE id = ?`).get(id);
}
function deleteChatAttachment(id) {
  db.prepare(`DELETE FROM chat_attachments WHERE id = ?`).run(id);
}
function getExpiredChatAttachments() {
  const cutoff = Date.now() - TWO_DAYS_MS;
  return db.prepare(`SELECT * FROM chat_attachments WHERE created_at < ?`).all(cutoff);
}

// ---- messages ----
function addMessage(sessionId, role, content) {
  db.prepare(`INSERT INTO messages (session_id, role, content, created_at) VALUES (?, ?, ?, ?)`)
    .run(sessionId, role, content, Date.now());
}
function getRecentHistory(sessionId, limit = 30) {
  const cutoff = Date.now() - TWO_DAYS_MS;
  return db.prepare(
    `SELECT role, content FROM messages WHERE session_id = ? AND created_at >= ? ORDER BY created_at ASC LIMIT ?`
  ).all(sessionId, cutoff, limit);
}

// ---- temp_files ----
function addTempFile({ sessionId, customerName, originalFilename, storedPath, hasHeader = true }) {
  const info = db.prepare(
    `INSERT INTO temp_files (session_id, customer_name, original_filename, stored_path, has_header, created_at)
     VALUES (?, ?, ?, ?, ?, ?)`
  ).run(sessionId, customerName, originalFilename, storedPath, hasHeader ? 1 : 0, Date.now());
  return Number(info.lastInsertRowid);
}
function setTempProcessedPath(tempId, processedPath) {
  db.prepare(`UPDATE temp_files SET processed_path = ? WHERE id = ?`).run(processedPath, tempId);
}
function getTempFile(tempId) {
  return db.prepare(`SELECT * FROM temp_files WHERE id = ?`).get(tempId);
}

// ---- customer_files & customer_rows ----
function saveCustomerFile({ sessionId, customerName, originalFilename, storedPath, processedPath, rows, hasHeader = true }) {
  const info = db.prepare(
    `INSERT INTO customer_files (session_id, customer_name, original_filename, stored_path, processed_path, has_header, created_at)
     VALUES (?, ?, ?, ?, ?, ?, ?)`
  ).run(sessionId, customerName, originalFilename, storedPath || null, processedPath || null, hasHeader ? 1 : 0, Date.now());
  const customerFileId = Number(info.lastInsertRowid);

  const insertRow = db.prepare(
    `INSERT INTO customer_rows (customer_file_id, row_index, data_json) VALUES (?, ?, ?)`
  );
  db.exec('BEGIN');
  try {
    rows.forEach((r, idx) => insertRow.run(customerFileId, idx, JSON.stringify(r)));
    db.exec('COMMIT');
  } catch (e) {
    db.exec('ROLLBACK');
    throw e;
  }

  return customerFileId;
}

function getDistinctCustomers() {
  return db.prepare(
    `SELECT DISTINCT customer_name FROM customer_files ORDER BY customer_name ASC`
  ).all().map(r => r.customer_name);
}

// ---- danh muc khach hang ----
function addCustomer(name) {
  const trimmed = name.trim();
  if (!trimmed) throw new Error('Tên khách hàng không được để trống');
  try {
    const info = db.prepare(`INSERT INTO customers (name, created_at) VALUES (?, ?)`)
      .run(trimmed, Date.now());
    return { id: Number(info.lastInsertRowid), name: trimmed };
  } catch (e) {
    if (e.message.includes('UNIQUE')) throw new Error('Khách hàng này đã có trong danh mục');
    throw e;
  }
}
function listCustomersCatalog() {
  return db.prepare(`SELECT * FROM customers ORDER BY name ASC`).all();
}
function deleteCustomer(id) {
  db.prepare(`DELETE FROM customers WHERE id = ?`).run(id);
}

// ---- danh muc nhan vien ----
function addEmployee(name) {
  const trimmed = name.trim();
  if (!trimmed) throw new Error('Tên nhân viên không được để trống');
  try {
    const info = db.prepare(`INSERT INTO employees (name, created_at) VALUES (?, ?)`)
      .run(trimmed, Date.now());
    return { id: Number(info.lastInsertRowid), name: trimmed };
  } catch (e) {
    if (e.message.includes('UNIQUE')) throw new Error('Nhân viên này đã có trong danh mục');
    throw e;
  }
}
function listEmployeesCatalog() {
  return db.prepare(`SELECT * FROM employees ORDER BY name ASC`).all();
}
function deleteEmployee(id) {
  db.prepare(`DELETE FROM employees WHERE id = ?`).run(id);
}

function listCustomerFiles(sessionId) {
  const files = db.prepare(
    `SELECT * FROM customer_files WHERE session_id = ? ORDER BY created_at DESC`
  ).all(sessionId);

  const totalStmt = db.prepare(`SELECT COUNT(*) c FROM customer_rows WHERE customer_file_id = ?`);
  const unassignedStmt = db.prepare(
    `SELECT COUNT(*) c FROM customer_rows WHERE customer_file_id = ? AND assigned_employee IS NULL`
  );

  return files.map(f => ({
    ...f,
    totalRows: totalStmt.get(f.id).c,
    unassignedRows: unassignedStmt.get(f.id).c,
  }));
}

function getCustomerFile(id) {
  return db.prepare(`SELECT * FROM customer_files WHERE id = ?`).get(id);
}

// Gan N dong CHUA gan (theo thu tu row_index) cho 1 nhan vien.
// Tra ve du lieu cua CAC DONG VUA GAN (khong chi so luong) de con dung sinh tin nhan SMS.
function assignRowsToEmployee(customerFileId, employeeName, rowCount) {
  const rows = db.prepare(
    `SELECT id, data_json FROM customer_rows WHERE customer_file_id = ? AND assigned_employee IS NULL
     ORDER BY row_index ASC LIMIT ?`
  ).all(customerFileId, rowCount);

  const update = db.prepare(`UPDATE customer_rows SET assigned_employee = ?, assigned_at = ? WHERE id = ?`);
  const now = Date.now();
  db.exec('BEGIN');
  try {
    rows.forEach(r => update.run(employeeName, now, r.id));
    db.exec('COMMIT');
  } catch (e) {
    db.exec('ROLLBACK');
    throw e;
  }

  return rows.map(r => JSON.parse(r.data_json));
}

// Thu hoi (huy gan) toan bo dong da giao cho 1 nhan vien cu the trong 1 ho so khach hang
function revokeAssignmentForEmployee(customerFileId, employeeName) {
  const result = db.prepare(
    `UPDATE customer_rows SET assigned_employee = NULL, assigned_at = NULL
     WHERE customer_file_id = ? AND assigned_employee = ?`
  ).run(customerFileId, employeeName);
  return result.changes;
}

// Thu hoi TOAN BO phan cong cua 1 ho so khach hang (tra tat ca ve trang thai chua gan)
function revokeAllAssignments(customerFileId) {
  const result = db.prepare(
    `UPDATE customer_rows SET assigned_employee = NULL, assigned_at = NULL
     WHERE customer_file_id = ? AND assigned_employee IS NOT NULL`
  ).run(customerFileId);
  return result.changes;
}

// Trang thai phan cong cua 1 file: tong dong, con lai, va breakdown theo nhan vien
function getAssignmentStatus(customerFileId) {
  const total = db.prepare(`SELECT COUNT(*) c FROM customer_rows WHERE customer_file_id = ?`).get(customerFileId).c;
  const unassigned = db.prepare(
    `SELECT COUNT(*) c FROM customer_rows WHERE customer_file_id = ? AND assigned_employee IS NULL`
  ).get(customerFileId).c;
  const breakdown = db.prepare(
    `SELECT assigned_employee as employee, COUNT(*) as rowCount
     FROM customer_rows WHERE customer_file_id = ? AND assigned_employee IS NOT NULL
     GROUP BY assigned_employee`
  ).all(customerFileId);
  return { total, unassigned, breakdown };
}

// Thong ke trong ngay: khach hang -> nhan vien -> so dong, thoi gian
function getTodaySummary() {
  const startOfDay = new Date();
  startOfDay.setHours(0, 0, 0, 0);
  const cutoff = startOfDay.getTime();

  return db.prepare(
    `SELECT cf.customer_name, cr.assigned_employee as employee_name,
            COUNT(*) as row_count, MIN(cr.assigned_at) as first_time, MAX(cr.assigned_at) as last_time
     FROM customer_rows cr
     JOIN customer_files cf ON cf.id = cr.customer_file_id
     WHERE cr.assigned_employee IS NOT NULL AND cr.assigned_at >= ?
     GROUP BY cf.id, cr.assigned_employee
     ORDER BY first_time ASC`
  ).all(cutoff);
}

// Chi tiet TUNG DONG da gan trong ngay: khach hang nao, dong so may, gan cho ai, luc nao
function getTodayDetailedAssignments() {
  const startOfDay = new Date();
  startOfDay.setHours(0, 0, 0, 0);
  const cutoff = startOfDay.getTime();

  return db.prepare(
    `SELECT cf.customer_name, cf.original_filename, cr.row_index, cr.data_json,
            cr.assigned_employee as employee_name, cr.assigned_at
     FROM customer_rows cr
     JOIN customer_files cf ON cf.id = cr.customer_file_id
     WHERE cr.assigned_employee IS NOT NULL AND cr.assigned_at >= ?
     ORDER BY cr.assigned_at ASC, cr.row_index ASC`
  ).all(cutoff);
}

module.exports = {
  db,
  purgeOldMessages,
  getExpiredTempFiles,
  deleteTempFile,
  addChatAttachment,
  getChatAttachment,
  deleteChatAttachment,
  getExpiredChatAttachments,
  addMessage,
  getRecentHistory,
  addTempFile,
  setTempProcessedPath,
  getTempFile,
  saveCustomerFile,
  getDistinctCustomers,
  addCustomer,
  listCustomersCatalog,
  deleteCustomer,
  addEmployee,
  listEmployeesCatalog,
  deleteEmployee,
  listCustomerFiles,
  getCustomerFile,
  assignRowsToEmployee,
  revokeAssignmentForEmployee,
  revokeAllAssignments,
  getAssignmentStatus,
  getTodaySummary,
  getTodayDetailedAssignments,
};
