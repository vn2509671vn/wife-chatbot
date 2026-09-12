require('dotenv').config();
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const fetch = require('node-fetch');

const db = require('./db');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const DATA_DIR = process.env.DATA_DIR || __dirname;
const UPLOAD_DIR = path.join(DATA_DIR, 'uploads');
const OUTPUT_DIR = path.join(DATA_DIR, 'outputs');
fs.mkdirSync(UPLOAD_DIR, { recursive: true });
fs.mkdirSync(OUTPUT_DIR, { recursive: true });

const upload = multer({ dest: UPLOAD_DIR });

// ---- don rac dinh ky: tin nhan cu > 2 ngay + file tam chua luu > 2 ngay ----
function purgeExpired() {
  db.purgeOldMessages();
  const expired = db.getExpiredTempFiles();
  for (const t of expired) {
    [t.stored_path, t.processed_path].forEach((p) => {
      if (p && fs.existsSync(p)) fs.unlinkSync(p);
    });
    db.deleteTempFile(t.id);
  }
  const expiredAttachments = db.getExpiredChatAttachments();
  for (const a of expiredAttachments) {
    if (a.stored_path && fs.existsSync(a.stored_path)) fs.unlinkSync(a.stored_path);
    db.deleteChatAttachment(a.id);
  }
}
purgeExpired();
setInterval(purgeExpired, 60 * 60 * 1000);

const OPENROUTER_API_KEY = process.env.OPENROUTER_API_KEY;
const MODEL = process.env.OPENROUTER_MODEL || 'anthropic/claude-sonnet-4.5';

async function callOpenRouter(messages) {
  if (!OPENROUTER_API_KEY) throw new Error('Thiếu OPENROUTER_API_KEY trong file .env');
  const res = await fetch('https://openrouter.ai/api/v1/chat/completions', {
    method: 'POST',
    headers: { Authorization: `Bearer ${OPENROUTER_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: MODEL, messages }),
  });
  if (!res.ok) throw new Error(`OpenRouter lỗi ${res.status}: ${await res.text()}`);
  const data = await res.json();
  return data.choices[0].message.content;
}

// Doc file xlsx thanh mang du lieu. hasHeader=true -> mang object (dong dau la ten cot).
// hasHeader=false -> mang cac mang (giu nguyen, khong coi dong nao la tieu de).
function readRows(filePath, hasHeader = true) {
  const wb = XLSX.readFile(filePath);
  const ws = wb.Sheets[wb.SheetNames[0]];
  return hasHeader
    ? XLSX.utils.sheet_to_json(ws, { defval: '' })
    : XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
}

// Tao worksheet tu du lieu, tu dong nhan dien: mang cac mang -> khong tieu de,
// mang cac object -> co tieu de (giu dung dinh dang nhu du lieu dau vao/AI tra ve).
function rowsToSheet(rows) {
  if (!rows || !rows.length) return XLSX.utils.aoa_to_sheet([[]]);
  return Array.isArray(rows[0]) ? XLSX.utils.aoa_to_sheet(rows) : XLSX.utils.json_to_sheet(rows);
}

function writeRowsToXlsxFile(rows, outPath) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, rowsToSheet(rows), 'Sheet1');
  XLSX.writeFile(wb, outPath);
}

// ================== QUY TAC CHUAN HOA SDT / SERI SIM ==================
// SDT: 9 so | 10 so bat dau bang 0 | 11 so bat dau bang 84  -> chuan hoa ve 9 so cuoi
// Seri sim: 10 so bat dau bang 1 hoac 2 (da chuan) | 20 so -> lay 11 so cuoi roi bo 1 so cuoi = 10 so
function onlyDigits(v) {
  return String(v).replace(/\D/g, '');
}
function isPhoneCandidate(raw) {
  const d = onlyDigits(raw);
  return d.length === 9 || (d.length === 10 && d[0] === '0') || (d.length === 11 && d.slice(0, 2) === '84');
}
function normalizePhone(raw) {
  const d = onlyDigits(raw);
  if (d.length === 9) return d;
  if (d.length === 10 && d[0] === '0') return d.slice(-9);
  if (d.length === 11 && d.slice(0, 2) === '84') return d.slice(-9);
  return null;
}
function isSimCandidate(raw) {
  const d = onlyDigits(raw);
  return (d.length === 10 && (d[0] === '1' || d[0] === '2')) || d.length === 20;
}
function normalizeSimSerial(raw) {
  const d = onlyDigits(raw);
  if (d.length === 10 && (d[0] === '1' || d[0] === '2')) return d;
  if (d.length === 20) return d.slice(-11, -1); // 11 so cuoi, bo 1 so cuoi cung -> con 10 so
  return null;
}
// Chi chuan hoa neu gia tri la chuoi CHI GOM CHU SO va khop dung 1 trong 2 dinh dang tren.
// Cac gia tri khac (tien, so luong, chu...) giu nguyen, khong dong cham.
function normalizeCellForPhoneSim(value) {
  if (value === null || value === undefined || value === '') return value;
  const str = String(value).trim();
  if (!/^\d+$/.test(str)) return value;
  if (isSimCandidate(str)) {
    const n = normalizeSimSerial(str);
    if (n !== null) return n;
  }
  if (isPhoneCandidate(str)) {
    const n = normalizePhone(str);
    if (n !== null) return n;
  }
  return value;
}
function normalizeRows(rows) {
  if (!Array.isArray(rows)) return rows;
  return rows.map((row) => {
    if (Array.isArray(row)) return row.map(normalizeCellForPhoneSim);
    if (row && typeof row === 'object') {
      const out = {};
      Object.entries(row).forEach(([k, v]) => { out[k] = normalizeCellForPhoneSim(v); });
      return out;
    }
    return normalizeCellForPhoneSim(row);
  });
}
// Tim SDT va Seri sim trong 1 dong du lieu (object hoac mang) DA duoc chuan hoa san.
// Luu y: seri sim rut gon tu 20 so co the KHONG con bat dau bang 1/2 nua (vi chi la cat chuoi),
// nen o day nhan dien theo DO DAI (da chuan hoa: SDT luon 9 so, Sim luon 10 so) thay vi kiem tra lai dinh dang goc.
function extractPhoneAndSim(rowData) {
  let phone = null;
  let sim = null;
  const values = Array.isArray(rowData) ? rowData : Object.values(rowData || {});
  for (const v of values) {
    if (v === null || v === undefined || v === '') continue;
    const str = String(v).trim();
    if (!/^\d+$/.test(str)) continue;
    if (sim === null && str.length === 10) { sim = str; continue; }
    if (phone === null && str.length === 9) { phone = str; continue; }
  }
  return { phone, sim };
}
// Sinh noi dung tin nhan de vo copy gui Zalo cho nhan vien: moi dong 1 cap SDT - Seri sim
function buildAssignmentSms(employeeName, rowsData) {
  const lines = rowsData.map((r) => {
    const { phone, sim } = extractPhoneAndSim(r);
    return `${phone || '(?)'} - ${sim || '(?)'}`;
  });
  return `Danh sách bàn giao cho ${employeeName} (${lines.length} số):\n${lines.join('\n')}`;
}
// Ghi chu quy tac nay de nhung vao prompt AI, giup AI hieu dung ngu canh nghiep vu
const PHONE_SIM_RULE_NOTE =
  'Lưu ý nghiệp vụ: dữ liệu trong hệ thống chỉ gồm SĐT (đã chuẩn hoá về đúng 9 số cuối) và Seri sim ' +
  '(đã chuẩn hoá về đúng 10 số, bắt đầu bằng 1 hoặc 2). Nếu xử lý/tạo ra các trường SĐT hoặc Seri sim, ' +
  'hãy giữ đúng định dạng đã chuẩn hoá này, không thêm/bớt số, không đổi lại về dạng gốc.';

// Phan tich van ban dan tay (copy tu Excel/Zalo/tin nhan...) thanh bang du lieu.
// Tu dong nhan dien dau phan cach: tab -> phay -> nhieu khoang trang -> cot don.
function parsePastedText(text, hasHeader) {
  const lines = text.replace(/\r\n/g, '\n').split('\n').filter((l) => l.trim().length > 0);
  if (!lines.length) return [];

  let splitFn;
  if (lines.every((l) => l.includes('\t'))) {
    splitFn = (l) => l.split('\t');
  } else {
    const commaCounts = lines.map((l) => (l.match(/,/g) || []).length);
    if (commaCounts[0] > 0 && commaCounts.every((c) => c === commaCounts[0])) {
      splitFn = (l) => l.split(',');
    } else if (lines.every((l) => /\s{2,}/.test(l))) {
      splitFn = (l) => l.split(/\s{2,}/);
    } else {
      splitFn = (l) => [l]; // moi dong 1 cot duy nhat
    }
  }

  const toCell = (c) => {
    const trimmed = c.trim();
    // Uu tien nhan dien SDT/Seri sim TRUOC khi thu chuyen thanh Number,
    // vi so 20 chu so se mat do chinh xac neu chuyen sang Number cua JS.
    if (isSimCandidate(trimmed)) {
      const n = normalizeSimSerial(trimmed);
      if (n !== null) return n;
    }
    if (isPhoneCandidate(trimmed)) {
      const n = normalizePhone(trimmed);
      if (n !== null) return n;
    }
    if (trimmed !== '' && !isNaN(trimmed) && !isNaN(parseFloat(trimmed))) return Number(trimmed);
    return trimmed;
  };

  const grid = lines.map((l) => splitFn(l).map(toCell));

  if (!hasHeader) return grid; // tra ve mang cac mang, khong tieu de

  const headers = grid[0].map((h, i) => (h === '' || h === undefined ? `Cột ${i + 1}` : String(h)));
  return grid.slice(1).map((row) => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
    return obj;
  });
}

// Trich khoi du lieu file (###FILE_DATA### ... ###END_FILE_DATA###) tu cau tra loi cua AI trong chat.
function extractFileBlock(text) {
  const match = text.match(/###FILE_DATA###([\s\S]*?)###END_FILE_DATA###/);
  if (!match) return { cleanText: text, rows: null };
  try {
    const rows = JSON.parse(match[1].trim());
    return { cleanText: text.replace(match[0], '').trim(), rows };
  } catch {
    return { cleanText: text, rows: null };
  }
}

// ================== DINH KEM FILE TRONG CHAT (khong lien quan khach hang) ==================
app.post('/api/chat/attach', upload.single('file'), (req, res) => {
  try {
    const { sessionId } = req.body;
    const hasHeader = req.body.hasHeader !== 'false';
    if (!req.file || !sessionId) return res.status(400).json({ error: 'Thiếu file hoặc sessionId' });

    const attachmentId = db.addChatAttachment({
      sessionId,
      originalFilename: req.file.originalname,
      storedPath: req.file.path,
      hasHeader,
    });
    const rows = readRows(req.file.path, hasHeader);
    res.json({ attachmentId, filename: req.file.originalname, rowCount: rows.length, preview: rows.slice(0, 5) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ================== CHAT (bo nho 2 ngay, co the tra ve file Excel) ==================
app.post('/api/chat', async (req, res) => {
  try {
    const { sessionId, message, attachmentId } = req.body;
    if (!sessionId || !message) return res.status(400).json({ error: 'Thiếu sessionId hoặc message' });

    // Neu co file dinh kem cho tin nhan nay: doc du lieu va nhung vao noi dung gui AI (chi dung 1 lan)
    let attachmentBlock = '';
    let attachmentNote = '';
    const MAX_ROWS = 500;
    if (attachmentId) {
      const attachment = db.getChatAttachment(attachmentId);
      if (attachment) {
        const hasHeader = !!attachment.has_header;
        const rows = normalizeRows(readRows(attachment.stored_path, hasHeader));
        const formatNote = hasHeader
          ? 'mảng object (mỗi object 1 dòng, tên cột là khoá)'
          : 'mảng các mảng (mỗi mảng con 1 dòng, KHÔNG có tiêu đề cột)';
        attachmentBlock =
          `\n\n[Dữ liệu đính kèm từ file "${attachment.original_filename}", ${rows.length} dòng, ` +
          `dạng ${formatNote}, KHÔNG liên quan gì tới khách hàng hay nhân viên trừ khi người dùng nói rõ]:\n` +
          JSON.stringify(rows.slice(0, MAX_ROWS));
        attachmentNote = ` [đã đính kèm: ${attachment.original_filename}]`;
        db.deleteChatAttachment(attachmentId); // dung xong xoa luon, khong luu lau dai
      }
    }

    db.addMessage(sessionId, 'user', message + attachmentNote);
    const history = db.getRecentHistory(sessionId);

    // Ngu canh du lieu that de AI co the tra loi/xuat file dua tren du lieu thuc te
    const todaySummary = db.getTodaySummary();
    const customerFiles = db.listCustomerFiles(sessionId).map((f) => ({
      khach_hang: f.customer_name,
      file: f.original_filename,
      tong_dong: f.totalRows,
      chua_gan: f.unassignedRows,
    }));
    // Chi tiet TUNG DONG da gan hom nay (kem SDT/Seri sim thuc te) - de AI tra loi duoc
    // cac cau hoi kieu "cho tôi chi tiết/tin nhắn Zalo của NV X cho KH Y", khong chi so luong.
    const MAX_DETAIL_ROWS = 1000;
    const todayDetails = db.getTodayDetailedAssignments()
      .slice(0, MAX_DETAIL_ROWS)
      .map((d) => {
        let rowData;
        try { rowData = JSON.parse(d.data_json); } catch { rowData = null; }
        const { phone, sim } = extractPhoneAndSim(rowData);
        return {
          khach_hang: d.customer_name,
          nhan_vien: d.employee_name,
          dong_so: d.row_index + 1,
          sdt: phone,
          seri_sim: sim,
          thoi_gian_gan: new Date(d.assigned_at).toLocaleString('vi-VN'),
        };
      });

    const systemPrompt = {
      role: 'system',
      content:
        'Bạn là trợ lý AI giúp xử lý công việc hàng ngày. Có 2 khu vực chính: ' +
        '(1) Xử lý dữ liệu khách hàng; (2) Phân chia từng dòng dữ liệu cho nhân viên. ' +
        'Ngoài ra, người dùng có thể đính kèm file bất kỳ ngay trong khung chat này để nhờ xử lý những việc KHÔNG liên quan gì đến khách hàng/nhân viên ' +
        '(ví dụ: tính lương, gộp danh sách, lọc số liệu bất kỳ...) — hãy xử lý bình thường theo đúng yêu cầu, không cần gắn nó vào khách hàng nào. ' +
        'Trả lời ngắn gọn, rõ ràng, bằng tiếng Việt.\n\n' +
        `${PHONE_SIM_RULE_NOTE}\n\n` +
        `Dữ liệu thực tế hiện có (chỉ dùng khi câu hỏi thực sự liên quan đến khách hàng/nhân viên):\n` +
        `Thống kê phân việc hôm nay (tổng hợp theo khách hàng + nhân viên): ${JSON.stringify(todaySummary)}\n` +
        `Danh sách hồ sơ khách hàng đã lưu: ${JSON.stringify(customerFiles)}\n` +
        `Chi tiết TỪNG DÒNG đã gán hôm nay, kèm SĐT/Seri sim thật (dùng để trả lời khi người dùng hỏi chi tiết, ` +
        `xin danh sách, hoặc xin soạn tin nhắn Zalo cho 1 nhân viên/khách hàng cụ thể — khi soạn tin nhắn Zalo, ` +
        `liệt kê MỖI DÒNG 1 CẶP theo đúng định dạng "SĐT - Seri sim"): ${JSON.stringify(todayDetails)}\n\n` +
        'Khi liệt kê danh sách kiểu "SĐT - Seri sim" để người dùng copy gửi Zalo: chỉ dùng văn bản thuần, ' +
        'mỗi cặp 1 dòng, TUYỆT ĐỐI không bọc trong dấu ```, không dùng markdown, không đánh số thứ tự, ' +
        'không thêm ký tự trang trí nào khác ngoài chính danh sách đó (để người dùng dán thẳng vào Zalo là dùng được ngay).\n\n' +
        'Nếu người dùng yêu cầu xuất/tải/gửi kết quả dưới dạng FILE EXCEL (dù là dữ liệu khách hàng hay dữ liệu đính kèm bất kỳ): ' +
        'hãy trả lời phần giải thích ngắn gọn trước, sau đó thêm ĐÚNG MỘT khối bắt đầu bằng dòng "###FILE_DATA###" và kết thúc bằng dòng "###END_FILE_DATA###", ' +
        'bên trong là JSON hợp lệ đại diện cho bảng cần xuất (mảng các object nếu có tiêu đề cột, hoặc mảng các mảng nếu không có tiêu đề — ' +
        'nếu xử lý từ file đính kèm thì giữ đúng định dạng có/không tiêu đề như dữ liệu đính kèm gốc). ' +
        'Nếu người dùng KHÔNG yêu cầu file, tuyệt đối không thêm khối này.',
    };

    const historyMessages = history.map((m) => ({ role: m.role, content: m.content }));
    // Neu co dinh kem, nhung du lieu vao NOI DUNG THAT su gui AI cho tin nhan cuoi cung (khong luu du lieu tho vao lich su)
    if (attachmentBlock && historyMessages.length) {
      const last = historyMessages[historyMessages.length - 1];
      if (last.role === 'user') last.content = message + attachmentBlock;
    }

    const chatMessages = [systemPrompt, ...historyMessages];

    const raw = await callOpenRouter(chatMessages);
    const { cleanText, rows } = extractFileBlock(raw);
    db.addMessage(sessionId, 'assistant', cleanText);

    let downloadUrl;
    if (rows && rows.length) {
      const outName = `chat_${Date.now()}.xlsx`;
      writeRowsToXlsxFile(normalizeRows(rows), path.join(OUTPUT_DIR, outName));
      downloadUrl = `/api/download/${outName}`;
    }

    res.json({ reply: cleanText, downloadUrl });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ================== DANH MUC KHACH HANG ==================
app.get('/api/customers', (req, res) => {
  res.json(db.listCustomersCatalog());
});
app.post('/api/customers', (req, res) => {
  try {
    res.json(db.addCustomer(req.body.name || ''));
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});
app.delete('/api/customers/:id', (req, res) => {
  db.deleteCustomer(req.params.id);
  res.json({ ok: true });
});

// ================== DANH MUC NHAN VIEN ==================
app.get('/api/employees', (req, res) => {
  res.json(db.listEmployeesCatalog());
});
app.post('/api/employees', (req, res) => {
  try {
    res.json(db.addEmployee(req.body.name || ''));
  } catch (err) {
    res.status(400).json({ error: err.message });
  }
});
app.delete('/api/employees/:id', (req, res) => {
  db.deleteEmployee(req.params.id);
  res.json({ ok: true });
});

// ================== UPLOAD FILE (TAM THOI - CHUA LUU) ==================
app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    const { sessionId, customerName } = req.body;
    const hasHeader = req.body.hasHeader !== 'false'; // mac dinh true
    if (!req.file || !sessionId || !customerName) {
      return res.status(400).json({ error: 'Thiếu file, sessionId hoặc customerName' });
    }
    const tempFileId = db.addTempFile({
      sessionId,
      customerName,
      originalFilename: req.file.originalname,
      storedPath: req.file.path,
      hasHeader,
    });

    const rows = normalizeRows(readRows(req.file.path, hasHeader));
    writeRowsToXlsxFile(rows, req.file.path); // ghi de bang du lieu da chuan hoa SDT/seri sim
    res.json({ tempFileId, customerName, rowCount: rows.length, preview: rows.slice(0, 5) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ================== DAN DU LIEU TAY (khong co file Excel) ==================
app.post('/api/upload-paste', (req, res) => {
  try {
    const { sessionId, customerName, pastedText } = req.body;
    const hasHeader = req.body.hasHeader !== false && req.body.hasHeader !== 'false';
    if (!sessionId || !customerName || !pastedText || !pastedText.trim()) {
      return res.status(400).json({ error: 'Thiếu sessionId, customerName hoặc dữ liệu dán vào' });
    }

    const rows = parsePastedText(pastedText, hasHeader);
    if (!rows.length) return res.status(400).json({ error: 'Không đọc được dữ liệu từ nội dung đã dán' });

    const originalFilename = `du-lieu-dan-tay-${Date.now()}.xlsx`;
    const storedPath = path.join(UPLOAD_DIR, originalFilename);
    writeRowsToXlsxFile(rows, storedPath);

    const tempFileId = db.addTempFile({ sessionId, customerName, originalFilename, storedPath, hasHeader });

    res.json({ tempFileId, customerName, rowCount: rows.length, preview: rows.slice(0, 5) });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ================== AI XU LY FILE TAM (CHUA LUU) ==================
app.post('/api/process', async (req, res) => {
  try {
    const { tempFileId, instruction } = req.body;
    const temp = db.getTempFile(tempFileId);
    if (!temp) return res.status(404).json({ error: 'Không tìm thấy file tạm' });

    const hasHeader = !!temp.has_header;
    const rows = readRows(temp.stored_path, hasHeader);
    const MAX_ROWS = 500;

    const formatNote = hasHeader
      ? 'Dữ liệu là mảng các object, mỗi object là 1 dòng với tên cột là khoá.'
      : 'Dữ liệu là mảng các mảng, mỗi mảng con là 1 dòng, KHÔNG có dòng tiêu đề (dữ liệu gốc không có tiêu đề, kết quả trả về CŨNG KHÔNG được thêm tiêu đề).';

    const messages = [
      {
        role: 'system',
        content:
          `Bạn nhận một bảng dữ liệu Excel dạng JSON và một yêu cầu xử lý. ${formatNote} ${PHONE_SIM_RULE_NOTE} ` +
          'Hãy thực hiện yêu cầu rồi CHỈ trả về JSON hợp lệ đúng CÙNG ĐỊNH DẠNG với dữ liệu đầu vào (mảng object nếu có tiêu đề, mảng các mảng nếu không có tiêu đề) đại diện cho bảng kết quả. ' +
          'Không thêm giải thích, không thêm markdown, không thêm dấu ```.',
      },
      {
        role: 'user',
        content: `Dữ liệu (khách hàng: ${temp.customer_name}):\n${JSON.stringify(
          rows.slice(0, MAX_ROWS)
        )}\n\nYêu cầu: ${instruction}`,
      },
    ];

    const raw = await callOpenRouter(messages);
    let resultRows;
    try {
      resultRows = normalizeRows(JSON.parse(raw.replace(/```json|```/g, '').trim()));
    } catch {
      return res.status(422).json({ error: 'AI không trả về JSON hợp lệ, thử lại với yêu cầu rõ ràng hơn.', raw });
    }

    const outName = `ketqua_${tempFileId}_${Date.now()}.xlsx`;
    const outPath = path.join(OUTPUT_DIR, outName);
    writeRowsToXlsxFile(resultRows, outPath);

    db.setTempProcessedPath(tempFileId, outPath);

    res.json({
      tempFileId,
      rowCount: resultRows.length,
      preview: resultRows.slice(0, 5),
      downloadUrl: `/api/download/${outName}`,
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ================== LUU FILE TAM VAO HO SO KHACH HANG (VO CHON "CO") ==================
// Neu chua qua AI xu ly (khong co processed_path) thi luu du lieu goc.
app.post('/api/save', (req, res) => {
  try {
    const { tempFileId } = req.body;
    const temp = db.getTempFile(tempFileId);
    if (!temp) return res.status(404).json({ error: 'Không tìm thấy file tạm' });

    const hasHeader = !!temp.has_header;
    const sourcePath = temp.processed_path || temp.stored_path;
    const rows = normalizeRows(readRows(sourcePath, hasHeader));

    const customerFileId = db.saveCustomerFile({
      sessionId: temp.session_id,
      customerName: temp.customer_name,
      originalFilename: temp.original_filename,
      storedPath: temp.stored_path,
      processedPath: temp.processed_path,
      rows,
      hasHeader,
    });

    db.deleteTempFile(tempFileId); // da luu chinh thuc, khong can giu ban tam nua

    res.json({ customerFileId, rowCount: rows.length });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Neu vo chon "KHONG luu" - chi xoa file tam de don rac som
app.post('/api/discard', (req, res) => {
  try {
    const { tempFileId } = req.body;
    const temp = db.getTempFile(tempFileId);
    if (temp) {
      [temp.stored_path, temp.processed_path].forEach((p) => {
        if (p && fs.existsSync(p)) fs.unlinkSync(p);
      });
      db.deleteTempFile(tempFileId);
    }
    res.json({ ok: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/download/:filename', (req, res) => {
  const filePath = path.join(OUTPUT_DIR, req.params.filename);
  if (!fs.existsSync(filePath)) return res.status(404).send('Không tìm thấy file');
  res.download(filePath);
});

// ================== DANH SACH FILE KHACH HANG DA LUU (kem trang thai phan cong) ==================
app.get('/api/files/:sessionId', (req, res) => {
  res.json(db.listCustomerFiles(req.params.sessionId));
});

app.get('/api/files/:customerFileId/status', (req, res) => {
  res.json(db.getAssignmentStatus(req.params.customerFileId));
});

// ================== PHAN VIEC THEO TUNG DONG DU LIEU ==================
// Gan N dong CHUA gan cua 1 file khach hang cho 1 nhan vien.
// Vi du: KH A co 100 dong -> goi 3 lan voi employeeName khac nhau va rowCount khac nhau
// se tu dong chia lan luot 100 dong do cho 3 nhan vien.
app.post('/api/assign-rows', (req, res) => {
  try {
    const { customerFileId, employeeName, rowCount } = req.body;
    if (!customerFileId || !employeeName || !rowCount) {
      return res.status(400).json({ error: 'Thiếu customerFileId, employeeName hoặc rowCount' });
    }
    const assignedRowsData = db.assignRowsToEmployee(customerFileId, employeeName, Number(rowCount));
    const status = db.getAssignmentStatus(customerFileId);
    const smsText = assignedRowsData.length ? buildAssignmentSms(employeeName, assignedRowsData) : '';
    res.json({ assigned: assignedRowsData.length, status, smsText });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Thu hoi (huy gan) toan bo du lieu da giao cho 1 nhan vien cu the trong 1 ho so khach hang
app.post('/api/revoke-assignment', (req, res) => {
  try {
    const { customerFileId, employeeName } = req.body;
    if (!customerFileId || !employeeName) {
      return res.status(400).json({ error: 'Thiếu customerFileId hoặc employeeName' });
    }
    const revoked = db.revokeAssignmentForEmployee(customerFileId, employeeName);
    const status = db.getAssignmentStatus(customerFileId);
    res.json({ revoked, status });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// Thu hoi TOAN BO phan cong cua 1 ho so khach hang (dat lai het ve trang thai chua gan)
app.post('/api/revoke-all', (req, res) => {
  try {
    const { customerFileId } = req.body;
    if (!customerFileId) return res.status(400).json({ error: 'Thiếu customerFileId' });
    const revoked = db.revokeAllAssignments(customerFileId);
    const status = db.getAssignmentStatus(customerFileId);
    res.json({ revoked, status });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// ================== THONG KE CUOI NGAY ==================
app.get('/api/summary', (req, res) => {
  res.json(db.getTodaySummary());
});

app.get('/api/summary/export', (req, res) => {
  const rows = db.getTodaySummary().map((r) => ({
    'Khách hàng': r.customer_name,
    'Nhân viên nhận': r.employee_name,
    'Số dòng dữ liệu': r.row_count,
    'Gửi lần đầu lúc': new Date(r.first_time).toLocaleString('vi-VN'),
    'Gửi lần cuối lúc': new Date(r.last_time).toLocaleString('vi-VN'),
  }));
  const sheet = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, 'ThongKe');
  const outName = `thongke_${Date.now()}.xlsx`;
  const outPath = path.join(OUTPUT_DIR, outName);
  XLSX.writeFile(wb, outPath);
  res.download(outPath, 'thong_ke_cuoi_ngay.xlsx');
});

// Chuyen 1 dong du lieu (co the la object hoac mang) + thong tin gan viec thanh 1 dong phang de hien thi/xuat file
function flattenDetailRow(d) {
  let rowData;
  try { rowData = JSON.parse(d.data_json); } catch { rowData = d.data_json; }

  const base = {
    'Khách hàng': d.customer_name,
    'Nhân viên': d.employee_name,
    'Dòng số': d.row_index + 1,
    'Thời gian gán': new Date(d.assigned_at).toLocaleString('vi-VN'),
  };

  if (Array.isArray(rowData)) {
    rowData.forEach((v, i) => { base[`Cột ${i + 1}`] = v; });
  } else if (rowData && typeof rowData === 'object') {
    Object.entries(rowData).forEach(([k, v]) => { base[k] = v; });
  } else if (rowData !== undefined) {
    base['Nội dung'] = rowData;
  }
  return base;
}

// ================== THONG KE CHI TIET TUNG DONG (xem trong app) ==================
app.get('/api/summary/detail', (req, res) => {
  const details = db.getTodayDetailedAssignments();
  res.json(details.map(flattenDetailRow));
});

// ================== THONG KE CHI TIET TUNG DONG (xuat Excel) ==================
app.get('/api/summary/detail/export', (req, res) => {
  const rows = db.getTodayDetailedAssignments().map(flattenDetailRow);
  const sheet = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, 'ChiTiet');
  const outName = `chitiet_${Date.now()}.xlsx`;
  const outPath = path.join(OUTPUT_DIR, outName);
  XLSX.writeFile(wb, outPath);
  res.download(outPath, 'chi_tiet_phan_viec_hom_nay.xlsx');
});

// ================== HOI AI BANG NGON NGU TU NHIEN VE THONG KE ==================
app.post('/api/summary/ask', async (req, res) => {
  try {
    const { question } = req.body;
    if (!question) return res.status(400).json({ error: 'Thiếu question' });

    const data = db.getTodaySummary().map((r) => ({
      khach_hang: r.customer_name,
      nhan_vien: r.employee_name,
      so_dong: r.row_count,
      gui_lan_dau: new Date(r.first_time).toLocaleString('vi-VN'),
      gui_lan_cuoi: new Date(r.last_time).toLocaleString('vi-VN'),
    }));

    const messages = [
      {
        role: 'system',
        content:
          'Bạn là trợ lý phân tích dữ liệu công việc. Dưới đây là bảng thống kê hôm nay ' +
          '(mỗi dòng là: khách hàng nào đã được phân bao nhiêu dòng dữ liệu cho nhân viên nào, và thời gian). ' +
          'Hãy trả lời câu hỏi của người dùng dựa hoàn toàn vào dữ liệu này, bằng tiếng Việt, ngắn gọn, dễ hiểu. ' +
          'Nếu dữ liệu không đủ để trả lời, hãy nói rõ là chưa có dữ liệu.',
      },
      { role: 'user', content: `Dữ liệu thống kê hôm nay:\n${JSON.stringify(data)}\n\nCâu hỏi: ${question}` },
    ];

    const answer = await callOpenRouter(messages);
    res.json({ answer });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server đang chạy tại http://localhost:${PORT}`));
