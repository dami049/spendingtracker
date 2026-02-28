/* ============================================================
   Spending Tracker — app.js
   ============================================================ */

// ============================================================
// State
// ============================================================
const state = {
  transactions: [],
  income: {},            // { "2024-01": 3000 }
  categoryOverrides: {}, // { txId: "category" }
  customCategories: [],  // ["dining", ...]
  categoryColors: {},    // { category: "#hex" }
  activeMonth: null,     // "YYYY-MM"
  activeTab: 'upload',
  activeYear: null,      // "2024"
  pendingGenericRows: null,  // raw rows for generic mapper
  pendingGenericHeaders: null,
  donutChart: null,
};

// ============================================================
// Constants
// ============================================================
const LS_TRANSACTIONS  = 'monzo_transactions';
const LS_INCOME        = 'monzo_income';
const LS_OVERRIDES     = 'monzo_category_overrides';
const LS_CUSTOM_CATS   = 'monzo_custom_categories';
const LS_CAT_COLORS    = 'monzo_category_colors';
const LS_THEME         = 'monzo_theme';

const COLOR_PALETTE = [
  '#6366f1','#ec4899','#f59e0b','#10b981','#3b82f6',
  '#8b5cf6','#06b6d4','#ef4444','#84cc16','#f97316',
  '#14b8a6','#a855f7','#e11d48','#64748b','#0ea5e9',
  '#d946ef','#22c55e','#fb923c','#38bdf8','#c084fc',
];

const MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

// Auto-categorisation rules for non-Monzo banks (keyword → category)
const AUTO_CAT_RULES = [
  { patterns: /tesco|sainsbury|waitrose|aldi|lidl|asda|morrisons|co-op|marks.{0,5}spencer|m&s food/i, category: 'groceries' },
  { patterns: /tfl|trainline|national rail|uber|bolt|heathrow|gatwick|parking|petrol|shell|bp |esso|fuel/i, category: 'transport' },
  { patterns: /netflix|spotify|apple.{0,6}tv|prime video|disney|cinema|odeon|vue|cineworld|sky cinema/i, category: 'entertainment' },
  { patterns: /amazon|ebay|asos|argos|currys|john lewis|next|zara|h&m|primark/i, category: 'shopping' },
  { patterns: /sky|bt |virgin media|broadband|council tax|water|severn|thames|electric|gas|bulb|eon|npower/i, category: 'bills' },
  { patterns: /salary|payroll|wages|bacs credit|direct credit|employer/i, category: 'income' },
  { patterns: /mcdonald|kfc|burger king|subway|nando|pizza|domino|just eat|deliveroo|uber eats|takeaway|restaurant|cafe|coffee|starbucks|costa|greggs/i, category: 'eating out' },
  { patterns: /gym|fitness|sports|swimming|pilates|yoga/i, category: 'fitness' },
  { patterns: /boots|pharmacy|chemist|superdrug|hospital|gp |nhs|dentist|optician/i, category: 'health' },
  { patterns: /hotel|airbnb|booking\.com|holiday|travel|flight|easyjet|ryanair|ba |british airways/i, category: 'travel' },
  { patterns: /school|college|university|tuition|course|udemy|coursera/i, category: 'education' },
];

function autoCategory(name) {
  for (const rule of AUTO_CAT_RULES) {
    if (rule.patterns.test(name)) return rule.category;
  }
  return 'other';
}

// ============================================================
// Storage
// ============================================================
function loadFromStorage() {
  try {
    const tx = localStorage.getItem(LS_TRANSACTIONS);
    state.transactions = tx ? JSON.parse(tx).map(deserializeTx) : [];
    state.income = JSON.parse(localStorage.getItem(LS_INCOME) || '{}');
    state.categoryOverrides = JSON.parse(localStorage.getItem(LS_OVERRIDES) || '{}');
    state.customCategories = JSON.parse(localStorage.getItem(LS_CUSTOM_CATS) || '[]');
    state.categoryColors = JSON.parse(localStorage.getItem(LS_CAT_COLORS) || '{}');
  } catch (e) {
    console.error('Storage load error', e);
  }
}

function saveToStorage() {
  localStorage.setItem(LS_TRANSACTIONS, JSON.stringify(state.transactions.map(serializeTx)));
  localStorage.setItem(LS_INCOME, JSON.stringify(state.income));
  localStorage.setItem(LS_OVERRIDES, JSON.stringify(state.categoryOverrides));
  localStorage.setItem(LS_CUSTOM_CATS, JSON.stringify(state.customCategories));
  localStorage.setItem(LS_CAT_COLORS, JSON.stringify(state.categoryColors));
}

function serializeTx(tx) {
  return { ...tx, date: tx.date instanceof Date ? tx.date.toISOString() : tx.date };
}
function deserializeTx(tx) {
  return { ...tx, date: new Date(tx.date) };
}

// ============================================================
// CSV Parsing
// ============================================================

/**
 * Split a line respecting quoted fields (CSV) or tabs (TSV).
 * For TSV we never have quoted fields.
 */
function splitLine(line, delimiter) {
  if (delimiter === '\t') return line.split('\t');
  const result = [];
  let cur = '';
  let inQuote = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') {
      if (inQuote && line[i+1] === '"') { cur += '"'; i++; }
      else inQuote = !inQuote;
    } else if (c === delimiter && !inQuote) {
      result.push(cur.trim()); cur = '';
    } else {
      cur += c;
    }
  }
  result.push(cur.trim());
  return result;
}

function parseRawCSV(text) {
  // Detect delimiter: if many tabs on first line → TSV
  const firstLine = text.split('\n')[0];
  const tabs = (firstLine.match(/\t/g) || []).length;
  const commas = (firstLine.match(/,/g) || []).length;
  const delimiter = tabs > commas ? '\t' : ',';

  const lines = text.split('\n').map(l => l.replace(/\r$/, ''));
  const headers = splitLine(lines[0], delimiter).map(h => h.trim().replace(/^"|"$/g, ''));
  const rows = lines.slice(1)
    .filter(l => l.trim().length > 0)
    .map(l => {
      const cols = splitLine(l, delimiter);
      const obj = {};
      headers.forEach((h, i) => { obj[h] = (cols[i] || '').trim().replace(/^"|"$/g, ''); });
      return obj;
    });
  return { headers, rows, delimiter };
}

function detectBankFormat(headers) {
  const h = headers.map(x => x.toLowerCase());
  const has = (...keys) => keys.every(k => h.some(x => x.includes(k)));

  // Monzo: tab-separated, has Transaction ID + Emoji + Money Out
  if (has('transaction id') && has('emoji') && has('money out')) return 'monzo';
  // HSBC: Date, Description, Amount, Balance (4 cols typically)
  if (has('description') && has('amount') && has('balance') && !has('debit')) return 'hsbc';
  // Halifax: Debit Amount + Credit Amount
  if (has('debit amount') || has('debit') && has('credit')) return 'halifax';
  // Amex: Date, Reference, Amount (no balance, no description variant)
  if (has('reference') && has('amount') && !has('balance')) return 'amex';

  return 'generic';
}

// Parse DD/MM/YYYY date string
function parseDDMMYYYY(str) {
  if (!str) return null;
  const parts = str.split('/');
  if (parts.length === 3) {
    const [d, m, y] = parts;
    const date = new Date(+y, +m - 1, +d);
    if (!isNaN(date)) return date;
  }
  // Fallback: try native
  const d = new Date(str);
  return isNaN(d) ? null : d;
}

// Parse YYYY-MM-DD
function parseYMD(str) {
  if (!str) return null;
  const d = new Date(str);
  return isNaN(d) ? null : d;
}

function parseAmount(str) {
  if (!str) return 0;
  return parseFloat(str.replace(/[£,\s]/g, '')) || 0;
}

let _txCounter = 0;
function genId(prefix) { return `${prefix}_${Date.now()}_${++_txCounter}`; }

// ---- Bank parsers ----

function parseMonzoCSV(rows) {
  return rows
    .filter(r => r['Transaction ID'] || r['Name'] || r['Amount'])
    .map(r => {
      const amount = parseAmount(r['Amount']);
      const date = parseDDMMYYYY(r['Date']);
      return {
        id: r['Transaction ID'] || genId('monzo'),
        date,
        name: r['Name'] || r['Description'] || '',
        category: (r['Category'] || 'other').toLowerCase(),
        monzoCategory: (r['Category'] || 'other').toLowerCase(),
        amount,
        type: amount >= 0 ? 'income' : 'expense',
        bank: 'monzo',
      };
    })
    .filter(t => t.date !== null);
}

function parseHSBCCSV(rows) {
  return rows.map(r => {
    const dateStr = r['Date'] || r['date'] || '';
    const name = r['Description'] || r['Payee'] || '';
    const amount = parseAmount(r['Amount'] || r['amount'] || '0');
    // HSBC negative = debit (expense), positive = credit
    const date = parseDDMMYYYY(dateStr) || parseYMD(dateStr);
    return {
      id: genId('hsbc'),
      date,
      name,
      category: autoCategory(name),
      amount,
      type: amount >= 0 ? 'income' : 'expense',
      bank: 'hsbc',
    };
  }).filter(t => t.date !== null);
}

function parseHalifaxCSV(rows) {
  return rows.map(r => {
    const dateStr = r['Date'] || '';
    const name = r['Transaction Description'] || r['Description'] || '';
    const debit = parseAmount(r['Debit Amount'] || r['Debit'] || '0');
    const credit = parseAmount(r['Credit Amount'] || r['Credit'] || '0');
    const amount = credit > 0 ? credit : -Math.abs(debit);
    const date = parseDDMMYYYY(dateStr) || parseYMD(dateStr);
    return {
      id: genId('halifax'),
      date,
      name,
      category: autoCategory(name),
      amount,
      type: amount >= 0 ? 'income' : 'expense',
      bank: 'halifax',
    };
  }).filter(t => t.date !== null);
}

function parseAmexCSV(rows) {
  return rows.map(r => {
    const dateStr = r['Date'] || '';
    const name = r['Description'] || r['Reference'] || '';
    // Amex: positive = charge (expense), negative = payment/refund
    const raw = parseAmount(r['Amount'] || '0');
    const amount = -raw; // flip so negative = expense
    const date = parseDDMMYYYY(dateStr) || parseYMD(dateStr);
    return {
      id: genId('amex'),
      date,
      name,
      category: autoCategory(name),
      amount,
      type: amount >= 0 ? 'income' : 'expense',
      bank: 'amex',
    };
  }).filter(t => t.date !== null);
}

function parseGenericCSV(rows, mapping) {
  // mapping: { date: colName, name: colName, amount: colName, category?: colName }
  return rows.map(r => {
    const dateStr = r[mapping.date] || '';
    const name = r[mapping.name] || '';
    const amount = parseAmount(r[mapping.amount] || '0');
    const category = mapping.category ? (r[mapping.category] || 'other').toLowerCase() : autoCategory(name);
    const date = parseDDMMYYYY(dateStr) || parseYMD(dateStr);
    return {
      id: genId('generic'),
      date,
      name,
      category,
      amount,
      type: amount >= 0 ? 'income' : 'expense',
      bank: 'generic',
    };
  }).filter(t => t.date !== null);
}

function parseCSV(text) {
  const { headers, rows } = parseRawCSV(text);
  const format = detectBankFormat(headers);

  if (format === 'monzo')   return { transactions: parseMonzoCSV(rows),  bank: 'monzo',   format };
  if (format === 'hsbc')    return { transactions: parseHSBCCSV(rows),   bank: 'hsbc',    format };
  if (format === 'halifax') return { transactions: parseHalifaxCSV(rows),bank: 'halifax', format };
  if (format === 'amex')    return { transactions: parseAmexCSV(rows),   bank: 'amex',    format };

  // Generic: return rows + headers for mapper
  return { transactions: null, bank: 'generic', format, rows, headers };
}

// ============================================================
// PDF Parsing
// ============================================================
const PDFJS_WORKER_URL = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
let _pdfJsReady = false;

async function initPDFJS() {
  if (typeof pdfjsLib === 'undefined' || _pdfJsReady) return;
  try {
    const resp = await fetch(PDFJS_WORKER_URL);
    const blob = await resp.blob();
    pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(blob);
  } catch (_e) {
    // Fallback: point at CDN directly (works on https:// pages)
    pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_URL;
  }
  _pdfJsReady = true;
}

async function extractPDFLines(pdfDoc) {
  const allLines = [];
  for (let pageNum = 1; pageNum <= pdfDoc.numPages; pageNum++) {
    const page = await pdfDoc.getPage(pageNum);
    const viewport = page.getViewport({ scale: 1 });
    const content = await page.getTextContent();

    const items = content.items
      .filter(item => item.str && item.str.trim())
      .map(item => ({
        text: item.str,
        x: item.transform[4],
        y: viewport.height - item.transform[5], // flip so top = 0
      }));

    if (!items.length) continue;

    // Group items within ±3 y-units into the same visual line
    const groups = [];
    for (const item of items) {
      const g = groups.find(g => Math.abs(g.y - item.y) <= 3);
      if (g) g.items.push(item);
      else groups.push({ y: item.y, items: [item] });
    }

    groups.sort((a, b) => a.y - b.y);
    for (const g of groups) {
      g.items.sort((a, b) => a.x - b.x);
      const text = g.items.map(i => i.text).join(' ').trim();
      if (text) allLines.push({ text, items: g.items, pageNum, y: g.y });
    }
  }
  return allLines;
}

function detectBankFromPDFText(text) {
  const t = text.toLowerCase();
  if (t.includes('hsbc')) return 'hsbc';
  if (t.includes('halifax')) return 'halifax';
  if (t.includes('american express') || t.includes('amex')) return 'amex';
  if (t.includes('barclays')) return 'barclays';
  if (t.includes('lloyds')) return 'lloyds';
  if (t.includes('natwest')) return 'natwest';
  if (t.includes('nationwide')) return 'nationwide';
  if (t.includes('santander')) return 'santander';
  return 'generic';
}

const PDF_MONTH_MAP = {
  jan:1, feb:2, mar:3, apr:4, may:5, jun:6,
  jul:7, aug:8, sep:9, oct:10, nov:11, dec:12,
};

function parseDateAtStart(text) {
  // DD/MM/YYYY or DD/MM/YY
  let m = text.match(/^(\d{1,2})\/(\d{2})\/(\d{2,4})/);
  if (m) {
    let y = parseInt(m[3], 10);
    if (y < 100) y += 2000;
    const d = new Date(y, parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    if (!isNaN(d)) return { date: d, end: m[0].length };
  }
  // DD Mon YY or DD Mon YYYY
  m = text.match(/^(\d{1,2})\s+([A-Za-z]{3})\s+(\d{2,4})/);
  if (m) {
    const mo = PDF_MONTH_MAP[m[2].toLowerCase()];
    if (mo) {
      let y = parseInt(m[3], 10);
      if (y < 100) y += 2000;
      const d = new Date(y, mo - 1, parseInt(m[1], 10));
      if (!isNaN(d)) return { date: d, end: m[0].length };
    }
  }
  return null;
}

function extractAmountsFromLine(text) {
  const re = /£?([\d,]+\.\d{2})\s*(CR|DR)?/gi;
  const results = [];
  let match;
  while ((match = re.exec(text)) !== null) {
    results.push({
      value: parseFloat(match[1].replace(/,/g, '')),
      crdr: match[2] ? match[2].toUpperCase() : null,
      index: match.index,
    });
  }
  return results;
}

const SKIP_LINE_RE = /^(page\s+\d|sort\s+code|account\s+(number|no\.?)|balance\s+brought\s+forward|opening\s+balance|closing\s+balance|total\s+(debit|credit)|statement\s+period|date\s+description|transaction\s+date|previous\s+balance|brought\s+forward|available\s+balance)/i;

const INCOME_KEYWORDS_RE = /salary|wages|bacs\s+credit|direct\s+credit|employer|refund|cashback|interest\s+paid|credit\s+interest|dividend|payment\s+received|transfer\s+in/i;

function applyBankSign(value, bank, description) {
  if (INCOME_KEYWORDS_RE.test(description)) return Math.abs(value);
  if (bank === 'amex') return -Math.abs(value);
  return -Math.abs(value); // default: expense
}

function finalisePDFTx(pending, bank, prevBalance) {
  const { date, description, amounts, rawText } = pending;
  if (!amounts.length) return null;

  const primaryAmt = amounts[0];
  let amount = null;

  // 1. CR/DR suffix on any amount
  const crdrAmt = amounts.find(a => a.crdr);
  if (crdrAmt) {
    amount = crdrAmt.crdr === 'CR' ? Math.abs(crdrAmt.value) : -Math.abs(crdrAmt.value);
  }

  // 2. Balance differential: last amount = running balance
  if (amount === null && amounts.length >= 2 && prevBalance !== null) {
    const newBal = amounts[amounts.length - 1].value;
    const delta = newBal - prevBalance;
    if (Math.abs(Math.abs(delta) - primaryAmt.value) < 0.02) {
      amount = delta;
    }
  }

  // 3. Explicit minus sign in raw text before the amount
  if (amount === null && /[-\u2212]\s*£?\d/.test(rawText)) {
    amount = -Math.abs(primaryAmt.value);
  }

  // 4. Heuristics: income keywords, bank-specific rules, default to expense
  if (amount === null) {
    amount = applyBankSign(primaryAmt.value, bank, description);
  }

  const name = description.replace(/\s+/g, ' ').trim() || 'Unknown';
  return {
    id: genId(`pdf_${bank}`),
    date,
    name,
    category: autoCategory(name),
    amount,
    type: amount >= 0 ? 'income' : 'expense',
    bank,
  };
}

function parsePDFTransactions(lines, bank) {
  const transactions = [];
  let pending = null;
  let prevBalance = null;

  const flush = () => {
    if (!pending) return;
    const tx = finalisePDFTx(pending, bank, prevBalance);
    if (tx) transactions.push(tx);
    if (pending.amounts.length >= 2) {
      prevBalance = pending.amounts[pending.amounts.length - 1].value;
    }
    pending = null;
  };

  for (const { text } of lines) {
    const t = text.trim();
    if (!t || SKIP_LINE_RE.test(t)) continue;

    const dateResult = parseDateAtStart(t);
    const amounts = extractAmountsFromLine(text);

    if (dateResult && amounts.length > 0) {
      flush();
      const desc = text.slice(dateResult.end, amounts[0].index).trim();
      pending = { date: dateResult.date, description: desc, amounts, rawText: text };
    } else if (dateResult) {
      // Date with no amounts = header date-range line — skip
    } else if (!dateResult && pending && amounts.length === 0 && t.length > 1) {
      // Continuation line: append to description
      pending.description += ' ' + t;
    }
  }

  flush();
  return transactions;
}

async function parsePDF(file) {
  if (!_pdfJsReady) await initPDFJS();
  if (typeof pdfjsLib === 'undefined') {
    throw new Error('PDF.js library not loaded. Check your internet connection.');
  }

  const buf = await file.arrayBuffer();
  let pdfDoc;
  try {
    pdfDoc = await pdfjsLib.getDocument({ data: buf }).promise;
  } catch (e) {
    if (e.name === 'PasswordException') {
      throw new Error('This PDF is password protected. Please unlock it first.');
    }
    throw new Error(`Could not open PDF: ${e.message}`);
  }

  const lines = await extractPDFLines(pdfDoc);
  const allText = lines.map(l => l.text).join(' ');

  if (!allText.trim()) {
    throw new Error("No text found in this PDF. It may be a scanned image — please use a text-based PDF or your bank's CSV export.");
  }

  const bank = detectBankFromPDFText(allText);
  const transactions = parsePDFTransactions(lines, bank);

  if (!transactions.length) {
    throw new Error("Could not extract any transactions from this PDF. Try your bank's CSV export instead.");
  }

  return { transactions, bank };
}

// ============================================================
// Deduplication
// ============================================================
function txKey(tx) {
  const d = tx.date instanceof Date ? tx.date.toISOString().slice(0,10) : String(tx.date).slice(0,10);
  return `${d}|${tx.name}|${tx.amount}`;
}

function mergeTransactions(existing, newBatch) {
  const seenIds = new Set(existing.map(t => t.id));
  const seenKeys = new Set(existing.map(txKey));
  const toAdd = newBatch.filter(t => {
    if (seenIds.has(t.id)) return false;
    const k = txKey(t);
    if (seenKeys.has(k)) return false;
    seenIds.add(t.id);
    seenKeys.add(k);
    return true;
  });
  return [...existing, ...toAdd];
}

// ============================================================
// Data Helpers
// ============================================================
function monthKey(date) {
  if (!(date instanceof Date)) date = new Date(date);
  const m = String(date.getMonth() + 1).padStart(2, '0');
  return `${date.getFullYear()}-${m}`;
}

function getMonthsPresent(transactions) {
  const keys = new Set(transactions.map(t => monthKey(t.date)));
  return Array.from(keys).sort();
}

function getYearsPresent(transactions) {
  const keys = new Set(transactions.map(t => String(t.date.getFullYear())));
  return Array.from(keys).sort();
}

function getTransactionsForMonth(month) {
  return state.transactions.filter(t => monthKey(t.date) === month);
}

function getEffectiveCategory(tx) {
  return state.categoryOverrides[tx.id] || tx.category || 'other';
}

function getCategoryTotals(transactions) {
  const totals = {};
  for (const tx of transactions) {
    if (tx.type !== 'expense') continue;
    const cat = getEffectiveCategory(tx);
    totals[cat] = (totals[cat] || 0) + Math.abs(tx.amount);
  }
  return totals;
}

function getYearlyMatrix(year) {
  // { category: { "YYYY-MM": amount } }
  const yearTx = state.transactions.filter(t => String(t.date.getFullYear()) === String(year));
  const matrix = {};
  for (const tx of yearTx) {
    if (tx.type !== 'expense') continue;
    const cat = getEffectiveCategory(tx);
    const mk = monthKey(tx.date);
    if (!matrix[cat]) matrix[cat] = {};
    matrix[cat][mk] = (matrix[cat][mk] || 0) + Math.abs(tx.amount);
  }
  return matrix;
}

function getAllCategories() {
  const fromTx = new Set(state.transactions.map(t => getEffectiveCategory(t)));
  return Array.from(new Set([...fromTx, ...state.customCategories])).sort();
}

// ============================================================
// Category Colors
// ============================================================
function getCategoryColor(cat) {
  if (!state.categoryColors[cat]) {
    const idx = Object.keys(state.categoryColors).length % COLOR_PALETTE.length;
    state.categoryColors[cat] = COLOR_PALETTE[idx];
    saveToStorage();
  }
  return state.categoryColors[cat];
}

// ============================================================
// Format helpers
// ============================================================
function fmtGBP(amount) {
  const abs = Math.abs(amount);
  const fmt = abs.toLocaleString('en-GB', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return `£${fmt}`;
}

function fmtDate(date) {
  if (!(date instanceof Date)) return '';
  return `${String(date.getDate()).padStart(2,'0')}/${String(date.getMonth()+1).padStart(2,'0')}`;
}

// ============================================================
// Renderers
// ============================================================

// --- Upload Tab ---
function renderUploadTab() {
  const months = getMonthsPresent(state.transactions);
  const overview = document.getElementById('dataOverview');
  const stats = document.getElementById('overviewStats');
  const badges = document.getElementById('bankBadges');

  if (state.transactions.length === 0) {
    overview.classList.add('hidden');
    return;
  }

  overview.classList.remove('hidden');
  document.getElementById('overviewTitle').textContent = 'Loaded data';

  // Stats
  const banks = [...new Set(state.transactions.map(t => t.bank))];
  const dateMin = months[0];
  const dateMax = months[months.length - 1];
  const cover = months.length === 1
    ? formatMonthLabel(dateMin)
    : `${formatMonthLabel(dateMin)} – ${formatMonthLabel(dateMax)}`;

  stats.innerHTML = `
    <div class="stat-item"><strong>${state.transactions.length}</strong> transactions</div>
    <div class="stat-item">covering <strong>${cover}</strong></div>
    <div class="stat-item"><strong>${months.length}</strong> month${months.length !== 1 ? 's' : ''}</div>
  `;

  badges.innerHTML = banks.map(b => `<span class="bank-badge ${b}">${b}</span>`).join('');
}

// --- Monthly Tab ---
function renderMonthlyTab() {
  populateMonthSelect();
  if (!state.activeMonth) {
    const months = getMonthsPresent(state.transactions);
    state.activeMonth = months[months.length - 1] || null;
  }

  const monthSel = document.getElementById('monthSelect');
  if (state.activeMonth) monthSel.value = state.activeMonth;

  const income = state.income[state.activeMonth] || 0;
  document.getElementById('incomeInput').value = income || '';

  renderMonthSummary();
  renderDonutChart();
  renderTransactionList();
}

function populateMonthSelect() {
  const months = getMonthsPresent(state.transactions);
  const sel = document.getElementById('monthSelect');
  sel.innerHTML = months.length === 0
    ? '<option value="">No data loaded</option>'
    : months.map(m => `<option value="${m}">${formatMonthLabel(m)}</option>`).join('');
}

function formatMonthLabel(ym) {
  if (!ym) return '';
  const [y, m] = ym.split('-');
  return `${MONTH_NAMES[+m - 1]} ${y}`;
}

function renderMonthSummary() {
  const txs = getTransactionsForMonth(state.activeMonth || '');
  const spent = txs.filter(t => t.type === 'expense').reduce((s, t) => s + Math.abs(t.amount), 0);
  const income = state.income[state.activeMonth] || 0;
  const savings = income - spent;
  const rate = income > 0 ? ((savings / income) * 100).toFixed(1) : 0;

  document.getElementById('cardIncome').textContent = fmtGBP(income);
  document.getElementById('cardSpent').textContent = fmtGBP(spent);

  const savingsEl = document.getElementById('cardSavings');
  savingsEl.textContent = fmtGBP(savings);
  savingsEl.className = 'card-value ' + (savings >= 0 ? 'positive' : 'negative');

  const rateEl = document.getElementById('cardRate');
  rateEl.textContent = `${rate}%`;
  rateEl.className = 'card-value ' + (rate >= 0 ? 'positive' : 'negative');
}

function renderDonutChart() {
  const txs = getTransactionsForMonth(state.activeMonth || '');
  const totals = getCategoryTotals(txs);
  const entries = Object.entries(totals).sort((a, b) => b[1] - a[1]);
  const cats = entries.map(e => e[0]);
  const amounts = entries.map(e => e[1]);
  const colors = cats.map(c => getCategoryColor(c));

  const ctx = document.getElementById('donutChart').getContext('2d');

  if (state.donutChart) {
    state.donutChart.destroy();
    state.donutChart = null;
  }

  if (entries.length === 0) {
    document.getElementById('chartLegend').innerHTML = '<p class="empty-state text-muted">No expenses this month</p>';
    return;
  }

  state.donutChart = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: cats,
      datasets: [{
        data: amounts,
        backgroundColor: colors,
        borderWidth: 0,
        hoverOffset: 6,
      }],
    },
    options: {
      responsive: true,
      cutout: '65%',
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => ` ${fmtGBP(ctx.parsed)}`,
          },
        },
      },
    },
  });

  const total = amounts.reduce((a, b) => a + b, 0);
  const legend = document.getElementById('chartLegend');
  legend.innerHTML = entries.map(([cat, amt]) => `
    <div class="legend-item">
      <span class="legend-dot" style="background:${getCategoryColor(cat)}"></span>
      <span class="legend-label">${cat}</span>
      <span class="legend-amount">${fmtGBP(amt)} <span class="text-muted">(${((amt/total)*100).toFixed(0)}%)</span></span>
    </div>
  `).join('');
}

function renderTransactionList() {
  const container = document.getElementById('transactionList');
  const txs = getTransactionsForMonth(state.activeMonth || '');

  if (txs.length === 0) {
    container.innerHTML = '<div class="empty-state">No transactions for this month.<br>Upload a CSV file to get started.</div>';
    return;
  }

  // Group by category
  const groups = {};
  for (const tx of txs) {
    const cat = getEffectiveCategory(tx);
    if (!groups[cat]) groups[cat] = [];
    groups[cat].push(tx);
  }

  // Sort groups by total spend descending
  const sortedCats = Object.keys(groups).sort((a, b) => {
    const sumA = groups[a].reduce((s, t) => s + Math.abs(t.amount), 0);
    const sumB = groups[b].reduce((s, t) => s + Math.abs(t.amount), 0);
    return sumB - sumA;
  });

  const allCats = getAllCategories();
  const catOptions = allCats.map(c => `<option value="${c}">${c}</option>`).join('');

  container.innerHTML = sortedCats.map(cat => {
    const txList = groups[cat];
    const total = txList.reduce((s, t) => s + Math.abs(t.amount), 0);
    const color = getCategoryColor(cat);

    const rows = txList
      .sort((a, b) => b.date - a.date)
      .map(tx => {
        const isExpense = tx.type === 'expense';
        const amtClass = isExpense ? 'expense' : 'income';
        const amtStr = (isExpense ? '−' : '+') + fmtGBP(tx.amount);
        const effectiveCat = getEffectiveCategory(tx);
        return `
          <div class="tx-row">
            <span class="tx-date">${fmtDate(tx.date)}</span>
            <span class="tx-name" title="${escHtml(tx.name)}">${escHtml(tx.name)}</span>
            <span class="tx-amount ${amtClass}">${amtStr}</span>
            <select class="tx-cat-select" data-txid="${tx.id}">
              ${allCats.map(c => `<option value="${c}"${c===effectiveCat?' selected':''}>${c}</option>`).join('')}
            </select>
          </div>
        `;
      }).join('');

    return `
      <div class="category-group" data-cat="${escHtml(cat)}">
        <div class="category-header">
          <span class="category-dot" style="background:${color}"></span>
          <span>${escHtml(cat)}</span>
          <span class="category-total">${fmtGBP(total)}</span>
          <span class="category-chevron">▾</span>
        </div>
        <div class="category-rows">${rows}</div>
      </div>
    `;
  }).join('');

  // Collapsible headers
  container.querySelectorAll('.category-header').forEach(h => {
    h.addEventListener('click', () => {
      h.closest('.category-group').classList.toggle('collapsed');
    });
  });

  // Category override dropdowns
  container.querySelectorAll('.tx-cat-select').forEach(sel => {
    sel.addEventListener('change', e => {
      handleCategoryOverride(e.target.dataset.txid, e.target.value);
    });
  });
}

// --- Yearly Tab ---
function renderYearlyTab() {
  populateYearSelect();
  if (!state.activeYear) {
    const years = getYearsPresent(state.transactions);
    state.activeYear = years[years.length - 1] || null;
  }
  const yearSel = document.getElementById('yearSelect');
  if (state.activeYear) yearSel.value = state.activeYear;

  renderYearlyGrid();
}

function populateYearSelect() {
  const years = getYearsPresent(state.transactions);
  const sel = document.getElementById('yearSelect');
  sel.innerHTML = years.length === 0
    ? '<option value="">No data loaded</option>'
    : years.map(y => `<option value="${y}">${y}</option>`).join('');
}

function renderYearlyGrid() {
  const container = document.getElementById('yearlyGrid');
  const year = state.activeYear;

  if (!year) {
    container.innerHTML = '<div class="empty-state">No data loaded. Upload a CSV file first.</div>';
    return;
  }

  const matrix = getYearlyMatrix(year);
  const cats = Object.keys(matrix).sort();

  if (cats.length === 0) {
    container.innerHTML = '<div class="empty-state">No expenses found for this year.</div>';
    return;
  }

  // Get months present for this year
  const yearTx = state.transactions.filter(t => String(t.date.getFullYear()) === String(year));
  const monthsPresent = getMonthsPresent(yearTx).sort();

  // Per-category max for color normalization
  const catMax = {};
  for (const cat of cats) {
    catMax[cat] = Math.max(...Object.values(matrix[cat] || {}), 0.01);
  }

  const headerCells = monthsPresent.map(m => {
    const [, mo] = m.split('-');
    return `<th>${MONTH_NAMES[+mo - 1]}</th>`;
  }).join('');

  const bodyRows = cats.map(cat => {
    const color = getCategoryColor(cat);
    const rowTotal = Object.values(matrix[cat] || {}).reduce((a, b) => a + b, 0);

    const cells = monthsPresent.map(m => {
      const amt = matrix[cat][m];
      if (!amt) return `<td class="empty-cell">—</td>`;
      const intensity = amt / catMax[cat]; // 0..1
      const bg = hexWithAlpha(color, 0.12 + intensity * 0.55);
      return `<td class="data-cell" style="background:${bg}" data-month="${m}" title="${cat}: ${fmtGBP(amt)}">${fmtGBP(amt)}</td>`;
    }).join('');

    return `
      <tr>
        <td><span class="legend-dot" style="background:${color};display:inline-block;margin-right:6px;width:8px;height:8px;border-radius:50%"></span>${escHtml(cat)}</td>
        ${cells}
        <td><strong>${fmtGBP(rowTotal)}</strong></td>
      </tr>
    `;
  }).join('');

  // Footer: Total Expenses, Income, Net Savings
  const expensesByMonth = {};
  const incomeByMonth = {};
  for (const m of monthsPresent) {
    const mTx = state.transactions.filter(t => monthKey(t.date) === m);
    expensesByMonth[m] = mTx.filter(t => t.type === 'expense').reduce((s, t) => s + Math.abs(t.amount), 0);
    incomeByMonth[m] = state.income[m] || 0;
  }

  const totalExpenses = monthsPresent.map(m => expensesByMonth[m]);
  const totalIncome = monthsPresent.map(m => incomeByMonth[m]);
  const totalSavings = monthsPresent.map((m, i) => totalIncome[i] - totalExpenses[i]);

  const grandExpense = totalExpenses.reduce((a, b) => a + b, 0);
  const grandIncome = totalIncome.reduce((a, b) => a + b, 0);
  const grandSavings = grandIncome - grandExpense;

  const footerExpenses = monthsPresent.map((m, i) =>
    `<td>${fmtGBP(totalExpenses[i])}</td>`).join('') + `<td><strong>${fmtGBP(grandExpense)}</strong></td>`;
  const footerIncome = monthsPresent.map((m, i) =>
    `<td class="income-row-cell">${totalIncome[i] > 0 ? fmtGBP(totalIncome[i]) : '—'}</td>`).join('') +
    `<td class="income-row-cell"><strong>${grandIncome > 0 ? fmtGBP(grandIncome) : '—'}</strong></td>`;
  const footerSavings = monthsPresent.map((m, i) => {
    const cls = totalSavings[i] >= 0 ? 'savings-row-cell' : 'negative';
    return `<td class="${cls}">${totalIncome[i] > 0 ? fmtGBP(totalSavings[i]) : '—'}</td>`;
  }).join('') + `<td class="savings-row-cell"><strong>${grandIncome > 0 ? fmtGBP(grandSavings) : '—'}</strong></td>`;

  container.innerHTML = `
    <table class="yearly-table">
      <thead>
        <tr>
          <th>Category</th>
          ${headerCells}
          <th>Total</th>
        </tr>
      </thead>
      <tbody>${bodyRows}</tbody>
      <tfoot>
        <tr><td>Total Expenses</td>${footerExpenses}</tr>
        <tr><td>Income</td>${footerIncome}</tr>
        <tr><td>Net Savings</td>${footerSavings}</tr>
      </tfoot>
    </table>
  `;

  // Click cell → navigate to monthly tab
  container.querySelectorAll('.data-cell').forEach(cell => {
    cell.addEventListener('click', () => {
      const month = cell.dataset.month;
      if (month) {
        state.activeMonth = month;
        switchTab('monthly');
        renderMonthlyTab();
        document.getElementById('monthSelect').value = month;
      }
    });
  });
}

// ============================================================
// Column mapper (generic format)
// ============================================================
function showColumnMapper(headers) {
  const mapper = document.getElementById('columnMapper');
  const grid = document.getElementById('mapperGrid');
  mapper.classList.remove('hidden');

  const fields = [
    { key: 'date', label: 'Date column', required: true },
    { key: 'name', label: 'Description / Merchant column', required: true },
    { key: 'amount', label: 'Amount column', required: true },
    { key: 'category', label: 'Category column (optional)', required: false },
  ];

  grid.innerHTML = fields.map(f => `
    <div class="mapper-field">
      <label>${f.label}</label>
      <select class="select-input" id="map_${f.key}">
        ${!f.required ? '<option value="">— none —</option>' : ''}
        ${headers.map(h => `<option value="${h}">${h}</option>`).join('')}
      </select>
    </div>
  `).join('');
}

// ============================================================
// Event Handlers
// ============================================================
function handleCSVUpload(file) {
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const result = parseCSV(e.target.result);
      if (result.format === 'generic') {
        state.pendingGenericRows = result.rows;
        state.pendingGenericHeaders = result.headers;
        showColumnMapper(result.headers);
        showStatus(`Unknown format — please map the columns below.`, 'info');
        return;
      }
      importTransactions(result.transactions, result.bank, file.name);
    } catch (err) {
      showStatus(`Error parsing file: ${err.message}`, 'error');
      console.error(err);
    }
  };
  reader.readAsText(file);
}

async function handleFileUpload(file) {
  if (/\.pdf$/i.test(file.name)) {
    const zone = document.getElementById('uploadZone');
    zone.classList.add('processing');
    showStatus('Reading PDF…', 'info');
    try {
      const { transactions, bank } = await parsePDF(file);
      importTransactions(transactions, bank, file.name);
    } catch (err) {
      showStatus(`PDF error: ${err.message}`, 'error');
      console.error(err);
    } finally {
      zone.classList.remove('processing');
    }
  } else {
    handleCSVUpload(file);
  }
}

function importTransactions(newTxs, bank, filename) {
  const before = state.transactions.length;
  state.transactions = mergeTransactions(state.transactions, newTxs);
  const added = state.transactions.length - before;

  saveToStorage();
  renderUploadTab();

  const months = getMonthsPresent(newTxs);
  const cover = months.length > 0
    ? `${formatMonthLabel(months[0])} – ${formatMonthLabel(months[months.length-1])}`
    : 'unknown period';

  showStatus(
    `${bank.charAt(0).toUpperCase()+bank.slice(1)} detected ✓ — ` +
    `${added} new transactions added (${before + added} total), covering ${cover}.`,
    'success'
  );

  document.getElementById('columnMapper').classList.add('hidden');
}

function handleMonthChange(month) {
  state.activeMonth = month;
  renderMonthlyTab();
}

function handleYearChange(year) {
  state.activeYear = year;
  renderYearlyGrid();
}

function handleCategoryOverride(txId, newCategory) {
  state.categoryOverrides[txId] = newCategory;
  saveToStorage();
  renderMonthSummary();
  renderDonutChart();
  renderTransactionList();
}

function handleIncomeInput(month, amount) {
  if (!month) return;
  const val = parseFloat(amount);
  if (isNaN(val) || val < 0) {
    delete state.income[month];
  } else {
    state.income[month] = val;
  }
  saveToStorage();
  renderMonthSummary();
}

function showStatus(msg, type = 'info') {
  const el = document.getElementById('uploadStatus');
  el.textContent = msg;
  el.className = `upload-status ${type}`;
  el.classList.remove('hidden');
}

// ============================================================
// Utility
// ============================================================
function escHtml(str) {
  return String(str)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;');
}

function hexWithAlpha(hex, alpha) {
  const r = parseInt(hex.slice(1,3),16);
  const g = parseInt(hex.slice(3,5),16);
  const b = parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${alpha.toFixed(2)})`;
}

// ============================================================
// Tab switching
// ============================================================
function switchTab(tab) {
  state.activeTab = tab;
  document.querySelectorAll('.tab-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.tab === tab);
  });
  document.querySelectorAll('.tab-panel').forEach(p => {
    p.classList.toggle('active', p.id === `tab-${tab}`);
  });
  if (tab === 'monthly') renderMonthlyTab();
  if (tab === 'yearly')  renderYearlyTab();
  if (tab === 'upload')  renderUploadTab();
}

// ============================================================
// Init
// ============================================================
function init() {
  loadFromStorage();
  initPDFJS(); // fire-and-forget blob-worker setup

  // Theme
  const savedTheme = localStorage.getItem(LS_THEME) || 'light';
  document.documentElement.setAttribute('data-theme', savedTheme);
  document.getElementById('themeToggle').addEventListener('click', () => {
    const cur = document.documentElement.getAttribute('data-theme');
    const next = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem(LS_THEME, next);
  });

  // Tab navigation
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => switchTab(btn.dataset.tab));
  });

  // Upload zone drag-and-drop
  const zone = document.getElementById('uploadZone');
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const files = Array.from(e.dataTransfer.files).filter(f => /\.(csv|tsv|pdf)$/i.test(f.name));
    files.forEach(handleFileUpload);
  });
  zone.addEventListener('click', () => document.getElementById('fileInput').click());

  // File picker
  document.getElementById('fileInput').addEventListener('change', e => {
    Array.from(e.target.files).forEach(handleFileUpload);
    e.target.value = '';
  });

  // Generic mapper apply button
  document.getElementById('applyMapping').addEventListener('click', () => {
    if (!state.pendingGenericRows) return;
    const mapping = {
      date:     document.getElementById('map_date').value,
      name:     document.getElementById('map_name').value,
      amount:   document.getElementById('map_amount').value,
      category: document.getElementById('map_category').value || null,
    };
    if (!mapping.date || !mapping.name || !mapping.amount) {
      showStatus('Please map all required fields (Date, Description, Amount).', 'error');
      return;
    }
    const txs = parseGenericCSV(state.pendingGenericRows, mapping);
    importTransactions(txs, 'generic', 'uploaded file');
    state.pendingGenericRows = null;
    state.pendingGenericHeaders = null;
  });

  // Clear data
  document.getElementById('clearDataBtn').addEventListener('click', () => {
    if (!confirm('Delete all loaded transactions and overrides? This cannot be undone.')) return;
    state.transactions = [];
    state.income = {};
    state.categoryOverrides = {};
    state.customCategories = [];
    state.categoryColors = {};
    state.activeMonth = null;
    state.activeYear = null;
    saveToStorage();
    document.getElementById('uploadStatus').classList.add('hidden');
    document.getElementById('columnMapper').classList.add('hidden');
    renderUploadTab();
  });

  // Month select
  document.getElementById('monthSelect').addEventListener('change', e => handleMonthChange(e.target.value));

  // Year select
  document.getElementById('yearSelect').addEventListener('change', e => handleYearChange(e.target.value));

  // Income input
  const incomeInput = document.getElementById('incomeInput');
  incomeInput.addEventListener('change', e => handleIncomeInput(state.activeMonth, e.target.value));

  // Add category modal
  document.getElementById('addCategoryBtn').addEventListener('click', () => {
    document.getElementById('categoryModal').classList.remove('hidden');
    document.getElementById('newCategoryInput').focus();
  });
  document.getElementById('cancelCategoryBtn').addEventListener('click', () => {
    document.getElementById('categoryModal').classList.add('hidden');
    document.getElementById('newCategoryInput').value = '';
  });
  document.getElementById('saveCategoryBtn').addEventListener('click', () => {
    const val = document.getElementById('newCategoryInput').value.trim().toLowerCase();
    if (val && !state.customCategories.includes(val)) {
      state.customCategories.push(val);
      getCategoryColor(val); // assign a color
      saveToStorage();
    }
    document.getElementById('categoryModal').classList.add('hidden');
    document.getElementById('newCategoryInput').value = '';
    renderTransactionList();
  });
  // Close modal on overlay click
  document.getElementById('categoryModal').addEventListener('click', e => {
    if (e.target === document.getElementById('categoryModal')) {
      document.getElementById('categoryModal').classList.add('hidden');
    }
  });
  // Enter key in modal
  document.getElementById('newCategoryInput').addEventListener('keydown', e => {
    if (e.key === 'Enter') document.getElementById('saveCategoryBtn').click();
    if (e.key === 'Escape') document.getElementById('cancelCategoryBtn').click();
  });

  // Render initial state
  renderUploadTab();
}

document.addEventListener('DOMContentLoaded', init);
