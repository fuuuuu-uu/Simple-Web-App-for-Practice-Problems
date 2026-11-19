/* 刷題小幫手 — 單檔前端程式（無後端、無網路需求）
 *
 * 結構：
 *   1. 設定與欄位對應
 *   2. 檔案處理（載入 Excel → 解析成題目物件）
 *   3. 資料儲存（localStorage：題目與作答紀錄；IndexedDB：圖片）
 *   4. 圖片處理（預留介面）
 *   5. 出題 / 作答邏輯（練習、模擬考試、尋找）
 *   6. UI 繪製
 *   7. 匯出 Excel
 *   8. 例外處理與啟動
 */
'use strict';

/* ============================================================
 * 1. 設定與欄位對應
 * ============================================================ */

// 題庫無標題列時，依「欄位位置」讀取（對應你目前的 Excel：A~J）
const DEFAULT_COLUMNS = {
  no: 0,        // A 題號
  img: 1,       // B 圖片：「題」「選」「題 選」
  explain: 2,   // C 解析：1 = 有解析圖片 p題號.jpg
  fam: 3,       // D 熟悉度：知 / 背 / 算（可多個，以空白分隔）
  answer: 4,    // E 解答：A~D
  text: 5,      // F 題目
  optA: 6, optB: 7, optC: 8, optD: 9, // G~J 選項
  wrong: 10,    // K 錯誤次數（匯出後才會有）
  attempts: 11, // L 作答次數（匯出後才會有）
  parent: 12,   // M 承上題號（選填）
};

// 若第一列是標題列，用這些名稱辨識欄位（可自行增加別名）
const HEADER_ALIASES = {
  no: ['題號', '編號', 'no', 'id'],
  img: ['圖片', '圖', 'image'],
  explain: ['解析', 'explain'],
  fam: ['熟悉度', '標籤', 'tag'],
  answer: ['解答', '答案', 'answer'],
  text: ['題目', 'question'],
  optA: ['選項a', 'a', '(a)'],
  optB: ['選項b', 'b', '(b)'],
  optC: ['選項c', 'c', '(c)'],
  optD: ['選項d', 'd', '(d)'],
  wrong: ['錯誤次數', '錯誤', 'wrong'],
  attempts: ['作答次數', 'attempts'],
  parent: ['承上題', '承上', 'parent'],
};

const LETTERS = ['A', 'B', 'C', 'D'];
const DEFAULT_FAM_TAGS = ['知', '背', '算'];        // 熟悉度可選標籤
const CHAIN_RE = /^\s*[（(]?\s*(承上題|續上題|接上題|承上|續上)/; // 「承上題」偵測
const LS_KEY = 'quiz.bank.v1';
const LS_PREF = 'quiz.pref.v1';
const IMG_EXT = ['jpg', 'jpeg', 'png', 'webp', 'gif'];

/* ============================================================
 * 狀態
 * ============================================================ */
const state = {
  bank: null,          // { name, loadedAt, famTags:[], questions:[...] }
  byId: new Map(),     // 題號(字串) → 題目
  mode: 'practice',    // practice | exam | search
  session: null,       // 目前作答階段，見 startXxx()
  images: new Map(),   // 檔名(小寫，不含副檔名) → objectURL
};

/* ============================================================
 * 2. 檔案處理：載入並解析 Excel
 * ============================================================ */

/** 讀取使用者選的題庫檔案。任何錯誤都只顯示「檢查檔案」，不讓程式中斷。 */
async function loadBankFile(file) {
  if (!file) { showFileError('沒有選擇檔案。'); return; }
  try {
    const buf = await file.arrayBuffer();
    const bank = parseWorkbook(buf, file.name);
    mergeOldStats(bank, state.bank);
    setBank(bank);
    saveBank();
    toast(`已載入 ${bank.questions.length} 題` + (bank.skipped ? `（略過 ${bank.skipped} 列格式不符）` : ''));
    showView('home');
  } catch (err) {
    console.error(err);
    showFileError(err && err.userMessage ? err.userMessage : '檔案無法讀取，可能已毀損或格式不正確。', err);
  }
}

function fileError(msg) { const e = new Error(msg); e.userMessage = msg; return e; }

/** ArrayBuffer → bank 物件 */
function parseWorkbook(buf, fileName) {
  if (typeof XLSX === 'undefined') throw fileError('Excel 函式庫載入失敗（lib/xlsx.full.min.js）。');
  let wb;
  try { wb = XLSX.read(buf, { type: 'array' }); }
  catch (e) { throw fileError('檔案已毀損或不是 Excel 格式。'); }
  if (!wb.SheetNames || !wb.SheetNames.length) throw fileError('檔案中沒有任何工作表。');

  // 取第一個有資料的工作表
  let rows = null;
  for (const name of wb.SheetNames) {
    const r = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: false });
    if (r.some(row => row.some(c => String(c).trim() !== ''))) { rows = r; break; }
  }
  if (!rows) throw fileError('工作表是空的。');

  const { cols, startRow } = detectColumns(rows[0]);
  const questions = [];
  const seen = new Set();
  let skipped = 0;

  for (let i = startRow; i < rows.length; i++) {
    const q = parseRow(rows[i], cols);
    if (!q) { if (rows[i].some(c => String(c).trim())) skipped++; continue; }
    if (seen.has(q.id)) { skipped++; continue; }
    seen.add(q.id);
    questions.push(q);
  }
  if (!questions.length) throw fileError('找不到有效題目（每題至少需有：題號、題目、4 個選項、解答）。');

  questions.sort((a, b) => a.no - b.no || a.id.localeCompare(b.id));
  linkParents(questions);

  const famTags = [...DEFAULT_FAM_TAGS];
  questions.forEach(q => q.fam.forEach(t => { if (!famTags.includes(t)) famTags.push(t); }));

  return { name: fileName.replace(/\.[^.]+$/, ''), loadedAt: Date.now(), famTags, skipped, questions };
}

/** 判斷第一列是不是標題列，回傳欄位對應 */
function detectColumns(firstRow) {
  const cells = (firstRow || []).map(c => String(c).trim().toLowerCase());
  const looksLikeHeader = cells.length && !/^\d+$/.test(cells[0]) && cells.some(c => ['題號', '題目', '解答'].includes(c));
  if (!looksLikeHeader) return { cols: { ...DEFAULT_COLUMNS }, startRow: 0 };

  const cols = {};
  for (const [key, aliases] of Object.entries(HEADER_ALIASES)) {
    const idx = cells.findIndex(c => aliases.includes(c));
    cols[key] = idx >= 0 ? idx : -1;
  }
  // 標題列缺少必要欄位時，退回預設欄位位置
  for (const k of ['no', 'answer', 'text', 'optA', 'optB', 'optC', 'optD']) {
    if (cols[k] < 0) cols[k] = DEFAULT_COLUMNS[k];
  }
  return { cols, startRow: 1 };
}

/** 一列 → 題目物件；缺必要欄位回傳 null */
function parseRow(row, cols) {
  const get = k => (cols[k] >= 0 && cols[k] < row.length) ? String(row[cols[k]] ?? '').trim() : '';

  const id = get('no');
  const text = get('text');
  const answer = normalizeAnswer(get('answer'));
  const options = ['optA', 'optB', 'optC', 'optD'].map(k => cleanOption(get(k)));
  const rawOptions = ['optA', 'optB', 'optC', 'optD'].map(get);
  const imgFlag = parseImageFlag(get('img'));

  if (!id || !answer) return null;
  if (!text && !imgFlag.question) return null;
  // 選項：每格至少要存在（純圖片選項可能只寫「(A)」）
  if (rawOptions.some(o => o === '') && !imgFlag.options) return null;

  const explainRaw = get('explain');
  return {
    id,
    no: parseFloat(id) || 0,
    text,
    options,
    answer,
    hasExplain: /^(1|y|yes|true|有|v|ˇ|o)$/i.test(explainRaw),
    img: imgFlag,
    fam: splitTags(get('fam')),
    wrong: toInt(get('wrong')),
    attempts: toInt(get('attempts')),
    last: null,                       // 'ok' | 'ng' | null
    parentId: get('parent') || null,  // 若 Excel 有指定承上題號
  };
}

function normalizeAnswer(s) {
  s = String(s).toUpperCase().replace(/[（）()\s.。]/g, '');
  s = s.replace(/[Ａ-Ｄ]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));
  if (/^[A-D]$/.test(s)) return s;
  if (/^[1-4]$/.test(s)) return LETTERS[+s - 1];
  return null;
}

/** 去掉開頭的 (A) / （A） / A. 與結尾句號 */
function cleanOption(s) {
  return String(s).replace(/^\s*[（(]\s*[A-DＡ-Ｄa-d]\s*[)）]\s*/, '').replace(/^\s*[A-D][.、．]\s*/, '').replace(/[。．]\s*$/, '').trim();
}

/** 「題」「選」「題 選」→ { question, options } */
function parseImageFlag(s) {
  s = String(s);
  return { question: s.includes('題'), options: s.includes('選') };
}

function splitTags(s) {
  return String(s).split(/[\s,，、/]+/).map(t => t.trim()).filter(Boolean);
}
function toInt(s) { const n = parseInt(s, 10); return Number.isFinite(n) && n > 0 ? n : 0; }

/** 偵測「承上題」：題目以「承上題／續上題」開頭 → 連到前一題 */
function linkParents(questions) {
  for (let i = 0; i < questions.length; i++) {
    const q = questions[i];
    if (q.parentId) continue;
    if (CHAIN_RE.test(q.text) && i > 0) q.parentId = questions[i - 1].id;
  }
}

/** 取得某題所有「被承上」的題目（由遠到近），防止循環 */
function getParentChain(q) {
  const chain = [];
  const seen = new Set([q.id]);
  let cur = q;
  while (cur.parentId && state.byId.has(cur.parentId) && !seen.has(cur.parentId)) {
    cur = state.byId.get(cur.parentId);
    seen.add(cur.id);
    chain.unshift(cur);
  }
  return chain;
}

/** 重新載入同一份題庫時，保留舊的作答紀錄（題號 + 題目相同才沿用） */
function mergeOldStats(newBank, oldBank) {
  if (!oldBank) return;
  const old = new Map(oldBank.questions.map(q => [q.id, q]));
  for (const q of newBank.questions) {
    const o = old.get(q.id);
    if (!o || o.text !== q.text) continue;
    q.wrong = Math.max(q.wrong, o.wrong || 0);
    q.attempts = Math.max(q.attempts, o.attempts || 0);
    q.last = o.last || null;
  }
}

/* ============================================================
 * 3. 資料儲存
 * ============================================================ */

function setBank(bank) {
  state.bank = bank;
  state.byId = new Map(bank.questions.map(q => [q.id, q]));
}

function saveBank() {
  if (!state.bank) return;
  try { localStorage.setItem(LS_KEY, JSON.stringify(state.bank)); }
  catch (e) { console.warn('儲存失敗', e); toast('⚠ 本機儲存空間不足，請匯出 Excel 備份'); }
}

function loadSavedBank() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (!raw) return;
    const bank = JSON.parse(raw);
    if (bank && Array.isArray(bank.questions) && bank.questions.length) setBank(bank);
  } catch (e) {
    console.warn('本機資料損毀，已忽略', e);
    try { localStorage.removeItem(LS_KEY); } catch (_) {}
  }
}

function loadPref() { try { return JSON.parse(localStorage.getItem(LS_PREF)) || {}; } catch (_) { return {}; } }
function savePref(p) { try { localStorage.setItem(LS_PREF, JSON.stringify({ ...loadPref(), ...p })); } catch (_) {} }

/* ---- IndexedDB：圖片 ---- */
const IDB = {
  db: null,
  open() {
    if (this.db) return Promise.resolve(this.db);
    return new Promise((resolve, reject) => {
      if (!('indexedDB' in window)) return reject(new Error('no indexedDB'));
      const req = indexedDB.open('quiz-images', 1);
      req.onupgradeneeded = () => req.result.createObjectStore('img');
      req.onsuccess = () => { this.db = req.result; resolve(this.db); };
      req.onerror = () => reject(req.error);
    });
  },
  async tx(mode, fn) {
    const db = await this.open();
    return new Promise((resolve, reject) => {
      const t = db.transaction('img', mode);
      const store = t.objectStore('img');
      const out = fn(store);
      t.oncomplete = () => resolve(out && out.result !== undefined ? out.result : out);
      t.onerror = () => reject(t.error);
    });
  },
  put(key, blob) { return this.tx('readwrite', s => s.put(blob, key)); },
  clear() { return this.tx('readwrite', s => s.clear()); },
  async all() {
    const db = await this.open();
    return new Promise((resolve, reject) => {
      const out = [];
      const req = db.transaction('img', 'readonly').objectStore('img').openCursor();
      req.onsuccess = () => { const c = req.result; if (c) { out.push([c.key, c.value]); c.continue(); } else resolve(out); };
      req.onerror = () => reject(req.error);
    });
  },
};

/* ============================================================
 * 4. 圖片處理（預留介面）
 *
 * 檔名規則（不分大小寫、副檔名 jpg/png/webp 皆可）：
 *   解析圖     p{題號}      例：p12.jpg
 *   題目圖     q{題號}      例：q12.jpg
 *   選項圖     q{題號}{A~D} 例：q12A.jpg、q12B.jpg ...
 * 若之後想改規則，只要改下面三個 xxxImageName() 即可。
 * ============================================================ */

const questionImageName = q => `q${q.id}`;
const optionImageName = (q, letter) => `q${q.id}${letter}`;
const explainImageName = q => `p${q.id}`;

function imageKey(fileName) {
  return fileName.toLowerCase().replace(/\.[^.]+$/, '');
}

/** 使用者選取多張圖片 → 存進 IndexedDB */
async function loadImageFiles(files) {
  const list = [...(files || [])].filter(f => IMG_EXT.includes((f.name.split('.').pop() || '').toLowerCase()));
  if (!list.length) { showFileError('沒有選到可用的圖片檔（jpg / png / webp）。'); return; }
  let ok = 0;
  for (const f of list) {
    const key = imageKey(f.name);
    try { await IDB.put(key, f); } catch (e) { console.warn('圖片存入 IndexedDB 失敗，僅本次可用', e); }
    setImageURL(key, f);
    ok++;
  }
  toast(`已載入 ${ok} 張圖片`);
  if (state.session) renderQuestion();
}

function setImageURL(key, blob) {
  const old = state.images.get(key);
  if (old) URL.revokeObjectURL(old);
  state.images.set(key, URL.createObjectURL(blob));
}

async function restoreImages() {
  try { for (const [k, blob] of await IDB.all()) setImageURL(k, blob); }
  catch (e) { console.warn('讀取已存圖片失敗', e); }
}

/** 依檔名（不含副檔名）取得圖片網址；沒有則回傳 null */
function getImageURL(name) {
  return state.images.get(String(name).toLowerCase()) || null;
}

/** 在 container 放入圖片；若應有圖但未載入，顯示提示 */
function renderImageInto(container, name, expected, label) {
  const url = getImageURL(name);
  if (url) {
    const img = document.createElement('img');
    img.src = url;
    img.alt = label;
    img.loading = 'lazy';
    img.onerror = () => { img.replaceWith(missingImageNote(name, label + '（圖片無法顯示）')); };
    container.appendChild(img);
    return true;
  }
  if (expected) container.appendChild(missingImageNote(name, label));
  return false;
}

function missingImageNote(name, label) {
  const d = document.createElement('div');
  d.className = 'imgmissing';
  d.textContent = `🖼️ ${label}：尚未載入圖片 ${name}.jpg`;
  return d;
}

function renderQuestionImage(q, container) {
  container.replaceChildren();
  renderImageInto(container, questionImageName(q), q.img.question, '題目圖');
}

function renderOptionImage(q, letter, container) {
  return renderImageInto(container, optionImageName(q, letter), q.img.options, `選項 ${letter} 圖`);
}

function renderExplainImage(q, container) {
  container.replaceChildren();
  if (!q.hasExplain && !getImageURL(explainImageName(q))) return;
  const h = document.createElement('div');
  h.className = 'hint';
  h.textContent = '解析：';
  container.appendChild(h);
  renderImageInto(container, explainImageName(q), true, '解析');
}

/* ============================================================
 * 5. 出題 / 作答邏輯
 * ============================================================ */

function shuffle(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function getPool() {
  const qs = state.bank.questions;
  const sel = $('#selPool').value;
  if (sel === 'unseen') return qs.filter(q => !q.attempts);
  if (sel === 'wrong') return qs.filter(q => q.wrong > 0);
  if (sel === 'tag') {
    const tags = [...document.querySelectorAll('#poolTags .chip.on')].map(b => b.dataset.tag);
    if (!tags.length) return qs;
    return qs.filter(q => q.fam.some(t => tags.includes(t)));
  }
  return qs;
}

/* 練習：隨機抽題，作答完立即公布 */
function startPractice() {
  const pool = getPool();
  if (!pool.length) { toast('這個範圍沒有題目'); return; }
  state.session = {
    type: 'practice',
    order: shuffle(pool).map(q => q.id), // 隨機、不重複，抽完再重洗
    idx: 0,
    picked: null,
    revealed: false,
  };
  showView('quiz');
  renderQuestion();
}

/* 模擬考試：全部作答完才公布 */
function startExam() {
  const n = Math.max(1, Math.min(parseInt($('#examCount').value, 10) || 50, state.bank.questions.length));
  savePref({ examCount: n });
  state.session = {
    type: 'exam',
    order: shuffle(state.bank.questions).slice(0, n).map(q => q.id),
    idx: 0,
    answers: {},       // id → 'A'~'D'
    submitted: false,
  };
  showView('quiz');
  renderQuestion();
}

/* 尋找 / 跳題：顯示指定題目，可直接作答 */
function openSingle(id) {
  if (!state.byId.has(id)) { toast(`找不到第 ${id} 題`); return; }
  state.session = { type: 'single', order: [id], idx: 0, picked: null, revealed: false };
  showView('quiz');
  renderQuestion();
}

function currentQ() {
  const s = state.session;
  return s ? state.byId.get(s.order[s.idx]) : null;
}

function pickOption(letter) {
  const s = state.session;
  if (!s) return;
  if (s.type === 'exam') {
    if (s.submitted) return;
    s.answers[s.order[s.idx]] = letter;
  } else {
    if (s.revealed) return;
    s.picked = letter;
  }
  renderQuestion();
}

function recordResult(q, correct) {
  q.attempts = (q.attempts || 0) + 1;
  if (!correct) q.wrong = (q.wrong || 0) + 1;
  q.last = correct ? 'ok' : 'ng';
}

function confirmAnswer() {
  const s = state.session, q = currentQ();
  if (!s || !q) return;
  if (s.type === 'exam') {
    if (s.submitted) return goNext();
    if (!s.answers[q.id]) { toast('請先選擇一個選項'); return; }
    // 跳到下一個未作答的題目
    const next = s.order.findIndex((id, i) => i > s.idx && !s.answers[id]);
    const any = s.order.findIndex(id => !s.answers[id]);
    if (next >= 0) s.idx = next;
    else if (any >= 0) s.idx = any;
    else { askSubmitExam(); return; }
    renderQuestion();
    return;
  }
  if (s.revealed) return goNext();
  if (!s.picked) { toast('請先選擇一個選項'); return; }
  s.revealed = true;
  recordResult(q, s.picked === q.answer);
  saveBank();
  renderQuestion();
}

function goNext() {
  const s = state.session;
  if (!s) return;
  if (s.type === 'practice') {
    s.idx++;
    if (s.idx >= s.order.length) { s.order = shuffle(s.order); s.idx = 0; toast('本輪題目已全部抽完，重新洗牌'); }
    s.picked = null; s.revealed = false;
  } else if (s.type === 'exam') {
    if (s.idx < s.order.length - 1) s.idx++;
  } else if (s.type === 'single') {
    const qs = state.bank.questions;
    const i = qs.findIndex(q => q.id === s.order[0]);
    if (i < qs.length - 1) { s.order = [qs[i + 1].id]; s.picked = null; s.revealed = false; }
  }
  renderQuestion();
}

function goPrev() {
  const s = state.session;
  if (!s) return;
  if (s.type === 'exam') { if (s.idx > 0) s.idx--; }
  else if (s.type === 'single') {
    const qs = state.bank.questions;
    const i = qs.findIndex(q => q.id === s.order[0]);
    if (i > 0) { s.order = [qs[i - 1].id]; s.picked = null; s.revealed = false; }
  } else if (s.type === 'practice') {
    if (s.idx > 0) { s.idx--; s.picked = null; s.revealed = false; }
  }
  renderQuestion();
}

function askSubmitExam() {
  const s = state.session;
  const left = s.order.filter(id => !s.answers[id]).length;
  showDialog('交卷', left ? `還有 ${left} 題未作答，確定要交卷嗎？` : '全部作答完成，確定交卷？', [
    { text: '繼續作答' },
    { text: '交卷', cls: 'warn', fn: submitExam },
  ]);
}

function submitExam() {
  const s = state.session;
  s.submitted = true;
  let correct = 0;
  for (const id of s.order) {
    const q = state.byId.get(id);
    const ok = s.answers[id] === q.answer;
    if (ok) correct++;
    recordResult(q, ok);
  }
  saveBank();
  s.score = correct;
  renderExamResult();
  showView('examResult');
}

function toggleFam(tag) {
  const q = currentQ();
  if (!q) return;
  q.fam = q.fam.includes(tag) ? q.fam.filter(t => t !== tag) : [...q.fam, tag];
  saveBank();
  renderFamChips(q);
}

/* ============================================================
 * 6. UI 繪製
 * ============================================================ */

const $ = sel => document.querySelector(sel);
function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}

const VIEWS = ['home', 'quiz', 'examResult', 'list', 'end'];
function showView(name) {
  for (const v of VIEWS) $('#view' + v[0].toUpperCase() + v.slice(1)).classList.toggle('hidden', v !== name);
  $('#menu').classList.add('hidden');
  if (name === 'home') renderHome();
  if (name === 'list') renderList();
  window.scrollTo(0, 0);
}

function renderHome() {
  const box = $('#bankStatus');
  box.replaceChildren();
  const b = state.bank;
  if (!b) {
    box.append(el('h2', null, '尚未載入題庫'));
    box.append(el('p', 'hint', '請點右上「⋯」→「載入題庫 Excel」。題庫只存在你的手機，不會上傳。'));
    const btn = el('button', 'primary', '📂 載入題庫');
    btn.onclick = () => $('#fileBank').click();
    box.append(btn);
  } else {
    const qs = b.questions;
    const done = qs.filter(q => q.attempts).length;
    const wrong = qs.filter(q => q.wrong).length;
    box.append(el('h2', null, b.name));
    const st = el('div', 'stats');
    st.innerHTML = '';
    for (const [n, t] of [[qs.length, '總題數'], [done, '已作答'], [wrong, '曾答錯']]) {
      const d = el('div'); d.append(el('b', null, String(n)), el('span', null, t)); st.append(d);
    }
    box.append(st);
    const imgCount = state.images.size;
    box.append(el('p', 'hint', `已載入圖片 ${imgCount} 張`));
  }
  $('#btnStart').disabled = !b;
  $('#appTitle').textContent = b ? b.name : '刷題小幫手';

  // 模式
  document.querySelectorAll('#modeSeg button').forEach(x => x.classList.toggle('on', x.dataset.mode === state.mode));
  $('#optPractice').classList.toggle('hidden', state.mode !== 'practice');
  $('#optExam').classList.toggle('hidden', state.mode !== 'exam');
  $('#optSearch').classList.toggle('hidden', state.mode !== 'search');
  $('#modeHint').textContent = {
    practice: '隨機抽題，每題按「確定」後立即公布解答與解析。',
    exam: '隨機抽出指定題數，全部作答完（交卷）後才公布解答。',
    search: '輸入題號直接顯示該題。',
  }[state.mode];
  $('#btnStart').textContent = state.mode === 'search' ? '顯示題目' : '開始';

  // 熟悉度標籤篩選
  const tagBox = $('#poolTags');
  tagBox.classList.toggle('hidden', $('#selPool').value !== 'tag');
  if (b && tagBox.dataset.built !== b.loadedAt + '') {
    tagBox.replaceChildren();
    for (const t of b.famTags) {
      const c = el('button', 'chip', t);
      c.dataset.tag = t;
      c.onclick = () => c.classList.toggle('on');
      tagBox.append(c);
    }
    tagBox.dataset.built = b.loadedAt + '';
  }
}

function renderQuestion() {
  const s = state.session, q = currentQ();
  if (!s || !q) return;
  const isExam = s.type === 'exam';
  const reveal = isExam ? s.submitted : s.revealed;
  const picked = isExam ? s.answers[q.id] : s.picked;

  // 頂部
  $('#qProgress').textContent =
    isExam ? `第 ${s.idx + 1} / ${s.order.length} 題　已答 ${Object.keys(s.answers).length}`
    : s.type === 'practice' ? `練習　本輪 ${s.idx + 1} / ${s.order.length}`
    : '尋找';
  $('#qNo').textContent = `#${q.id}`;

  // 承上題
  const pb = $('#parentBlock');
  pb.replaceChildren();
  const chain = getParentChain(q);
  pb.classList.toggle('hidden', !chain.length);
  for (const p of chain) {
    const card = el('div', 'card parent');
    card.append(el('div', 'ptitle', `被承上題 #${p.id}`));
    card.append(el('div', 'qtext', p.text));
    const pimg = el('div', 'imgslot'); renderQuestionImage(p, pimg); card.append(pimg);
    const po = el('div', 'popts');
    p.options.forEach((o, i) => po.append(el('div', null, `(${LETTERS[i]}) ${o}`)));
    card.append(po);
    pb.append(card);
  }

  // 題目
  $('#qText').textContent = q.text;
  renderQuestionImage(q, $('#qImage'));

  // 選項（單選）
  const box = $('#qOptions');
  box.replaceChildren();
  q.options.forEach((text, i) => {
    const L = LETTERS[i];
    const b = el('button', 'opt');
    b.setAttribute('role', 'radio');
    b.setAttribute('aria-checked', picked === L ? 'true' : 'false');
    b.append(el('span', 'lt', `(${L})`));
    const body = el('span', 'ot');
    if (text) body.append(el('div', null, text));
    const imgWrap = el('div', 'oimg');
    renderOptionImage(q, L, imgWrap);
    if (imgWrap.childNodes.length) body.append(imgWrap);
    b.append(body);
    if (picked === L) b.classList.add('sel');
    if (reveal) {
      if (L === q.answer) b.classList.add('right');
      else if (picked === L) b.classList.add('wrong');
      b.disabled = false; // 保留樣式可讀性
    }
    b.onclick = () => pickOption(L);
    box.append(b);
  });

  // 解答 / 解析
  const rb = $('#resultBox');
  rb.classList.toggle('hidden', !reveal);
  if (reveal) {
    const line = $('#resultLine');
    line.replaceChildren();
    if (picked) {
      const ok = picked === q.answer;
      line.append(el('span', ok ? 'ok' : 'ng', ok ? '✔ 答對了' : `✘ 答錯，你選 (${picked})`));
    } else line.append(el('span', 'ng', '未作答'));
    line.append(el('span', null, `　正確答案：(${q.answer})`));
    line.append(el('div', 'hint', `累計作答 ${q.attempts} 次、錯誤 ${q.wrong} 次`));
    renderExplainImage(q, $('#explainBox'));
  }

  renderFamChips(q);

  // 按鈕
  $('#btnConfirm').textContent = reveal ? '下一題' : '確定';
  $('#btnSubmitExam').classList.toggle('hidden', !isExam || s.submitted);
  $('#btnPrev').disabled = (isExam || s.type === 'practice') ? s.idx === 0 : state.bank.questions[0].id === q.id;
  $('#btnNext').disabled = isExam ? s.idx >= s.order.length - 1
    : s.type === 'single' ? state.bank.questions[state.bank.questions.length - 1].id === q.id : false;
  if (isExam && s.submitted) $('#btnConfirm').textContent = '下一題';
  $('#btnBackHome').textContent = isExam && s.submitted ? '回成績' : '回首頁';
}

function renderFamChips(q) {
  const box = $('#famChips');
  box.replaceChildren();
  for (const t of state.bank.famTags) {
    const c = el('button', 'chip' + (q.fam.includes(t) ? ' on' : ''), t);
    c.onclick = () => toggleFam(t);
    box.append(c);
  }
}

function renderExamResult() {
  const s = state.session;
  $('#examScore').textContent = `${s.score} / ${s.order.length}　(${Math.round(s.score / s.order.length * 100)} 分)`;
  const grid = $('#examGrid');
  grid.replaceChildren();
  s.order.forEach((id, i) => {
    const q = state.byId.get(id);
    const b = el('button', s.answers[id] === q.answer ? 'ok' : 'ng', id);
    b.onclick = () => { s.idx = i; showView('quiz'); renderQuestion(); };
    grid.append(b);
  });
}

function renderList() {
  const grid = $('#listGrid');
  grid.replaceChildren();
  if (!state.bank) { grid.append(el('p', 'hint', '尚未載入題庫')); return; }
  const s = state.session;
  const inExam = s && s.type === 'exam' && !s.submitted;
  const ids = inExam ? s.order : state.bank.questions.map(q => q.id);
  const cur = currentQ();
  ids.forEach((id, i) => {
    const q = state.byId.get(id);
    let cls = inExam ? (s.answers[id] ? 'ok' : '') : (q.last || '');
    if (cur && cur.id === id) cls += ' cur';
    const b = el('button', cls.trim(), id);
    b.onclick = () => {
      if (inExam) { s.idx = i; showView('quiz'); renderQuestion(); }
      else openSingle(id);
    };
    grid.append(b);
  });
}

/* ---- 對話框 / 提示 ---- */
function showDialog(title, body, buttons) {
  $('#dlgTitle').textContent = title;
  const b = $('#dlgBody');
  b.replaceChildren();
  if (typeof body === 'string') b.append(el('p', null, body)); else b.append(body);
  const bb = $('#dlgBtns');
  bb.replaceChildren();
  for (const btn of buttons || [{ text: '關閉' }]) {
    const x = el('button', btn.cls || '', btn.text);
    x.onclick = () => { $('#dialog').classList.add('hidden'); btn.fn && btn.fn(); };
    bb.append(x);
  }
  $('#dialog').classList.remove('hidden');
}

/** 檔案錯誤：顯示「檢查檔案」並等使用者自行關閉 */
function showFileError(msg, err) {
  const body = el('div');
  body.append(el('p', 'err', msg));
  body.append(el('p', null, '請確認：副檔名為 .xlsx、檔案能用 Excel 正常開啟、每題至少有題號／題目／4 個選項／解答。'));
  if (err && err.message && err.message !== msg) body.append(el('pre', null, err.message));
  showDialog('⚠ 檢查檔案', body, [{ text: '關閉' }]);
}

let toastTimer;
function toast(msg) {
  const t = $('#toast');
  t.textContent = msg;
  t.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.add('hidden'), 2200);
}

/* ============================================================
 * 7. 匯出 Excel（建立／覆寫進度表）
 *   瀏覽器無法直接覆寫原檔，會下載一份新檔；
 *   之後用這份新檔「載入題庫」即可接續（含錯誤次數、熟悉度）。
 * ============================================================ */
function exportExcel() {
  if (!state.bank) { toast('尚未載入題庫'); return; }
  try {
    const header = ['題號', '圖片', '解析', '熟悉度', '解答', '題目', '選項A', '選項B', '選項C', '選項D', '錯誤次數', '作答次數', '承上題'];
    const rows = state.bank.questions.map(q => [
      isNaN(+q.id) ? q.id : +q.id,
      [q.img.question ? '題' : '', q.img.options ? '選' : ''].filter(Boolean).join(' '),
      q.hasExplain ? 1 : '',
      q.fam.join(' '),
      q.answer,
      q.text,
      ...q.options.map((o, i) => `(${LETTERS[i]})${o}`),
      q.wrong || 0,
      q.attempts || 0,
      q.parentId || '',
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    ws['!cols'] = [5, 6, 5, 7, 5, 50, 20, 20, 20, 20, 8, 8, 7].map(w => ({ wch: w }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '題庫');
    XLSX.writeFile(wb, `${state.bank.name}.xlsx`);
    toast('已匯出 Excel（在「下載」資料夾）');
  } catch (e) {
    console.error(e);
    showDialog('匯出失敗', String(e && e.message || e));
  }
}

/* ---- 結束 ---- */
function endApp() {
  showDialog('結束', '要結束刷題嗎？作答紀錄已自動存在本機。', [
    { text: '取消' },
    { text: '匯出 Excel 後結束', fn: () => { exportExcel(); setTimeout(doEnd, 600); } },
    { text: '結束', cls: 'danger', fn: doEnd },
  ]);
}
function doEnd() {
  saveBank();
  state.session = null;
  showView('end');
  try { window.close(); } catch (_) {}
}

/* ============================================================
 * 8. 事件綁定、例外處理、啟動
 * ============================================================ */

function bindEvents() {
  $('#btnMenu').onclick = e => { e.stopPropagation(); $('#menu').classList.toggle('hidden'); };
  document.addEventListener('click', e => { if (!e.target.closest('#menu')) $('#menu').classList.add('hidden'); });
  $('#btnHome').onclick = () => showView('home');
  $('#btnList').onclick = () => showView('list');

  $('#menu').onclick = e => {
    const act = e.target.closest('button')?.dataset.act;
    $('#menu').classList.add('hidden');
    if (act === 'loadBank') $('#fileBank').click();
    if (act === 'loadImages') $('#fileImages').click();
    if (act === 'export') exportExcel();
    if (act === 'end') endApp();
    if (act === 'resetStats') showDialog('清除作答紀錄', '將所有題目的錯誤次數、作答次數歸零（熟悉度保留）。', [
      { text: '取消' },
      { text: '清除', cls: 'danger', fn: () => { state.bank?.questions.forEach(q => { q.wrong = 0; q.attempts = 0; q.last = null; }); saveBank(); showView('home'); } },
    ]);
    if (act === 'clearAll') showDialog('清除所有本機資料', '題庫、作答紀錄與圖片都會從這支手機刪除（建議先匯出 Excel）。', [
      { text: '取消' },
      { text: '全部清除', cls: 'danger', fn: async () => {
        try { localStorage.removeItem(LS_KEY); } catch (_) {}
        try { await IDB.clear(); } catch (_) {}
        state.images.forEach(u => URL.revokeObjectURL(u)); state.images.clear();
        state.bank = null; state.byId = new Map(); state.session = null;
        showView('home');
      } },
    ]);
  };

  $('#fileBank').onchange = e => { const f = e.target.files[0]; e.target.value = ''; loadBankFile(f); };
  $('#fileImages').onchange = e => { const f = e.target.files; loadImageFiles(f).finally(() => { e.target.value = ''; }); };

  $('#modeSeg').onclick = e => {
    const m = e.target.closest('button')?.dataset.mode;
    if (!m) return;
    state.mode = m; savePref({ mode: m }); renderHome();
  };
  $('#selPool').onchange = renderHome;
  $('#btnStart').onclick = () => {
    if (!state.bank) { showFileError('尚未載入題庫檔案。'); return; }
    if (state.mode === 'practice') startPractice();
    else if (state.mode === 'exam') startExam();
    else {
      const no = $('#searchNo').value.trim();
      if (!no) { toast('請輸入題號'); return; }
      openSingle(no);
    }
  };
  $('#searchNo').onkeydown = e => { if (e.key === 'Enter') $('#btnStart').click(); };

  $('#btnConfirm').onclick = confirmAnswer;
  $('#btnNext').onclick = goNext;
  $('#btnPrev').onclick = goPrev;
  $('#btnSubmitExam').onclick = askSubmitExam;
  $('#btnBackHome').onclick = () => {
    const s = state.session;
    if (s && s.type === 'exam' && s.submitted) showView('examResult');
    else showView('home');
  };
  $('#btnExamHome').onclick = () => showView('home');
  $('#btnReopen').onclick = () => showView('home');

  // 離開頁面前存檔
  document.addEventListener('visibilitychange', () => { if (document.hidden) saveBank(); });
}

// 全域例外：任何未預期錯誤都不讓畫面卡死
window.addEventListener('error', e => { console.error(e.error || e.message); toast('發生錯誤：' + (e.message || '未知')); });
window.addEventListener('unhandledrejection', e => { console.error(e.reason); toast('發生錯誤：' + (e.reason?.message || e.reason)); });

async function init() {
  bindEvents();
  const pref = loadPref();
  if (pref.mode) state.mode = pref.mode;
  if (pref.examCount) $('#examCount').value = pref.examCount;
  loadSavedBank();
  await restoreImages();
  showView('home');

  if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
    navigator.serviceWorker.register('sw.js').catch(e => console.warn('SW 註冊失敗', e));
  }
}

init();
