/***************************************
 * TRAVESSIA SAGRADA – APP SCRIPT
 * Backend (Code.gs)
 ***************************************/

const DATE_FMT_BR = 'dd/MM/yyyy';

/** ===== Template / HTTP ===== */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  return t.evaluate()
    .setTitle('Travessia Sagrada')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** ===== Utils ===== */
function _normalize_(s) {
  return String(s || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, ''); // remove acentos
}
function _toBool_(v) {
  const n = _normalize_(v);
  return v === true || v === 1 || n === 'true' || n === 'x' || n === 'ok' || n === 'sim';
}
function _toDateStr_(v) {
  if (!v) return '';
  try {
    const tz = Session.getScriptTimeZone();
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return Utilities.formatDate(v, tz, DATE_FMT_BR);
    }
    const m = String(v).match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (m) {
      const d = new Date(String(m[3]).padStart(4,'20'), +m[2]-1, +m[1]);
      return Utilities.formatDate(d, tz, DATE_FMT_BR);
    }
    const d2 = new Date(v);
    if (!isNaN(d2)) return Utilities.formatDate(d2, tz, DATE_FMT_BR);
  } catch(e){}
  return '';
}

/** ===== PLAN SHEET (Livros & capítulos) ===== */

function _getPlanSheet_() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();
  const forced = props.getProperty('SHEET_NAME');

  if (forced) {
    const sh = ss.getSheetByName(forced);
    if (!sh) throw new Error(`Aba "${forced}" não encontrada. Abas: ${ss.getSheets().map(s=>s.getName()).join(', ')}`);
    _ensurePlanHeaders_(sh);
    return sh;
  }

  const candidates = [];
  for (const sh of ss.getSheets()) {
    const vals = sh.getDataRange().getValues();
    if (!vals.length) continue;
    const top = (vals[0] || []).map(_normalize_);
    const hasLivro = top.some(h => h.includes('livro'));
    const hasCap   = top.some(h => h.includes('capit') || h === 'cap' || h.startsWith('cap'));
    if (hasLivro && hasCap) candidates.push(sh);
  }

  if (candidates.length === 0) {
    const names = ss.getSheets().map(s => s.getName()).join(', ');
    throw new Error(`Nenhuma aba com cabeçalhos "Livro / Capítulo" foi encontrada.
Abas existentes: ${names}
Dica: use a aba "Livros & capítulos" ou defina a propriedade SHEET_NAME com o nome da aba que tem Livro/Capítulo.`);
  }

  const preferred = candidates.find(s => _normalize_(s.getName()).includes('livros'));
  const sh = preferred || candidates[0];
  _ensurePlanHeaders_(sh);
  return sh;
}

/** Garante colunas Lido/Data */
function _ensurePlanHeaders_(sh) {
  const rng = sh.getDataRange();
  const vals = rng.getValues();
  if (!vals.length) throw new Error(`Aba "${sh.getName()}" está vazia.`);

  const header = vals[0].map(String);
  const norm = header.map(_normalize_);

  const colLivro = norm.findIndex(h => h.includes('livro'));
  const colCap   = norm.findIndex(h => h.includes('capit') || h === 'cap' || h.startsWith('cap'));
  if (colLivro < 0 || colCap < 0) {
    throw new Error(`Aba "${sh.getName()}" não possui colunas "Livro" e "Capítulo".`);
  }

  let colLido = norm.findIndex(h => h.includes('lido'));
  let colData = norm.findIndex(h => h.includes('data'));
  let nextCol = header.length + 1;

  if (colLido < 0) {
    sh.getRange(1, nextCol).setValue('Lido');
    nextCol++;
  }
  if (colData < 0) {
    sh.getRange(1, nextCol).setValue('Data');
  }
  SpreadsheetApp.flush();
}

function _getHeaderIdx_(sh) {
  const first = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(_normalize_);
  const idx = (fnArr) => {
    const i = first.findIndex(h => fnArr.some(f => typeof f === 'string' ? h.includes(f) : f(h)));
    if (i < 0) throw new Error('Cabeçalho obrigatório ausente.');
    return i;
  };
  return {
    book: idx(['livro']),
    chap: idx([h => h.includes('capit') || h === 'cap' || h.startsWith('cap')]),
    done: idx(['lido']),
    date: idx(['data'])
  };
}

/** ===== API: Plano ===== */

function getPlan() {
  const sh = _getPlanSheet_();
  const H = _getHeaderIdx_(sh);
  const data = sh.getDataRange().getValues();

  const rows = [];
  const books = new Set();

  for (let r = 2; r <= sh.getLastRow(); r++) {
    const v = data[r-1] || [];
    const livro = String(v[H.book] || '').trim();
    const cap   = Number(v[H.chap] || '');
    if (!livro || !cap) continue;

    const lido = _toBool_(v[H.done]);
    const dataStr = _toDateStr_(v[H.date]);
    books.add(livro);
    rows.push({ i: r, livro, capitulo: cap, lido, data: dataStr });
  }

  return { rows, books: Array.from(books), stats: _stats_(rows), sheetName: sh.getName() };
}

function toggleChapter(rowIndex) {
  const sh = _getPlanSheet_();
  const H = _getHeaderIdx_(sh);

  if (rowIndex < 2 || rowIndex > sh.getLastRow()) throw new Error('Índice de linha inválido.');
  const current = _toBool_(sh.getRange(rowIndex, H.done + 1).getValue());
  const next = !current;

  sh.getRange(rowIndex, H.done + 1).setValue(next);
  sh.getRange(rowIndex, H.date + 1).setValue(next ? new Date() : '');
  SpreadsheetApp.flush();
  return getPlan();
}

function _stats_(rows) {
  const total = rows.length;
  const lidos = rows.filter(r => r.lido).length;
  const pct   = total ? Math.round((lidos/total)*100) : 0;
  return { total, lidos, faltam: total - lidos, pct };
}

/** (Opcional) Fixar aba específica via propriedade */
function setSheetName(name) {
  if (!name) throw new Error('Informe o nome da aba.');
  PropertiesService.getScriptProperties().setProperty('SHEET_NAME', name);
  return getPlan();
}
function clearSheetName() {
  PropertiesService.getScriptProperties().deleteProperty('SHEET_NAME');
  return getPlan();
}

/** ===== ATLAS SHEET (Atlas guia) =====
 * Detecta aba que tenha colunas Livro + (Cap/Capítulo/Capítulos) + uma coluna contendo 'atlas' ou 'seção'
 */
function _getAtlasSheet_() {
  const ss = SpreadsheetApp.getActive();
  const candidates = [];
  for (const sh of ss.getSheets()) {
    const vals = sh.getDataRange().getValues();
    if (!vals.length) continue;
    const head = (vals[0] || []).map(_normalize_);
    const hasLivro = head.some(h => h.includes('livro'));
    const hasCap   = head.some(h => h.includes('cap'));
    const hasAtlas = head.some(h => h.includes('atlas') || h.includes('secao') || h.includes('seção'));
    if (hasLivro && hasCap && hasAtlas) candidates.push(sh);
  }
  if (candidates.length === 0) throw new Error('Não encontrei uma aba de Atlas (com colunas Livro/Capítulo/Seção do Atlas).');
  // preferir nome contendo 'atlas'
  const preferred = candidates.find(s => _normalize_(s.getName()).includes('atlas'));
  return preferred || candidates[0];
}

function _getAtlasHeaders_(sh) {
  const first = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(_normalize_);
  const pos = (arr) => {
    const i = first.findIndex(h => arr.some(k => typeof k==='string' ? h.includes(k) : k(h)));
    if (i < 0) throw new Error('Cabeçalho do Atlas ausente.');
    return i;
  };
  return {
    book: pos(['livro']),
    chap: pos(['capit','cap','capitulo','capítulos','capitulos']),
    atlas: pos(['atlas','secao','seção'])
  };
}

/** API: retorna mapeamento do Atlas para o modal */
function getAtlasMap() {
  const sh = _getAtlasSheet_();
  const H = _getAtlasHeaders_(sh);
  const vals = sh.getDataRange().getValues();

  const rows = [];
  const books = new Set();

  for (let r = 2; r <= sh.getLastRow(); r++) {
    const v = vals[r-1] || [];
    const livro = String(v[H.book] || '').trim();
    const capStr = String(v[H.chap] || '').trim(); // pode ser "1–14", "3-5,7", etc.
    const atlas = String(v[H.atlas] || '').trim();
    if (!livro) continue;
    rows.push({ livro, capitulos: capStr, atlas });
    books.add(livro);
  }

  return { rows, books: Array.from(books), sheetName: sh.getName() };
}
