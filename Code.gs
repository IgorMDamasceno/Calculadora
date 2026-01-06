const USD_BRL_RATE = 5.30;
const SPREADSHEET_ID = '1Bqj1KnaZ14KjQDfZ3RczJ7Z9kRyD4Y73PeERyUsTjvs';

const SHEET_CATS = 'FUNIL_CATEGORIAS';
const SHEET_OPS  = 'FUNIL_OPERACOES';

function doGet() {
  ensureSheets_();
  const t = HtmlService.createTemplateFromFile('Index');
  t.usdRate = USD_BRL_RATE;
  return t.evaluate()
    .setTitle('Funil Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSS_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'COLE_AQUI_O_ID_DA_SUA_PLANILHA') {
    throw new Error('Defina o SPREADSHEET_ID no Code.gs.');
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function apiListCategorias() {
  ensureSheets_();
  const sh = getSS_().getSheetByName(SHEET_CATS);
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1).filter(r => r[0]).map(r => ({
    id: String(r[0]),
    nome: String(r[1] || ''),
    createdAt: String(r[2] || ''),
    updatedAt: String(r[3] || '')
  })).sort((a,b) => a.nome.localeCompare(b.nome, 'pt-BR'));
}

function apiUpsertCategoria(payload) {
  ensureSheets_();
  const nome = String(payload?.nome || '').trim();
  if (!nome) throw new Error('Nome da categoria é obrigatório.');

  const sh = getSS_().getSheetByName(SHEET_CATS);
  const now = new Date().toISOString();
  const id = payload?.id ? String(payload.id) : Utilities.getUuid();

  const row = findRowById_(sh, id);
  if (row) {
    sh.getRange(row, 2, 1, 1).setValue(nome);
    sh.getRange(row, 4, 1, 1).setValue(now);
  } else {
    sh.appendRow([id, nome, now, now]);
  }
  return { id, nome };
}

function apiDeleteCategoria(payload) {
  ensureSheets_();
  const id = String(payload?.id || '');
  if (!id) throw new Error('Categoria inválida.');

  const ss = getSS_();
  const shCats = ss.getSheetByName(SHEET_CATS);
  const row = findRowById_(shCats, id);
  if (row) shCats.deleteRow(row);

  const shOps = ss.getSheetByName(SHEET_OPS);
  const values = shOps.getDataRange().getValues();
  if (values.length <= 1) return true;

  const toDelete = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]) === id) toDelete.push(i + 1);
  }
  toDelete.sort((a,b) => b - a).forEach(r => shOps.deleteRow(r));
  return true;
}

function apiListOperacoes(payload) {
  ensureSheets_();
  const categoriaId = String(payload?.categoriaId || '');
  const sh = getSS_().getSheetByName(SHEET_OPS);
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];

  return values.slice(1)
    .filter(r => r[0] && String(r[1]) === categoriaId)
    .map(r => ({ id: String(r[0]), nome: String(r[2] || '') }))
    .sort((a,b) => a.nome.localeCompare(b.nome, 'pt-BR'));
}

function apiListOperacoesFull(payload) {
  ensureSheets_();
  const categoriaId = String(payload?.categoriaId || '');
  const sh = getSS_().getSheetByName(SHEET_OPS);
  const values = sh.getDataRange().getValues();
  if (values.length <= 1) return [];
  return values.slice(1)
    .filter(r => r[0] && String(r[1]) === categoriaId)
    .map(r => rowToOperacao_(r))
    .sort((a,b) => a.nome.localeCompare(b.nome, 'pt-BR'));
}

function apiGetOperacao(payload) {
  ensureSheets_();
  const id = String(payload?.id || '');
  if (!id) throw new Error('Operação inválida.');

  const sh = getSS_().getSheetByName(SHEET_OPS);
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === id) return rowToOperacao_(values[i]);
  }
  throw new Error('Operação não encontrada.');
}

function apiUpsertOperacao(payload) {
  ensureSheets_();
  const categoriaId = String(payload?.categoriaId || '').trim();
  const nome = String(payload?.nome || '').trim();
  if (!categoriaId) throw new Error('Selecione uma categoria.');
  if (!nome) throw new Error('Nome da operação é obrigatório.');

  const sh = getSS_().getSheetByName(SHEET_OPS);
  const now = new Date().toISOString();
  const id = payload?.id ? String(payload.id) : Utilities.getUuid();

  const investimento = asNumber_(payload?.investimento);
  const cpa = asNumber_(payload?.cpa);
  const roasPercent = asNumber_(payload?.roasPercent);
  const aproveitamentoLeadsPercent = asNumber_(payload?.aproveitamentoLeadsPercent);
  const aproveitamentoCicloInicialPercent = asNumber_(payload?.aproveitamentoCicloInicialPercent);
  const rpsCicloInicialUsd = asNumber_(payload?.rpsCicloInicialUsd);

  const dataRow = [
    id,
    categoriaId,
    nome,
    investimento,
    cpa,
    roasPercent,
    aproveitamentoLeadsPercent,
    aproveitamentoCicloInicialPercent,
    rpsCicloInicialUsd,
    payload?.createdAt ? String(payload.createdAt) : now,
    now
  ];

  const row = findRowById_(sh, id);
  if (row) sh.getRange(row, 1, 1, dataRow.length).setValues([dataRow]);
  else sh.appendRow(dataRow);

  return { id };
}

function apiDeleteOperacao(payload) {
  ensureSheets_();
  const id = String(payload?.id || '');
  if (!id) throw new Error('Operação inválida.');

  const sh = getSS_().getSheetByName(SHEET_OPS);
  const row = findRowById_(sh, id);
  if (row) sh.deleteRow(row);
  return true;
}

function ensureSheets_() {
  const ss = getSS_();

  const catsHeaders = ['id', 'nome', 'createdAt', 'updatedAt'];
  const opsHeaders = [
    'id', 'categoriaId', 'nome',
    'investimento', 'cpa', 'roasPercent',
    'aproveitamentoLeadsPercent', 'aproveitamentoCicloInicialPercent',
    'rpsCicloInicialUsd',
    'createdAt', 'updatedAt'
  ];

  ensureSheetWithHeaders_(ss, SHEET_CATS, catsHeaders);
  ensureSheetWithHeaders_(ss, SHEET_OPS, opsHeaders);

  const cats = ss.getSheetByName(SHEET_CATS);
  if (cats.getLastRow() === 1) {
    const now = new Date().toISOString();
    cats.appendRow([Utilities.getUuid(), 'Estados Unidos', now, now]);
  }
}

function ensureSheetWithHeaders_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const firstRow = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeaders = firstRow.join('||') !== headers.join('||');

  if (sh.getLastRow() === 0 || needsHeaders) {
    sh.clear();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
}

function findRowById_(sheet, id) {
  const last = sheet.getLastRow();
  if (last <= 1) return 0;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i]) === String(id)) return i + 2;
  }
  return 0;
}

function rowToOperacao_(r) {
  return {
    id: String(r[0]),
    categoriaId: String(r[1]),
    nome: String(r[2] || ''),
    investimento: asNumber_(r[3]),
    cpa: asNumber_(r[4]),
    roasPercent: asNumber_(r[5]),
    aproveitamentoLeadsPercent: asNumber_(r[6]),
    aproveitamentoCicloInicialPercent: asNumber_(r[7]),
    rpsCicloInicialUsd: asNumber_(r[8]),
    createdAt: String(r[9] || ''),
    updatedAt: String(r[10] || '')
  };
}

function asNumber_(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
}
