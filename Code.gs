/**
 * Backend Apontamento de Fundações — DC Pecém / SS&B
 *
 * Fluxo:
 *   1. Operador envia (action: "submit") → grava na aba PENDENTES
 *   2. Admin lista pendentes (action: "list_pending") com idToken Google
 *   3. Admin aprova (action: "approve") → dados migram para LANÇAMENTO
 *   4. Admin rejeita (action: "reject") → marca como REJEITADO em PENDENTES
 *
 * Implantação:
 *   - Implantar → Nova implantação → Aplicativo Web
 *   - Executar como: Eu (sua conta)
 *   - Quem tem acesso: Qualquer pessoa
 *   - Cole a URL no HTML (constante WEB_APP_URL)
 */

// ===== CONFIGURAÇÃO ===== ⚠ AJUSTE ESTES VALORES ANTES DE USAR
const ADMIN_EMAILS = [
  "leandro@exemplo.com",       // ← seu email Google
  "gestor@exemplo.com",        // ← email do gestor da obra
];

// ID OAuth do Google Cloud Console — DEVE bater com OAUTH_CLIENT_ID no HTML
// Crie em: https://console.cloud.google.com/apis/credentials
const OAUTH_CLIENT_ID = "SEU_CLIENT_ID.apps.googleusercontent.com";
// =========================

const SHEET_NAME = "LANÇAMENTO";
const PENDING_SHEET_NAME = "PENDENTES";
const HEADER_ROW = 1;

const STAGE_COLUMNS = {
  ESCAVACAO:   { idCol: 2,  dateCol: 3  },
  MAGRO:       { idCol: 8,  dateCol: 9  },
  ARMADURA:    { idCol: 14, dateCol: 15 },
  FORMAS:      { idCol: 20, dateCol: 21 },
  CONCRETAGEM: { idCol: 26, dateCol: 27 },
  REATERRO:    { idCol: 32, dateCol: 33 },
  REATERRO_T:  { idCol: 38, dateCol: 39 },
};

// ===== ROTEADOR =====
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || "submit";
    if (action === "submit")       return handleSubmit(payload);
    if (action === "list_pending") return handleListPending(payload);
    if (action === "approve")      return handleApprove(payload);
    if (action === "reject")       return handleReject(payload);
    throw new Error("Ação desconhecida: " + action);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doGet(e) {
  return jsonResponse({
    ok: true,
    message: "Backend ativo. Use POST com action=submit|list_pending|approve|reject."
  });
}

// ===== HANDLERS =====

function handleSubmit(payload) {
  const { data, lancamentos, operador } = payload;
  if (!data) throw new Error("Data não informada");
  if (!operador || !String(operador).trim()) throw new Error("Operador não informado");
  if (!lancamentos) throw new Error("Sem lançamentos");

  let total = 0;
  const stages = [];
  Object.keys(lancamentos).forEach(k => {
    const n = (lancamentos[k] || []).length;
    if (n > 0) { total += n; stages.push(k + ": " + n); }
  });
  if (total === 0) throw new Error("Nenhum pilar com status concluído");

  const id = "S" + Date.now() + "-" + Math.floor(Math.random() * 1000);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PENDING_SHEET_NAME);
  if (!sheet) sheet = createPendingSheet(ss);

  sheet.appendRow([
    id,
    new Date(),
    data,
    String(operador).trim(),
    total,
    stages.join(", "),
    "PENDENTE",
    "", "", "",
    JSON.stringify(lancamentos)
  ]);

  return jsonResponse({ ok: true, id: id, total: total });
}

function handleListPending(payload) {
  const email = verifyAdmin(payload.idToken);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(PENDING_SHEET_NAME);
  if (!sheet) return jsonResponse({ ok: true, admin: email, items: [] });

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ ok: true, admin: email, items: [] });

  const rows = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  const items = rows
    .filter(r => String(r[6]) === "PENDENTE")
    .map(r => ({
      id: String(r[0]),
      timestamp: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
      data: r[2] instanceof Date
        ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), "yyyy-MM-dd")
        : String(r[2]),
      operador: String(r[3]),
      total: Number(r[4]) || 0,
      resumo: String(r[5]),
      payload: r[10] ? JSON.parse(String(r[10])) : {}
    }))
    .sort((a, b) => a.timestamp.localeCompare(b.timestamp));

  return jsonResponse({ ok: true, admin: email, items: items });
}

function handleApprove(payload) {
  const email = verifyAdmin(payload.idToken);
  const { id, observacao } = payload;
  if (!id) throw new Error("ID não informado");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pendingSheet = ss.getSheetByName(PENDING_SHEET_NAME);
  if (!pendingSheet) throw new Error("Aba PENDENTES não encontrada");

  const rowNumber = findRowById(pendingSheet, id);
  if (rowNumber < 0) throw new Error("Submissão " + id + " não encontrada");

  const status = String(pendingSheet.getRange(rowNumber, 7).getValue());
  if (status !== "PENDENTE") throw new Error("Já processada como " + status);

  const dataRaw = pendingSheet.getRange(rowNumber, 3).getValue();
  const payloadJson = String(pendingSheet.getRange(rowNumber, 11).getValue());
  const lancamentos = JSON.parse(payloadJson);

  writeToLancamento(ss, dataRaw, lancamentos);

  pendingSheet.getRange(rowNumber, 7).setValue("APROVADO");
  pendingSheet.getRange(rowNumber, 8).setValue(email);
  pendingSheet.getRange(rowNumber, 9).setValue(new Date());
  if (observacao) pendingSheet.getRange(rowNumber, 10).setValue(observacao);

  return jsonResponse({ ok: true });
}

function handleReject(payload) {
  const email = verifyAdmin(payload.idToken);
  const { id, observacao } = payload;
  if (!id) throw new Error("ID não informado");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pendingSheet = ss.getSheetByName(PENDING_SHEET_NAME);
  if (!pendingSheet) throw new Error("Aba PENDENTES não encontrada");

  const rowNumber = findRowById(pendingSheet, id);
  if (rowNumber < 0) throw new Error("Submissão " + id + " não encontrada");

  const status = String(pendingSheet.getRange(rowNumber, 7).getValue());
  if (status !== "PENDENTE") throw new Error("Já processada como " + status);

  pendingSheet.getRange(rowNumber, 7).setValue("REJEITADO");
  pendingSheet.getRange(rowNumber, 8).setValue(email);
  pendingSheet.getRange(rowNumber, 9).setValue(new Date());
  if (observacao) pendingSheet.getRange(rowNumber, 10).setValue(observacao);

  return jsonResponse({ ok: true });
}

// ===== AUTORIZAÇÃO =====

function verifyAdmin(idToken) {
  if (!idToken) throw new Error("Token de autenticação ausente");
  if (!OAUTH_CLIENT_ID || OAUTH_CLIENT_ID.indexOf("SEU_CLIENT_ID") === 0) {
    throw new Error("OAUTH_CLIENT_ID não configurado no Apps Script");
  }
  const url = "https://oauth2.googleapis.com/tokeninfo?id_token=" + encodeURIComponent(idToken);
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) throw new Error("Token inválido (HTTP " + resp.getResponseCode() + ")");
  const info = JSON.parse(resp.getContentText());
  if (info.aud !== OAUTH_CLIENT_ID) throw new Error("Token de outra aplicação (audience mismatch)");
  if (parseInt(info.exp, 10) < Date.now() / 1000) throw new Error("Token expirado, faça login novamente");
  const email = String(info.email || "").toLowerCase();
  const allowed = ADMIN_EMAILS.map(x => String(x).toLowerCase());
  if (allowed.indexOf(email) === -1) {
    throw new Error("Email " + email + " não tem permissão de admin");
  }
  return email;
}

// ===== UTILITÁRIOS =====

function createPendingSheet(ss) {
  const sheet = ss.insertSheet(PENDING_SHEET_NAME);
  const headers = [
    "ID", "Timestamp Envio", "Data Execução", "Operador", "Total Pilares",
    "Resumo Etapas", "Status", "Aprovado Por", "Aprovado Em", "Observação Admin",
    "Payload JSON"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold").setBackground("#0E0F10").setFontColor("#FBFAF6");
  sheet.setFrozenRows(1);
  const widths = [120, 160, 110, 150, 90, 220, 100, 220, 160, 220, 320];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  return sheet;
}

function findRowById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const idx = ids.findIndex(x => String(x) === String(id));
  return idx >= 0 ? idx + 2 : -1;
}

function writeToLancamento(ss, dataRaw, lancamentos) {
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Aba '" + SHEET_NAME + "' não encontrada");
  const dateObj = parseDate(dataRaw);

  Object.keys(STAGE_COLUMNS).forEach(stageKey => {
    const ids = (lancamentos[stageKey] || []).filter(x => x && String(x).trim().length > 0);
    if (ids.length === 0) return;
    const cols = STAGE_COLUMNS[stageKey];
    const firstEmptyRow = findFirstEmptyRow(sheet, cols.idCol);
    const idValues = ids.map(id => [String(id).toUpperCase().trim()]);
    const dateValues = ids.map(() => [dateObj]);
    sheet.getRange(firstEmptyRow, cols.idCol, ids.length, 1).setValues(idValues);
    const dateRange = sheet.getRange(firstEmptyRow, cols.dateCol, ids.length, 1);
    dateRange.setValues(dateValues);
    dateRange.setNumberFormat("dd/mm/yyyy");
  });
}

function findFirstEmptyRow(sheet, columnIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < HEADER_ROW + 1) return HEADER_ROW + 1;
  const values = sheet.getRange(HEADER_ROW + 1, columnIndex, lastRow - HEADER_ROW, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === "" || values[i][0] === null) return HEADER_ROW + 1 + i;
  }
  return lastRow + 1;
}

function parseDate(raw) {
  if (raw instanceof Date) return raw;
  const s = String(raw).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
  throw new Error("Formato de data inválido: " + s);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== TESTES (rode no editor) =====

function testSubmit() {
  const e = { postData: { contents: JSON.stringify({
    action: "submit",
    data: "2026-05-15",
    operador: "Teste",
    lancamentos: { ESCAVACAO: ["P999", "P998"] }
  })}};
  Logger.log(doPost(e).getContent());
}

function testListPending() {
  // Para testar autorizar, precisa de um idToken Google válido — gere via o app HTML
  // ou pelo Google Playground. Aqui apenas validar a chamada genérica:
  Logger.log("Configure ADMIN_EMAILS e OAUTH_CLIENT_ID, depois teste pelo app HTML.");
}
