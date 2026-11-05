/**
 * Agenda online da Barbearia Dom Carlos (Google Apps Script).
 * Passo a passo:
 * 1. Crie uma planilha no Google Sheets e abra Extensões > Apps Script.
 * 2. Cole este arquivo, salve e escolha Implantar > Nova implantação > App da Web.
 *    Executar como: você. Quem tem acesso: qualquer pessoa.
 * 3. Copie a URL gerada e cole em agenda.url no CONFIG do index.html.
 * A planilha recebe as colunas: Registro | Data | Hora | Barbeiro | Serviços.
 */
const ABA = "Agenda";

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) { sh = ss.insertSheet(ABA); sh.appendRow(["Registro", "Data", "Hora", "Barbeiro", "Serviços"]); }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// GET: lista os horários ocupados de hoje em diante
function doGet() {
  const v = aba_().getDataRange().getValues().slice(1);
  const hoje = Utilities.formatDate(new Date(), "America/Fortaleza", "yyyy-MM-dd");
  const ocupados = v.map(r => ({
    data: Utilities.formatDate(new Date(r[1]), "America/Fortaleza", "yyyy-MM-dd") || String(r[1]),
    hora: String(r[2]), barbeiro: String(r[3])
  })).filter(o => o.data >= hoje);
  return json_({ ocupados: ocupados });
}

// POST: registra o pedido, validando conflito no servidor
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d.data) || !/^\d{2}:00$/.test(d.hora)) return json_({ ok: false, motivo: "invalido" });
    const sh = aba_();
    const rows = sh.getDataRange().getValues().slice(1);
    const fmt = x => Utilities.formatDate(new Date(x), "America/Fortaleza", "yyyy-MM-dd");
    const mesmos = rows.filter(r => fmt(r[1]) === d.data && String(r[2]) === d.hora);
    const cap = Math.max(1, Math.min(Number(d.capacidade) || 1, 5));
    const mesmoBarbeiro = d.barbeiro !== "Sem preferência" && mesmos.some(r => String(r[3]) === d.barbeiro);
    if (mesmos.length >= cap || mesmoBarbeiro) return json_({ ok: false, motivo: "ocupado" });
    sh.appendRow([new Date(), d.data, d.hora, String(d.barbeiro).slice(0, 60), String(d.servicos).slice(0, 200)]);
    return json_({ ok: true });
  } finally { lock.releaseLock(); }
}
