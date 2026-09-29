// Google Apps Script: recebe as respostas do formulário e grava uma linha por resposta na planilha.
const CAMPOS = [
  "enviado_em", "curso", "periodo", "perfil", "identificacao", "nome",
  "us1", "us2", "us3", "us4", "us5",
  "fo1", "fo2", "fo3", "fo4",
  "co1", "co2", "co3", "co4",
  "vo1", "vo2", "vo3",
  "obs_us", "obs_fo", "obs_co", "obs_vo",
  "canal_duvidas", "mais_util", "uso", "comentario"
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const dados = JSON.parse(e.postData.contents);
    const aba = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (aba.getLastRow() === 0) aba.appendRow(CAMPOS);
    aba.appendRow(CAMPOS.map(function (c) { return dados[c] === undefined ? "" : dados[c]; }));
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}
