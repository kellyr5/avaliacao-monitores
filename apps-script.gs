// Google Apps Script: recebe as respostas dos formulários e grava uma linha por resposta.
// Cada formulário vai para a sua própria aba. Rode configurar() uma vez para criar as abas e a aba de estatísticas.
var PLANILHA_ID = "1o_9G0b1Een8qLrXMjcuufIu7TsB4tuTNvRwonq-Va94";
var FORMS = {
  monitores: {
    aba: "Monitores",
    campos: ["enviado_em", "curso", "periodo", "perfil", "identificacao", "nome",
      "us1", "us2", "us3", "us4", "us5", "fo1", "fo2", "fo3", "fo4", "co1", "co2", "co3", "co4", "vo1", "vo2", "vo3",
      "obs_us", "obs_fo", "obs_co", "obs_vo", "canal_duvidas", "mais_util", "uso", "comentario"],
    itens: ["us1", "us2", "us3", "us4", "us5", "fo1", "fo2", "fo3", "fo4", "co1", "co2", "co3", "co4", "vo1", "vo2", "vo3"]
  },
  alunos: {
    aba: "Alunos",
    campos: ["enviado_em", "curso", "periodo", "canal_duvidas",
      "ag1", "ag2", "ag3", "ag4", "ag5", "ag6", "cp1", "cp2", "cp3", "cp4", "cp5", "cp6", "cp7", "cp8", "us1", "us2",
      "mais_util", "uso", "falta", "comentario"],
    itens: ["ag1", "ag2", "ag3", "ag4", "ag5", "ag6", "cp1", "cp2", "cp3", "cp4", "cp5", "cp6", "cp7", "cp8", "us1", "us2"]
  }
};

function abaDe(ss, cfg) {
  var aba = ss.getSheetByName(cfg.aba);
  if (!aba) aba = ss.insertSheet(cfg.aba);
  if (aba.getLastRow() === 0) aba.appendRow(cfg.campos);
  return aba;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var dados = JSON.parse(e.postData.contents);
    var cfg = FORMS[dados.formulario];
    if (!cfg) return ContentService.createTextOutput("formulario desconhecido");
    var aba = abaDe(SpreadsheetApp.openById(PLANILHA_ID), cfg);
    aba.appendRow(cfg.campos.map(function (c) {
      var v = dados[c] === undefined ? "" : dados[c];
      return cfg.itens.indexOf(c) >= 0 && v !== "" ? Number(v) : v;
    }));
    return ContentService.createTextOutput("ok");
  } finally {
    lock.releaseLock();
  }
}

function letra(n) { // 1 -> A
  var s = "";
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - m) / 26); }
  return s;
}

function configurar() {
  var ss = SpreadsheetApp.openById(PLANILHA_ID);
  var est = ss.getSheetByName("Estatisticas");
  if (est) est.clear(); else est = ss.insertSheet("Estatisticas");
  var linha = 1;
  ["monitores", "alunos"].forEach(function (nome) {
    var cfg = FORMS[nome];
    abaDe(ss, cfg);
    est.getRange(linha, 1).setValue(cfg.aba).setFontWeight("bold").setFontSize(13);
    est.getRange(linha, 3).setValue("Respostas:");
    est.getRange(linha, 4).setFormula("=COUNTA(" + cfg.aba + "!A2:A)");
    linha++;
    est.getRange(linha, 1, 1, 6).setValues([["Item", "Media", "Mediana", "Desvio padrao", "Concordancia (4-5)", "Respostas"]]).setFontWeight("bold");
    linha++;
    cfg.itens.forEach(function (k) {
      var col = letra(cfg.campos.indexOf(k) + 1);
      var r = cfg.aba + "!" + col + "2:" + col;
      est.getRange(linha, 1).setValue(k);
      est.getRange(linha, 2).setFormula('=IFERROR(AVERAGE(' + r + '),"-")').setNumberFormat("0.00");
      est.getRange(linha, 3).setFormula('=IFERROR(MEDIAN(' + r + '),"-")').setNumberFormat("0.0");
      est.getRange(linha, 4).setFormula('=IFERROR(STDEV(' + r + '),"-")').setNumberFormat("0.00");
      est.getRange(linha, 5).setFormula('=IFERROR(COUNTIF(' + r + ',">=4")/COUNT(' + r + '),"-")').setNumberFormat("0%");
      est.getRange(linha, 6).setFormula("=COUNT(" + r + ")");
      linha++;
    });
    linha++;
    ["uso", "mais_util", "canal_duvidas", "curso"].forEach(function (campo) {
      var col = letra(cfg.campos.indexOf(campo) + 1);
      est.getRange(linha, 1).setValue("Distribuicao: " + campo).setFontWeight("bold");
      est.getRange(linha, 2).setFormula('=IFERROR(QUERY(' + cfg.aba + "!" + col + '2:' + col + ',"select Col1, count(Col1) where Col1 is not null group by Col1 label Col1 \'\', count(Col1) \'\'",0),"sem respostas")');
      linha += 9;
    });
    linha++;
  });
  est.setColumnWidth(1, 220);
  est.autoResizeColumns(2, 5);
}
