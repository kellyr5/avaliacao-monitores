// Google Apps Script: recebe as respostas dos formulários e grava uma linha por resposta.
// Cada formulário vai para a sua própria aba. Rode configurar() uma vez para criar as abas e a aba de estatísticas.
var PLANILHA_ID = "1o_9G0b1Een8qLrXMjcuufIu7TsB4tuTNvRwonq-Va94";
var FORMS = {
  monitores: {
    aba: "Monitores",
    campos: ["enviado_em", "curso", "periodo", "perfil", "identificacao", "nome",
      "us1", "us2", "us3", "us4", "us5", "fo1", "fo2", "fo3", "fo4", "co1", "co2", "co3", "co4", "vo1", "vo2", "vo3",
      "obs_us", "obs_fo", "obs_co", "obs_vo", "canal_duvidas", "mais_util", "uso", "comentario", "testou"],
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
    abaDe(ss, cfg).getRange(1, 1, 1, cfg.campos.length).setValues([cfg.campos]);
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
      est.getRange(linha, 2).setFormula('=IFERROR(AVERAGE(' + r + ');"-")').setNumberFormat("0.00");
      est.getRange(linha, 3).setFormula('=IFERROR(MEDIAN(' + r + ');"-")').setNumberFormat("0.0");
      est.getRange(linha, 4).setFormula('=IFERROR(STDEV(' + r + ');"-")').setNumberFormat("0.00");
      est.getRange(linha, 5).setFormula('=IFERROR(COUNTIF(' + r + ';">=4")/COUNT(' + r + ');"-")').setNumberFormat("0%");
      est.getRange(linha, 6).setFormula("=COUNT(" + r + ")");
      linha++;
    });
    linha++;
    ["uso", "mais_util", "canal_duvidas", "curso"].forEach(function (campo) {
      var col = letra(cfg.campos.indexOf(campo) + 1);
      est.getRange(linha, 1).setValue("Distribuicao: " + campo).setFontWeight("bold");
      est.getRange(linha, 2).setFormula('=IFERROR(QUERY(' + cfg.aba + "!" + col + '2:' + col + ';"select Col1, count(Col1) where Col1 is not null group by Col1 label Col1 \'\', count(Col1) \'\'";0);"sem respostas")');
      linha += 9;
    });
    linha++;
  });
  est.setColumnWidth(1, 220);
  est.autoResizeColumns(2, 5);
  analisar();
}

// ---- Analise cruzada: escores por tema, correlacoes, medias por grupo e alfa de Cronbach ----
var MAXL = 300; // le as linhas 2 a 301 de cada aba de respostas
var PERFIL = ["Sou monitor(a)", "Já fui monitor(a)", "Nunca fui"];
var USO = ["Usaria", "Talvez usaria", "Não usaria"];
var CURSOS_L = ["Ciência da Computação", "Engenharia de Computação", "Sistemas de Informação", "Engenharia Elétrica", "Engenharia Mecânica", "Engenharia de Produção", "Engenharia Civil", "Engenharia Ambiental", "Engenharia Hídrica", "Engenharia Química", "Engenharia de Materiais", "Engenharia de Controle e Automação", "Engenharia de Energia", "Física", "Matemática", "Química", "Outro"];
var PERIODOS_L = ["1º", "2º", "3º", "4º", "5º", "6º", "7º", "8º", "9º", "10º ou mais"];
var ANALISE = {
  monitores: {
    aba: "Monitores", esc: "Escores_Monitores",
    temas: [["Usabilidade", "G", "K"], ["Forum e monitoria", "L", "O"], ["Grupos, chats e duvidas", "P", "S"], ["Voluntariado", "T", "V"]],
    grupos: [["Curso", "B", CURSOS_L], ["Periodo", "C", PERIODOS_L], ["Relacao com monitoria", "D", PERFIL], ["Usaria na rotina", "AC", USO],
      ["Funcionalidade mais util", "AB", ["Chat privado de grupo", "Chat monitor e professor", "Fórum por disciplina", "Pedidos de ajuda e monitoria", "Busca de dúvidas parecidas", "Trabalhos em grupo", "Voluntariado e certificado", "Acervo de arquivos"]],
      ["Canal atual das duvidas", "AA", ["WhatsApp", "E-mail", "Pessoalmente", "Fórum ou plataforma da disciplina", "Outro"]],
      ["Testou a plataforma", "AE", ["Testei na plataforma", "Só vi as telas"]]]
  },
  alunos: {
    aba: "Alunos", esc: "Escores_Alunos",
    temas: [["Agregacao da plataforma", "E", "J"], ["Utilidade dos componentes", "K", "R"], ["Primeira impressao", "S", "T"]],
    grupos: [["Curso", "B", CURSOS_L], ["Periodo", "C", PERIODOS_L], ["Canal atual das duvidas", "D", ["Grupo de WhatsApp da turma", "Perguntando ao professor", "Perguntando ao monitor", "Perguntando a colegas", "Pesquisando na internet", "Outro"]],
      ["Usaria nas disciplinas", "V", USO],
      ["Funcionalidade mais util", "U", ["Fórum por disciplina", "Busca de dúvidas parecidas", "Pedidos de ajuda ao monitor", "Chat privado do grupo", "Acervo de arquivos", "Notificações", "Voluntariado e certificado"]]]
  }
};
var COMPONENTES = ["Forum por disciplina", "Pedidos de ajuda ao monitor", "Busca de duvidas parecidas", "Chat privado do grupo", "Acervo de arquivos", "Notificacoes", "Oportunidades de voluntariado", "Certificados com codigo de validacao"];

function letras(de, ate) { // "G","K" -> [G,H,I,J,K]
  var a = [], i, n1 = 0, n2 = 0;
  for (i = 0; i < de.length; i++) n1 = n1 * 26 + de.charCodeAt(i) - 64;
  for (i = 0; i < ate.length; i++) n2 = n2 * 26 + ate.charCodeAt(i) - 64;
  for (i = n1; i <= n2; i++) a.push(letra(i));
  return a;
}

function escoresDe(ss, a) {
  var sh = ss.getSheetByName(a.esc);
  if (sh) sh.clear(); else sh = ss.insertSheet(a.esc);
  var cab = ["n"];
  a.temas.forEach(function (t) { cab.push(t[0]); });
  cab.push("Geral");
  a.grupos.forEach(function (g) { cab.push(g[0]); });
  var rows = [cab];
  var T = a.temas.length;
  for (var r = 2; r <= MAXL + 1; r++) {
    var l = ['=IF(' + a.aba + '!A' + r + '="";"";ROW()-1)'];
    a.temas.forEach(function (t) { l.push('=IF($A' + r + '="";"";AVERAGE(' + a.aba + '!' + t[1] + r + ':' + t[2] + r + '))'); });
    l.push('=IF($A' + r + '="";"";AVERAGE(' + letra(2) + r + ':' + letra(1 + T) + r + '))');
    a.grupos.forEach(function (g) { l.push('=IF($A' + r + '="";"";' + a.aba + '!' + g[1] + r + ')'); });
    rows.push(l);
  }
  sh.getRange(1, 1, rows.length, cab.length).setValues(rows);
  sh.getRange(1, 1, 1, cab.length).setFontWeight("bold");
}

function analisar() {
  var ss = SpreadsheetApp.openById(PLANILHA_ID);
  var an = ss.getSheetByName("Analise");
  if (an) an.clear(); else an = ss.insertSheet("Analise");
  var L = [];
  var R = function (x) { return 'IFERROR(ROUND(' + x + ';2);"-")'; };
  ["monitores", "alunos"].forEach(function (nome) {
    var a = ANALISE[nome], T = a.temas.length;
    escoresDe(ss, a);
    var rg = function (col) { return a.esc + '!' + col + '2:' + col + (MAXL + 1); };
    var ab = function (col) { return a.aba + '!' + col + '2:' + col + (MAXL + 1); };
    var colTema = function (i) { return letra(2 + i); };
    var colGeral = letra(2 + T);
    L.push(["ANALISE - " + a.aba.toUpperCase()]);
    L.push(["Escala 1 a 5. Cada escore de tema e a media dos itens do respondente. Alfa de Cronbach acima de 0,70 indica consistencia aceitavel."]);
    L.push(["Tema", "n", "Media", "Mediana", "Desvio padrao", "% escore >= 4", "Alfa de Cronbach"]);
    a.temas.forEach(function (t, i) {
      var c = rg(colTema(i)), its = letras(t[1], t[2]), k = its.length;
      var somaVar = its.map(function (x) { return 'VAR(' + ab(x) + ')'; }).join('+');
      L.push([t[0], '=COUNT(' + c + ')', '=' + R('AVERAGE(' + c + ')'), '=' + R('MEDIAN(' + c + ')'), '=' + R('STDEV(' + c + ')'),
        '=' + R('100*COUNTIF(' + c + ';">=4")/COUNT(' + c + ')'), '=' + R(k + '/(' + k + '-1)*(1-(' + somaVar + ')/(' + (k * k) + '*VAR(' + c + ')))')]);
    });
    var cg = rg(colGeral);
    L.push(["Geral (media dos temas)", '=COUNT(' + cg + ')', '=' + R('AVERAGE(' + cg + ')'), '=' + R('MEDIAN(' + cg + ')'), '=' + R('STDEV(' + cg + ')'), '=' + R('100*COUNTIF(' + cg + ';">=4")/COUNT(' + cg + ')'), ""]);
    L.push([""]);
    L.push(["Correlacao de Pearson entre temas (-1 a 1; com poucas respostas e so indicativa)"]);
    L.push([""].concat(a.temas.map(function (t) { return t[0]; })));
    a.temas.forEach(function (t, i) {
      L.push([t[0]].concat(a.temas.map(function (u, j) { return '=' + R('CORREL(' + rg(colTema(i)) + ';' + rg(colTema(j)) + ')'); })));
    });
    L.push([""]);
    a.grupos.forEach(function (g, gi) {
      var gc = rg(letra(3 + T + gi));
      L.push(["Media por: " + g[0], "n"].concat(a.temas.map(function (t) { return t[0]; })).concat(["Geral"]));
      g[2].forEach(function (cat) {
        var row = L.length + 1, ref = '$A' + row;
        var l = [cat, '=COUNTIF(' + gc + ';' + ref + ')'];
        a.temas.forEach(function (t, i) { l.push('=' + R('AVERAGEIFS(' + rg(colTema(i)) + ';' + gc + ';' + ref + ')')); });
        l.push('=' + R('AVERAGEIFS(' + cg + ';' + gc + ';' + ref + ')'));
        L.push(l);
      });
      L.push([""]);
    });
    if (nome === "alunos") {
      L.push(["Ranking dos componentes (utilidade 1 a 5)", "Media", "% nota >= 4", "Posicao"]);
      var ini = L.length + 1;
      letras("K", "R").forEach(function (x, i) {
        var row = L.length + 1;
        L.push([COMPONENTES[i], '=' + R('AVERAGE(' + ab(x) + ')'), '=' + R('100*COUNTIF(' + ab(x) + ';">=4")/COUNT(' + ab(x) + ')'), '=IFERROR(RANK(B' + row + ';B$' + ini + ':B$' + (ini + 7) + ');"-")']);
      });
      L.push([""]);
    }
    L.push([""]);
  });
  L.push(["COMPARATIVO MONITORES x ALUNOS (itens de primeira impressao)", "Monitores", "Alunos", "Diferenca"]);
  [["Entendi para que serve a plataforma", "G", "S"], ["As telas parecem faceis de usar", "H", "T"]].forEach(function (x) {
    var row = L.length + 1;
    L.push([x[0], '=' + R('AVERAGE(' + ANALISE.monitores.aba + '!' + x[1] + '2:' + x[1] + (MAXL + 1) + ')'), '=' + R('AVERAGE(' + ANALISE.alunos.aba + '!' + x[2] + '2:' + x[2] + (MAXL + 1) + ')'), '=IFERROR(ROUND(B' + row + '-C' + row + ';2);"-")']);
  });
  var w = 0;
  L.forEach(function (l) { if (l.length > w) w = l.length; });
  L = L.map(function (l) { while (l.length < w) l.push(""); return l; });
  an.getRange(1, 1, L.length, w).setValues(L);
  an.setColumnWidth(1, 300);
  var titulos = ["ANALISE", "Media por", "Tema", "Correlacao", "Ranking", "COMPARATIVO"];
  L.forEach(function (l, i) {
    if (titulos.some(function (t) { return String(l[0]).indexOf(t) === 0; })) an.getRange(i + 1, 1, 1, w).setFontWeight("bold");
  });
}
