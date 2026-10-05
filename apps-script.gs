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

// Atalho: roda a configuracao completa (abas, estatisticas, analise e Painel). Pode rodar quantas vezes quiser;
// nunca apaga as respostas das abas Monitores e Alunos.
function gerarPainel() {
  configurar();
}

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
  estilizarRespostas(ss);
  painel(ss);
  organizarAbas(ss);
}

// ---- Analise cruzada: escores por tema, correlacoes, medias por grupo e alfa de Cronbach ----
var MAXL = 300; // le as linhas 2 a 301 de cada aba de respostas
var PERFIL = ["Sou monitor(a)", "Já fui monitor(a)", "Nunca fui"];
var USO = ["Usaria", "Talvez usaria", "Não usaria"];
var CURSOS_L = ["Administração", "Ciência da Computação", "Ciência de Dados Aplicada", "Ciências Atmosféricas", "Ciências Biológicas Licenciatura", "Design", "Engenharia Ambiental", "Engenharia Civil", "Engenharia de Bioprocessos", "Engenharia de Computação", "Engenharia de Controle e Automação", "Engenharia de Energia", "Engenharia de Materiais", "Engenharia de Produção", "Engenharia Elétrica", "Engenharia Eletrônica", "Engenharia Hídrica", "Engenharia Mecânica", "Engenharia Mecânica Aeronáutica", "Engenharia Química", "Física Bacharelado", "Física Licenciatura", "Matemática Bacharelado", "Matemática Licenciatura", "Química Bacharelado", "Química Licenciatura", "Sistemas de Informação"];
var PERIODOS_L = ["1º", "2º", "3º", "4º", "5º", "6º", "7º", "8º", "9º", "10º ou mais"];
var ANALISE = {
  monitores: {
    aba: "Monitores", esc: "Escores_Monitores",
    temas: [["Usabilidade", "G", "K"], ["Forum por disciplina", "L", "O"], ["Grupos e pedido de ajuda", "P", "S"], ["Voluntariado", "T", "V"]],
    grupos: [["Curso", "B", CURSOS_L], ["Periodo", "C", PERIODOS_L], ["Relacao com monitoria", "D", PERFIL], ["Usaria na rotina", "AC", USO],
      ["Funcionalidade mais util", "AB", ["Fórum por disciplina", "Busca de dúvidas parecidas", "Painel da monitoria (dúvidas sem resposta)", "Trabalhos em grupo e chat do grupo", "Pedidos de ajuda ao monitor", "Conversa da monitoria com o professor", "Acervo de arquivos", "Notificações", "Voluntariado e certificado"]],
      ["Canal atual das duvidas", "AA", ["WhatsApp", "E-mail", "Pessoalmente", "Fórum ou plataforma da disciplina", "Outro"]],
      ["Testou a plataforma", "AE", ["Testei na plataforma", "Só vi as telas"]]]
  },
  alunos: {
    aba: "Alunos", esc: "Escores_Alunos",
    temas: [["Agregacao da plataforma", "E", "J"], ["Utilidade dos componentes", "K", "R"], ["Primeira impressao", "S", "T"]],
    grupos: [["Curso", "B", CURSOS_L], ["Periodo", "C", PERIODOS_L], ["Canal atual das duvidas", "D", ["Grupo de WhatsApp da turma", "Perguntando ao professor", "Perguntando ao monitor", "Perguntando a colegas", "Pesquisando na internet", "Outro"]],
      ["Usaria nas disciplinas", "V", USO],
      ["Funcionalidade mais util", "U", ["Fórum por disciplina", "Busca de dúvidas parecidas", "Trabalhos em grupo e chat do grupo", "Pedidos de ajuda ao monitor", "Acervo de arquivos", "Notificações", "Voluntariado e certificado"]]]
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

// =====================================================================
// PAINEL VISUAL + ESTILO DAS ABAS
// Nao altera nenhuma coluna nem apaga respostas. Rode configurar() para (re)gerar.
// =====================================================================
var COR = {
  azul: "#12305c", azul2: "#1b4f9c", claro: "#eef3fb", linha: "#d9dfe8",
  texto: "#152033", mudo: "#5b6678", branco: "#ffffff", oliva: "#8a9a00",
  ruim: "#f4c7c3", medio: "#fce8b2", bom: "#b7e1cd"
};

var ROT_MON = {
  us1: "Entendi para que serve a plataforma", us2: "As telas parecem fáceis de usar",
  us3: "Usaria sozinho(a), sem ajuda", us4: "É claro como ir de uma área a outra",
  us5: "Textos e botões fáceis de entender", fo1: "Um fórum por disciplina ajuda a tirar dúvidas",
  fo2: "As funções do fórum fazem sentido", fo3: "Busca de dúvidas parecidas é útil",
  fo4: "Ver dúvidas sem resposta ajuda no meu trabalho", co1: "Chat privado do grupo ajuda a organizar",
  co2: "Receber só o recorte do pedido facilita", co3: "Dúvidas pela plataforma > WhatsApp/e-mail",
  co4: "Canal monitoria-professor facilita", vo1: "Achar e se inscrever é simples",
  vo2: "Informações da oportunidade estão claras", vo3: "Certificado com código é útil"
};
var ROT_ALU = {
  ag1: "Dúvidas da disciplina em um só lugar ajudam a estudar", ag2: "Respostas de colegas, monitores e professores ajudam a aprender",
  ag3: "Achar dúvida parecida já respondida economiza tempo", ag4: "Pedir ajuda ao monitor aqui é melhor que no WhatsApp",
  ag5: "Chat privado para o meu grupo seria útil", ag6: "Voluntariado e certificado me motivariam",
  cp1: "Fórum por disciplina", cp2: "Pedido de ajuda ao monitor", cp3: "Busca de dúvidas parecidas",
  cp4: "Trabalhos em grupo e chat do grupo", cp5: "Acervo de arquivos", cp6: "Notificações",
  cp7: "Oportunidades de voluntariado", cp8: "Certificado com código de validação",
  us1: "Entendi para que serve a plataforma", us2: "As telas parecem fáceis de usar"
};
var UTIL_L = ["Fórum por disciplina", "Busca de dúvidas parecidas", "Painel da monitoria (dúvidas sem resposta)", "Trabalhos em grupo e chat do grupo",
  "Pedidos de ajuda ao monitor", "Conversa da monitoria com o professor", "Acervo de arquivos", "Notificações", "Voluntariado e certificado"];
var CANAIS_M = ["WhatsApp", "E-mail", "Pessoalmente", "Fórum ou plataforma da disciplina", "Outro"];
var CANAIS_A = ["Grupo de WhatsApp da turma", "Perguntando ao professor", "Perguntando ao monitor", "Perguntando a colegas", "Pesquisando na internet", "Outro"];

// n = maior numero de respostas entre as perguntas do bloco (pessoas que responderam ao tema)
function nDoBloco(r) {
  var m = /^(.+!)([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(r);
  if (!m) return "=COUNT(" + r + ")";
  return "=MAX(" + letras(m[2], m[4]).map(function (c) { return "COUNT(" + m[1] + c + m[3] + ":" + c + m[5] + ")"; }).join(";") + ")";
}
function faixa(aba, col) { return aba + "!" + col + "2:" + col + (MAXL + 1); }
function bloco(aba, c1, c2) { return aba + "!" + c1 + "2:" + c2 + (MAXL + 1); }
function colDe(form, campo) { return letra(FORMS[form].campos.indexOf(campo) + 1); }

// ---------------- abas de respostas ----------------
function estilizarRespostas(ss) {
  ["monitores", "alunos"].forEach(function (nome) {
    var cfg = FORMS[nome], sh = ss.getSheetByName(cfg.aba);
    if (!sh) return;
    var n = cfg.campos.length;
    var linhas = Math.max(sh.getMaxRows(), 2);
    var cab = sh.getRange(1, 1, 1, n);
    cab.setBackground(COR.azul).setFontColor(COR.branco).setFontWeight("bold").setFontFamily("Arial")
      .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    sh.setRowHeight(1, 34);
    sh.setFrozenRows(1);
    sh.setTabColor(COR.oliva);
    sh.getBandings().forEach(function (b) { b.remove(); });
    var corpo = sh.getRange(1, 1, linhas, n);
    var b = corpo.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
    b.setHeaderRowColor(COR.azul).setFirstRowColor(COR.branco).setSecondRowColor(COR.claro);
    cab.setFontColor(COR.branco);
    sh.getRange(2, 1, linhas - 1, n).setFontFamily("Arial").setFontSize(10).setVerticalAlignment("top");
    cfg.campos.forEach(function (c, i) {
      var col = i + 1, largo = (c.indexOf("obs_") === 0 || c === "comentario" || c === "falta" || c === "nome");
      if (c === "enviado_em") { sh.setColumnWidth(col, 170); }
      else if (cfg.itens.indexOf(c) >= 0) { sh.setColumnWidth(col, 52); sh.getRange(2, col, linhas - 1, 1).setHorizontalAlignment("center"); }
      else if (largo) { sh.setColumnWidth(col, 280); sh.getRange(2, col, linhas - 1, 1).setWrap(true); }
      else { sh.setColumnWidth(col, 170); }
    });
    // escala de cor nas notas (1 vermelho, 3 amarelo, 5 verde)
    var p = cfg.itens[0], u = cfg.itens[cfg.itens.length - 1];
    var iniC = FORMS[nome].campos.indexOf(p) + 1, fimC = FORMS[nome].campos.indexOf(u) + 1;
    var notas = sh.getRange(2, iniC, linhas - 1, fimC - iniC + 1);
    sh.setConditionalFormatRules([gradiente(notas)]);
  });
}

function gradiente(faixaNotas) {
  return SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue(COR.ruim, SpreadsheetApp.InterpolationType.NUMBER, "1")
    .setGradientMidpointWithValue(COR.medio, SpreadsheetApp.InterpolationType.NUMBER, "3")
    .setGradientMaxpointWithValue(COR.bom, SpreadsheetApp.InterpolationType.NUMBER, "5")
    .setRanges([faixaNotas]).build();
}

// ---------------- blocos do painel ----------------
function titulo(sh, linha, texto) {
  var r = sh.getRange(linha, 2, 1, 12).merge();
  r.setValue(texto).setBackground(COR.azul2).setFontColor(COR.branco).setFontWeight("bold").setFontSize(12)
    .setVerticalAlignment("middle");
  sh.setRowHeight(linha, 30);
}

function cartao(sh, linha, col, rotulo, formula, nota, formato) {
  var area = sh.getRange(linha, col, 3, 2);
  area.setBackground(COR.claro).setBorder(true, true, true, true, false, false, COR.linha, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(linha, col, 1, 2).merge().setValue(rotulo).setFontSize(9).setFontColor(COR.mudo).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("bottom");
  var v = sh.getRange(linha + 1, col, 1, 2).merge();
  v.setFormula(formula).setFontSize(26).setFontWeight("bold").setFontColor(COR.azul2)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  if (formato) v.setNumberFormat(formato);
  sh.getRange(linha + 2, col, 1, 2).merge().setValue(nota).setFontSize(9).setFontColor(COR.mudo)
    .setHorizontalAlignment("center").setVerticalAlignment("top");
}

// tabela de notas: linhas = [[rotulo, faixa]], com cabecalho; devolve a proxima linha livre
function tabelaNotas(sh, linha, cabecalho, linhas, geral) {
  var cab = sh.getRange(linha, 2, 1, 11);
  sh.getRange(linha, 2, 1, 4).merge().setValue(cabecalho);
  sh.getRange(linha, 6).setValue("Média");
  sh.getRange(linha, 7, 1, 4).merge().setValue("Barra (escala 1 a 5)");
  sh.getRange(linha, 11).setValue("Concordância");
  sh.getRange(linha, 12).setValue("n");
  cab.setBackground(COR.azul).setFontColor(COR.branco).setFontWeight("bold").setHorizontalAlignment("center");
  sh.getRange(linha, 2).setHorizontalAlignment("left");
  var r = linha + 1, primeira = r;
  linhas.concat(geral ? [geral] : []).forEach(function (it, i) {
    var ehGeral = geral && i === linhas.length;
    sh.getRange(r, 2, 1, 4).merge().setValue(it[0]).setHorizontalAlignment("left").setWrap(true);
    sh.getRange(r, 6).setFormula("=IFERROR(ROUND(AVERAGE(" + it[1] + ");2);\"-\")").setNumberFormat("0.00");
    sh.getRange(r, 7, 1, 4).merge().setFormula("=IFERROR(REPT(\"█\";ROUND(F" + r + "*4;0));\"\")")
      .setFontColor(COR.azul2).setHorizontalAlignment("left");
    sh.getRange(r, 11).setFormula("=IFERROR(COUNTIF(" + it[1] + ";\">=4\")/COUNT(" + it[1] + ");\"-\")").setNumberFormat("0%");
    sh.getRange(r, 12).setFormula(nDoBloco(it[1]));
    sh.getRange(r, 6, 1, 1).setFontWeight("bold").setHorizontalAlignment("center");
    sh.getRange(r, 11, 1, 2).setHorizontalAlignment("center");
    sh.getRange(r, 2, 1, 11).setBorder(null, null, true, null, null, null, COR.linha, SpreadsheetApp.BorderStyle.SOLID);
    if (ehGeral) sh.getRange(r, 2, 1, 11).setFontWeight("bold").setBackground(COR.claro);
    sh.setRowHeight(r, 24);
    r++;
  });
  var rules = sh.getConditionalFormatRules();
  rules.push(gradiente(sh.getRange(primeira, 6, r - primeira, 1)));
  sh.setConditionalFormatRules(rules);
  return r + 1;
}

// tabela de contagens por categoria (para graficos); devolve {proxima, topo, base}
function tabelaContagem(sh, linha, cabecalho, categorias, series) {
  // series = [[nomeSerie, aba, colunaDaResposta]]
  sh.getRange(linha, 2, 1, 3).merge().setValue(cabecalho);
  series.forEach(function (s, i) { sh.getRange(linha, 5 + i).setValue(s[0]); });
  var cab = sh.getRange(linha, 2, 1, 3 + series.length);
  cab.setBackground(COR.azul).setFontColor(COR.branco).setFontWeight("bold").setHorizontalAlignment("center");
  sh.getRange(linha, 2).setHorizontalAlignment("left");
  categorias.forEach(function (c, i) {
    var r = linha + 1 + i;
    sh.getRange(r, 2, 1, 3).merge().setValue(c).setHorizontalAlignment("left").setWrap(true);
    series.forEach(function (s, j) {
      sh.getRange(r, 5 + j).setFormula("=COUNTIF(" + faixa(s[1], s[2]) + ";" + "\"" + c + "\")").setHorizontalAlignment("center");
    });
    sh.getRange(r, 2, 1, 3 + series.length).setBorder(null, null, true, null, null, null, COR.linha, SpreadsheetApp.BorderStyle.SOLID);
    sh.setRowHeight(r, 24);
  });
  return { topo: linha, base: linha + categorias.length, proxima: linha + categorias.length + 2 };
}

function grafico(sh, tipo, titulo_, t, colRotulo, colsSeries, linhaAncora, cores) {
  var b = sh.newChart().setChartType(tipo)
    .addRange(sh.getRange(t.topo, colRotulo, t.base - t.topo + 1, 1));
  colsSeries.forEach(function (c) { b.addRange(sh.getRange(t.topo, c, t.base - t.topo + 1, 1)); });
  b.setNumHeaders(1)
    .setOption("title", titulo_)
    .setOption("titleTextStyle", { color: COR.texto, fontSize: 12, bold: true })
    .setOption("colors", cores)
    .setOption("legend", { position: colsSeries.length > 1 ? "bottom" : "none" })
    .setOption("width", 470).setOption("height", Math.max(220, (t.base - t.topo + 2) * 26 + 70))
    .setOption("chartArea", tipo === Charts.ChartType.BAR ? { left: 200, width: "48%", height: "68%" } : { width: "72%", height: "62%" })
    .setPosition(linhaAncora, 9, 10, 4);
  sh.insertChart(b.build());
}

// ---------------- painel ----------------
function painel(ss) {
  var sh = ss.getSheetByName("Painel");
  if (sh) { sh.getCharts().forEach(function (c) { sh.removeChart(c); }); sh.clear(); sh.clearConditionalFormatRules(); sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); }
  else sh = ss.insertSheet("Painel", 0);
  sh.setHiddenGridlines(true);
  sh.setTabColor(COR.azul2);
  if (sh.getMaxColumns() < 14) sh.insertColumnsAfter(sh.getMaxColumns(), 14 - sh.getMaxColumns());
  sh.setColumnWidth(1, 20);
  for (var c = 2; c <= 13; c++) sh.setColumnWidth(c, 96);
  sh.getRange(1, 1, 400, 14).setFontFamily("Arial").setFontColor(COR.texto).setVerticalAlignment("middle");

  // cabecalho
  sh.getRange(1, 2, 1, 12).merge().setValue("Avaliação da Plataforma UNIFEI")
    .setBackground(COR.azul).setFontColor(COR.branco).setFontSize(22).setFontWeight("bold").setVerticalAlignment("middle");
  sh.setRowHeight(1, 54);
  sh.getRange(2, 2, 1, 7).merge().setValue("Painel de resultados · monitores e alunos · atualiza sozinho a cada resposta")
    .setBackground(COR.azul).setFontColor("#c9d7ee").setFontSize(10).setVerticalAlignment("top");
  sh.getRange(2, 9, 1, 3).merge().setValue("última resposta em").setBackground(COR.azul).setFontColor("#c9d7ee")
    .setFontSize(10).setHorizontalAlignment("right").setVerticalAlignment("top");
  var dM = "ARRAYFORMULA(IFERROR(DATEVALUE(LEFT(" + faixa("Monitores", "A") + ";10));0))";
  var dA = "ARRAYFORMULA(IFERROR(DATEVALUE(LEFT(" + faixa("Alunos", "A") + ";10));0))";
  sh.getRange(2, 12, 1, 2).merge().setFormula("=IFERROR(IF(MAX(" + dM + ";" + dA + ")=0;\"-\";MAX(" + dM + ";" + dA + "));\"-\")")
    .setNumberFormat("dd/MM/yyyy").setBackground(COR.azul).setFontColor(COR.branco).setFontSize(10)
    .setFontWeight("bold").setHorizontalAlignment("left").setVerticalAlignment("top");
  sh.setRowHeight(2, 26);
  sh.setRowHeight(3, 10);

  var M = "Monitores", A = "Alunos";
  var gM = bloco(M, "G", "V"), gA = bloco(A, "E", "T");
  var uM = faixa(M, "AC"), uA = faixa(A, "V");

  // cartoes
  [4, 8].forEach(function (l) { sh.setRowHeight(l, 22); sh.setRowHeight(l + 1, 40); sh.setRowHeight(l + 2, 22); });
  cartao(sh, 4, 2, "RESPOSTAS · MONITORES", "=COUNTA(" + faixa(M, "A") + ")", "formulário dos monitores", "0");
  cartao(sh, 4, 4, "RESPOSTAS · ALUNOS", "=COUNTA(" + faixa(A, "A") + ")", "formulário dos alunos", "0");
  cartao(sh, 4, 6, "NOTA GERAL · MONITORES", "=IFERROR(ROUND(AVERAGE(" + gM + ");2);\"-\")", "média dos itens, de 1 a 5", "0.00");
  cartao(sh, 4, 8, "NOTA GERAL · ALUNOS", "=IFERROR(ROUND(AVERAGE(" + gA + ");2);\"-\")", "média dos itens, de 1 a 5", "0.00");
  cartao(sh, 4, 10, "CONCORDÂNCIA · MONITORES", "=IFERROR(COUNTIF(" + gM + ";\">=4\")/COUNT(" + gM + ");\"-\")", "itens com nota 4 ou 5", "0%");
  cartao(sh, 4, 12, "CONCORDÂNCIA · ALUNOS", "=IFERROR(COUNTIF(" + gA + ";\">=4\")/COUNT(" + gA + ");\"-\")", "itens com nota 4 ou 5", "0%");
  sh.setRowHeight(7, 10);
  cartao(sh, 8, 2, "MONITORES QUE USARIAM", "=IFERROR(COUNTIF(" + uM + ";\"Usaria\")/COUNTA(" + uM + ");\"-\")", "usaria na rotina", "0%");
  cartao(sh, 8, 4, "ALUNOS QUE USARIAM", "=IFERROR(COUNTIF(" + uA + ";\"Usaria\")/COUNTA(" + uA + ");\"-\")", "usaria nas disciplinas", "0%");
  cartao(sh, 8, 6, "MONITORES QUE TESTARAM", "=IFERROR(COUNTIF(" + faixa(M, "AE") + ";\"Testei na plataforma\")/COUNTA(" + faixa(M, "AE") + ");\"-\")", "usaram o login de demonstração", "0%");
  cartao(sh, 8, 8, "MONITORES ATUAIS", "=COUNTIF(" + faixa(M, "D") + ";\"Sou monitor(a)\")", "entre os respondentes", "0");
  cartao(sh, 8, 10, "TOTAL DE RESPOSTAS", "=B5+D5", "monitores + alunos", "0");
  cartao(sh, 8, 12, "NOTA GERAL · TODOS", "=IFERROR(ROUND((SUM(" + gM + ")+SUM(" + gA + "))/(COUNT(" + gM + ")+COUNT(" + gA + "));2);\"-\")", "média ponderada por item", "0.00");
  sh.setRowHeight(11, 12);

  // notas
  var l = 12;
  titulo(sh, l++, "MONITORES · nota média por tema (1 = discordo totalmente · 5 = concordo totalmente)");
  var temasM = ANALISE.monitores.temas.map(function (t) { return [t[0].replace("Forum", "Fórum"), bloco(M, t[1], t[2])]; });
  l = tabelaNotas(sh, l, "Tema", temasM, ["Geral", gM]);
  titulo(sh, l++, "ALUNOS · nota média por tema");
  var temasA = ANALISE.alunos.temas.map(function (t) { return [t[0].replace("Agregacao", "Agregação").replace("Primeira impressao", "Primeira impressão"), bloco(A, t[1], t[2])]; });
  l = tabelaNotas(sh, l, "Tema", temasA, ["Geral", gA]);

  titulo(sh, l++, "MONITORES · nota por pergunta");
  var itM = FORMS.monitores.itens.map(function (k) { return [ROT_MON[k], faixa(M, colDe("monitores", k))]; });
  l = tabelaNotas(sh, l, "Pergunta", itM, null);
  titulo(sh, l++, "ALUNOS · nota por pergunta");
  var itA = FORMS.alunos.itens.map(function (k) { return [ROT_ALU[k], faixa(A, colDe("alunos", k))]; });
  l = tabelaNotas(sh, l, "Pergunta", itA, null);

  // distribuicoes + graficos
  titulo(sh, l++, "COMO RESPONDERAM · quem usaria, o que seria mais útil e como as dúvidas chegam hoje");
  var USO3 = ["Usaria", "Talvez usaria", "Não usaria"];
  var t1 = tabelaContagem(sh, l, "Usaria a plataforma?", USO3, [["Monitores", M, "AC"], ["Alunos", A, "V"]]);
  grafico(sh, Charts.ChartType.COLUMN, "Usaria a plataforma?", t1, 2, [5, 6], l, [COR.azul2, COR.oliva]);
  l = Math.max(t1.proxima, l + 11);
  var t2 = tabelaContagem(sh, l, "Funcionalidade mais útil", UTIL_L, [["Monitores", M, "AB"], ["Alunos", A, "U"]]);
  grafico(sh, Charts.ChartType.BAR, "Funcionalidade mais útil", t2, 2, [5, 6], l, [COR.azul2, COR.oliva]);
  l = Math.max(t2.proxima, l + 16);
  var t3 = tabelaContagem(sh, l, "Por onde as dúvidas chegam hoje (monitores)", CANAIS_M, [["Monitores", M, "AA"]]);
  grafico(sh, Charts.ChartType.BAR, "Canal atual das dúvidas · monitores", t3, 2, [5], l, [COR.azul2]);
  l = Math.max(t3.proxima, l + 12);
  var t4 = tabelaContagem(sh, l, "Como tiram dúvidas hoje (alunos)", CANAIS_A, [["Alunos", A, "D"]]);
  grafico(sh, Charts.ChartType.BAR, "Canal atual das dúvidas · alunos", t4, 2, [5], l, [COR.oliva]);
  l = Math.max(t4.proxima, l + 12);
  var t5 = tabelaContagem(sh, l, "Período do curso", PERIODOS_L, [["Monitores", M, "C"], ["Alunos", A, "C"]]);
  grafico(sh, Charts.ChartType.COLUMN, "Respostas por período", t5, 2, [5, 6], l, [COR.azul2, COR.oliva]);
  l = Math.max(t5.proxima, l + 17);

  sh.getRange(l, 2, 1, 12).merge().setValue("Fontes: abas Monitores e Alunos (respostas), Estatisticas e Analise (cálculos detalhados). Este painel é recalculado sozinho a cada nova resposta.")
    .setFontSize(9).setFontColor(COR.mudo).setWrap(true);
  sh.setFrozenRows(2);
  sh.setActiveRange(sh.getRange("B4"));
}

function organizarAbas(ss) {
  var ordem = ["Painel", "Monitores", "Alunos", "Estatisticas", "Analise", "Escores_Monitores", "Escores_Alunos"];
  ordem.forEach(function (nome, i) {
    var s = ss.getSheetByName(nome);
    if (!s) return;
    ss.setActiveSheet(s);
    ss.moveActiveSheet(i + 1);
    if (nome === "Estatisticas" || nome === "Analise") s.setTabColor("#8d99ae");
    if (nome.indexOf("Escores") === 0) s.setTabColor("#c5cbd6");
  });
  ss.setActiveSheet(ss.getSheetByName("Painel"));
}
