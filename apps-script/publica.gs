// 🌐 BOLÃO COPA 2026 - PLANILHA PÚBLICA (ESPELHO)
// ============================================
// Esta planilha é o espelho READ-ONLY que o dashboard consome via gviz.
// Ela NÃO recebe respostas: puxa tudo da planilha MASTER (privada) por
// QUERY(IMPORTRANGE(...)), expondo só nome + palpites — sem e-mail, sem timestamp
// de quem respondeu, e sem permitir edição.
//
// QUANDO USAR: rode se as fórmulas dos espelhos quebrarem, se as colunas
// saírem do lugar, ou se quiser recriar a planilha pública do zero com o
// layout EXATO que o index.html espera.
//
// ⚠️ LAYOUTS QUE O DASHBOARD EXIGE (não altere sem mexer no HTML também):
//   • Estatisticas: SEM cabeçalho, dados a partir da LINHA 1.
//       A=Nome | B=Unidade | C=J1Bra | D=J1Mar | E=J2Bra | F=J2Hai
//       G=J3Esc | H=J3Bra | I/J=Grupo A ... AE/AF=Grupo L
//   • Ranking: A=# | B=Nome | C=Total | D=Placares | E=Posições | F=Timestamp
//   • Gabarito: espelho 1:1 da master (mantém os rótulos e as linhas).
//
// COMO USAR:
//   1) Cole MASTER_ID abaixo (ID da planilha privada que recebe os Forms).
//   2) Extensões > Apps Script (NA PLANILHA PÚBLICA) > cole este arquivo.
//   3) Rode configurarPublica().
//   4) Abra cada aba e clique em "Permitir acesso" no aviso do IMPORTRANGE
//      (autorização única; o script não consegue clicar por você).
//   5) Compartilhe a planilha como "Qualquer pessoa com o link: Leitor".
// ============================================

// ⬇️ ID da planilha MASTER (privada). Pega da URL:
//    https://docs.google.com/spreadsheets/d/ESTE_PEDACO/edit
var MASTER_ID = PropertiesService.getScriptProperties().getProperty('MASTER_ID');

// Nome da aba de respostas na MASTER (a aba vinculada ao Forms).
var MASTER_RESPOSTAS = 'Respostas';

function configurarPublica() {
  if (!MASTER_ID || MASTER_ID.indexOf('COLE_O_ID') > -1) {
    Logger.log('❌ Preencha MASTER_ID antes de rodar.');
    return;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  espelharEstatisticas_(ss);
  espelharRanking_(ss);
  espelharGabarito_(ss);

  Logger.log('');
  Logger.log('✅ Espelhos configurados.');
  Logger.log('👉 Agora abra cada aba e clique em "Permitir acesso" no aviso do IMPORTRANGE.');
  Logger.log('👉 Depois, opcional: rode protegerPublica() para travar contra edição acidental.');
}

// ===== ESTATISTICAS (sem cabeçalho — começa na linha 1) =====
// Master Respostas: A=Timestamp B=Email C=Nome D=Unidade E..J=jogos K..AH=grupos
// Pegamos C..AH (Col3..Col34) = exatamente Nome..Grupo L, fora e-mail/timestamp.
function espelharEstatisticas_(ss) {
  var sh = recriarAba_(ss, 'Estatisticas');
  var formula =
    '=QUERY(IMPORTRANGE("' + MASTER_ID + '","' + MASTER_RESPOSTAS + '!C2:AH"),' +
    '"select * where Col1 is not null",0)';
  sh.getRange('A1').setFormula(formula);
  Logger.log('  ✅ Estatisticas espelhada (A1, sem cabeçalho)');
}

// ===== RANKING (espelha as linhas de dados da master, sem cabeçalho) =====
// Master Ranking: dados em A4:F53 (#, Nome, Total, Placares, Posições, Timestamp).
function espelharRanking_(ss) {
  var sh = recriarAba_(ss, 'Ranking');
  var formula =
    '=QUERY(IMPORTRANGE("' + MASTER_ID + '","Ranking!A4:F53"),' +
    '"select * where Col2 is not null",0)';
  sh.getRange('A1').setFormula(formula);
  Logger.log('  ✅ Ranking espelhado (A1)');
}

// ===== GABARITO (espelho 1:1 — mantém rótulos e posições) =====
function espelharGabarito_(ss) {
  var sh = recriarAba_(ss, 'Gabarito');
  sh.getRange('A1').setFormula('=IMPORTRANGE("' + MASTER_ID + '","Gabarito!A1:C25")');
  Logger.log('  ✅ Gabarito espelhado (1:1)');
}

// ===== util: recria a aba zerada =====
function recriarAba_(ss, nome) {
  var existing = ss.getSheetByName(nome);
  if (existing) ss.deleteSheet(existing);
  return ss.insertSheet(nome);
}

// ===== PROTEÇÃO (opcional) =====
// Trava as abas para ninguém apagar as fórmulas por engano.
function protegerPublica() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var owner = Session.getEffectiveUser();
  ['Estatisticas', 'Ranking', 'Gabarito'].forEach(function (nome) {
    var sh = ss.getSheetByName(nome);
    if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });
    var prot = sh.protect().setDescription('🔒 Espelho - Bolão Copa 2026');
    prot.addEditor(owner);
    prot.removeEditors(prot.getEditors());
    if (prot.canDomainEdit()) prot.setDomainEdit(false);
    prot.setWarningOnly(false);
    Logger.log('🔒 Aba "' + nome + '" protegida');
  });
  Logger.log('✅ Proteção aplicada.');
}

// ===== DIAGNÓSTICO: confere se os espelhos têm dados =====
function diagnosticarPublica() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ['Estatisticas', 'Ranking', 'Gabarito'].forEach(function (nome) {
    var sh = ss.getSheetByName(nome);
    if (!sh) { Logger.log('❌ Aba ' + nome + ' não existe'); return; }
    var a1 = sh.getRange('A1').getValue();
    var linhas = sh.getDataRange().getNumRows();
    var refErro = String(a1).indexOf('#REF') > -1;
    Logger.log(nome + ': ' + linhas + ' linha(s) | A1="' + a1 + '"' +
      (refErro ? '  ⚠️ IMPORTRANGE sem autorização — clique em "Permitir acesso"' : ''));
  });
}
