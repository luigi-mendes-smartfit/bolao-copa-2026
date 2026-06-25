// 🏆 BOLÃO COPA DO MUNDO 2026 - SMART FIT
// ============================================
// Script exclusivo para utilização no google sheets na planilha principal gerada pelo forms
// COMO USAR:
// 1) Delete as abas: Pontuação, Ranking, Controle Pagamento (se existirem)
// 2) Mantenha as abas: Respostas (do Forms) e Gabarito (se já estiver ok)
// 3) Vá em: Extensões > Apps Script
// 4) Cole ESTE script inteiro (substitua o anterior)
// 5) Execute: configurarPlanilha()
// ============================================

function configurarPlanilha() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Identificar aba de respostas
  var sheets = ss.getSheets();
  var respostasName = null;
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf('Respostas') > -1 || name.indexOf('Form') > -1 || name.indexOf('Response') > -1) {
      respostasName = name;
      break;
    }
  }
  if (!respostasName) {
    respostasName = sheets[0].getName();
  }

  Logger.log('📋 Aba de respostas detectada: ' + respostasName);

  criarAbaGabarito_(ss);
  criarAbaPontuacao_(ss, respostasName);
  criarAbaRanking_(ss);
  criarAbaRegras_(ss);
  criarAbaControlePagamento_(ss, respostasName);

  Logger.log('');
  Logger.log('✅ Planilha configurada com sucesso!');
  Logger.log('📝 Abas criadas: Gabarito, Pontuação, Ranking, Regras, Controle Pagamento');
  Logger.log('');
  Logger.log('👉 Próximos passos:');
  Logger.log('   1) criarGatilho() → auto-refresh a cada resposta');
  Logger.log('   2) criarGatilhoPagamento() → verificação diária de pgto após 22/06');
  Logger.log('   3) corrigirTimestamp() → formata datas');
  Logger.log('   4) protegerAbas() → trava abas para auditoria');
}

// ===== GATILHO AUTOMÁTICO =====
function criarGatilho() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();

  Logger.log('✅ Gatilho criado! A cada nova resposta, a planilha atualizará automaticamente.');
}

function onFormSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ptsSheet = ss.getSheetByName('Pontuação');
  if (ptsSheet) {
    SpreadsheetApp.flush();
  }
  Logger.log('📬 Nova resposta recebida! Pontuação atualizada.');
}

// ===== ABA GABARITO =====
function criarAbaGabarito_(ss) {
  var existing = ss.getSheetByName('Gabarito');
  if (existing) ss.deleteSheet(existing);

  var gab = ss.insertSheet('Gabarito');

  gab.getRange('A1:C1').merge().setValue('📝 GABARITO — Preencha com os resultados reais')
    .setBackground('#000000').setFontColor('#FFD100').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  gab.getRange('A2').setValue('Preencha após cada rodada. A aba Pontuação recalcula automaticamente.')
    .setFontStyle('italic').setFontColor('#666666');

  gab.getRange('A4:C4').setValues([['Jogo', 'Gols Time 1', 'Gols Time 2']])
    .setBackground('#006400').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  gab.getRange('A5').setValue('Brasil x Marrocos');
  gab.getRange('A6').setValue('Brasil x Haiti');
  gab.getRange('A7').setValue('Escócia x Brasil');
  gab.getRange('B5:C7').setBackground('#E8FFE8').setHorizontalAlignment('center');

  var rule0a10 = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 10).setAllowInvalid(false).build();
  gab.getRange('B5:C7').setDataValidation(rule0a10);

  gab.getRange('A9:C9').setValues([['Grupo', '1º Colocado Real', '2º Colocado Real']])
    .setBackground('#006400').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');

  var grupos = {
    'A': ['🇿🇦 África do Sul', '🇰🇷 Coreia do Sul', '🇲🇽 México', '🇨🇿 Tchéquia'],
    'B': ['🇧🇦 Bósnia e Herz.', '🇨🇦 Canadá', '🇶🇦 Catar', '🇨🇭 Suíça'],
    'C': ['🇧🇷 Brasil', '🏴󠁧󠁢󠁳󠁣󠁴󠁿 Escócia', '🇭🇹 Haiti', '🇲🇦 Marrocos'],
    'D': ['🇦🇺 Austrália', '🇺🇸 Estados Unidos', '🇵🇾 Paraguai', '🇹🇷 Turquía'],
    'E': ['🇩🇪 Alemanha', '🇨🇮 Costa do Marfim', '🇨🇼 Curaçao', '🇪🇨 Equador'],
    'F': ['🇳🇱 Holanda', '🇯🇵 Japão', '🇸🇪 Suécia', '🇹🇳 Tunísia'],
    'G': ['🇧🇪 Bélgica', '🇪🇬 Egito', '🇮🇷 Irã', '🇳🇿 Nova Zelândia'],
    'H': ['🇸🇦 Arábia Saudita', '🇨🇻 Cabo Verde', '🇪🇸 Espanha', '🇺🇾 Uruguai'],
    'I': ['🇫🇷 França', '🇮🇶 Iraque', '🇳🇴 Noruega', '🇸🇳 Senegal'],
    'J': ['🇩🇿 Argélia', '🇦🇷 Argentina', '🇦🇹 Áustria', '🇯🇴 Jordânia'],
    'K': ['🇨🇴 Colômbia', '🇵🇹 Portugal', '🇨🇩 Rep. Dem. Congo', '🇺🇿 Uzbequistão'],
    'L': ['🇭🇷 Croácia', '🇬🇭 Gana', '🏴󠁧󠁢󠁥󠁮󠁧󠁿 Inglaterra', '🇵🇦 Panamá']
  };

  var grupoKeys = Object.keys(grupos);
  for (var i = 0; i < grupoKeys.length; i++) {
    var grp = grupoKeys[i];
    var teams = grupos[grp];
    var row = 10 + i;
    gab.getRange(row, 1).setValue('Grupo ' + grp);
    gab.getRange(row, 2).setBackground('#E8FFE8').setHorizontalAlignment('center');
    gab.getRange(row, 3).setBackground('#E8FFE8').setHorizontalAlignment('center');

    var ruleTeams = SpreadsheetApp.newDataValidation()
      .requireValueInList(teams, true).setAllowInvalid(false).build();
    gab.getRange(row, 2).setDataValidation(ruleTeams);
    gab.getRange(row, 3).setDataValidation(ruleTeams);
  }

  gab.setColumnWidth(1, 100);
  gab.setColumnWidth(2, 180);
  gab.setColumnWidth(3, 180);

  Logger.log('  ✅ Aba Gabarito criada');
}

// ===== ABA PONTUAÇÃO (INDIRECT — IMUNE A INSERÇÃO DE LINHAS) =====
function criarAbaPontuacao_(ss, respostasName) {
  var existing = ss.getSheetByName('Pontuação');
  if (existing) ss.deleteSheet(existing);

  var pts = ss.insertSheet('Pontuação');
  var R = respostasName;

  // Header
  pts.getRange('A1').setValue('🏆 PONTUAÇÃO AUTOMÁTICA')
    .setBackground('#000000').setFontColor('#FFD100').setFontWeight('bold').setFontSize(14);
  pts.getRange('A2').setValue('Tudo automático! Preencha apenas o Gabarito após os jogos.')
    .setFontStyle('italic').setFontColor('#666666');

  // Column headers
  var headers = ['Nome', 'Jogo 1', 'Jogo 2', 'Jogo 3', 'Sub Jogos',
    'Grp A', 'Grp B', 'Grp C', 'Grp D', 'Grp E', 'Grp F',
    'Grp G', 'Grp H', 'Grp I', 'Grp J', 'Grp K', 'Grp L',
    'Sub Grupos', 'TOTAL', 'Placares Exatos', 'Posições Cravadas', 'Timestamp'];

  pts.getRange(3, 1, 1, headers.length).setValues([headers])
    .setBackground('#FFD100').setFontWeight('bold').setHorizontalAlignment('center').setFontSize(9);

  // Colunas dos grupos na aba Respostas
  var grpCols = [
    ['K','L'],  // A
    ['M','N'],  // B
    ['O','P'],  // C
    ['Q','R'],  // D
    ['S','T'],  // E
    ['U','V'],  // F
    ['W','X'],  // G
    ['Y','Z'],  // H
    ['AA','AB'], // I
    ['AC','AD'], // J
    ['AE','AF'], // K
    ['AG','AH']  // L
  ];

  var NUM = 50;

  for (var p = 0; p < NUM; p++) {
    var row = p + 4;
    var rr = p + 2;

    // INDIRECT() constrói a referência como string → NUNCA se desloca
    // Formato: INDIRECT("'Form_Responses'!C2")

    // Nome (Col C)
    pts.getRange(row, 1).setFormula(
      '=IFERROR(INDIRECT("\'" & "' + R + '" & "\'!C' + rr + '"),"")'
    );

    // Jogo 1: E,F vs Gabarito $B$5,$C$5
    pts.getRange(row, 2).setFormula(
      '=IF(OR(Gabarito!$B$5="",INDIRECT("\'" & "' + R + '" & "\'!E' + rr + '")=""),"",IF(AND(INDIRECT("\'" & "' + R + '" & "\'!E' + rr + '")=Gabarito!$B$5,INDIRECT("\'" & "' + R + '" & "\'!F' + rr + '")=Gabarito!$C$5),5,IF(SIGN(INDIRECT("\'" & "' + R + '" & "\'!E' + rr + '")-INDIRECT("\'" & "' + R + '" & "\'!F' + rr + '"))=SIGN(Gabarito!$B$5-Gabarito!$C$5),3,0)))'
    );

    // Jogo 2: G,H vs Gabarito $B$6,$C$6
    pts.getRange(row, 3).setFormula(
      '=IF(OR(Gabarito!$B$6="",INDIRECT("\'" & "' + R + '" & "\'!G' + rr + '")=""),"",IF(AND(INDIRECT("\'" & "' + R + '" & "\'!G' + rr + '")=Gabarito!$B$6,INDIRECT("\'" & "' + R + '" & "\'!H' + rr + '")=Gabarito!$C$6),5,IF(SIGN(INDIRECT("\'" & "' + R + '" & "\'!G' + rr + '")-INDIRECT("\'" & "' + R + '" & "\'!H' + rr + '"))=SIGN(Gabarito!$B$6-Gabarito!$C$6),3,0)))'
    );

    // Jogo 3: I,J vs Gabarito $B$7,$C$7
    pts.getRange(row, 4).setFormula(
      '=IF(OR(Gabarito!$B$7="",INDIRECT("\'" & "' + R + '" & "\'!I' + rr + '")=""),"",IF(AND(INDIRECT("\'" & "' + R + '" & "\'!I' + rr + '")=Gabarito!$B$7,INDIRECT("\'" & "' + R + '" & "\'!J' + rr + '")=Gabarito!$C$7),5,IF(SIGN(INDIRECT("\'" & "' + R + '" & "\'!I' + rr + '")-INDIRECT("\'" & "' + R + '" & "\'!J' + rr + '"))=SIGN(Gabarito!$B$7-Gabarito!$C$7),3,0)))'
    );

    // Sub Jogos (referências internas — não se deslocam)
    pts.getRange(row, 5).setFormula('=IF(A' + row + '="","",IFERROR(B' + row + '+C' + row + '+D' + row + ',0))');

    // Grupos A-L
    for (var g = 0; g < 12; g++) {
      var gabRow = 10 + g;
      var col1 = grpCols[g][0];
      var col2 = grpCols[g][1];

      pts.getRange(row, 6 + g).setFormula(
        '=IF(OR(Gabarito!$B$' + gabRow + '="",INDIRECT("\'" & "' + R + '" & "\'!' + col1 + rr + '")=""),"",IF(INDIRECT("\'" & "' + R + '" & "\'!' + col1 + rr + '")=INDIRECT("\'" & "' + R + '" & "\'!' + col2 + rr + '"),0,IF(AND(INDIRECT("\'" & "' + R + '" & "\'!' + col1 + rr + '")=Gabarito!$B$' + gabRow + ',INDIRECT("\'" & "' + R + '" & "\'!' + col2 + rr + '")=Gabarito!$C$' + gabRow + '),5,IF(AND(INDIRECT("\'" & "' + R + '" & "\'!' + col1 + rr + '")=Gabarito!$C$' + gabRow + ',INDIRECT("\'" & "' + R + '" & "\'!' + col2 + rr + '")=Gabarito!$B$' + gabRow + '),3,IF(OR(INDIRECT("\'" & "' + R + '" & "\'!' + col1 + rr + '")=Gabarito!$B$' + gabRow + ',INDIRECT("\'" & "' + R + '" & "\'!' + col2 + rr + '")=Gabarito!$C$' + gabRow + '),2,IF(OR(INDIRECT("\'" & "' + R + '" & "\'!' + col1 + rr + '")=Gabarito!$C$' + gabRow + ',INDIRECT("\'" & "' + R + '" & "\'!' + col2 + rr + '")=Gabarito!$B$' + gabRow + '),1,0))))))'
      );
    }

    // Sub Grupos
    pts.getRange(row, 18).setFormula('=IF(A' + row + '="","",IFERROR(SUM(F' + row + ':Q' + row + '),0))');

    // TOTAL
    pts.getRange(row, 19).setFormula('=IF(A' + row + '="","",IFERROR(E' + row + '+R' + row + ',0))');

    // Desempate 1: Placares exatos
    pts.getRange(row, 20).setFormula('=IF(A' + row + '="","",COUNTIF(B' + row + ':D' + row + ',5))');

    // Desempate 2: Posições cravadas
    pts.getRange(row, 21).setFormula('=IF(A' + row + '="","",COUNTIF(F' + row + ':Q' + row + ',5))');

    // Desempate 3: Timestamp
    pts.getRange(row, 22).setFormula(
      '=IFERROR(INDIRECT("\'" & "' + R + '" & "\'!A' + rr + '"),"")'
    );
    pts.getRange(row, 22).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }

  // Formatting
  pts.getRange('S4:S53').setBackground('#FFD100').setFontWeight('bold');
  pts.getRange('T4:U53').setBackground('#E6F3FF');
  pts.setColumnWidth(1, 180);
  for (var c = 2; c <= 18; c++) pts.setColumnWidth(c, 65);
  pts.setColumnWidth(19, 75);
  pts.setColumnWidth(20, 100);
  pts.setColumnWidth(21, 110);
  pts.setColumnWidth(22, 140);

  pts.setFrozenRows(3);

  Logger.log('  ✅ Aba Pontuação criada (INDIRECT — imune a inserção de linhas)');
}

// ===== ABA RANKING =====
function criarAbaRanking_(ss) {
  var existing = ss.getSheetByName('Ranking');
  if (existing) ss.deleteSheet(existing);

  var rank = ss.insertSheet('Ranking');

  rank.getRange('A1:F1').merge().setValue('🏅 RANKING FINAL')
    .setBackground('#000000').setFontColor('#FFD100').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  rank.getRange('A2:F2').merge()
    .setValue('Ordene: TOTAL (Z→A) → Placares Exatos (Z→A) → Posições Cravadas (Z→A) → Timestamp (A→Z)')
    .setFontStyle('italic').setFontColor('#666666').setFontSize(9);

  rank.getRange('A3:F3').setValues([['#', 'Nome', 'TOTAL', 'Placares\nExatos', 'Posições\nCravadas', 'Timestamp']])
    .setBackground('#FFD100').setFontWeight('bold').setHorizontalAlignment('center');

  // Ranking usa INDIRECT para puxar da Pontuação (também imune)
  for (var p = 0; p < 50; p++) {
    var row = p + 4;
    var ptsRow = p + 4;
    rank.getRange(row, 1).setFormula('=IF(INDIRECT("Pontuação!A' + ptsRow + '")="","",ROW()-3)');
    rank.getRange(row, 2).setFormula('=IFERROR(INDIRECT("Pontuação!A' + ptsRow + '"),"")');
    rank.getRange(row, 3).setFormula('=IFERROR(INDIRECT("Pontuação!S' + ptsRow + '"),"")');
    rank.getRange(row, 4).setFormula('=IFERROR(INDIRECT("Pontuação!T' + ptsRow + '"),"")');
    rank.getRange(row, 5).setFormula('=IFERROR(INDIRECT("Pontuação!U' + ptsRow + '"),"")');
    rank.getRange(row, 6).setFormula('=IFERROR(INDIRECT("Pontuação!V' + ptsRow + '"),"")');
    rank.getRange(row, 6).setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }

  rank.getRange('C4:C53').setBackground('#FFD100').setFontWeight('bold');
  rank.getRange('D4:E53').setBackground('#E6F3FF');
  rank.setColumnWidth(1, 35);
  rank.setColumnWidth(2, 200);
  rank.setColumnWidth(3, 70);
  rank.setColumnWidth(4, 100);
  rank.setColumnWidth(5, 110);
  rank.setColumnWidth(6, 160);
  rank.setFrozenRows(3);

  Logger.log('  ✅ Aba Ranking criada');
}

// ===== ABA REGRAS =====
function criarAbaRegras_(ss) {
  var existing = ss.getSheetByName('Regras');
  if (existing) ss.deleteSheet(existing);

  var regras = ss.insertSheet('Regras');

  regras.getRange('A1').setValue('🏆 BOLÃO COPA DO MUNDO 2026 - SMART FIT')
    .setBackground('#000000').setFontColor('#FFD100').setFontWeight('bold').setFontSize(16);

  var rules = [
    [''],
    ['📋 SISTEMA DE PONTUAÇÃO'],
    [''],
    ['⚽ PLACARES DOS JOGOS DO BRASIL:'],
    ['   • Acertou o placar exato → 5 pontos'],
    ['   • Acertou vencedor/empate (errou placar) → 3 pontos'],
    ['   • Errou quem ganha → 0 pontos'],
    [''],
    ['🌍 CLASSIFICAÇÃO DOS GRUPOS (1º e 2º colocado):'],
    ['   • Acertou 1º E 2º na ordem correta → 5 pontos'],
    ['   • Acertou os 2 classificados, errou a ordem → 3 pontos'],
    ['   • Acertou só 1 classificado na posição correta → 2 pontos'],
    ['   • Acertou só 1 classificado na posição incorreta → 1 ponto'],
    ['   • Errou ambos → 0 pontos'],
    ['   ⚠️ Mesmo país em 1º e 2º → pontuação ZERADA nesse grupo!'],
    [''],
    ['📊 PONTUAÇÃO MÁXIMA: 75 pontos (Jogos: 15 + Grupos: 60)'],
    [''],
    ['🏆 PREMIAÇÃO:'],
    ['   🥇 1º Colocado → 50% da premiação'],
    ['   🥈 2º Colocado → 30% da premiação'],
    ['   🥉 3º Colocado → 20% da premiação'],
    [''],
    ['⚖️ CRITÉRIOS DE DESEMPATE:'],
    ['   1º) Maior quantidade de placares exatos dos jogos do Brasil'],
    ['   2º) Quem cravou mais posições dos classificados (5 pts)'],
    ['   3º) Quem enviou o formulário primeiro (timestamp mais antigo)'],
    [''],
    ['📅 JOGOS DO BRASIL (Grupo C):'],
    ['   • 13/06 (Sáb) 19h — Brasil x Marrocos'],
    ['   • 19/06 (Sex) 22h — Brasil x Haiti'],
    ['   • 24/06 (Qua) 19h — Escócia x Brasil'],
    [''],
    ['⏰ PRAZO: até 12/06 às 16h']
  ];

  regras.getRange(3, 1, rules.length, 1).setValues(rules);
  regras.setColumnWidth(1, 500);

  Logger.log('  ✅ Aba Regras criada');
}

// ===== ABA CONTROLE DE PAGAMENTO =====
function criarAbaControlePagamento_(ss, respostasName) {
  var existing = ss.getSheetByName('Controle Pagamento');
  if (existing) ss.deleteSheet(existing);

  var pgto = ss.insertSheet('Controle Pagamento');
  var R = respostasName;

  pgto.getRange('A1:E1').merge().setValue('💰 CONTROLE DE PAGAMENTO — Bolão Copa 2026')
    .setBackground('#000000').setFontColor('#FFD100').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  pgto.getRange('A2:E2').merge()
    .setValue('Prazo: 22/06/2026. Após essa data, status muda para ELIMINADO automaticamente.')
    .setFontStyle('italic').setFontColor('#666666').setFontSize(9);

  pgto.getRange('A3:E3').setValues([['Nome', 'Email', 'Status Pagamento', 'Data Pagamento', 'Observação']])
    .setBackground('#FFD100').setFontWeight('bold').setHorizontalAlignment('center');

  for (var p = 0; p < 50; p++) {
    var row = p + 4;
    var rr = p + 2;

    // INDIRECT — imune a inserção de linhas
    pgto.getRange(row, 1).setFormula('=IFERROR(INDIRECT("\'" & "' + R + '" & "\'!C' + rr + '"),"")');
    pgto.getRange(row, 2).setFormula('=IFERROR(INDIRECT("\'" & "' + R + '" & "\'!B' + rr + '"),"")');
    pgto.getRange(row, 3).setValue('');
    pgto.getRange(row, 4).setValue('');
    pgto.getRange(row, 5).setValue('');
  }

  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pago', 'Pendente', 'Eliminado'], true)
    .setAllowInvalid(false).build();
  pgto.getRange('C4:C53').setDataValidation(statusRule);
  pgto.getRange('D4:D53').setNumberFormat('dd/MM/yyyy');

  // Conditional formatting
  var rules = pgto.getConditionalFormatRules();

  var rulePago = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pago').setBackground('#C6EFCE').setFontColor('#006400')
    .setRanges([pgto.getRange('C4:C53')]).build();
  rules.push(rulePago);

  var rulePendente = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pendente').setBackground('#FFEB9C').setFontColor('#9C5700')
    .setRanges([pgto.getRange('C4:C53')]).build();
  rules.push(rulePendente);

  var ruleEliminado = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Eliminado').setBackground('#FFC7CE').setFontColor('#9C0006')
    .setRanges([pgto.getRange('C4:C53')]).build();
  rules.push(ruleEliminado);

  pgto.setConditionalFormatRules(rules);

  pgto.setColumnWidth(1, 200);
  pgto.setColumnWidth(2, 220);
  pgto.setColumnWidth(3, 140);
  pgto.setColumnWidth(4, 120);
  pgto.setColumnWidth(5, 200);
  pgto.setFrozenRows(3);

  Logger.log('  ✅ Aba Controle Pagamento criada');
}

// ===== VERIFICAR PRAZO DE PAGAMENTO =====
function verificarPrazoPagamento() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pgto = ss.getSheetByName('Controle Pagamento');
  if (!pgto) { Logger.log('❌ Aba "Controle Pagamento" não encontrada'); return; }

  var prazo = new Date('2026-06-22T23:59:59');
  var hoje = new Date();

  if (hoje < prazo) {
    Logger.log('⏰ Prazo ainda não venceu (22/06/2026).');
    return;
  }

  var lastRow = pgto.getLastRow();
  var eliminados = 0;

  for (var row = 4; row <= lastRow; row++) {
    var nome = pgto.getRange(row, 1).getValue();
    var status = pgto.getRange(row, 3).getValue();

    if (nome && (status === '' || status === 'Pendente')) {
      pgto.getRange(row, 3).setValue('Eliminado');
      pgto.getRange(row, 5).setValue('Não pagou até 22/06 - eliminado automaticamente');
      eliminados++;
    }
  }

  Logger.log('✅ ' + eliminados + ' participante(s) eliminado(s) por falta de pagamento.');
}

function criarGatilhoPagamento() {
  ScriptApp.newTrigger('verificarPrazoPagamento')
    .timeBased().atHour(0).everyDays(1).create();
  Logger.log('✅ Gatilho diário criado para verificar pagamentos após 22/06.');
}

// ===== CORRIGIR TIMESTAMP =====
function corrigirTimestamp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pts = ss.getSheetByName('Pontuação');
  if (pts) pts.getRange('V4:V53').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  var rank = ss.getSheetByName('Ranking');
  if (rank) rank.getRange('F4:F53').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  Logger.log('✅ Formato de timestamp corrigido.');
}

// ===== PROTEÇÃO DE ABAS =====
function protegerAbas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var owner = Session.getEffectiveUser();
  var abasParaProteger = ['Pontuação', 'Ranking'];

  for (var i = 0; i < abasParaProteger.length; i++) {
    var sheet = ss.getSheetByName(abasParaProteger[i]);
    if (sheet) {
      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (var j = 0; j < protections.length; j++) protections[j].remove();

      var protection = sheet.protect().setDescription('🔒 Protegido - Bolão Copa 2026');
      protection.addEditor(owner);
      protection.removeEditors(protection.getEditors());
      protection.setWarningOnly(false);
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
      Logger.log('🔒 Aba "' + abasParaProteger[i] + '" protegida');
    }
  }

  // Proteger aba de Respostas
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf('Respostas') > -1 || name.indexOf('Form') > -1 || name.indexOf('Response') > -1) {
      var protections = sheets[i].getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (var j = 0; j < protections.length; j++) protections[j].remove();
      var protection = sheets[i].protect().setDescription('🔒 Respostas - Não editar!');
      protection.addEditor(owner);
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
      protection.setWarningOnly(false);
      Logger.log('🔒 Aba "' + name + '" protegida');
    }
  }
  Logger.log('✅ Todas as abas protegidas!');
}

function desprotegerAbas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var protections = sheets[i].getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var j = 0; j < protections.length; j++) protections[j].remove();
  }
  Logger.log('🔓 Todas as proteções removidas.');
}

// ===== API FOOTBALL-DATA.ORG =====
var API_BASE = 'https://api.football-data.org/v4';

function SETUP_API_KEY() {
  var MINHA_KEY = '';  // ← COLE SEU TOKEN AQUI
  if (!MINHA_KEY) { Logger.log('❌ Cole sua API key!'); return; }
  PropertiesService.getScriptProperties().setProperty('FOOTBALL_DATA_KEY', MINHA_KEY);
  Logger.log('✅ API Key salva! Execute: testarAPI()');
}

function testarAPI() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('FOOTBALL_DATA_KEY');
  if (!apiKey) { Logger.log('❌ Execute SETUP_API_KEY() primeiro.'); return; }

  var url = API_BASE + '/competitions/WC';
  var response = UrlFetchApp.fetch(url, { headers: { 'X-Auth-Token': apiKey }, muteHttpExceptions: true });
  var data = JSON.parse(response.getContentText());

  if (data.name) {
    Logger.log('✅ Conexão OK! Competição: ' + data.name);
    Logger.log('📅 Temporada: ' + (data.currentSeason ? data.currentSeason.startDate + ' a ' + data.currentSeason.endDate : 'N/A'));
  } else {
    Logger.log('❌ Erro: ' + JSON.stringify(data));
  }
}

// ===== DIAGNÓSTICO DE STANDINGS =====
// Roda isto quando quiser ver EXATAMENTE o que a API devolve:
// HTTP, season, qtd de grupos, o valor literal do campo "group" e o top-2 de cada grupo.
function diagnosticarStandings() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('FOOTBALL_DATA_KEY');
  if (!apiKey) { Logger.log('❌ Sem API key'); return; }

  var resp = UrlFetchApp.fetch(API_BASE + '/competitions/WC/standings',
    { headers: { 'X-Auth-Token': apiKey }, muteHttpExceptions: true });

  Logger.log('HTTP ' + resp.getResponseCode());
  var data = JSON.parse(resp.getContentText());
  Logger.log('Season: ' + (data.season ? data.season.startDate + ' → ' + data.season.endDate : 'N/A'));
  Logger.log('Blocos de standings: ' + (data.standings ? data.standings.length : 0));

  if (!data.standings) { Logger.log('Corpo: ' + resp.getContentText().slice(0, 500)); return; }

  data.standings.forEach(function(s) {
    if (s.type !== 'TOTAL') return;
    var t = s.table || [];
    var l1 = t[0] ? t[0].team.name + ' (' + t[0].playedGames + 'j)' : '-';
    var l2 = t[1] ? t[1].team.name + ' (' + t[1].playedGames + 'j)' : '-';
    Logger.log('group="' + s.group + '" | 1º ' + l1 + ' | 2º ' + l2);
  });
}

function atualizarGabarito() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gab = ss.getSheetByName('Gabarito');
  if (!gab) { Logger.log('❌ Aba Gabarito não encontrada'); return; }
  var apiKey = PropertiesService.getScriptProperties().getProperty('FOOTBALL_DATA_KEY');
  if (!apiKey) { Logger.log('❌ API Key não configurada'); return; }

  atualizarJogosBrasil_(apiKey, gab);
  atualizarClassificacaoGrupos_(apiKey, gab);
}

function atualizarJogosBrasil_(apiKey, gab) {
  var url = API_BASE + '/competitions/WC/matches?status=FINISHED';
  var response = UrlFetchApp.fetch(url, { headers: { 'X-Auth-Token': apiKey }, muteHttpExceptions: true });
  var data = JSON.parse(response.getContentText());

  if (!data.matches || data.matches.length === 0) { Logger.log('ℹ️ Nenhum jogo finalizado'); return; }

  var brasilMatches = data.matches.filter(function(m) {
    return m.stage === 'GROUP_STAGE' && (m.homeTeam.id === 764 || m.awayTeam.id === 764);
  });
  brasilMatches.sort(function(a, b) { return new Date(a.utcDate) - new Date(b.utcDate); });

  for (var i = 0; i < brasilMatches.length && i < 3; i++) {
    var match = brasilMatches[i];
    var gabRow = 5 + i;

    if (i === 2) {
      // Jogo 3: Escócia x Brasil (B7 = gols Escócia, C7 = gols Brasil)
      if (match.homeTeam.id === 764) {
        gab.getRange(gabRow, 2).setValue(match.score.fullTime.away);
        gab.getRange(gabRow, 3).setValue(match.score.fullTime.home);
      } else {
        gab.getRange(gabRow, 2).setValue(match.score.fullTime.home);
        gab.getRange(gabRow, 3).setValue(match.score.fullTime.away);
      }
    } else {
      // Jogos 1 e 2: Brasil é Time 1 (B = gols Brasil, C = gols adversário)
      if (match.homeTeam.id === 764) {
        gab.getRange(gabRow, 2).setValue(match.score.fullTime.home);
        gab.getRange(gabRow, 3).setValue(match.score.fullTime.away);
      } else {
        gab.getRange(gabRow, 2).setValue(match.score.fullTime.away);
        gab.getRange(gabRow, 3).setValue(match.score.fullTime.home);
      }
    }
    Logger.log('✅ Jogo ' + (i+1) + ': ' + match.homeTeam.name + ' ' + match.score.fullTime.home + 'x' + match.score.fullTime.away + ' ' + match.awayTeam.name);
  }
}

// ===== ATUALIZAR CLASSIFICAÇÃO DOS GRUPOS =====
// ✅ Normaliza qualquer formato de grupo ("Group C", "GROUP_C", "Grupo C", "C").
// ✅ Atualiza assim que houver pelo menos 1 jogo disputado no grupo (incremental).
// ✅ Loga o motivo quando NÃO atualiza, e o HTTP em caso de erro.
function atualizarClassificacaoGrupos_(apiKey, gab) {
  var resp = UrlFetchApp.fetch(API_BASE + '/competitions/WC/standings',
    { headers: { 'X-Auth-Token': apiKey }, muteHttpExceptions: true });
  var code = resp.getResponseCode();
  var data = JSON.parse(resp.getContentText());

  if (code !== 200) {
    Logger.log('❌ Standings HTTP ' + code + ': ' + resp.getContentText().slice(0, 300));
    return;
  }
  if (!data.standings || data.standings.length === 0) {
    Logger.log('ℹ️ data.standings vazio. Campos recebidos: ' + Object.keys(data).join(', '));
    return;
  }

  var teamNameMap = {
    'South Africa': '🇿🇦 África do Sul', 'South Korea': '🇰🇷 Coreia do Sul',
    'Korea Republic': '🇰🇷 Coreia do Sul', 'Mexico': '🇲🇽 México',
    'Czech Republic': '🇨🇿 Tchéquia', 'Czechia': '🇨🇿 Tchéquia',
    'Bosnia-Herzegovina': '🇧🇦 Bósnia e Herz.', 'Bosnia and Herzegovina': '🇧🇦 Bósnia e Herz.',
    'Canada': '🇨🇦 Canadá', 'Qatar': '🇶🇦 Catar', 'Switzerland': '🇨🇭 Suíça',
    'Brazil': '🇧🇷 Brasil', 'Scotland': '🏴󠁧󠁢󠁳󠁣󠁴󠁿 Escócia', 'Haiti': '🇭🇹 Haiti',
    'Morocco': '🇲🇦 Marrocos', 'Australia': '🇦🇺 Austrália',
    'United States': '🇺🇸 Estados Unidos', 'USA': '🇺🇸 Estados Unidos',
    'Türkiye': '🇹🇷 Turquía', 'Turkey': '🇹🇷 Turquía', 'Paraguay': '🇵🇾 Paraguai',
    'Germany': '🇩🇪 Alemanha', "Côte d'Ivoire": '🇨🇮 Costa do Marfim',
    'Ivory Coast': '🇨🇮 Costa do Marfim', 'Curaçao': '🇨🇼 Curaçao',
    'Ecuador': '🇪🇨 Equador',
    'Netherlands': '🇳🇱 Holanda', 'Japan': '🇯🇵 Japão', 'Sweden': '🇸🇪 Suécia',
    'Tunisia': '🇹🇳 Tunísia', 'Belgium': '🇧🇪 Bélgica', 'Egypt': '🇪🇬 Egito',
    'Iran': '🇮🇷 Irã', 'New Zealand': '🇳🇿 Nova Zelândia',
    'Saudi Arabia': '🇸🇦 Arábia Saudita',
    'Cape Verde Islands': '🇨🇻 Cabo Verde', 'Cape Verde': '🇨🇻 Cabo Verde',
    'Spain': '🇪🇸 Espanha', 'Uruguay': '🇺🇾 Uruguai', 'France': '🇫🇷 França',
    'Iraq': '🇮🇶 Iraque', 'Norway': '🇳🇴 Noruega', 'Senegal': '🇸🇳 Senegal',
    'Algeria': '🇩🇿 Argélia', 'Argentina': '🇦🇷 Argentina', 'Austria': '🇦🇹 Áustria',
    'Jordan': '🇯🇴 Jordânia', 'Colombia': '🇨🇴 Colômbia', 'Portugal': '🇵🇹 Portugal',
    'DR Congo': '🇨🇩 Rep. Dem. Congo', 'Congo DR': '🇨🇩 Rep. Dem. Congo',
    'Uzbekistan': '🇺🇿 Uzbequistão', 'Croatia': '🇭🇷 Croácia',
    'Ghana': '🇬🇭 Gana', 'England': '🏴󠁧󠁢󠁥󠁮󠁧󠁿 Inglaterra', 'Panama': '🇵🇦 Panamá'
  };

  // Aceita "GROUP_C", "Group C", "Grupo C" ou "C" → "C"
  function letraDoGrupo(g) {
    if (!g) return null;
    var m = String(g).toUpperCase().match(/([A-L])\s*$/);
    return m ? m[1] : null;
  }
  var letraToRow = { A:10,B:11,C:12,D:13,E:14,F:15,G:16,H:17,I:18,J:19,K:20,L:21 };

  var atualizados = 0;
  for (var i = 0; i < data.standings.length; i++) {
    var standing = data.standings[i];
    if (standing.type !== 'TOTAL') continue;

    var letra = letraDoGrupo(standing.group);
    if (!letra || !letraToRow[letra]) {
      Logger.log('⚠️ Grupo não reconhecido: "' + standing.group + '" (ignorado)');
      continue;
    }
    var gabRow = letraToRow[letra];
    var table = standing.table;
    if (!table || table.length < 2) continue;

    // ✅ Incremental: basta 1 jogo disputado no grupo
    var jogosNoGrupo = 0;
    for (var t = 0; t < table.length; t++) jogosNoGrupo += (table[t].playedGames || 0);
    if (jogosNoGrupo === 0) {
      Logger.log('⏳ Grupo ' + letra + ': nenhum jogo ainda.');
      continue;
    }

    var prim = teamNameMap[table[0].team.name] || table[0].team.name;
    var seg  = teamNameMap[table[1].team.name] || table[1].team.name;

    gab.getRange(gabRow, 2).setValue(prim);
    gab.getRange(gabRow, 3).setValue(seg);
    atualizados++;
    Logger.log('✅ Grupo ' + letra + ' (' + (jogosNoGrupo / 2) + ' jogo(s)): 1º ' + prim + ' | 2º ' + seg);
  }
  Logger.log('📊 Grupos atualizados: ' + atualizados + '/12');
}

function criarGatilhoGabarito() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'atualizarGabarito') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('atualizarGabarito').timeBased().everyHours(2).create();
  Logger.log('✅ Gatilho criado: verificar resultados a cada 2 horas');
}

function forcarAtualizacaoGrupos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gab = ss.getSheetByName('Gabarito');
  var apiKey = PropertiesService.getScriptProperties().getProperty('FOOTBALL_DATA_KEY');
  if (!apiKey || !gab) { Logger.log('❌ Key ou Gabarito não encontrado'); return; }
  Logger.log('🔄 Forçando atualização...');
  atualizarClassificacaoGrupos_(apiKey, gab);
  Logger.log('✅ Concluído!');
}
