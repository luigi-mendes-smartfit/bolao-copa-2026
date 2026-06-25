// 🏆 BOLÃO COPA DO MUNDO 2026 - SMART FIT
// ============================================
// COMECE POR AQUI
// SERÁ GERADO O FORMS, CONFIGURE ELE, na planilha que egerar as respostas do forms, rode master.gs, crie uma outra planilha e rode publica.gs
// SCRIPT PARA CONFIGURAR A PLANILHA DE RESPOSTAS DO GOOGLE FORMS
// 
// COMO USAR:
// 1) Abra a planilha de respostas vinculada ao seu Google Forms
// 2) Vá em: Extensões > Apps Script
// 3) Cole ESTE script inteiro
// 4) Execute a função: configurarPlanilha()
// 5) Pronto! As abas Gabarito, Pontuação e Ranking serão criadas
//
// ESTRUTURA ESPERADA DA ABA DE RESPOSTAS (verificada):
// Col A = Timestamp
// Col B = Email Address
// Col C = Nome completo
// Col D = Área/unidade
// Col E = Gols BRA (J1: Brasil x Marrocos)
// Col F = Gols MAR (J1)
// Col G = Gols BRA (J2: Brasil x Haiti)
// Col H = Gols HAI (J2)
// Col I = Gols SCO (J3: Escócia x Brasil)
// Col J = Gols BRA (J3)
// Col K-L = Grupo A [Primeiro/Segundo]
// Col M-N = Grupo B ... até Col AG-AH = Grupo L
//
// GATILHO: após configurar, execute criarGatilho() para auto-atualizar
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
  
  Logger.log('');
  Logger.log('✅ Planilha configurada com sucesso!');
  Logger.log('📝 Abas criadas: Gabarito, Pontuação, Ranking, Regras');
  Logger.log('');
  Logger.log('👉 Agora execute criarGatilho() para atualizar automaticamente a cada resposta!');
}

// ===== GATILHO AUTOMÁTICO =====
function criarGatilho() {
  // Remove triggers antigos
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Cria novo trigger
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  
  Logger.log('✅ Gatilho criado! A cada nova resposta, a planilha atualizará automaticamente.');
}

function onFormSubmit(e) {
  // Trigger chamado a cada resposta nova
  // As fórmulas já são dinâmicas, mas podemos forçar refresh e notificar
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ptsSheet = ss.getSheetByName('Pontuação');
  if (ptsSheet) {
    // Força recalcular
    SpreadsheetApp.flush();
  }
  Logger.log('📬 Nova resposta recebida! Pontuação atualizada.');
}

// ===== ABA GABARITO =====
function criarAbaGabarito_(ss) {
  var existing = ss.getSheetByName('Gabarito');
  if (existing) ss.deleteSheet(existing);
  
  var gab = ss.insertSheet('Gabarito');
  
  // Header
  gab.getRange('A1:C1').merge().setValue('📝 GABARITO — Preencha com os resultados reais')
    .setBackground('#000000').setFontColor('#FFD100').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  
  gab.getRange('A2').setValue('Preencha após cada rodada. A aba Pontuação recalcula automaticamente.')
    .setFontStyle('italic').setFontColor('#666666');
  
  // Jogos
  gab.getRange('A4:C4').setValues([['Jogo', 'Gols Time 1', 'Gols Time 2']])
    .setBackground('#006400').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  
  gab.getRange('A5').setValue('Brasil x Marrocos');
  gab.getRange('A6').setValue('Brasil x Haiti');
  gab.getRange('A7').setValue('Escócia x Brasil');
  gab.getRange('B5:C7').setBackground('#E8FFE8').setHorizontalAlignment('center');
  
  var rule0a10 = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(0, 10).setAllowInvalid(false).build();
  gab.getRange('B5:C7').setDataValidation(rule0a10);
  
  // Grupos — COM BANDEIRAS (mesmo formato das respostas!)
  gab.getRange('A9:C9').setValues([['Grupo', '1º Colocado Real', '2º Colocado Real']])
    .setBackground('#006400').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  
  // IMPORTANTE: As opções devem ter BANDEIRAS igual às respostas do Forms!
  var grupos = {
    'A': ['🇿🇦 África do Sul', '🇰🇷 Coreia do Sul', '🇲🇽 México', '🇨🇿 Tchéquia'],
    'B': ['🇧🇦 Bósnia e Herz.', '🇨🇦 Canadá', '🇶🇦 Catar', '🇨🇭 Suíça'],
    'C': ['🇧🇷 Brasil', '🏴󠁧󠁢󠁳󠁣󠁴󠁿 Escócia', '🇭🇹 Haiti', '🇲🇦 Marrocos'],
    'D': ['🇦🇺 Austrália', '🇺🇸 Estados Unidos', '🇵🇾 Paraguai', '🇹🇷 Turquia'],
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

// ===== ABA PONTUAÇÃO (fórmulas dinâmicas) =====
function criarAbaPontuacao_(ss, respostasName) {
  var existing = ss.getSheetByName('Pontuação');
  if (existing) ss.deleteSheet(existing);
  
  var pts = ss.insertSheet('Pontuação');
  var R = "'" + respostasName + "'";
  
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
  
  // MAPEAMENTO REAL DAS COLUNAS (confirmado via export):
  // Col C = Nome
  // Col E,F = Jogo 1 (BRA, MAR)
  // Col G,H = Jogo 2 (BRA, HAI)
  // Col I,J = Jogo 3 (SCO, BRA)
  // Col K,L = Grupo A [Primeiro, Segundo]
  // Col M,N = Grupo B ... até Col AG,AH = Grupo L
  
  // Gabarito references:
  // B5,C5 = Jogo 1 (BRA gols, MAR gols)
  // B6,C6 = Jogo 2 (BRA gols, HAI gols)
  // B7,C7 = Jogo 3 (SCO gols, BRA gols)
  // B10,C10 = Grupo A (1º, 2º) ... B21,C21 = Grupo L
  
  // Group column pairs in the Respostas sheet:
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
  
  var NUM = 50; // max participantes
  
  for (var p = 0; p < NUM; p++) {
    var row = p + 4;   // row in Pontuação
    var rr = p + 2;    // row in Respostas (row 1 = header, row 2 = first response)
    
    // Nome (Col C na aba de respostas)
    pts.getRange(row, 1).setFormula("=IFERROR(" + R + "!C" + rr + ',"")');
    
    // Jogo 1: E,F vs Gabarito B5,C5
    pts.getRange(row, 2).setFormula(
      '=IF(OR(Gabarito!B5="",' + R + '!E' + rr + '=""),"",IF(AND(' + R + '!E' + rr + '=Gabarito!B5,' + R + '!F' + rr + '=Gabarito!C5),5,IF(SIGN(' + R + '!E' + rr + '-' + R + '!F' + rr + ')=SIGN(Gabarito!B5-Gabarito!C5),3,0)))'
    );
    
    // Jogo 2: G,H vs Gabarito B6,C6
    pts.getRange(row, 3).setFormula(
      '=IF(OR(Gabarito!B6="",' + R + '!G' + rr + '=""),"",IF(AND(' + R + '!G' + rr + '=Gabarito!B6,' + R + '!H' + rr + '=Gabarito!C6),5,IF(SIGN(' + R + '!G' + rr + '-' + R + '!H' + rr + ')=SIGN(Gabarito!B6-Gabarito!C6),3,0)))'
    );
    
    // Jogo 3: I,J vs Gabarito B7,C7
    pts.getRange(row, 4).setFormula(
      '=IF(OR(Gabarito!B7="",' + R + '!I' + rr + '=""),"",IF(AND(' + R + '!I' + rr + '=Gabarito!B7,' + R + '!J' + rr + '=Gabarito!C7),5,IF(SIGN(' + R + '!I' + rr + '-' + R + '!J' + rr + ')=SIGN(Gabarito!B7-Gabarito!C7),3,0)))'
    );
    
    // Sub Jogos
    pts.getRange(row, 5).setFormula('=IF(A' + row + '="","",IFERROR(B' + row + '+C' + row + '+D' + row + ',0))');
    
    // Grupos A-L
    for (var g = 0; g < 12; g++) {
      var gabRow = 10 + g;
      var col1 = grpCols[g][0]; // Primeiro
      var col2 = grpCols[g][1]; // Segundo
      
      // Pontuação grupos:
      // 5 = ambos corretos na ordem
      // 3 = ambos corretos, ordem trocada
      // 2 = 1 certo na posição correta
      // 1 = 1 certo na posição errada
      // 0 = errou ambos ou duplicata
      pts.getRange(row, 6 + g).setFormula(
        '=IF(OR(Gabarito!B' + gabRow + '="",' + R + '!' + col1 + rr + '=""),"",'+
        'IF(' + R + '!' + col1 + rr + '=' + R + '!' + col2 + rr + ',0,' +
        'IF(AND(' + R + '!' + col1 + rr + '=Gabarito!B' + gabRow + ',' + R + '!' + col2 + rr + '=Gabarito!C' + gabRow + '),5,' +
        'IF(AND(' + R + '!' + col1 + rr + '=Gabarito!C' + gabRow + ',' + R + '!' + col2 + rr + '=Gabarito!B' + gabRow + '),3,' +
        'IF(OR(' + R + '!' + col1 + rr + '=Gabarito!B' + gabRow + ',' + R + '!' + col2 + rr + '=Gabarito!C' + gabRow + '),2,' +
        'IF(OR(' + R + '!' + col1 + rr + '=Gabarito!C' + gabRow + ',' + R + '!' + col2 + rr + '=Gabarito!B' + gabRow + '),1,' +
        '0))))))'
      );
    }
    
    // Sub Grupos (sum F:Q)
    pts.getRange(row, 18).setFormula('=IF(A' + row + '="","",IFERROR(SUM(F' + row + ':Q' + row + '),0))');
    
    // TOTAL
    pts.getRange(row, 19).setFormula('=IF(A' + row + '="","",IFERROR(E' + row + '+R' + row + ',0))');
    
    // Desempate 1: Placares exatos (quantos 5 em B:D)
    pts.getRange(row, 20).setFormula('=IF(A' + row + '="","",COUNTIF(B' + row + ':D' + row + ',5))');
    
    // Desempate 2: Posições cravadas (quantos 5 em F:Q)
    pts.getRange(row, 21).setFormula('=IF(A' + row + '="","",COUNTIF(F' + row + ':Q' + row + ',5))');
    
    // Desempate 3: Timestamp (Col A na aba de respostas)
    pts.getRange(row, 22).setFormula("=IFERROR(" + R + "!A" + rr + ',"")');
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
  
  // Freeze header
  pts.setFrozenRows(3);
  
  Logger.log('  ✅ Aba Pontuação criada (com fórmulas dinâmicas)');
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
  
  for (var p = 0; p < 50; p++) {
    var row = p + 4;
    var ptsRow = p + 4;
    rank.getRange(row, 1).setFormula('=IF(B' + row + '="","",ROW()-3)');
    rank.getRange(row, 2).setFormula("=IFERROR(Pontuação!A" + ptsRow + ',"")');
    rank.getRange(row, 3).setFormula("=IFERROR(Pontuação!S" + ptsRow + ',"")');
    rank.getRange(row, 4).setFormula("=IFERROR(Pontuação!T" + ptsRow + ',"")');
    rank.getRange(row, 5).setFormula("=IFERROR(Pontuação!U" + ptsRow + ',"")');
    rank.getRange(row, 6).setFormula("=IFERROR(Pontuação!V" + ptsRow + ',"")');
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
  rank.getRange('F4:F53').setNumberFormat('dd/MM/yyyy HH:mm:ss');
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


// ===== CORRIGIR FORMATO DO TIMESTAMP NO RANKING =====
// Execute após configurarPlanilha() se o timestamp aparecer como número
function corrigirTimestamp() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Formatar timestamp na aba Pontuação
  var pts = ss.getSheetByName('Pontuação');
  if (pts) {
    pts.getRange('V4:V53').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }
  
  // Formatar timestamp na aba Ranking
  var rank = ss.getSheetByName('Ranking');
  if (rank) {
    rank.getRange('F4:F53').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }
  
  Logger.log('✅ Formato de timestamp corrigido (dd/MM/yyyy HH:mm:ss)');
}

// ===== PROTEGER ABAS (AUDITORIA) =====
// Protege as abas para que ninguém edite manualmente.
// Fórmulas e scripts continuam funcionando normalmente.
// Apenas o owner da planilha pode desproteger.

function protegerAbas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var owner = Session.getEffectiveUser();
  
  var abasParaProteger = ['Pontuação', 'Ranking'];
  
  for (var i = 0; i < abasParaProteger.length; i++) {
    var sheet = ss.getSheetByName(abasParaProteger[i]);
    if (sheet) {
      // Remove proteções existentes
      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (var j = 0; j < protections.length; j++) {
        protections[j].remove();
      }
      
      // Criar nova proteção
      var protection = sheet.protect().setDescription('🔒 Protegido - Bolão Copa 2026 (auditoria)');
      
      // Apenas o owner pode editar
      protection.addEditor(owner);
      protection.removeEditors(protection.getEditors());
      
      // Se quiser que NINGUÉM edite (nem o owner via interface):
      protection.setWarningOnly(false);
      
      // Manter apenas o owner
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
      
      Logger.log('🔒 Aba "' + abasParaProteger[i] + '" protegida');
    }
  }
  
  // Proteger aba de Respostas (a do Forms)
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf('Respostas') > -1 || name.indexOf('Form') > -1 || name.indexOf('Response') > -1) {
      var protections = sheets[i].getProtections(SpreadsheetApp.ProtectionType.SHEET);
      for (var j = 0; j < protections.length; j++) {
        protections[j].remove();
      }
      
      var protection = sheets[i].protect().setDescription('🔒 Respostas do Forms - Não editar!');
      protection.addEditor(owner);
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
      protection.setWarningOnly(false);
      Logger.log('🔒 Aba "' + name + '" protegida');
    }
  }
  
  Logger.log('');
  Logger.log('✅ Todas as abas protegidas!');
  Logger.log('   Apenas o owner pode desproteger via: Dados > Proteger planilhas e intervalos');
  Logger.log('   Fórmulas e o trigger onFormSubmit continuam funcionando normalmente.');
}

// ===== DESPROTEGER (caso precise editar algo) =====
function desprotegerAbas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  for (var i = 0; i < sheets.length; i++) {
    var protections = sheets[i].getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var j = 0; j < protections.length; j++) {
      protections[j].remove();
    }
  }
  
  Logger.log('🔓 Todas as proteções removidas.');
}
