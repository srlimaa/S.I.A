// PERÍODOS LETIVOS: aba "Periodos" da planilha mestre (A período | B ID da planilha | C ocorrências UNILEÃO | D ocorrências UNIVS).
// Cada período tem a sua própria planilha de registros (Presencial, Ramal, Email, Virtual). Também cria um período novo.
// ══════════════════════════════════════════════════════════════════════
// PERÍODOS LETIVOS — aba "Periodos" da planilha mestre e criação de novos períodos.
// ══════════════════════════════════════════════════════════════════════

// Converte o valor da aba "Periodos" para "AAAA.S".
// O Sheets pode transformar "2026.2" em data (fev/2026) ou número; aqui o MÊS da data é o próprio semestre.
function formatPeriodo_(val) {
  if (val instanceof Date) return val.getFullYear() + '.' + (val.getMonth() + 1);
  return formatSemestre(val);
}

// Grava uma linha na aba Periodos sempre como texto
function gravarPeriodo_(sh, periodo, id) {
  var r = sh.getLastRow() + 1;
  var rng = sh.getRange(r, 1, 1, 2);
  rng.setNumberFormat('@');
  rng.setValues([[periodo, id]]);
}

// ── Períodos letivos ─────────────────────────────────────────────────────────
// Aba "Periodos" (na planilha mestre): A-Período | B-ID da planilha
// Abre a aba Periodos (criando-a, com o período inicial, se não existir). Colunas em texto para "2026.2" não virar data.
function getPeriodosSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('Periodos');
  if (!sh) {
    sh = ss.insertSheet('Periodos');
    sh.getRange('A:B').setNumberFormat('@');        // texto puro: evita "2026.2" virar número/data
    sh.appendRow(['Período', 'ID da Planilha', 'Ocorrências UNILEÃO', 'Ocorrências UNIVS']);
    gravarPeriodo_(sh, PERIODO_INICIAL, SPREADSHEET_ID);
  }
  return sh;
}

// Lista os períodos cadastrados (mais recente primeiro) para os seletores da tela.
function getPeriodos() {
  try {
    var data = getPeriodosSheet_().getDataRange().getValues();
    var lista = [];
    for (var i = 1; i < data.length; i++) {
      var p = formatPeriodo_(data[i][0]);
      if (p && String(data[i][1]).trim()) lista.push(p);
    }
    lista.sort().reverse();                          // mais recente primeiro
    return lista;
  } catch(e) { return []; }
}

// Devolve a planilha do período (o ID nunca vai para o navegador)
function abrirPlanilhaPeriodo_(periodo) {
  if (!periodo) throw new Error('Nenhum período letivo selecionado.');
  var data = getPeriodosSheet_().getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (formatPeriodo_(data[i][0]) === String(periodo).trim())
      return SpreadsheetApp.openById(String(data[i][1]).trim());
  }
  throw new Error('Período letivo "' + periodo + '" não encontrado.');
}

// Cria as abas de registro (com cabeçalho) que ainda não existirem na planilha do período
function prepararPlanilhaPeriodo_(ss) {
  var abas = [
    { nomes: ['Presencial'], cab: ['ID','Data','CPF','Nome Completo','Curso','Assunto','Semestre','IES','Colaborador','Data do Registro'] },
    { nomes: ['Ramal'],      cab: ['ID','Data','CPF','Nome Completo','Telefone','Assunto','Observação','Colaborador','Data do Registro'] },
    { nomes: ['Email','E-mail'], cab: ['ID','Data de Recebimento','CPF','Nome Completo','E-mail do Aluno','Assunto','Semestre','IES','Data de Retorno','Observação','Colaborador','Data do Registro'] },
    { nomes: ['Virtual'],    cab: ['ID','Data','Colaborador','Conversas Atribuídas','Conversas Fechadas','Registrado Por','Data do Registro'] }
  ];
  abas.forEach(function(a) {
    var existe = a.nomes.some(function(n) { return ss.getSheetByName(n); });
    if (!existe) {
      var sh = ss.insertSheet(a.nomes[0]);
      sh.appendRow(a.cab);
      sh.setFrozenRows(1);
    }
  });
}

// Admin: cadastra um período (AAAA.S) com o ID ou link da planilha, cria nela as abas que faltam e grava em Periodos.
function criarPeriodo(periodo, planilhaId, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode cadastrar períodos.' };
    periodo = String(periodo || '').trim();
    planilhaId = String(planilhaId || '').trim();
    // aceita também o link completo da planilha
    var m = planilhaId.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (m) planilhaId = m[1];
    if (!/^\d{4}\.[12]$/.test(periodo))
      return { success: false, message: 'Formato do período inválido. Use AAAA.S (ex.: 2027.1).' };
    if (!planilhaId) return { success: false, message: 'Informe o ID da planilha.' };
    if (getPeriodos().indexOf(periodo) !== -1)
      return { success: false, message: 'Este período já está cadastrado.' };

    var ss;
    try { ss = SpreadsheetApp.openById(planilhaId); }
    catch(e) { return { success: false, message: 'Não foi possível abrir a planilha. Confira o ID e o acesso da conta que executa o script.' }; }

    prepararPlanilhaPeriodo_(ss);
    gravarPeriodo_(getPeriodosSheet_(), periodo, planilhaId);
    return { success: true, periodo: periodo };
  } catch(e) { return { success: false, message: e.message }; }
}
