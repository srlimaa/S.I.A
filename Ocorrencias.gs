// ══════════════════════════════════════════════════════════════════════
// OCORRÊNCIAS / ANEXOS (UNILEÃO e UNIVS) — links das planilhas, pendentes, marcação de responsável e confirmação.
// ══════════════════════════════════════════════════════════════════════

// ── Planilhas de ocorrências/anexos (registro manual; apenas leitura) ────────
// Aba "Periodos": C-ID da planilha UNILEÃO | D-ID da planilha UNIVS
// Em cada planilha (1ª aba): Ocorrências → nome na col C, data na col D | Anexos → nome na col E, data na col F
// Linha inicial: UNILEÃO = 6 | UNIVS = 8
function extrairIdPlanilha_(txt) {
  txt = String(txt || '').trim();
  var m = txt.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : txt;
}

function urlPlanilha_(id) { return id ? 'https://docs.google.com/spreadsheets/d/' + id : ''; }

function linhaPeriodo_(periodo) {
  var sh = getPeriodosSheet_();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (formatPeriodo_(data[i][0]) === String(periodo).trim()) return { sh: sh, linha: i + 1, data: data[i] };
  }
  return null;
}

function getLinksPeriodo_(periodo) {
  var r = linhaPeriodo_(periodo);
  if (!r) return { unileao: '', univs: '' };
  return { unileao: String(r.data[2] || '').trim(), univs: String(r.data[3] || '').trim() };
}

function getLinksOcorrencias(periodo, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin.' };
    var l = getLinksPeriodo_(periodo);
    return { success: true, unileao: urlPlanilha_(l.unileao), univs: urlPlanilha_(l.univs) };
  } catch(e) { return { success: false, message: e.message }; }
}

function salvarLinksOcorrencias(periodo, linkUnileao, linkUnivs, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode alterar os links.' };
    var r = linhaPeriodo_(periodo);
    if (!r) return { success: false, message: 'Período não encontrado.' };
    var ids = [extrairIdPlanilha_(linkUnileao), extrairIdPlanilha_(linkUnivs)];
    var nomes = ['UNILEÃO', 'UNIVS'];
    for (var k = 0; k < 2; k++) {
      if (!ids[k]) continue;
      try { SpreadsheetApp.openById(ids[k]); }
      catch(e) { return { success: false, message: 'Não foi possível abrir a planilha ' + nomes[k] + '. Confira o link e se a conta que executa o script tem acesso.' }; }
    }
    var cab = r.sh.getRange(1, 1, 1, 4).getValues()[0];
    if (!String(cab[2] || '').trim()) r.sh.getRange(1, 3).setValue('Ocorrências UNILEÃO');
    if (!String(cab[3] || '').trim()) r.sh.getRange(1, 4).setValue('Ocorrências UNIVS');
    var rng = r.sh.getRange(r.linha, 3, 1, 2);
    rng.setNumberFormat('@');
    rng.setValues([[ids[0], ids[1]]]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// Lê a 1ª aba da planilha e soma 1 por linha preenchida: Ocorrência (C/D) e Anexo (E/F)
function lerOcorrenciasPlanilha_(id, linhaInicial, rotulo, ctx) {
  var livro = SpreadsheetApp.openById(id), tzLivro = livro.getSpreadsheetTimeZone();
  var sh = livro.getSheets()[0];
  var ultima = sh.getLastRow();
  if (ultima < linhaInicial) return;
  var dados = sh.getRange(linhaInicial, 3, ultima - linhaInicial + 1, 4).getValues();   // colunas C..F
  var conf = [
    { tipo: 'Ocorrência ' + rotulo, nome: 0, data: 1, letraData: 'D' },
    { tipo: 'Anexo ' + rotulo,      nome: 2, data: 3, letraData: 'F' }
  ];
  conf.forEach(function(cf) {
    var semData = 0;
    for (var r = 0; r < dados.length; r++) {
      var nome = String(dados[r][cf.nome] || '').trim();
      if (!nome) continue;
      var d = dataBR_(dados[r][cf.data], tzLivro);
      if (!d) { semData++; continue; }
      var c = nomeCanonico_(nome, ctx.canon);
      ctx.addColab(c);
      ctx.ensure(d, cf.tipo);
      ctx.ind[d][cf.tipo][c] = (ctx.ind[d][cf.tipo][c] || 0) + 1;
    }
    if (semData > 0)
      ctx.avisos.push(semData + ' linha(s) de "' + cf.tipo + '" ignorada(s) por não ter data válida na coluna ' + cf.letraData + '.');
  });
}

// ── Ocorrências pendentes (planilhas UNILEÃO / UNIVS) ────────────────────────
// Colunas: A CPF | B Aluno | C/D colaborador+data da OCORRÊNCIA | E/F colaborador+data do ANEXO |
//          H data a informar | I semestre | J tipo | K obs. mala | L nome do arquivo (link) | N arquivo de anexo (link; "--" = sem anexo)
// Linha inicial: UNILEÃO = 6 | UNIVS = 8. Cada confirmação é gravada na planilha de ocorrências e na aba "Ocorrencias" do período.
var OCORR_IES = { unileao: { linha: 6, rotulo: 'UNILEÃO' }, univs: { linha: 8, rotulo: 'UNIVS' } };

// Texto + endereço do link de uma célula (hiperlink da célula ou URL escrita no texto)
function linkInfo_(rv, plain) {
  var texto = String(plain == null ? '' : plain).trim(), url = '';
  try {
    if (rv) {
      url = rv.getLinkUrl() || '';
      if (!url) {
        var runs = rv.getRuns();
        for (var k = 0; k < runs.length && !url; k++) url = runs[k].getLinkUrl() || '';
      }
    }
  } catch(e) {}
  if (!url && /^https?:\/\//i.test(texto)) url = texto;
  if (url && !/^https?:\/\//i.test(url)) url = '';
  return { texto: texto, url: url };
}

function chaveLogOcorr_(ies, linha, cpf, etapa) { return ies + '|' + linha + '|' + cpfNum_(cpf) + '|' + etapa; }

function lerLogOcorrencias_(periodo) {
  var set = {};
  try {
    var sh = abrirPlanilhaPeriodo_(periodo).getSheetByName('Ocorrencias');
    if (!sh || sh.getLastRow() < 2) return set;
    var d = sh.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var iesKey = String(d[i][2]).toUpperCase() === 'UNIVS' ? 'univs' : 'unileao';
      set[chaveLogOcorr_(iesKey, d[i][3], d[i][4], String(d[i][6]))] = true;
    }
  } catch(e) {}
  return set;
}

// Marcação de "responsável": enquanto alguém está com a janela de um aluno aberta, ele fica marcado (some sozinho
// se a janela for fechada ou o navegador parar de avisar). Evita que duas pessoas gerem a mesma ocorrência.
var OC_CLAIM_TTL = 150;   // segundos sem "sinal de vida" até a marcação expirar

function ocChaveClaim_(periodo, ies, linha, cpf) { return 'ocr|' + periodo + '|' + ies + '|' + linha + '|' + cpfNum_(cpf); }

function ocClaimLer_(cache, chave) {
  var v = cache.get(chave);
  if (!v) return null;
  try { return JSON.parse(v); } catch(e) { return null; }
}

// acao: 'abrir' | 'bater' (renova) | 'forcar' (assume mesmo ocupado) | 'liberar'
function ocClaim(periodo, ies, linha, cpf, acao, userId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    var usr = usuarioPorId_(userId);
    if (!usr) return { success: false, message: 'Usuário não identificado.' };
    var cache = CacheService.getScriptCache(), k = ocChaveClaim_(periodo, ies, linha, cpf);
    var atual = ocClaimLer_(cache, k), meu = !!(atual && String(atual.id) === String(userId));
    if (acao === 'liberar') { if (meu) cache.remove(k); return { success: true }; }
    if (atual && !meu && acao !== 'forcar')
      return { success: false, ocupado: true, resp: { nome: atual.nome, desde: atual.desde } };
    var novo = meu ? atual : { id: String(userId), nome: usr.nome, desde: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm') };
    cache.put(k, JSON.stringify(novo), OC_CLAIM_TTL);
    return { success: true, resp: { nome: novo.nome, desde: novo.desde } };
  } catch(e) { return { success: false, message: e.message }; }
  finally { try { lock.releaseLock(); } catch(x) {} }
}

// Quem está com cada aluno da lista aberto agora. itens: [{linha, cpf}]  →  { linha: {nome, desde} }
function getOcorrenciasResponsaveis(periodo, ies, itens) {
  try {
    var cache = CacheService.getScriptCache(), chaves = [], porChave = {};
    (itens || []).forEach(function(it) { var k = ocChaveClaim_(periodo, ies, it.linha, it.cpf); chaves.push(k); porChave[k] = it.linha; });
    var out = {};
    if (!chaves.length) return { success: true, resp: out };
    var achados = cache.getAll(chaves);
    Object.keys(achados || {}).forEach(function(k) {
      try { var c = JSON.parse(achados[k]); out[porChave[k]] = { nome: c.nome, desde: c.desde }; } catch(e) {}
    });
    return { success: true, resp: out };
  } catch(e) { return { success: false, message: e.message }; }
}

// Lista só os pendentes. Ocorrência pendente: a DATA (coluna D) está vazia — mesmo que haja nome na coluna C —, com algo na coluna L.
// Anexo pendente: coluna N com link (diferente de "--") e a DATA do anexo (coluna F) vazia. Uma linha com ocorrência já registrada
// só continua na lista se essa ocorrência foi confirmada por este sistema e o anexo ainda está pendente.
function getOcorrenciasPendentes(periodo, ies) {
  try {
    var conf = OCORR_IES[ies];
    if (!conf) return { success: false, message: 'Planilha inválida.' };
    var id = getLinksPeriodo_(periodo)[ies];
    if (!id) return { success: false, message: 'A planilha de ocorrências da ' + conf.rotulo + ' ainda não foi cadastrada neste período. Peça a um Admin para cadastrá-la em Configurações › Planilhas.' };
    var book = SpreadsheetApp.openById(id);
    var tz = book.getSpreadsheetTimeZone();           // datas lidas no fuso da planilha
    var sh = book.getSheets()[0];
    var ultima = sh.getLastRow();
    if (ultima < conf.linha) return { success: true, itens: [], rotulo: conf.rotulo };
    var n = ultima - conf.linha + 1;
    var vals = sh.getRange(conf.linha, 1, n, 14).getValues();           // A..N
    var rich = sh.getRange(conf.linha, 12, n, 3).getRichTextValues();   // L..N
    var log = lerLogOcorrencias_(periodo);
    var itens = [];
    for (var i = 0; i < n; i++) {
      var r = vals[i], linha = conf.linha + i;
      if (vazio_(r[11])) continue;                                       // sem informação na coluna L
      var ocPend = vazio_(r[3]);                                         // sem data na D = ocorrência não gerada
      var temAnexo = !semAnexo_(r[13]);
      var anPend = temAnexo && vazio_(r[5]);                             // sem data na F = anexo não gerado
      var ocFeitaPeloSistema = !ocPend && log[chaveLogOcorr_(ies, linha, r[0], 'Ocorrência')];
      if (!(ocPend || (anPend && ocFeitaPeloSistema))) continue;
      itens.push({
        linha: linha, cpf: String(r[0] == null ? '' : r[0]).trim(), nome: String(r[1] == null ? '' : r[1]).trim(),
        dataInfo: dataTxt_(r[7], tz), semestre: dataTxt_(r[8], tz), tipo: String(r[9] == null ? '' : r[9]).trim(),
        obs: String(r[10] == null ? '' : r[10]).trim(),
        arquivo: linkInfo_(rich[i][0], r[11]),
        anexo: temAnexo ? linkInfo_(rich[i][2], r[13]) : null,
        temAnexo: temAnexo, ocPend: ocPend, anPend: anPend, resp: null
      });
    }
    var rs = getOcorrenciasResponsaveis(periodo, ies, itens.map(function(it) { return { linha: it.linha, cpf: it.cpf }; }));
    if (rs.success) itens.forEach(function(it) { it.resp = rs.resp[it.linha] || null; });
    return { success: true, itens: itens, rotulo: conf.rotulo };
  } catch(e) { return { success: false, message: 'Não foi possível ler a planilha: ' + e.message }; }
}

// etapa: 'Ocorrência' (grava C/D) ou 'Anexo' (grava E/F). Só a data, sem hora.
// dataCliente (AAAA-MM-DD): data do computador de quem confirmou; só é aceita se estiver a no máximo 1 dia da data do servidor.
function confirmarOcorrencia(periodo, ies, linha, cpf, etapa, obs, userId, dataCliente) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var conf = OCORR_IES[ies];
    if (!conf) return { success: false, message: 'Planilha inválida.' };
    if (etapa !== 'Ocorrência' && etapa !== 'Anexo') return { success: false, message: 'Etapa inválida.' };
    var usr = usuarioPorId_(userId);
    if (!usr) return { success: false, message: 'Usuário não identificado.' };
    linha = parseInt(linha, 10);
    if (!linha || linha < conf.linha) return { success: false, message: 'Linha inválida.' };
    var id = getLinksPeriodo_(periodo)[ies];
    if (!id) return { success: false, message: 'Planilha de ocorrências não cadastrada.' };
    var book = SpreadsheetApp.openById(id), tz = book.getSpreadsheetTimeZone(), sh = book.getSheets()[0];
    var row = sh.getRange(linha, 1, 1, 14).getValues()[0];
    if (cpfNum_(row[0]) !== cpfNum_(cpf))
      return { success: false, mudou: true, message: 'A planilha foi alterada e esta linha mudou de lugar. Clique em Atualizar e tente novamente.' };

    // outra pessoa está com este aluno aberto?
    var cl = ocClaimLer_(CacheService.getScriptCache(), ocChaveClaim_(periodo, ies, linha, cpf));
    if (cl && String(cl.id) !== String(userId))
      return { success: false, ocupado: true, resp: { nome: cl.nome, desde: cl.desde }, message: cl.nome + ' está com este aluno aberto. Aguarde ou assuma o aluno antes de confirmar.' };

    var colNome = etapa === 'Ocorrência' ? 3 : 5;       // C ou E (a data fica na coluna seguinte: D ou F)
    if (etapa === 'Anexo' && semAnexo_(row[13])) return { success: false, jaFeito: true, message: 'Este aluno não possui anexo.' };
    if (!vazio_(row[colNome])) {                         // só a DATA indica que foi gerado; nome sem data = não gerado
      var quem = String(row[colNome - 1] || '').trim();
      return { success: false, jaFeito: true, message: (etapa === 'Ocorrência' ? 'A ocorrência já foi registrada' : 'O anexo já foi registrado') + (quem ? ' por ' + quem : '') + '.' };
    }

    // data de hoje (sem hora): a do computador de quem confirmou, se coerente; senão a do fuso da planilha
    var agora = new Date();
    var hojeStr = Utilities.formatDate(agora, tz, 'yyyy-MM-dd'), dataStr = hojeStr;
    if (/^\d{4}-\d{2}-\d{2}$/.test(String(dataCliente || ''))) {
      var pc = String(dataCliente).split('-'), ph = hojeStr.split('-');
      var dif = Math.abs(serialDia_(+pc[0], +pc[1], +pc[2]) - serialDia_(+ph[0], +ph[1], +ph[2]));
      if (dif <= 1) dataStr = dataCliente;
    }
    var pd = dataStr.split('-'), dataBR = pd[2] + '/' + pd[1] + '/' + pd[0];
    sh.getRange(linha, colNome).setValue(usr.nome);
    var cData = sh.getRange(linha, colNome + 1);
    cData.setNumberFormat('dd/MM/yyyy');
    cData.setValue(serialDia_(+pd[0], +pd[1], +pd[2]));

    // cópia no sistema (planilha do período)
    var aviso = '';
    try {
      var ss = abrirPlanilhaPeriodo_(periodo);
      var lg = ss.getSheetByName('Ocorrencias');
      if (!lg) {
        lg = ss.insertSheet('Ocorrencias');
        lg.appendRow(['ID', 'Data/Hora', 'IES', 'Linha', 'CPF', 'Aluno', 'Etapa', 'Colaborador', 'Data informada', 'Observação']);
        lg.setFrozenRows(1);
      }
      var rg = lg.getRange(lg.getLastRow() + 1, 1, 1, 10);
      rg.setNumberFormat('@');
      rg.setValues([[String(agora.getTime()), Utilities.formatDate(agora, tz, 'dd/MM/yyyy HH:mm:ss'), conf.rotulo, linha,
                     String(row[0]), String(row[1]), etapa, usr.nome, dataBR, String(obs || '').trim()]]);
    } catch(e2) { aviso = 'Registrado na planilha de ocorrências, mas não foi possível gravar a cópia no sistema: ' + e2.message; }

    return { success: true, etapa: etapa, colaborador: usr.nome, data: dataBR, aviso: aviso };
  } catch(e) { return { success: false, message: e.message }; }
  finally { try { lock.releaseLock(); } catch(x) {} }
}
