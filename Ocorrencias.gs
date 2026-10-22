// OCORRÊNCIAS / ANEXOS (UNILEÃO e UNIVS): links das planilhas, leitura para os Indicadores, lista de pendentes, responsável pelo aluno,
// sincronização entre colaboradores e confirmação (grava na planilha de ocorrências e na aba "Ocorrencias" do período).
// Telas correspondentes: html_modal_ocorrencias.html e js_ocorrencias.html.
// ══════════════════════════════════════════════════════════════════════
// OCORRÊNCIAS / ANEXOS (UNILEÃO e UNIVS) — links das planilhas, pendentes, responsável, sincronização e confirmação.
// ══════════════════════════════════════════════════════════════════════

// ── Planilhas de ocorrências/anexos (registro manual; apenas leitura) ────────
// Aba "Periodos": C-ID da planilha UNILEÃO | D-ID da planilha UNIVS
// Em cada planilha (1ª aba): Ocorrências → nome na col C, data na col D | Anexos → nome na col E, data na col F
// Linha inicial: UNILEÃO = 6 | UNIVS = 8
// Extrai o ID de um link do Google Sheets (ou devolve o texto se já for o ID).
function extrairIdPlanilha_(txt) {
  txt = String(txt || '').trim();
  var m = txt.match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : txt;
}

// Monta o link completo da planilha a partir do ID.
function urlPlanilha_(id) { return id ? 'https://docs.google.com/spreadsheets/d/' + id : ''; }

// Localiza a linha do período na aba Periodos (para ler/gravar os IDs das planilhas de ocorrências).
function linhaPeriodo_(periodo) {
  var sh = getPeriodosSheet_();
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (formatPeriodo_(data[i][0]) === String(periodo).trim()) return { sh: sh, linha: i + 1, data: data[i] };
  }
  return null;
}

// IDs das planilhas UNILEÃO e UNIVS do período (colunas C e D da aba Periodos).
function getLinksPeriodo_(periodo) {
  var r = linhaPeriodo_(periodo);
  if (!r) return { unileao: '', univs: '' };
  return { unileao: String(r.data[2] || '').trim(), univs: String(r.data[3] || '').trim() };
}

// Admin: devolve os links das planilhas de ocorrências do período (para a tela de Configurações).
function getLinksOcorrencias(periodo, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin.' };
    var l = getLinksPeriodo_(periodo);
    return { success: true, unileao: urlPlanilha_(l.unileao), univs: urlPlanilha_(l.univs) };
  } catch(e) { return { success: false, message: e.message }; }
}

// Admin: confere se as planilhas abrem e grava os IDs em Periodos.
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

// Chave única (IES, linha, CPF, etapa) de uma confirmação feita pelo sistema.
function chaveLogOcorr_(ies, linha, cpf, etapa) { return ies + '|' + linha + '|' + cpfNum_(cpf) + '|' + etapa; }

// Lê a aba "Ocorrencias" do período e devolve o conjunto de etapas já confirmadas pelo sistema.
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

// ── Responsável pelo aluno e sincronização rápida entre colaboradores ────────
// O estado de cada período/planilha fica num único item do cache (lido em milissegundos, sem tocar na planilha):
//   c = responsáveis  { "linha|cpf": { id, nome, desde, tipo: 'aberto' | 'atribuido', aberto: validade da janela aberta, ate: validade } }
//   d = etapas confirmadas há pouco  { "linha|cpf": { oc: {nome,data,t}, an: {nome,data,t} } }  → os outros veem na hora
// A planilha continua sendo a fonte da verdade; o cache só acelera a visualização.
//   tipo 'aberto'    = alguém está com a janela do aluno aberta (expira em 2,5 min sem sinal de vida)
//   tipo 'atribuido' = aluno reservado para o colaborador (vale 3 h, renovado enquanto ele usa a tela; some ao concluir ou liberar)
// Tempos de validade (segundos): janela aberta, aluno atribuído e estado guardado no cache.
var OC_ABERTO_TTL = 150;     // s
var OC_ATRIB_TTL = 10800;    // s
var OC_ST_TTL = 21600;       // s (máximo do cache: 6 h)

// Chave "linha|cpf" de um aluno na lista de pendentes.
function ocItemKey_(linha, cpf) { return String(linha) + '|' + cpfNum_(cpf); }

// Chave de cache do estado de um período/planilha.
function ocStKey_(periodo, ies) { return 'ocs|' + periodo + '|' + ies; }

// Lê o estado do cache (responsáveis c e etapas concluídas d) já descartando o que expirou.
function ocStLer_(cache, chave, agora) {
  var st = { c: {}, d: {} };
  try {
    var v = cache.get(chave);
    if (v) { var o = JSON.parse(v); st.c = o.c || {}; st.d = o.d || {}; }
  } catch (e) {}
  Object.keys(st.c).forEach(function(k) { if (!(st.c[k].ate > agora)) delete st.c[k]; });
  Object.keys(st.d).forEach(function(k) {
    var d = st.d[k], t = Math.max((d.oc && d.oc.t) || 0, (d.an && d.an.t) || 0);
    if (agora - t > OC_ST_TTL * 1000) delete st.d[k];
  });
  return st;
}

// Grava o estado no cache (se passar de ~90 KB, descarta as etapas concluídas).
function ocStGravar_(cache, chave, st) {
  var txt = JSON.stringify({ c: st.c, d: st.d });
  if (txt.length > 90000) txt = JSON.stringify({ c: st.c, d: {} });      // o cache limita cada item a ~100 KB
  cache.put(chave, txt, OC_ST_TTL);
}

// Formato público do responsável enviado ao navegador.
function ocRespPub_(c, agora) {
  return { id: String(c.id), nome: c.nome, desde: c.desde, tipo: c.tipo, aberto: !!(c.aberto && c.aberto > agora) };
}

// Hora atual HH:mm no fuso do script.
function ocHora_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm'); }

// Janela de um aluno. acao: 'abrir' | 'bater' (renova) | 'forcar' (assume mesmo ocupado) | 'liberar' (fechou a janela) | 'soltar'
// Fechar a janela NÃO tira um aluno atribuído ao colaborador; só 'soltar' (ou concluir) libera.
function ocClaim(periodo, ies, linha, cpf, acao, userId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);
    var usr = usuarioPorId_(userId);
    if (!usr) return { success: false, message: 'Usuário não identificado.' };
    var cache = CacheService.getScriptCache(), sk = ocStKey_(periodo, ies), agora = Date.now();
    var st = ocStLer_(cache, sk, agora), k = ocItemKey_(linha, cpf), atual = st.c[k];
    var meu = !!(atual && String(atual.id) === String(userId));
    if (acao === 'liberar') {
      if (meu) { if (atual.tipo === 'atribuido') atual.aberto = 0; else delete st.c[k]; ocStGravar_(cache, sk, st); }
      return { success: true };
    }
    if (acao === 'soltar') { if (meu) { delete st.c[k]; ocStGravar_(cache, sk, st); } return { success: true }; }
    if (atual && !meu && acao !== 'forcar') return { success: false, ocupado: true, resp: ocRespPub_(atual, agora) };
    var aberto = agora + OC_ABERTO_TTL * 1000;
    var c = meu ? atual : { id: String(userId), nome: usr.nome, desde: ocHora_(), tipo: 'aberto' };
    c.aberto = aberto;
    c.ate = c.tipo === 'atribuido' ? agora + OC_ATRIB_TTL * 1000 : aberto;
    st.c[k] = c;
    ocStGravar_(cache, sk, st);
    return { success: true, resp: ocRespPub_(c, agora) };
  } catch(e) { return { success: false, message: e.message }; }
  finally { try { lock.releaseLock(); } catch(x) {} }
}

// Atribui (ou libera) vários alunos de uma vez ao colaborador. itens: [{linha, cpf}]
// modo: 'assumir' (pula quem já está com outra pessoa) | 'forcar' (assume mesmo assim) | 'soltar' (libera os que são dele)
function ocAtribuir(periodo, ies, itens, userId, modo) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var usr = usuarioPorId_(userId);
    if (!usr) return { success: false, message: 'Usuário não identificado.' };
    if (modo !== 'assumir' && modo !== 'forcar' && modo !== 'soltar') return { success: false, message: 'Ação inválida.' };
    if (!OCORR_IES[ies]) return { success: false, message: 'Planilha inválida.' };
    var cache = CacheService.getScriptCache(), sk = ocStKey_(periodo, ies), agora = Date.now();
    var st = ocStLer_(cache, sk, agora), ok = 0, soltos = 0, ocupados = [], desde = ocHora_();
    (itens || []).slice(0, 500).forEach(function(it) {
      var k = ocItemKey_(it.linha, it.cpf), atual = st.c[k], meu = !!(atual && String(atual.id) === String(userId));
      if (modo === 'soltar') { if (meu) { delete st.c[k]; soltos++; } return; }
      if (atual && !meu && modo !== 'forcar') { ocupados.push({ linha: it.linha, nome: atual.nome }); return; }
      if (meu) { atual.tipo = 'atribuido'; atual.ate = agora + OC_ATRIB_TTL * 1000; ok++; return; }
      st.c[k] = { id: String(userId), nome: usr.nome, desde: desde, tipo: 'atribuido', aberto: 0, ate: agora + OC_ATRIB_TTL * 1000 };
      ok++;
    });
    ocStGravar_(cache, sk, st);
    return { success: true, ok: ok, soltos: soltos, ocupados: ocupados };
  } catch(e) { return { success: false, message: e.message }; }
  finally { try { lock.releaseLock(); } catch(x) {} }
}

// Sincronização leve (chamada a cada poucos segundos): quem está com cada aluno da lista e o que já foi concluído por outras pessoas.
// Só lê o cache; renova as atribuições de quem está sincronizando.  itens: [{linha, cpf}]
function ocSync(periodo, ies, itens, userId) {
  try {
    if (!OCORR_IES[ies]) return { success: false, message: 'Planilha inválida.' };
    var cache = CacheService.getScriptCache(), sk = ocStKey_(periodo, ies), agora = Date.now();
    var st = ocStLer_(cache, sk, agora);
    var renovar = Object.keys(st.c).some(function(k) {
      var c = st.c[k];
      return c.tipo === 'atribuido' && String(c.id) === String(userId) && c.ate - agora < OC_ATRIB_TTL * 500;
    });
    if (renovar) {
      var lock = LockService.getScriptLock();
      try {
        lock.waitLock(3000);
        st = ocStLer_(cache, sk, agora);
        Object.keys(st.c).forEach(function(k) {
          var c = st.c[k];
          if (c.tipo === 'atribuido' && String(c.id) === String(userId)) c.ate = agora + OC_ATRIB_TTL * 1000;
        });
        ocStGravar_(cache, sk, st);
      } catch(e0) {} finally { try { lock.releaseLock(); } catch(x) {} }
    }
    var resp = {}, feitos = {};
    (itens || []).forEach(function(it) {
      var k = ocItemKey_(it.linha, it.cpf), c = st.c[k], d = st.d[k];
      if (c) resp[it.linha] = ocRespPub_(c, agora);
      if (d) feitos[it.linha] = { oc: d.oc ? { nome: d.oc.nome, data: d.oc.data } : null, an: d.an ? { nome: d.an.nome, data: d.an.data } : null };
    });
    return { success: true, resp: resp, feitos: feitos };
  } catch(e) { return { success: false, message: e.message }; }
}

// Quem está com cada aluno da lista agora (mantida por compatibilidade). itens: [{linha, cpf}]  →  { linha: {nome, desde} }
function getOcorrenciasResponsaveis(periodo, ies, itens) {
  var r = ocSync(periodo, ies, itens, '');
  return r.success ? { success: true, resp: r.resp } : r;
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
    var agoraP = Date.now(), stP = ocStLer_(CacheService.getScriptCache(), ocStKey_(periodo, ies), agoraP);
    itens.forEach(function(it) { var c = stP.c[ocItemKey_(it.linha, it.cpf)]; it.resp = c ? ocRespPub_(c, agoraP) : null; });
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
    var cacheOc = CacheService.getScriptCache(), skOc = ocStKey_(periodo, ies), kOc = ocItemKey_(linha, cpf);
    var cl = ocStLer_(cacheOc, skOc, Date.now()).c[kOc];
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

    // avisa os outros colaboradores na hora (cache) e libera o aluno quando não restar nenhuma etapa pendente
    try {
      var stOc = ocStLer_(cacheOc, skOc, Date.now());
      var dm = stOc.d[kOc] || (stOc.d[kOc] = {});
      dm[etapa === 'Ocorrência' ? 'oc' : 'an'] = { nome: usr.nome, data: dataBR, t: Date.now() };
      var restam = etapa === 'Ocorrência' ? (!semAnexo_(row[13]) && vazio_(row[5])) : vazio_(row[3]);
      var meuCl = stOc.c[kOc];
      if (!restam && meuCl && String(meuCl.id) === String(userId)) delete stOc.c[kOc];
      ocStGravar_(cacheOc, skOc, stOc);
    } catch(e3) {}

    return { success: true, etapa: etapa, colaborador: usr.nome, data: dataBR, aviso: aviso };
  } catch(e) { return { success: false, message: e.message }; }
  finally { try { lock.releaseLock(); } catch(x) {} }
}
