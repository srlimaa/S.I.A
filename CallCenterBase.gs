// CALL CENTER: abas, usuários unificados, semestres, cache e leitura de dados (funções "cc...").
// Todos os dados do Call Center ficam na MESMA estrutura do Atendimentos: planilha MESTRE (abas CC_USUARIOS, CC_CONFIG, CC_OPCOES e colunas E e F da aba Periodos)
// e planilha de cada PERÍODO (abas CC_REGISTROS e CC_RESPONSAVEIS). A planilha antiga do Call Center só é usada pela migração (CallCenterMigracao.gs).
// Desempenho: memória da execução (_CC_MEM) + cache compartilhado (CacheService) com a base de alunos compactada.
// ══════════════════════════════════════════════════════════════════════
// CALL CENTER — planilha, login, utilitários e leitura de dados (funções "cc...").
// ══════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
// CALL CENTER - ADITAMENTOS (integrado a este projeto)
// Todos os dados do Call Center ficam na MESMA estrutura do sistema de Atendimentos:
//   • Planilha MESTRE (SPREADSHEET_ID): abas CC_USUARIOS (tema e semestre de cada pessoa), CC_CONFIG e CC_OPCOES, e as colunas
//     E (Base Geral) e F (Planilha Call Center) da aba Periodos — é a lista de semestres do Call Center.
//   • Planilha de cada PERÍODO (a mesma onde o Atendimentos registra): abas CC_REGISTROS e CC_RESPONSAVEIS.
// Assim, cadastrar um período no Atendimentos já o torna disponível no Call Center; lá só se informam os dois links.
// A tela do Call Center está no arquivo HTML "CallCenter". Os nomes das funções começam com "cc".
// ═══════════════════════════════════════════════════════════════════════════
// Planilha ANTIGA do Call Center: usada só pela migração (CallCenterMigracao.gs)
// ID da planilha ANTIGA do Call Center. Hoje só é usado pela migração (CallCenterMigracao.gs).
var CC_PLANILHA_PADRAO = '1L5wclKLy7MZvVhDgR3KUDPla2Q7GI50FifWv0VXN480';

/* ---------- velocidade: memória da execução + cache compartilhado ----------
   Cada chamada do navegador é uma execução nova do script. Dentro dela, planilhas e abas já lidas ficam em memória
   (_CC_MEM) em vez de serem abertas/lidas de novo. Entre execuções, o CacheService guarda a base de alunos (compactada). */
// Memória da execução atual: evita abrir/ler a mesma planilha várias vezes.
var _CC_MEM = {};

// Executa f uma única vez por chave durante a execução.
function ccMem_(k, f) { if (!(k in _CC_MEM)) _CC_MEM[k] = f(); return _CC_MEM[k]; }

// Descarta da memória as chaves com o prefixo informado.
function ccMemLimpa_(prefixo) { Object.keys(_CC_MEM).forEach(function(k) { if (k.indexOf(prefixo) === 0) delete _CC_MEM[k]; }); }

// Abre uma planilha pelo ID (com memória).
function ccOpen_(id) { return ccMem_('ss|' + id, function() { return SpreadsheetApp.openById(id); }); }

// Abre a planilha mestre do Atendimentos.
function ccMestre_() { return ccMem_('mestre', function() { return SpreadsheetApp.openById(SPREADSHEET_ID); }); }

// Depois de gravar em uma aba, descarta a cópia em memória dela
function ccInval_(ss, aba) { ccMemLimpa_('tab|' + ss.getId() + '|' + aba); }

// Guarda um objeto no cache compartilhado (JSON compactado e dividido em partes de ~90 KB). Falha em silêncio.
function ccCachePut_(chave, obj, seg) {
  try {
    var c = CacheService.getScriptCache();
    var b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify(obj), 'application/octet-stream')).getBytes());
    var TAM = 90000, n = Math.ceil(b64.length / TAM), tudo = {};
    for (var i = 0; i < n; i++) tudo[chave + '_' + i] = b64.substr(i * TAM, TAM);
    tudo[chave + '_n'] = String(n);
    c.putAll(tudo, seg || 600);
    return true;
  } catch (e) { return false; }
}

// Lê do cache um objeto guardado em partes (ccCachePut_); null se faltar alguma parte.
function ccCacheGet_(chave) {
  try {
    var c = CacheService.getScriptCache(), n = parseInt(c.get(chave + '_n'), 10);
    if (!n) return null;
    var ks = [];
    for (var i = 0; i < n; i++) ks.push(chave + '_' + i);
    var m = c.getAll(ks), b64 = '';
    for (var j = 0; j < n; j++) { var p = m[ks[j]]; if (p == null) return null; b64 += p; }
    return JSON.parse(Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString());
  } catch (e) { return null; }
}

// "Versão" dos dados vivos do semestre (registros e responsáveis). Muda a cada gravação; as telas abertas consultam
// essa versão (chamada leve) e só baixam os dados quando ela mudou.
// Chave da versão dos dados de um semestre.
function ccVerChave_(sem) { return 'cc_ver_' + String(sem).replace(/[^\w.-]/g, ''); }

// Versão atual dos dados do semestre (muda a cada gravação; as telas conferem e só baixam se mudou).
function ccVer_(sem) {
  var c = CacheService.getScriptCache(), k = ccVerChave_(sem), v = c.get(k);
  if (!v) { v = Utilities.getUuid().slice(0, 8); c.put(k, v, 21600); }
  return v;
}

// Muda a versão do semestre (chamar depois de gravar).
function ccBump_(sem) { CacheService.getScriptCache().put(ccVerChave_(sem), Utilities.getUuid().slice(0, 8), 21600); }

// Mantida só por compatibilidade: o Call Center agora usa a planilha mestre do sistema.
// Mantida por compatibilidade: devolve a planilha mestre (o Call Center não tem mais planilha própria).
function ccPlanilhaId_() { return SPREADSHEET_ID; }

// Planilha onde ficam as abas CC_ do Call Center: a planilha mestre.
function ccSs_() { return ccMestre_(); }

// Fuso horário do script.
function ccTz_() { return Session.getScriptTimeZone(); }

// Cabeçalhos (colunas) de cada aba lógica do Call Center. Os semestres não têm mais aba própria: vêm da aba Periodos.
const CC_HEAD = {
  USUARIOS: ['Email', 'Nome', 'Senha', 'Perfil', 'Foto', 'Tema', 'Semestre'],
  CONFIG: ['Chave', 'Valor'],
  OPCOES: ['Tipo', 'Valor'],
  REGISTROS: ['Id', 'Semestre', 'CPF', 'Aluno', 'DataHora', 'Objetivo', 'Meio', 'Resultado', 'Observacao', 'ColabEmail', 'ColabNome', 'Status', 'EditadoEm'],
  RESPONSAVEIS: ['Semestre', 'CPF', 'Email']
};

// Nome real de cada aba (prefixo CC_ para não se misturar com as abas do Atendimentos)
// Nome real de cada aba no Google Sheets (prefixo CC_ para não misturar com as abas do Atendimentos).
const CC_ABAS = { USUARIOS: 'CC_USUARIOS', CONFIG: 'CC_CONFIG', OPCOES: 'CC_OPCOES', REGISTROS: 'CC_REGISTROS', RESPONSAVEIS: 'CC_RESPONSAVEIS' };
// Abas lógicas que ficam na planilha mestre.
const CC_MESTRE = ['USUARIOS', 'CONFIG', 'OPCOES'];          // abas da planilha mestre
// Abas lógicas que ficam na planilha de cada período.
const CC_PERIODO = ['REGISTROS', 'RESPONSAVEIS'];            // abas da planilha de cada período

// Converte o nome lógico da aba (ex.: REGISTROS) no nome real (CC_REGISTROS).
function ccAbaNome_(n) { return CC_ABAS[n] || n; }

// Status que entram na lista de trabalho.
const CC_ST = ['SOLICITADO / NÃO ADITOU', 'SOLICITADO / ENVIADO AO BANCO'];

// Nome da aba de aditamento na planilha de comparecimento ao banco.
const CC_ABA = 'Aditamento';      // aba da planilha "Call Center - Aditamentos"

// Coluna (índice 0) com o período do banco na planilha de comparecimento.
const CC_COL_BANCO = 45;             // coluna AT (índice 0)

// Cria, na planilha mestre, as abas do Call Center que ainda não existirem. vazio=true: sem as opções/configurações padrão (migração).
// Cria, na planilha mestre, as abas CC_ que faltam (com cabeçalho). vazio=true: sem as configurações e opções padrão (usado na migração).
function ccEnsure_(vazio) {
  const ss = ccSs_();
  CC_MESTRE.forEach(n => {
    const nome = ccAbaNome_(n);
    let sh = ss.getSheetByName(nome);
    if (sh) return;
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, CC_HEAD[n].length).setValues([CC_HEAD[n]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), CC_HEAD[n].length).setNumberFormat('@');
    if (vazio) return;
    if (n === 'CONFIG') { sh.appendRow(['nomeLogin', 'Call Center']); sh.appendRow(['nomeSistema', 'Call Center - Aditamentos']); sh.appendRow(['logo', '']); }
    if (n === 'OPCOES') {
      ['Informar sobre o aditamento', 'Orientar comparecimento ao banco', 'Cobrar documentação', 'Verificar andamento'].forEach(v => sh.appendRow(['Objetivo', v]));
      ['Ligação', 'WhatsApp', 'E-mail'].forEach(v => sh.appendRow(['Meio', v]));
      ['Contato realizado', 'Não atendeu', 'Número inválido', 'Retornar depois', 'Aluno vai comparecer', 'Aluno já compareceu'].forEach(v => sh.appendRow(['Resultado', v]));
    }
  });
  ccMemLimpa_('tab|');
}

/* ---------- semestres = aba "Periodos" do Atendimentos ----------
   Colunas: A Período | B ID da planilha do período | C/D ocorrências | E Base Geral (Call Center) | F Planilha Call Center */
// Abre a aba Periodos garantindo as colunas E (Base Geral) e F (Planilha Call Center) com cabeçalho.
function ccPerSheet_() {
  return ccMem_('persh', function() {
    const sh = getPeriodosSheet_();
    if (sh.getMaxColumns() < 6) sh.insertColumnsAfter(sh.getMaxColumns(), 6 - sh.getMaxColumns());
    const h = sh.getRange(1, 5, 1, 2).getValues()[0];
    if (String(h[0]).trim() !== 'Base Geral (Call Center)' || String(h[1]).trim() !== 'Planilha Call Center') {
      sh.getRange(1, 5, 1, 2).setValues([['Base Geral (Call Center)', 'Planilha Call Center']]).setFontWeight('bold');
      sh.getRange(2, 5, Math.max(1, sh.getMaxRows() - 1), 2).setNumberFormat('@');
    }
    return sh;
  });
}

// Semestres do Call Center (do mais antigo ao mais recente): { Semestre, IdBase, IdCallCenter, IdPeriodo, _row }
// Semestres do Call Center = períodos letivos do Atendimentos, com os links de Base Geral e Call Center (do mais antigo ao mais recente).
function ccSems_() {
  return ccMem_('sems', function() {
    const v = ccPerSheet_().getDataRange().getValues(), out = [];
    for (let i = 1; i < v.length; i++) {
      const p = formatPeriodo_(v[i][0]);
      if (!p || !String(v[i][1]).trim()) continue;
      out.push({ Semestre: p, IdBase: ccId_(v[i][4]), IdCallCenter: ccId_(v[i][5]), IdPeriodo: ccId_(v[i][1]), _row: i + 1 });
    }
    out.sort((a, b) => a.Semestre < b.Semestre ? -1 : a.Semestre > b.Semestre ? 1 : 0);
    return out;
  });
}

// Só os nomes dos semestres.
function ccSemNomes_() { return ccSems_().map(s => String(s.Semestre)); }

// Grava os links da Base Geral e do Call Center de um semestre (o semestre precisa existir em Períodos letivos)
// Grava os links da Base Geral e do Call Center de um semestre (ele precisa existir em Períodos letivos). Links vazios = limpar.
function ccSemSalvar_(sem, base, cc) {
  const r = ccSems_().find(x => String(x.Semestre) === String(sem));
  if (!r) throw new Error('O semestre ' + sem + ' ainda não está cadastrado em Configurações › Períodos letivos (sistema de Atendimentos). Cadastre-o lá primeiro.');
  const g = ccPerSheet_().getRange(r._row, 5, 1, 2);
  g.setNumberFormat('@'); g.setValues([[ccId_(base), ccId_(cc)]]);
  ccMemLimpa_('sems');
}

/* ---------- utilidades de planilha ---------- */

// Devolve a aba (pelo nome lógico), criando sozinha as abas do Call Center que ainda não existirem
// Devolve a aba pelo nome lógico, criando sozinha as abas do Call Center que ainda não existirem.
function ccAba_(n, s0) {
  let sh = s0.getSheetByName(ccAbaNome_(n));
  if (!sh) {
    if (CC_MESTRE.indexOf(n) >= 0) ccEnsure_(); else if (CC_PERIODO.indexOf(n) >= 0) ccPrepararRegistros_(s0);
    sh = s0.getSheetByName(ccAbaNome_(n));
  }
  if (!sh) throw new Error('Aba "' + ccAbaNome_(n) + '" não encontrada na planilha.');
  return sh;
}

// Lê uma aba (pelo nome lógico, via ccAba_) como lista de objetos com _row = linha na planilha; com memória; datas viram texto.
function ccTab_(n, ss) {
  const s0 = ss || ccSs_();
  return ccMem_('tab|' + s0.getId() + '|' + n, function() {
    const sh = ccAba_(n, s0);
    const v = sh.getDataRange().getValues(), h = v.shift();
    return v.filter(r => r.join('') !== '').map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = (r[j] instanceof Date) ? Utilities.formatDate(r[j], ccTz_(), k === 'Semestre' ? 'yyyy.M' : 'yyyy-MM-dd HH:mm') : r[j]); return o; });
  });
}

// Acrescenta uma linha (sempre como texto), ampliando a aba se necessário, e descarta a memória da aba.
function ccAdd_(n, o, ss) { const s0 = ss || ccSs_(), sh = ccAba_(n, s0), r = sh.getLastRow() + 1; if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1); const g = sh.getRange(r, 1, 1, CC_HEAD[n].length); g.setNumberFormat('@'); g.setValues([CC_HEAD[n].map(k => o[k] === undefined ? '' : String(o[k]))]); ccInval_(s0, n); }

// Altera campos de uma linha existente.
function ccUpd_(n, row, o, ss) { const s0 = ss || ccSs_(), sh = ccAba_(n, s0); CC_HEAD[n].forEach((k, j) => { if (o[k] !== undefined) { const c = sh.getRange(row, j + 1); c.setNumberFormat('@'); c.setValue(String(o[k])); } }); ccInval_(s0, n); }

// Exclui uma linha.
function ccDel_(n, row) { const s0 = ccSs_(); ccAba_(n, s0).deleteRow(row); ccInval_(s0, n); }

// Extrai o ID de um link do Google Sheets.
function ccId_(s) { s = String(s || '').trim(); const m = s.match(/\/d\/([a-zA-Z0-9-_]+)/); return m ? m[1] : s; }

// Chave de comparação de textos (sem acento, maiúsculas, espaços padronizados).
function ccKey_(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/\s*\/\s*/g, ' / ').replace(/\s+/g, ' ').trim(); }

// Os 7 status possíveis do aditamento na base.
const CC_STATUS7 = ['ADITOU', 'NÃO SOLICITADO', 'REJEITADO', 'SOLICITADO / CANCELADO', 'SOLICITADO / EM PROCESSAMENTO', 'SOLICITADO / ENVIADO AO BANCO', 'SOLICITADO / NÃO ADITOU'];

// Converte um status para a grafia oficial.
function ccCanon_(s) { const k = ccKey_(s); return CC_STATUS7.find(x => ccKey_(x) === k) || k; }

// Padroniza texto em maiúsculas.
function ccNorm_(s) { return String(s || '').toUpperCase().replace(/\s+/g, ' ').replace(/\s*\/\s*/g, ' / ').trim(); }

// CPF só com dígitos, completado com zeros a esquerda (11 dígitos).
function ccCpf_(s) { const d = String(s || '').replace(/\D/g, ''); return d ? ('00000000000' + d).slice(-11) : ''; }

// Formata data como DD/MM/AAAA.
function ccFmtD_(v) { if (v instanceof Date) return Utilities.formatDate(v, ccTz_(), 'dd/MM/yyyy'); return String(v || '').trim(); }

// Formata percentual.
function ccFmtP_(v) { if (typeof v === 'number') return (v <= 1 ? Math.round(v * 10000) / 100 : v) + '%'; return String(v || '').trim(); }

/* ---------- senha ---------- */

// Hash SHA-256 com sal (resto do cadastro antigo do Call Center; o login atual é o do Atendimentos).
function ccHash_(s, salt) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + s).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }

function ccHashNew_(s) { const salt = Utilities.getUuid().slice(0, 8); return salt + '$' + ccHash_(s, salt); }

function ccCheck_(s, stored) { const p = String(stored).split('$'); return p.length === 2 && ccHash_(s, p[0]) === p[1]; }

/* ---------- login / sessão ---------- */

// Dados públicos (nome do login e logo) para o uso isolado da tela.
function ccGetPublic() { ccEnsure_(); const c = ccCfg_(); return { nomeLogin: c.nomeLogin, logo: c.logo }; }

// Login do Call Center com e-mail e senha do Atendimentos; exige acesso liberado.
function ccApiLogin(email, senha) {
  // login único: o usuário é o do sistema de Atendimentos; precisa ter acesso liberado ao Call Center
  const r = login(email, senha);
  if (!r || !r.success) return { err: 'E-mail ou senha inválidos.' };
  const c = ccLoginInterno_(r.email);
  if (!c) return { err: 'Seu usuário não tem acesso ao Call Center. Peça a um Admin para liberar.' };
  return { ok: { token: c.token } };
}

// Encerra a sessão do Call Center (chave s_ no cache).
function ccApiLogout(t) { CacheService.getScriptCache().remove('s_' + t); }

// Configuração do Call Center como objeto.
function ccCfg_() { const o = {}; ccTab_('CONFIG').forEach(r => o[r.Chave] = r.Valor); return o; }

// Opções de Objetivo, Meio e Resultado.
function ccOpc_() { const o = { Objetivo: [], Meio: [], Resultado: [] }; ccTab_('OPCOES').forEach(r => { if (o[r.Tipo]) o[r.Tipo].push(r.Valor); }); return o; }

// Dados públicos do usuário enviados à tela.
function ccPub_(u) { return { email: u.Email, nome: u.Nome, perfil: u.Perfil, foto: u.Foto, tema: u.Tema || 'Padrão', semestre: ccSemU_(u) }; }

// Semestre do usuário (o último cadastrado se o dele não existir mais).
function ccSemU_(u) { const s = ccSemNomes_(); return s.indexOf(String(u.Semestre)) >= 0 ? String(u.Semestre) : s[s.length - 1]; }

/* ---------- usuários unificados ----------
   Os usuários são os do sistema de Atendimentos (aba "Usuarios" da planilha mestre). A coluna H dessa aba ("CallCenter",
   Sim/Não) libera o acesso. Nome, perfil (Admin/Colab) e foto vêm de lá. A aba USUARIOS da planilha do Call Center
   guarda só as preferências de cada pessoa (Tema e Semestre) e é preenchida sozinha no primeiro acesso. */
function ccFlagAcesso_(flag, temPrefs) {
  var f = String(flag == null ? '' : flag).trim().toLowerCase();
  if (f === 'sim' || f === 's' || f === 'true') return true;
  if (f === 'nao' || f === 'não' || f === 'n' || f === 'false') return false;
  return !!temPrefs;      // coluna em branco: quem já usava o Call Center continua com acesso
}

// Usuários do sistema de Atendimentos
// Usuários do Atendimentos (com memória).
function ccAtend_() { return ccMem_('atend', ccAtendLer_); }

// Lê a aba Usuarios da planilha mestre.
function ccAtendLer_() {
  var data = ccMestre_().getSheetByName('Usuarios').getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var email = String(data[i][2] == null ? '' : data[i][2]).trim();
    if (!email) continue;
    out.push({ id: String(data[i][0]), nome: String(data[i][1] == null ? '' : data[i][1]).trim(), email: email,
               perfil: String(data[i][4] == null ? '' : data[i][4]).trim(),
               avatar: { tipo: String(data[i][5] || '').trim(), valor: String(data[i][6] || '') }, flag: data[i][7] });
  }
  return out;
}

// Foto no formato da tela do Call Center: 'emoji:X' ou imagem (data URL)
// Converte o avatar para o formato da tela do Call Center.
function ccFoto_(av) {
  if (av && av.tipo === 'emoji' && av.valor) return 'emoji:' + av.valor;
  if (av && av.tipo === 'imagem' && av.valor) return av.valor;
  return '';
}

// E-mails que já têm preferências na aba USUARIOS do Call Center.
function ccPrefsEmails_() {
  try { var o = {}; ccTab_('USUARIOS').forEach(function(x) { o[String(x.Email).toLowerCase()] = 1; }); return o; }
  catch (e) { return {}; }
}

// Mapa e-mail > nome atual (do Atendimentos).
function ccNomesPorEmail_() {
  var o = {};
  try { ccAtend_().forEach(function(a) { o[a.email.toLowerCase()] = a.nome; }); } catch (e) {}
  return o;
}

// Usuário do Call Center = usuário do Atendimentos com acesso liberado (null se não existe ou não tem acesso)
function ccUser_(email) {
  var em = String(email || '').trim().toLowerCase();
  if (!em) return null;
  var a = ccAtend_().find(function(x) { return x.email.toLowerCase() === em; });
  if (!a) return null;
  var find = function() { return ccTab_('USUARIOS').find(function(x) { return String(x.Email).toLowerCase() === em; }); };
  var p = find();
  if (!ccFlagAcesso_(a.flag, !!p)) return null;
  var perfil = a.perfil === 'Admin' ? 'Admin' : 'Colab';
  if (!p) { ccAdd_('USUARIOS', { Email: em, Nome: a.nome, Perfil: perfil, Tema: 'Padrão' }); p = find(); }
  return { Email: em, Nome: a.nome, Perfil: perfil, Foto: ccFoto_(a.avatar), Tema: p.Tema || 'Padrão', Semestre: p.Semestre || '', _row: p._row, _p: p };
}

// Lista de usuários com acesso (para responsáveis e cartões) + nomes/fotos de quem aparece nos registros mas não tem mais acesso
function ccUsuariosInfo_(emails) {
  var atend = ccAtend_(), tem = ccPrefsEmails_(), usuarios = [], vistos = {}, mapa = {};
  atend.forEach(function(a) {
    var em = a.email.toLowerCase();
    mapa[em] = a;
    if (!ccFlagAcesso_(a.flag, tem[em])) return;
    vistos[em] = 1;
    usuarios.push({ email: em, nome: a.nome, foto: ccFoto_(a.avatar), perfil: a.perfil === 'Admin' ? 'Admin' : 'Colab' });
  });
  usuarios.sort(function(x, y) { return x.nome.localeCompare(y.nome, 'pt', { sensitivity: 'base' }); });
  var extras = [], prefs = null;
  (emails || []).forEach(function(e) {
    var em = String(e || '').trim().toLowerCase();
    if (!em || vistos[em]) return;
    vistos[em] = 1;
    if (mapa[em]) { extras.push({ email: em, nome: mapa[em].nome, foto: ccFoto_(mapa[em].avatar) }); return; }
    if (prefs === null) { try { prefs = ccTab_('USUARIOS'); } catch (x) { prefs = []; } }
    var p = prefs.find(function(q) { return String(q.Email).toLowerCase() === em; });
    if (p) extras.push({ email: em, nome: String(p.Nome || em), foto: String(p.Foto || '') });
  });
  return { usuarios: usuarios, extras: extras };
}

/* ---------- planilha de registros de cada semestre ----------
   Os registros ficam na planilha do próprio período letivo (a mesma onde o Atendimentos registra), nas abas
   CC_REGISTROS e CC_RESPONSAVEIS, criadas sozinhas quando faltam. */
// Planilha do PERÍODO onde ficam CC_REGISTROS e CC_RESPONSAVEIS do semestre (a mesma do Atendimentos); cria as abas se faltarem.
function ccSsReg_(sem) {
  return ccMem_('ssreg|' + sem, function() {
    const r = ccSems_().find(x => String(x.Semestre) === String(sem));
    if (!r) throw new Error('O semestre ' + sem + ' não está cadastrado nos Períodos letivos do sistema de Atendimentos.');
    const ss = ccOpen_(r.IdPeriodo);
    ccPrepararRegistros_(ss);
    return ss;
  });
}

// Cria (se ainda não existirem) as abas CC_REGISTROS e CC_RESPONSAVEIS com os cabeçalhos na planilha informada
// Cria CC_REGISTROS e CC_RESPONSAVEIS com cabeçalhos na planilha do período, se ainda não existirem.
function ccPrepararRegistros_(ss) {
  CC_PERIODO.forEach(n => {
    const nome = ccAbaNome_(n);
    if (ss.getSheetByName(nome)) return;
    const sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, CC_HEAD[n].length).setValues([CC_HEAD[n]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), CC_HEAD[n].length).setNumberFormat('@');
  });
}

// Registros de contato do semestre, com o nome atual do colaborador.
function ccRegs_(sem) {
  const nomes = ccNomesPorEmail_();       // nome atual do colaborador (cadastro de usuários do Atendimentos)
  return ccTab_('REGISTROS', ccSsReg_(sem)).filter(r => String(r.Semestre) === String(sem)).map(r => ({
    id: r.Id, cpf: r.CPF, aluno: r.Aluno, data: String(r.DataHora), objetivo: r.Objetivo, meio: r.Meio, resultado: r.Resultado,
    obs: r.Observacao, colab: String(r.ColabEmail || '').toLowerCase(), colabNome: nomes[String(r.ColabEmail || '').toLowerCase()] || r.ColabNome, status: r.Status, editado: r.EditadoEm
  }));
}

// Lê a planilha de comparecimento ao banco e devolve o período (início e fim) por CPF.
function ccComp_(idc) {
  const sc = SpreadsheetApp.openById(ccId_(idc));
  const sh = sc.getSheetByName(CC_ABA) || sc.getSheets().find(x => /aditamento/i.test(x.getName())) || sc.getSheets()[1];
  if (!sh) return {};
  const v = sh.getDataRange().getDisplayValues();
  let ci = 2;
  for (let i = 0; i < Math.min(3, v.length); i++) { const j = v[i].findIndex(c => /^cpf/i.test(String(c).trim())); if (j >= 0) { ci = j; break; } }
  const o = {};
  for (let i = 2; i < v.length; i++) {
    const k = ccCpf_(v[i][ci]), t = String(v[i][CC_COL_BANCO] || '').trim();
    if (!k || !t) continue;
    const m = t.match(/(\d{2}\/\d{2}\/\d{4})\s*(?:até|ate|a)\s*(\d{2}\/\d{2}\/\d{4})/i);
    o[k] = m ? { ini: m[1], fim: m[2] } : { ini: t, fim: '' };
  }
  return o;
}

/* ---------- API ---------- */

// Confere se as abas CC_ da planilha mestre existem.
function ccTemAbas_() {
  var ss = ccSs_();
  return CC_MESTRE.every(function(n) { return !!ss.getSheetByName(ccAbaNome_(n)); });
}

// Abre a sessão do Call Center de um usuário já autenticado; null sem acesso. Cria a sessão (chave s_, 6 h) e devolve dados do sistema.
function ccLoginInterno_(email) {
  if (!ccTemAbas_()) ccEnsure_();
  var u = ccUser_(email);
  if (!u) return null;                                  // sem cadastro ou sem acesso liberado ao Call Center
  var p = u._p || {};
  if (String(p.Nome || '') !== u.Nome || String(p.Perfil || '') !== u.Perfil || String(p.Email || '') !== u.Email)
    ccUpd_('USUARIOS', u._row, { Email: u.Email, Nome: u.Nome, Perfil: u.Perfil });   // só para ficar legível na planilha
  var t = Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + t, u.Email, 21600);
  var c = ccCfg_();
  var logo = String(c.logo || '');
  return {
    token: t, nome: u.Nome, perfil: u.Perfil, email: u.Email,
    nomeSistema: String(c.nomeSistema || 'Call Center - Aditamentos'),
    logo: /^data:image\//.test(logo) ? logo : '',
    sems: ccSemNomes_(),
    semAtual: String(ccSemU_(u) || '')
  };
}

// Define o semestre do usuário no Call Center (período escolhido no sistema integrado)
function ccAbrirPeriodo(token, periodo) {
  try {
    var email = CacheService.getScriptCache().get('s_' + token);
    var sessao = { success: false, sessao: true, message: 'Sua sessão no Call Center expirou. Entre novamente.' };
    if (!email) return sessao;
    var u = ccUser_(email);
    if (!u) return sessao;
    var p = String(periodo == null ? '' : periodo).trim();
    var sems = ccSemNomes_();
    if (sems.indexOf(p) < 0)
      return { success: false, message: 'O período ' + p + ' não está cadastrado nos Períodos letivos do sistema de Atendimentos.' };
    if (String(u.Semestre || '') !== p) ccUpd_('USUARIOS', u._row, { Semestre: p });   // só grava quando mudou
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// Conteúdo da tela do Call Center (arquivo HTML "CallCenter" do projeto)
function ccGetHtml() {
  var chave = 'cc_html_' + String(typeof VERSAO_SISTEMA === 'undefined' ? '0' : VERSAO_SISTEMA);
  var c = ccCacheGet_(chave);
  if (c && c.html) return c.html;
  var html = HtmlService.createTemplateFromFile('CallCenter').evaluate().getContent();
  ccCachePut_(chave, { html: html }, 21600);
  return html;
}

// Leitura pesada da base de alunos do semestre (aba BASE GERAL + comparecimento ao banco).
// O resultado fica em cache por 10 min (ou até um Admin pedir "Recarregar base").
function ccBase_(r, sem, fresh) {
  const chave = 'cb_' + String(sem).replace(/[^\w.-]/g, '') + '_' + ccId_(r.IdBase).slice(-8) + ccId_(r.IdCallCenter).slice(-8);
  if (!fresh) { const c = ccCacheGet_(chave); if (c && c.alunos) return c; }
  const sb = ccOpen_(ccId_(r.IdBase));
  const bsh = sb.getSheetByName('BASE GERAL') || sb.getSheets()[0];
  const base = bsh.getDataRange().getValues();
  const h0 = 1; // cabeçalhos na linha 2; dados a partir da linha 3
  const comp = r.IdCallCenter ? ccComp_(r.IdCallCenter) : {};
  const alunos = [], light = [], st0 = {};
  let tot = 0;
  for (let i = h0 + 1; i < base.length; i++) {
    const v = base[i], nome = String(v[3]).trim();
    if (!nome) continue;
    const st = ccCanon_(v[40]);
    st0[st] = (st0[st] || 0) + 1;
    tot++;
    // Regra do Indicador - Geral: coluna AB = Matriculado e coluna AL contendo a palavra FIES
    const mat = /^MATRICULAD[OA]\b/.test(ccKey_(v[27])), fies = ccKey_(v[37]).indexOf('FIES') >= 0;
    if (mat && fies) light.push([st, String(v[4]).trim(), String(v[5]).trim()]);
    if (CC_ST.indexOf(st) < 0) continue;
    const cpf = ccCpf_(v[2]);
    const al = {
      cpf: cpf, cpfFmt: String(v[2]).trim(), nome: nome, curso: String(v[4]).trim(), turno: String(v[5]).trim(), status: st,
      semC: String(v[8]).trim(), qtd: String(v[9]).trim(), perc: ccFmtP_(v[10]), dil: ccFmtD_(v[29]),
      tel: [12, 13, 14, 15, 16].map(k => String(v[k]).trim()).filter(x => /\d/.test(x)),
      email: [17, 18].map(k => String(v[k]).trim()).filter(x => x.indexOf('@') > 0)
    };
    if (st === CC_ST[1] && comp[cpf]) al.banco = comp[cpf];
    alunos.push(al);
  }
  const out = { alunos: alunos, light: light, diag: { aba: bsh.getName(), linhas: base.length, cab: h0 + 1, status: st0, total: tot, indicador: light.length }, ts: Date.now() };
  ccCachePut_(chave, out, 600);
  return out;
}

// Registros do Call Center (aba CC_REGISTROS: B semestre, E data, K colaborador) para os Indicadores
function lerRegistrosCallCenter_(periodo, ctx) {
  var sh = ccSsReg_(periodo).getSheetByName(ccAbaNome_('REGISTROS'));   // planilha do período
  if (!sh) return;
  var ultima = sh.getLastRow();
  if (ultima < 2) return;
  var vals = sh.getRange(2, 1, ultima - 1, 11).getValues();          // A..K
  var alvo = String(periodo || '').trim(), tz = ccTz_(), semData = 0, nomes = ccNomesPorEmail_();
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    if (formatPeriodo_(r[1]) !== alvo) continue;                      // só o período selecionado
    var nome = String(nomes[String(r[9] || '').trim().toLowerCase()] || (r[10] == null ? '' : r[10])).trim();   // nome atual pelo e-mail
    if (!nome) continue;
    var d = dataBR_(r[4], tz);
    if (!d) { semData++; continue; }
    var c = nomeCanonico_(nome, ctx.canon);
    ctx.addColab(c);
    ctx.ensure(d, 'Call Center');
    ctx.ind[d]['Call Center'][c] = (ctx.ind[d]['Call Center'][c] || 0) + 1;
  }
  if (semData > 0) ctx.avisos.push(semData + ' registro(s) do Call Center ignorado(s) por não terem data válida na coluna E.');
}
