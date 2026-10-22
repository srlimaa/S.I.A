// ══════════════════════════════════════════════════════════════════════
// CALL CENTER — planilha, login, utilitários e leitura de dados (funções "cc...").
// ══════════════════════════════════════════════════════════════════════

// ARQUIVO: CallCenterBase.gs
// O QUE FAZ: base do sistema de Call Center: abre as planilhas, lê/grava abas, valida usuários, abre a sessão e lê a base de alunos.
// ONDE É USADO: CallCenterApi.gs (ações da tela), Autenticacao.gs (login único), Indicadores.gs (contagem de registros) e js_callcenter.html (abertura).
// DICA: todas as funções cc... pertencem ao Call Center; as que terminam com _ são internas (a tela não as chama direto).

// ═══════════════════════════════════════════════════════════════════════════
// CALL CENTER - ADITAMENTOS (integrado a este projeto)
// A planilha do Call Center (abas USUARIOS, CONFIG, SEMESTRES, OPCOES, REGISTROS, RESPONSAVEIS) é aberta pelo ID.
// A tela do Call Center está no arquivo HTML "CallCenter". Os nomes das funções começam com "cc".
// ═══════════════════════════════════════════════════════════════════════════
// ID da planilha do Call Center usada enquanto nenhuma outra for configurada (Configurações › Planilhas).
var CC_PLANILHA_PADRAO = '1L5wclKLy7MZvVhDgR3KUDPla2Q7GI50FifWv0VXN480';

/* ---------- velocidade: memória da execução + cache compartilhado ----------
   Cada chamada do navegador é uma execução nova do script. Dentro dela, planilhas e abas já lidas ficam em memória
   (_CC_MEM) em vez de serem abertas/lidas de novo. Entre execuções, o CacheService guarda a base de alunos (compactada). */
var _CC_MEM = {};

// Guarda em memória, durante a execução, o resultado de f() com a chave k (evita reabrir/reler a mesma planilha várias vezes).
function ccMem_(k, f) { if (!(k in _CC_MEM)) _CC_MEM[k] = f(); return _CC_MEM[k]; }

// Apaga da memória as cópias cujo nome começa por "prefixo" (usada depois de gravar numa aba).
function ccMemLimpa_(prefixo) { Object.keys(_CC_MEM).forEach(function(k) { if (k.indexOf(prefixo) === 0) delete _CC_MEM[k]; }); }

// Abre uma planilha pelo ID (uma única vez por execução).
function ccOpen_(id) { return ccMem_('ss|' + id, function() { return SpreadsheetApp.openById(id); }); }

// Abre a planilha MESTRE do Atendimentos (usuários, períodos, config) uma única vez por execução.
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

// Lê do cache compartilhado o objeto guardado por ccCachePut_ (junta as partes, descompacta e converte). null se não houver.
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
function ccVerChave_(sem) { return 'cc_ver_' + String(sem).replace(/[^\w.-]/g, ''); }

// Versão atual dos dados vivos do semestre. Se não existir, cria uma nova.
function ccVer_(sem) {
  var c = CacheService.getScriptCache(), k = ccVerChave_(sem), v = c.get(k);
  if (!v) { v = Utilities.getUuid().slice(0, 8); c.put(k, v, 21600); }
  return v;
}

// Troca a versão do semestre. É chamada a cada gravação (registro, exclusão, responsável) para avisar as outras telas abertas.
function ccBump_(sem) { CacheService.getScriptCache().put(ccVerChave_(sem), Utilities.getUuid().slice(0, 8), 21600); }

// ID da planilha do Call Center em uso (a configurada em Planilhas ou a padrão). Fica em cache para não reler a aba Config.
function ccPlanilhaId_() {
  return ccMem_('plan', function() {
    var c = CacheService.getScriptCache(), v = c.get('cc_planilha');
    if (v === null) { v = String(lerConfig_().cc_planilha || '').trim(); c.put('cc_planilha', v || '-', 21600); }
    v = (v === '-') ? '' : v;
    return v || CC_PLANILHA_PADRAO;
  });
}

// Planilha principal do Call Center, já aberta.
function ccSs_() { return ccOpen_(ccPlanilhaId_()); }

// Fuso horário do projeto (usado nas datas gravadas).
function ccTz_() { return Session.getScriptTimeZone(); }

// Colunas de cada aba da planilha do Call Center (a ordem aqui é a ordem das colunas na planilha).
const CC_HEAD = {
  USUARIOS: ['Email', 'Nome', 'Senha', 'Perfil', 'Foto', 'Tema', 'Semestre'],
  CONFIG: ['Chave', 'Valor'],
  SEMESTRES: ['Semestre', 'IdBase', 'IdCallCenter', 'IdRegistros'],
  OPCOES: ['Tipo', 'Valor'],
  REGISTROS: ['Id', 'Semestre', 'CPF', 'Aluno', 'DataHora', 'Objetivo', 'Meio', 'Resultado', 'Observacao', 'ColabEmail', 'ColabNome', 'Status', 'EditadoEm'],
  RESPONSAVEIS: ['Semestre', 'CPF', 'Email']
};

// Os dois status que aparecem na lista de alunos do Call Center.
const CC_ST = ['SOLICITADO / NÃO ADITOU', 'SOLICITADO / ENVIADO AO BANCO'];

// Nome da aba da planilha "Call Center - Aditamentos" onde está o comparecimento ao banco.
const CC_ABA = 'Aditamento';      // aba da planilha "Call Center - Aditamentos"

// Coluna (AT, contando de 0) com o texto do comparecimento ao banco, no formato "dd/mm/aaaa até dd/mm/aaaa".
const CC_COL_BANCO = 45;             // coluna AT (índice 0)

// Cria as abas que faltarem (USUARIOS, CONFIG, SEMESTRES, OPCOES, REGISTROS, RESPONSAVEIS) com cabeçalho e opções iniciais.
function ccEnsure_() {
  const ss = ccSs_();
  Object.keys(CC_HEAD).forEach(n => {
    let sh = ss.getSheetByName(n);
    if (sh) return;
    sh = ss.insertSheet(n);
    sh.getRange(1, 1, 1, CC_HEAD[n].length).setValues([CC_HEAD[n]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, sh.getMaxRows() - 1, CC_HEAD[n].length).setNumberFormat('@');
    if (n === 'CONFIG') { sh.appendRow(['nomeLogin', 'Call Center']); sh.appendRow(['nomeSistema', 'Call Center - Aditamentos']); sh.appendRow(['logo', '']); }
    if (n === 'SEMESTRES') ccAdd_('SEMESTRES', { Semestre: '2026.2' });
    if (n === 'OPCOES') {
      ['Informar sobre o aditamento', 'Orientar comparecimento ao banco', 'Cobrar documentação', 'Verificar andamento'].forEach(v => sh.appendRow(['Objetivo', v]));
      ['Ligação', 'WhatsApp', 'E-mail'].forEach(v => sh.appendRow(['Meio', v]));
      ['Contato realizado', 'Não atendeu', 'Número inválido', 'Retornar depois', 'Aluno vai comparecer', 'Aluno já compareceu'].forEach(v => sh.appendRow(['Resultado', v]));
    }
  });
  [['SEMESTRES', 1], ['USUARIOS', 7], ['REGISTROS', 2], ['RESPONSAVEIS', 1]].forEach(x => { const sh = ss.getSheetByName(x[0]); sh.getRange(1, x[1], sh.getMaxRows(), 1).setNumberFormat('@'); });
  const d = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1') || ss.getSheetByName('Planilha1');
  if (d && ss.getSheets().length > 1 && d.getLastRow() === 0) ss.deleteSheet(d);
}

/* ---------- utilidades de planilha ---------- */

// Lê uma aba inteira e devolve uma lista de objetos {coluna: valor, _row: nº da linha}.
// O resultado fica em memória durante a execução; as gravações (ccAdd_/ccUpd_/ccDel_) descartam essa cópia.
function ccTab_(n, ss) {
  const s0 = ss || ccSs_();
  return ccMem_('tab|' + s0.getId() + '|' + n, function() {
    const sh = s0.getSheetByName(n), v = sh.getDataRange().getValues(), h = v.shift();
    return v.filter(r => r.join('') !== '').map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = (r[j] instanceof Date) ? Utilities.formatDate(r[j], ccTz_(), k === 'Semestre' ? 'yyyy.M' : 'yyyy-MM-dd HH:mm') : r[j]); return o; });
  });
}

// Acrescenta uma linha no fim da aba n (todas as células como texto).
function ccAdd_(n, o, ss) { const s0 = ss || ccSs_(), sh = s0.getSheetByName(n), r = sh.getLastRow() + 1, g = sh.getRange(r, 1, 1, CC_HEAD[n].length); g.setNumberFormat('@'); g.setValues([CC_HEAD[n].map(k => o[k] === undefined ? '' : String(o[k]))]); ccInval_(s0, n); }

// Altera, na linha informada, só as colunas presentes em o.
function ccUpd_(n, row, o, ss) { const s0 = ss || ccSs_(), sh = s0.getSheetByName(n); CC_HEAD[n].forEach((k, j) => { if (o[k] !== undefined) { const c = sh.getRange(row, j + 1); c.setNumberFormat('@'); c.setValue(String(o[k])); } }); ccInval_(s0, n); }

// Exclui uma linha da aba n.
function ccDel_(n, row) { const s0 = ccSs_(); s0.getSheetByName(n).deleteRow(row); ccInval_(s0, n); }

// Extrai o ID de uma planilha a partir do link completo (ou devolve o próprio texto, se já for o ID).
function ccId_(s) { s = String(s || '').trim(); const m = s.match(/\/d\/([a-zA-Z0-9-_]+)/); return m ? m[1] : s; }

// Texto "limpo" para comparar: sem acento, maiúsculo e com espaços padronizados.
function ccKey_(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/\s*\/\s*/g, ' / ').replace(/\s+/g, ' ').trim(); }

// Os 7 status possíveis da coluna de aditamento (usados nos gráficos do Indicador - Geral).
const CC_STATUS7 = ['ADITOU', 'NÃO SOLICITADO', 'REJEITADO', 'SOLICITADO / CANCELADO', 'SOLICITADO / EM PROCESSAMENTO', 'SOLICITADO / ENVIADO AO BANCO', 'SOLICITADO / NÃO ADITOU'];

// Converte o status lido da planilha para a grafia oficial de CC_STATUS7.
function ccCanon_(s) { const k = ccKey_(s); return CC_STATUS7.find(x => ccKey_(x) === k) || k; }

// Maiúsculo com espaços padronizados (comparações simples de texto).
function ccNorm_(s) { return String(s || '').toUpperCase().replace(/\s+/g, ' ').replace(/\s*\/\s*/g, ' / ').trim(); }

// CPF só com dígitos e sempre com 11 posições (completa com zeros à esquerda). É a chave que liga aluno, registro e responsável.
function ccCpf_(s) { const d = String(s || '').replace(/\D/g, ''); return d ? ('00000000000' + d).slice(-11) : ''; }

// Data em texto dd/mm/aaaa.
function ccFmtD_(v) { if (v instanceof Date) return Utilities.formatDate(v, ccTz_(), 'dd/MM/yyyy'); return String(v || '').trim(); }

// Percentual em texto (ex.: 0,85 vira 85%).
function ccFmtP_(v) { if (typeof v === 'number') return (v <= 1 ? Math.round(v * 10000) / 100 : v) + '%'; return String(v || '').trim(); }

/* ---------- senha ---------- */

// Calcula o hash SHA-256 de uma senha com sal (uso do login antigo do Call Center).
function ccHash_(s, salt) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + s).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }

// Cria um hash novo (sal + hash) para guardar uma senha.
function ccHashNew_(s) { const salt = Utilities.getUuid().slice(0, 8); return salt + '$' + ccHash_(s, salt); }

// Confere se a senha corresponde ao hash guardado.
function ccCheck_(s, stored) { const p = String(stored).split('$'); return p.length === 2 && ccHash_(s, p[0]) === p[1]; }

/* ---------- login / sessão ---------- */

// Dados públicos (nome e logo) do Call Center, mostrados antes do login.
function ccGetPublic() { ccEnsure_(); const c = ccCfg_(); return { nomeLogin: c.nomeLogin, logo: c.logo }; }

// Login direto no Call Center com e-mail/senha do Atendimentos. Devolve o token da sessão (hoje o login usa loginUnificado).
function ccApiLogin(email, senha) {
  // login único: o usuário é o do sistema de Atendimentos; precisa ter acesso liberado ao Call Center
  const r = login(email, senha);
  if (!r || !r.success) return { err: 'E-mail ou senha inválidos.' };
  const c = ccLoginInterno_(r.email);
  if (!c) return { err: 'Seu usuário não tem acesso ao Call Center. Peça a um Admin para liberar.' };
  return { ok: { token: c.token } };
}

// Encerra a sessão do Call Center (apaga o token do cache).
function ccApiLogout(t) { CacheService.getScriptCache().remove('s_' + t); }

// Lê a aba CONFIG (nome do sistema, logo) como um objeto.
function ccCfg_() { const o = {}; ccTab_('CONFIG').forEach(r => o[r.Chave] = r.Valor); return o; }

// Lê a aba OPCOES: listas de Objetivo, Meio e Resultado dos contatos.
function ccOpc_() { const o = { Objetivo: [], Meio: [], Resultado: [] }; ccTab_('OPCOES').forEach(r => { if (o[r.Tipo]) o[r.Tipo].push(r.Valor); }); return o; }

// Dados do usuário que vão para a tela (e-mail, nome, perfil, foto, tema, semestre).
function ccPub_(u) { return { email: u.Email, nome: u.Nome, perfil: u.Perfil, foto: u.Foto, tema: u.Tema || 'Padrão', semestre: ccSemU_(u) }; }

// Semestre do usuário (o último escolhido ou, se não houver, o mais recente cadastrado).
function ccSemU_(u) { const s = ccTab_('SEMESTRES').map(x => String(x.Semestre)); return s.indexOf(String(u.Semestre)) >= 0 ? String(u.Semestre) : s[s.length - 1]; }

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
function ccAtend_() { return ccMem_('atend', ccAtendLer_); }

// Lê a aba Usuarios da planilha mestre: id, nome, e-mail, perfil, avatar e flag de acesso ao Call Center.
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
function ccFoto_(av) {
  if (av && av.tipo === 'emoji' && av.valor) return 'emoji:' + av.valor;
  if (av && av.tipo === 'imagem' && av.valor) return av.valor;
  return '';
}

// E-mails que já têm preferências na aba USUARIOS do Call Center (quem já usava o sistema).
function ccPrefsEmails_() {
  try { var o = {}; ccTab_('USUARIOS').forEach(function(x) { o[String(x.Email).toLowerCase()] = 1; }); return o; }
  catch (e) { return {}; }
}

// Mapa e-mail → nome atual do colaborador (para mostrar sempre o nome atualizado).
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
   Cada semestre pode ter a sua própria planilha para as abas REGISTROS e RESPONSAVEIS
   (coluna IdRegistros da aba SEMESTRES). Vazio = continua usando a planilha principal do Call Center. */
function ccSsReg_(sem) {
  return ccMem_('ssreg|' + sem, function() {
    const r = ccTab_('SEMESTRES').find(x => String(x.Semestre) === String(sem));
    const id = r ? ccId_(r.IdRegistros) : '';
    if (!id || id === ccPlanilhaId_()) return ccSs_();
    const ss = ccOpen_(id);
    ccPrepararRegistros_(ss);
    return ss;
  });
}

// Cria (se ainda não existirem) as abas REGISTROS e RESPONSAVEIS com os cabeçalhos na planilha informada
function ccPrepararRegistros_(ss) {
  ['REGISTROS', 'RESPONSAVEIS'].forEach(n => {
    if (ss.getSheetByName(n)) return;
    const sh = ss.insertSheet(n);
    sh.getRange(1, 1, 1, CC_HEAD[n].length).setValues([CC_HEAD[n]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), CC_HEAD[n].length).setNumberFormat('@');
  });
  const d = ss.getSheetByName('Sheet1') || ss.getSheetByName('Página1') || ss.getSheetByName('Planilha1');
  if (d && ss.getSheets().length > 1 && d.getLastRow() === 0) ss.deleteSheet(d);
}

// Garante a coluna D (IdRegistros) no cabeçalho da aba SEMESTRES já existente
function ccGarantirColunaSem_() {
  const sh = ccSs_().getSheetByName('SEMESTRES');
  if (!sh) return;
  if (sh.getMaxColumns() < 4) sh.insertColumnsAfter(sh.getMaxColumns(), 4 - sh.getMaxColumns());
  if (String(sh.getRange(1, 4).getValue()).trim() !== 'IdRegistros') {
    sh.getRange(1, 4, sh.getMaxRows(), 1).setNumberFormat('@');
    sh.getRange(1, 4).setValue('IdRegistros').setFontWeight('bold');
  }
}

// Registros de contato do semestre, já no formato da tela (com o nome atual do colaborador).
function ccRegs_(sem) {
  const nomes = ccNomesPorEmail_();       // nome atual do colaborador (cadastro de usuários do Atendimentos)
  return ccTab_('REGISTROS', ccSsReg_(sem)).filter(r => String(r.Semestre) === String(sem)).map(r => ({
    id: r.Id, cpf: r.CPF, aluno: r.Aluno, data: String(r.DataHora), objetivo: r.Objetivo, meio: r.Meio, resultado: r.Resultado,
    obs: r.Observacao, colab: String(r.ColabEmail || '').toLowerCase(), colabNome: nomes[String(r.ColabEmail || '').toLowerCase()] || r.ColabNome, status: r.Status, editado: r.EditadoEm
  }));
}

// Lê o "comparecimento ao banco" da planilha Call Center - Aditamentos: devolve {cpf: {ini, fim}}.
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

// Confere se todas as abas do Call Center existem.
function ccTemAbas_() {
  var ss = ccSs_();
  return Object.keys(CC_HEAD).every(function(n) { return !!ss.getSheetByName(n); });
}

// Abre a sessão do Call Center para um usuário do Atendimentos: confere acesso, cria o token (válido por 6 h) e devolve nome, perfil, logo e semestres.
// Chamada por loginUnificado e retomarSessao (Autenticacao.gs).
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
    sems: ccTab_('SEMESTRES').map(function(s) { return String(s.Semestre); }),
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
    var sems = ccTab_('SEMESTRES').map(function(s) { return String(s.Semestre); });
    if (sems.indexOf(p) < 0)
      return { success: false, message: 'O período ' + p + ' não está cadastrado no Call Center. Peça a um Admin do Call Center para cadastrá-lo em Configurações › Semestres.' };
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

// Planilha do Call Center (configurável em ⚙️ › Planilhas)
function getPlanilhaCC(adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin.' };
    return { success: true, link: 'https://docs.google.com/spreadsheets/d/' + ccPlanilhaId_() };
  } catch(e) { return { success: false, message: e.message }; }
}

// Só Admin. Valida e grava a planilha do Call Center (e limpa os caches).
function salvarPlanilhaCC(link, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode alterar a planilha do Call Center.' };
    var id = ccId_(link);
    if (id) {
      var ss;
      try { ss = SpreadsheetApp.openById(id); }
      catch(e) { return { success: false, message: 'Não foi possível abrir a planilha. Confira o link e se a conta que executa o script tem acesso.' }; }
      if (!ss.getSheetByName('USUARIOS') || !ss.getSheetByName('REGISTROS'))
        return { success: false, message: 'Esta planilha não parece ser a do Call Center (faltam as abas USUARIOS e REGISTROS).' };
    }
    setConfig_('cc_planilha', id);          // vazio = volta para a planilha padrão
    CacheService.getScriptCache().remove('cc_planilha'); ccMemLimpa_('plan'); ccMemLimpa_('ss|'); ccMemLimpa_('tab|'); ccMemLimpa_('ssreg|');
    return { success: true, link: 'https://docs.google.com/spreadsheets/d/' + ccPlanilhaId_() };
  } catch(e) { return { success: false, message: e.message }; }
}

// Registros do Call Center (aba REGISTROS: B semestre, E data, K colaborador) para os Indicadores
function lerRegistrosCallCenter_(periodo, ctx) {
  var sh = ccSsReg_(periodo).getSheetByName('REGISTROS');   // planilha de registros do semestre
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
