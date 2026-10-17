// ══════════════════════════════════════════════════════════════════════
// CALL CENTER — planilha, login, utilitários e leitura de dados (funções "cc...").
// ══════════════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════════════════
// CALL CENTER - ADITAMENTOS (integrado a este projeto)
// A planilha do Call Center (abas USUARIOS, CONFIG, SEMESTRES, OPCOES, REGISTROS, RESPONSAVEIS) é aberta pelo ID.
// A tela do Call Center está no arquivo HTML "CallCenter". Os nomes das funções começam com "cc".
// ═══════════════════════════════════════════════════════════════════════════
var CC_PLANILHA_PADRAO = '1L5wclKLy7MZvVhDgR3KUDPla2Q7GI50FifWv0VXN480';

function ccPlanilhaId_() {
  var id = String(lerConfig_().cc_planilha || '').trim();
  return id || CC_PLANILHA_PADRAO;
}

function ccSs_() { return SpreadsheetApp.openById(ccPlanilhaId_()); }

function ccTz_() { return Session.getScriptTimeZone(); }

const CC_HEAD = {
  USUARIOS: ['Email', 'Nome', 'Senha', 'Perfil', 'Foto', 'Tema', 'Semestre'],
  CONFIG: ['Chave', 'Valor'],
  SEMESTRES: ['Semestre', 'IdBase', 'IdCallCenter', 'IdRegistros'],
  OPCOES: ['Tipo', 'Valor'],
  REGISTROS: ['Id', 'Semestre', 'CPF', 'Aluno', 'DataHora', 'Objetivo', 'Meio', 'Resultado', 'Observacao', 'ColabEmail', 'ColabNome', 'Status', 'EditadoEm'],
  RESPONSAVEIS: ['Semestre', 'CPF', 'Email']
};

const CC_ST = ['SOLICITADO / NÃO ADITOU', 'SOLICITADO / ENVIADO AO BANCO'];

const CC_ABA = 'Aditamento';      // aba da planilha "Call Center - Aditamentos"

const CC_COL_BANCO = 45;             // coluna AT (índice 0)

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

function ccTab_(n, ss) {
  const sh = (ss || ccSs_()).getSheetByName(n), v = sh.getDataRange().getValues(), h = v.shift();
  return v.filter(r => r.join('') !== '').map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = (r[j] instanceof Date) ? Utilities.formatDate(r[j], ccTz_(), k === 'Semestre' ? 'yyyy.M' : 'yyyy-MM-dd HH:mm') : r[j]); return o; });
}

function ccAdd_(n, o, ss) { const sh = (ss || ccSs_()).getSheetByName(n), r = sh.getLastRow() + 1, g = sh.getRange(r, 1, 1, CC_HEAD[n].length); g.setNumberFormat('@'); g.setValues([CC_HEAD[n].map(k => o[k] === undefined ? '' : String(o[k]))]); }

function ccUpd_(n, row, o, ss) { const sh = (ss || ccSs_()).getSheetByName(n); CC_HEAD[n].forEach((k, j) => { if (o[k] !== undefined) { const c = sh.getRange(row, j + 1); c.setNumberFormat('@'); c.setValue(String(o[k])); } }); }

function ccDel_(n, row) { ccSs_().getSheetByName(n).deleteRow(row); }

function ccId_(s) { s = String(s || '').trim(); const m = s.match(/\/d\/([a-zA-Z0-9-_]+)/); return m ? m[1] : s; }

function ccKey_(s) { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/\s*\/\s*/g, ' / ').replace(/\s+/g, ' ').trim(); }

const CC_STATUS7 = ['ADITOU', 'NÃO SOLICITADO', 'REJEITADO', 'SOLICITADO / CANCELADO', 'SOLICITADO / EM PROCESSAMENTO', 'SOLICITADO / ENVIADO AO BANCO', 'SOLICITADO / NÃO ADITOU'];

function ccCanon_(s) { const k = ccKey_(s); return CC_STATUS7.find(x => ccKey_(x) === k) || k; }

function ccNorm_(s) { return String(s || '').toUpperCase().replace(/\s+/g, ' ').replace(/\s*\/\s*/g, ' / ').trim(); }

function ccCpf_(s) { const d = String(s || '').replace(/\D/g, ''); return d ? ('00000000000' + d).slice(-11) : ''; }

function ccFmtD_(v) { if (v instanceof Date) return Utilities.formatDate(v, ccTz_(), 'dd/MM/yyyy'); return String(v || '').trim(); }

function ccFmtP_(v) { if (typeof v === 'number') return (v <= 1 ? Math.round(v * 10000) / 100 : v) + '%'; return String(v || '').trim(); }

/* ---------- senha ---------- */

function ccHash_(s, salt) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + s).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }

function ccHashNew_(s) { const salt = Utilities.getUuid().slice(0, 8); return salt + '$' + ccHash_(s, salt); }

function ccCheck_(s, stored) { const p = String(stored).split('$'); return p.length === 2 && ccHash_(s, p[0]) === p[1]; }

/* ---------- login / sessão ---------- */

function ccGetPublic() { ccEnsure_(); const c = ccCfg_(); return { nomeLogin: c.nomeLogin, logo: c.logo }; }

function ccApiLogin(email, senha) {
  // login único: o usuário é o do sistema de Atendimentos; precisa ter acesso liberado ao Call Center
  const r = login(email, senha);
  if (!r || !r.success) return { err: 'E-mail ou senha inválidos.' };
  const c = ccLoginInterno_(r.email);
  if (!c) return { err: 'Seu usuário não tem acesso ao Call Center. Peça a um Admin para liberar.' };
  return { ok: { token: c.token } };
}

function ccApiLogout(t) { CacheService.getScriptCache().remove('s_' + t); }

function ccCfg_() { const o = {}; ccTab_('CONFIG').forEach(r => o[r.Chave] = r.Valor); return o; }

function ccOpc_() { const o = { Objetivo: [], Meio: [], Resultado: [] }; ccTab_('OPCOES').forEach(r => { if (o[r.Tipo]) o[r.Tipo].push(r.Valor); }); return o; }

function ccPub_(u) { return { email: u.Email, nome: u.Nome, perfil: u.Perfil, foto: u.Foto, tema: u.Tema || 'Padrão', semestre: ccSemU_(u) }; }

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
function ccAtend_() {
  var data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Usuarios').getDataRange().getValues();
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

function ccPrefsEmails_() {
  try { var o = {}; ccTab_('USUARIOS').forEach(function(x) { o[String(x.Email).toLowerCase()] = 1; }); return o; }
  catch (e) { return {}; }
}

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
  const r = ccTab_('SEMESTRES').find(x => String(x.Semestre) === String(sem));
  const id = r ? ccId_(r.IdRegistros) : '';
  if (!id || id === ccPlanilhaId_()) return ccSs_();
  const ss = SpreadsheetApp.openById(id);
  ccPrepararRegistros_(ss);
  return ss;
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

function ccRegs_(sem) {
  const nomes = ccNomesPorEmail_();       // nome atual do colaborador (cadastro de usuários do Atendimentos)
  return ccTab_('REGISTROS', ccSsReg_(sem)).filter(r => String(r.Semestre) === String(sem)).map(r => ({
    id: r.Id, cpf: r.CPF, aluno: r.Aluno, data: String(r.DataHora), objetivo: r.Objetivo, meio: r.Meio, resultado: r.Resultado,
    obs: r.Observacao, colab: String(r.ColabEmail || '').toLowerCase(), colabNome: nomes[String(r.ColabEmail || '').toLowerCase()] || r.ColabNome, status: r.Status, editado: r.EditadoEm
  }));
}

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

function ccTemAbas_() {
  var ss = ccSs_();
  return Object.keys(CC_HEAD).every(function(n) { return !!ss.getSheetByName(n); });
}

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
    ccUpd_('USUARIOS', u._row, { Semestre: p });
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// Conteúdo da tela do Call Center (arquivo HTML "CallCenter" do projeto)
function ccGetHtml() { return HtmlService.createTemplateFromFile('CallCenter').evaluate().getContent(); }

// Planilha do Call Center (configurável em ⚙️ › Planilhas)
function getPlanilhaCC(adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin.' };
    return { success: true, link: 'https://docs.google.com/spreadsheets/d/' + ccPlanilhaId_() };
  } catch(e) { return { success: false, message: e.message }; }
}

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
