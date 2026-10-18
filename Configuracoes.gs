// ══════════════════════════════════════════════════════════════════════
// CONFIGURAÇÕES DO SISTEMA — ícone, nome, fundo do login (aba "Config").
// ══════════════════════════════════════════════════════════════════════

// ── Configurações do sistema (logo/emoji) e avatares dos usuários ────────────
// Aba "Config" (planilha mestre): A-Chave | B-Valor   (icone_tipo, icone_valor)
// Aba "Usuarios": colunas F-AvatarTipo | G-AvatarValor
// As imagens são guardadas como texto (data URL); limite do Sheets: 50.000 caracteres por célula.
var AV_EMOJI_MAX = 16;

var AV_IMG_MAX_SISTEMA = 48000;

var AV_IMG_MAX_USER = 24000;

function validarIcone_(tipo, valor, maxImg) {
  if (tipo === 'emoji') {
    valor = String(valor || '').trim();
    if (!valor || valor.length > AV_EMOJI_MAX || /[<>&"'\\]/.test(valor))
      return { ok: false, message: 'Emoji inválido.' };
    return { ok: true, tipo: 'emoji', valor: valor };
  }
  if (tipo === 'imagem') {
    valor = String(valor || '');
    if (valor.length > maxImg)
      return { ok: false, message: 'Imagem grande demais. Reduza o tamanho ou use uma imagem mais simples.' };
    if (!/^data:image\/(png|jpeg|webp);base64,[A-Za-z0-9+\/=]+$/.test(valor))
      return { ok: false, message: 'Formato de imagem inválido.' };
    return { ok: true, tipo: 'imagem', valor: valor };
  }
  return { ok: false, message: 'Tipo inválido.' };
}

function lerConfig_() {
  var cfg = {};
  try {
    var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Config');
    if (!sh) return cfg;
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) cfg[String(data[i][0]).trim()] = data[i][1];
  } catch(e) {}
  return cfg;
}

// Fundo da tela de login: JSON guardado em pedaços (login_fundo_0, _1, ...) por causa do limite de 50.000 caracteres por célula
var FUNDO_CHUNK = 40000;

var FUNDO_MAX_CHUNKS = 6;

var AV_IMG_MAX_FUNDO = 150000;

function lerFundoLogin_(cfg) {
  try {
    var n = parseInt(cfg.login_fundo_n, 10) || 0;
    if (!n) return null;
    var txt = '';
    for (var i = 0; i < n; i++) txt += String(cfg['login_fundo_' + i] || '');
    var f = JSON.parse(txt);
    return (f && f.tipo) ? f : null;
  } catch(e) { return null; }
}

var NOME_PADRAO = 'Sistema de Atendimentos';

var SUBTITULO_PADRAO = 'UNILEÃO — Integração / FIES';

function getConfigSistema() {
  var padrao = { tipo: 'emoji', valor: '🏛️', fundo: null, nome: NOME_PADRAO, subtitulo: SUBTITULO_PADRAO };
  try {
    var cfg = lerConfig_();
    var tipo = String(cfg.icone_tipo || ''), valor = String(cfg.icone_valor || '');
    var out = { tipo: 'emoji', valor: '🏛️', fundo: lerFundoLogin_(cfg), nome: NOME_PADRAO, subtitulo: SUBTITULO_PADRAO };
    if ((tipo === 'emoji' || tipo === 'imagem') && valor) { out.tipo = tipo; out.valor = valor; }
    if (String(cfg.nome_sistema || '').trim()) out.nome = String(cfg.nome_sistema).trim();
    if (cfg.subtitulo_login !== undefined) out.subtitulo = String(cfg.subtitulo_login).trim();   // vazio = ocultar
    return out;
  } catch(e) { return padrao; }
}

// Nome do sistema (topo do sistema, escolha do sistema e aba do navegador) — só Admin.
// subtitulo = null/undefined: não altera o subtítulo da tela de login (ele é editado em salvarTelaLogin).
function salvarNomeSistema(nome, subtitulo, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode alterar o nome do sistema.' };
    nome = String(nome == null ? '' : nome).replace(/\s+/g, ' ').trim();
    var mudaSub = (subtitulo !== null && subtitulo !== undefined);
    subtitulo = mudaSub ? String(subtitulo).replace(/\s+/g, ' ').trim() : '';
    if (!nome) return { success: false, message: 'Informe o nome do sistema.' };
    if (nome.length > 60) return { success: false, message: 'O nome pode ter no máximo 60 caracteres.' };
    if (subtitulo.length > 100) return { success: false, message: 'O subtítulo pode ter no máximo 100 caracteres.' };
    if (/[<>]/.test(nome + subtitulo)) return { success: false, message: 'O texto não pode conter < ou >.' };
    setConfig_('nome_sistema', nome);
    if (mudaSub) setConfig_('subtitulo_login', subtitulo);
    return { success: true, nome: nome };
  } catch(e) { return { success: false, message: e.message }; }
}

function validarFundo_(f) {
  var cor = /^#[0-9a-fA-F]{6}$/;
  if (!f || typeof f !== 'object') return { ok: false, message: 'Configuração inválida.' };
  var tipo = String(f.tipo || '');
  if (tipo === 'padrao') return { ok: true, valor: { tipo: 'padrao' } };
  if (tipo === 'cor') {
    if (!cor.test(String(f.c1))) return { ok: false, message: 'Cor inválida.' };
    return { ok: true, valor: { tipo: 'cor', c1: f.c1 } };
  }
  if (tipo === 'degrade') {
    if (!cor.test(String(f.c1)) || !cor.test(String(f.c2))) return { ok: false, message: 'Cor inválida.' };
    var ang = parseInt(f.ang, 10);
    if (isNaN(ang) || ang < 0 || ang > 360) ang = 135;
    return { ok: true, valor: { tipo: 'degrade', c1: f.c1, c2: f.c2, ang: ang } };
  }
  if (tipo === 'imagem') {
    var v = validarIcone_('imagem', f.img, AV_IMG_MAX_FUNDO);
    if (!v.ok) return v;
    return { ok: true, valor: { tipo: 'imagem', img: v.valor, c1: cor.test(String(f.c1)) ? f.c1 : '#1e3a5f' } };
  }
  return { ok: false, message: 'Tipo de fundo inválido.' };
}

function salvarFundoLogin(fundo, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode alterar o fundo do login.' };
    var f = validarFundo_(fundo);
    if (!f.ok) return { success: false, message: f.message };
    var txt = JSON.stringify(f.valor);
    var n = Math.ceil(txt.length / FUNDO_CHUNK);
    if (n > FUNDO_MAX_CHUNKS) return { success: false, message: 'Imagem grande demais.' };
    var antigo = parseInt(lerConfig_().login_fundo_n, 10) || 0;
    for (var i = 0; i < n; i++) setConfig_('login_fundo_' + i, txt.substr(i * FUNDO_CHUNK, FUNDO_CHUNK));
    for (var j = n; j < antigo; j++) setConfig_('login_fundo_' + j, '');
    setConfig_('login_fundo_n', String(n));
    return { success: true, fundo: f.valor };
  } catch(e) { return { success: false, message: e.message }; }
}

function setConfig_(chave, valor) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName('Config');
  if (!sh) {
    sh = ss.insertSheet('Config');
    sh.getRange('A:B').setNumberFormat('@');
    sh.appendRow(['Chave', 'Valor']);
  }
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === chave) {
      sh.getRange(i + 1, 2).setNumberFormat('@').setValue(valor);
      return;
    }
  }
  var rng = sh.getRange(sh.getLastRow() + 1, 1, 1, 2);
  rng.setNumberFormat('@');
  rng.setValues([[chave, valor]]);
}

function salvarConfigSistema(tipo, valor, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode alterar as configurações do sistema.' };
    var v = validarIcone_(tipo, valor, AV_IMG_MAX_SISTEMA);
    if (!v.ok) return { success: false, message: v.message };
    setConfig_('icone_tipo', v.tipo);
    setConfig_('icone_valor', v.valor);
    return { success: true, tipo: v.tipo, valor: v.valor };
  } catch(e) { return { success: false, message: e.message }; }
}
