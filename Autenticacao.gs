// ══════════════════════════════════════════════════════════════════════
// AUTENTICAÇÃO — login (registro), login unificado e checagem de perfil Admin.
// ══════════════════════════════════════════════════════════════════════

// ARQUIVO: Autenticacao.gs
// O QUE FAZ: login (e-mail/senha), checagem de Admin, login único Atendimentos + Call Center e sessão que sobrevive ao recarregar a página.
// ONDE É USADO: chamado pelas telas via google.script.run (js_login.html, js_boot.html) e por outros .gs (ehAdmin_).

// ── Login ────────────────────────────────────────────────────────────────────
// Confere e-mail e senha na aba Usuarios (planilha mestre). Devolve nome, perfil, ID, e-mail e avatar, ou a mensagem de erro.
// É usada por loginUnificado() logo abaixo.
function login(email, senha) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Usuarios');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim() === String(email).trim() &&
          String(data[i][3]).trim() === String(senha).trim()) {
        return { success: true, nome: data[i][1],
                 perfil: (String(data[i][4]).trim() === 'Atend' ? 'Colab' : data[i][4]), id: String(data[i][0]), email: String(data[i][2]).trim(),
                 avatar: { tipo: String(data[i][5] || '').trim(), valor: String(data[i][6] || '') } };
      }
    }
    return { success: false, message: 'E-mail ou senha incorretos.' };
  } catch(e) { return { success: false, message: 'Erro: ' + e.message }; }
}

// Confere se o usuário (pelo ID) tem perfil Admin. É usada por quase todas as funções que só o Admin pode executar.
function ehAdmin_(userId) {
  var data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Usuarios').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId)) return String(data[i][4]).trim() === 'Admin';
  }
  return false;
}

// ── Editar / excluir registros do histórico ─────────────────────────────────
// Admin altera qualquer registro; os demais perfis, apenas os próprios.
function usuarioPorId_(userId) {
  var data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Usuarios').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(userId))
      return { nome: String(data[i][1]).trim(), perfil: String(data[i][4]).trim() };
  }
  return null;
}

// Login único: valida e-mail/senha no cadastro de usuários do sistema de Atendimentos (aba "Usuarios").
// O acesso ao Call Center é liberado por usuário (coluna "CallCenter" = Sim/Não) e usa o mesmo nome, perfil e foto.
function loginUnificado(email, senha) {
  email = String(email == null ? '' : email).trim();
  senha = String(senha == null ? '' : senha).trim();
  var reg = null, cc = null, msg = '', ccErro = '';
  try {
    var r = login(email, senha);
    if (r && r.success) reg = r; else if (r && r.message) msg = r.message;
  } catch(e) { msg = 'Falha ao consultar o cadastro de usuários: ' + e.message; }
  if (!reg) return { success: false, message: msg || 'E-mail ou senha incorretos.' };
  try { cc = ccLoginInterno_(reg.email); }          // null = sem acesso ao Call Center
  catch(e2) { ccErro = e2.message; }
  return { success: true, registro: reg, cc: cc, ccErro: ccErro, sessao: sessaoCriar_(reg.email) };
}

// ── Sessão do sistema (guardada no navegador) ───────────────────────────────
// Depois do login, o navegador recebe um código de sessão (válido por 6 h, renovado a cada retomada).
// Ao recarregar a página, ele é usado em retomarSessao() para voltar logado, sem pedir a senha de novo.
function sessaoCriar_(email) {
  var t = Utilities.getUuid();
  CacheService.getScriptCache().put('ss_' + t, String(email), 21600);
  return t;
}

// Monta os dados do usuário (mesmo formato de login()) a partir do e-mail, SEM conferir senha.
// Uso exclusivo de retomarSessao().
function regPorEmail_(email) {
  var em = String(email || '').trim().toLowerCase();
  var data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Usuarios').getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim().toLowerCase() === em) {
      return { success: true, nome: data[i][1], perfil: (String(data[i][4]).trim() === 'Atend' ? 'Colab' : data[i][4]),
               id: String(data[i][0]), email: String(data[i][2]).trim(),
               avatar: { tipo: String(data[i][5] || '').trim(), valor: String(data[i][6] || '') } };
    }
  }
  return null;
}

// Chamada quando a página é recarregada (js_boot.html): troca o código de sessão guardado no navegador pelos dados do usuário,
// sem pedir senha. Renova a sessão por 6 h e já devolve um novo token do Call Center.
function retomarSessao(tk) {
  try {
    tk = String(tk || '').trim();
    if (!tk) return { success: false };
    var cache = CacheService.getScriptCache(), email = cache.get('ss_' + tk);
    if (!email) return { success: false };
    var reg = regPorEmail_(email);
    if (!reg) return { success: false };
    cache.put('ss_' + tk, email, 21600);                      // renova
    var cc = null, ccErro = '';
    try { cc = ccLoginInterno_(reg.email); } catch(e) { ccErro = e.message; }
    return { success: true, registro: reg, cc: cc, ccErro: ccErro, sessao: tk };
  } catch(e) { return { success: false }; }
}

// Apaga o código de sessão do navegador (botão Sair, em js_login.html > logout()).
function encerrarSessao(tk) {
  try { CacheService.getScriptCache().remove('ss_' + String(tk || '').trim()); } catch(e) {}
}
