// ══════════════════════════════════════════════════════════════════════
// AUTENTICAÇÃO — login (registro), login unificado e checagem de perfil Admin.
// ══════════════════════════════════════════════════════════════════════

// ── Login ────────────────────────────────────────────────────────────────────
function login(email, senha) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Usuarios');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim() === String(email).trim() &&
          String(data[i][3]).trim() === String(senha).trim()) {
        return { success: true, nome: data[i][1],
                 perfil: data[i][4], id: String(data[i][0]),
                 avatar: { tipo: String(data[i][5] || '').trim(), valor: String(data[i][6] || '') } };
      }
    }
    return { success: false, message: 'E-mail ou senha incorretos.' };
  } catch(e) { return { success: false, message: 'Erro: ' + e.message }; }
}

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

// Login único: confere o e-mail/senha nas duas bases. Cada sistema só é liberado onde a senha confere.
function loginUnificado(email, senha) {
  email = String(email == null ? '' : email).trim();
  senha = String(senha == null ? '' : senha).trim();
  var reg = null, cc = null, msg = '', ccErro = '';
  try {
    var r = login(email, senha);
    if (r && r.success) reg = r; else if (r && r.message) msg = r.message;
  } catch(e) { msg = 'Falha ao consultar o sistema de registro: ' + e.message; }
  try { cc = ccLoginInterno_(email, senha); }
  catch(e2) { ccErro = e2.message; }
  if (!reg && !cc) return { success: false, message: msg || 'E-mail ou senha incorretos.' };
  return { success: true, registro: reg, cc: cc, ccErro: ccErro };
}
