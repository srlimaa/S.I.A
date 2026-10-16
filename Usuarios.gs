// ══════════════════════════════════════════════════════════════════════
// USUÁRIOS — cadastro, edição, exclusão e avatares.
// ══════════════════════════════════════════════════════════════════════

// Lista de avatares (só quem tem) para o histórico
function getAvatares() {
  try {
    var data = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Usuarios').getDataRange().getValues();
    var lista = [];
    for (var i = 1; i < data.length; i++) {
      var t = String(data[i][5] || '').trim(), v = String(data[i][6] || '');
      if (t && v) lista.push({ nome: String(data[i][1]).trim(), tipo: t, valor: v });
    }
    return lista;
  } catch(e) { return []; }
}

// Cada usuário altera apenas o próprio avatar. tipo: 'emoji' | 'imagem' | 'nenhum'
function salvarMeuAvatar(userId, tipo, valor) {
  try {
    var sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Usuarios');
    var data = sh.getDataRange().getValues();
    var linha = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(userId)) { linha = i; break; }
    }
    if (linha === -1) return { success: false, message: 'Usuário não encontrado.' };

    var t = '', v = '';
    if (tipo && tipo !== 'nenhum') {
      var val = validarIcone_(tipo, valor, AV_IMG_MAX_USER);
      if (!val.ok) return { success: false, message: val.message };
      t = val.tipo; v = val.valor;
    }
    if (!String(data[0][5] || '').trim()) sh.getRange(1, 6).setValue('AvatarTipo');
    if (!String(data[0][6] || '').trim()) sh.getRange(1, 7).setValue('AvatarValor');
    var rng = sh.getRange(linha + 1, 6, 1, 2);
    rng.setNumberFormat('@');
    rng.setValues([[t, v]]);
    return { success: true, tipo: t, valor: v };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── Usuários ─────────────────────────────────────────────────────────────────

function getUsuarios() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Usuarios');
    var data = sheet.getDataRange().getValues();
    var usuarios = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== '')
        usuarios.push({ id: String(data[i][0]), nome: data[i][1],
                        email: data[i][2], perfil: data[i][4] });
    }
    return usuarios;
  } catch(e) { return []; }
}

function criarUsuario(nome, email, senha, perfil) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Usuarios');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim() === String(email).trim())
        return { success: false, message: 'Este e-mail já está cadastrado.' };
    }
    sheet.appendRow([new Date().getTime().toString(), nome, email, senha, perfil]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

function editarUsuario(id, nome, email, senha, perfil, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode editar usuários.' };
    nome  = String(nome  || '').trim();
    email = String(email || '').trim();
    senha = String(senha || '').trim();          // vazio = mantém a senha atual
    perfil = String(perfil || '').trim();
    if (perfil !== 'Admin' && perfil !== 'Atend') return { success: false, message: 'Perfil inválido.' };
    if (!nome || !email) return { success: false, message: 'Nome e e-mail são obrigatórios.' };
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { success: false, message: 'E-mail inválido.' };

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Usuarios');
    var data = sheet.getDataRange().getValues();
    var linha = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) linha = i;
      else if (String(data[i][2]).trim().toLowerCase() === email.toLowerCase())
        return { success: false, message: 'Este e-mail já está cadastrado para outro usuário.' };
    }
    if (linha === -1) return { success: false, message: 'Usuário não encontrado.' };

    // Não deixa o sistema ficar sem nenhum Admin
    var perfilAntigo = String(data[linha][4]).trim();
    if (perfilAntigo === 'Admin' && perfil !== 'Admin') {
      var admins = 0;
      for (var k = 1; k < data.length; k++) if (String(data[k][4]).trim() === 'Admin') admins++;
      if (admins <= 1) return { success: false, message: 'Este é o único Admin. Defina outro Admin antes de alterar este perfil.' };
    }

    var nomeAntigo = String(data[linha][1]).trim();
    sheet.getRange(linha + 1, 2).setValue(nome);
    sheet.getRange(linha + 1, 3).setValue(email);
    sheet.getRange(linha + 1, 5).setValue(perfil);
    if (senha) sheet.getRange(linha + 1, 4).setNumberFormat('@').setValue(senha);   // texto: preserva zeros à esquerda

    // Se o nome mudou, acompanha o nome nos registros já salvos (todos os períodos)
    var renomeados = 0;
    if (nomeAntigo && nomeAntigo !== nome) renomeados = renomearColaborador_(nomeAntigo, nome);

    return { success: true, nome: nome, email: email, perfil: perfil, renomeados: renomeados };
  } catch(e) { return { success: false, message: e.message }; }
}

// Troca o nome do colaborador nas colunas de cada aba de registro, em todas as planilhas de período
function renomearColaborador_(antigo, novo) {
  var cols = { 'Presencial': [9], 'Ramal': [8], 'Email': [11], 'E-mail': [11], 'Virtual': [3, 6] };
  var ids = {};
  var per = getPeriodosSheet_().getDataRange().getValues();
  for (var i = 1; i < per.length; i++) {
    var pid = String(per[i][1]).trim();
    if (pid) ids[pid] = true;
  }
  var total = 0;
  Object.keys(ids).forEach(function(pid) {
    try {
      var ss = SpreadsheetApp.openById(pid);
      Object.keys(cols).forEach(function(aba) {
        var sh = ss.getSheetByName(aba);
        if (!sh || sh.getLastRow() < 2) return;
        cols[aba].forEach(function(c) {
          var rng = sh.getRange(2, c, sh.getLastRow() - 1, 1);
          var vals = rng.getValues(), mudou = false;
          for (var r = 0; r < vals.length; r++) {
            if (String(vals[r][0]).trim() === antigo) { vals[r][0] = novo; mudou = true; total++; }
          }
          if (mudou) rng.setValues(vals);
        });
      });
    } catch(e) { Logger.log('Renomear falhou em ' + pid + ': ' + e.message); }
  });
  return total;
}

function excluirUsuario(id) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Usuarios');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: 'Usuário não encontrado.' };
  } catch(e) { return { success: false, message: e.message }; }
}
