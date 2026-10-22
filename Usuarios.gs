// USUÁRIOS: cadastro, edição, exclusão, avatares e liberação de acesso ao Call Center (coluna H da aba Usuarios).
// Perfis válidos: Admin e Colab (o antigo "Atend" é convertido automaticamente).
// Atenção: criarUsuario, excluirUsuario e salvarMeuAvatar não conferem o perfil no servidor; a restrição a Admin hoje vem da tela.
// ══════════════════════════════════════════════════════════════════════
// USUÁRIOS — cadastro, edição, exclusão e avatares.
// ══════════════════════════════════════════════════════════════════════

// Lista de avatares (só quem tem) para o histórico
// Lista nome/tipo/valor dos usuários que têm avatar, para mostrar nas telas de Histórico e Ranque.
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

// Lista os usuários para a tela de gerenciamento, já com o indicador de acesso ao Call Center.
function getUsuarios() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Usuarios');
    var data = sheet.getDataRange().getValues();
    migrarPerfilColab_(sheet, data);
    var usuarios = [];
    var tem = ccPrefsEmails_();          // quem já usava o Call Center (coluna em branco = mantém o acesso)
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] !== '')
        usuarios.push({ id: String(data[i][0]), nome: data[i][1],
                        email: data[i][2], perfil: data[i][4],
                        cc: ccFlagAcesso_(data[i][7], tem[String(data[i][2]).trim().toLowerCase()]) });
    }
    return usuarios;
  } catch(e) { return []; }
}

// Coluna H da aba "Usuarios": CallCenter (Sim/Não) — libera o acesso ao sistema do Call Center
function ccFlagTexto_(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  return (v === true || s === 'sim' || s === 's' || s === 'true' || s === '1') ? 'Sim' : 'Não';
}

// Garante a coluna H "CallCenter" na aba Usuarios.
function garantirColunaCC_(sheet) {
  if (sheet.getMaxColumns() < 8) sheet.insertColumnsAfter(sheet.getMaxColumns(), 8 - sheet.getMaxColumns());
  if (!String(sheet.getRange(1, 8).getValue()).trim()) sheet.getRange(1, 8).setValue('CallCenter');
}

// Usuários que existem só na planilha do Call Center (cadastro antigo) e ainda não foram criados aqui — só Admin
function getUsuariosSoCallCenter(adminId) {
  try {
    if (!ehAdmin_(adminId)) return [];
    var tem = {};
    ccAtend_().forEach(function(a) { tem[a.email.toLowerCase()] = 1; });
    return ccTab_('USUARIOS')
      .filter(function(x) { return String(x.Email || '').trim() && !tem[String(x.Email).trim().toLowerCase()]; })
      .map(function(x) { return { nome: String(x.Nome || ''), email: String(x.Email) }; });
  } catch(e) { return []; }
}

// Perfis: "Admin" e "Colab" (o antigo "Atend" passou a se chamar "Colab"). Vazio = valor inválido.
function perfilNorm_(v) {
  v = String(v == null ? '' : v).trim();
  if (v === 'Admin') return 'Admin';
  if (v === 'Colab' || v === 'Atend') return 'Colab';
  return '';
}

// Renomeia, na aba Usuarios, o perfil antigo "Atend" para "Colab" (acontece sozinho, uma vez)
function migrarPerfilColab_(sheet, data) {
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim() === 'Atend') {
      sheet.getRange(i + 1, 5).setValue('Colab');
      data[i][4] = 'Colab';
    }
  }
}

// Cadastra um usuário novo (e-mail único), com perfil e acesso ao Call Center (Sim/Não).
function criarUsuario(nome, email, senha, perfil, cc) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Usuarios');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim() === String(email).trim())
        return { success: false, message: 'Este e-mail já está cadastrado.' };
    }
    garantirColunaCC_(sheet);
    sheet.appendRow([new Date().getTime().toString(), nome, email, senha, (perfilNorm_(perfil) || 'Colab'), '', '', ccFlagTexto_(cc)]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// Admin: altera nome, e-mail, perfil, senha (vazia = mantém) e acesso ao Call Center.
// Impede ficar sem nenhum Admin e, se o nome mudou, renomeia o colaborador nos registros dos períodos.
function editarUsuario(id, nome, email, senha, perfil, adminId, cc) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode editar usuários.' };
    nome  = String(nome  || '').trim();
    email = String(email || '').trim();
    senha = String(senha || '').trim();          // vazio = mantém a senha atual
    perfil = perfilNorm_(perfil);
    if (!perfil) return { success: false, message: 'Perfil inválido.' };
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
    if (cc !== undefined && cc !== null && cc !== '') {                              // libera/retira o acesso ao Call Center
      garantirColunaCC_(sheet);
      sheet.getRange(linha + 1, 8).setValue(ccFlagTexto_(cc));
    }

    // Se o nome mudou, acompanha o nome nos registros já salvos (todos os períodos)
    var renomeados = 0;
    if (nomeAntigo && nomeAntigo !== nome) renomeados = renomearColaborador_(nomeAntigo, nome);

    return { success: true, nome: nome, email: email, perfil: perfil, renomeados: renomeados, cc: ccFlagTexto_(cc) === 'Sim' };
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

// Remove o usuário da aba Usuarios (sem checagem de perfil no servidor).
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
