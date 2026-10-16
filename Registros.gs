// ══════════════════════════════════════════════════════════════════════
// REGISTROS — gravar, editar, excluir e histórico (Presencial, Ramal, E-mail, Virtual).
// ══════════════════════════════════════════════════════════════════════

// ── Registros ─────────────────────────────────────────────────────────────────
// PRESENCIAL — TURNO REMOVIDO
// Colunas: A-ID | B-Data | C-CPF | D-NomeCompleto | E-Curso | F-Assunto |
//          G-Semestre | H-IES | I-Colaborador | J-DataRegistro
function registrarPresencial(dados) {
  try {
    var ss = abrirPlanilhaPeriodo_(dados.periodo);
    var sheet = ss.getSheetByName('Presencial');
    sheet.appendRow([
      new Date().getTime().toString(),
      dados.data,
      dados.cpf,
      dados.nomeCompleto,
      dados.curso||'',
      dados.assunto,
      dados.semestre||'',
      dados.ies||'',
      dados.colaborador,          // col I (índice 8)
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
    ]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// RAMAL — CPF opcional
// Colunas: A-ID | B-Data | C-CPF | D-NomeCompleto | E-Telefone | F-Assunto |
//          G-Observacao | H-Colaborador | I-DataRegistro
function registrarRamal(dados) {
  try {
    var ss = abrirPlanilhaPeriodo_(dados.periodo);
    var sheet = ss.getSheetByName('Ramal');
    sheet.appendRow([
      new Date().getTime().toString(),
      dados.data,
      dados.cpf||'',
      dados.nomeCompleto,
      dados.telefone,
      dados.assunto,
      dados.observacao||'',
      dados.colaborador,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
    ]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// EMAIL
// Colunas: A-ID | B-DataRecebimento | C-CPF | D-NomeCompleto | E-EmailAluno |
//          F-Assunto | G-Semestre | H-IES | I-DataRetorno |
//          J-Observacao | K-Colaborador | L-DataRegistro
function registrarEmail(dados) {
  try {
    var ss = abrirPlanilhaPeriodo_(dados.periodo);
    var sheet = ss.getSheetByName('Email');
    sheet.appendRow([
      new Date().getTime().toString(),
      dados.dataRecebimento,
      dados.cpf||'',
      dados.nomeCompleto||'',
      dados.emailAluno,
      dados.assunto,
      dados.semestre||'',
      dados.ies||'',
      dados.dataRetorno,
      dados.observacao||'',
      dados.colaborador,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
    ]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// VIRTUAL
// Colunas: A-ID | B-Data | C-Colaborador | D-ConversasAtribuidas |
//          E-ConversasFechadas | F-RegistradoPor | G-DataRegistro
function registrarVirtual(dados) {
  try {
    var ss = abrirPlanilhaPeriodo_(dados.periodo);
    var sheet = ss.getSheetByName('Virtual');
    sheet.appendRow([
      new Date().getTime().toString(),
      dados.data,
      dados.colaborador||'',
      parseInt(dados.conversasAtribuidas)||0,
      parseInt(dados.conversasFechadas)||0,
      dados.registradoPor,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss')
    ]);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
}

// abas: nomes possíveis | colabs: colunas (índice 0) que identificam o dono | campos: nome do campo -> índice da coluna
var REG_CONF_ = {
  'Presencial': { abas: ['Presencial'], colabs: [8],
    campos: { data: 1, cpf: 2, nome: 3, curso: 4, assunto: 5, semestre: 6, ies: 7 },
    obrig: ['data', 'cpf', 'nome', 'assunto'] },
  'Ramal': { abas: ['Ramal'], colabs: [7],
    campos: { data: 1, cpf: 2, nome: 3, telefone: 4, assunto: 5, observacao: 6 },
    obrig: ['data', 'nome', 'telefone', 'assunto'] },
  'E-mail': { abas: ['Email', 'E-mail'], colabs: [10],
    campos: { data: 1, cpf: 2, nome: 3, emailAluno: 4, assunto: 5, semestre: 6, ies: 7, dataRetorno: 8, observacao: 9 },
    obrig: ['data', 'emailAluno', 'assunto', 'dataRetorno'] },
  'Virtual': { abas: ['Virtual'], colabs: [2, 5],
    campos: { data: 1, colaborador: 2, atrib: 3, fech: 4 },
    obrig: ['data'] }
};

function localizarRegistro_(ss, conf, id) {
  var sh = null;
  for (var a = 0; a < conf.abas.length && !sh; a++) sh = ss.getSheetByName(conf.abas[a]);
  if (!sh) return null;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || i) === String(id)) return { sh: sh, linha: i + 1, valores: data[i] };
  }
  return null;
}

function podeMexerRegistro_(usr, conf, valores) {
  if (usr.perfil === 'Admin') return true;
  var meu = normNome_(usr.nome);
  for (var k = 0; k < conf.colabs.length; k++) {
    if (normNome_(valores[conf.colabs[k]]) === meu) return true;
  }
  return false;
}

function prepararRegistro_(periodo, tipo, id, userId) {
  var usr = usuarioPorId_(userId);
  if (!usr) return { erro: 'Usuário não identificado.' };
  var conf = REG_CONF_[tipo];
  if (!conf) return { erro: 'Tipo de registro inválido.' };
  var ss = abrirPlanilhaPeriodo_(periodo);
  var reg = localizarRegistro_(ss, conf, id);
  if (!reg) return { erro: 'Registro não encontrado (talvez já tenha sido excluído).' };
  if (!podeMexerRegistro_(usr, conf, reg.valores)) return { erro: 'Você só pode alterar os seus próprios registros.' };
  return { usr: usr, conf: conf, ss: ss, reg: reg };
}

function editarRegistro(periodo, tipo, id, campos, userId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var p = prepararRegistro_(periodo, tipo, id, userId);
    if (p.erro) return { success: false, message: p.erro };
    campos = campos || {};
    for (var o = 0; o < p.conf.obrig.length; o++) {
      var kk = p.conf.obrig[o];
      if (!String(campos[kk] == null ? '' : campos[kk]).trim())
        return { success: false, message: 'Preencha os campos obrigatórios.' };
    }
    Object.keys(p.conf.campos).forEach(function(k) {
      if (!(k in campos)) return;
      var v = campos[k];
      v = (k === 'atrib' || k === 'fech') ? (parseInt(v, 10) || 0) : String(v == null ? '' : v).trim();
      p.reg.sh.getRange(p.reg.linha, p.conf.campos[k] + 1).setValue(v);
    });
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
  finally { try { lock.releaseLock(); } catch(x) {} }
}

// Antes de apagar, guarda uma cópia na aba "Excluidos" da planilha do período
function excluirRegistro(periodo, tipo, id, userId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var p = prepararRegistro_(periodo, tipo, id, userId);
    if (p.erro) return { success: false, message: p.erro };
    var tz = Session.getScriptTimeZone();
    var lx = p.ss.getSheetByName('Excluidos');
    if (!lx) {
      lx = p.ss.insertSheet('Excluidos');
      lx.appendRow(['Tipo', 'Excluído por', 'Data da exclusão', 'Dados originais (colunas da aba de origem)']);
    }
    var copia = p.reg.valores.map(function(v) {
      return v instanceof Date ? Utilities.formatDate(v, tz, 'dd/MM/yyyy HH:mm:ss') : String(v);
    });
    lx.appendRow([tipo, p.usr.nome, Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss')].concat(copia));
    p.reg.sh.deleteRow(p.reg.linha);
    return { success: true };
  } catch(e) { return { success: false, message: e.message }; }
  finally { try { lock.releaseLock(); } catch(x) {} }
}

// ── Histórico ─────────────────────────────────────────────────────────────────
function getHistorico(periodo) {
  try {
    var ss = abrirPlanilhaPeriodo_(periodo);
    var resultado = [];

    function getSheet(n1, n2) {
      var s = ss.getSheetByName(n1);
      return s || (n2 ? ss.getSheetByName(n2) : null);
    }

    // PRESENCIAL — Colaborador: col I (índice 8)
    try {
      var shP = getSheet('Presencial');
      if (shP && shP.getLastRow() > 1) {
        var dtP = shP.getDataRange().getValues();
        for (var i = 1; i < dtP.length; i++) {
          if (dtP[i].length < 9) continue;
          var colabP = String(dtP[i][8]||'').trim();
          resultado.push({
            id: String(dtP[i][0]||i), tipo: 'Presencial',
            colaborador: colabP, registradoPor: colabP,
            data: formatDateValue(dtP[i][1]),
            detalhes: {
              CPF:      String(dtP[i][2]||''),
              Nome:     String(dtP[i][3]||''),
              Curso:    String(dtP[i][4]||''),
              Assunto:  String(dtP[i][5]||''),
              Semestre: formatSemestre(dtP[i][6]),
              IES:      String(dtP[i][7]||'')
            }
          });
        }
      }
    } catch(eP) { Logger.log('Presencial err: ' + eP.message); }

    // RAMAL — Colaborador: col H (índice 7)
    try {
      var shR = getSheet('Ramal');
      if (shR && shR.getLastRow() > 1) {
        var dtR = shR.getDataRange().getValues();
        for (var i = 1; i < dtR.length; i++) {
          if (dtR[i].length < 8) continue;
          var colabR = String(dtR[i][7]||'').trim();
          resultado.push({
            id: String(dtR[i][0]||i), tipo: 'Ramal',
            colaborador: colabR, registradoPor: colabR,
            data: formatDateValue(dtR[i][1]),
            detalhes: {
              CPF:        String(dtR[i][2]||''),
              Nome:       String(dtR[i][3]||''),
              Telefone:   String(dtR[i][4]||''),
              Assunto:    String(dtR[i][5]||''),
              Observacao: String(dtR[i][6]||'')
            }
          });
        }
      }
    } catch(eR) { Logger.log('Ramal err: ' + eR.message); }

    // EMAIL — Colaborador: col K (índice 10) | DataRetorno: col I (índice 8)
    try {
      var shE = getSheet('Email','E-mail');
      if (shE && shE.getLastRow() > 1) {
        var dtE = shE.getDataRange().getValues();
        for (var i = 1; i < dtE.length; i++) {
          if (dtE[i].length < 11) continue;
          var colabE = String(dtE[i][10]||'').trim();
          resultado.push({
            id: String(dtE[i][0]||i), tipo: 'E-mail',
            colaborador: colabE, registradoPor: colabE,
            data: formatDateValue(dtE[i][1]),
            detalhes: {
              CPF:         String(dtE[i][2]||''),
              Nome:        String(dtE[i][3]||''),
              EmailAluno:  String(dtE[i][4]||''),
              Assunto:     String(dtE[i][5]||''),
              Semestre:    formatSemestre(dtE[i][6]),
              IES:         String(dtE[i][7]||''),
              DataRetorno: formatDateValue(dtE[i][8]),
              Observacao:  String(dtE[i][9]||'')
            }
          });
        }
      }
    } catch(eE) { Logger.log('Email err: ' + eE.message); }

    // VIRTUAL — Colaborador: col C (índice 2) | RegistradoPor: col F (índice 5)
    try {
      var shV = getSheet('Virtual');
      if (shV && shV.getLastRow() > 1) {
        var dtV = shV.getDataRange().getValues();
        for (var i = 1; i < dtV.length; i++) {
          if (dtV[i].length < 6) continue;
          resultado.push({
            id: String(dtV[i][0]||i), tipo: 'Virtual',
            colaborador:   String(dtV[i][2]||'').trim(),
            registradoPor: String(dtV[i][5]||'').trim(),
            data: formatDateValue(dtV[i][1]),
            detalhes: {
              Colaborador:         String(dtV[i][2]||''),
              ConversasAtribuidas: String(dtV[i][3]||'0'),
              ConversasFechadas:   String(dtV[i][4]||'0'),
              RegistradoPor:       String(dtV[i][5]||'')
            }
          });
        }
      }
    } catch(eV) { Logger.log('Virtual err: ' + eV.message); }

    return resultado;
  } catch(e) { return { error: e.message }; }
}
