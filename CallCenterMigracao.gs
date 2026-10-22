// Chamado por: botão em Configurações > Planilhas (lkMigrarCC em js_periodos.html > migrarCallCenterAdmin) ou manualmente pelo editor (migrarCallCenter).
// Depende de CallCenterBase.gs (ccAba_, ccTab_, ccSems_, ccSemSalvar_, ccSsReg_, ccEnsure_) e de CC_HEAD para as colunas de cada aba.
// ══════════════════════════════════════════════════════════════════════
// CALL CENTER — migração ÚNICA dos dados da planilha antiga do Call Center para a estrutura do Atendimentos.
//   • CONFIG, OPCOES e USUARIOS  → abas CC_CONFIG, CC_OPCOES e CC_USUARIOS da planilha mestre
//   • SEMESTRES (Base Geral e Call Center) → colunas E e F da aba Periodos
//   • REGISTROS e RESPONSAVEIS de cada semestre → abas CC_REGISTROS e CC_RESPONSAVEIS da planilha do período
// Pode ser executada quantas vezes for preciso: só copia o que ainda não existe (registros pelo Id). Nada é apagado
// na planilha antiga. Execute ANTES do primeiro acesso ao Call Center depois da atualização (ou use o botão em
// ⚙️ › Planilhas), para que as opções e configurações antigas não se misturem com as padrão.
// ══════════════════════════════════════════════════════════════════════

// Para rodar pelo editor do Apps Script: escolha "migrarCallCenter" e clique em Executar (o resultado vai para o Registro de execução)
// Execução manual: roda a migração e grava o relatório no Registro de execução.
function migrarCallCenter() {
  var r = ccMigrar_();
  Logger.log(r.message + '\n' + r.linhas.join('\n'));
  return r;
}

// Para o botão em ⚙️ › Planilhas (somente Admin)
// Execução pelo botão: só Admin; devolve a mensagem e as linhas do relatório para a tela.
function migrarCallCenterAdmin(adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode migrar os dados.' };
    var r = ccMigrar_();
    return { success: true, message: r.message, linhas: r.linhas };
  } catch (e) { return { success: false, message: e.message }; }
}

// Lê uma aba da planilha antiga como lista de objetos (datas viram texto), ignorando linhas vazias.
function ccLerAntiga_(ss, nome) {
  var sh = ss.getSheetByName(nome);
  if (!sh || sh.getLastRow() < 1) return [];
  var v = sh.getDataRange().getValues(), h = v.shift(), tz = ccTz_();
  return v.filter(function(r) { return r.join('') !== ''; }).map(function(r) {
    var o = {};
    h.forEach(function(k, j) { o[k] = (r[j] instanceof Date) ? Utilities.formatDate(r[j], tz, k === 'Semestre' ? 'yyyy.M' : 'yyyy-MM-dd HH:mm') : r[j]; });
    return o;
  });
}

// Acrescenta em lote as linhas (objetos) que ainda não existem na aba de destino. Devolve quantas foram copiadas.
// Acrescenta em lote as linhas que ainda não existem no destino (comparando pela chave informada). Devolve quantas copiou.
function ccMesclar_(nome, linhas, chave, ss) {
  var tem = {}, sh = ccAba_(nome, ss);
  ccTab_(nome, ss).forEach(function(r) { tem[chave(r)] = 1; });
  var novos = linhas.filter(function(r) { var k = chave(r); if (tem[k]) return false; tem[k] = 1; return true; });
  if (!novos.length) return 0;
  var cab = CC_HEAD[nome], ini = sh.getLastRow() + 1, fim = ini + novos.length - 1;
  if (fim > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), fim - sh.getMaxRows());
  var g = sh.getRange(ini, 1, novos.length, cab.length);
  g.setNumberFormat('@');
  g.setValues(novos.map(function(r) { return cab.map(function(k) { return r[k] === undefined || r[k] === null ? '' : String(r[k]); }); }));
  ccInval_(ss, nome);
  return novos.length;
}

// Faz a migração: CONFIG, OPCOES e USUARIOS vão para a planilha mestre; links de Base Geral e Call Center de cada semestre vão para a aba Periodos;
// REGISTROS e RESPONSAVEIS (da planilha principal ou da planilha própria do semestre) vão para a planilha do período.
// Só copia semestres que já existem em Períodos letivos; os demais geram aviso. Usa trava (lock) e limpa as memórias ao final.
function ccMigrar_() {
  var L = [], lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var idAntigo = ccId_(String(lerConfig_().cc_planilha || '').trim()) || CC_PLANILHA_PADRAO;
    if (idAntigo === SPREADSHEET_ID) throw new Error('A planilha antiga do Call Center é a própria planilha mestre. Não há o que migrar.');
    var velha;
    try { velha = SpreadsheetApp.openById(idAntigo); }
    catch (e) { throw new Error('Não foi possível abrir a planilha antiga do Call Center (' + idAntigo + '). Confira o acesso da conta que executa o script.'); }
    if (!velha.getSheetByName('REGISTROS') && !velha.getSheetByName('OPCOES')) throw new Error('A planilha antiga não tem as abas do Call Center (REGISTROS/OPCOES).');

    _CC_MEM = {};
    ccEnsure_(true);                       // cria as abas CC_ vazias na mestre (sem opções padrão)
    var mestre = ccSs_(), n;

    n = ccMesclar_('CONFIG', ccLerAntiga_(velha, 'CONFIG'), function(r) { return String(r.Chave); }, mestre);
    L.push('Configurações copiadas: ' + n);
    n = ccMesclar_('OPCOES', ccLerAntiga_(velha, 'OPCOES'), function(r) { return r.Tipo + '|' + r.Valor; }, mestre);
    L.push('Opções de registro copiadas: ' + n);
    n = ccMesclar_('USUARIOS', ccLerAntiga_(velha, 'USUARIOS'), function(r) { return String(r.Email).toLowerCase(); }, mestre);
    L.push('Preferências de usuários copiadas (tema/semestre): ' + n);

    // Semestres: Base Geral e Call Center vão para a aba Periodos; registros vão para a planilha de cada período
    var antigos = ccLerAntiga_(velha, 'SEMESTRES'), conhecidos = {}, pend = 0;
    var regsPrincipal = ccLerAntiga_(velha, 'REGISTROS'), respPrincipal = ccLerAntiga_(velha, 'RESPONSAVEIS');
    antigos.forEach(function(s) {
      var sem = String(s.Semestre).trim();
      if (!sem) return;
      conhecidos[sem] = 1;
      var atual = ccSems_().find(function(x) { return String(x.Semestre) === sem; });
      if (!atual) { pend++; L.push('⚠️ Semestre ' + sem + ': não existe em Períodos letivos do Atendimentos. Cadastre-o lá e execute a migração de novo.'); return; }
      var base = atual.IdBase || ccId_(s.IdBase), cc = atual.IdCallCenter || ccId_(s.IdCallCenter);
      if (base !== atual.IdBase || cc !== atual.IdCallCenter) { ccSemSalvar_(sem, base, cc); L.push('Semestre ' + sem + ': links da Base Geral/Call Center copiados.'); }

      var regs = regsPrincipal.filter(function(r) { return String(r.Semestre) === sem; });
      var resp = respPrincipal.filter(function(r) { return String(r.Semestre) === sem; });
      var idReg = ccId_(s.IdRegistros);
      if (idReg && idReg !== idAntigo) {                          // semestre que tinha planilha própria de registros
        try {
          var propria = SpreadsheetApp.openById(idReg);
          regs = regs.concat(ccLerAntiga_(propria, 'REGISTROS').filter(function(r) { return String(r.Semestre) === sem; }));
          resp = resp.concat(ccLerAntiga_(propria, 'RESPONSAVEIS').filter(function(r) { return String(r.Semestre) === sem; }));
        } catch (e) { L.push('⚠️ Semestre ' + sem + ': não foi possível abrir a planilha própria de registros; só os da planilha principal foram copiados.'); }
      }
      var destino = ccSsReg_(sem);
      var nr = ccMesclar_('REGISTROS', regs, function(r) { return String(r.Id); }, destino);
      var np = ccMesclar_('RESPONSAVEIS', resp, function(r) { return r.Semestre + '|' + r.CPF; }, destino);
      L.push('Semestre ' + sem + ': ' + nr + ' registro(s) e ' + np + ' responsável(is) copiado(s).');
      ccBump_(sem);
    });
    var orfaos = regsPrincipal.filter(function(r) { return !conhecidos[String(r.Semestre).trim()]; }).length;
    if (orfaos) L.push('⚠️ ' + orfaos + ' registro(s) da planilha antiga são de semestres que não constam na aba SEMESTRES antiga e não foram copiados.');

    _CC_MEM = {};
    CacheService.getScriptCache().remove('cc_planilha');
    return { message: pend ? 'Migração concluída com pendências (veja os avisos).' : 'Migração concluída. A planilha antiga não foi alterada nem apagada.', linhas: L };
  } finally { lock.releaseLock(); }
}
