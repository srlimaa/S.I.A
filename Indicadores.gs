// ══════════════════════════════════════════════════════════════════════
// INDICADORES — soma dos atendimentos por dia/atividade/colaborador.
// ══════════════════════════════════════════════════════════════════════

// ── Indicadores ───────────────────────────────────────────────────────────────
function getIndicadores(periodo) {
  try {
    var ss = abrirPlanilhaPeriodo_(periodo);
    var ind = {};
    var colaboradores = [];

    function addColab(c) {
      if (c && String(c).trim() && colaboradores.indexOf(String(c).trim()) === -1)
        colaboradores.push(String(c).trim());
    }
    function ensure(d, t) {
      if (!ind[d]) ind[d] = {};
      if (!ind[d][t]) ind[d][t] = {};
    }
    function toStr(val) {
      if (!val) return null;
      if (val instanceof Date)
        return ('0'+val.getDate()).slice(-2)+'/'+
               ('0'+(val.getMonth()+1)).slice(-2)+'/'+val.getFullYear();
      var s = String(val);
      if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
      if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
        var p = s.split('-'); return p[2]+'/'+p[1]+'/'+p[0];
      }
      var d2 = new Date(val);
      if (!isNaN(d2.getTime()))
        return ('0'+d2.getDate()).slice(-2)+'/'+
               ('0'+(d2.getMonth()+1)).slice(-2)+'/'+d2.getFullYear();
      return s;
    }

    var sh, dt, d, c;

    // PRESENCIAL — Data: col B (índice 1) | Colaborador: col I (índice 8)
    sh = ss.getSheetByName('Presencial');
    if (sh) {
      dt = sh.getDataRange().getValues();
      for (var i = 1; i < dt.length; i++) {
        if (!dt[i][0]) continue;
        d = toStr(dt[i][1]); c = String(dt[i][8]||'').trim();
        if (!d||!c) continue;
        addColab(c); ensure(d,'Presencial');
        ind[d]['Presencial'][c] = (ind[d]['Presencial'][c]||0) + 1;
      }
    }

    // RAMAL — Data: col B (índice 1) | Colaborador: col H (índice 7)
    sh = ss.getSheetByName('Ramal');
    if (sh) {
      dt = sh.getDataRange().getValues();
      for (var i = 1; i < dt.length; i++) {
        if (!dt[i][0]) continue;
        d = toStr(dt[i][1]); c = String(dt[i][7]||'').trim();
        if (!d||!c) continue;
        addColab(c); ensure(d,'Ramal');
        ind[d]['Ramal'][c] = (ind[d]['Ramal'][c]||0) + 1;
      }
    }

    // EMAIL — DataRetorno: col I (índice 8) | Colaborador: col K (índice 10)
    sh = ss.getSheetByName('Email') || ss.getSheetByName('E-mail');
    if (sh) {
      dt = sh.getDataRange().getValues();
      for (var i = 1; i < dt.length; i++) {
        if (!dt[i][0]) continue;
        d = toStr(dt[i][8]); c = String(dt[i][10]||'').trim();
        if (!d||!c) continue;
        addColab(c); ensure(d,'E-mail');
        ind[d]['E-mail'][c] = (ind[d]['E-mail'][c]||0) + 1;
      }
    }

    // VIRTUAL — Data: col B (índice 1) | Colaborador: col C (índice 2)
    //           Atribuídas: col D (índice 3) | Fechadas: col E (índice 4)
    sh = ss.getSheetByName('Virtual');
    if (sh) {
      dt = sh.getDataRange().getValues();
      for (var i = 1; i < dt.length; i++) {
        if (!dt[i][0]) continue;
        d = toStr(dt[i][1]); c = String(dt[i][2]||'').trim();
        var atrib = parseInt(dt[i][3])||0;
        var fech  = parseInt(dt[i][4])||0;
        if (!d||!c) continue;
        addColab(c);
        ensure(d,'Conversa Atribuída'); ensure(d,'Conversa Fechada');
        ind[d]['Conversa Atribuída'][c] = (ind[d]['Conversa Atribuída'][c]||0) + atrib;
        ind[d]['Conversa Fechada'][c]   = (ind[d]['Conversa Fechada'][c]||0)   + fech;
      }
    }

    // ── Planilhas de ocorrências/anexos (registro manual; só leitura) ──
    var avisos = [];
    var canon = {}, ctx = null;
    try {
      colaboradores.forEach(function(n) { canon[normNome_(n)] = n; });
      try {
        var us = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Usuarios').getDataRange().getValues();
        for (var u = 1; u < us.length; u++) {
          var un = String(us[u][1]).trim(), uk = normNome_(un);
          if (un && !canon[uk]) canon[uk] = un;
        }
      } catch(eu) {}
      ctx = { ind: ind, addColab: addColab, ensure: ensure, canon: canon, avisos: avisos };
    } catch(ex0) { avisos.push('Erro ao preparar leitura externa: ' + ex0.message); }

    // planilhas de ocorrências/anexos (registro manual; só leitura)
    try {
      var links = getLinksPeriodo_(periodo);
      if (ctx && links.unileao) {
        try { lerOcorrenciasPlanilha_(links.unileao, 6, 'UNILEÃO', ctx); }
        catch(e1) { avisos.push('Não foi possível ler a planilha de ocorrências da UNILEÃO: ' + e1.message); }
      }
      if (ctx && links.univs) {
        try { lerOcorrenciasPlanilha_(links.univs, 8, 'UNIVS', ctx); }
        catch(e2) { avisos.push('Não foi possível ler a planilha de ocorrências da UNIVS: ' + e2.message); }
      }
    } catch(ex) { avisos.push('Erro ao ler as planilhas de ocorrências: ' + ex.message); }

    // registros feitos no sistema de Call Center
    try { if (ctx) lerRegistrosCallCenter_(periodo, ctx); }
    catch(ex2) { avisos.push('Não foi possível ler os registros do Call Center: ' + ex2.message); }

    colaboradores.sort();
    return { indicadores: ind, colaboradores: colaboradores, avisos: avisos };
  } catch(e) {
    return { indicadores: {}, colaboradores: [], error: e.message };
  }
}
