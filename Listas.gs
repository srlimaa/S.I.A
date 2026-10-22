// LISTAS GERENCIADAS PELO ADMIN (aba Config, em JSON): colaboradores da opção Virtual, assuntos extras e semestres extras.
// Somam-se às opções fixas dos formulários (html_form_*.html); a tela de configuração está em js_listas_admin.html.
// ══════════════════════════════════════════════════════════════════════
// LISTAS GERENCIADAS PELO ADMIN — colaboradores Virtual, assuntos extras e semestres extras.
// ══════════════════════════════════════════════════════════════════════

// ── Colaboradores da opção Virtual (lista gerenciada pelo Admin) ─────────────
// Guardada na aba "Config" (chave virtual_colabs, JSON). Sem configuração, vale a lista padrão abaixo.
// Lista padrão de colaboradores Virtual, usada quando não há configuração salva.
var VIRTUAL_PADRAO = ['Amanda', 'Andreza', 'Cesar', 'Israel', 'Luana', 'Magno', 'Miriã', 'Patrick'];

// Limpa a lista: tira espaços extras, repetidos (sem acento/maiúscula), nomes longos ou com < >; máximo 100.
function limparNomesVirtual_(lista) {
  var out = [], vistos = {};
  (lista || []).forEach(function(n) {
    n = String(n == null ? '' : n).replace(/\s+/g, ' ').trim();
    if (!n || n.length > 60 || /[<>]/.test(n)) return;
    var k = normNome_(n);
    if (vistos[k]) return;
    vistos[k] = true;
    out.push(n);
  });
  return out.slice(0, 100);
}

// Devolve a lista de colaboradores Virtual (configurada ou padrão).
function getColaboradoresVirtual() {
  try {
    var cfg = lerConfig_();
    if (cfg.virtual_colabs !== undefined && String(cfg.virtual_colabs).trim() !== '') {
      var l = JSON.parse(String(cfg.virtual_colabs));
      if (Object.prototype.toString.call(l) === '[object Array]') return limparNomesVirtual_(l);
    }
  } catch(e) {}
  return VIRTUAL_PADRAO.slice();
}

// Admin: salva a lista de colaboradores Virtual.
function salvarColaboradoresVirtual(lista, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode alterar a lista.' };
    if (Object.prototype.toString.call(lista) !== '[object Array]') return { success: false, message: 'Lista inválida.' };
    var limpa = limparNomesVirtual_(lista);
    setConfig_('virtual_colabs', JSON.stringify(limpa));
    return { success: true, lista: limpa };
  } catch(e) { return { success: false, message: e.message }; }
}

// Troca o nome do colaborador nos registros Virtuais já salvos (todos os períodos cadastrados)
function trocarNomeVirtual(antigo, novo, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin.' };
    antigo = String(antigo || '').trim(); novo = String(novo || '').replace(/\s+/g, ' ').trim();
    if (!antigo || !novo) return { success: false, message: 'Informe o nome antigo e o novo.' };
    if (novo.length > 60 || /[<>]/.test(novo)) return { success: false, message: 'Nome novo inválido.' };
    var alvo = normNome_(antigo), ids = {};
    var per = getPeriodosSheet_().getDataRange().getValues();
    for (var i = 1; i < per.length; i++) { var pid = String(per[i][1]).trim(); if (pid) ids[pid] = true; }
    var total = 0;
    Object.keys(ids).forEach(function(pid) {
      try {
        var sh = SpreadsheetApp.openById(pid).getSheetByName('Virtual');
        if (!sh || sh.getLastRow() < 2) return;
        var rng = sh.getRange(2, 3, sh.getLastRow() - 1, 1);       // coluna C = Colaborador
        var vals = rng.getValues(), mudou = false;
        for (var r = 0; r < vals.length; r++) {
          if (normNome_(vals[r][0]) === alvo && String(vals[r][0]).trim() !== novo) { vals[r][0] = novo; mudou = true; total++; }
        }
        if (mudou) rng.setValues(vals);
      } catch(e) { Logger.log('trocarNomeVirtual falhou em ' + pid + ': ' + e.message); }
    });
    return { success: true, atualizados: total };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── Assuntos extras (Presencial, Ramal e E-mail) — gerenciados pelo Admin ────
// Guardados na aba "Config" (chave assuntos_extra, JSON: [{t:'presencial'|'ramal'|'email', g:'Grupo', n:'Nome'}]).
// Somam-se às opções que já existem no formulário.
// Tipos de formulário que aceitam assuntos extras.
var ASSUNTO_TIPOS_ = ['presencial', 'ramal', 'email'];

// Valida a lista de assuntos extras ({t tipo, g grupo, n nome}), sem repetir dentro do mesmo tipo; máximo 300.
function limparAssuntosExtra_(lista) {
  var out = [], vistos = {};
  (lista || []).forEach(function(a) {
    if (!a) return;
    var t = String(a.t || '');
    if (ASSUNTO_TIPOS_.indexOf(t) === -1) return;
    var n = String(a.n == null ? '' : a.n).replace(/\s+/g, ' ').trim();
    var gr = String(a.g == null ? '' : a.g).replace(/\s+/g, ' ').trim();
    if (!n || n.length > 120 || /[<>]/.test(n)) return;
    if (gr.length > 60 || /[<>]/.test(gr)) return;
    var k = t + '|' + normNome_(n);
    if (vistos[k]) return;
    vistos[k] = true;
    out.push({ t: t, g: gr, n: n });
  });
  return out.slice(0, 300);
}

// Devolve os assuntos extras salvos.
function getAssuntosExtra() {
  try {
    var cfg = lerConfig_();
    if (cfg.assuntos_extra !== undefined && String(cfg.assuntos_extra).trim() !== '') {
      var l = JSON.parse(String(cfg.assuntos_extra));
      if (Object.prototype.toString.call(l) === '[object Array]') return limparAssuntosExtra_(l);
    }
  } catch(e) {}
  return [];
}

// Admin: salva os assuntos extras (limite de tamanho do JSON).
function salvarAssuntosExtra(lista, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode alterar os assuntos.' };
    if (Object.prototype.toString.call(lista) !== '[object Array]') return { success: false, message: 'Lista inválida.' };
    var limpa = limparAssuntosExtra_(lista);
    var txt = JSON.stringify(limpa);
    if (txt.length > 48000) return { success: false, message: 'Limite de assuntos adicionais atingido.' };
    setConfig_('assuntos_extra', txt);
    return { success: true, lista: limpa };
  } catch(e) { return { success: false, message: e.message }; }
}

// ── Semestres extras (Presencial e E-mail) — gerenciados pelo Admin ──────────
// Guardados na aba "Config" (chave semestres_extra, JSON: [{t:'presencial'|'email', n:'2028.1'}]).
// Formato AAAA.S (S = 1 ou 2), o mesmo usado em todo o sistema.
// Tipos de formulário que aceitam semestres extras.
var SEMESTRE_TIPOS_ = ['presencial', 'email'];

// Valida a lista de semestres extras (formato AAAA.1 ou AAAA.2).
function limparSemestresExtra_(lista) {
  var out = [], vistos = {};
  (lista || []).forEach(function(a) {
    if (!a) return;
    var t = String(a.t || '');
    if (SEMESTRE_TIPOS_.indexOf(t) === -1) return;
    var n = String(a.n == null ? '' : a.n).trim();
    if (!/^\d{4}\.[12]$/.test(n)) return;
    var k = t + '|' + n;
    if (vistos[k]) return;
    vistos[k] = true;
    out.push({ t: t, n: n });
  });
  return out.slice(0, 100);
}

// Devolve os semestres extras salvos.
function getSemestresExtra() {
  try {
    var cfg = lerConfig_();
    if (cfg.semestres_extra !== undefined && String(cfg.semestres_extra).trim() !== '') {
      var l = JSON.parse(String(cfg.semestres_extra));
      if (Object.prototype.toString.call(l) === '[object Array]') return limparSemestresExtra_(l);
    }
  } catch(e) {}
  return [];
}

// Admin: salva os semestres extras.
function salvarSemestresExtra(lista, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode alterar os semestres.' };
    if (Object.prototype.toString.call(lista) !== '[object Array]') return { success: false, message: 'Lista inválida.' };
    var limpa = limparSemestresExtra_(lista);
    setConfig_('semestres_extra', JSON.stringify(limpa));
    return { success: true, lista: limpa };
  } catch(e) { return { success: false, message: e.message }; }
}
