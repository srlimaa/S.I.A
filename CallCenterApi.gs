// CALL CENTER: tabela de ações da API (CC_API) e despachante ccApi(). Cada ação tem f(usuário, argumentos); "admin: 1" exige perfil Admin e "light: 1" é chamada leve.
// Quem chama: cc_js.html (função call), pela ponte do iframe e js_callcenter.html. Planilhas, semestres e cache ficam em CallCenterBase.gs.
// Semestres = períodos letivos do Atendimentos; aqui só se informam os links da Base Geral e do Call Center de cada um.
// ══════════════════════════════════════════════════════════════════════
// CALL CENTER — tabela de ações da API (CC_API) e despachante ccApi().
// ══════════════════════════════════════════════════════════════════════

const CC_API = {
  // Primeira chamada da tela: usuário, configuração (nomes e logo), semestres (ccSemNomes_) e opções de registro.
  boot: { f: u => ({ user: ccPub_(u), cfg: ccCfg_(), sems: ccSemNomes_(), opc: ccOpc_() }) },

  // Dados do semestre: base de alunos (cache de 10 min; Admin pode forçar), registros, responsáveis, usuários e versão. Falha se o semestre não tem Base Geral.
  getData: {
    f: (u, a) => {
      const sem = a.sem || ccSemU_(u);
      const r = ccSems_().find(x => String(x.Semestre) === String(sem));
      if (!r || !r.IdBase) throw new Error('Semestre sem Base Geral configurada. Peça ao Admin: Configurações > Semestres.');
      const b = ccBase_(r, sem, !!a.fresh && u.Perfil === 'Admin');   // base de alunos: do cache (10 min) ou lida da planilha
      const resp = {};
      const regs = ccRegs_(sem), info = ccUsuariosInfo_(regs.map(r => r.colab));
      ccTab_('RESPONSAVEIS', ccSsReg_(sem)).filter(x => String(x.Semestre) === String(sem)).forEach(x => resp[x.CPF] = x.Email);
      return {
        sem: sem, diag: b.diag, alunos: b.alunos, light: u.Perfil === 'Admin' ? b.light : [], regs: regs, resp: resp,
        usuarios: info.usuarios, extras: info.extras, v: ccVer_(sem), bts: b.ts
      };
    }
  },

  // Atualização em tempo real: chamada LEVE (não relê a planilha quando nada mudou). Devolve os registros e os
  // responsáveis só quando a versão do semestre mudou (ou quando a tela pede conferência completa, v vazio).
  sync: {
    light: 1, f: (u, a) => {
      const sem = String(a.sem || '').trim();
      if (!sem) throw new Error('Semestre não informado.');
      const v = ccVer_(sem);
      if (a.v && String(a.v) === v) return { v: v };
      const resp = {};
      ccTab_('RESPONSAVEIS', ccSsReg_(sem)).filter(x => String(x.Semestre) === sem).forEach(x => resp[x.CPF] = x.Email);
      return { v: v, regs: ccRegs_(sem), resp: resp };
    }
  },

  // Admin: define o responsável de um aluno (aba CC_RESPONSAVEIS da planilha do período).
  setResp: {
    admin: 1, f: (u, a) => {   // só Admin escolhe/altera o responsável (Colab apenas visualiza)
      const l = LockService.getScriptLock(); l.waitLock(20000);
      try {
        const ssR = ccSsReg_(a.sem);
        const x = ccTab_('RESPONSAVEIS', ssR).find(r => String(r.Semestre) === String(a.sem) && r.CPF === a.cpf);
        if (x) ccUpd_('RESPONSAVEIS', x._row, { Email: a.email }, ssR); else ccAdd_('RESPONSAVEIS', { Semestre: a.sem, CPF: a.cpf, Email: a.email }, ssR);
        ccBump_(a.sem);
      } finally { l.releaseLock(); }
      return true;
    }
  },

  // Atribui (ou remove) o responsável de vários alunos de uma vez. email vazio = remover o responsável.
  setRespLote: {
    admin: 1, f: (u, a) => {
      const cpfs = (Array.isArray(a.cpfs) ? a.cpfs : []).map(x => String(x)).filter(Boolean);
      if (!cpfs.length) throw new Error('Nenhum aluno selecionado.');
      if (cpfs.length > 5000) throw new Error('Seleção grande demais. Aplique em partes menores.');
      const email = String(a.email || '').trim();
      const sem = String(a.sem || '').trim();
      if (!sem) throw new Error('Semestre não informado.');
      if (email && !ccUsuariosInfo_([]).usuarios.some(x => x.email === email.toLowerCase())) throw new Error('Responsável inválido.');
      const alvo = {}; cpfs.forEach(c => alvo[c] = 1);
      const l = LockService.getScriptLock(); l.waitLock(30000);
      try {
        const ssR = ccSsReg_(sem), sh = ssR.getSheetByName(ccAbaNome_('RESPONSAVEIS'));
        const ultima = sh.getLastRow(), achados = {};
        if (ultima >= 2) {                                           // atualiza as linhas existentes de uma só vez
          const vals = sh.getRange(2, 1, ultima - 1, 3).getValues();  // Semestre, CPF, Email
          vals.forEach(r => { if (String(r[0]) === sem && alvo[String(r[1])]) { r[2] = email; achados[String(r[1])] = 1; } });
          const col = sh.getRange(2, 3, ultima - 1, 1); col.setNumberFormat('@'); col.setValues(vals.map(r => [r[2]]));
        }
        if (email) {                                                 // quem ainda não tinha linha ganha uma nova
          const novos = cpfs.filter(c => !achados[c]).map(c => [sem, c, email]);
          if (novos.length) { const g2 = sh.getRange(sh.getLastRow() + 1, 1, novos.length, 3); g2.setNumberFormat('@'); g2.setValues(novos); }
        }
        ccInval_(ssR, 'RESPONSAVEIS'); ccBump_(sem);
      } finally { l.releaseLock(); }
      return cpfs.length;
    }
  },

  // Grava um novo contato ou edita um existente (somente quem registrou, ou Admin). Exige objetivo, meio e resultado.
  saveReg: {
    f: (u, a) => {
      if (!a.objetivo || !a.meio || !a.resultado) throw new Error('Preencha objetivo, meio e resultado.');
      const l = LockService.getScriptLock(); l.waitLock(20000);
      try {
        const agora = Utilities.formatDate(new Date(), ccTz_(), 'yyyy-MM-dd HH:mm');
        const ssR = ccSsReg_(a.sem);
        if (a.id) {
          const r = ccTab_('REGISTROS', ssR).find(x => x.Id === a.id);
          if (!r) throw new Error('Registro não encontrado.');
          if (String(r.ColabEmail).toLowerCase() !== u.Email && u.Perfil !== 'Admin') throw new Error('Só quem registrou (ou Admin) pode editar.');
          ccUpd_('REGISTROS', r._row, { Objetivo: a.objetivo, Meio: a.meio, Resultado: a.resultado, Observacao: a.obs || '', EditadoEm: agora }, ssR);
        } else {
          ccAdd_('REGISTROS', { Id: Utilities.getUuid(), Semestre: a.sem, CPF: a.cpf, Aluno: a.aluno, DataHora: agora, Objetivo: a.objetivo, Meio: a.meio, Resultado: a.resultado, Observacao: a.obs || '', ColabEmail: u.Email, ColabNome: u.Nome, Status: a.status }, ssR);
        }
        ccBump_(a.sem);
      } finally { l.releaseLock(); }
      return ccRegs_(a.sem);
    }
  },

  // Exclui um registro de contato (somente Admin). Localiza a linha pelo Id na própria planilha (não depende de _row).
  regDel: {
    admin: 1, f: (u, a) => {
      const id = String(a.id || '').trim();
      if (!id) throw new Error('Registro não informado.');
      const l = LockService.getScriptLock(); l.waitLock(20000);
      try {
        const ssR = ccSsReg_(a.sem), sh = ssR.getSheetByName(ccAbaNome_('REGISTROS'));
        const ultima = sh ? sh.getLastRow() : 0;
        if (ultima < 2) throw new Error('Registro não encontrado.');
        const ids = sh.getRange(2, 1, ultima - 1, 1).getValues();
        let linha = 0;
        for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === id) { linha = i + 2; break; } }
        if (!linha) throw new Error('Registro não encontrado (talvez já tenha sido excluído).');
        sh.deleteRow(linha);
        ccInval_(ssR, 'REGISTROS'); ccBump_(a.sem);
      } finally { l.releaseLock(); }
      return ccRegs_(a.sem);
    }
  },

  // Guarda o semestre escolhido pelo usuário.
  setSem: { f: (u, a) => { ccUpd_('USUARIOS', u._row, { Semestre: a.sem }); return true; } },
  // Guarda o tema do usuário (padrão ou escuro).
  setTema: { f: (u, a) => { ccUpd_('USUARIOS', u._row, { Tema: a.tema }); return true; } },
  // Admin: adiciona uma opção de Objetivo, Meio ou Resultado.
  opcAdd: { admin: 1, f: (u, a) => { if (!a.valor) throw new Error('Informe o texto.'); ccAdd_('OPCOES', { Tipo: a.tipo, Valor: a.valor }); return ccOpc_(); } },
  // Salva a nova ordem das opções de um grupo (Objetivo, Meio ou Resultado). Só reorganiza as linhas desse grupo,
  // dentro das posições que elas já ocupavam; os outros grupos e opções fora da lista não mudam.
  opcOrdem: {
    admin: 1, f: (u, a) => {
      const tipo = String(a.tipo || ''), novo = (Array.isArray(a.valores) ? a.valores : []).map(x => String(x));
      if (['Objetivo', 'Meio', 'Resultado'].indexOf(tipo) < 0) throw new Error('Grupo inválido.');
      const l = LockService.getScriptLock(); l.waitLock(20000);
      try {
        ccTab_('OPCOES');   // garante a aba
        const ss = ccSs_(), sh = ss.getSheetByName(ccAbaNome_('OPCOES')), ultima = sh.getLastRow();
        if (ultima >= 2) {
          const vals = sh.getRange(2, 1, ultima - 1, 2).getValues();
          const idx = [], atuais = [];
          vals.forEach((r, i) => { if (String(r[0]) === tipo) { idx.push(i); atuais.push(String(r[1])); } });
          const ordem = novo.filter(v => atuais.indexOf(v) >= 0).filter((v, i, arr) => arr.indexOf(v) === i);
          atuais.forEach(v => { if (ordem.indexOf(v) < 0) ordem.push(v); });   // opção que alguém criou nesse meio tempo vai para o fim
          idx.forEach((i, k) => { vals[i][1] = ordem[k]; });
          const g = sh.getRange(2, 1, vals.length, 2); g.setNumberFormat('@'); g.setValues(vals);
          ccInval_(ss, 'OPCOES');
        }
      } finally { l.releaseLock(); }
      return ccOpc_();
    }
  },
  // Admin: exclui uma opção.
  opcDel: { admin: 1, f: (u, a) => { const r = ccTab_('OPCOES').find(x => x.Tipo === a.tipo && x.Valor === a.valor); if (r) ccDel_('OPCOES', r._row); return ccOpc_(); } },

  // Admin: salva nome do login, nome do sistema ou logo (limite de tamanho).
  cfgSave: {
    admin: 1, f: (u, a) => {
      if (['nomeLogin', 'nomeSistema', 'logo'].indexOf(a.chave) < 0) throw new Error('Chave inválida.');
      if (String(a.valor).length > 45000) throw new Error('Imagem muito grande.');
      const r = ccTab_('CONFIG').find(x => x.Chave === a.chave);
      if (r) ccUpd_('CONFIG', r._row, { Valor: a.valor }); else ccAdd_('CONFIG', { Chave: a.chave, Valor: a.valor });
      return ccCfg_();
    }
  },

  // Semestres = Períodos letivos do Atendimentos. Aqui só se informam a Base Geral e a planilha do Call Center de cada um.
  // Admin: lista os semestres (períodos letivos) com os links de Base Geral e Call Center.
  semList: {
    admin: 1, f: () => ccSems_().map(x => ({ sem: String(x.Semestre), base: x.IdBase, cc: x.IdCallCenter }))
  },
  // Admin: grava os links de Base Geral e Call Center de um semestre que já existe em Períodos letivos (ccSemSalvar_).
  semSave: {
    admin: 1, f: (u, a) => {
      const s = String(a.sem || '').trim();
      if (!s) throw new Error('Informe o semestre (ex.: 2027.1).');
      ccSemSalvar_(s, a.base, a.cc);
      return CC_API.semList.f();
    }
  },
  // "Excluir" só limpa os links do Call Center; o período em si continua no Atendimentos
  // Admin: "excluir" só limpa os links do Call Center; o período e os registros continuam.
  semDel: { admin: 1, f: (u, a) => { ccSemSalvar_(String(a.sem || ''), '', ''); return CC_API.semList.f(); } }
};

// Despachante único chamado pelo navegador: confere a sessão (token no cache), a existência da ação e a permissão Admin; devolve {ok} ou {err}.
// Ações "light" só conferem a sessão (não leem as planilhas de usuários) e renovam a validade da sessão.
function ccApi(token, fn, a) {
  const cache = CacheService.getScriptCache(), email = cache.get('s_' + token);
  if (!email) return { err: 'SESSAO' };
  const h = CC_API[fn];
  if (!h) return { err: 'Função inválida.' };
  // chamadas "light" (sincronização) só conferem a sessão: não releem as planilhas de usuários
  const u = h.light ? { Email: email, Perfil: '' } : ccUser_(email);   // usuário do Atendimentos com acesso liberado ao Call Center
  if (!u) return { err: 'SESSAO' };
  if (h.admin && u.Perfil !== 'Admin') return { err: 'Sem permissão.' };
  if (h.light) cache.put('s_' + token, email, 21600);                 // sessão renovada enquanto a tela estiver aberta
  try { return { ok: h.f(u, a || {}) }; } catch (e) { return { err: String(e.message || e) }; }
}
