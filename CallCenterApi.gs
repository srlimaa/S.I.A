// ══════════════════════════════════════════════════════════════════════
// CALL CENTER — tabela de ações da API (CC_API) e despachante ccApi().
// ══════════════════════════════════════════════════════════════════════

const CC_API = {
  boot: { f: u => ({ user: ccPub_(u), cfg: ccCfg_(), sems: ccTab_('SEMESTRES').map(s => String(s.Semestre)), opc: ccOpc_() }) },

  getData: {
    f: (u, a) => {
      const sem = a.sem || ccSemU_(u);
      const r = ccTab_('SEMESTRES').find(x => String(x.Semestre) === String(sem));
      if (!r || !r.IdBase) throw new Error('Semestre sem planilha configurada. Peça ao Admin: Configurações > Semestres.');
      const sb = SpreadsheetApp.openById(ccId_(r.IdBase));
      const bsh = sb.getSheetByName('BASE GERAL') || sb.getSheets()[0];
      const base = bsh.getDataRange().getValues();
      const h0 = 1; // cabeçalhos na linha 2; dados a partir da linha 3
      const comp = r.IdCallCenter ? ccComp_(r.IdCallCenter) : {};
      const alunos = [], light = [], st0 = {};
      let tot = 0;
      for (let i = h0 + 1; i < base.length; i++) {
        const v = base[i], nome = String(v[3]).trim();
        if (!nome) continue;
        const st = ccCanon_(v[40]);
        st0[st] = (st0[st] || 0) + 1;
        tot++;
        // Regra do Indicador - Geral: coluna AB = Matriculado e coluna AL contendo a palavra FIES
        const mat = /^MATRICULAD[OA]\b/.test(ccKey_(v[27])), fies = ccKey_(v[37]).indexOf('FIES') >= 0;
        if (mat && fies) light.push([st, String(v[4]).trim(), String(v[5]).trim()]);
        if (CC_ST.indexOf(st) < 0) continue;
        const cpf = ccCpf_(v[2]);
        const al = {
          cpf: cpf, cpfFmt: String(v[2]).trim(), nome: nome, curso: String(v[4]).trim(), turno: String(v[5]).trim(), status: st,
          semC: String(v[8]).trim(), qtd: String(v[9]).trim(), perc: ccFmtP_(v[10]), dil: ccFmtD_(v[29]),
          tel: [12, 13, 14, 15, 16].map(k => String(v[k]).trim()).filter(x => /\d/.test(x)),
          email: [17, 18].map(k => String(v[k]).trim()).filter(x => x.indexOf('@') > 0)
        };
        if (st === CC_ST[1] && comp[cpf]) al.banco = comp[cpf];
        alunos.push(al);
      }
      const resp = {};
      ccTab_('RESPONSAVEIS', ccSsReg_(sem)).filter(x => String(x.Semestre) === String(sem)).forEach(x => resp[x.CPF] = x.Email);
      return {
        sem: sem, diag: { aba: bsh.getName(), linhas: base.length, cab: h0 + 1, status: st0, total: tot, indicador: light.length }, alunos: alunos, light: u.Perfil === 'Admin' ? light : [], regs: ccRegs_(sem), resp: resp,
        usuarios: ccTab_('USUARIOS').map(x => ({ email: x.Email, nome: x.Nome, foto: x.Foto }))
      };
    }
  },

  setResp: {
    f: (u, a) => {
      const l = LockService.getScriptLock(); l.waitLock(20000);
      try {
        const ssR = ccSsReg_(a.sem);
        const x = ccTab_('RESPONSAVEIS', ssR).find(r => String(r.Semestre) === String(a.sem) && r.CPF === a.cpf);
        if (x) ccUpd_('RESPONSAVEIS', x._row, { Email: a.email }, ssR); else ccAdd_('RESPONSAVEIS', { Semestre: a.sem, CPF: a.cpf, Email: a.email }, ssR);
      } finally { l.releaseLock(); }
      return true;
    }
  },

  // Atribui (ou remove) o responsável de vários alunos de uma vez. email vazio = remover o responsável.
  setRespLote: {
    f: (u, a) => {
      const cpfs = (Array.isArray(a.cpfs) ? a.cpfs : []).map(x => String(x)).filter(Boolean);
      if (!cpfs.length) throw new Error('Nenhum aluno selecionado.');
      if (cpfs.length > 5000) throw new Error('Seleção grande demais. Aplique em partes menores.');
      const email = String(a.email || '').trim();
      const sem = String(a.sem || '').trim();
      if (!sem) throw new Error('Semestre não informado.');
      if (email && !ccTab_('USUARIOS').some(x => String(x.Email) === email)) throw new Error('Responsável inválido.');
      const alvo = {}; cpfs.forEach(c => alvo[c] = 1);
      const l = LockService.getScriptLock(); l.waitLock(30000);
      try {
        const sh = ccSsReg_(sem).getSheetByName('RESPONSAVEIS');
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
      } finally { l.releaseLock(); }
      return cpfs.length;
    }
  },

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
          if (r.ColabEmail !== u.Email && u.Perfil !== 'Admin') throw new Error('Só quem registrou (ou Admin) pode editar.');
          ccUpd_('REGISTROS', r._row, { Objetivo: a.objetivo, Meio: a.meio, Resultado: a.resultado, Observacao: a.obs || '', EditadoEm: agora }, ssR);
        } else {
          ccAdd_('REGISTROS', { Id: Utilities.getUuid(), Semestre: a.sem, CPF: a.cpf, Aluno: a.aluno, DataHora: agora, Objetivo: a.objetivo, Meio: a.meio, Resultado: a.resultado, Observacao: a.obs || '', ColabEmail: u.Email, ColabNome: u.Nome, Status: a.status }, ssR);
        }
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
        const sh = ccSsReg_(a.sem).getSheetByName('REGISTROS');
        const ultima = sh ? sh.getLastRow() : 0;
        if (ultima < 2) throw new Error('Registro não encontrado.');
        const ids = sh.getRange(2, 1, ultima - 1, 1).getValues();
        let linha = 0;
        for (let i = 0; i < ids.length; i++) { if (String(ids[i][0]) === id) { linha = i + 2; break; } }
        if (!linha) throw new Error('Registro não encontrado (talvez já tenha sido excluído).');
        sh.deleteRow(linha);
      } finally { l.releaseLock(); }
      return ccRegs_(a.sem);
    }
  },

  setSem: { f: (u, a) => { ccUpd_('USUARIOS', u._row, { Semestre: a.sem }); return true; } },
  setTema: { f: (u, a) => { ccUpd_('USUARIOS', u._row, { Tema: a.tema }); return true; } },
  setFoto: { f: (u, a) => { if (String(a.foto).length > 45000) throw new Error('Imagem muito grande.'); ccUpd_('USUARIOS', u._row, { Foto: a.foto }); return true; } },
  opcAdd: { admin: 1, f: (u, a) => { if (!a.valor) throw new Error('Informe o texto.'); ccAdd_('OPCOES', { Tipo: a.tipo, Valor: a.valor }); return ccOpc_(); } },
  opcDel: { admin: 1, f: (u, a) => { const r = ccTab_('OPCOES').find(x => x.Tipo === a.tipo && x.Valor === a.valor); if (r) ccDel_('OPCOES', r._row); return ccOpc_(); } },

  cfgSave: {
    admin: 1, f: (u, a) => {
      if (['nomeLogin', 'nomeSistema', 'logo'].indexOf(a.chave) < 0) throw new Error('Chave inválida.');
      if (String(a.valor).length > 45000) throw new Error('Imagem muito grande.');
      const r = ccTab_('CONFIG').find(x => x.Chave === a.chave);
      if (r) ccUpd_('CONFIG', r._row, { Valor: a.valor }); else ccAdd_('CONFIG', { Chave: a.chave, Valor: a.valor });
      return ccCfg_();
    }
  },

  usersList: { admin: 1, f: () => ccTab_('USUARIOS').map(x => ({ email: x.Email, nome: x.Nome, perfil: x.Perfil })) },
  userSave: {
    admin: 1, f: (u, a) => {
      const email = String(a.email || '').trim().toLowerCase();
      if (!email || !a.nome) throw new Error('Informe nome e e-mail.');
      if (['Admin', 'Colab'].indexOf(a.perfil) < 0) throw new Error('Perfil inválido.');
      const us = ccTab_('USUARIOS');
      const old = a.old ? us.find(x => String(x.Email).toLowerCase() === String(a.old).toLowerCase()) : null;
      if (us.some(x => String(x.Email).toLowerCase() === email && (!old || x._row !== old._row))) throw new Error('E-mail já cadastrado.');
      if (old) {
        if (old.Perfil === 'Admin' && a.perfil !== 'Admin' && us.filter(x => x.Perfil === 'Admin').length < 2) throw new Error('Precisa existir ao menos um Admin.');
        const o = { Email: email, Nome: a.nome, Perfil: a.perfil };
        if (a.senha) o.Senha = ccHashNew_(a.senha);
        ccUpd_('USUARIOS', old._row, o);
      } else {
        if (!a.senha) throw new Error('Informe a senha.');
        ccAdd_('USUARIOS', { Email: email, Nome: a.nome, Senha: ccHashNew_(a.senha), Perfil: a.perfil, Tema: 'Padrão' });
      }
      return CC_API.usersList.f();
    }
  },
  userDel: {
    admin: 1, f: (u, a) => {
      const us = ccTab_('USUARIOS'), r = us.find(x => x.Email === a.email);
      if (!r) return CC_API.usersList.f();
      if (r.Email === u.Email) throw new Error('Você não pode excluir seu próprio usuário.');
      if (r.Perfil === 'Admin' && us.filter(x => x.Perfil === 'Admin').length < 2) throw new Error('Precisa existir ao menos um Admin.');
      ccDel_('USUARIOS', r._row); return CC_API.usersList.f();
    }
  },

  semList: {
    admin: 1, f: () => {
      const principal = ccPlanilhaId_(), regsPrincipal = ccTab_('REGISTROS');
      return ccTab_('SEMESTRES').map(x => {
        const reg = ccId_(x.IdRegistros), propria = !!reg && reg !== principal;
        // registros antigos do semestre que ficaram na planilha principal (não são movidos automaticamente)
        const legado = propria ? regsPrincipal.filter(r => String(r.Semestre) === String(x.Semestre)).length : 0;
        return { sem: String(x.Semestre), base: x.IdBase, cc: x.IdCallCenter, reg: reg, legado: legado };
      });
    }
  },
  semSave: {
    admin: 1, f: (u, a) => {
      const s = String(a.sem || '').trim();
      if (!s) throw new Error('Informe o semestre (ex.: 2027.1).');
      const reg = ccId_(a.reg);
      if (reg) {                                  // planilha própria de registros deste semestre
        let ssR;
        try { ssR = SpreadsheetApp.openById(reg); }
        catch (e) { throw new Error('Não foi possível abrir a planilha de registros. Confira o link e se a conta que executa o script tem acesso (editor).'); }
        ccPrepararRegistros_(ssR);                // cria as abas REGISTROS e RESPONSAVEIS com os cabeçalhos
      }
      ccGarantirColunaSem_();
      const o = { Semestre: s, IdBase: ccId_(a.base), IdCallCenter: ccId_(a.cc), IdRegistros: reg };
      const r = ccTab_('SEMESTRES').find(x => String(x.Semestre) === s);
      if (r) ccUpd_('SEMESTRES', r._row, o); else ccAdd_('SEMESTRES', o);
      return CC_API.semList.f();
    }
  },
  semDel: { admin: 1, f: (u, a) => { const r = ccTab_('SEMESTRES').find(x => String(x.Semestre) === String(a.sem)); if (r) ccDel_('SEMESTRES', r._row); return CC_API.semList.f(); } }
};

function ccApi(token, fn, a) {
  const email = CacheService.getScriptCache().get('s_' + token);
  if (!email) return { err: 'SESSAO' };
  const u = ccTab_('USUARIOS').find(x => x.Email === email);
  if (!u) return { err: 'SESSAO' };
  const h = CC_API[fn];
  if (!h) return { err: 'Função inválida.' };
  if (h.admin && u.Perfil !== 'Admin') return { err: 'Sem permissão.' };
  try { return { ok: h.f(u, a || {}) }; } catch (e) { return { err: String(e.message || e) }; }
}
