// ══════════════════════════════════════════════════════════════════════
// NOME NA TELA DE LOGIN — texto próprio para o título do login.
// Não altera o nome do sistema (escolha do sistema, topo e aba do navegador).
// Vazio = o login continua usando o nome do sistema.
// Guardado na configuração do sistema, na chave "nome_login".
// ══════════════════════════════════════════════════════════════════════

// Pública (chamada antes do login, como getConfigSistema)
function getNomeLogin() {
  try { return String(lerConfig_().nome_login || ''); }
  catch (e) { return ''; }
}

function salvarNomeLogin(nome, adminId) {
  try {
    if (!ehAdmin_(adminId)) return { success: false, message: 'Apenas Admin pode alterar o nome da tela de login.' };
    nome = String(nome == null ? '' : nome).replace(/\s+/g, ' ').trim();
    if (nome.length > 60) return { success: false, message: 'Use no máximo 60 caracteres.' };
    if (/[<>]/.test(nome)) return { success: false, message: 'O texto não pode conter < ou >.' };
    setConfig_('nome_login', nome);          // vazio = volta a usar o nome do sistema
    return { success: true, nome: nome };
  } catch (e) { return { success: false, message: e.message }; }
}
