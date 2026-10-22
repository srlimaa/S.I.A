// Planilha MESTRE: guarda apenas as abas "Usuarios" e "Periodos".
// Os registros (Presencial, Ramal, Email, Virtual) ficam na planilha de cada período letivo.
var SPREADSHEET_ID = '18EeUvCq5RFdqynYoK3rtRsdFgQh3lUAaafdHrKiuE7I';

// Período que já usa a planilha acima; é cadastrado automaticamente na primeira execução.
var PERIODO_INICIAL = '2026.2';

// Número da versão exibido no rodapé das Configurações. Suba +1 a cada atualização.
// Versão 77 - 2026-10-17 - Call Center: atribuição em lote agrupada por curso/turno/status, com escolha dos grupos e do responsável de cada um.
// Versão 78 - 2026-10-18 - Call Center: colunas da lista ordenáveis por clique (aluno, curso, status, início e fim do banco) e data final vencida com fundo vermelho.
// Versão 79 - 2026-10-18 - Call Center: botão de download (JPG e Excel) nos gráficos, cartões e tabelas dos Indicadores; cartões do colaborador só com Responsável por, Contatos realizados e não realizados.
// Versão 80 - 2026-10-18 - Call Center: Indicador - Atendimentos e Alunos contatados (tabela por data, atividade e colaborador, com filtros de data/mês/semestre e download em JPG/Excel).
// Versão 81 - 2026-10-19 - Call Center: filtro Responsável lista só os responsáveis ativos (com ao menos um aluno atribuído).
// Versão 82 - 2026-10-19 - Call Center: filtro Resultado do contato (aba Call Center e Indicador - Colaborador); cartões só de colaboradores responsáveis ou com registro.
// Versão 83 - 2026-10-19 - Atendimentos: botões de download só com o ícone (Indicadores, Ranque e Reunião) e botão de olho para ver a senha no login.
// Versão 84 - 2026-10-19 - Call Center: atualização em tempo real sem recarregar, abertura bem mais rápida (cache da base e leituras em memória) e sessão mantida ao recarregar a página.
// Versão 85 - 2026-10-20 - Atendimentos: em Indicadores, ao filtrar mais de uma data (De/Até) aparece a janela Semana com a soma dos dias por atividade e colaborador, com download em Excel e JPG.
// Versão 86 - 2026-10-20 - Call Center: o campo Responsável da lista só pode ser alterado pelo Admin; Colab apenas visualiza (também bloqueado no servidor).
// Versão 87 - 2026-10-20 - Atendimentos: removido da aba Reunião o alerta sobre linhas de ocorrências e anexos (continua em Indicadores).
var VERSAO_SISTEMA = '90';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('Sistema de Atendimentos');
}

// Insere dentro da página o conteúdo de outro arquivo HTML (CSS, JS ou trecho da tela).
// Uso nos arquivos HTML:  <?!= include('nome_do_arquivo') ?>
// Os arquivos js_*, css_* e cc_* ficam embrulhados em <script>/<style> (o Apps Script só aceita HTML válido).
// Aqui o embrulho é retirado, porque a página (Index/CallCenter) já tem a sua própria tag <script>/<style>.
function include(nome) {
  var t = HtmlService.createHtmlOutputFromFile(nome).getContent();
  return t.replace(/^\s*<(script|style)>\r?\n?/, '').replace(/\r?\n?<\/(script|style)>\s*$/, '');
}

function getVersao() { return VERSAO_SISTEMA; }
