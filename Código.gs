// Planilha MESTRE: guarda apenas as abas "Usuarios" e "Periodos".
// Os registros (Presencial, Ramal, Email, Virtual) ficam na planilha de cada período letivo.
var SPREADSHEET_ID = '18EeUvCq5RFdqynYoK3rtRsdFgQh3lUAaafdHrKiuE7I';

// Período que já usa a planilha acima; é cadastrado automaticamente na primeira execução.
var PERIODO_INICIAL = '2026.2';

// Versão 47 - 2026-10-15 - <resumo curto da mudança>.
var VERSAO_SISTEMA = '48';

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
