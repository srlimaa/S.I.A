// UTILITÁRIOS compartilhados: formatação de datas e semestres, normalização de nomes e CPF e conversões de data para planilha.
// ══════════════════════════════════════════════════════════════════════
// UTILITÁRIOS — formatação de datas, semestres e nomes usados por vários arquivos.
// ══════════════════════════════════════════════════════════════════════

// ── Helper: converte qualquer valor de data para DD/MM/AAAA ──────────────────
function formatDateValue(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return ('0'+val.getDate()).slice(-2) + '/' +
           ('0'+(val.getMonth()+1)).slice(-2) + '/' + val.getFullYear();
  }
  var s = String(val);
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    var p = s.split('-'); return p[2]+'/'+p[1]+'/'+p[0];
  }
  var d = new Date(val);
  if (!isNaN(d.getTime()))
    return ('0'+d.getDate()).slice(-2)+'/'+
           ('0'+(d.getMonth()+1)).slice(-2)+'/'+d.getFullYear();
  return s;
}

// ── Helper: converte valor de semestre ───────────────────────────────────────
// Problema: Google Sheets auto-converte "2026.1" para objeto Date ou número float.
// Esta função detecta todos os casos e retorna "AAAA.S" (ex.: "2026.1").
function formatSemestre(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var year = val.getFullYear();
    var month = val.getMonth() + 1;
    return year + '.' + (month <= 6 ? '1' : '2');
  }
  var s = String(val).trim();
  // Já no formato correto
  if (/^\d{4}\.\d$/.test(s)) return s;
  // Número float armazenado pelo Sheets (ex.: 2026.1)
  var n = parseFloat(s);
  if (!isNaN(n) && n > 2000 && n < 2100) {
    var yr = Math.floor(n);
    var dec = Math.round((n - yr) * 10);
    if (dec === 1 || dec === 2) return yr + '.' + dec;
  }
  return s;
}

// Normaliza nome para comparação: minúsculas, sem acento e sem espaços repetidos.
function normNome_(s) {
  s = String(s || '').toLowerCase();
  var mapa = {'á':'a','à':'a','â':'a','ã':'a','ä':'a','é':'e','è':'e','ê':'e','ë':'e','í':'i','ì':'i','î':'i','ï':'i',
              'ó':'o','ò':'o','ô':'o','õ':'o','ö':'o','ú':'u','ù':'u','û':'u','ü':'u','ç':'c','ñ':'n'};
  s = s.replace(/[áàâãäéèêëíìîïóòôõöúùûüçñ]/g, function(ch) { return mapa[ch]; });
  return s.replace(/\s+/g, ' ').trim();
}

// Devolve sempre a mesma grafia para um nome (primeira encontrada), para somar corretamente nos indicadores.
function nomeCanonico_(nome, canon) {
  var k = normNome_(nome);
  if (!canon[k]) canon[k] = String(nome).replace(/\s+/g, ' ').trim();
  return canon[k];
}

// Converte data (Date, DD/MM/AAAA ou AAAA-MM-DD) para DD/MM/AAAA; null se inválida. Com tz usa o fuso da planilha.
function dataBR_(val, tz) {
  if (!val) return null;
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return null;
    if (tz) return Utilities.formatDate(val, tz, 'dd/MM/yyyy');   // data no fuso da planilha
    return ('0' + val.getDate()).slice(-2) + '/' + ('0' + (val.getMonth() + 1)).slice(-2) + '/' + val.getFullYear();
  }
  var s = String(val).trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) {
    var ano = m[3].length === 2 ? '20' + m[3] : m[3];
    return ('0' + m[1]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + ano;
  }
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[3] + '/' + m[2] + '/' + m[1];
  return null;
}

// true se o valor é nulo ou só espaços.
function vazio_(v) { return String(v == null ? '' : v).trim() === ''; }

// true se a célula de anexo está vazia ou só tem traços ("--"), isto é, o aluno não tem anexo.
function semAnexo_(v) { var t = String(v == null ? '' : v).trim(); return !t || /^[-–—\s]+$/.test(t); }

// Mantém só os dígitos do CPF.
function cpfNum_(v) { return String(v == null ? '' : v).replace(/\D/g, ''); }

// Texto de data DD/MM/AAAA no fuso informado (ou o texto da célula).
function dataTxt_(v, tz) {
  if (v instanceof Date) return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz, 'dd/MM/yyyy');
  return String(v == null ? '' : v).trim();
}

// Número de série da data (dias desde 30/12/1899, sem hora): grava só a DATA, sem depender de fuso horário
function serialDia_(y, m, d) { return Math.round(Date.UTC(y, m - 1, d) / 86400000) + 25569; }
