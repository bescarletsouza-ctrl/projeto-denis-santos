/**
 * Recebe as respostas da pesquisa e grava uma linha por resposta na planilha.
 *
 * Como instalar / atualizar:
 * 1. Na planilha, menu Extensões > Apps Script. Apague o conteúdo e cole este arquivo.
 * 2. Primeira vez: Implantar > Nova implantação > tipo "App da Web"
 *    (Executar como: Eu · Quem pode acessar: Qualquer pessoa).
 *    Para ATUALIZAR mantendo a mesma URL: Implantar > Gerenciar implantações >
 *    lápis (editar) > Versão: "Nova versão" > Implantar.
 * 3. A URL (termina em /exec) vai em ENDPOINT no index.html.
 *
 * A aba "Respostas" é criada com as colunas abaixo. Se já existir uma aba
 * "Respostas" com colunas diferentes (versão antiga da pesquisa), ela é
 * renomeada para "Respostas (antiga dd-MM HHhmm)" e uma nova é criada.
 */

const SHEET_NAME = "Respostas";

// [chave enviada pelo formulário, título da coluna na planilha]
const COLUMNS = [
  ["data_hora", "Data/hora"],
  // Bloco 1 — Perfil e atuação profissional
  ["faixa_etaria", "1. Faixa etária"],
  ["ocupacao", "2. Ocupação"],
  ["email", "3. E-mail"],
  ["segmento", "4. Segmento"],
  ["faturamento", "5. Faturamento mensal"],
  // Bloco 2 — Dores, desafios e desejos
  ["area_desenvolver", "6. Área que deseja desenvolver"],
  ["maior_desafio", "7. Maior desafio"],
  ["transformacao", "8. Transformação desejada"],
  // Bloco 3 — Conhecimento e interesse em psicanálise
  ["nivel_psicanalise", "9. Conhecimento em psicanálise"],
  ["interesse_psicanalise", "10. Interesse em estudar psicanálise"],
  ["assunto_interesse", "11. Assunto de interesse"],
  ["formato_aprendizado", "12. Formato preferido"],
  // Origem do tráfego
  ["utm_source", "utm_source"],
  ["utm_medium", "utm_medium"],
  ["utm_campaign", "utm_campaign"],
  ["utm_content", "utm_content"],
  ["utm_term", "utm_term"],
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // evita linhas sobrepostas com envios simultâneos

  try {
    const sheet = getSheet_();
    const data = Object.assign({ data_hora: new Date() }, e.parameter);
    sheet.appendRow(COLUMNS.map(([key]) => data[key] ?? ""));

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const titles = COLUMNS.map(([, title]) => title);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (sheet && sheet.getLastColumn()) {
    const current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (current.join("|") === titles.join("|")) return sheet;
    const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd-MM HH'h'mm");
    sheet.setName(`${SHEET_NAME} (antiga ${stamp})`);
    sheet = null;
  }

  sheet = sheet || ss.insertSheet(SHEET_NAME, 0);
  sheet.getRange(1, 1, 1, titles.length).setValues([titles])
    .setFontWeight("bold").setBackground("#f1efe9");
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(2, titles.length - 1, 220);
  return sheet;
}
