function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mail Merge')
    .addItem('Abrir Panel', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Mail Merge')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getGmailAliases() {
  return GmailApp.getAliases();
}

// Obtiene todos los borradores como plantillas
function getGmailDraftTemplates() {
  const drafts = GmailApp.getDrafts();
  // Solo selecciona los borradores con asunto (puedes filtrar mejor si deseas)
  return drafts.map(d => ({
    id: d.getId(),
    subject: d.getMessage().getSubject(),
    body: d.getMessage().getBody()
  }));
}

function getSheetHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return sheet.getDataRange().getValues()[0];
}

// Obtiene el borrador por ID
function getDraftById(draftId) {
  const draft = GmailApp.getDraft(draftId);
  return {
    subject: draft.getMessage().getSubject(),
    body: draft.getMessage().getBody()
  };
}

// Envía mails usando la plantilla seleccionada
function sendBulkEmailsWithTemplate(config) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  // Recupera el borrador
  const draft = GmailApp.getDraft(config.draftId);
  const draftMsg = draft.getMessage();
  const draftSubject = draftMsg.getSubject();
  const draftBody = draftMsg.getBody();

  let count = 0;
  let sent = 0;
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let email = row[headers.indexOf('Email')];
    if (!email) continue;
    let subject = fillTemplate(draftSubject, headers, row);
    let body = fillTemplate(draftBody, headers, row);
    let mailOptions = {
      name: config.senderName,
      from: config.senderEmail
    };
    try {
      GmailApp.sendEmail(email, subject, '', {
        ...mailOptions,
        htmlBody: body
      });
      sent++;
    } catch (e) {
      // Manejar errores si quieres
    }
    count++;
  }
  return {count, sent};
}

function fillTemplate(template, headers, row) {
  let result = template;
  headers.forEach((header, i) => {
    const re = new RegExp(`{{${header}}}`, 'g');
    result = result.replace(re, row[i]);
  });
  return result;
}
