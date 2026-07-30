function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    if (payload.action === 'sendEmailPdf') {
      const ss = getOrCreateSpreadsheet('Checklist TI Lojas - Respostas');
      const sheet = getOrCreateSheet(ss, 'Envios por E-mail');
      const headers = ['Data', 'Destinatários', 'Assunto', 'Resumo', 'Status'];
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(headers);
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setBackground('#1a56db');
        headerRange.setFontColor('#ffffff');
        headerRange.setFontWeight('bold');
        sheet.setFrozenRows(1);
      }
      const row = [
        new Date().toLocaleString('pt-BR'),
        (payload.to || []).join(', '),
        payload.subject || '',
        payload.summary ? `${payload.summary.loja} · ${payload.summary.conformidade}%` : '',
        'Enviado'
      ];
      sheet.appendRow(row);
      sheet.autoResizeColumns(1, headers.length);
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'OK', action: 'sendEmailPdf' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sheetName = payload.sheetName || 'Checklist TI Lojas';
    const headers   = payload.headers || [];
    const rows      = payload.rows || [];

    const ss = getOrCreateSpreadsheet('Checklist TI Lojas - Respostas');
    const sheet = getOrCreateSheet(ss, sheetName);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#1a56db');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    rows.forEach(row => sheet.appendRow(row));
    sheet.autoResizeColumns(1, headers.length);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'OK', rows: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ERROR', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSpreadsheet(name) {
  const files = DriveApp.getFilesByName(name);
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return SpreadsheetApp.create(name);
}

function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

// Teste manual: execute esta função para verificar se o script funciona
function testeManual() {
  const payload = {
    sheetName: 'Checklist TI Lojas',
    headers: ['Loja','Data','Responsável','Regional','Seção','Item','Status','Quantidade','Observação','Score','Progresso','Salvo em'],
    rows: [
      ['Loja Teste','2025-05-19','João Silva','Norte','CFTV','Gravação ativa','ok','','','85%','70/83','19/05/2025 10:30']
    ]
  };
  const e = { postData: { contents: JSON.stringify(payload) } };
  const result = doPost(e);
  Logger.log(result.getContent());
}
