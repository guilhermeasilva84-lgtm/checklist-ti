function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'OK', service: 'Checklist TI Lojas', timestamp: new Date().toISOString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: 'ERROR', message: 'Requisição sem corpo JSON.' });
    }

    const payload = JSON.parse(e.postData.contents);

    if (payload.action === 'sendEmailPdf') {
      return handleEmailLog_(payload);
    }

    return handleChecklistSave_(payload);
  } catch (err) {
    return jsonResponse({
      status: 'ERROR',
      message: err && err.message ? err.message : String(err)
    });
  }
}

function handleEmailLog_(payload) {
  const ss = getOrCreateSpreadsheet('Checklist TI Lojas - Respostas');
  const sheet = getOrCreateSheet(ss, 'Envios por E-mail');
  const headers = ['Data', 'Destinatários', 'Assunto', 'Resumo', 'Status'];
  ensureHeaders_(sheet, headers);

  const recipients = Array.isArray(payload.to)
    ? payload.to.filter(Boolean).join(', ')
    : String(payload.to || '');

  const summary = payload.summary || {};
  const loja = summary.loja || payload.loja || '';
  const conformidade = summary.conformidade !== undefined ? summary.conformidade : '';

  sheet.appendRow([
    new Date(),
    recipients,
    payload.subject || '',
    loja ? loja + (conformidade !== '' ? ' · ' + conformidade + '%' : '') : '',
    'Enviado'
  ]);

  sheet.autoResizeColumns(1, headers.length);
  return jsonResponse({ status: 'OK', action: 'sendEmailPdf' });
}

function handleChecklistSave_(payload) {
  const sheetName = sanitizeSheetName_(payload.sheetName || 'Checklist TI Lojas');
  const headers = Array.isArray(payload.headers) ? payload.headers : [];
  const rows = Array.isArray(payload.rows) ? payload.rows : [];

  if (!headers.length) {
    return jsonResponse({ status: 'ERROR', message: 'Nenhum cabeçalho foi informado.' });
  }

  const ss = getOrCreateSpreadsheet('Checklist TI Lojas - Respostas');
  const sheet = getOrCreateSheet(ss, sheetName);
  ensureHeaders_(sheet, headers);

  if (rows.length) {
    const normalizedRows = rows.map(function(row) {
      const values = Array.isArray(row) ? row.slice(0, headers.length) : [];
      while (values.length < headers.length) values.push('');
      return values;
    });
    sheet.getRange(sheet.getLastRow() + 1, 1, normalizedRows.length, headers.length).setValues(normalizedRows);
  }

  sheet.autoResizeColumns(1, headers.length);
  return jsonResponse({ status: 'OK', rows: rows.length, sheetName: sheetName });
}

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() > 0) return;

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange
    .setBackground('#1a56db')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
}

function getOrCreateSpreadsheet(name) {
  const files = DriveApp.getFilesByName(name);
  if (files.hasNext()) {
    return SpreadsheetApp.open(files.next());
  }
  return SpreadsheetApp.create(name);
}

function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function sanitizeSheetName_(name) {
  const cleaned = String(name)
    .replace(/[\\/?*\[\]:]/g, '-')
    .trim()
    .substring(0, 90);
  return cleaned || 'Checklist TI Lojas';
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Teste manual no Apps Script.
function testeManual() {
  const payload = {
    sheetName: 'Checklist TI Lojas',
    headers: ['Loja', 'Data', 'Responsável', 'Regional', 'Seção', 'Item', 'Status', 'Quantidade', 'Observação', 'Score', 'Progresso', 'Salvo em'],
    rows: [
      ['Loja Teste', '2026-08-24', 'João Silva', 'Norte', 'CFTV', 'Gravação ativa', 'ok', '', '', '85%', '70/83', '24/08/2026 10:30']
    ]
  };
  const e = { postData: { contents: JSON.stringify(payload) } };
  const result = doPost(e);
  Logger.log(result.getContent());
}
