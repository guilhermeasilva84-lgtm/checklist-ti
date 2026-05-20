function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheetName = payload.sheetName || 'Checklist TI Lojas';
    const headers   = payload.headers || [];
    const rows      = payload.rows || [];

    // Abre ou cria a planilha
    let ss;
    const files = DriveApp.getFilesByName('Checklist TI Lojas - Respostas');
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create('Checklist TI Lojas - Respostas');
    }

    // Abre ou cria a aba
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    // Cria cabeçalho se a aba estiver vazia
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      // Formata cabeçalho
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#1a56db');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Adiciona as linhas
    rows.forEach(row => sheet.appendRow(row));

    // Auto-resize colunas
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
