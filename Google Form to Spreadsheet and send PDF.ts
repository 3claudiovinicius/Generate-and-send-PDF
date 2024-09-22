const EMAIL_OVERRIDE = true;
const EMAIL_ADDRESS_OVERRIDE = 'claudio_soares@whirlpool.com';
const APP_TITLE = 'Generate and send PDFs';
const OUTPUT_FOLDER_NAME = "Component Requests PDFs";

const DATA_SHEET_NAME = 'Form Responses 1';
const REQUEST_TEMPLATE_SHEET_NAME = 'Request Template';

/**
 * Trigger function to process data when a new row is added.
 */
function onEdit(e) {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() === DATA_SHEET_NAME && e.range.getRow() > 1) {
      const lock = LockService.getScriptLock();
      try {
        // Tenta obter o lock, esperando até 10 segundos se necessário
        lock.waitLock(10000);
        
        // Aqui você pode adicionar um delay se necessário
        Utilities.sleep(600000); // 10 minutos
        
        // Processa a linha
        processRow(e.range.getRow());
      } catch (error) {
        Logger.log('Erro ao obter lock: ' + error);
      } finally {
        // Libera o lock, independentemente de ter ocorrido um erro
        lock.releaseLock();
      }
    }
  }
  

/**
 * Processes the specific row to generate a PDF and send an email.
 * 
 * @param {number} row - The row number to process
 */
function processRow(row) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    const templateSheet = ss.getSheetByName(REQUEST_TEMPLATE_SHEET_NAME);
  
    // Check if the row is already processed
    const status = dataSheet.getRange(row, 40).getValue(); // Supondo que a coluna 40 seja a coluna de status
    if (status === "Processed") {
      return; // Skip if already processed
    }
  
    // Get data from the specified row (C to AL)
    const data = dataSheet.getRange(row, 3, 1, 38).getValues()[0]; // C até AL
  
    // Validate data (you can customize the validation logic as needed)
    if (!data[0] || !data[1] || !data[2]) { // Exemplo: verificar colunas C, D e E
      Logger.log('Missing data in row ' + row);
      return;
    }
  
    // Fill in the template with the data
    updateTemplateSheet(templateSheet, data);
  
    // Generate and send PDF
    const pdf = generatePDF(ss.getId(), templateSheet, `Request-${data[0]}`); // Nome do PDF baseado na coluna C
    sendEmail(data[1], pdf); // Enviar para o email na coluna D
  
    // Mark row as processed
    dataSheet.getRange(row, 40).setValue("Processed"); // Marcar como processada
}

/**
 * Updates the template sheet with data.
 * 
 * @param {object} templateSheet - Template sheet to update
 * @param {array} data - Data to set in the template
 */
function updateTemplateSheet(templateSheet, data) {
  const ranges = {
    'B10': data[0],  // ColumC
    'B11': data[1],  // ColumD
    'B12': data[2],  // ColumE
    'B13': data[3],  // ColumF
    'B14': data[4],  // ColumG
    'B15': data[5],  // ColumH
    'B16': data[6],  // ColumI
    'B17': data[7],  // ColumJ
    'B18': data[8],  // ColumK
    'B19': data[9],  // ColumL
    'B20': data[10], // ColumM
    'B21': data[11], // ColumN
    'B22': data[12], // ColumO
    'B23': data[13], // ColumP
    'B24': data[14], // ColumQ
    'B25': data[15], // ColumR
    'B26': data[16], // ColumS
    'B27': data[17], // ColumT
    'B28': data[18], // ColumU
    'B29': data[19], // ColumV
    'B30': data[20], // ColumW
    'B31': data[21], // ColumX
    'B32': data[22], // ColumY
    'B33': data[23], // ColumZ
    'B34': data[24], // ColumAA
    'B35': data[25], // ColumAB
    'B36': data[26], // ColumAC
    'B37': data[27], // ColumAD
    'B38': data[28], // ColumAE
    'B39': data[29], // ColumAF
    'B40': data[30], // ColumAG
    'B41': data[31], // ColumAH
    'B42': data[32], // ColumAI
    'B43': data[33], // ColumAJ
    'B44': data[34], // ColumAK
    'B45': data[35]  // ColumAL
  };
  
  for (const range in ranges) {
    templateSheet.getRange(range).setValue(ranges[range]);
  }
}

/**
 * Generates a PDF for the given sheet.
 * 
 * @param {string} ssId - Google Spreadsheet ID
 * @param {object} sheet - Sheet to be converted to PDF
 * @param {string} pdfName - File name for the PDF
 * @return {blob} PDF file as a blob
 */
function generatePDF(ssId, sheet, pdfName) {
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=pdf&size=7&fzr=true&portrait=true&fitw=true&gridlines=false&printtitle=false&top_margin=0.5&bottom_margin=0.25&left_margin=0.5&right_margin=0.5&sheetnames=false&pagenum=UNDEFINED&attachment=true&gid=${sheet.getSheetId()}`;
  
  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(`${pdfName}.pdf`);

  return savePDF(blob);
}

/**
 * Saves the PDF to a folder in Google Drive.
 * 
 * @param {blob} pdfBlob - PDF file blob
 * @return {file} PDF file
 */
function savePDF(pdfBlob) {
  const folder = getOrCreateFolder(OUTPUT_FOLDER_NAME);
  return folder.createFile(pdfBlob);
}

/**
 * Sends an email with the PDF attachment.
 * 
 * @param {string} recipient - Email address of the recipient
 * @param {blob} pdfBlob - PDF file blob
 */
function sendEmail(recipient, pdfBlob) {
  GmailApp.sendEmail(
    EMAIL_OVERRIDE ? EMAIL_ADDRESS_OVERRIDE : recipient,
    'Component Request Notification',
    'Hello!\rPlease see the attached PDF document.',
    { attachments: [pdfBlob], name: APP_TITLE }
  );
}

/**
 * Gets or creates a folder by name.
 * 
 * @param {string} name - Name of the folder
 * @return {folder} Folder object
 */
function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

