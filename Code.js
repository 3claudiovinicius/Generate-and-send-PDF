/**
 * Projeto: Gerador e Envio de PDFs para Solicitações de Componentes
 * 
 * Descrição:
 * Este script automatiza o processo de geração de PDFs a partir de respostas de um formulário
 * no Google Sheets e envia esses PDFs por email para os destinatários especificados.
 * O código é acionado automaticamente quando uma linha do formulário é editada.
 * 
 * Principais Funcionalidades:
 * 1. Processa automaticamente as respostas de um formulário no Google Sheets.
 * 2. Preenche um modelo de planilha com dados do formulário.
 * 3. Gera um PDF do modelo preenchido.
 * 4. Envia o PDF por email para os destinatários configurados.
 * 5. Controla o processo de envio e evita duplicações, marcando as solicitações como "Processadas".
 * 
 * Variáveis:
 * - EMAIL_OVERRIDE: Ativa ou desativa a substituição de email com uma lista predefinida.
 * - EMAIL_ADDRESS_OVERRIDE: Lista de emails para os quais os PDFs serão enviados (se EMAIL_OVERRIDE for verdadeiro).
 * - APP_TITLE: Título do aplicativo usado no envio do email.
 * - OUTPUT_FOLDER_NAME: Nome da pasta no Google Drive onde os PDFs serão armazenados.
 * - DATA_SHEET_NAME: Nome da planilha de dados onde as respostas do formulário são armazenadas.
 * - REQUEST_TEMPLATE_SHEET_NAME: Nome da planilha de modelo usada para gerar o PDF.
 * 
 * Funções:
 * - atEdit(e): Função principal que é chamada ao editar uma célula no Google Sheets.
 * - processRow(row): Processa uma linha específica do formulário e gera o PDF.
 * - updateTemplateSheet(templateSheet, formattedDate, data): Atualiza o modelo com os dados da linha.
 * - generateAndSendPDF(ssId, sheetId, pdfName, recipient): Gera e envia o PDF por email.
 * - getOrCreateFolder(name): Cria uma pasta no Google Drive se ela não existir.
 * 
 * @author [Claudio Ohara]
 * @date [2024-Nov-09]
 * @version 1.0
 */

const EMAIL_OVERRIDE = true;
const EMAIL_ADDRESS_OVERRIDE = ['amily_han@whirlpool.com', 'regis_t_feroldi@whirlpool.com', 'swift_xu@whirlpool.com', 'claudio_soares@whirlpool.com'];
//const EMAIL_ADDRESS_OVERRIDE = ['claudio_soares@whirlpool.com'];
const APP_TITLE = 'Generate and send PDFs';
const OUTPUT_FOLDER_NAME = "Component Requests PDFs";
const DATA_SHEET_NAME = 'Form Responses 1';
const REQUEST_TEMPLATE_SHEET_NAME = 'Request Template';

function atEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() === DATA_SHEET_NAME && e.range.getRow() > 1) {
    processRow(e.range.getRow());
  }
}

function processRow(row) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);  // Tenta adquirir o lock por até 10 segundos

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(DATA_SHEET_NAME);
    const templateSheet = ss.getSheetByName(REQUEST_TEMPLATE_SHEET_NAME);
    
    const dateTime = dataSheet.getRange(row, 1).getValue();
    const date = new Date(dateTime);
    const formattedDate = `${String(date.getDate()).padStart(2, '0')}/${String(date.getMonth() + 1).padStart(2, '0')}/${date.getFullYear()}`;

    const status = dataSheet.getRange(row, 40).getValue();
    if (status === "Processed") return;

    const data = dataSheet.getRange(row, 3, 1, 38).getValues()[0];
    const requiredColumns = [0, 1, 2];
    if (requiredColumns.some(index => !data[index])) {
      Logger.log(`Missing required data in row ${row}`);
      return;
    }

    const originalTemplateValues = templateSheet.getDataRange().getValues();

    updateTemplateSheet(templateSheet, formattedDate, data);
    Utilities.sleep(5000);

    const testCell = templateSheet.getRange("A1").getValue();  // Leitura do valor
    templateSheet.getRange("A1").setValue(testCell); // Escreve o mesmo valor para forçar a atualização

    const sheetID = templateSheet.getSheetId();
    generateAndSendPDF(ss.getId(), sheetID, `Request-${data[0]}`, data[1]);

    templateSheet.getDataRange().setValues(originalTemplateValues);

    dataSheet.getRange(row, 40).setValue("Processed");

  } catch (e) {
    Logger.log("Failed to process row ${row} due to locking: " + e);
  } finally {
    lock.releaseLock();  // Libera o lock após o processamento
  }
}

function updateTemplateSheet(templateSheet, formattedDate, data) {
  // Mapeia os aliases (placeholders) com os dados correspondentes de C a AL
  const aliases = {
    '{{Date}}': formattedDate,
    '{{ColumC}}': data[0],
    '{{ColumD}}': data[1],
    '{{ColumE}}': data[2],
    '{{ColumF}}': data[3],
    '{{ColumG}}': data[4],
    '{{ColumH}}': data[5],
    '{{ColumI}}': data[6],
    '{{ColumJ}}': data[7],
    '{{ColumK}}': data[8],
    '{{ColumL}}': data[9],
    '{{ColumM}}': data[10],
    '{{ColumN}}': data[11],
    '{{ColumO}}': data[12],
    '{{ColumP}}': data[13],
    '{{ColumQ}}': data[14],
    '{{ColumR}}': data[15],
    '{{ColumS}}': data[16],
    '{{ColumT}}': data[17],
    '{{ColumU}}': data[18],
    '{{ColumV}}': data[19],
    '{{ColumW}}': data[20],
    '{{ColumX}}': data[21],
    '{{ColumY}}': data[22],
    '{{ColumZ}}': data[23],
    '{{ColumAA}}': data[24],
    '{{ColumAB}}': data[25],
    '{{ColumAC}}': data[26],
    '{{ColumAD}}': data[27],
    '{{ColumAE}}': data[28],
    '{{ColumAF}}': data[29],
    '{{ColumAG}}': data[30],
    '{{ColumAH}}': data[31],
    '{{ColumAI}}': data[32],
    '{{ColumAJ}}': data[33],
    '{{ColumAK}}': data[34],
    '{{ColumAL}}': data[35]
  };

  // Substitui os placeholders pelo valor correspondente no template
  const range = templateSheet.getDataRange();
  const values = range.getValues();

  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[row].length; col++) {   

      let cellValue = values[row][col];

      if (cellValue === null || cellValue === undefined) {
        cellValue = "N/A";
      } else {
        // Converte cellValue em string para garantir que `includes` funcione
        cellValue = String(cellValue);
      }


      for (let alias in aliases) {
        if (cellValue.includes(alias)) {
          cellValue = cellValue.replace(alias, aliases[alias]);
        }
      }
      values[row][col] = cellValue;
    }
  }

  // Atualiza a planilha com os valores modificados
  range.setValues(values);
}


function generateAndSendPDF(ssId, sheetId, pdfName, recipient) {
  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=pdf&size=A4&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false&gid=${sheetId}`;


  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(`${pdfName}.pdf`);

  const folder = getOrCreateFolder(OUTPUT_FOLDER_NAME);
  folder.createFile(blob);

  const finalRecipient = EMAIL_OVERRIDE ? EMAIL_ADDRESS_OVERRIDE.join(',') : recipient;
  GmailApp.sendEmail(finalRecipient, 'Component Request Notification', 'Hello!\rPlease see the attached PDF document.', { attachments: [blob], name: APP_TITLE });
}


function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}
