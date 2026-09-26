// Projeto separado, executado exclusivamente pela conta nutri4nutri@gmail.com.
// Ele processa a fila criada pelo backend principal sem alterar a API ou os acessos.
const SHEET_ID = '1Bw9RxJ872eb4-f3hciloXyb1wZmSTKPs9Bw_LoYeurQ';
const FILA_EMAILS_SHEET = 'FILA EMAILS';
const EMAILS_ENVIADOS_SHEET = 'EMAILS ENVIADOS';
const EMAIL_SENDER = 'nutri4nutri@gmail.com';

function instalarEnvioPelaContaDaPriscila() {
  const effectiveEmail = normalizeEmail(Session.getEffectiveUser().getEmail());
  if (effectiveEmail !== EMAIL_SENDER) {
    throw new Error('Execute esta função conectada como ' + EMAIL_SENDER + '. Conta atual: ' + (effectiveEmail || 'não identificada'));
  }
  ensureEmailQueueSheet();
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'processarFilaEmailsPriscila') ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger('processarFilaEmailsPriscila').timeBased().everyMinutes(1).create();
  return { ok: true, remetenteReal: effectiveEmail, intervaloMinutos: 1 };
}

function processarFilaEmailsPriscila() {
  const effectiveEmail = normalizeEmail(Session.getEffectiveUser().getEmail());
  if (effectiveEmail !== EMAIL_SENDER) {
    throw new Error('Envio bloqueado: o acionador deve pertencer a ' + EMAIL_SENDER + '. Conta atual: ' + (effectiveEmail || 'não identificada'));
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: true, skipped: 'locked' };
  try {
    const sheet = ensureEmailQueueSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: true, processados: 0 };
    const rows = sheet.getRange(2, 1, lastRow - 1, 13).getValues();
    let sent = 0;
    let failed = 0;

    rows.forEach(function(values, index) {
      const row = index + 2;
      const status = String(values[6] || '').toUpperCase();
      const attempts = Number(values[10] || 0);
      const lastAttempt = values[12] instanceof Date ? values[12] : new Date(values[12] || 0);
      const staleProcessing = status === 'PROCESSANDO' && (!lastAttempt.getTime() || Date.now() - lastAttempt.getTime() > 5 * 60000);
      const eligible = status === 'PENDENTE' || status === 'ERRO' || staleProcessing;
      if (!eligible || attempts >= 3) return;

      const now = new Date();
      sheet.getRange(row, 7).setValue('PROCESSANDO');
      sheet.getRange(row, 11).setValue(attempts + 1);
      sheet.getRange(row, 13).setValue(now);
      try {
        MailApp.sendEmail({
          to: String(values[2] || ''),
          subject: String(values[3] || ''),
          body: String(values[4] || ''),
          htmlBody: String(values[5] || ''),
          name: 'Priscila Leite',
          replyTo: EMAIL_SENDER
        });
        sheet.getRange(row, 7, 1, 4).setValues([['ENVIADO', new Date(), effectiveEmail, '']]);
        updateWelcomeEmailLogStatus(String(values[11] || ''), 'ENVIADO', '');
        sent++;
      } catch (err) {
        const errorText = String(err && err.message ? err.message : err);
        sheet.getRange(row, 7).setValue('ERRO');
        sheet.getRange(row, 10).setValue(errorText);
        updateWelcomeEmailLogStatus(String(values[11] || ''), 'ERRO', errorText);
        failed++;
      }
    });
    return { ok: failed === 0, processados: sent + failed, enviados: sent, falhas: failed, remetenteReal: effectiveEmail };
  } finally {
    lock.releaseLock();
  }
}

function ensureEmailQueueSheet() {
  const sheet = ensureSheet(FILA_EMAILS_SHEET, [
    'Criado Em','Chave','Destinatário','Assunto','Texto','HTML','Status',
    'Enviado Em','Remetente Real','Erro','Tentativas','Chave do Log','Última Tentativa'
  ]);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,13).setBackground('#123D31').setFontColor('#FFFFFF').setFontWeight('bold');
  return sheet;
}

function updateWelcomeEmailLogStatus(logKey, status, errorText) {
  if (!logKey) return;
  const log = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EMAILS_ENVIADOS_SHEET);
  if (!log || log.getLastRow() < 2) return;
  const row = findRowByValue(log, 2, logKey);
  if (row < 2) return;
  log.getRange(row, 8, 1, 2).setValues([[status, String(errorText || '')]]);
}

function ensureSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0 && headers && headers.length) sheet.appendRow(headers);
  return sheet;
}

function findRowByValue(sheet, column, value) {
  if (!value || sheet.getLastRow() < 2) return -1;
  const finder = sheet.getRange(2, column, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(value)).matchEntireCell(true).findNext();
  return finder ? finder.getRow() : -1;
}

function normalizeEmail(value) {
  return String(value || '').toLowerCase().trim();
}
