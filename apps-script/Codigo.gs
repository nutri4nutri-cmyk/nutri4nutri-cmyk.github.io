const SHEET_ID = '1Bw9RxJ872eb4-f3hciloXyb1wZmSTKPs9Bw_LoYeurQ';
const SHEET_NAME = 'CRM';
const COMU_FREE_SHEET = 'COMU FREE';
const EVENTOS_SHEET = 'EVENTOS';
const ALUNAS_WORK_SHEET = 'ALUNAS WORK';
const LEADS_QUENTE_SHEET = 'LEADS QUENTE';
const PAGAMENTOS_SHEET = 'Pagamentos'; // compatibilidade com compras antigas
const ASAAS_BASE_URL = 'https://api.asaas.com/v3';
const WORKSHOP_PUBLIC_LINK_SLUG = 'lreonttfy8mnzycj';

function getAsaasKey() {
  return PropertiesService.getScriptProperties().getProperty('ASAAS_API_KEY');
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'verificar') {
    const email = e.parameter.email || '';
    const curso = e.parameter.curso || '';
    return jsonOutput(verificarAcesso(email, curso));
  }
  return jsonOutput({ status: 'ok' });
}

function doPost(e) {
  try {
    if (e && e.parameter && e.parameter.email) {
      gravarLead(
        e.parameter.name || '',
        e.parameter.email || '',
        e.parameter.phone || '',
        e.parameter.course || '',
        e.parameter.type || ''
      );
      return jsonOutput({ ok: true });
    }

    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // Webhooks antigos do Asaas podem continuar chegando. A sincronização financeira
    // é feita exclusivamente pela API e pelo acionador de tempo.
    if (data.event && data.payment) {
      return jsonOutput({ ok: true, ignored: true, source: 'webhook-disabled' });
    }

    const sheet = ensureSheet(SHEET_NAME, [
      'Data / Hora', 'Nome', 'Email', 'Telefone', 'Origem',
      'Perfil Profissional', 'Principal Dificuldade', 'Objetivo',
      'Texto Livre', 'Curso Sugerido'
    ]);
    sheet.appendRow([
      new Date(),
      data.nome || '',
      data.email || '',
      data.telefone || '',
      data.tag || 'quiz',
      data.pergunta1 || '',
      data.pergunta2 || '',
      data.pergunta3 || '',
      data.textoLivre || '',
      data.cursoSugerido || ''
    ]);
    refreshDashboard();
    return jsonOutput({ ok: true });
  } catch (err) {
    console.error('doPost:', err);
    return jsonOutput({ ok: false, err: String(err) });
  }
}

function gravarLead(name, email, phone, course, type) {
  email = normalizeEmail(email);
  if (!email) return;

  const now = new Date();
  const origem = course ? (type ? type + ':' + course : course) : (type || 'site');
  const crm = ensureSheet(SHEET_NAME, [
    'Data / Hora', 'Nome', 'Email', 'Telefone', 'Origem',
    'Perfil Profissional', 'Principal Dificuldade', 'Objetivo',
    'Texto Livre', 'Curso Sugerido'
  ]);
  crm.appendRow([now, name, email, phone, origem, '', '', '', '', '']);

  if (course === 'comunidade') {
    upsertComuFree(now, name, email, phone, origem, type || 'acesso');
  }
  refreshDashboard();
}

function upsertComuFree(now, name, email, phone, origem, evento) {
  const sheet = ensureSheet(COMU_FREE_SHEET, [
    'Data / Hora', 'Nome', 'Email', 'Telefone', 'Origem',
    'Primeiro Evento', 'Último Evento'
  ]);
  const row = findRowByValue(sheet, 3, email);
  if (row > 1) {
    const current = sheet.getRange(row, 1, 1, 7).getValues()[0];
    sheet.getRange(row, 1, 1, 7).setValues([[
      current[0] || now,
      name || current[1] || '',
      email,
      phone || current[3] || '',
      origem || current[4] || 'comunidade',
      current[5] || evento,
      evento
    ]]);
  } else {
    sheet.appendRow([now, name, email, phone, origem || 'comunidade', evento, evento]);
  }
}

function verificarAcesso(email, curso) {
  email = normalizeEmail(email);
  if (!email) return { acesso: false, diasDesdeCompra: 0 };
  if (curso === 'workshop') return verificarWorkshopAPI(email);

  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(PAGAMENTOS_SHEET);
    if (!sheet) return { acesso: false, diasDesdeCompra: 0 };
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const rowEmail = normalizeEmail(rows[i][2]);
      const rowCourse = String(rows[i][4] || '').toLowerCase().trim();
      if (rowEmail === email && (!curso || rowCourse === String(curso).toLowerCase().trim())) {
        return { acesso: true, diasDesdeCompra: daysSince(rows[i][0]) };
      }
    }
  } catch (err) {
    console.error('verificarAcesso:', err);
  }
  return { acesso: false, diasDesdeCompra: 0 };
}

function verificarWorkshopAPI(email) {
  try {
    const customers = asaasGet('/customers?email=' + encodeURIComponent(email)).data || [];
    for (let i = 0; i < customers.length; i++) {
      const payments = asaasGet('/payments?limit=100&customer=' + encodeURIComponent(customers[i].id)).data || [];
      for (let j = 0; j < payments.length; j++) {
        const payment = payments[j];
        if (isWorkshopPayment(payment) && isPaidStatus(payment.status)) {
          return {
            acesso: true,
            diasDesdeCompra: daysSince(payment.paymentDate || payment.clientPaymentDate || payment.dateCreated)
          };
        }
      }
    }
  } catch (err) {
    console.error('verificarWorkshopAPI:', err);
  }
  return { acesso: false, diasDesdeCompra: 0 };
}

function setupSistema() {
  ensureAllSheets();
  syncAsaasEvents();
  instalarSincronizacao();
}

function instalarSincronizacao() {
  const exists = ScriptApp.getProjectTriggers().some(function(trigger) {
    return trigger.getHandlerFunction() === 'syncAsaasEvents';
  });
  if (!exists) {
    ScriptApp.newTrigger('syncAsaasEvents').timeBased().everyMinutes(15).create();
  }
}

function syncAsaasEvents() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    ensureAllSheets();
    const payments = listAsaasPayments();
    const eventSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(EVENTOS_SHEET);
    const eventState = buildEventState(eventSheet);
    const customerCache = {};

    payments.forEach(function(payment) {
      const workshop = isWorkshopPayment(payment);
      if (!workshop) return;

      const customer = getCustomer(payment.customer, customerCache);
      const statusPt = translateStatus(payment.status);
      const billingTypePt = translateBillingType(payment.billingType);
      const signature = [statusPt, billingTypePt].join('|');
      const previous = eventState[payment.id];

      if (previous !== signature) {
        const eventName = previous ? 'Status atualizado' : 'Cobrança identificada';
        eventSheet.appendRow([
          new Date(), payment.id || '', eventName, statusPt, billingTypePt,
          payment.dateCreated || '', payment.paymentDate || payment.clientPaymentDate || '',
          customer.name || '', customer.email || '', customer.phone || customer.mobilePhone || '',
          payment.description || '', payment.externalReference || '', payment.value || 0,
          'Workshop Seletividade Alimentar'
        ]);
        eventState[payment.id] = signature;
      }

      if (isPaidStatus(payment.status)) {
        upsertAlunaWorkshop(payment, customer);
        markLeadAsConverted(payment, customer);
      } else {
        upsertLeadQuente(payment, customer, 'NÃO CONVERTIDO');
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function listAsaasPayments() {
  const result = [];
  let offset = 0;
  for (let page = 0; page < 30; page++) {
    const response = asaasGet('/payments?limit=100&offset=' + offset);
    const data = response.data || [];
    Array.prototype.push.apply(result, data);
    if (!response.hasMore || data.length === 0) break;
    offset += data.length;
  }
  return result;
}

function getCustomer(customerId, cache) {
  if (!customerId) return {};
  if (!cache[customerId]) {
    try {
      cache[customerId] = asaasGet('/customers/' + encodeURIComponent(customerId)) || {};
    } catch (err) {
      cache[customerId] = {};
    }
  }
  return cache[customerId];
}

function isWorkshopPayment(payment) {
  const linkId = getWorkshopPaymentLinkId();
  const paymentLink = payment.paymentLink || payment.paymentLinkId || '';
  if (linkId && paymentLink && String(paymentLink) === String(linkId)) return true;

  const text = [payment.description, payment.externalReference].join(' ').toLowerCase();
  return text.indexOf('workshop') >= 0 && text.indexOf('seletividade') >= 0;
}

function getWorkshopPaymentLinkId() {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('WORKSHOP_PAYMENT_LINK_ID');
  if (cached) return cached;
  try {
    const links = asaasGet('/paymentLinks?limit=100&active=true').data || [];
    for (let i = 0; i < links.length; i++) {
      const link = links[i];
      const text = [link.name, link.description, link.url].join(' ').toLowerCase();
      const slugMatch = text.indexOf(WORKSHOP_PUBLIC_LINK_SLUG.toLowerCase()) >= 0;
      const nameMatch = text.indexOf('workshop') >= 0 && text.indexOf('seletividade') >= 0;
      if (slugMatch || nameMatch) {
        props.setProperty('WORKSHOP_PAYMENT_LINK_ID', String(link.id));
        return String(link.id);
      }
    }
  } catch (err) {
    console.error('getWorkshopPaymentLinkId:', err);
  }
  return '';
}

function upsertAlunaWorkshop(payment, customer) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ALUNAS_WORK_SHEET);
  const values = [
    payment.paymentDate || payment.clientPaymentDate || payment.dateCreated || '',
    customer.name || '', normalizeEmail(customer.email), customer.phone || customer.mobilePhone || '',
    payment.id || '', translateBillingType(payment.billingType), translateStatus(payment.status), payment.value || 0,
    payment.description || '', payment.externalReference || '', new Date()
  ];
  upsertById(sheet, 5, payment.id, values);
}

function upsertLeadQuente(payment, customer, situation) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LEADS_QUENTE_SHEET);
  const values = [
    payment.dateCreated || new Date(), customer.name || '', normalizeEmail(customer.email),
    customer.phone || customer.mobilePhone || '', payment.id || '', translateBillingType(payment.billingType),
    translateStatus(payment.status), payment.value || 0, payment.description || '', payment.externalReference || '',
    situation, new Date()
  ];
  upsertById(sheet, 5, payment.id, values);
}

function markLeadAsConverted(payment, customer) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LEADS_QUENTE_SHEET);
  const row = findRowByValue(sheet, 5, payment.id);
  if (row > 1) upsertLeadQuente(payment, customer, 'CONVERTIDO');
}

function buildEventState(sheet) {
  const state = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return state;
  const rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  rows.forEach(function(row) {
    if (row[1]) state[String(row[1])] = [row[3] || '', row[4] || ''].join('|');
  });
  return state;
}

function upsertById(sheet, idColumn, id, values) {
  const row = findRowByValue(sheet, idColumn, id);
  if (row > 1) sheet.getRange(row, 1, 1, values.length).setValues([values]);
  else sheet.appendRow(values);
}

function findRowByValue(sheet, column, value) {
  if (!value || sheet.getLastRow() < 2) return -1;
  const finder = sheet.getRange(2, column, sheet.getLastRow() - 1, 1)
    .createTextFinder(String(value)).matchEntireCell(true).findNext();
  return finder ? finder.getRow() : -1;
}

function ensureAllSheets() {
  ensureSheet(COMU_FREE_SHEET, ['Data / Hora','Nome','Email','Telefone','Origem','Primeiro Evento','Último Evento']);
  ensureSheet(EVENTOS_SHEET, ['Data / Hora Sync','ID Pagamento','Evento Derivado','Status','Tipo Cobrança','Data Criação','Data Pagamento','Nome','Email','Telefone','Descrição','Referência','Valor','Link / Produto']);
  ensureSheet(ALUNAS_WORK_SHEET, ['Data da Compra','Nome','Email','Telefone','ID Pagamento','Tipo Cobrança','Status','Valor','Descrição','Referência','Última Atualização']);
  ensureSheet(LEADS_QUENTE_SHEET, ['Data da Tentativa','Nome','Email','Telefone','ID Pagamento','Tipo Cobrança','Status','Valor','Descrição','Referência','Situação Remarketing','Última Atualização']);
}

function ensureSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0 && headers && headers.length) sheet.appendRow(headers);
  return sheet;
}

function asaasGet(path) {
  const key = getAsaasKey();
  if (!key) throw new Error('ASAAS_API_KEY não configurada');
  const response = UrlFetchApp.fetch(ASAAS_BASE_URL + path, {
    method: 'get',
    headers: { access_token: key },
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  const body = response.getContentText();
  if (code < 200 || code >= 300) throw new Error('Asaas HTTP ' + code + ': ' + body.slice(0, 300));
  return JSON.parse(body);
}

function isPaidStatus(status) {
  return ['CONFIRMED', 'RECEIVED', 'RECEIVED_IN_CASH'].indexOf(String(status || '').toUpperCase()) >= 0;
}

function translateStatus(status) {
  const labels = {
    PENDING: 'Pendente', RECEIVED: 'Recebido', CONFIRMED: 'Confirmado',
    OVERDUE: 'Vencido', REFUNDED: 'Estornado', RECEIVED_IN_CASH: 'Recebido em dinheiro',
    REFUND_REQUESTED: 'Estorno solicitado', REFUND_IN_PROGRESS: 'Estorno em andamento',
    CHARGEBACK_REQUESTED: 'Contestação solicitada', CHARGEBACK_DISPUTE: 'Contestação em análise',
    AWAITING_CHARGEBACK_REVERSAL: 'Aguardando reversão da contestação',
    DUNNING_REQUESTED: 'Negativação solicitada', DUNNING_RECEIVED: 'Negativação recebida',
    AWAITING_RISK_ANALYSIS: 'Aguardando análise de risco', CANCELED: 'Cancelado',
    DELETED: 'Excluído'
  };
  const key = String(status || '').toUpperCase();
  return labels[key] || key || 'Não informado';
}

function translateBillingType(type) {
  const labels = {
    PIX: 'Pix', BOLETO: 'Boleto', CREDIT_CARD: 'Cartão de crédito',
    DEBIT_CARD: 'Cartão de débito', TRANSFER: 'Transferência',
    DEPOSIT: 'Depósito', UNDEFINED: 'A definir'
  };
  const key = String(type || '').toUpperCase();
  return labels[key] || key || 'Não informado';
}

function daysSince(dateValue) {
  const date = new Date(dateValue);
  if (isNaN(date.getTime())) return 0;
  return Math.max(0, Math.floor((new Date().getTime() - date.getTime()) / 86400000));
}

function normalizeEmail(email) {
  return String(email || '').toLowerCase().trim();
}

function refreshDashboard() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Dashboard');
    if (sheet) sheet.getRange('B3').setValue('Atualizado em ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
  } catch (err) {
    console.error('refreshDashboard:', err);
  }
}

function jsonOutput(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
