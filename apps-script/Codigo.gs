const SHEET_ID = '1Bw9RxJ872eb4-f3hciloXyb1wZmSTKPs9Bw_LoYeurQ';
const SHEET_NAME = 'CRM';
const COMU_FREE_SHEET = 'COMU FREE';
const EVENTOS_SHEET = 'EVENTOS';
const ALUNAS_WORK_SHEET = 'ALUNAS WORK';
const LEADS_QUENTE_SHEET = 'LEADS QUENTE';
const ALUNAS_SELET_SHEET = 'ALUNAS SELET';
const LEADS_SELET_SHEET = 'LEADS SELET';
const ADMIN_AULAS_SHEET = 'ADMIN AULAS';
const ADMIN_COMENTARIOS_SHEET = 'ADMIN COMENTARIOS';
const ADMIN_EMAILS = ['nutri4nutri@gmail.com', 'divarebel.on@gmail.com'];
const PAGAMENTOS_SHEET = 'Pagamentos'; // compatibilidade com compras antigas
const ASAAS_BASE_URL = 'https://api.asaas.com/v3';
const WORKSHOP_PUBLIC_LINK_SLUG = 'lreonttfy8mnzycj';
const SELETIVIDADE_PUBLIC_LINK_SLUG = '0gbq24ep6hqvqsh9';

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
  if (action === 'communityContent') return jsonOutput(getCommunityContent(e.parameter.email || ''));
  if (action === 'lessonComments') return jsonOutput(getLessonComments(e.parameter.lessonId || ''));
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

    if (data.action === 'requestAdminCode') return jsonOutput(requestAdminCode(data.email));
    if (data.action === 'verifyAdminCode') return jsonOutput(verifyAdminCode(data.email, data.code));
    if (data.action === 'listAdminData') return jsonOutput(listAdminData(data.token));
    if (data.action === 'saveLesson') return jsonOutput(saveAdminLesson(data.token, data.lesson || {}));
    if (data.action === 'archiveLesson') return jsonOutput(archiveAdminLesson(data.token, data.id));
    if (data.action === 'replyComment') return jsonOutput(replyAdminComment(data.token, data.id, data.reply));
    if (data.action === 'uploadPdf') return jsonOutput(uploadAdminPdf(data.token, data));
    if (data.action === 'addComment') return jsonOutput(addLessonComment(data));

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
  if (curso === 'workshop') return verificarProdutoAPI(email, 'workshop');
  if (curso === 'seletividade') return verificarProdutoAPI(email, 'seletividade');

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
  return verificarProdutoAPI(email, 'workshop');
}

function verificarProdutoAPI(email, produto) {
  try {
    const customers = asaasGet('/customers?email=' + encodeURIComponent(email)).data || [];
    for (let i = 0; i < customers.length; i++) {
      const payments = asaasGet('/payments?limit=100&customer=' + encodeURIComponent(customers[i].id)).data || [];
      for (let j = 0; j < payments.length; j++) {
        const payment = payments[j];
        if (isProductPayment(payment, produto) && isPaidStatus(payment.status)) {
          return {
            acesso: true,
            diasDesdeCompra: daysSince(payment.paymentDate || payment.clientPaymentDate || payment.dateCreated)
          };
        }
      }
    }
  } catch (err) {
    console.error('verificarProdutoAPI:', produto, err);
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
      const produto = identifyPaymentProduct(payment);
      if (!produto) return;

      const customer = getCustomer(payment.customer, customerCache);
      const statusPt = translateStatus(payment.status);
      const billingTypePt = translateBillingType(payment.billingType);
      const signature = [statusPt, billingTypePt].join('|');
      const previous = eventState[payment.id];

      if (previous !== signature) {
        const eventName = deriveEventName(payment, previous);
        eventSheet.appendRow([
          new Date(), payment.id || '', eventName, statusPt, billingTypePt,
          payment.dateCreated || '', payment.paymentDate || payment.clientPaymentDate || '',
          customer.name || '', customer.email || '', customer.phone || customer.mobilePhone || '',
          payment.description || '', payment.externalReference || '', payment.value || 0,
          productLabel(produto)
        ]);
        eventState[payment.id] = signature;
      }

      if (isPaidStatus(payment.status)) {
        upsertAlunaProduto(payment, customer, produto);
        markLeadAsConverted(payment, customer, produto);
      } else {
        upsertLeadProduto(payment, customer, produto, 'NÃO CONVERTIDO');
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

function isSeletividadePayment(payment) {
  const linkId = getSeletividadePaymentLinkId();
  const paymentLink = payment.paymentLink || payment.paymentLinkId || '';
  if (linkId && paymentLink && String(paymentLink) === String(linkId)) return true;
  const text = [payment.description, payment.externalReference].join(' ').toLowerCase();
  return text.indexOf('seletividade') >= 0 && text.indexOf('workshop') < 0;
}

function isProductPayment(payment, produto) {
  return produto === 'workshop' ? isWorkshopPayment(payment) : isSeletividadePayment(payment);
}

function identifyPaymentProduct(payment) {
  if (isWorkshopPayment(payment)) return 'workshop';
  if (isSeletividadePayment(payment)) return 'seletividade';
  return '';
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

function getSeletividadePaymentLinkId() {
  return getPaymentLinkId('SELETIVIDADE_PAYMENT_LINK_ID', SELETIVIDADE_PUBLIC_LINK_SLUG, function(text) {
    return text.indexOf('seletividade') >= 0 && text.indexOf('workshop') < 0;
  });
}

function testarIntegracaoSeletividade() {
  const linkId = getSeletividadePaymentLinkId();
  const payments = listAsaasPayments();
  const matches = payments.filter(isSeletividadePayment).length;
  const result = { linkId: linkId, linkEncontrado: !!linkId, pagamentosRelacionados: matches };
  console.log(JSON.stringify(result));
  return result;
}

function getPaymentLinkId(propertyName, publicSlug, nameMatcher) {
  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty(propertyName);
  if (cached) return cached;
  try {
    const links = asaasGet('/paymentLinks?limit=100&active=true').data || [];
    for (let i = 0; i < links.length; i++) {
      const link = links[i];
      const text = [link.name, link.description, link.url].join(' ').toLowerCase();
      if (text.indexOf(String(publicSlug).toLowerCase()) >= 0 || nameMatcher(text)) {
        props.setProperty(propertyName, String(link.id));
        return String(link.id);
      }
    }
  } catch (err) { console.error('getPaymentLinkId:', propertyName, err); }
  return '';
}

function productLabel(produto) {
  return produto === 'workshop' ? 'Workshop Seletividade Alimentar' : 'Curso Seletividade Alimentar';
}

function deriveEventName(payment, previous) {
  const status = String(payment.status || '').toUpperCase();
  if (isPaidStatus(status)) return 'Pagamento confirmado';
  if (status === 'OVERDUE') return 'Pagamento vencido';
  if (status === 'REFUNDED') return 'Pagamento estornado';
  if (status === 'CANCELED' || status === 'DELETED') return 'Cobrança cancelada';
  if (status === 'PENDING') {
    if (String(payment.billingType || '').toUpperCase() === 'BOLETO') return 'Boleto emitido';
    if (String(payment.billingType || '').toUpperCase() === 'PIX') return 'Pix gerado';
    return 'Tentativa de compra / pagamento pendente';
  }
  return previous ? 'Status atualizado' : 'Cobrança identificada';
}

function upsertAlunaWorkshop(payment, customer) {
  return upsertAlunaProduto(payment, customer, 'workshop');
}

function upsertAlunaProduto(payment, customer, produto) {
  const sheetName = produto === 'workshop' ? ALUNAS_WORK_SHEET : ALUNAS_SELET_SHEET;
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
  const values = [
    payment.paymentDate || payment.clientPaymentDate || payment.dateCreated || '',
    customer.name || '', normalizeEmail(customer.email), customer.phone || customer.mobilePhone || '',
    payment.id || '', translateBillingType(payment.billingType), translateStatus(payment.status), payment.value || 0,
    payment.description || '', payment.externalReference || '', new Date()
  ];
  upsertById(sheet, 5, payment.id, values);
}

function upsertLeadQuente(payment, customer, situation) {
  return upsertLeadProduto(payment, customer, 'workshop', situation);
}

function upsertLeadProduto(payment, customer, produto, situation) {
  const sheetName = produto === 'workshop' ? LEADS_QUENTE_SHEET : LEADS_SELET_SHEET;
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
  const values = [
    payment.dateCreated || new Date(), customer.name || '', normalizeEmail(customer.email),
    customer.phone || customer.mobilePhone || '', payment.id || '', translateBillingType(payment.billingType),
    translateStatus(payment.status), payment.value || 0, payment.description || '', payment.externalReference || '',
    situation, new Date()
  ];
  upsertById(sheet, 5, payment.id, values);
}

function markLeadAsConverted(payment, customer, produto) {
  const sheetName = produto === 'seletividade' ? LEADS_SELET_SHEET : LEADS_QUENTE_SHEET;
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);
  const row = findRowByValue(sheet, 5, payment.id);
  if (row > 1) upsertLeadProduto(payment, customer, produto || 'workshop', 'CONVERTIDO');
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
  ensureSheet(ALUNAS_SELET_SHEET, ['Data da Compra','Nome','Email','Telefone','ID Pagamento','Tipo Cobrança','Status','Valor','Descrição','Referência','Última Atualização']);
  ensureSheet(LEADS_SELET_SHEET, ['Data da Tentativa','Nome','Email','Telefone','ID Pagamento','Tipo Cobrança','Status','Valor','Descrição','Referência','Situação Remarketing','Última Atualização']);
  ensureAdminSheets();
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

/* ─── COMUNIDADE E PAINEL ADMINISTRATIVO ─── */
function ensureAdminSheets() {
  const lessons = ensureSheet(ADMIN_AULAS_SHEET, ['ID','Módulo','Título','Duração','Tipo','Conteúdo HTML','URL Material','Público','Email Específico','Ordem','Status','Criado Em','Atualizado Em','Atualizado Por']);
  ensureSheet(ADMIN_COMENTARIOS_SHEET, ['ID','Aula ID','Data','Nome','Email','Comentário','Resposta','Respondido Em','Respondido Por','Status']);
  if (lessons.getLastRow() < 2) {
    const now = new Date();
    lessons.getRange(2, 1, 2, 14).setValues([
      ['aula-checklist-terapia','Material de Apoio','Checklist Terapia Alimentar — PDF','PDF','PDF','Baixe agora o <strong>Checklist de Terapia Alimentar</strong>, um material prático para apoiar sua conduta clínica.','/materiais/checklist_terapia_alimentar.pdf','TODAS','',1,'ATIVA',now,now,'sistema'],
      ['aula-checklist-anamnese','Material de Apoio','Checklist de Anamnese e Raciocínio Clínico — PDF','PDF','PDF','Baixe o <strong>Checklist: Anamnese para Organizar o Raciocínio Clínico</strong>, com perguntas essenciais e sinais de alerta.','/materiais/checklist_anamnese_raciocinio_clinico.pdf','TODAS','',2,'ATIVA',now,now,'sistema']
    ]);
  }
  [lessons, SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_COMENTARIOS_SHEET)].forEach(function(sheet) {
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,sheet.getLastColumn()).setBackground('#0D0C0A').setFontColor('#C7A16A').setFontWeight('bold');
  });
}

function getCommunityContent(email) {
  ensureAdminSheets();
  email = normalizeEmail(email);
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_AULAS_SHEET);
  const rows = sheet.getLastRow() < 2 ? [] : sheet.getRange(2,1,sheet.getLastRow()-1,14).getValues();
  const lessons = rows.filter(function(r) {
    const audience = String(r[7] || 'TODAS').toUpperCase();
    return String(r[10]).toUpperCase() === 'ATIVA' && (audience === 'TODAS' || (audience === 'ESPECÍFICA' && normalizeEmail(r[8]) === email));
  }).sort(function(a,b){ return Number(a[9]||0)-Number(b[9]||0); }).map(function(r){
    return {id:String(r[0]),module:String(r[1]),name:String(r[2]),dur:String(r[3]||''),type:String(r[4]||''),body:String(r[5]||''),materialUrl:String(r[6]||'')};
  });
  return {ok:true, lessons:lessons};
}

function getLessonComments(lessonId) {
  ensureAdminSheets();
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_COMENTARIOS_SHEET);
  const rows = sheet.getLastRow() < 2 ? [] : sheet.getRange(2,1,sheet.getLastRow()-1,10).getValues();
  return {ok:true, comments:rows.filter(function(r){ return String(r[1]) === String(lessonId) && String(r[9]||'ATIVO').toUpperCase() === 'ATIVO'; }).map(function(r){
    return {id:String(r[0]),date:r[2],name:String(r[3]),text:String(r[5]),reply:String(r[6]||'')};
  })};
}

function addLessonComment(data) {
  ensureAdminSheets();
  const lessonId = cleanText(data.lessonId, 100), name = cleanText(data.name, 100), email = normalizeEmail(data.email), comment = cleanText(data.comment, 1200);
  if (!lessonId || !name || !email || !comment) return {ok:false,error:'Preencha nome, e-mail e comentário.'};
  const id = 'com-' + Utilities.getUuid();
  SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_COMENTARIOS_SHEET).appendRow([id,lessonId,new Date(),name,email,comment,'','','','ATIVO']);
  return {ok:true,id:id};
}

function requestAdminCode(email) {
  email = normalizeEmail(email);
  if (ADMIN_EMAILS.indexOf(email) < 0) return {ok:false,error:'E-mail sem permissão administrativa.'};
  const code = String(Math.floor(100000 + Math.random()*900000));
  const record = {code:code, expiresAt:Date.now()+600000, attempts:0};
  PropertiesService.getScriptProperties().setProperty(adminCodeKey(email), JSON.stringify(record));
  MailApp.sendEmail({
    to:email,
    name:'Priscila Leite',
    replyTo:'nutri4nutri@gmail.com',
    subject:'Seu código de acesso — Nutri For Nutri',
    htmlBody:'<div style="font-family:Arial,sans-serif;color:#17130f"><p>Olá!</p><p>Seu código de acesso ao painel da <strong>Priscila Leite</strong> é:</p><p style="font-size:30px;font-weight:bold;letter-spacing:6px;color:#a47855">'+code+'</p><p>Ele expira em 10 minutos e pode ser usado uma vez.</p><p>Nutri For Nutri · Priscila Leite</p></div>'
  });
  return {ok:true};
}

// Executar uma vez pelo editor para autorizar o serviço de e-mail do projeto.
function autorizarEmailPainel() {
  return MailApp.getRemainingDailyQuota();
}

function verifyAdminCode(email, code) {
  email = normalizeEmail(email);
  if (ADMIN_EMAILS.indexOf(email) < 0 || !code) return {ok:false,error:'Código inválido ou expirado.'};
  const props = PropertiesService.getScriptProperties();
  const key = adminCodeKey(email);
  const raw = props.getProperty(key);
  if (!raw) return {ok:false,error:'Código inválido ou expirado.'};
  let record;
  try { record = JSON.parse(raw); } catch (err) { props.deleteProperty(key); return {ok:false,error:'Código inválido ou expirado.'}; }
  if (Date.now() > Number(record.expiresAt || 0)) { props.deleteProperty(key); return {ok:false,error:'Código expirado. Solicite um novo.'}; }
  if (String(record.code) !== String(code).trim()) {
    record.attempts = Number(record.attempts || 0) + 1;
    if (record.attempts >= 5) props.deleteProperty(key); else props.setProperty(key, JSON.stringify(record));
    return {ok:false,error:'Código incorreto.'};
  }
  props.deleteProperty(key);
  const token = Utilities.getUuid() + Utilities.getUuid();
  props.setProperty('ADMIN_TOKEN_' + token, JSON.stringify({email:email,expiresAt:Date.now()+21600000}));
  return {ok:true,token:token,email:email};
}

function requireAdmin(token) {
  const props = PropertiesService.getScriptProperties();
  const key = token ? 'ADMIN_TOKEN_' + token : '';
  const raw = key && props.getProperty(key);
  let session = null;
  try { session = raw ? JSON.parse(raw) : null; } catch (err) {}
  if (session && Date.now() > Number(session.expiresAt || 0)) { props.deleteProperty(key); session = null; }
  const email = session && normalizeEmail(session.email);
  if (!email || ADMIN_EMAILS.indexOf(email) < 0) throw new Error('Sessão administrativa inválida ou expirada.');
  return email;
}

function adminCodeKey(email) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, normalizeEmail(email));
  return 'ADMIN_CODE_' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/,'');
}

function listAdminData(token) {
  const admin = requireAdmin(token); ensureAdminSheets();
  const ss = SpreadsheetApp.openById(SHEET_ID), ls = ss.getSheetByName(ADMIN_AULAS_SHEET), cs = ss.getSheetByName(ADMIN_COMENTARIOS_SHEET);
  const lessons = ls.getLastRow()<2?[]:ls.getRange(2,1,ls.getLastRow()-1,14).getValues().map(function(r){return {id:r[0],module:r[1],title:r[2],duration:r[3],type:r[4],body:r[5],materialUrl:r[6],audience:r[7],specificEmail:r[8],order:r[9],status:r[10]};});
  const comments = cs.getLastRow()<2?[]:cs.getRange(2,1,cs.getLastRow()-1,10).getValues().map(function(r){return {id:r[0],lessonId:r[1],date:r[2],name:r[3],email:r[4],text:r[5],reply:r[6],status:r[9]};});
  return {ok:true,admin:admin,lessons:lessons,comments:comments};
}

function saveAdminLesson(token, lesson) {
  const admin=requireAdmin(token); ensureAdminSheets();
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_AULAS_SHEET), now=new Date();
  const id=cleanText(lesson.id,100)||('aula-'+Utilities.getUuid()), row=findRowByValue(sheet,1,id);
  const values=[id,cleanText(lesson.module,100)||'Conteúdos',cleanText(lesson.title,180),cleanText(lesson.duration,30),cleanText(lesson.type,30)||'CONTEÚDO',String(lesson.body||'').slice(0,15000),cleanText(lesson.materialUrl,1000),String(lesson.audience||'TODAS').toUpperCase(),normalizeEmail(lesson.specificEmail),Number(lesson.order)||1,'ATIVA',now,now,admin];
  if(!values[2]) return {ok:false,error:'Informe o título da aula.'};
  if(row>1){values[11]=sheet.getRange(row,12).getValue()||now;sheet.getRange(row,1,1,14).setValues([values]);}else sheet.appendRow(values);
  return {ok:true,id:id};
}

function archiveAdminLesson(token,id){const admin=requireAdmin(token),sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_AULAS_SHEET),row=findRowByValue(sheet,1,id);if(row<2)return {ok:false,error:'Aula não encontrada.'};sheet.getRange(row,11).setValue('ARQUIVADA');sheet.getRange(row,13,1,2).setValues([[new Date(),admin]]);return {ok:true};}
function replyAdminComment(token,id,reply){const admin=requireAdmin(token),sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName(ADMIN_COMENTARIOS_SHEET),row=findRowByValue(sheet,1,id);if(row<2)return {ok:false,error:'Comentário não encontrado.'};sheet.getRange(row,7,1,3).setValues([[cleanText(reply,1200),new Date(),admin]]);return {ok:true};}

function uploadAdminPdf(token,data){requireAdmin(token);if(String(data.mimeType)!=='application/pdf')return {ok:false,error:'Envie um arquivo PDF.'};const bytes=Utilities.base64Decode(String(data.base64||''));if(bytes.length>10*1024*1024)return {ok:false,error:'O PDF deve ter no máximo 10 MB.'};const props=PropertiesService.getScriptProperties();let folderId=props.getProperty('COMMUNITY_MATERIALS_FOLDER_ID'),folder;try{folder=folderId?DriveApp.getFolderById(folderId):null;}catch(e){folder=null;}if(!folder){folder=DriveApp.createFolder('Nutri4Nutri - Materiais da Comunidade');props.setProperty('COMMUNITY_MATERIALS_FOLDER_ID',folder.getId());}const file=folder.createFile(Utilities.newBlob(bytes,'application/pdf',cleanText(data.fileName,180)||'material.pdf'));file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);return {ok:true,url:'https://drive.google.com/uc?export=download&id='+file.getId()};}
function cleanText(value,max){return String(value||'').replace(/[<>]/g,'').trim().slice(0,max||500);}
