/**
 * Webhook Receiver for Respondi -> Google Sheets (VERSÃO 16.2.5)
 * --------------------------------------------------------------
 * NOVO: CRM agora é preenchido por cabeçalhos (linha 2), não por número de coluna.
 * Patch no processHandles_ para evitar transformar “Gabriela” em “@Gabriela", por exemplo.
 * - Assinatura gravada na RAW (coluna R = processed_by)
 * REMOVIDO Cálculo de datas de Follow-up (serão usadas pelo usuário para texto referente aos follow-ups)
 * v16.2.0:
 * - RAW passa a gravar "resumo" em Q (col 17) já pronto (sem custom function)
 * - RESUMO_LEAD() agora só lê a coluna Q (resumo) por raw_id (bem mais leve)
 *
 * v16.1.5 base:
 * - Interesse dinâmico por funil
 * - Filtro pro CRM (só entra se tiver contato) + flag na RAW (SEM_CONTATO)
 * - respondent_id incluído na RAW
 *
 * Contato “acionável” = WhatsApp OU Email válido OU Instagram preenchido.
 * - Sempre grava na RAW
 * - Só grava no CRM se tiver contato
 * - Marca RAW col P (flags) com "SEM_CONTATO" quando não tiver contato
 */

const SHEET_RAW = 'RAW_RESPONDI';
const SHEET_MAP = 'MAPEAMENTO_FORMS';
const SHEET_CRM = 'CRM_GERAL';

// Assinatura gravada na RAW (coluna R = processed_by)
const PROCESSED_BY = 'webhook_v16.2.5';

const RAW_COLS = [
  'raw_id', 'received_at', 'form_id', 'form_name',
  'lead_name', 'lead_instagram', 'lead_whatsapp', 'lead_email',
  'lead_key', 'source', 'funnel', 'interest', 'notes', 'payload_json',
  'respondent_id',
  'flags',  // Coluna P (16)
  'resumo', // Coluna Q (17)
  'processed_by' // Coluna R (18)  <-- NOVO
];

/**
 * =========================
 * CRM: Mapeamento por cabeçalho (linha 2)
 * - Chaves normalizadas: minúsculas e sem sufixos em "(...)" ou "[...]"
 *   Ex: "Interesse [Infoproduto/Serviço]" -> "interesse"
 * =========================
 */
function normalizeCrmHeader_(h) {
  let s = String(h || '').trim().toLowerCase();
  s = s.split(' (')[0];
  s = s.split(' [')[0];
  return s;
}

function getCrmCtx_(shCrm) {
  const lastCol = Math.max(1, shCrm.getLastColumn());
  const headers = shCrm.getRange(2, 1, 1, lastCol).getValues()[0];
  const colIndexByName = {};
  for (let c = 0; c < headers.length; c++) {
    const key = normalizeCrmHeader_(headers[c]);
    if (key) colIndexByName[key] = c; // 0-based
  }
  return { shCrm, lastCol, headers, colIndexByName };
}

function setCrmByHeader_(ctx, rowArray, headerName, value) {
  const key = normalizeCrmHeader_(headerName);
  const idx = ctx.colIndexByName[key];
  if (idx == null) return;
  rowArray[idx] = value;
}

function getCrmColIndex1Based_(ctx, headerName) {
  const key = normalizeCrmHeader_(headerName);
  const idx0 = ctx.colIndexByName[key];
  return (idx0 == null) ? null : (idx0 + 1);
}


/**
 * GATILHO DE EDIÇÃO: Sincroniza dados quando você mexe na planilha
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const col = range.getColumn();
  const row = range.getRow();

  if (sheet.getName() !== SHEET_CRM || row < 3) return;

  const ctx = getCrmCtx_(sheet);

  const colLeadKey = getCrmColIndex1Based_(ctx, 'lead_key');
  const colInstagram = getCrmColIndex1Based_(ctx, '@ Instagram');
  const colWhatsapp = getCrmColIndex1Based_(ctx, 'WhatsApp');
  const colEmail = getCrmColIndex1Based_(ctx, 'Email');

  // Se os cabeçalhos não existirem, não faz nada
  if (!colLeadKey || !colInstagram || !colWhatsapp || !colEmail) return;

  // Sincronização de ID quando mexer em IG/WhatsApp/Email
  if ([colInstagram, colWhatsapp, colEmail].includes(col)) {
    const idAtual = String(sheet.getRange(row, colLeadKey).getValue()).trim();
    const instagram = String(sheet.getRange(row, colInstagram).getValue()).trim();
    const whatsapp = String(sheet.getRange(row, colWhatsapp).getValue()).trim();
    const email = String(sheet.getRange(row, colEmail).getValue()).trim();
    if (!instagram && !whatsapp && !email) return;

    const novoIdCalculado = buildLeadKey_(normalizeWhatsapp_(whatsapp), normalizeEmail_(email), instagram, "");
    if (novoIdCalculado && idAtual !== novoIdCalculado) {
      sheet.getRange(row, colLeadKey).setValue(novoIdCalculado);
    }
  }
}


/**
 * FUNÇÃO PERSONALIZADA (agora leve):
 * =RESUMO_LEAD(raw_id)
 *
 * OBS: ideal é não usar mais, mas se ainda existir em algum funil,
 * ela agora só puxa o valor pronto da coluna Q (resumo).
 */
function RESUMO_LEAD(ids) {
  if (!ids) return "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shRaw = ss.getSheetByName(SHEET_RAW);
  if (!shRaw) return "Erro: RAW não encontrada";

  const lastRow = shRaw.getLastRow();
  if (lastRow < 2) return "";

  // Lê só A (raw_id) e Q (resumo)
  const data = shRaw.getRange(2, 1, lastRow - 1, 17).getValues(); // A..Q
  const db = new Map();
  for (let i = 0; i < data.length; i++) {
    const rid = String(data[i][0] || '').trim();     // A
    const resumo = String(data[i][16] || '').trim(); // Q (17) => idx 16
    if (rid) db.set(rid, resumo);
  }

  const one = (id) => {
    const rid = String(id).trim();
    if (!rid || rid === "raw_id") return "";
    return db.get(rid) || "";
  };

  return Array.isArray(ids) ? ids.map(row => [one(row[0])]) : one(ids);
}

/**
 * APOIO: Gera IDs para quem está sem
 */
function gerarIdsFaltantesCrm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_CRM);
  if (!sh) return;

  const last = sh.getLastRow();
  if (last < 3) return;

  const ctx = getCrmCtx_(sh);

  const colLeadKey = getCrmColIndex1Based_(ctx, 'lead_key');
  const colInstagram = getCrmColIndex1Based_(ctx, '@ Instagram');
  const colWhatsapp = getCrmColIndex1Based_(ctx, 'WhatsApp');
  const colEmail = getCrmColIndex1Based_(ctx, 'Email');

  if (!colLeadKey || !colInstagram || !colWhatsapp || !colEmail) return;

  const range = sh.getRange(3, 1, last - 2, ctx.lastCol);
  const data = range.getValues();

  const novos = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    let id = row[colLeadKey - 1];
    if (!id) {
      const whatsapp = row[colWhatsapp - 1];
      const email = row[colEmail - 1];
      const instagram = row[colInstagram - 1];
      id = buildLeadKey_(normalizeWhatsapp_(whatsapp), normalizeEmail_(email), String(instagram || ''), "");
    }
    novos.push([id]);
  }

  // Atualiza somente a coluna lead_key
  sh.getRange(3, colLeadKey, novos.length, 1).setValues(novos);
}


/* --- CORE WEBHOOK --- */

function doGet() {
  return ContentService
    .createTextOutput('Webhook V16.2.3 Online')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return json_(429, { ok: false, error: 'Ocupado' });

  try {
    const token = (e?.parameter?.token) ? String(e.parameter.token) : '';
    const expected = getProp_('WEBHOOK_TOKEN');
    if (expected && token !== expected) return json_(401, { ok: false, error: 'Não autorizado' });

    const parsed = parsePayload_(e);
    const payload = parsed.payload, payloadJson = parsed.payloadJson;

    // Permite backfill/histórico (importação): usa data do payload se existir.
    const receivedAt = parseIncomingDate_(payload) || new Date();
    const formId = String(payload?.form?.form_id || ''), formName = String(payload?.form?.form_name || '');
    const map = loadFormMap_(formId, formName);

    // Extração robusta
    const rawData = extractFromRespondi_(payload, map);

    // Tratamento de Instagrams extras
    const igProc = processHandles_(rawData.instagram);
    let finalInstagram = igProc.primary, extraIgs = igProc.secondaries;
    const nameProc = processHandles_(rawData.name);

    if (!finalInstagram && nameProc.primary) {
      finalInstagram = nameProc.primary;
      extraIgs = extraIgs.concat(nameProc.secondaries);
    } else if (nameProc.primary) {
      extraIgs = extraIgs.concat(nameProc.primary).concat(nameProc.secondaries);
    }

    const finalName = cleanNameWithHandles_(rawData.name, nameProc.matchesRaw.concat(igProc.matchesRaw));
    let finalNotes = rawData.notes;
    const uniqueExtras = [...new Set(extraIgs.map(i => String(i).toLowerCase()))].filter(i => i && i !== String(finalInstagram || '').toLowerCase());
    if (uniqueExtras.length > 0) finalNotes += (finalNotes ? ' | Outros: ' : 'Outros: ') + uniqueExtras.join(', ');

    const leadEmail = normalizeEmail_(rawData.email);
    const leadWhatsapp = normalizeWhatsapp_(rawData.whatsapp);

    // Observação: finalInstagram pode ser email (por isso o isEmail_)
    const leadKey = buildLeadKey_(
      leadWhatsapp,
      (isEmail_(finalInstagram) ? finalInstagram : leadEmail),
      finalInstagram,
      String(payload?.respondent?.respondent_id || '')
    );

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const finalSource = validarContraLista_(ss, map.source_default || pickFirst_(payload, ['source', 'origem']) || 'Formulário', 'Origem_Lead', 'Formulário');
    const finalFunnel = validarContraLista_(ss, map.funnel_default || formName, 'Funis_Entrada', '');

    // === TRATAMENTO DE ESTADO: Recuperação vs Completo ===
    const respStatus = String(payload?.respondent?.status || '').toLowerCase();
    const isRecuperacao = (respStatus === 'incomplete' || respStatus === 'recovered' || respStatus === 'abandoned');

    // Rótulo explícito para a operação
    const statusLabel = isRecuperacao ? 'Recuperação - Formulário Abandonado' : 'Novo Lead - Formulário';

    // Validação contra a aba CONFIG
    const finalStatus = validarContraLista_(ss, statusLabel, 'Status_Funil', 'Novo Lead - Formulário');

    const rawId = Utilities.getUuid();

    // === Interesse dinâmico por funil ===
    const rawInterest = String(map.interest_default || rawData.interest || '').trim();
    const interestCfg = getInterestConfigByFunnel_(finalFunnel);
    const crmInterest = normalizeToListOrFallback_(ss, rawInterest, interestCfg.rangeName, interestCfg.fallback);

    // Se precisou cair no fallback, guarda o valor original nas notas
    if (rawInterest && crmInterest !== rawInterest) {
      finalNotes += (finalNotes ? ' | ' : '') + 'Interesse (original): ' + rawInterest;
    }

    // === Filtro CRM por contato ===
    const hasContact = hasAnyContact_(finalInstagram, leadWhatsapp, leadEmail);
    const flags = hasContact ? '' : 'SEM_CONTATO';

    // === NOVO: resumo pronto gravado na RAW ===
    const resumo = buildResumoFromPayloadJson_(payloadJson);

    // 1) RAW sempre
    const shRaw = ensureSheet_(ss, SHEET_RAW, RAW_COLS);
    shRaw.appendRow([
      rawId, receivedAt, formId, formName,
      finalName, finalInstagram, leadWhatsapp, leadEmail,
      leadKey, finalSource, finalFunnel,
      rawInterest,
      finalNotes,
      payloadJson,
      String(payload?.respondent?.respondent_id || ''),
      flags,
      resumo,
      PROCESSED_BY
    ]);

    // 2) CRM só se tiver contato (dedupe por respondent_id dentro de inserirNoCrmGeral_)
    const insertedCrm = hasContact ? inserirNoCrmGeral_(ss, {
      leadKey,
      rawId,
      respondentId: String(payload?.respondent?.respondent_id || ''),
      finalName,
      finalInstagram,
      leadWhatsapp,
      leadEmail,
      notes: finalNotes,
      source: finalSource,
      funnel: finalFunnel,
      interest: crmInterest,
      status: finalStatus,
      receivedAt
    }) : false;

    return json_(200, { ok: true, raw_id: rawId, inserted_crm: insertedCrm, flags });

  } catch (err) {
    console.error("ERRO:", err);
    return json_(500, { ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function inserirNoCrmGeral_(ss, d) {
  const sh = ss.getSheetByName(SHEET_CRM);
  if (!sh) return false;

  const ctx = getCrmCtx_(sh);

  const respondentId = String(d.respondentId || '').trim();

  // Dedupe: respondent_id (por cabeçalho)
  const colRespondentId = getCrmColIndex1Based_(ctx, 'respondent_id');
  if (respondentId && colRespondentId) {
    const last = sh.getLastRow();
    if (last >= 3) {
      const existing = sh.getRange(3, colRespondentId, last - 2, 1).getValues().flat()
        .map(v => String(v || '').trim())
        .filter(Boolean);
      if (existing.includes(respondentId)) return false;
    }
  }

  const baseDate = new Date();

  // Cria uma linha com o mesmo tamanho do CRM (última coluna)
  const row = new Array(ctx.lastCol).fill("");

  // Campos básicos
  setCrmByHeader_(ctx, row, 'lead_key', d.leadKey || "");
  setCrmByHeader_(ctx, row, 'raw_id', d.rawId || "");
  setCrmByHeader_(ctx, row, 'respondent_id', respondentId);

  setCrmByHeader_(ctx, row, 'Ação / Tarefa', "");
  setCrmByHeader_(ctx, row, 'Status_Funil', d.status || "");
  setCrmByHeader_(ctx, row, 'Tipo_Relacao', 'Lead');
  setCrmByHeader_(ctx, row, 'Temperatura', 'Quente');

  setCrmByHeader_(ctx, row, 'Nome', d.finalName || "");
  setCrmByHeader_(ctx, row, '@ Instagram', d.finalInstagram || "");
  setCrmByHeader_(ctx, row, 'WhatsApp', d.leadWhatsapp || "");
  setCrmByHeader_(ctx, row, 'Email', d.leadEmail || "");
  setCrmByHeader_(ctx, row, 'Anotações / Outros IGs', d.notes || "");

  setCrmByHeader_(ctx, row, 'Origem_Lead', d.source || "");
  setCrmByHeader_(ctx, row, 'Funil_Entrada', d.funnel || "");
  setCrmByHeader_(ctx, row, 'Interesse [Infoproduto/Serviço]', d.interest || "");
  setCrmByHeader_(ctx, row, 'Data_Entrada', baseDate);

  sh.insertRowBefore(3);
  const rg = sh.getRange(3, 1, 1, ctx.lastCol);
  rg.setValues([row]);
  rg.setVerticalAlignment("middle");

  // Formata Data_Entrada, se existir
  const colDataEntrada = getCrmColIndex1Based_(ctx, 'Data_Entrada');
  if (colDataEntrada) {
    sh.getRange(3, colDataEntrada).setNumberFormat("dd/mm/yyyy HH:mm");
  }

  return true;
}


/* --- EXTRAÇÃO --- */

function extractFromRespondi_(payload, map) {
  const out = { name: '', instagram: '', whatsapp: '', email: '', interest: '', notes: '' };

  const respondent = payload.respondent || {};
  const rawArr = Array.isArray(respondent.raw_answers) ? respondent.raw_answers : [];
  const answersObj = (respondent.answers && typeof respondent.answers === 'object') ? respondent.answers : {};

  const getQ = (ks) => {
    if (!ks || !String(ks).trim()) return null;
    const list = String(ks).split('|').map(s => s.trim().toLowerCase()).filter(Boolean);
    return rawArr.find(r => {
      const t = String(r?.question?.question_title || '').toLowerCase();
      return list.some(k => t.includes(k));
    });
  };

  // 1) Tipos nativos
  for (const ra of rawArr) {
    const qt = String(ra?.question?.question_type || '').toLowerCase();
    if (qt === 'name' && !out.name) out.name = valueFromRawAnswer_(ra);
    if (qt === 'phone' && !out.whatsapp) out.whatsapp = phoneFromRawAnswer_(ra);
    if (qt === 'email' && !out.email) out.email = valueFromRawAnswer_(ra);

    if ((qt === 'instagram' || qt === 'social' || qt === 'url') && !out.instagram) {
      out.instagram = valueFromRawAnswer_(ra);
    }
  }

  // 2) Mapeamento por keys
  if (!out.name) out.name = valueFromRawAnswer_(getQ(map.field_name_keys));
  if (!out.instagram) out.instagram = valueFromRawAnswer_(getQ(map.field_instagram_keys));
  if (!out.whatsapp) out.whatsapp = phoneFromRawAnswer_(getQ(map.field_whatsapp_keys));
  if (!out.email) out.email = valueFromRawAnswer_(getQ(map.field_email_keys));
  if (!out.interest) out.interest = valueFromRawAnswer_(getQ(map.field_interest_keys));
  if (!out.notes) out.notes = valueFromRawAnswer_(getQ(map.field_notes_keys));

  // 3) Heurística pelo answersObj
  const pickFromAnswersObj_ = (obj, re) => {
    for (const k in obj) if (re.test(k)) return String(obj[k]);
    return '';
  };

  if (!out.name) out.name = pickFromAnswersObj_(answersObj, /(nome|name|completo)/i);
  if (!out.whatsapp) out.whatsapp = normalizeWhatsapp_(pickFromAnswersObj_(answersObj, /(whats|telefone|celular)/i));
  if (!out.email) out.email = pickFromAnswersObj_(answersObj, /(email|e-mail)/i);

  // Instagram: chave que tenha instagram/insta/ig e NÃO tenha nome
  if (!out.instagram) {
    const igOnlyKey = Object.keys(answersObj).find(k => /instagram|insta|ig/i.test(k) && !/nome|name/i.test(k));
    if (igOnlyKey) out.instagram = String(answersObj[igOnlyKey] || '');
  }

  // 4) Interesse: fallback pela primeira pergunta tipo radio
  if (!out.interest) {
    const radio = rawArr.find(r => String(r?.question?.question_type || '').toLowerCase() === 'radio');
    if (radio) out.interest = valueFromRawAnswer_(radio);
  }

  return out;
}

/* --- RESUMO (mesma lógica do backfill, sem ler RAW inteira) --- */

function buildResumoFromPayloadJson_(payloadJson) {
  let p;
  try {
    p = JSON.parse(payloadJson);
  } catch (_) {
    return '❌ Erro JSON';
  }

  const resp = p?.respondent || {};
  let ans = [];

  // 1) raw_answers (webhook)
  if (Array.isArray(resp.raw_answers) && resp.raw_answers.length) {
    ans = resp.raw_answers.map(r => ({
      q: String(r?.question?.question_title || '').toLowerCase(),
      a: Array.isArray(r?.answer) ? r.answer.join(', ') : String(r?.answer ?? '')
    }));
  }
  // 2) answers (importados)
  else if (resp.answers && typeof resp.answers === 'object') {
    const IGNORAR = new Set(["pontuação", "pontuacao", "data", "id", "score"]);
    ans = Object.entries(resp.answers)
      .filter(([k]) => !IGNORAR.has(String(k).trim().toLowerCase()))
      .map(([k, v]) => ({
        q: String(k || '').toLowerCase(),
        a: Array.isArray(v) ? v.join(', ') : String(v ?? '')
      }));
  }

  const buscar = (ks) => {
    const m = ans.find(it => ks.some(k => it.q.includes(k)));
    return m ? String(m.a).trim() : '';
  };

  const desafio = buscar(["maior desafio", "desafio", "dificuldade", "dor", "problema"]);
  const objetivo = buscar(["maior objetivo", "o seu maior objetivo", "vale a pena", "valeu a pena", "final do consumo", "precisa acontecer"]);

  const fase = buscar(["estágio da sua carreira", "estágio da carreira", "fase atual", "em que fase"]);
  const faseFallback = buscar(["nível", "momento"]);
  const inv = buscar(["investimento", "você pode realizar", "pode investir", "parcelado", "à vista", "cartão", "pix"]);
  const faseFinal = fase || (!inv ? faseFallback : "");

  const preco = buscar(["quanto você cobra", "quanto cobra", "ticket", "valor", "preço", "precificar"]);
  const clientes = buscar(["clientes ativos", "quantos clientes", "número de clientes", "clientes hoje"]);
  const servicos = buscar(["quais serviços", "serviços você entrega", "o que você entrega", "entrega dentro do seu trabalho"]);

  const faturamento = buscar(["faturamento", "ganha", "renda", "r$"]);
  const escala = buscar(["estágio", "equipe", "volume"]);

  let res = "";
  if (faseFinal) res += "📍 Fase: " + faseFinal + "\n";
  if (inv) res += "💳 Investimento: " + inv + "\n";
  if (desafio) res += "⚠️ Desafio: " + desafio + "\n";
  if (objetivo) res += "🎯 Objetivo: " + objetivo + "\n";
  if (preco) res += "💵 Preço: " + preco + "\n";
  if (clientes) res += "👥 Clientes: " + clientes + "\n";
  if (servicos) res += "🧩 Serviços: " + servicos + "\n";

  if (!res.trim()) {
    if (faturamento) res += "💰 " + faturamento + "\n";
    if (escala) res += "🚀 " + escala + "\n";
  }

  return res.trim();
}

/* --- HELPERS --- */

function hasAnyContact_(instagram, whatsapp, email) {
  const w = normalizeWhatsapp_(whatsapp);
  const e = String(email || '').trim().toLowerCase();
  const i = String(instagram || '').trim();
  return (w && w.length >= 10) || isEmail_(e) || !!i;
}

function getInterestConfigByFunnel_(funnel) {
  const f = String(funnel || '').trim().toLowerCase();
  if (f === 'serviços agência' || f === 'servicos agencia') {
    return { rangeName: 'Servicos_Agencia', fallback: 'Serviços Diversos' };
  }
  return { rangeName: 'Infoprodutos', fallback: 'Infoprodutos [precisa analisar o lead]' };
}

function buildLeadKey_(w, e, i, r) {
  if (w && w.length > 5) return 'w:' + w;
  if (e && e.includes('@')) return 'e:' + e.toLowerCase().trim();
  if (i && i.trim()) return 'i:' + String(i).toLowerCase().replace('@', '').trim();
  return r ? 'r:' + r : '';
}
function normalizeWhatsapp_(v) { return String(v || '').replace(/\D/g, ''); }
function normalizeEmail_(v) { const s = String(v || '').trim().toLowerCase(); return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s) ? s : ''; }
function valueFromRawAnswer_(ra) { if (!ra || ra.answer == null) return ''; return Array.isArray(ra.answer) ? ra.answer.join(', ') : String(ra.answer); }
function phoneFromRawAnswer_(ra) {
  if (!ra) return '';
  const a = ra.answer;
  if (a && typeof a === 'object' && a.country && a.phone) return String(a.country).replace(/\D/g, '') + String(a.phone).replace(/\D/g, '');
  return normalizeWhatsapp_(valueFromRawAnswer_(ra));
}
function isEmail_(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || '').trim().toLowerCase()); }
function pickFirst_(p, ks) { for (const k of ks) if (p && k in p) return p[k]; return ''; }
function getProp_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function json_(c, d) { return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON); }

function ensureSheet_(ss, n, h) {
  let s = ss.getSheetByName(n);
  if (!s) s = ss.insertSheet(n);

  const current = s.getRange(1, 1, 1, Math.max(1, s.getLastColumn())).getValues()[0].map(x => String(x || '').trim());
  const needsHeader = current.length < h.length || current[0] !== h[0] || current[h.length - 1] !== h[h.length - 1];

  if (needsHeader) {
    s.getRange(1, 1, 1, h.length).setValues([h]);
    s.setFrozenRows(1);
  }
  return s;
}

function loadFormMap_(fid, fname) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(SHEET_MAP);
  if (!sh) return {};
  const data = sh.getDataRange().getValues(), h = data[0].map(s => String(s).trim()), idx = (c) => h.indexOf(c);
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if ((fid && String(r[idx('form_id')]).trim() === fid) || (!fid && String(r[idx('form_name')]).trim() === fname)) {
      const g = k => (idx(k) >= 0 ? String(r[idx(k)]).trim() : '');
      return {
        funnel_default: g('funnel_default'),
        interest_default: g('interest_default'),
        source_default: g('source_default'),
        field_name_keys: g('field_name_keys'),
        field_instagram_keys: g('field_instagram_keys'),
        field_whatsapp_keys: g('field_whatsapp_keys'),
        field_email_keys: g('field_email_keys'),
        field_interest_keys: g('field_interest_keys'),
        field_notes_keys: g('field_notes_keys')
      };
    }
  }
  return {};
}

function validarContraLista_(ss, v, rangeName, fallback) {
  if (v == null) return fallback;
  const raw = String(v).trim();
  if (!raw) return fallback;

  try {
    const r = ss.getRangeByName(rangeName);
    if (!r) return raw;

    const vs = r.getValues().flat().map(x => String(x).trim()).filter(Boolean);
    const rawLower = raw.toLowerCase();

    const hit = vs.find(x => x.toLowerCase() === rawLower);
    return hit ? hit : raw;
  } catch (e) {
    return raw;
  }
}

function normalizeToListOrFallback_(ss, v, rangeName, fallback) {
  const raw = String(v || '').trim();
  if (!raw) return fallback;
  try {
    const r = ss.getRangeByName(rangeName);
    if (!r) return fallback;
    const vs = r.getValues().flat().map(x => String(x).trim()).filter(Boolean);
    const rawLower = raw.toLowerCase();
    const hit = vs.find(x => x.toLowerCase() === rawLower);
    return hit ? hit : fallback;
  } catch (e) {
    return fallback;
  }
}

function processHandles_(text) {
  const regex = /([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})|@\s*([a-zA-Z0-9._]{2,30})/gi;
  const matchesRaw = [], normalized = [];
  const str = String(text || '');
  for (const m of str.matchAll(regex)) {
    matchesRaw.push(m[0]);
    if (m[1]) normalized.push(m[1].toLowerCase()); else if (m[2]) normalized.push('@' + m[2]);
  }
  // Fallback SEM "@": só considera handle se realmente "parecer" um user.
  // Evita transformar nomes simples (ex: "Gabriela") em @gabriela.
  if (normalized.length === 0) {
    const s = str.trim();
    if (s.length > 2 && !s.includes(' ')) {
      const clean = s.replace(/^@+/, '').trim();
      const hasDotOrUnderscore = /[._]/.test(clean);
      const hasDigit = /\d/.test(clean);
      const isAllLower = clean === clean.toLowerCase();
      const isOnlyLetters = /^[A-Za-zÀ-ÿ]+$/.test(clean);
      const looksLikeHandle = (hasDotOrUnderscore || hasDigit || isAllLower) && !(isOnlyLetters && !isAllLower);

      if (looksLikeHandle) {
        normalized.push('@' + clean.toLowerCase());
        matchesRaw.push(str);
      }
    }
  }

  return { primary: normalized[0] || '', secondaries: normalized.slice(1), matchesRaw };
}

function cleanNameWithHandles_(text, handles) {
  let s = String(text || '');
  for (const h of handles || []) s = s.replace(new RegExp(h.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), '');
  return s.replace(/[|/\\•·]/g, ' ').replace(/\s{2,}/g, ' ').trim();
}

function parsePayload_(e) {
  let p = {}, pj = '';
  const c = e?.postData?.contents;
  if (c) {
    try { p = JSON.parse(c); pj = JSON.stringify(p); }
    catch (_) { p = { raw: c }; }
  } else if (e?.parameter) {
    p = e.parameter; pj = JSON.stringify(p);
  }
  return { payload: p, payloadJson: pj };
}

function parseIncomingDate_(payload) {
  const candidates = [
    payload?.received_at,
    payload?.responded_at,
    payload?.respondent?.responded_at,
    payload?.respondent?.created_at,
    payload?.respondent?.submitted_at,
  ].filter(Boolean);

  for (const c of candidates) {
    const d = (c instanceof Date) ? c : new Date(String(c));
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}
