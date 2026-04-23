/**
 * IMPORTADOR INTELIGENTE DE CSV (RESPONDI) - v3.2.3
 * -------------------------------------------------
 * CRM_GERAL v9:
 * - Cabeçalho na linha 2 (A2..AH2) conforme mapeamento v9.
 * - Dados começam na linha 3.
 * - CRM é mapeado por TEXTO DO CABEÇALHO (não por índice fixo).
* - Normalização de cabeçalhos: ignora sufixos em parênteses OU colchetes.
 * - Dedupe no CRM: SOMENTE por respondent_id (para evitar duplicar a MESMA resposta).
 *   -> Se a pessoa responder de novo, respondent_id muda e entra de novo (o que você quer).
 * - CRM nunca reescreve: só insere novas linhas no topo (linha 3).
 *
 * RAW_RESPONDI v9:
 * - A..R (18 cols) com processed_by na coluna R.
 *
 * Mantido:
 * - Lotes (BATCH_SIZE)
 * - Dedupe RAW por (form_id|respondent_id)
 * - Interesse dinâmico por funil
 * - Notes com "Outros: ..." e "Interesse (original)"
 * - Parsing seguro de @ (não transforma nome simples em @)
 */

/** =========================
 * CONFIG
 * ========================= */
const RESPONDI_IMPORT = {
  BATCH_SIZE: 400,
  STATE_KEY: 'RESPONDI_IMPORT_STATE_V3',
  CACHE_KEYS_PREFIX: 'RESPONDI_KEYS_V3',
  TRIGGER_HANDLER: 'importarCsvRespondiContinue_'
};

/** =========================
 * PROCESSED_BY (somente na RAW)
 * ========================= */
const RESPONDI_IMPORT_VERSION = 'CSV3.2.2';
function buildProcessedByTag_() {
  return `RI:${RESPONDI_IMPORT_VERSION}`;
}

/** =========================
 * MENU
 * ========================= */
function importarCsvRespondiIniciar() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const state = getImportState_();
    if (state && state.status && (state.status === 'RUNNING' || state.status === 'SCHEDULING')) {
      ui.alert('Importação já em andamento', 'Use "Estado da Importação" ou "Cancelar importação".', ui.ButtonSet.OK);
      return;
    }

    const resp = ui.prompt(
      'Importar CSV (Respondi)',
      'Cole a URL do CSV no Drive (link de compartilhamento).',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    const input = String(resp.getResponseText() || '').trim();
    const fileId = extractDriveFileIdFromUrl_(input);
    if (!fileId) {
      ui.alert('Erro', 'Não consegui extrair o File ID da URL. Cole um link do Drive válido.', ui.ButtonSet.OK);
      return;
    }

    const file = getDriveFileByIdSafe_(fileId);
    const fileName = file.getName();
    const formIdFromFileName = fileName.replace(/\.csv$/i, '').trim();

    const csvContent = file.getBlob().getDataAsString('UTF-8');
    const csvData = Utilities.parseCsv(csvContent);
    if (!csvData || csvData.length < 2) throw new Error('O arquivo CSV está vazio ou inválido.');

    const headers = csvData[0];
    const totalRows = csvData.length - 1;

    const config = getFormConfig_(ss, formIdFromFileName, fileName);

    // RAW dedupe cache
    const existingKeys = loadExistingRespondentsForForm_(ss, config.formId);
    cacheExistingKeysForForm_(config.formId, existingKeys);

    const now = new Date();
    setImportState_({
      jobId: Utilities.getUuid(),
      fileId,
      fileName,
      formIdFromFileName,
      config,
      totalRows,
      offset: 0,
      processed: 0,
      duplicatesCount: 0,
      crmCount: 0,
      rawOnlyCount: 0,
      status: 'SCHEDULING',
      cancelled: false,
      startedAt: now.toISOString(),
      updatedAt: now.toISOString(),
      lastError: ''
    });

    clearImportTriggers_();
    scheduleNextBatch_();
    ss.toast('Importação iniciada (lotes).', 'Respondi CSV', 6);

    importarCsvRespondiContinue_();

  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro na Importação', `Detalhes: ${String(e)}`, SpreadsheetApp.getUi().ButtonSet.OK);
    failImportState_(e);
  }
}

function importarCsvRespondiStatus() {
  const ui = SpreadsheetApp.getUi();
  const state = getImportState_();

  if (!state || !state.status) {
    ui.alert('Estado da Importação', 'Nenhuma importação registrada.', ui.ButtonSet.OK);
    return;
  }

  const pct = state.totalRows ? Math.min(100, Math.round((state.offset / state.totalRows) * 100)) : 0;
  ui.alert(
    'Estado da Importação',
    `Status: ${state.status}\n` +
    `Arquivo: ${state.fileName || ''}\n` +
    `Form: ${state.config?.formName || ''}\n` +
    `Total linhas: ${state.totalRows || 0}\n` +
    `Processadas no arquivo: ${state.offset || 0} (${pct}%)\n\n` +
    `Inseridas (não duplicadas): ${state.processed || 0}\n` +
    `Duplicadas puladas (RAW): ${state.duplicatesCount || 0}\n` +
    `✅ CRM: ${state.crmCount || 0}\n` +
    `🟨 Só RAW (SEM_CONTATO): ${state.rawOnlyCount || 0}\n\n` +
    (state.lastError ? `Último erro: ${state.lastError}\n` : '') +
    `Atualizado em: ${state.updatedAt ? new Date(state.updatedAt).toLocaleString() : ''}`,
    ui.ButtonSet.OK
  );
}

function importarCsvRespondiCancelar() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const state = getImportState_();

  if (!state || !state.status || (state.status !== 'RUNNING' && state.status !== 'SCHEDULING')) {
    ui.alert('Cancelar importação', 'Não há importação em andamento.', ui.ButtonSet.OK);
    return;
  }

  state.cancelled = true;
  state.status = 'CANCELLED';
  state.updatedAt = new Date().toISOString();
  setImportState_(state);

  clearImportTriggers_();
  ss.toast('Importação cancelada.', 'Respondi CSV', 6);
  ui.alert('Cancelado', 'Importação cancelada com sucesso.', ui.ButtonSet.OK);
}

/** =========================
 * LOOP EM LOTES
 * ========================= */
function importarCsvRespondiContinue_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const processed_by = buildProcessedByTag_();

  try {
    let state = getImportState_();
    if (!state || !state.status) return;

    if (state.cancelled || state.status === 'CANCELLED') {
      clearImportTriggers_();
      return;
    }
    if (state.status === 'COMPLETED' || state.status === 'FAILED') {
      clearImportTriggers_();
      return;
    }

    state.status = 'RUNNING';
    state.updatedAt = new Date().toISOString();
    setImportState_(state);

    const file = getDriveFileByIdSafe_(state.fileId);
    const csvContent = file.getBlob().getDataAsString('UTF-8');
    const csvData = Utilities.parseCsv(csvContent);
    if (!csvData || csvData.length < 2) throw new Error('O arquivo CSV está vazio ou inválido.');

    const headers = csvData[0];
    const rows = csvData.slice(1);

    state.totalRows = rows.length;

    const map = buildCsvMap_(headers);
    const config = state.config;

    // RAW dedupe
    const existingKeys = getCachedExistingKeysForForm_(config.formId) || loadExistingRespondentsForForm_(ss, config.formId);

    // CRM header map + respondent_id dedupe set
    const crmCtx = getCrmContextV9_(); // { sh, headerRow, startRow, colIndexByName, numCols, respondentColIndex }
    const crmRidSet = buildCrmRespondentIdSet_(crmCtx);

    const start = state.offset || 0;
    const end = Math.min(rows.length, start + RESPONDI_IMPORT.BATCH_SIZE);
    const batchRows = rows.slice(start, end);

    const rawBatch = [];
    const crmBatch = [];

    let processedDelta = 0;
    let duplicatesDelta = 0;
    let crmDelta = 0;
    let rawOnlyDelta = 0;

    for (let idx = 0; idx < batchRows.length; idx++) {
      if (state.cancelled) break;
      const row = batchRows[idx];

      let rawName = map.nome !== -1 ? String(row[map.nome] || '').trim() : '';
      let rawInstagram = map.instagram !== -1 ? String(row[map.instagram] || '').trim() : '';
      let rawWhatsapp = map.whatsapp !== -1 ? normalizeWhatsapp_(row[map.whatsapp]) : '';
      let rawEmail = map.email !== -1 ? String(row[map.email] || '').trim() : '';
      rawEmail = normalizeEmail_(rawEmail);

      let csvDate = map.data !== -1 && row[map.data] ? new Date(row[map.data]) : new Date();
      let respondentId = map.id !== -1 ? String(row[map.id] || '').trim() : '';
      let rawInterestCol = map.interesse !== -1 ? String(row[map.interesse] || '').trim() : '';
      let rawNotesCol = map.notes !== -1 ? String(row[map.notes] || '').trim() : '';

      let extraIgs = [];

      // Nome + IG combinados
      if (map.nomeIg !== -1) {
        const combined = String(row[map.nomeIg] || '').trim();
        const proc = processHandles_(combined);
        if (proc.primary) rawInstagram = proc.primary;
        if (proc.secondaries?.length) extraIgs = extraIgs.concat(proc.secondaries);
        rawName = cleanNameWithHandles_(combined, proc.matchesRaw);
      }

      // Fallback: IG dentro do nome
      if (!rawInstagram && rawName) {
        const procName = processHandles_(rawName);
        if (procName.primary) {
          rawInstagram = procName.primary;
          if (procName.secondaries?.length) extraIgs = extraIgs.concat(procName.secondaries);
          rawName = cleanNameWithHandles_(rawName, procName.matchesRaw);
        }
      }

      if (!rawName && !rawWhatsapp && !rawInstagram && !rawEmail) continue;

      // DEDUPE RAW por (form_id|respondent_id)
      if (respondentId) {
        const k = config.formId + '|' + respondentId;
        if (existingKeys.has(k)) {
          duplicatesDelta++;
          continue;
        }
        existingKeys.add(k);
      }

      processedDelta++;

      const finalFunnel = String(config.funnelDefault || '').trim();

      // Interesse dinâmico por funil
      const rawInterest = String(config.interestDefault || rawInterestCol || '').trim();
      const interestCfg = getInterestConfigByFunnel_(finalFunnel);
      const crmInterest = normalizeToListOrFallback_(ss, rawInterest, interestCfg.rangeName, interestCfg.fallback);

      let finalNotes = String(rawNotesCol || '').trim();
      if (rawInterest && crmInterest !== rawInterest) {
        finalNotes += (finalNotes ? ' | ' : '') + 'Interesse (original): ' + rawInterest;
      }

      // Notes: Outros IGs
      const uniqueExtras = [...new Set((extraIgs || []).map(i => String(i).toLowerCase()))]
        .filter(i => i && i !== String(rawInstagram || '').toLowerCase());
      if (uniqueExtras.length > 0) {
        finalNotes += (finalNotes ? ' | Outros: ' : 'Outros: ') + uniqueExtras.join(', ');
      }

      // Contato acionável
      const hasContact = hasAnyContact_(rawInstagram, rawWhatsapp, rawEmail);
      const flags = hasContact ? '' : 'SEM_CONTATO';
      if (hasContact) crmDelta++; else rawOnlyDelta++;

      // lead_key (IG pode ser email)
      const leadKey = buildLeadKey_(
        rawWhatsapp,
        (isEmail_(rawInstagram) ? rawInstagram : rawEmail),
        rawInstagram,
        respondentId
      );

      // answersObj para payload
      const answersObj = {};
      for (let h = 0; h < headers.length; h++) answersObj[String(headers[h] || '')] = row[h] || '';

      // Se IG não parece handle, tenta extrair de answers
      if (!looksLikeHandle_(String(rawInstagram || '').trim())) {
        const igFromAnswers = extractInstagramFromAnswers_(answersObj);
        if (igFromAnswers) rawInstagram = igFromAnswers;
      }

      // Normaliza IG + extras e limpa nome
      const igProcFinal = processHandles_(String(rawInstagram || '').trim());
      if (igProcFinal.primary) rawInstagram = igProcFinal.primary;
      if (igProcFinal.secondaries?.length) extraIgs = extraIgs.concat(igProcFinal.secondaries);

      const nameProcFinal = processHandles_(String(rawName || '').trim());
      rawName = cleanNameWithHandles_(rawName, (nameProcFinal.matchesRaw || []).concat(igProcFinal.matchesRaw || []));

      const payloadJson = JSON.stringify({
        import: true,
        file: state.fileName,
        form: { form_name: config.formName, form_id: config.formId },
        respondent: { respondent_id: respondentId, responded_at: csvDate, answers: answersObj }
      });

      const resumo = buildResumoFromPayloadJson_IMPORT_(payloadJson);

      // raw_id único (também vai para o CRM)
      const rawId = Utilities.getUuid();

      // RAW A..R (18)
      rawBatch.push([
        rawId,                 // A raw_id
        csvDate,               // B received_at
        config.formId,         // C form_id
        config.formName,       // D form_name
        rawName,               // E lead_name
        rawInstagram,          // F lead_instagram
        rawWhatsapp,           // G lead_whatsapp
        rawEmail,              // H lead_email
        leadKey,               // I lead_key
        config.sourceDefault,  // J source
        finalFunnel,           // K funnel
        rawInterest,           // L interest (cru)
        finalNotes,            // M notes
        payloadJson,           // N payload_json
        respondentId,          // O respondent_id
        flags,                 // P flags
        resumo,                // Q resumo
        processed_by           // R processed_by
      ]);

      // CRM só se tiver contato + dedupe somente por respondent_id
      if (hasContact && crmCtx.sh) {
        const rid = String(respondentId || '').trim();

        // Se rid vazio, NÃO deduplica (entra de qualquer jeito). Se você preferir bloquear vazio, me fala.
        if (rid && crmRidSet.has(rid)) {
          // já existe a MESMA resposta no CRM, não insere
        } else {
          if (rid) crmRidSet.add(rid);

          const crmRow = new Array(crmCtx.numCols).fill('');

          setCrmByHeader_(crmCtx, crmRow, 'lead_key', leadKey);
          setCrmByHeader_(crmCtx, crmRow, 'raw_id', rawId);
          setCrmByHeader_(crmCtx, crmRow, 'respondent_id', respondentId);

          setCrmByHeader_(crmCtx, crmRow, 'Ação / Tarefa', '');
          setCrmByHeader_(crmCtx, crmRow, 'Status_Funil', 'Novo Lead - Formulário');
          setCrmByHeader_(crmCtx, crmRow, 'Tipo_Relacao', 'Lead');
          setCrmByHeader_(crmCtx, crmRow, 'Temperatura', 'Quente');

          setCrmByHeader_(crmCtx, crmRow, 'Nome', rawName);
          setCrmByHeader_(crmCtx, crmRow, '@ Instagram', rawInstagram);
          setCrmByHeader_(crmCtx, crmRow, 'WhatsApp', rawWhatsapp);
          setCrmByHeader_(crmCtx, crmRow, 'Email', rawEmail);
          setCrmByHeader_(crmCtx, crmRow, 'Anotações / Outros IGs', finalNotes);

          setCrmByHeader_(crmCtx, crmRow, 'Origem_Lead', config.sourceDefault);
          setCrmByHeader_(crmCtx, crmRow, 'Funil_Entrada', finalFunnel);
          setCrmByHeader_(crmCtx, crmRow, 'Interesse [Infoproduto/Serviço]', crmInterest);

          setCrmByHeader_(crmCtx, crmRow, 'Data_Entrada', csvDate);

          crmBatch.push(crmRow);
        }
      }
    }

    // Writes
    if (rawBatch.length > 0) {
      const shRaw = ss.getSheetByName('RAW_RESPONDI');
      shRaw.getRange(shRaw.getLastRow() + 1, 1, rawBatch.length, 18).setValues(rawBatch);
    }

    if (crmBatch.length > 0 && crmCtx.sh) {
      const shCrm = crmCtx.sh;
      shCrm.insertRowsBefore(crmCtx.startRow, crmBatch.length);
      shCrm.getRange(crmCtx.startRow, 1, crmBatch.length, crmCtx.numCols).setValues(crmBatch);

      // Formata Data_Entrada (se existir)
      const dtCol = crmCtx.colIndexByName['data_entrada'];
      if (dtCol != null) {
        shCrm.getRange(crmCtx.startRow, dtCol + 1, crmBatch.length, 1).setNumberFormat("dd/mm/yyyy HH:mm");
      }
      shCrm.getRange(crmCtx.startRow, 1, crmBatch.length, crmCtx.numCols).setVerticalAlignment("middle");
    }

    // Update state + cache
    state = getImportState_() || state;
    state.offset = end;
    state.processed = (state.processed || 0) + processedDelta;
    state.duplicatesCount = (state.duplicatesCount || 0) + duplicatesDelta;
    state.crmCount = (state.crmCount || 0) + crmDelta;
    state.rawOnlyCount = (state.rawOnlyCount || 0) + rawOnlyDelta;
    state.updatedAt = new Date().toISOString();
    state.lastError = '';
    setImportState_(state);

    cacheExistingKeysForForm_(config.formId, existingKeys);

    if (state.cancelled) {
      state.status = 'CANCELLED';
      state.updatedAt = new Date().toISOString();
      setImportState_(state);
      clearImportTriggers_();
      return;
    }

    if (end >= rows.length) {
      state.status = 'COMPLETED';
      state.updatedAt = new Date().toISOString();
      setImportState_(state);
      clearImportTriggers_();
      ss.toast(
        `Importação finalizada! Importados: ${state.processed} | Duplicados RAW: ${state.duplicatesCount} | CRM: ${state.crmCount} | Só RAW: ${state.rawOnlyCount}`,
        'Respondi CSV',
        10
      );
      return;
    }

    state.status = 'SCHEDULING';
    state.updatedAt = new Date().toISOString();
    setImportState_(state);
    scheduleNextBatch_();

  } catch (e) {
    failImportState_(e);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Erro na importação: ${String(e)}`, 'Respondi CSV', 10);
    clearImportTriggers_();
  } finally {
    lock.releaseLock();
  }
}

/** =========================
 * CRM v9: header mapping
 * ========================= */
function getCrmContextV9_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('CRM_GERAL');
  if (!sh) return { sh: null, headerRow: 2, startRow: 3, colIndexByName: {}, numCols: 0 };

  const headerRow = 2;
  const startRow = 3;

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];

  const colIndexByName = {};
  for (let i = 0; i < headers.length; i++) {
    const key = normalizeCrmHeader_(headers[i]);
    if (key) colIndexByName[key] = i; // 0-based
  }

  return { sh, headerRow, startRow, colIndexByName, numCols: lastCol };
}

function normalizeCrmHeader_(v) {
  // Normaliza cabeçalhos para mapear por "nome lógico" (robusto a sufixos explicativos).
  // Ex.: "Interesse (Infoproduto/Serviço)" e "Interesse [Infoproduto/Serviço]" -> "interesse"
  let s = String(v || '').trim();
  if (!s) return '';

  // corta antes de " (" (espaço + abre parênteses)
  let ix = s.indexOf(' (');
  if (ix >= 0) s = s.slice(0, ix);

  // corta antes de " [" (espaço + abre colchetes)
  ix = s.indexOf(' [');
  if (ix >= 0) s = s.slice(0, ix);

  return s.trim().toLowerCase();
}

function setCrmByHeader_(ctx, rowArray, headerName, value) {
  // Usa a mesma normalização do mapa de colunas (evita mismatch quando o cabeçalho tem "(...)" ou "[...]")
  const key = normalizeCrmHeader_(headerName);
  const idx = ctx.colIndexByName[key];
  if (idx == null) return;
  rowArray[idx] = value;
}

function buildCrmRespondentIdSet_(crmCtx) {
  const set = new Set();
  if (!crmCtx.sh) return set;

  const idx = crmCtx.colIndexByName['respondent_id'];
  if (idx == null) return set;

  const lastRow = crmCtx.sh.getLastRow();
  if (lastRow < crmCtx.startRow) return set;

  const num = lastRow - (crmCtx.startRow - 1);
  const values = crmCtx.sh.getRange(crmCtx.startRow, idx + 1, num, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const rid = String(values[i][0] || '').trim();
    if (rid) set.add(rid);
  }
  return set;
}

/** =========================
 * STATE + TRIGGERS
 * ========================= */
function getImportState_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(RESPONDI_IMPORT.STATE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (_) { return null; }
}

function setImportState_(state) {
  PropertiesService.getScriptProperties().setProperty(RESPONDI_IMPORT.STATE_KEY, JSON.stringify(state || {}));
}

function failImportState_(err) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const state = getImportState_() || {};
  state.status = 'FAILED';
  state.lastError = String(err && err.stack ? err.stack : err);
  state.updatedAt = new Date().toISOString();
  setImportState_(state);
  clearImportTriggers_();
  ss.toast(`Falhou: ${state.lastError}`, 'Respondi CSV', 10);
}

function scheduleNextBatch_() {
  clearImportTriggers_();
  ScriptApp.newTrigger(RESPONDI_IMPORT.TRIGGER_HANDLER)
    .timeBased()
    .after(60 * 1000)
    .create();
}

function clearImportTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === RESPONDI_IMPORT.TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(t);
    }
  });
}

/** =========================
 * RAW dedupe cache
 * ========================= */
function cacheExistingKeysForForm_(formId, keySet) {
  try {
    const cache = CacheService.getScriptCache();
    const baseKey = RESPONDI_IMPORT.CACHE_KEYS_PREFIX + '::' + String(formId || '').trim();
    const payload = Array.from(keySet || []).join('\n');
    const b64 = Utilities.base64Encode(Utilities.gzip(payload));
    setLargeCache_(cache, baseKey, b64);
  } catch (_) {}
}

function getCachedExistingKeysForForm_(formId) {
  try {
    const cache = CacheService.getScriptCache();
    const baseKey = RESPONDI_IMPORT.CACHE_KEYS_PREFIX + '::' + String(formId || '').trim();
    const b64 = getLargeCache_(cache, baseKey);
    if (!b64) return null;

    const text = Utilities.newBlob(Utilities.ungzip(Utilities.base64Decode(b64))).getDataAsString('UTF-8');
    const set = new Set();
    if (!text) return set;
    text.split('\n').forEach(s => { if (s) set.add(s); });
    return set;
  } catch (_) {
    return null;
  }
}

function setLargeCache_(cache, baseKey, value) {
  const CHUNK = 90000;
  const totalChunks = Math.ceil(value.length / CHUNK);

  cache.put(baseKey + '::chunks', String(totalChunks), 21600);
  for (let i = 0; i < totalChunks; i++) {
    cache.put(baseKey + '::' + i, value.slice(i * CHUNK, (i + 1) * CHUNK), 21600);
  }
}

function getLargeCache_(cache, baseKey) {
  const chunksStr = cache.get(baseKey + '::chunks');
  const n = chunksStr ? parseInt(chunksStr, 10) : 0;
  if (!n || isNaN(n) || n <= 0) return null;

  let out = '';
  for (let i = 0; i < n; i++) {
    const part = cache.get(baseKey + '::' + i);
    if (part == null) return null;
    out += part;
  }
  return out;
}

/** =========================
 * CONFIG / MAPEAMENTO_FORMS
 * ========================= */
function getFormConfig_(ss, formIdFromFileName, fileName) {
  const sheetMap = ss.getSheetByName('MAPEAMENTO_FORMS');
  if (!sheetMap) {
    return { formId: 'CSV_IMPORT', formName: fileName, funnelDefault: 'FUNIL_AGENCIA', interestDefault: '', sourceDefault: 'Formulário' };
  }

  const mapData = sheetMap.getDataRange().getValues();
  mapData.shift();

  const formMeta = mapData.find(row => String(row[0]).trim() === String(formIdFromFileName).trim());

  return {
    formId: formMeta ? String(formMeta[0]) : 'CSV_IMPORT',
    formName: formMeta ? String(formMeta[1]) : fileName,
    funnelDefault: formMeta ? String(formMeta[2]) : 'FUNIL_AGENCIA',
    interestDefault: formMeta ? String(formMeta[3]) : '',
    sourceDefault: formMeta ? String(formMeta[4]) : 'Formulário'
  };
}

function buildCsvMap_(headers) {
  return {
    nomeIg: headers.findIndex(h => {
      const s = String(h || '').toLowerCase();
      const temNome = s.includes('nome') || s.includes('name');
      const temInsta = s.includes('instagram') || s.includes('insta') || hasWord_(s, 'ig');
      const temArroba = s.includes('@');
      return temNome && temInsta && temArroba;
    }),
    instagram: headers.findIndex(h => {
      const s = String(h || '').toLowerCase();
      const temInsta = s.includes('instagram') || s.includes('insta') || hasWord_(s, 'ig');
      const temArroba = s.includes('@');
      const ehInteresse = s.includes('buscando') || s.includes('interesse');
      return temInsta && temArroba && !ehInteresse;
    }),
    nome: headers.findIndex(h => {
      const s = String(h || '').toLowerCase();
      return s.includes('nome') && !s.includes('@') && !s.includes('instagram');
    }),
    whatsapp: headers.findIndex(h => {
      const s = String(h || '').toLowerCase();
      return s.includes('whatsapp') || s.includes('contato') || s.includes('número') || s.includes('telefone');
    }),
    email: headers.findIndex(h => {
      const s = String(h || '').toLowerCase();
      return s.includes('email') || s.includes('e-mail') || s.includes('mail');
    }),
    interesse: headers.findIndex(h => h && (String(h).toLowerCase().includes('interesse') || String(h).toLowerCase().includes('buscando') || String(h).toLowerCase().includes('produto'))),
    notes: headers.findIndex(h => h && (String(h).toLowerCase().includes('nota') || String(h).toLowerCase().includes('observa') || String(h).toLowerCase().includes('mensagem'))),
    data: headers.findIndex(h => h && (String(h).toLowerCase() === 'data' || String(h).toLowerCase() === 'responded_at')),
    id: headers.findIndex(h => h && (String(h).toLowerCase() === 'id' || String(h).toLowerCase() === 'respondent_id'))
  };
}

/** =========================
 * RAW dedupe: (form_id|respondent_id)
 * ========================= */
function loadExistingRespondentsForForm_(ss, formId) {
  const shRaw = ss.getSheetByName('RAW_RESPONDI');
  const set = new Set();
  if (!shRaw) return set;

  const lastRow = shRaw.getLastRow();
  if (lastRow < 2) return set;

  // Lê C..O (13 cols): form_id (idx 0) e respondent_id (idx 12)
  const data = shRaw.getRange(2, 3, lastRow - 1, 13).getValues();

  const fid = String(formId || '').trim();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rFormId = String(row[0] || '').trim();
    const respondentId = String(row[12] || '').trim();
    if (!respondentId) continue;
    if (rFormId === fid) set.add(fid + '|' + respondentId);
  }
  return set;
}

/** =========================
 * RESUMO
 * ========================= */
function buildResumoFromPayloadJson_IMPORT_(payloadJson) {
  let p;
  try { p = JSON.parse(payloadJson); } catch (_) { return '❌ Erro JSON'; }

  const resp = p?.respondent || {};
  let ans = [];

  if (Array.isArray(resp.raw_answers) && resp.raw_answers.length) {
    ans = resp.raw_answers.map(r => ({
      q: String(r?.question?.question_title || '').toLowerCase(),
      a: Array.isArray(r?.answer) ? r.answer.join(', ') : String(r?.answer ?? '')
    }));
  } else if (resp.answers && typeof resp.answers === 'object') {
    const IGNORAR = new Set(["pontuação", "pontuacao", "data", "id", "score"]);
    ans = Object.entries(resp.answers)
      .filter(([k]) => !IGNORAR.has(String(k).trim().toLowerCase()))
      .map(([k, v]) => ({ q: String(k || '').toLowerCase(), a: Array.isArray(v) ? v.join(', ') : String(v ?? '') }));
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

  let res = "";
  if (faseFinal) res += "📍 Fase: " + faseFinal + "\n";
  if (inv) res += "💳 Investimento: " + inv + "\n";
  if (desafio) res += "⚠️ Desafio: " + desafio + "\n";
  if (objetivo) res += "🎯 Objetivo: " + objetivo + "\n";
  if (preco) res += "💵 Preço: " + preco + "\n";
  if (clientes) res += "👥 Clientes: " + clientes + "\n";
  if (servicos) res += "🧩 Serviços: " + servicos + "\n";

  return res.trim();
}

/** =========================
 * HELPERS
 * ========================= */
function hasWord_(text, word) {
  const s = String(text || '').toLowerCase();
  const re = new RegExp(`(^|[^a-z0-9_])${word}([^a-z0-9_]|$)`, 'i');
  return re.test(s);
}

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
  return { rangeName: 'Infoprodutos', fallback: 'Infoprodutos (precisa analisar o lead)' };
}

function buildLeadKey_(w, e, i, r) {
  if (w && String(w).length > 5) return 'w:' + String(w);
  if (e && String(e).includes('@')) return 'e:' + String(e).toLowerCase().trim();
  if (i && String(i).trim()) return 'i:' + String(i).toLowerCase().replace('@', '').trim();
  return r ? 'r:' + String(r) : '';
}

function normalizeWhatsapp_(v) {
  return String(v || '').replace(/\D/g, '');
}

function normalizeEmail_(v) {
  return String(v || '').trim().toLowerCase();
}

function isEmail_(v) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || '').trim().toLowerCase());
}

function processHandles_(text) {
  const regex = /([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})|@\s*([a-zA-Z0-9._]{2,30})/gi;
  const matchesRaw = [], normalized = [];
  const str = String(text || '');

  for (const m of str.matchAll(regex)) {
    matchesRaw.push(m[0]);
    if (m[1]) normalized.push(String(m[1]).toLowerCase());
    else if (m[2]) normalized.push('@' + String(m[2]));
  }

  // Fallback SEM "@": só considera handle se realmente "parecer" um user (evita "Gabriela" virar @gabriela)
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

function looksLikeHandle_(v) {
  const s = String(v || '').trim();
  if (!s) return false;
  if (/\s/.test(s)) return false;
  return /^@?[a-zA-Z0-9._]{2,30}$/.test(s);
}

function extractInstagramFromAnswers_(answersObj) {
  for (const [k, v] of Object.entries(answersObj || {})) {
    const key = String(k || '').toLowerCase();
    const val = String(v || '').trim();
    if (key.includes('@') && key.includes('instagram')) {
      if (!val) continue;
      const proc = processHandles_(val);
      if (proc.primary && looksLikeHandle_(proc.primary)) return proc.primary;
    }
  }
  return '';
}

function cleanNameWithHandles_(text, handles) {
  let s = String(text || '');
  for (const h of handles || []) {
    s = s.replace(new RegExp(String(h).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), '');
  }
  return s.replace(/[|/\\•·]/g, ' ').replace(/\s{2,}/g, ' ').trim();
}

function normalizeToListOrFallback_(ss, v, rangeName, fallback) {
  const raw = String(v || '').trim();
  if (!raw) return fallback;
  try {
    const r = ss.getRangeByName(rangeName);
    if (!r) return fallback;
    const vs = r.getValues().flat().map(x => String(x).trim()).filter(Boolean);
    const hit = vs.find(x => x.toLowerCase() === raw.toLowerCase());
    return hit ? hit : fallback;
  } catch (_) {
    return fallback;
  }
}

function extractDriveFileIdFromUrl_(input) {
  const s = String(input || '').trim();
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s) && !s.includes('/')) return s;

  const m =
    s.match(/\/d\/([a-zA-Z0-9_-]{20,})/i) ||
    s.match(/[?&]id=([a-zA-Z0-9_-]{20,})/i) ||
    s.match(/\/file\/d\/([a-zA-Z0-9_-]{20,})/i);

  return m ? m[1] : '';
}

function getDriveFileByIdSafe_(fileId) {
  const id = String(fileId || '').trim();
  if (!id) throw new Error('File ID vazio.');
  return DriveApp.getFileById(id);
}
