/**
 * IMPORTADOR NOTION -> CRM_GERAL — v1.3 (Data_Entrada como Date + parse dd/mm/yyyy)
 * 
 * CORREÇÃO na data de entrada que estava sendo ignorada e escrita a data da importação.
 * O que mudou vs versão antiga:
 * - Antes: montava um array fixo (29 colunas) e escrevia por índice.
 * - Agora: lê os cabeçalhos do CRM_GERAL na LINHA 2 e escreve por NOME.
 *   Assim você pode reordenar/inserir colunas no CRM sem quebrar o importador.
 *
 * Observações:
 * - Este importador NÃO cria lead_key automaticamente.
 * - Ele insere os leads logo abaixo dos cabeçalhos (linha 2).
 * - Ele tenta herdar FORMATO + DATA VALIDATION (chips/dropdowns) do primeiro lead antigo.
 */

function importarDadosNotion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shImport = ss.getSheetByName('IMPORTAR_NOTION');
  const shCrm = ss.getSheetByName('CRM_GERAL');

  if (!shImport || !shCrm) {
    SpreadsheetApp.getUi().alert("Erro: Certifique-se de que as abas 'IMPORTAR_NOTION' e 'CRM_GERAL' existem.");
    return;
  }

  const data = shImport.getDataRange().getValues();
  if (data.length < 2) {
    SpreadsheetApp.getUi().alert('A aba IMPORTAR_NOTION parece estar vazia.');
    return;
  }

  const crmCtx = getCrmHeaderContext_(shCrm);
  if (!crmCtx || crmCtx.numCols < 1) {
    SpreadsheetApp.getUi().alert('Erro: não consegui ler os cabeçalhos da linha 2 do CRM_GERAL.');
    return;
  }

  const leadsParaInserir = [];
  const dataAtual = new Date();

  // Percorre dados ignorando cabeçalho da IMPORTAR_NOTION (linha 1)
  for (let i = 1; i < data.length; i++) {
    const r = data[i];

    // --- MAPEAMENTO DOS DADOS DO NOTION (mantido igual ao seu script original) ---
    const nome        = String(r[0] || '');
    const instagram   = String(r[1] || '');
    const status      = String(r[2] || 'Importado Notion');
    const dataEntrada = parseNotionDate_(r[8]);
    const responsavel = String(r[11] || '');
    const notas       = String(r[12] || '');
    const whatsapp    = String(r[13] || '');
    const observacoes = String(r[15] || '');

    // --- CONSTRUÇÃO DA LINHA DO CRM (tamanho dinâmico) ---
    const novaLinha = new Array(crmCtx.numCols).fill('');

    // Campos principais (pelos cabeçalhos reais do CRM)
    setCrmByHeader_(crmCtx, novaLinha, 'lead_key', '');
    setCrmByHeader_(crmCtx, novaLinha, 'raw_id', '');
    setCrmByHeader_(crmCtx, novaLinha, 'respondent_id', '');

    setCrmByHeader_(crmCtx, novaLinha, 'Ação / Tarefa', '');
    setCrmByHeader_(crmCtx, novaLinha, 'Status_Funil', status);
    setCrmByHeader_(crmCtx, novaLinha, 'Tipo_Relacao', 'Lead');
    setCrmByHeader_(crmCtx, novaLinha, 'Temperatura', 'Morno');

    setCrmByHeader_(crmCtx, novaLinha, 'Nome', nome);
    setCrmByHeader_(crmCtx, novaLinha, '@ Instagram', instagram);
    setCrmByHeader_(crmCtx, novaLinha, 'WhatsApp', whatsapp);
    setCrmByHeader_(crmCtx, novaLinha, 'Email', '');
    setCrmByHeader_(crmCtx, novaLinha, 'Anotações / Outros IGs', notas);

    setCrmByHeader_(crmCtx, novaLinha, 'Origem_Lead', 'Notion Antigo');
    setCrmByHeader_(crmCtx, novaLinha, 'Funil_Entrada', 'Importado Notion');
    // Interesse / Segmento_Perfil ficam em branco por padrão

    setCrmByHeader_(crmCtx, novaLinha, 'Data_Entrada', dataEntrada);

    // Observações e responsável
    setCrmByHeader_(crmCtx, novaLinha, 'Observações', observacoes);
    setCrmByHeader_(crmCtx, novaLinha, 'Responsável', responsavel);

    leadsParaInserir.push(novaLinha);
  }

  const numNovosLeads = leadsParaInserir.length;
  if (numNovosLeads === 0) {
    SpreadsheetApp.getUi().alert('Nada para importar.');
    return;
  }

  // Existe algum lead já na planilha? (linha 3 é o primeiro lead)
  const temTemplate = shCrm.getLastRow() >= 3;

  // Se tem template, guardo a altura atual da linha 3 (antes de inserir)
  const alturaTemplateAntes = temTemplate ? shCrm.getRowHeight(3) : 35;

  // 1) Insere as linhas abaixo do cabeçalho (linha 2)
  shCrm.insertRowsAfter(2, numNovosLeads);

  // 2) Intervalo alvo (as novas linhas inseridas)
  const rangeAlvo = shCrm.getRange(3, 1, numNovosLeads, crmCtx.numCols);

  // 3) Preenche valores
  rangeAlvo.setValues(leadsParaInserir);

  // 4) Aplica altura (herdando do template; se não existir, usa 35)
  shCrm.setRowHeights(3, numNovosLeads, alturaTemplateAntes);

  // 5) Se já havia leads, copia FORMATO + VALIDAÇÃO do primeiro lead antigo
  //    Depois do insert, o antigo "primeiro lead" (linha 3) virou (3 + numNovosLeads)
  if (temTemplate) {
    const linhaTemplate = shCrm.getRange(3 + numNovosLeads, 1, 1, crmCtx.numCols);

    // Copia só o formato (fonte, cor, alinhamento, bordas, etc.)
    linhaTemplate.copyTo(rangeAlvo, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

    // Copia só validações (chips/dropdowns)
    linhaTemplate.copyTo(rangeAlvo, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  } else {
    // Se não tem template (planilha vazia), aplica um mínimo opcional
    rangeAlvo
      .setFontFamily('Arial')
      .setFontSize(10)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('left');
    rangeAlvo.setBorder(true, true, true, true, true, true, '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);
  }

  // Formata Data_Entrada (se existir). Importante para validação "A data é válida"
  const dtCol = crmCtx.colIndexByName['data_entrada'];
  if (dtCol != null) {
    shCrm.getRange(3, dtCol + 1, numNovosLeads, 1).setNumberFormat("dd/MM/yyyy");
  }

  SpreadsheetApp.getUi().alert(`${numNovosLeads} leads importados com formato/chips herdados da linha abaixo!`);
}



/**
 * Converte data vinda do IMPORTAR_NOTION para um Date real.
 * Aceita:
 * - Date (quando a célula de origem já é tipo data)
 * - string "dd/mm/yyyy" (ou "dd/mm/yyyy hh:mm")
 * Retorna '' se inválido/vazio.
 */
function parseNotionDate_(v) {
  if (v == null || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    // já é Date válido
    return v;
  }

  const s = String(v).trim();
  if (!s) return '';

  // dd/mm/yyyy [hh:mm]
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?$/);
  if (!m) return '';

  const dd = Number(m[1]);
  const mm = Number(m[2]);
  const yyyy = Number(m[3]);
  const hh = m[4] != null ? Number(m[4]) : 0;
  const mi = m[5] != null ? Number(m[5]) : 0;

  const d = new Date(yyyy, mm - 1, dd, hh, mi, 0, 0);
  if (isNaN(d.getTime())) return '';
  return d;
}


/**
 * Cria um contexto de mapeamento de colunas do CRM por cabeçalho (linha 2).
 */
function getCrmHeaderContext_(shCrm) {
  const lastCol = shCrm.getLastColumn();
  if (lastCol < 1) return null;

  const headers = shCrm.getRange(2, 1, 1, lastCol).getValues()[0];
  const colIndexByName = {};

  for (let c = 0; c < headers.length; c++) {
    const raw = headers[c];
    if (raw === '' || raw == null) continue;
    const key = normalizeCrmHeader_(raw);
    if (!key) continue;
    // Se existir duplicado, mantém o PRIMEIRO (mais conservador)
    if (colIndexByName[key] == null) colIndexByName[key] = c;
  }

  return {
    numCols: lastCol,
    headers,
    colIndexByName,
  };
}

/**
 * Normaliza cabeçalho do CRM:
 * - lower + trim
 * - corta sufixo tipo " ( ..." e " [ ..." para tolerar variações
 */
function normalizeCrmHeader_(h) {
  let s = String(h || '').trim().toLowerCase();
  s = s.split(' (')[0];
  s = s.split(' [')[0];
  return s;
}

/**
 * Define valor na linha do CRM usando nome do cabeçalho.
 */
function setCrmByHeader_(ctx, rowArray, headerName, value) {
  const key = normalizeCrmHeader_(headerName);
  const idx = ctx.colIndexByName[key];
  if (idx == null) return;
  rowArray[idx] = value;
}
