/**
 * UI_CrmSidebar.gs - v3.3
 * - NOVO: Botão que leva o usuário para a linha do lead (selecionado) na aba CRM_GERAL agora funciona a partir das abas INFOPRODUTOS e SERVICOS_AGENCIA.
 *  
 *  MANTIDO
 * - Submenu para Importar Notion
 * - Seção da sidebar (Bases / Config) COMENTADA
 * - CHAVE LIGA/DELISGA blocos do menu personalizado [DESLIGADOS]
 * 
 * - AMPLIAÇÃO DA BUSCA: 
 * Agora é possível pesquisar formulários por nome ou parte do nome, além do ID que já era possível
 * - MANTIDO:
 * - Menus Backfill_RAW_CRM_v16.2.1_bf3
 * - Servicos_Agencia na atualização de listas
 * - Menus e submenus reorganizados "🖥️ CRM Bia Meneses"
 * - BACKFILL da RAW e do CRM
 * - Sidebar CRM (HTML: CRM_Sidebar.html) com abas: Menu, Respondi e Config
 * - Funções de navegação, visualização de payload e gestão de listas
 */

function abrirCrmSidebar() {
  const tpl = HtmlService.createTemplateFromFile("CRM_Sidebar");
  SpreadsheetApp.getUi().showSidebar(tpl.evaluate().setTitle("Menu lateral"));
}

/** * NAVEGAÇÃO ENTRE ABAS DA PLANILHA 
 */
function navGoToSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("Aba não encontrada: " + sheetName);
  ss.setActiveSheet(sh);
  return { ok: true, sheet: sheetName };
}

/** * CONFIGURAÇÃO DO MENU LATERAL 
 */
function navGetConfig() {
  return {
    activeSheet: SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName(),
    sections: [
      {
        title: "Vistas",
        items: [
          { label: "CRM (Home)", icon: "🏠", sheet: "CRM_GERAL" },
          { label: "Infoprodutos", icon: "📱", sheet: "INFOPRODUTOS" },
          { label: "Serviços Agência", icon: "🧩", sheet: "SERVICOS_AGENCIA" },
        ]
      },   
      {
        title: "Funis",
        items: [
          { label: "Diagnóstico", icon: "🩺", sheet: "FUNIL_DIAGNOSTICO" },
          { label: "Lista de Espera", icon: "⏳", sheet: "FUNIL_LISTA_ESPERA" },
          { label: "Lista VIP / Pré-lançamento", icon: "⭐", sheet: "FUNIL_LISTA_VIP" },
          { label: "Site / Orgânico", icon: "🌐", sheet: "FUNIL_SITE_ORGANICO" },
          { label: "Direct / Inbound", icon: "📥", sheet: "FUNIL_DIRECT" },
          { label: "Eventos / Aulas / Black", icon: "🎟️", sheet: "FUNIL_EVENTOS" },
          { label: "Funil Agência", icon: "🧩", sheet: "FUNIL_AGENCIA" }
        ]
      },
              /** {
                title: "Bases / Config",
                items: [
                  { label: "Mapeamento Forms", icon: "🧭", sheet: "MAPEAMENTO_FORMS" },
                  { label: "RAW Respondi", icon: "🗃️", sheet: "RAW_RESPONDI" },
                  { label: "Config", icon: "⚙️", sheet: "CONFIG" }
                ]
              }*/
    ]
  };
}

// --- CONFIGURAÇÕES ---
const UI_SHEET_RAW = "RAW_RESPONDI";
const UI_COL_RAW_ID_RAW = 0;     // Coluna A (onde fica o ID na base RAW) - Índice 0

// ====== FLAGS (liga/desliga) ======
const ENABLE_ID_FORMAT = false;
const ENABLE_IMPORTAR_RESPONDI = false;
const ENABLE_BACKFILL_RESUMO = false;
const ENABLE_BACKFILL_AVANCADO = false; // <- desabilita sem remover

/**
 * MENU PERSONALIZADO: Menu e Funções de Gestão Manual
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const menu = ui.createMenu("🖥️ CRM Bia Meneses")
    .addItem("📱 Abrir menu lateral", "abrirCrmSidebar")


  if (ENABLE_ID_FORMAT) addMenuIdFormat_(menu);
  if (ENABLE_IMPORTAR_RESPONDI) addMenuImportarRespondi_(ui, menu);
  if (ENABLE_BACKFILL_RESUMO) addMenuBackfillResumo_(ui, menu);
  if (ENABLE_BACKFILL_AVANCADO) addMenuBackfillAvancado_(ui, menu);

  menu.addToUi();
}

// ====== BLOCOS DO MENU (funções auxiliares) ======

function addMenuIdFormat_(menu) {
  menu.addSeparator()
  .addItem("🆔 Gerar IDs para leads manuais/importados", "gerarIdsFaltantesCrm")
  .addItem("✨ Clonar formatação", "copiarFormatacaoAbaModelo")
  .addItem("📩 Importar Notion", "importarDadosNotion");
}

function addMenuImportarRespondi_(ui, menu) {
  menu.addSeparator()
    .addSubMenu(
      ui.createMenu("📥 Importar Respondi")
        .addItem("📩 Importar CSV (Respondi)", "importarCsvRespondiIniciar")
        .addItem("ℹ️ Estado da Importação", "importarCsvRespondiStatus")
        .addItem("⛔ Cancelar importação", "importarCsvRespondiCancelar")
    );
}

function addMenuBackfillResumo_(ui, menu) {
  menu.addSeparator()
    .addSubMenu(
      ui.createMenu("🔄 Backfill RAW Resumo")
        .addItem("▶️ Iniciar backfill (Resumo RAW)", "backfillResumoIniciar")
        .addItem("ℹ️ Status do backfill", "backfillResumoStatus")
        .addItem("⛔ Cancelar backfill", "backfillResumoCancelar")
    );
}

function addMenuBackfillAvancado_(ui, menu) {
  menu.addSeparator()
    .addSubMenu(
      ui.createMenu("🧩 Backfill Avançado")
        .addItem("ℹ️ Status do backfill", "backfillRawCrmStatus")
        .addItem("⛔ Cancelar backfill", "backfillRawCrmCancelar")
        .addItem("🗃️ Backfill apenas RAW", "backfillRAWOnlyReal")
        .addItem("🧪🗃️ Backfill RAW SIMULADO", "backfillRAWOnlyDryRun")
        .addItem("🗃️➡️📊 Backfill RAW + CRM", "backfillRAWandCRMReal")
        .addItem("🧪🗃️➡️📊 Backfill RAW + CRM SIMULADO", "backfillRAWandCRMDryRun")
        .addItem("📊 Backfill apenas CRM", "backfillCRMOnlyReal")
        .addItem("🧪📊 Backfill apenas CRM SIMULADO", "backfillCRMOnlyDryRun")
    );
}

function abrirSidebar() {
  const tpl = HtmlService.createTemplateFromFile("Sidebar");
  SpreadsheetApp.getUi().showSidebar(tpl.evaluate().setTitle("Formulário do lead"));
}

function uiFetchLeadFromActiveSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAtiva = ss.getActiveSheet();
  const linhaAtiva = sheetAtiva.getActiveCell().getRow();
  
  if (linhaAtiva < 2) return { error: "Selecione uma linha de lead válida." }; 

  // ARQUITETURA ROBUSTA: Busca a coluna dinamicamente, ignorando a constante hardcoded.
  // Isso garante funcionamento no CRM_GERAL, espelhos e Funis, tolerando adição de colunas.
  const colRawId = findColByHeader_(sheetAtiva, 'raw_id', 5);
  
  if (!colRawId) {
    return { error: `Cabeçalho 'raw_id' não encontrado na aba '${sheetAtiva.getName()}'. Certifique-se de estar em uma aba de lead.` };
  }

  const rawId = String(sheetAtiva.getRange(linhaAtiva, colRawId).getValue()).trim();
  if (!rawId) return { error: "ID não encontrado nesta linha. O raw_id pode estar vazio." };

  const dadosLead = getLeadDataByRawId_(rawId);
  dadosLead._sheet = sheetAtiva.getName();
  dadosLead._row = linhaAtiva;
  dadosLead._raw_id = rawId;

  return dadosLead;
}

function getLeadDataByRawId_(rawId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRaw = ss.getSheetByName(UI_SHEET_RAW);
  if (!sheetRaw) return { error: `Aba '${UI_SHEET_RAW}' não encontrada.` };

  const lastRow = sheetRaw.getLastRow();
  const idsRaw = sheetRaw.getRange(2, 1, lastRow - 1, 1).getDisplayValues().flat();
  const indexEncontrado = idsRaw.findIndex(id => String(id).trim() === String(rawId).trim());
  if (indexEncontrado === -1) return { error: `ID não encontrado na base RAW.` };

  const linhaRaw = indexEncontrado + 2;
  // IMPORTANTE: Verifique se o JSON está na Coluna N (14) ou O (15) na aba RAW_RESPONDI
  const jsonString = sheetRaw.getRange(linhaRaw, 14).getValue();

  let payload = {};
  try { payload = JSON.parse(jsonString); } catch (e) { return { error: "Erro ao ler JSON." }; }

  const formName = payload?.form?.form_name || "Formulário Desconhecido";
  const respondent = payload?.respondent || {};
  const rawAnswers = Array.isArray(respondent.raw_answers) ? respondent.raw_answers : [];
  
  let qa = rawAnswers.map(ra => ({
    q: ra?.question?.question_title || ra?.question?.question_name || "(pergunta)",
    a: formatRawAnswer_(ra?.answer)
  }));

  if (qa.length === 0 && respondent.answers) {
    qa = Object.entries(respondent.answers).map(([k, v]) => ({
      q: k || "(pergunta)",
      a: formatRawAnswer_(v)
    }));
  }

  return { qa, form_name: formName };
}

function formatRawAnswer_(ans) {
  if (ans == null) return "-";

  // arrays (ex.: radio multi)
  if (Array.isArray(ans)) return ans.join(", ") || "-";

  // objeto phone do Respondi: {country:"55", phone:"6798..."}
  if (typeof ans === "object") {
    const country = ans.country != null ? String(ans.country).replace(/\D/g, "") : "";
    const phone   = ans.phone   != null ? String(ans.phone).replace(/\D/g, "")   : "";

    if (country || phone) return (country + phone) || "-";

    // fallback genérico para outros objetos
    try { return JSON.stringify(ans); } catch (e) { return "-"; }
  }

  // string/number/etc
  const s = String(ans).trim();
  return s || "-";
}


/**
 * GESTÃO DE MAPEAMENTO E LISTAS (CONFIG) - BUSCA AMPLIADA
 * Agora pesquisa por ID exato ou por parte do nome do formulário.
 */
function uiGetFormMapping(searchTerm) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("MAPEAMENTO_FORMS");
  if (!sh) throw new Error("Aba MAPEAMENTO_FORMS não encontrada.");
  
  const data = sh.getDataRange().getValues();
  const body = data.slice(1); // Remove cabeçalho
  const term = String(searchTerm).trim().toLowerCase();
  
  if (!term) return [];

  // Filtra resultados: busca ID exato OU nome que contém o termo
  const results = body.filter(r => {
    const id = String(r[0]).trim().toLowerCase();
    const name = String(r[1]).trim().toLowerCase();
    return id === term || name.includes(term);
  });
  
  // Retorna array de objetos mapeados
  return results.map(row => ({
    form_id: row[0],
    form_name: row[1],
    funnel_default: row[2],
    interest_default: row[3]
  }));
}

function uiGetAllForms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("MAPEAMENTO_FORMS");
  if (!sh) return [];
  
  const data = sh.getDataRange().getValues();
  const body = data.slice(1);
  
  return body.map(row => ({
    form_id: row[0],
    form_name: row[1]
  })).filter(f => f.form_id || f.form_name);
}

function uiSaveFormMapping(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("MAPEAMENTO_FORMS");
  const data = sh.getDataRange().getValues();
  const rowIndex = data.findIndex(r => String(r[0]).trim() === String(formData.form_id).trim());
  
  const newRow = [formData.form_id, formData.form_name, formData.funnel_default, formData.interest_default, "Formulário"];
  
  if (rowIndex === -1) {
    sh.appendRow(newRow);
    return "✅ Formulário cadastrado com sucesso!";
  } else {
    sh.getRange(rowIndex + 1, 1, 1, 5).setValues([newRow]);
    return "🔄 Mapeamento atualizado com sucesso!";
  }
}

function uiGetConfigLists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listNames = ["Infoprodutos","Objeções","Origem_Lead","Responsável","Segmento_Perfil","Servicos_Agencia","Status_Funil","Tipo_Relacao","Interesse","Funis_Entrada"];
  const results = {};
  listNames.forEach(name => {
    try {
      const range = ss.getRangeByName(name);
      results[name] = range ? range.getValues().flat().filter(v => v !== "" && v !== null) : [];
    } catch(e) { results[name] = []; }
  });
  return results;
}

function uiSaveConfigList(listName, content) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const namedRange = ss.getRangeByName(listName);
  if (!namedRange) throw new Error("Intervalo '" + listName + "' não encontrado.");

  const sheet = namedRange.getSheet();
  const col = namedRange.getColumn();
  const items = content.split('\n').map(i => [i.trim()]).filter(i => i[0] !== "");

  sheet.getRange(2, col, sheet.getMaxRows() - 2, 1).clearContent();
  if (items.length > 0) {
    const newRange = sheet.getRange(2, col, items.length, 1);
    newRange.setValues(items);
    ss.setNamedRange(listName, newRange);
  }
  return "✅ Lista '" + listName + "' atualizada!";
}

function getCurrentSheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

/** =============================================
 * SIDEBAR: Ir para o lead no CRM_GERAL usando respondent_id da LINHA ATIVA
 * - Lê o respondent_id na aba/linha onde o usuário está clicado agora
 * - Procura no CRM_GERAL via TextFinder (rápido)
 * - Se encontrar, ativa a planilha e a célula da linha encontrada
 *
 * Regras de cabeçalho:
 * - Procura "respondent_id" nos primeiros 5 rows do sheet ativo (case-insensitive)
 * - No CRM, também procura "respondent_id" nos primeiros 5 rows
 * - Dados no CRM normalmente começam na linha 3
 * ==============================================
 */

function uiGoToCrmByRespondentIdFromSelection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shFrom = ss.getActiveSheet();
  const activeCell = shFrom.getActiveCell();
  if (!activeCell) return { ok: false, error: 'Selecione uma célula primeiro.' };

  const row = activeCell.getRow();
  if (row < 2) return { ok: false, error: 'Selecione uma linha de lead válida.' };

  // 1) Ler respondent_id na linha ativa (por cabeçalho "respondent_id")
  const colRidFrom = findColByHeader_(shFrom, 'respondent_id', 5);
  if (!colRidFrom) {
    return { ok: false, error: `Não achei o cabeçalho "respondent_id" na aba atual (${shFrom.getName()}).` };
  }

  const respondentId = String(shFrom.getRange(row, colRidFrom).getDisplayValue() || '').trim();
  if (!respondentId) return { ok: false, error: 'respondent_id está vazio na linha selecionada.' };

  // 2) Procurar no CRM_GERAL por TextFinder
  const shCrm = ss.getSheetByName('CRM_GERAL');
  if (!shCrm) return { ok: false, error: 'Aba CRM_GERAL não encontrada.' };

  const colRidCrm = findColByHeader_(shCrm, 'respondent_id', 5);
  if (!colRidCrm) return { ok: false, error: 'Não achei o cabeçalho "respondent_id" no CRM_GERAL.' };

  const lastRow = shCrm.getLastRow();
  if (lastRow < 3) return { ok: false, error: 'CRM_GERAL parece vazio.' };

  // Começa do 1º row de dados mais provável: 3
  const startRow = 3;
  const numRows = Math.max(0, lastRow - startRow + 1);
  if (numRows === 0) return { ok: false, error: 'CRM_GERAL não tem linhas de dados.' };

  const rangeRid = shCrm.getRange(startRow, colRidCrm, numRows, 1);

  const finder = rangeRid
    .createTextFinder(respondentId)
    .matchEntireCell(true)
    .matchCase(false);

  const found = finder.findNext();
  if (!found) {
    return { ok: false, error: `respondent_id não encontrado no CRM_GERAL: ${respondentId}` };
  }

  // 3) Levar o usuário até lá
  ss.setActiveSheet(shCrm);
  shCrm.getRange(found.getRow(), 1).activate(); // ativa a linha (col A)
  return { ok: true, respondent_id: respondentId, row: found.getRow() };
}

/**
 * Acha a coluna de um cabeçalho (case-insensitive) procurando nos primeiros N rows.
 * Retorna número da coluna (1-based) ou null.
 */
function findColByHeader_(sh, headerName, maxHeaderRows) {
  const target = String(headerName || '').trim().toLowerCase();
  const lastCol = Math.max(1, sh.getLastColumn());
  const rowsToScan = Math.max(1, Number(maxHeaderRows || 5));

  for (let r = 1; r <= rowsToScan; r++) {
    const vals = sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    for (let c = 0; c < vals.length; c++) {
      const v = String(vals[c] || '').trim().toLowerCase();
      if (v === target) return c + 1; // 1-based
    }
  }
  return null;
}
