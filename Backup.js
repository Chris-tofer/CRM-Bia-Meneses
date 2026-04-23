// BACKUP CRM - v2.1 (Incremental + Snapshot diário com retenção)
// - Mantém 1 aba incremental (append só do que é novo) para minimizar perda
// - Mantém 1 snapshot por dia (sobrescreve no mesmo dia) + apaga snapshots antigos (retenção)
// - Detecta automaticamente a coluna-chave pelo cabeçalho (ex: raw_id / respondent_id)
// Observação: este script NÃO depende de mapeamento fixo de colunas; ele copia a aba inteira.

// --- CONFIGURAÇÕES DO BACKUP ---
const CONFIG_BACKUP = {
  ID_PLANILHA_DESTINO: "1mC29QNNBqStvZhS498cLfjaQ81Iei9kELRCIlZX027Q",
  NOME_ABA_ORIGEM: "RAW_RESPONDI",

  // Abas no arquivo de backup:
  NOME_ABA_INCREMENTAL: "RAW_BACKUP_INCREMENTAL",
  NOME_ABA_LOGS: "📝 LOGS",

  // Snapshots diários (1 aba por dia, sobrescreve no mesmo dia):
  ENABLE_DAILY_SNAPSHOT: true,
  SNAPSHOT_PREFIX: "RAW_SNAP_",
  SNAPSHOT_RETENTION_DAYS: 21,

  // Estrutura da RAW:
  HEADER_ROW: 2, // seus cabeçalhos "reais" estão na linha 2 (ignore a linha 1)
  KEY_FIELDS_PRIORITY: ["raw_id", "respondent_id"], // tenta nessa ordem

  // Alertas:
  EMAIL_NOTIFICACAO: "chrisalexsander1@gmail.com"
};

/**
 * Roda o backup:
 * 1) Incremental (append apenas linhas novas, dedupe por raw_id/respondent_id)
 * 2) Snapshot diário (opcional) com retenção
 */
function executarBackup() {
  const ssOrigem = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOrigem = ssOrigem.getSheetByName(CONFIG_BACKUP.NOME_ABA_ORIGEM);

  if (!sheetOrigem) {
    enviarAlerta("Erro Crítico", "A aba de origem '" + CONFIG_BACKUP.NOME_ABA_ORIGEM + "' não foi encontrada no CRM.");
    return;
  }

  let ssDestino;
  try {
    ssDestino = SpreadsheetApp.openById(CONFIG_BACKUP.ID_PLANILHA_DESTINO);
  } catch (e) {
    enviarAlerta("Erro Crítico", "Não consegui abrir a planilha de backup (ID inválido ou sem permissão).\n\n" + e);
    return;
  }

  try {
    const dadosOrigem = sheetOrigem.getDataRange().getValues();
    if (!dadosOrigem || dadosOrigem.length === 0) {
      registrarLog(ssDestino, "AVISO", "A aba de origem está vazia. Nada para backupar.");
      return;
    }

    // 1) Incremental
    const incResult = backupIncremental_(ssDestino, sheetOrigem, dadosOrigem);

    // 2) Snapshot diário + retenção
    let snapMsg = "Snapshot diário DESLIGADO.";
    if (CONFIG_BACKUP.ENABLE_DAILY_SNAPSHOT) {
      const snapResult = backupSnapshotDiario_(ssDestino, dadosOrigem);
      limpezaRetencaoSnapshots_(ssDestino);
      snapMsg = `Snapshot diário OK: ${snapResult.nomeAba} (linhas: ${snapResult.linhas}). Retenção: ${CONFIG_BACKUP.SNAPSHOT_RETENTION_DAYS} dias.`;
    }

    registrarLog(
      ssDestino,
      "SUCESSO",
      `Incremental OK: +${incResult.inseridas} linhas novas (chave: ${incResult.chaveHeader || "N/D"}). Total origem: ${dadosOrigem.length}. ` + snapMsg
    );
  } catch (erro) {
    console.error(erro);
    try {
      registrarLog(ssDestino, "ERRO", erro.toString());
    } catch (e) {
      console.error("Falha ao registrar log: " + e.toString());
    }
    enviarAlerta("FALHA no Backup CRM", "Ocorreu um erro ao tentar realizar o backup:\n\n" + erro.toString());
  }
}

/**
 * Backup incremental: mantém 1 aba e só APPENDA linhas novas.
 * Dedupe pela coluna-chave identificada no HEADER_ROW.
 */
function backupIncremental_(ssDestino, sheetOrigem, dadosOrigem) {
  const headerRow = CONFIG_BACKUP.HEADER_ROW;
  const headers = dadosOrigem[headerRow - 1] || [];
  const keyColIndex1Based = findKeyColumnIndex_(headers, CONFIG_BACKUP.KEY_FIELDS_PRIORITY); // 1-based
  const keyHeader = keyColIndex1Based ? String(headers[keyColIndex1Based - 1]).trim() : "";

  // Garante aba incremental
  let sheetInc = ssDestino.getSheetByName(CONFIG_BACKUP.NOME_ABA_INCREMENTAL);
  if (!sheetInc) {
    sheetInc = ssDestino.insertSheet(CONFIG_BACKUP.NOME_ABA_INCREMENTAL, 0);
    // Copia as linhas 1..HEADER_ROW (inclui linha 1 "qualquer coisa" + linha 2 cabeçalhos)
    const headerBlock = dadosOrigem.slice(0, headerRow);
    sheetInc.getRange(1, 1, headerBlock.length, headerBlock[0].length).setValues(headerBlock);
    sheetInc.setFrozenRows(headerRow);
  }

  // Se a estrutura (nº de colunas) mudou, ajusta para caber (não perde conteúdo existente)
  const origemCols = dadosOrigem[0].length;
  if (sheetInc.getMaxColumns() < origemCols) {
    sheetInc.insertColumnsAfter(sheetInc.getMaxColumns(), origemCols - sheetInc.getMaxColumns());
  }

  // Atualiza cabeçalhos (linhas 1..HEADER_ROW) para refletir a versão atual da RAW
  const headerBlockAtual = dadosOrigem.slice(0, headerRow);
  sheetInc.getRange(1, 1, headerBlockAtual.length, headerBlockAtual[0].length).setValues(headerBlockAtual);

  // Se não achou coluna-chave, cai para append por "novas linhas" (pior caso)
  if (!keyColIndex1Based) {
    const lastRow = sheetInc.getLastRow();
    const startRow = Math.max(lastRow + 1, headerRow + 1);
    const novos = dadosOrigem.slice(startRow - 1); // 0-based
    if (novos.length > 0) {
      sheetInc.getRange(startRow, 1, novos.length, origemCols).setValues(novos);
    }
    return { inseridas: novos.length, chaveHeader: "" };
  }

  // Carrega chaves já existentes no incremental para dedupe
  const lastRowInc = sheetInc.getLastRow();
  const firstDataRow = headerRow + 1;
  const existingKeys = new Set();

  if (lastRowInc >= firstDataRow) {
    const keyRange = sheetInc.getRange(firstDataRow, keyColIndex1Based, lastRowInc - headerRow, 1).getValues();
    keyRange.forEach(r => {
      const k = String(r[0] || "").trim();
      if (k) existingKeys.add(k);
    });
  }

  // Seleciona linhas novas da origem
  const newRows = [];
  for (let r = headerRow; r < dadosOrigem.length; r++) {
    const row = dadosOrigem[r];
    const key = String(row[keyColIndex1Based - 1] || "").trim();
    if (!key) continue; // sem chave, ignora
    if (!existingKeys.has(key)) {
      existingKeys.add(key);
      newRows.push(row);
    }
  }

  // Append em bloco
  if (newRows.length > 0) {
    const appendAt = sheetInc.getLastRow() + 1;
    sheetInc.getRange(appendAt, 1, newRows.length, origemCols).setValues(newRows);
  }

  return { inseridas: newRows.length, chaveHeader: keyHeader };
}

/**
 * Snapshot diário: 1 aba por dia, sempre sobrescreve a do dia.
 */
function backupSnapshotDiario_(ssDestino, dadosOrigem) {
  const dataDia = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const nomeAba = CONFIG_BACKUP.SNAPSHOT_PREFIX + dataDia;

  const abaExistente = ssDestino.getSheetByName(nomeAba);
  if (abaExistente) {
    // Limpa e regrava (mais rápido/menos agressivo que deletar)
    abaExistente.clearContents();
    if (abaExistente.getMaxColumns() < dadosOrigem[0].length) {
      abaExistente.insertColumnsAfter(abaExistente.getMaxColumns(), dadosOrigem[0].length - abaExistente.getMaxColumns());
    }
    if (abaExistente.getMaxRows() < dadosOrigem.length) {
      abaExistente.insertRowsAfter(abaExistente.getMaxRows(), dadosOrigem.length - abaExistente.getMaxRows());
    }
    abaExistente.getRange(1, 1, dadosOrigem.length, dadosOrigem[0].length).setValues(dadosOrigem);
    abaExistente.setFrozenRows(CONFIG_BACKUP.HEADER_ROW);
    return { nomeAba, linhas: dadosOrigem.length };
  }

  const novaAba = ssDestino.insertSheet(nomeAba, 0);
  novaAba.getRange(1, 1, dadosOrigem.length, dadosOrigem[0].length).setValues(dadosOrigem);
  novaAba.setFrozenRows(CONFIG_BACKUP.HEADER_ROW);
  return { nomeAba, linhas: dadosOrigem.length };
}

/**
 * Apaga snapshots antigos além da retenção.
 * Refatorado para maior precisão na comparação de datas e logs de execução.
 */
function limpezaRetencaoSnapshots_(ssDestino) {
  const prefix = CONFIG_BACKUP.SNAPSHOT_PREFIX;
  const retentionDays = CONFIG_BACKUP.SNAPSHOT_RETENTION_DAYS;

  const hoje = new Date();
  // Normaliza o cutoff para o início do dia para evitar problemas com milissegundos
  const cutoff = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - retentionDays);
  const cutoffTime = cutoff.getTime();

  const sheets = ssDestino.getSheets();
  let deletadasCount = 0;

  sheets.forEach(sh => {
    const name = sh.getName();
    
    // 1. Filtro por prefixo
    if (!name.startsWith(prefix)) return;

    // 2. Extração e Parsing da data da aba
    const dateStr = name.substring(prefix.length);
    const parsedDate = parseDateYmd_(dateStr);
    
    if (!parsedDate) {
      console.warn(`Aba ignorada (formato de data inválido): ${name}`);
      return;
    }

    // 3. Comparação de timestamp
    if (parsedDate.getTime() < cutoffTime) {
      // Segurança adicional: nunca deletar abas fixas de configuração
      if (name === CONFIG_BACKUP.NOME_ABA_INCREMENTAL || name === CONFIG_BACKUP.NOME_ABA_LOGS) return;
      
      try {
        ssDestino.deleteSheet(sh);
        deletadasCount++;
      } catch (e) {
        console.error(`Erro ao deletar aba ${name}: ${e.message}`);
      }
    }
  });

  if (deletadasCount > 0) {
    registrarLog(ssDestino, "LIMPEZA", `Removidos ${deletadasCount} snapshots antigos (anteriores a ${Utilities.formatDate(cutoff, Session.getScriptTimeZone(), "yyyy-MM-dd")}).`);
  }
}

/**
 * Encontra a coluna-chave pelo cabeçalho, respeitando prioridade.
 * Retorna índice 1-based ou null.
 */
function findKeyColumnIndex_(headersRow, keyPriority) {
  if (!headersRow || headersRow.length === 0) return null;

  const normalized = headersRow.map(h => String(h || "").trim());
  for (let i = 0; i < keyPriority.length; i++) {
    const target = keyPriority[i];
    const idx = normalized.findIndex(h => h === target);
    if (idx !== -1) return idx + 1; // 1-based
  }
  return null;
}

/**
 * Parse seguro de "yyyy-MM-dd" em Date (timezone local do script).
 */
function parseDateYmd_(ymd) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || "").trim());
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]) - 1;
  const d = Number(m[3]);
  const dt = new Date(y, mo, d);
  return isNaN(dt.getTime()) ? null : dt;
}

// --- FUNÇÕES AUXILIARES ---

function registrarLog(ssDestino, status, mensagem) {
  let sheetLog = ssDestino.getSheetByName(CONFIG_BACKUP.NOME_ABA_LOGS);

  if (!sheetLog) {
    sheetLog = ssDestino.insertSheet(CONFIG_BACKUP.NOME_ABA_LOGS);
    sheetLog.appendRow(["Data/Hora", "Status", "Mensagem"]);
    sheetLog.setFrozenRows(1);
    sheetLog.getRange("A1:C1").setFontWeight("bold");
    sheetLog.setColumnWidth(1, 150);
    sheetLog.setColumnWidth(3, 600);
  }

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  sheetLog.appendRow([timestamp, status, mensagem]);
}

function enviarAlerta(assunto, corpo) {
  if (CONFIG_BACKUP.EMAIL_NOTIFICACAO && CONFIG_BACKUP.EMAIL_NOTIFICACAO.includes("@")) {
    MailApp.sendEmail({
      to: CONFIG_BACKUP.EMAIL_NOTIFICACAO,
      subject: "[ALERTA CRM] " + assunto,
      body: corpo
    });
  }
}

// --- (Opcional) Atalho para manter compatibilidade com seu trigger atual ---
// Se seu gatilho chama "executarBackupSnapshot", você pode manter este wrapper.
function executarBackupSnapshot() {
  executarBackup();
}
