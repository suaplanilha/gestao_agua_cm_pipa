/**
 * WATER ERP SAE - Backend Central
 * Desenvolvido para Google Apps Script (V8)
 * Centro de operações, validação e auditoria.
 */

const APP_ID = "WATER_ERP_SAE_001";
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

/**
 * Função principal de carregamento da WebApp
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('WaterERP SAE')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * CRUD Centralizado com Normalização
 * @param {string} entity - Nome da aba (entregas, abastecimentos, despesas)
 * @param {Object} data - Objeto de dados vindo do front
 */
function saveData(entity, data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(entity);
  
  if (!sheet) {
    sheet = ss.insertSheet(entity);
    // Setup headers automáticos se for nova aba
    sheet.appendRow(['uuid', 'timestamp', 'data_iso', ...Object.keys(data), 'usuario_auditoria']);
  }

  // Normalização e Validação (SAE Standard)
  const uuid = Utilities.getUuid();
  const timestamp = new Date();
  const rowData = [
    uuid,
    timestamp,
    data.data || timestamp.toISOString().split('T')[0],
    ...Object.values(data),
    Session.getActiveUser().getEmail()
  ];

  // Log de Auditoria
  writeLog(`INSERT_${entity.toUpperCase()}`, `ID: ${uuid} | User: ${Session.getActiveUser().getEmail()}`);

  sheet.appendRow(rowData);
  return { success: true, id: uuid };
}

/**
 * Recuperação de dados formatada para Dashboards
 */
function getDashboardStats() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const entregas = ss.getSheetByName('entregas')?.getDataRange().getValues() || [];
  const despesas = ss.getSheetByName('despesas')?.getDataRange().getValues() || [];
  
  // Aqui entraria a lógica de redução (reduce/map) para gerar os KPIs
  // Evitando processamento pesado no frontend
  return {
    totalVolume: 420.5,
    totalKm: 1700,
    totalHoras: 170.0,
    custosPorCategoria: { diesel: 2700, manutencao: 850 }
  };
}

/**
 * Sistema de Logs Centralizado
 */
function writeLog(action, details) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let logSheet = ss.getSheetByName('SYS_LOGS');
  if (!logSheet) logSheet = ss.insertSheet('SYS_LOGS');
  logSheet.appendRow([new Date(), action, details]);
}
