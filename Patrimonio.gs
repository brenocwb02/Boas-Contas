/**
 * @file Patrimonio.gs
 * @description Contém a lógica para calcular o Patrimônio Líquido (Ativos - Passivos).
 */

const SHEET_ATIVOS = "Ativos";
const SHEET_PASSIVOS = "Passivos";

/**
 * @private
 * Lê e calcula o valor total dos ativos da planilha.
 * @returns {number} O valor total dos ativos.
 */
function _getAssets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ATIVOS);
  if (!sheet || sheet.getLastRow() < 2) {
    logToSheet("[Patrimonio] Aba 'Ativos' não encontrada ou vazia.", "INFO");
    return 0;
  }

  const data = sheet.getRange("C2:C" + sheet.getLastRow()).getValues();
  let totalAssets = 0;
  data.forEach(row => {
    const value = parseBrazilianFloat(String(row[0]));
    if (!isNaN(value)) {
      totalAssets += value;
    }
  });
  return totalAssets;
}

/**
 * @private
 * Lê e calcula o valor total dos passivos da planilha.
 * @returns {number} O valor total dos passivos.
 */
function _getLiabilities() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PASSIVOS);
  if (!sheet || sheet.getLastRow() < 2) {
    logToSheet("[Patrimonio] Aba 'Passivos' não encontrada ou vazia.", "INFO");
    return 0;
  }

  const data = sheet.getRange("C2:C" + sheet.getLastRow()).getValues();
  let totalLiabilities = 0;
  data.forEach(row => {
    const value = parseBrazilianFloat(String(row[0]));
    if (!isNaN(value)) {
      totalLiabilities += value;
    }
  });
  return totalLiabilities;
}

/**
 * Calcula o Patrimônio Líquido total.
 * @returns {object} Um objeto contendo o total de ativos, passivos e o patrimônio líquido.
 */
function calculateNetWorth() {
  const totalAssets = _getAssets();
  const totalLiabilities = _getLiabilities();
  const netWorth = totalAssets - totalLiabilities;

  logToSheet(`[Patrimonio] Cálculo concluído: Ativos=${totalAssets}, Passivos=${totalLiabilities}, PL=${netWorth}`, "INFO");

  return {
    assets: totalAssets,
    liabilities: totalLiabilities,
    netWorth: netWorth
  };
}
