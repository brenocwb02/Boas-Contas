/**
 * @file Patrimonio.gs
 * @description Contém a lógica para calcular o Patrimônio Líquido (Ativos - Passivos).
 * VERSÃO MELHORADA: Agora considera todas as fontes de ativos e passivos.
 */

const SHEET_ATIVOS_MANUAIS = "Ativos";
const SHEET_PASSIVOS_MANUAIS = "Passivos";

/**
 * @private
 * Lê e calcula o valor total de todos os ativos.
 * ATIVOS = (Saldos positivos em Contas) + (Valor dos Investimentos) + (Outros Ativos Manuais)
 * @returns {number} O valor total dos ativos.
 */
function _getAssets() {
  // 1. Garante que os saldos de todas as contas estão atualizados
  atualizarSaldosDasContas(); 
  
  let totalSaldosPositivos = 0;
  for (const key in globalThis.saldosCalculados) {
      const conta = globalThis.saldosCalculados[key];
      if ((conta.tipo === 'conta corrente' || conta.tipo === 'dinheiro físico') && conta.saldo > 0) {
          totalSaldosPositivos += conta.saldo;
      }
  }

  // 2. Soma o valor total dos investimentos da aba "Investimentos"
  const totalInvestments = getTotalInvestmentsValue();

  // 3. Soma outros ativos (imóveis, veículos) da aba "Ativos"
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ATIVOS_MANUAIS);
  let totalAtivosManuais = 0;
  if (sheet && sheet.getLastRow() >= 2) {
      const data = sheet.getRange("C2:C" + sheet.getLastRow()).getValues();
      data.forEach(row => {
          totalAtivosManuais += parseBrazilianFloat(String(row[0]));
      });
  }
  
  const totalAssets = totalSaldosPositivos + totalInvestments + totalAtivosManuais;
  logToSheet(`[Patrimonio] Ativos Calculados: Saldos(${totalSaldosPositivos}) + Investimentos(${totalInvestments}) + Manuais(${totalAtivosManuais}) = ${totalAssets}`, "DEBUG");
  return totalAssets;
}

/**
 * @private
 * Lê e calcula o valor total de todos os passivos.
 * PASSIVOS = (Dívida total de cartões) + (Saldos negativos em contas) + (Outros Passivos Manuais)
 * @returns {number} O valor total dos passivos.
 */
function _getLiabilities() {
  // 1. Garante que os saldos estão atualizados
  atualizarSaldosDasContas();
  
  let totalDividas = 0;
  for (const key in globalThis.saldosCalculados) {
      const conta = globalThis.saldosCalculados[key];
      // Soma a dívida total de todos os cartões de crédito e faturas consolidadas
      if (conta.tipo === 'cartão de crédito' || conta.tipo === 'fatura consolidada') {
          totalDividas += conta.saldoTotalPendente;
      }
      // Soma o saldo negativo (cheque especial) de contas correntes
      if (conta.tipo === 'conta corrente' && conta.saldo < 0) {
          totalDividas += Math.abs(conta.saldo);
      }
  }

  // 2. Soma outros passivos (financiamentos, empréstimos) da aba "Passivos"
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PASSIVOS_MANUAIS);
  let totalPassivosManuais = 0;
  if (sheet && sheet.getLastRow() >= 2) {
      const data = sheet.getRange("C2:C" + sheet.getLastRow()).getValues();
      data.forEach(row => {
          totalPassivosManuais += parseBrazilianFloat(String(row[0]));
      });
  }

  const totalLiabilities = totalDividas + totalPassivosManuais;
  logToSheet(`[Patrimonio] Passivos Calculados: Dívidas Contas/Cartões(${totalDividas}) + Manuais(${totalPassivosManuais}) = ${totalLiabilities}`, "DEBUG");
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

  logToSheet(`[Patrimonio] Cálculo finalizado: Ativos=${totalAssets}, Passivos=${totalLiabilities}, PL=${netWorth}`, "INFO");

  return {
    assets: totalAssets,
    liabilities: totalLiabilities,
    netWorth: netWorth
  };
}

