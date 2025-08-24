/**
 * @file Budget.gs
 * @description Contém a lógica para gerir e atualizar a aba de Orçamento.
 */

/**
 * Atualiza os valores gastos na aba 'Orcamento' com base nos dados da aba 'Transacoes'.
 * Esta versão é mais eficiente e corrige a lógica de agregação de dados.
 */
function updateBudgetSpentValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const budgetSheet = ss.getSheetByName("Orcamento");
  const transacoesSheet = ss.getSheetByName("Transacoes");

  if (!budgetSheet || !transacoesSheet) {
    logToSheet("Aba 'Orcamento' ou 'Transacoes' não encontrada para atualização.", "WARN");
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('A atualizar os gastos do orçamento...', 'Orçamento', 10);

  const budgetData = budgetSheet.getRange("A2:F" + budgetSheet.getLastRow()).getValues();
  const transacoesData = transacoesSheet.getDataRange().getValues();
  const transacoesHeaders = transacoesData[0];
  const transacoesColMap = getColumnMap(transacoesHeaders); // Utiliza a função de Utils.gs

  const spentByCategoryMonth = {};

  // Itera sobre as transações UMA VEZ para agregar os gastos
  for (let i = 1; i < transacoesData.length; i++) {
    const row = transacoesData[i];
    const tipo = row[transacoesColMap["Tipo"]];
    const categoria = row[transacoesColMap["Categoria"]];
    const dataRelevante = parseData(row[transacoesColMap["Data de Vencimento"]]); // Utiliza a função de Utils.gs
    const valor = parseBrazilianFloat(String(row[transacoesColMap["Valor"]])); // Utiliza a função de Utils.gs

    if (tipo === "Despesa" && dataRelevante && categoria) {
      const monthYearKey = Utilities.formatDate(dataRelevante, Session.getScriptTimeZone(), "MMMM/yyyy").toLowerCase();
      const categoryKey = normalizarTexto(categoria); // Utiliza a função de Utils.gs
      const compositeKey = `${monthYearKey}|${categoryKey}`;

      if (!spentByCategoryMonth[compositeKey]) {
        spentByCategoryMonth[compositeKey] = 0;
      }
      spentByCategoryMonth[compositeKey] += valor;
    }
  }

  const newSpentValues = [];

  // Itera sobre o orçamento para encontrar os gastos correspondentes
  for (let i = 0; i < budgetData.length; i++) {
    const row = budgetData[i];
    const mesReferencia = (row[1] || "").toString().toLowerCase();
    const categoriaOrcamento = normalizarTexto(row[2]); // Utiliza a função de Utils.gs
    
    const compositeKey = `${mesReferencia}|${categoriaOrcamento}`;
    const spentValue = spentByCategoryMonth[compositeKey] || 0;
    
    newSpentValues.push([spentValue]);
  }

  // Escreve todos os novos valores na planilha de uma só vez
  if (newSpentValues.length > 0) {
    budgetSheet.getRange(2, 5, newSpentValues.length, 1).setValues(newSpentValues);
    logToSheet("Valores gastos na aba 'Orcamento' atualizados com sucesso via script.", "INFO");
    SpreadsheetApp.getActiveSpreadsheet().toast('Orçamento atualizado!', 'Sucesso', 5);
  }
}


// ... (o seu ficheiro Budget.gs existente) ...

/**
 * NOVO: Busca os dados de progresso do orçamento para serem enviados pelo Telegram.
 * @param {number} mes O mês para o resumo (1-12).
 * @param {number} ano O ano para o resumo.
 * @returns {Array<Object>} Um array de objetos, cada um representando uma categoria do orçamento.
 */
function getBudgetProgressForTelegram(mes, ano) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const orcamentoSheet = ss.getSheetByName(SHEET_ORCAMENTO);
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);

    if (!orcamentoSheet || !transacoesSheet) return [];

    const dadosOrcamento = orcamentoSheet.getDataRange().getValues();
    const dadosTransacoes = transacoesSheet.getDataRange().getValues();
    
    // Reutiliza a função _getBudgetProgress do seu Code_Dashboard.gs
    // Certifique-se de que a função _extractIconAndCleanCategory também esteja acessível
    return _getBudgetProgress(dadosOrcamento, dadosTransacoes, mes - 1, ano, {});
}


/**
 * NOVO: Busca os dados de progresso das metas de poupança da nova aba "Metas".
 * @returns {Array<Object>} Um array de objetos, cada um representando uma meta de poupança.
 */
function getGoalsStatusForTelegram() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const metasSheet = ss.getSheetByName(SHEET_METAS); // A nossa nova aba!

    if (!metasSheet || metasSheet.getLastRow() < 2) {
        return [];
    }

    const dadosMetas = metasSheet.getRange("A2:C" + metasSheet.getLastRow()).getValues();
    
    const status = [];
    dadosMetas.forEach(row => {
        const nome = row[0];
        const objetivo = parseBrazilianFloat(String(row[1] || '0'));
        const salvo = parseBrazilianFloat(String(row[2] || '0'));

        if (nome && objetivo > 0) {
            const percentage = (salvo / objetivo) * 100;
            status.push({
                nome: nome,
                objetivo: objetivo,
                salvo: salvo,
                percentage: percentage
            });
        }
    });

    return status;
}
