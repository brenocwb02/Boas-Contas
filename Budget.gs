/**
 * @file Budget.gs
 * @description Contém a lógica para gerir e atualizar a aba de Orçamento.
 */

/**
 * **FUNÇÃO ATUALIZADA E CORRIGIDA**
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
  const transacoesColMap = getColumnMap(transacoesHeaders);

  const spentByCategoryMonth = {}; // Objeto para armazenar os gastos agregados

  // Itera sobre as transações UMA VEZ para agregar os gastos
  for (let i = 1; i < transacoesData.length; i++) {
    const row = transacoesData[i];
    const tipo = row[transacoesColMap["Tipo"]];
    const categoria = row[transacoesColMap["Categoria"]];
    // **CORREÇÃO:** Usa a Data de Vencimento para despesas, que é o correto para o fluxo de caixa
    const dataRelevante = parseData(row[transacoesColMap["Data de Vencimento"]]);
    const valor = parseBrazilianFloat(String(row[transacoesColMap["Valor"]]));

    if (tipo === "Despesa" && dataRelevante && categoria) {
      // Cria uma chave única combinando mês/ano e categoria (ex: "julho/2025|vida espiritual")
      const monthYearKey = Utilities.formatDate(dataRelevante, Session.getScriptTimeZone(), "MMMM/yyyy").toLowerCase();
      const categoryKey = normalizarTexto(categoria);
      const compositeKey = `${monthYearKey}|${categoryKey}`;

      if (!spentByCategoryMonth[compositeKey]) {
        spentByCategoryMonth[compositeKey] = 0;
      }
      spentByCategoryMonth[compositeKey] += valor;
    }
  }

  const newSpentValues = []; // Array para armazenar os novos valores a serem escritos na planilha

  // Itera sobre o orçamento para encontrar os gastos correspondentes
  for (let i = 0; i < budgetData.length; i++) {
    const row = budgetData[i];
    const mesReferencia = (row[1] || "").toString().toLowerCase(); // Coluna B: Mes referencia
    const categoriaOrcamento = normalizarTexto(row[2]); // Coluna C: Categoria
    
    const compositeKey = `${mesReferencia}|${categoriaOrcamento}`;
    const spentValue = spentByCategoryMonth[compositeKey] || 0;
    
    newSpentValues.push([spentValue]); // Adiciona o valor ao array
  }

  // Escreve todos os novos valores na planilha de uma só vez (muito mais eficiente)
  if (newSpentValues.length > 0) {
    // Coluna E (índice 4) é a 'Valor Gasto'
    budgetSheet.getRange(2, 5, newSpentValues.length, 1).setValues(newSpentValues);
    logToSheet("Valores gastos na aba 'Orcamento' atualizados com sucesso via script.", "INFO");
    SpreadsheetApp.getActiveSpreadsheet().toast('Orçamento atualizado!', 'Sucesso', 5);
  }
}
