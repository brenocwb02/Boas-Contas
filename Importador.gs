/**
 * @file Importador.gs
 * @description Contém a lógica para importar, fazer o parse e conciliar extratos bancários (CSV/OFX).
 */

/**
 * Ponto de entrada principal para processar um arquivo de extrato.
 * @param {string} fileContent O conteúdo do arquivo como uma string de texto.
 * @param {string} fileName O nome do arquivo para determinar o tipo (csv ou ofx).
 * @returns {object} Um objeto com as listas de transações: 'notFound', 'reconciled', 'duplicates'.
 */
function _processBankStatement(fileContent, fileName) {
  logToSheet(`[Importador] Iniciando processamento para o arquivo: ${fileName}`, "INFO");
  try {
    const fileType = fileName.split('.').pop().toLowerCase();
    let importedTransactions = [];

    if (fileType === 'csv') {
      importedTransactions = parseCsv(fileContent);
    } else if (fileType === 'ofx') {
      importedTransactions = parseOfx(fileContent);
    } else {
      throw new Error("Formato de arquivo não suportado. Use .csv ou .ofx.");
    }
    
    logToSheet(`[Importador] ${importedTransactions.length} transações extraídas do arquivo.`, "INFO");

    if (importedTransactions.length === 0) {
      throw new Error("Nenhuma transação válida foi encontrada no arquivo. Verifique o conteúdo e o formato do arquivo.");
    }

    const result = reconcileTransactions(importedTransactions);
    logToSheet(`[Importador] Conciliação concluída. Não encontradas: ${result.notFound.length}, Conciliadas: ${result.reconciled.length}`, "INFO");
    return result;

  } catch (e) {
    logToSheet(`[Importador] ERRO ao processar extrato bancário: ${e.message}`, "ERROR");
    return { error: e.message };
  }
}

/**
 * Importa um lote de novas transações para a planilha.
 * @param {Array<object>} transactionsToImport Array de transações selecionadas pelo usuário.
 * @returns {object} Objeto de sucesso com os dados do dashboard atualizados.
 */
function _importNewTransactions(transactionsToImport) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TRANSACOES);
    if (!sheet) throw new Error("Planilha 'Transacoes' não encontrada.");

    const timezone = ss.getSpreadsheetTimeZone();
    const dadosPalavras = getSheetDataWithCache(SHEET_PALAVRAS_CHAVE, CACHE_KEY_PALAVRAS);
    
    const rowsToAdd = transactionsToImport.map(tx => {
      const { categoria, subcategoria } = extrairCategoriaSubcategoria(normalizarTexto(tx.description), tx.type, dadosPalavras);
      
      return [
        Utilities.formatDate(new Date(tx.date), timezone, "dd/MM/yyyy"),
        tx.description,
        categoria,
        subcategoria,
        tx.type,
        Math.abs(tx.value),
        '', // Metodo de Pagamento
        '', // Conta / Cartão
        1,  // Parcelas
        1,
        Utilities.formatDate(new Date(tx.date), timezone, "dd/MM/yyyy"), // Data Vencimento
        'Importado', // Usuario
        'Ativo',
        Utilities.getUuid(),
        new Date()
      ];
    });

    if (rowsToAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    }

    atualizarSaldosDasContas();
    
    const today = new Date();
    const dashboardData = getDashboardData(today.getMonth() + 1, today.getFullYear());
    
    return { 
      success: true, 
      message: `${transactionsToImport.length} transações importadas com sucesso.`,
      dashboardData: dashboardData 
    };

  } catch (e) {
    handleError(e, "importNewTransactions");
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Faz o parse de um conteúdo de texto em formato OFX.
 * @param {string} ofxContent O conteúdo do arquivo OFX.
 * @returns {Array<object>} Um array de transações.
 */
function parseOfx(ofxContent) {
  const transactions = [];
  const transactionBlocks = ofxContent.match(/<STMTTRN>([\s\S]*?)<\/STMTTRN>/g);

  if (!transactionBlocks) {
    throw new Error("Nenhum bloco de transação (<STMTTRN>) encontrado no arquivo OFX.");
  }

  transactionBlocks.forEach(block => {
    const getValue = (tag) => {
      const match = block.match(new RegExp(`<${tag}>([^<]*)`));
      return match ? match[1].trim() : null;
    };

    const dateStr = getValue('DTPOSTED');
    const amountStr = getValue('TRNAMT');
    const description = getValue('MEMO');
    
    if (dateStr && amountStr && description) {
      const year = parseInt(dateStr.substring(0, 4), 10);
      const month = parseInt(dateStr.substring(4, 6), 10) - 1;
      const day = parseInt(dateStr.substring(6, 8), 10);
      const date = new Date(year, month, day);
      const value = parseFloat(amountStr);

      if (date && !isNaN(value) && value !== 0) {
        transactions.push({
          date: date,
          description: description,
          value: value,
          type: value > 0 ? 'Receita' : 'Despesa'
        });
      }
    }
  });

  return transactions;
}

/**
 * Faz o parse de um conteúdo de texto em formato CSV.
 * @param {string} csvContent O conteúdo do arquivo CSV.
 * @returns {Array<object>} Um array de transações.
 */
function parseCsv(csvContent) {
  const transactions = [];
  const lines = csvContent.split(/\r\n|\n/);

  if (lines.length < 2) {
    throw new Error("O arquivo CSV está vazio ou não contém transações.");
  }

  const separator = lines[0].includes(';') ? ';' : ',';
  const headers = lines[0].split(separator).map(h => normalizarTexto(h.replace(/"/g, '')));
  
  const dateIndex = headers.findIndex(h => h.includes('data'));
  const descriptionIndex = headers.findIndex(h => h.includes('descricao') || h.includes('historico'));
  const valueIndex = headers.findIndex(h => h.includes('valor') || h.includes('montante'));

  if (dateIndex === -1 || descriptionIndex === -1 || valueIndex === -1) {
    throw new Error("Não foi possível encontrar as colunas 'Data', 'Descricao' e 'Valor' no cabeçalho do arquivo CSV.");
  }

  for (let i = 1; i < lines.length; i++) {
    const line = lines[i];
    if (!line.trim()) continue;

    const parts = line.split(separator);
    
    const dateStr = parts[dateIndex] ? parts[dateIndex].trim() : '';
    const descStr = parts[descriptionIndex] ? parts[descriptionIndex].trim().replace(/"/g, '') : '';
    const valueStr = parts[valueIndex] ? parts[valueIndex].trim() : '';

    const date = parseData(dateStr);
    const description = descStr;
    const value = parseBrazilianFloat(valueStr);

    if (date && description && !isNaN(value) && value !== 0) {
      transactions.push({
        date: date,
        description: description,
        value: value,
        type: value > 0 ? 'Receita' : 'Despesa'
      });
    } else {
       logToSheet(`Linha ${i + 1} do CSV ignorada: dados inválidos. Data: ${dateStr}, Desc: ${descStr}, Valor: ${valueStr}`, "WARN");
    }
  }
  return transactions;
}

/**
 * Compara as transações importadas com as existentes na planilha.
 * @param {Array<object>} importedTransactions Transações lidas do arquivo de extrato.
 * @returns {object} Um objeto contendo as listas de transações conciliadas e não encontradas.
 */
function reconcileTransactions(importedTransactions) {
  logToSheet(`[Importador] Iniciando conciliação de ${importedTransactions.length} transações importadas.`, "INFO");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TRANSACOES);
  if (!sheet) {
      logToSheet("[Importador] Aba 'Transacoes' não encontrada para conciliação.", "ERROR");
      return { notFound: importedTransactions, reconciled: [], duplicates: [] };
  }
  const existingTransactions = sheet.getDataRange().getValues();
  const headers = existingTransactions[0];
  const colMap = getColumnMap(headers);

  const notFound = [];
  const reconciled = [];

  // CORREÇÃO: Converte as datas para strings antes de retornar, para evitar problemas de serialização.
  const toJSONSafe = (tx) => {
    return {
      ...tx,
      date: tx.date.toISOString() // Converte para string
    };
  };

  importedTransactions.forEach(importedTx => {
    let foundMatch = false;
    const importedDate = new Date(importedTx.date);
    
    for (let i = 1; i < existingTransactions.length; i++) {
      const existingTx = existingTransactions[i];
      const existingDate = parseData(existingTx[colMap['Data']]);
      
      if (!existingDate) continue; 
      
      const existingValue = parseBrazilianFloat(String(existingTx[colMap['Valor']]));
      const existingType = existingTx[colMap['Tipo']];

      const isDateMatch = (d1, d2) => d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth() && d1.getDate() === d2.getDate();

      const isValueMatch = Math.abs(existingValue - Math.abs(importedTx.value)) < 0.01;
      const isTypeMatch = existingType === importedTx.type;

      if (isDateMatch(existingDate, importedDate) && isValueMatch && isTypeMatch) {
        const existingDescription = normalizarTexto(existingTx[colMap['Descricao']]);
        const importedDescription = normalizarTexto(importedTx.description);
        const similarity = calculateSimilarity(existingDescription, importedDescription);

        if (similarity > 0.6) {
          logToSheet(`[Importador] MATCH ENCONTRADO: Importado "${importedDescription}" com existente "${existingDescription}" (Similaridade: ${similarity.toFixed(2)})`, "DEBUG");
          foundMatch = true;
          reconciled.push({ imported: toJSONSafe(importedTx), existing: existingTx });
          break; 
        }
      }
    }
    if (!foundMatch) {
      notFound.push(toJSONSafe(importedTx));
    }
  });

  return { notFound, reconciled, duplicates: [] };
}
