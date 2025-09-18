/**
 * @file Investimentos.gs
 * @description Contém a lógica para gerir a carteira de investimentos,
 * incluindo a criação de transações de fluxo de caixa correspondentes.
 */

const SHEET_INVESTIMENTOS = "Investimentos";

/**
 * ATUALIZADO: Lida com a compra de um ativo, recebendo os dados já analisados.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} ticker O código do ativo (ex: ITSA4).
 * @param {number} quantidade A quantidade de ativos comprados.
 * @param {number} preco O preço unitário de compra.
 * @param {string} nomeCorretora O nome da conta/corretora de onde o dinheiro saiu.
 * @param {string} usuario O nome do usuário que fez a compra.
 */
function handleComprarAtivo(chatId, ticker, quantidade, preco, nomeCorretora, usuario) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Validação dos dados de entrada
    if (!ticker || isNaN(quantidade) || quantidade <= 0 || isNaN(preco) || preco < 0 || !nomeCorretora) {
      enviarMensagemTelegram(chatId, "❌ Dados inválidos para a compra. Verifique o ticker, quantidade, preço e corretora.");
      return;
    }

    const valorTotal = quantidade * preco;
    const tickerUpper = ticker.toUpperCase();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const investimentosSheet = ss.getSheetByName(SHEET_INVESTIMENTOS);
    const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);

    const contaOrigemInfo = obterInformacoesDaConta(nomeCorretora, dadosContas);
    if (!contaOrigemInfo) {
      enviarMensagemTelegram(chatId, `❌ Conta de origem "${escapeMarkdown(nomeCorretora)}" não encontrada.`);
      return;
    }

    // Adiciona a transação de DESPESA
    registrarTransacaoNaPlanilha(new Date(), `Compra de ${quantidade} ${tickerUpper}`, '📈 Investimentos / Futuro', 'Compra de Ativo', 'Despesa', valorTotal, 'Transferência', contaOrigemInfo.nomeOriginal, 1, 1, new Date(), usuario, 'Ativo', Utilities.getUuid(), new Date());

    // Adiciona/Atualiza o ativo na aba Investimentos
    const dadosInvestimentos = investimentosSheet.getDataRange().getValues();
    let ativoRow = -1;
    for (let i = 1; i < dadosInvestimentos.length; i++) {
        // Agora procura pelo ticker com ou sem .SA para garantir a correspondência
        const tickerPlanilha = dadosInvestimentos[i][0].toUpperCase().replace(".SA", "");
        if (tickerPlanilha === tickerUpper) {
            ativoRow = i + 1;
            break;
        }
    }

    if (ativoRow > -1) {
        // Atualiza ativo existente
        const qtdAtual = parseFloat(investimentosSheet.getRange(ativoRow, 3).getValue()) || 0;
        const valorInvestidoAtual = parseBrazilianFloat(String(dadosInvestimentos[ativoRow-1][4]));
        
        const novaQtd = qtdAtual + quantidade;
        const novoValorInvestido = valorInvestidoAtual + valorTotal;
        const novoPrecoMedio = novoValorInvestido / novaQtd;

        investimentosSheet.getRange(ativoRow, 3).setValue(novaQtd);
        investimentosSheet.getRange(ativoRow, 4).setValue(novoPrecoMedio);
    } else {
        // --- INÍCIO DA CORREÇÃO ---
        // Adiciona o sufixo ".SA" ao ticker que será inserido na planilha
        const tickerComSufixo = `${tickerUpper}.SA`;
        // --- FIM DA CORREÇÃO ---

        // Adiciona novo ativo
        const proximaLinha = investimentosSheet.getLastRow() + 1;
        
        // ATUALIZADO para usar a nova variável na fórmula GOOGLEFINANCE
        // A fórmula agora refere-se diretamente à célula do ticker na coluna A
        investimentosSheet.appendRow([
          tickerComSufixo, // Coluna A: Ticker com ".SA"
          'Ação/FII',      // Coluna B
          quantidade,      // Coluna C
          preco,           // Coluna D
          `=C${proximaLinha}*D${proximaLinha}`, // Coluna E: Valor Investido
          `=GOOGLEFINANCE(A${proximaLinha})`,  // Coluna F: Preço Atual (referencia a Coluna A)
          `=C${proximaLinha}*F${proximaLinha}`, // Coluna G: Valor Atual
          `=G${proximaLinha}-E${proximaLinha}`, // Coluna H: Lucro/Prejuízo
          'Aberta'         // Coluna I
        ]);
    }

    atualizarSaldosDasContas();
    enviarMensagemTelegram(chatId, `✅ Compra de *${quantidade} ${escapeMarkdown(tickerUpper)}* registada com sucesso! O valor de ${formatCurrency(valorTotal)} foi debitado de ${contaOrigemInfo.nomeOriginal}.`);

  } catch (e) {
    handleError(e, "handleComprarAtivo", chatId);
  } finally {
    lock.releaseLock();
  }
}

/**
 * ATUALIZADO: Lida com a venda de um ativo, recebendo os dados já analisados.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} ticker O código do ativo (ex: ITSA4).
 * @param {number} quantidade A quantidade de ativos vendidos.
 * @param {number} preco O preço unitário de venda.
 * @param {string} nomeContaDestino O nome da conta/corretora para onde o dinheiro foi.
 * @param {string} usuario O nome do usuário que fez a venda.
 */
function handleVenderAtivo(chatId, ticker, quantidade, preco, nomeContaDestino, usuario) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Validação dos dados
    if (!ticker || isNaN(quantidade) || quantidade <= 0 || isNaN(preco) || preco <= 0 || !nomeContaDestino) {
      enviarMensagemTelegram(chatId, "❌ Dados inválidos. Verifique o ticker, quantidade, preço e conta de destino.");
      return;
    }

    const valorTotal = quantidade * preco;
    const tickerUpper = ticker.toUpperCase();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const investimentosSheet = ss.getSheetByName(SHEET_INVESTIMENTOS);
    const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);
    
    const contaDestinoInfo = obterInformacoesDaConta(nomeContaDestino, dadosContas);
    if (!contaDestinoInfo) {
      enviarMensagemTelegram(chatId, `❌ Conta de destino "${escapeMarkdown(nomeContaDestino)}" não encontrada.`);
      return;
    }

    const dadosInvestimentos = investimentosSheet.getDataRange().getValues();
    let ativoRow = -1;
    for (let i = 1; i < dadosInvestimentos.length; i++) {
        const tickerPlanilha = dadosInvestimentos[i][0].toUpperCase().replace(".SA", "");
        if (tickerPlanilha === tickerUpper) {
            ativoRow = i + 1;
            break;
        }
    }

    if (ativoRow === -1) {
        enviarMensagemTelegram(chatId, `❌ Ativo ${tickerUpper} não encontrado na sua carteira.`);
        return;
    }
    
    const qtdAtual = parseFloat(investimentosSheet.getRange(ativoRow, 3).getValue()) || 0;
    if (quantidade > qtdAtual) {
        enviarMensagemTelegram(chatId, `❌ Você não pode vender ${quantidade} de ${tickerUpper}, pois só possui ${qtdAtual}.`);
        return;
    }

    // Adiciona a transação de RECEITA
    registrarTransacaoNaPlanilha(new Date(), `Venda de ${quantidade} ${tickerUpper}`, '💸 Renda Extra e Investimentos', 'Venda de Ativo', 'Receita', valorTotal, 'Transferência', contaDestinoInfo.nomeOriginal, 1, 1, new Date(), usuario, 'Ativo', Utilities.getUuid(), new Date());
    
    // Atualiza a quantidade na aba Investimentos
    const novaQtd = qtdAtual - quantidade;
    investimentosSheet.getRange(ativoRow, 3).setValue(novaQtd);

    if (novaQtd === 0) {
        investimentosSheet.getRange(ativoRow, 9).setValue('Fechada'); // Coluna J: Status
    }

    atualizarSaldosDasContas();
    enviarMensagemTelegram(chatId, `✅ Venda de *${quantidade} ${escapeMarkdown(tickerUpper)}* registada com sucesso! O valor de ${formatCurrency(valorTotal)} foi creditado em ${contaDestinoInfo.nomeOriginal}.`);

  } catch (e) {
    handleError(e, "handleVenderAtivo", chatId);
  } finally {
    lock.releaseLock();
  }
}


/**
 * @private
 * Lê e processa os dados da aba "Investimentos".
 * @returns {Array<Object>} Um array de objetos, onde cada objeto representa um ativo.
 */
function _getInvestmentsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INVESTIMENTOS);
  
  if (!sheet || sheet.getLastRow() < 2) {
    logToSheet("[Investimentos] Aba 'Investimentos' não encontrada ou vazia.", "INFO");
    return [];
  }

  // --- INÍCIO DA CORREÇÃO ---
  // A leitura dos dados agora vai até a coluna J (índice 9) para incluir a coluna "Status".
  const data = sheet.getRange("A2:J" + sheet.getLastRow()).getValues();
  const investments = [];

  data.forEach(row => {
    const ativo = row[0]; // Coluna A (índice 0)
    // A coluna "Status" está na posição 9 do array (correspondente à coluna J)
    const status = row[9]; 

    if (ativo && normalizarTexto(status) === 'aberta') {
      investments.push({
        ativo: ativo,
        tipo: row[1],
        quantidade: parseFloat(row[2]) || 0,
        precoMedio: parseBrazilianFloat(String(row[3])),
        valorInvestido: parseBrazilianFloat(String(row[4])),
        precoAtual: parseBrazilianFloat(String(row[5])),
        valorAtual: parseBrazilianFloat(String(row[6])),
        lucroPrejuizo: parseBrazilianFloat(String(row[7]))
      });
    }
  });

  return investments;
  // --- FIM DA CORREÇÃO ---
}

/**
 * Calcula o valor total atualizado de todos os investimentos.
 * @returns {number} O valor total de todos os ativos de investimento.
 */
function getTotalInvestmentsValue() {
    const investments = _getInvestmentsData();
    if (investments.length === 0) {
        return 0;
    }
    const totalValue = investments.reduce((sum, asset) => sum + asset.valorAtual, 0);
    return totalValue;
}


// ===================================================================================
// ### INÍCIO DA NOVA FUNÇÃO PARA REGISTAR PROVENTOS ###
// ===================================================================================

/**
 * Regista um provento (dividendo, jcp, etc.), atualizando a planilha de investimentos e
 * criando a transação de receita correspondente.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} ticker O código do ativo que pagou o provento.
 * @param {number} valorProvento O valor total recebido.
 * @param {string} nomeContaDestino O nome da conta/corretora onde o valor foi creditado.
 * @param {string} usuario O nome do usuário.
 */
function registrarProvento(chatId, ticker, valorProvento, nomeContaDestino, usuario) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!ticker || isNaN(valorProvento) || valorProvento <= 0 || !nomeContaDestino) {
      enviarMensagemTelegram(chatId, "❌ Dados inválidos para registar o provento. Verifique o ticker, o valor e a conta de destino.");
      return;
    }

    const tickerUpper = ticker.toUpperCase();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const investimentosSheet = ss.getSheetByName(SHEET_INVESTIMENTOS);
    const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);

    // Verifica se a conta de destino existe
    const contaDestinoInfo = obterInformacoesDaConta(nomeContaDestino, dadosContas);
    if (!contaDestinoInfo) {
      enviarMensagemTelegram(chatId, `❌ Conta de destino "${escapeMarkdown(nomeContaDestino)}" não encontrada.`);
      return;
    }

    // Encontra o ativo e a coluna de proventos
    const dadosInvestimentos = investimentosSheet.getDataRange().getValues();
    const headers = dadosInvestimentos[0];
    const colTicker = headers.indexOf('Ativo');
    const colProventos = headers.indexOf('Total de Proventos');

    if (colProventos === -1) {
        throw new Error("A coluna 'Total de Proventos' não foi encontrada na aba 'Investimentos'. Por favor, adicione-a.");
    }

    let ativoRow = -1;
    for (let i = 1; i < dadosInvestimentos.length; i++) {
        const tickerPlanilha = dadosInvestimentos[i][colTicker].toUpperCase().replace(".SA", "");
        if (tickerPlanilha === tickerUpper) {
            ativoRow = i + 1;
            break;
        }
    }

    if (ativoRow === -1) {
        enviarMensagemTelegram(chatId, `❌ Ativo *${escapeMarkdown(tickerUpper)}* não encontrado na sua carteira. Registe primeiro uma compra para ele.`);
        return;
    }

    // Atualiza o valor na coluna de proventos
    const proventosRange = investimentosSheet.getRange(ativoRow, colProventos + 1);
    const proventosAtuais = parseBrazilianFloat(String(proventosRange.getValue()));
    proventosRange.setValue(proventosAtuais + valorProvento);

    // Regista a transação de RECEITA
    registrarTransacaoNaPlanilha(
      new Date(),
      `Proventos de ${tickerUpper}`,
      '💸 Renda Extra e Investimentos',
      'Proventos',
      'Receita',
      valorProvento,
      'Transferência',
      contaDestinoInfo.nomeOriginal,
      1, 1, new Date(),
      usuario,
      'Ativo',
      Utilities.getUuid(),
      new Date()
    );

    // Ajusta o saldo da conta de destino
    const contasSheet = ss.getSheetByName(SHEET_CONTAS);
    ajustarSaldoIncrementalmente(contasSheet, contaDestinoInfo.nomeOriginal, valorProvento);

    enviarMensagemTelegram(chatId, `✅ Provento de ${formatCurrency(valorProvento)} de *${escapeMarkdown(tickerUpper)}* registado com sucesso na conta ${contaDestinoInfo.nomeOriginal}!`);

  } catch (e) {
    handleError(e, "registrarProvento", chatId);
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

// ===================================================================================
// ### FIM DA NOVA FUNÇÃO PARA REGISTAR PROVENTOS ###
// ===================================================================================

