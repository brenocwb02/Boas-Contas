/**
 * @file SheetUtils.gs
 * @description Funções para interação específica com o Google Planilhas (leitura, escrita, busca).
 */

// --- INÍCIO DA MELHORIA DE PERFORMANCE DE LOGS ---
// Buffer global para agrupar os logs antes de os escrever na planilha.
let logBuffer = [];
// --- FIM DA MELHORIA ---


/**
 * Obtém o nível de log configurado na aba "Configuracoes" da planilha.
 * Este valor define o quão detalhados serão os logs escritos na aba "Logs_Sistema".
 * @returns {string} O nível de log (DEBUG, INFO, WARN, ERROR) ou "INFO" se a configuração não for encontrada ou for inválida.
 */
function getLogLevelConfig() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SHEET_CONFIGURACOES);
    if (!configSheet) {
      Logger.log("ERRO: Aba 'Configuracoes' não encontrada para obter nível de log. Usando nível padrão 'INFO'.");
      return "INFO"; // Retorno padrão se a aba estiver faltando
    }
    const configData = configSheet.getDataRange().getValues();
    for (let i = 0; i < configData.length; i++) {
      // Procura pela chave "LOG_LEVEL" na primeira coluna
      if ((configData[i][0] || "").toString().toUpperCase().trim() === "LOG_LEVEL") {
        const level = (configData[i][1] || "").toString().toUpperCase().trim();
        // Valida se o nível lido é um dos níveis esperados
        if (["DEBUG", "INFO", "WARN", "ERROR", "NONE"].includes(level)) {
          Logger.log(`Nível de log configurado na planilha: ${level}`);
          return level; // Retorna o nível configurado
        }
      }
    }
    // Se "LOG_LEVEL" não for encontrado ou for inválido, retorna o padrão
    Logger.log("Configuracao 'LOG_LEVEL' não encontrada ou inválida na aba 'Configuracoes'. Usando nível padrão 'INFO'.");
    return "INFO";
  } catch (e) {
    Logger.log(`ERRO ao tentar obter nível de log da planilha: ${e.message}. Usando nível padrão 'INFO'.`);
    return "INFO"; // Em caso de erro, retorna o padrão
  }
}

/**
 * ATUALIZADO: Função de log centralizada que agora agrupa os logs em memória.
 * Em vez de escrever na planilha a cada chamada, adiciona a mensagem a um buffer.
 * O console.log é mantido para depuração em tempo real.
 * @param {string} message A mensagem a ser logada.
 * @param {string} level O nível do log (DEBUG, INFO, WARN, ERROR).
 */
function logToSheet(message, level) {
  // Garante que currentLogLevel foi inicializado.
  if (typeof currentLogLevel === 'undefined' || currentLogLevel === null) {
    currentLogLevel = getLogLevelConfig();
  }

  const numericCurrentLevel = LOG_LEVEL_MAP[currentLogLevel] || LOG_LEVEL_MAP["INFO"];
  const numericMessageLevel = LOG_LEVEL_MAP[level] || LOG_LEVEL_MAP["INFO"];

  // Loga no Cloud Logs (console) para depuração imediata
  if (numericMessageLevel >= numericCurrentLevel) {
    if (level === "ERROR") console.error(`[${level}] ${message}`);
    else if (level === "WARN") console.warn(`[${level}] ${message}`);
    else if (level === "INFO") console.info(`[${level}] ${message}`);
    else if (level === "DEBUG") console.log(`[${level}] ${message}`);
    Logger.log(`[${level}] ${message}`);
  }

  // Adiciona ao buffer para escrita em lote no final da execução
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
  logBuffer.push([timestamp, level, message]);
}


/**
 * NOVO: Escreve todos os logs agrupados do buffer para a planilha de uma só vez.
 * Esta função deve ser chamada no final da execução do script (ex: no `doPost`).
 */
function flushLogs() {
  if (logBuffer.length === 0) return;

  // Filtra os logs que devem ser escritos com base no nível de configuração
  const numericCurrentLevel = LOG_LEVEL_MAP[currentLogLevel] || LOG_LEVEL_MAP["INFO"];
  const logsToWrite = logBuffer.filter(row => {
    const numericMessageLevel = LOG_LEVEL_MAP[row[1]] || LOG_LEVEL_MAP["INFO"];
    return numericMessageLevel >= numericCurrentLevel;
  });

  if (logsToWrite.length === 0) {
    logBuffer = []; // Limpa o buffer mesmo que nada seja escrito
    return;
  }
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(SHEET_LOGS_SISTEMA);
    if (logSheet) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, logsToWrite.length, logsToWrite[0].length)
              .setValues(logsToWrite);
    }
  } catch (e) {
    console.error(`ERRO CRÍTICO ao descarregar logs: ${e.message}`);
    Logger.log(`ERRO CRÍTICO ao descarregar logs: ${e.message}`);
  } finally {
    // Limpa o buffer para a próxima execução
    logBuffer = [];
  }
}

/**
 * Registra uma nova linha na aba "Transacoes".
 * @param {Date} data Data da transação.
 * @param {string} descricao Descrição da transação.
 * @param {string} categoria Categoria da transação.
 * @param {string} subcategoria Subcategoria da transação.
 * @param {string} tipo Tipo da transação (Receita, Despesa, Transferência).
 * @param {number} valor Valor da transação.
 * @param {string} metodoPagamento Método de pagamento.
 * @param {string} conta Conta ou cartão utilizado.
 * @param {number} parcelasTotais Número total de parcelas.
 * @param {number} parcelaAtual Parcela atual (se parcelado).
 * @param {Date} dataVencimento Data de vencimento (para cartões de crédito).
 * @param {string} usuario Nome do usuário que registrou.
 * @param {string} status Status da transação (Ativo, Pago, Cancelado).
 * @param {string} idTransacao ID único da transação.
 * @param {Date} dataRegistro Timestamp do registro.
 */
function registrarTransacaoNaPlanilha(data, descricao, categoria, subcategoria, tipo, valor, metodoPagamento, conta, parcelasTotais, parcelaAtual, dataVencimento, usuario, status, idTransacao, dataRegistro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TRANSACOES);

  // Garante que a data de vencimento é um objeto Date válido
  const vencimentoFormatado = dataVencimento instanceof Date && !isNaN(dataVencimento) ? Utilities.formatDate(dataVencimento, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy") : "";

  // Formata a data para dd/MM/yyyy
  const dataFormatada = Utilities.formatDate(data, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy");
  const dataRegistroFormatada = Utilities.formatDate(dataRegistro, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");

  const newRow = [
    dataFormatada,
    descricao,
    categoria,
    subcategoria,
    tipo,
    valor,
    metodoPagamento,
    conta,
    parcelasTotais,
    parcelaAtual,
    vencimentoFormatado,
    usuario,
    status,
    idTransacao,
    dataRegistroFormatada
  ];

  sheet.appendRow(newRow);
  logToSheet(`Transação registrada na planilha: ${JSON.stringify(newRow)}`, "INFO");
}

/**
 * REMOVIDA: getTelegramToken
 * Esta função foi removida pois o token do Telegram agora é obtido diretamente das propriedades do script
 * através da função `getTelegramBotTokenFromProperties()` em `Constants.gs`.
 */

/**
 * Obtém o nome do usuário associado a um Chat ID na aba "Configuracoes".
 * @param {string} chatId O ID do chat do Telegram.
 * @param {Array<Array<any>>} config Os dados da aba "Configuracoes".
 * @returns {string} O nome do usuário ou "Desconhecido" se o Chat ID não for encontrado.
 */
function getUsuarioPorChatId(chatId, config) {
  for (let i = 0; i < config.length; i++) {
    const chave = config[i][0];
    const id = config[i][1];
    const nome = config[i][2];
    if (chave === "chatId" && id.toString() === chatId.toString()) {
      return nome ? nome.toString().trim() : "Usuário"; // Garante que retorna uma string válida
    }
  }
  return "Desconhecido";
}

/**
 * Obtém o Chat ID de um usuário específico a partir dos dados de configuração.
 * @param {Array<Array<any>>} config Os dados da aba "Configuracoes".
 * @param {string} usuario O nome do usuário.
 * @returns {string|null} O Chat ID do usuário ou null se não for encontrado.
 */
function getChatId(config, usuario) {
  for (let i = 0; i < config.length; i++) {
    if (config[i][2] && config[i][2].toString().trim() === usuario.trim()) {
      return config[i][1];
    }
  }
  return null;
}

/**
 * Obtém o nome do grupo associado a um Chat ID a partir dos dados de configuração.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {Array<Array<any>>} config Os dados da aba "Configuracoes".
 * @returns {string|null} O nome do grupo ou null se não for encontrado.
 */
function getGrupoPorChatId(chatId, config) {
  for (let i = 0; i < config.length; i++) {
    const chave = config[i][0];
    const id = config[i][1];
    const grupo = config[i][3]; // Coluna 4 (índice 3).
    if (chave === "chatId" && id.toString() === chatId.toString()) {
      return grupo || null;
    }
  }
  return null;
}

/**
 * Obtém o nome do grupo de um usuário específico a partir dos dados de configuração.
 * @param {string} usuario O nome do usuário.
 * @param {Array<Array<any>>} config Os dados da aba "Configuracoes".
 * @returns {string|null} O nome do grupo ou null se não for encontrado.
 */
function getGrupoPorChatIdByUsuario(usuario, config) {
  for (let i = 0; i < config.length; i++) {
    const nome = config[i][2];
    const grupo = config[i][3];
    if (nome && normalizarTexto(nome) === normalizarTexto(usuario)) {
      return grupo || null;
    }
  }
  return null;
}

/**
 * Obtém o tipo de uma conta (ex: "conta corrente", "cartão de crédito") a partir do seu nome.
 * @param {string} nomeConta O nome da conta.
 * @param {Array<Array<any>>} dadosContas Os dados da aba "Contas".
 * @returns {string|null} O tipo da conta em minúsculas ou null se não for encontrada.
 */
function getTipoDaConta(nomeConta, dadosContas) {
  const nomeContaNormalizado = normalizarTexto(nomeConta);
  for (let i = 1; i < dadosContas.length; i++) {
    const nomeNaPlanilha = normalizarTexto(dadosContas[i][0]);
    if (nomeNaPlanilha === nomeContaNormalizado) {
      return (dadosContas[i][1] || "").toLowerCase().trim(); // Retorna o tipo da conta (Coluna B).
    }
  }
  return null;
}

/**
 * Obtém informações detalhadas de uma conta ou cartão da aba "Contas".
 * Esta função é crucial para recuperar os atributos de uma conta (tipo, limite, vencimento, etc.).
 * @param {string} nomeBusca O nome da conta/cartão a ser buscado.
 * @param {Array<Array<any>>} dadosContas Os dados da aba "Contas" (passados para evitar múltiplas leituras).
 * @returns {Object|null} Um objeto com as informações da conta (nomeOriginal, nomeNormalizado, tipo, limite, vencimento, diaFechamento, tipoFechamento, contaPaiAgrupador) ou `null` se não encontrada.
 */
function obterInformacoesDaConta(nomeBusca, dadosContas) {
  logToSheet(`obterInformacoesDaConta chamada com nomeBusca: "${nomeBusca}"`, "DEBUG");
  // A aba 'Contas' já deve ter sido lida e passada como 'dadosContas'.
  if (!dadosContas || dadosContas.length === 0) {
    logToSheet("Dados da aba 'Contas' não fornecidos ou vazios para obterInformacoesDaConta.", "ERROR");
    return null;
  }
  
  const nomeBuscaNormalizado = normalizarTexto(nomeBusca);

  let bestMatch = null;
  let bestScore = -1; // Higher score is better.

  // Iterate through all accounts to find the best match
  for (let i = 1; i < dadosContas.length; i++) {
    const linha = dadosContas[i];
    const nomeNaPlanilha = (linha[0] || "").toString().trim();
    const tipoNaPlanilha = (linha[1] || "").toString().toLowerCase().trim();
    const nomeNormalizadoPlanilha = normalizarTexto(nomeNaPlanilha);
    const tipoNormalizadoPlanilha = normalizarTexto(tipoNaPlanilha);

    // --- CRÍTICO: SE ENCONTRAR UMA CORRESPONDÊNCIA EXATA, RETORNA IMEDIATAMENTE ---
    if (nomeNormalizadoPlanilha === nomeBuscaNormalizado) {
      logToSheet(`obterInformacoesDaConta: Correspondencia EXATA encontrada para "${nomeBusca}" em "${nomeNaPlanilha}". Retornando imediatamente.`, "DEBUG");
      return {
        nomeOriginal: nomeNaPlanilha,
        nomeNormalizado: nomeNormalizadoPlanilha,
        tipo: tipoNaPlanilha,
        limite: parseBrazilianFloat(String(linha[5] || '0')),
        vencimento: parseInt(linha[6]) || null,
        diaFechamento: parseInt(linha[9]) || null,
        tipoFechamento: (linha[10] || "").toString().trim(),
        contaPaiAgrupador: normalizarTexto((linha[12] || "").toString().trim()),
        pessoa: (linha[13] || "").toString().trim()
      };
    }

    let currentScore = 0;

    // Pontuação para correspondências parciais (se a busca contém o nome da planilha)
    if (nomeBuscaNormalizado.includes(nomeNormalizadoPlanilha)) {
      currentScore += nomeNormalizadoPlanilha.length * 10; // Maior peso para partes maiores
      logToSheet(`  - Score por inclusao (busca contem planilha): +${nomeNormalizadoPlanilha.length * 10} para "${nomeNaPlanilha}"`, "DEBUG");
    }
    // Pontuação para correspondências parciais (se o nome da planilha contém a busca)
    else if (nomeNormalizadoPlanilha.includes(nomeBuscaNormalizado)) {
      currentScore += nomeBuscaNormalizado.length * 8; // Um pouco menos de peso
      logToSheet(`  - Score por inclusao (planilha contem busca): +${nomeBuscaNormalizado.length * 8} para "${nomeNaPlanilha}"`, "DEBUG");
    }

    // Boost para cartões de crédito, especialmente se o termo de busca implica um cartão
    if (tipoNormalizadoPlanilha === "cartao de credito") {
        currentScore += 500; // Boost significativo para cartões de crédito
        logToSheet(`  - Score por ser cartao de credito: +500 para "${nomeNaPlanilha}"`, "DEBUG");

        if (nomeBuscaNormalizado.includes("cartao")) {
            currentScore += 200; // Boost ainda maior se o termo de busca mencionar "cartao"
            logToSheet(`  - Score por busca incluir "cartao": +200 para "${nomeNaPlanilha}"`, "DEBUG");
        }

        // NOVO: Boost para "Cartão [Nome do Banco]" quando a busca é apenas "[Nome do Banco]"
        // Ex: busca "inter", nome na planilha "cartao inter"
        const commonBankNames = ["inter", "nubank", "santander", "itau", "bradesco", "caixa", "original", "c6 bank", "picpay", "mercado pago", "sicoob", "sicredi", "banrisul", "neon", "next", "digio", "will bank", "bs2", "ame digital", "pagbank", "safra", "xp", "rico", "clear", "modalmais", "btg pactual", "creditas"];
        if (commonBankNames.includes(nomeBuscaNormalizado) && nomeNormalizadoPlanilha.startsWith("cartao " + nomeBuscaNormalizado)) {
            currentScore += 300; // Boost forte para essa correspondência específica
            logToSheet(`  - Score por "Cartao [Banco]" e busca "[Banco]": +300 para "${nomeNaPlanilha}"`, "DEBUG");
        }
        // NOVO: Boost se o nome da planilha é "Cartão [Banco]" e a busca é "Cartão [Banco]" (mesmo que não seja exato, mas muito próximo)
        else if (nomeBuscaNormalizado.startsWith("cartao ") && nomeNormalizadoPlanilha.startsWith("cartao ") && nomeBuscaNormalizado.includes(nomeNormalizadoPlanilha.replace("cartao ", ""))) {
            currentScore += 150; // Boost para correspondencia de cartao mais especifica
            logToSheet(`  - Score por "Cartao [Banco]" e busca "Cartao [Banco]" (parcial): +150 para "${nomeNaPlanilha}"`, "DEBUG");
        }
    }

    logToSheet(`obterInformacoesDaConta: Avaliando "${nomeNaPlanilha}" (Tipo: ${tipoNaPlanilha}) para busca "${nomeBusca}". Score: ${currentScore}, Melhor Score Atual: ${bestScore}`, "DEBUG");

    if (currentScore > bestScore) {
        bestScore = currentScore;
        bestMatch = linha;
        logToSheet(`obterInformacoesDaConta: Nova melhor correspondencia: "${nomeNaPlanilha}" (Score: ${bestScore})`, "DEBUG");
    }
  }

  // Se uma conta foi encontrada (exata ou flexível), formata o objeto de retorno.
  if (bestMatch) {
    const nomeConta = (bestMatch[0] || "").toString().trim();
    const nomeContaNormalizado = normalizarTexto(nomeConta);
    const tipoConta = (bestMatch[1] || "").toString().toLowerCase().trim();
    const limite = parseBrazilianFloat(String(bestMatch[5] || '0'));
    const vencimento = parseInt(bestMatch[6]) || null;
    const diaFechamento = parseInt(bestMatch[9]) || null;
    const tipoFechamento = (bestMatch[10] || "").toString().trim();
    const contaPaiAgrupador = (bestMatch[12] || "").toString().trim();
    const pessoa = (bestMatch[13] || "").toString().trim(); // Coluna N

    logToSheet(`obterInformacoesDaConta: Melhor correspondencia final para "${nomeBusca}": "${nomeConta}" (Tipo: ${tipoConta})`, "INFO");
    return {
      nomeOriginal: nomeConta,
      nomeNormalizado: nomeContaNormalizado,
      tipo: tipoConta,
      limite: limite,
      vencimento: vencimento,
      diaFechamento: diaFechamento,
      tipoFechamento: tipoFechamento,
      contaPaiAgrupador: normalizarTexto(contaPaiAgrupador),
      pessoa: pessoa
    };
  }

  logToSheet(`obterInformacoesDaConta: Nenhuma conta encontrada para "${nomeBusca}".`, "DEBUG");
  return null;
}

/**
 * Obtém as informações de uma conta a pagar da aba 'Contas_a_Pagar' pelo seu ID.
 * @param {string} idContaAPagar O ID único da conta a pagar.
 * @returns {Object|null} Um objeto com os detalhes da conta a pagar, incluindo a linha na planilha, ou null se não encontrada.
 */
function obterInformacoesDaContaAPagar(idContaAPagar) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);

  if (!contasAPagarSheet) {
    logToSheet("Aba 'Contas_a_Pagar' não encontrada para obterInformacoesDaContaAPagar.", "ERROR");
    return null;
  }

  const dados = contasAPagarSheet.getDataRange().getValues();
  const headers = dados[0]; // Primeira linha são os cabeçalhos

  const colID = headers.indexOf('ID');
  const colDescricao = headers.indexOf('Descricao');
  const colCategoria = headers.indexOf('Categoria');
  const colValor = headers.indexOf('Valor');
  const colDataVencimento = headers.indexOf('Data de Vencimento');
  const colStatus = headers.indexOf('Status');
  const colRecorrente = headers.indexOf('Recorrente');
  const colContaSugeria = headers.indexOf('Conta de Pagamento Sugerida');
  const colObservacoes = headers.indexOf('Observacoes');
  const colIDTransacaoVinculada = headers.indexOf('ID Transacao Vinculada');

  if (colID === -1 || colDescricao === -1 || colValor === -1 || colStatus === -1) {
    logToSheet("Colunas essenciais (ID, Descricao, Valor, Status) não encontradas na aba 'Contas_a_Pagar'.", "ERROR");
    return null;
  }

  for (let i = 1; i < dados.length; i++) {
    const row = dados[i];
    if (row[colID] === idContaAPagar) {
      let valorConta = parseBrazilianFloat(String(row[colValor]));
      
      return {
        id: row[colID],
        descricao: (row[colDescricao] || "").toString().trim(),
        categoria: (row[colCategoria] || "").toString().trim(),
        valor: valorConta,
        dataVencimento: parseData(row[colDataVencimento]),
        status: (row[colStatus] || "").toString().trim(),
        recorrente: (row[colRecorrente] || "").toString().trim(),
        contaDePagamentoSugeria: (row[colContaSugeria] || "").toString().trim(),
        observacoes: (row[colObservacoes] || "").toString().trim(),
        idTransacaoVinculada: (row[colIDTransacaoVinculada] || "").toString().trim(),
        linha: i + 1, // Retorna a linha real na planilha (base 1)
        headers: headers // Retorna os cabeçalhos para referencia de coluna
      };
    }
  }
  return null;
}
/**
 * Calcula o fluxo de caixa (receitas e despesas totais) para um determinado mês e ano.
 * @param {Array<Array<any>>} transacoesRaw Os dados brutos das transações.
 * @param {number} mes O mês para filtrar (1-12).
 * @param {number} ano O ano para filtrar.
 * @param {Array<Array<any>>} dadosContas Os dados da aba 'Contas'.
 * @returns {Object} Um objeto contendo receitas totais, despesas totais e saldo líquido.
 */
function calcularFluxoDeCaixa(transacoesRaw, mes, ano, dadosContas) {
  let receitasTotais = 0;
  let despesasTotais = 0;

  // Mapeia os nomes das contas para seus tipos (Débito/Crédito) e IDs
  const infoContas = {};
  for (let i = 1; i < dadosContas.length; i++) { // Ignora o cabeçalho
    const row = dadosContas[i];
    const nomeConta = row[0]; // Coluna A: Nome da Conta
    const tipoConta = row[1]; // Coluna B: Tipo (Débito/Crédito)
    const idConta = row[2]; // Coluna C: ID (para cartões)
    infoContas[nomeConta] = { tipo: tipoConta, id: idConta };
  }

  transacoesRaw.forEach(row => {
    const data = new Date(row[0]);
    const valor = parseFloat(row[2]);
    const tipo = row[3];
    const conta = row[4];

    if (data.getMonth() + 1 === mes && data.getFullYear() === ano) {
      // Ignora pagamentos de fatura e transferências para o cálculo do fluxo de caixa
      if (tipo === 'Receita') {
        receitasTotais += valor;
      } else if (tipo === 'Despesa') {
        const contaInfo = infoContas[conta];
        // Se for uma despesa e não for um pagamento de fatura ou transferência
        if (contaInfo && contaInfo.tipo !== 'Cartão de Crédito' && tipo !== 'Transferência') {
          despesasTotais += valor;
        } else if (!contaInfo && tipo !== 'Transferência') { // Para despesas que não estão mapeadas como cartão de crédito
          despesasTotais += valor;
        }
      }
    }
  });

  const saldoLiquido = receitasTotais - despesasTotais;

  return {
    receitasTotais: receitasTotais,
    despesasTotais: despesasTotais,
    saldoLiquido: saldoLiquido
  };
}
/**
 * Calcula os gastos brutos de cartão de crédito para um determinado mês e ano.
 * @param {Array<Array<any>>} transacoesRaw Os dados brutos das transações.
 * @param {number} mes O mês para filtrar (1-12).
 * @param {number} ano O ano para filtrar.
 * @param {Array<Array<any>>} dadosContas Os dados da aba 'Contas'.
 * @returns {Object} Um objeto com os gastos por cartão de crédito e seus limites.
 */
function calcularGastosCartaoCredito(transacoesRaw, mes, ano, dadosContas) {
  const gastosCartao = {};
  const limitesCartao = {};
  const datasVencimento = {};

  // Mapeia os IDs dos cartões para seus nomes, limites e datas de vencimento
  for (let i = 1; i < dadosContas.length; i++) { // Ignora o cabeçalho
    const row = dadosContas[i];
    const nomeConta = row[0]; // Coluna A: Nome da Conta
    const tipoConta = row[1]; // Coluna B: Tipo (Débito/Crédito)
    const idCartao = row[2]; // Coluna C: ID (para cartões)
    const limite = parseFloat(row[3]) || 0; // Coluna D: Limite
    const vencimento = row[4]; // Coluna E: Dia de Vencimento

    if (tipoConta === 'Cartão de Crédito' && idCartao) {
      limitesCartao[idCartao] = limite;
      gastosCartao[idCartao] = { nome: nomeConta, totalGasto: 0 };
      datasVencimento[idCartao] = vencimento;
    }
  }

  transacoesRaw.forEach(row => {
    const data = new Date(row[0]);
    const valor = parseFloat(row[2]);
    const tipo = row[3];
    const idCartaoTransacao = row[7]; // Coluna H: ID do Cartão

    if (data.getMonth() + 1 === mes && data.getFullYear() === ano && tipo === 'Despesa' && idCartaoTransacao) {
      if (gastosCartao[idCartaoTransacao]) {
        gastosCartao[idCartaoTransacao].totalGasto += valor;
      }
    }
  });

  const resultado = {};
  for (const id in gastosCartao) {
    resultado[gastosCartao[id].nome] = {
      totalGasto: gastosCartao[id].totalGasto,
      limite: limitesCartao[id],
      vencimento: datasVencimento[id]
    };
  }

  return resultado;
}

