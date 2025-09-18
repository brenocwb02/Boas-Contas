/**
 * @file FinancialLogic.gs
 * @description Este arquivo contém a lógica de negócio central do bot financeiro.
 * Inclui interpretação de mensagens, cálculos financeiros, categorização e atualização de saldos.
 * VERSÃO OTIMIZADA E CORRIGIDA.
 */

// As constantes de estado do tutorial (TUTORIAL_STATE_WAITING_DESPESA, etc.) foram movidas para Management.gs
// para evitar redeclaração e garantir um ponto único de verdade.

// Variáveis globais para os dados da planilha que são acessados frequentemente
// Serão populadas e armazenadas em cache.
let cachedPalavrasChave = null;
let cachedCategorias = null;
let cachedContas = null;
let cachedConfig = null;

/**
 * Obtém dados de uma aba da planilha e os armazena em cache.
 * @param {string} sheetName O nome da aba.
 * @param {string} cacheKey A chave para o cache.
 * @param {number} [expirationInSeconds=300] Tempo de expiração do cache em segundos.
 * @returns {Array<Array<any>>} Os dados da aba (incluindo cabeçalhos).
 */
function getSheetDataWithCache(sheetName, cacheKey, expirationInSeconds = 300) {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);

  if (cachedData) {
    // logToSheet(`Dados da aba '${sheetName}' recuperados do cache.`, "DEBUG");
    return JSON.parse(cachedData);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    logToSheet(`ERRO: Aba '${sheetName}' não encontrada.`, "ERROR");
    throw new Error(`Aba '${sheetName}' não encontrada.`);
  }

  const data = sheet.getDataRange().getValues();
  cache.put(cacheKey, JSON.stringify(data), expirationInSeconds);
  // logToSheet(`Dados da aba '${sheetName}' lidos da planilha e armazenados em cache.`, "DEBUG");
  return data;
}

/**
 * ATUALIZADO: Interpreta uma mensagem do Telegram para extrair informações de transação.
 * Agora salva a categoria original para acionar o aprendizado.
 * @param {string} mensagem O texto da mensagem recebida.
 * @param {string} usuario O nome do usuário que enviou a mensagem.
 * @param {string} chatId O ID do chat do Telegram.
 * @returns {Object} Um objeto contendo os detalhes da transação ou uma mensagem de erro/status.
 */
function interpretarMensagemTelegram(mensagem, usuario, chatId) {
  logToSheet(`Interpretando mensagem: "${mensagem}" para usuário: ${usuario}`, "INFO");

  const dadosPalavras = getSheetDataWithCache(SHEET_PALAVRAS_CHAVE, CACHE_KEY_PALAVRAS);
  const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);
  
  // ### INÍCIO DA CORREÇÃO DE LÓGICA ###
  // Texto para parsing de números, que mantém a pontuação (vírgulas, pontos) e remove acentos.
  const textoParaParseNumeros = mensagem.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  // Texto para busca de palavras-chave, que remove toda a pontuação.
  const textoNormalizado = normalizarTexto(mensagem);
  
  // ### INÍCIO DA NOVA LÓGICA DE DIVISÃO E EMPRÉSTIMO ###
  const dividirMatch = textoNormalizado.match(/(?:dividi|dividir)\s+([\d.,]+)\s+(?:de|do|da)?\s*(.*?)\s+com\s+(.+)/i);
  const empresteiMatch = textoNormalizado.match(/(?:emprestei|adiantei)\s+([\d.,]+)\s+(?:para)?\s*(.*?)(?=\s+com|\s+pelo|\s+pela|$)/i);
  
  if (dividirMatch) {
    const valor = parseBrazilianFloat(dividirMatch[1]);
    const descricao = dividirMatch[2];
    const pessoa = dividirMatch[3].split(/\s+(?:com|pelo|pela)\s+/)[0]; // Pega o nome da pessoa
    const restoDaFrase = dividirMatch[3]; // O resto da frase pode conter a conta
    handleDividirDespesa(chatId, usuario, valor, descricao, pessoa, restoDaFrase);
    return { handled: true };
  }
  
  if (empresteiMatch) {
    const valor = parseBrazilianFloat(empresteiMatch[1]);
    const pessoaEDesc = empresteiMatch[2]; // Pode conter a pessoa e a descrição
    const { conta } = extrairContaMetodoPagamento(textoNormalizado, dadosContas, dadosPalavras);
    handleEmprestarValor(chatId, usuario, valor, pessoaEDesc, conta);
    return { handled: true };
  }
  // ### FIM DA NOVA LÓGICA ###


  // Verifica PRIMEIRO se a mensagem é sobre compra ou venda de ativos, usando o texto com pontuação.
  const investmentMatch = textoParaParseNumeros.match(/(comprei|vendi)\s+(\d+[\d.,]*)\s+(?:acoes de|de|do|da)?\s*([a-zA-Z0-9]+)\s*(?:a|por)\s+([\d.,]+(?:[\s]reais)?(?:[\s]e[\s][\d]+)?(?:[\s]centavos)?)\s*(?:pela|pelo|na|no|da|do)?\s*(.+)/i);

  if (investmentMatch) {
    const acao = investmentMatch[1].toLowerCase();
    const quantidade = parseInt(investmentMatch[2].replace(/\./g, ''));
    const ticker = investmentMatch[3].toUpperCase();
    const precoTexto = investmentMatch[4];
    const nomeCorretora = investmentMatch[5].trim();
    const preco = parseBrazilianFloat(precoTexto); // CORREÇÃO: Usa a função parseBrazilianFloat diretamente
    
    logToSheet(`[Investimento Detectado] Ação: ${acao}, Qtd: ${quantidade}, Ticker: ${ticker}, Preço: ${preco}, Corretora: ${nomeCorretora}`, "INFO");

    if (acao === 'comprei') {
      handleComprarAtivo(chatId, ticker, quantidade, preco, nomeCorretora, usuario);
    } else if (acao === 'vendi') {
      handleVenderAtivo(chatId, ticker, quantidade, preco, nomeCorretora, usuario);
    }
    return { handled: true }; // Indica que a mensagem foi tratada e interrompe o processamento
  }
  
  // O resto da função continua usando a variável apropriada para cada tarefa.
  const tipoInfo = detectarTipoTransacao(textoNormalizado, dadosPalavras);
  if (!tipoInfo) {
    return { errorMessage: "Não consegui identificar se é uma despesa, receita ou transferência. Tente ser mais claro." };
  }
  const tipoTransacao = tipoInfo.tipo;
  const keywordTipo = tipoInfo.keyword;
  logToSheet(`Tipo de transação detectado: ${tipoTransacao} (keyword: ${keywordTipo})`, "DEBUG");

  // Usa o texto com pontuação para extrair o valor corretamente.
  const valor = extrairValor(textoParaParseNumeros);
  const transactionId = Utilities.getUuid().substring(0, 8);
  // ### FIM DA CORREÇÃO DE LÓGICA ###

  if (tipoTransacao === "Transferência") {
      if (isNaN(valor) || valor <= 0) {
        return { errorMessage: "Não consegui identificar o valor da transferência." };
      }
      const { contaOrigem, contaDestino } = extrairContasTransferencia(textoNormalizado, dadosContas, dadosPalavras);
      
      const transacaoParcialTransfer = { // Renomeada para evitar conflito de escopo
        id: transactionId,
        tipo: "Transferência",
        valor: valor,
        contaOrigem: contaOrigem,
        contaDestino: contaDestino,
        usuario: usuario
      };

      if (contaOrigem === "Não Identificada") {
        return solicitarInformacaoFaltante("conta_origem", transacaoParcialTransfer, chatId);
      }
      if (contaDestino === "Não Identificada") {
        return solicitarInformacaoFaltante("conta_destino", transacaoParcialTransfer, chatId);
      }
      
      return prepararConfirmacaoTransferencia(transacaoParcialTransfer, chatId);
  }

  const { conta, infoConta, metodoPagamento } = extrairContaMetodoPagamento(textoNormalizado, dadosContas, dadosPalavras);
  const { categoria, subcategoria } = extrairCategoriaSubcategoria(textoNormalizado, tipoTransacao, dadosPalavras);
  const parcelasTotais = extrairParcelas(textoNormalizado);
  const descricao = extrairDescricao(textoNormalizado, String(valor), [keywordTipo, conta, metodoPagamento]);

  const transacaoParcial = {
    id: transactionId,
    data: new Date(),
    descricao: descricao,
    categoria: categoria,
    subcategoria: subcategoria,
    originalCategory: categoria, // <-- ADICIONADO: Salva a categoria original detectada
    tipo: tipoTransacao,
    valor: valor,
    metodoPagamento: metodoPagamento,
    conta: conta,
    infoConta: infoConta,
    parcelasTotais: parcelasTotais,
    parcelaAtual: 1,
    dataVencimento: new Date(),
    usuario: usuario,
    status: "Pendente",
    dataRegistro: new Date()
  };

  if (isNaN(valor) || valor <= 0) {
    return solicitarInformacaoFaltante("valor", transacaoParcial, chatId);
  }
  if (conta === "Não Identificada") {
    return solicitarInformacaoFaltante("conta", transacaoParcial, chatId);
  }
  if (categoria === "Não Identificada") {
    return solicitarInformacaoFaltante("categoria", transacaoParcial, chatId);
  }
  if (metodoPagamento === "Não Identificado") {
    return solicitarInformacaoFaltante("metodo", transacaoParcial, chatId);
  }

  let dataVencimentoFinal = new Date();
  let isCreditCardTransaction = false;
  if (infoConta && normalizarTexto(infoConta.tipo) === "cartao de credito") {
    isCreditCardTransaction = true;
    dataVencimentoFinal = calcularVencimentoCartao(infoConta, new Date(), dadosContas);
  }
  transacaoParcial.dataVencimento = dataVencimentoFinal;
  transacaoParcial.isCreditCardTransaction = isCreditCardTransaction;
  transacaoParcial.finalId = Utilities.getUuid();
  
  // Lógica de Nudge movida para antes da confirmação
  const nudgeMessage = getNudgeMessage(chatId, transacaoParcial);
  if (nudgeMessage) {
      transacaoParcial.nudge = nudgeMessage;
  }

  if (parcelasTotais > 1) {
    return prepararConfirmacaoParcelada(transacaoParcial, chatId);
  } else {
    return prepararConfirmacaoSimples(transacaoParcial, chatId);
  }
}


/**
 * NOVO: Centraliza a lógica para solicitar informações faltantes ao usuário.
 * @param {string} campoFaltante O nome do campo que precisa ser preenchido ('valor', 'conta', 'categoria', 'conta_origem', 'conta_destino').
 * @param {Object} transacaoParcial O objeto de transação com os dados já coletados.
 * @param {string} chatId O ID do chat do Telegram.
 * @returns {Object} Um objeto indicando que uma ação do assistente está pendente.
 */
function solicitarInformacaoFaltante(campoFaltante, transacaoParcial, chatId) {
  let mensagem = "";
  let teclado = { inline_keyboard: [] };
  let optionsList = [];
  const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);

  // Adiciona a informação que estamos aguardando ao estado
  transacaoParcial.waitingFor = campoFaltante;

  switch (campoFaltante) {
    case "valor":
      mensagem = `Sou eu, o Zaq! Entendi o lançamento, mas não encontrei o valor. Pode dizer-me qual foi, por favor?`;
      transacaoParcial.waitingFor = 'valor';
      setAssistantState(chatId, transacaoParcial);
      enviarMensagemTelegram(chatId, mensagem);
      break;

    case "conta":
      mensagem = `Entendido! Para que este(a) *${escapeMarkdown(transacaoParcial.tipo)}* fique registado corretamente, de qual conta ou cartão devo usar?`;
      optionsList = dadosContas.slice(1).map(row => row[0]).filter(Boolean);
      optionsList.forEach((option, index) => {
        const button = { text: option, callback_data: `complete_conta_${transacaoParcial.id}_${index}` };
        if (index % 2 === 0) teclado.inline_keyboard.push([button]);
        else teclado.inline_keyboard[teclado.inline_keyboard.length - 1].push(button);
      });
      transacaoParcial.assistantOptions = optionsList;
      setAssistantState(chatId, transacaoParcial);
      enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
      break;

    case "conta_origem":
      mensagem = `Ok, entendi uma transferência de *${formatCurrency(transacaoParcial.valor)}* para *${escapeMarkdown(transacaoParcial.contaDestino)}*. De qual conta o dinheiro saiu?`;
      optionsList = dadosContas.slice(1).map(row => row[0]).filter(Boolean);
      optionsList.forEach((option, index) => {
        const button = { text: option, callback_data: `complete_conta_origem_${transacaoParcial.id}_${index}` };
        if (index % 2 === 0) teclado.inline_keyboard.push([button]);
        else teclado.inline_keyboard[teclado.inline_keyboard.length - 1].push(button);
      });
      transacaoParcial.assistantOptions = optionsList;
      setAssistantState(chatId, transacaoParcial);
      enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
      break;
    
    case "conta_destino":
       mensagem = `Ok, entendi uma transferência de *${formatCurrency(transacaoParcial.valor)}* de *${escapeMarkdown(transacaoParcial.contaOrigem)}*. Para qual conta o dinheiro foi?`;
      optionsList = dadosContas.slice(1).map(row => row[0]).filter(Boolean);
      optionsList.forEach((option, index) => {
        const button = { text: option, callback_data: `complete_conta_destino_${transacaoParcial.id}_${index}` };
        if (index % 2 === 0) teclado.inline_keyboard.push([button]);
        else teclado.inline_keyboard[teclado.inline_keyboard.length - 1].push(button);
      });
      transacaoParcial.assistantOptions = optionsList;
      setAssistantState(chatId, transacaoParcial);
      enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
      break;

    case "categoria":
      mensagem = `Em qual categoria este lançamento se encaixa?`;
      const dadosCategorias = getSheetDataWithCache(SHEET_CATEGORIAS, 'categorias_cache');
      optionsList = [...new Set(dadosCategorias.slice(1).map(row => row[0]))].filter(Boolean);
      optionsList.forEach((option, index) => {
        const button = { text: option, callback_data: `complete_categoria_${transacaoParcial.id}_${index}` };
        if (index % 2 === 0) teclado.inline_keyboard.push([button]);
        else teclado.inline_keyboard[teclado.inline_keyboard.length - 1].push(button);
      });
      transacaoParcial.assistantOptions = optionsList;
      setAssistantState(chatId, transacaoParcial);
      enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
      break;
    
    case "metodo":
      mensagem = `Qual foi o método de pagamento?`;
      const dadosPalavras = getSheetDataWithCache(SHEET_PALAVRAS_CHAVE, CACHE_KEY_PALAVRAS);
      optionsList = dadosPalavras.slice(1).filter(row => row[0].toLowerCase() === 'meio_pagamento').map(row => row[2]).filter(Boolean);
      optionsList.forEach((option, index) => {
        const button = { text: option, callback_data: `complete_metodo_${transacaoParcial.id}_${index}` };
        if (index % 2 === 0) teclado.inline_keyboard.push([button]);
        else teclado.inline_keyboard[teclado.inline_keyboard.length - 1].push(button);
      });
      transacaoParcial.assistantOptions = optionsList;
      setAssistantState(chatId, transacaoParcial);
      enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
      break;
  }

  logToSheet(`Assistente solicitando '${campoFaltante}' para transação ID ${transacaoParcial.id}`, "INFO");
  return { status: "PENDING_ASSISTANT_ACTION", transactionId: transacaoParcial.id };
}

/**
 * NOVO: Continua o fluxo do assistente após o usuário fornecer uma informação.
 * @param {Object} transacaoParcial O objeto de transação atualizado.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 */
function processAssistantCompletion(transacaoParcial, chatId, usuario) {
  logToSheet(`Continuando fluxo do assistente para transação ID ${transacaoParcial.id}`, "INFO");

  // Se for uma transferência, verifica se ambas as contas estão preenchidas
  if (transacaoParcial.tipo === "Transferência") {
    if (transacaoParcial.contaOrigem === "Não Identificada") {
      return solicitarInformacaoFaltante("conta_origem", transacaoParcial, chatId);
    }
    if (transacaoParcial.contaDestino === "Não Identificada") {
      return solicitarInformacaoFaltante("conta_destino", transacaoParcial, chatId);
    }
    // Se ambas estiverem ok, vai para a confirmação
    return prepararConfirmacaoTransferencia(transacaoParcial, chatId);
  }

  // Fluxo para Despesa e Receita
  if (transacaoParcial.conta === "Não Identificada") {
    return solicitarInformacaoFaltante("conta", transacaoParcial, chatId);
  }
  if (transacaoParcial.categoria === "Não Identificada") {
    return solicitarInformacaoFaltante("categoria", transacaoParcial, chatId);
  }
  if (transacaoParcial.subcategoria === "Não Identificada") {
      const dadosCategorias = getSheetDataWithCache(SHEET_CATEGORIAS, 'categorias_cache');
      const subcategoriasParaCategoria = dadosCategorias.slice(1).filter(row => normalizarTexto(row[0]) === normalizarTexto(transacaoParcial.categoria)).map(row => row[1]);
      if (subcategoriasParaCategoria.length > 1) {
          return solicitarSubcategoria(transacaoParcial, subcategoriasParaCategoria, chatId);
      } else if (subcategoriasParaCategoria.length === 1) {
          transacaoParcial.subcategoria = subcategoriasParaCategoria[0];
      } else {
          transacaoParcial.subcategoria = transacaoParcial.categoria;
      }
  }
  if (transacaoParcial.metodoPagamento === "Não Identificado") {
    return solicitarInformacaoFaltante("metodo", transacaoParcial, chatId);
  }
  
  // Se tudo estiver completo, prossegue para a confirmação
  const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);
  let dataVencimentoFinal = new Date();
  let isCreditCardTransaction = false;
  if (transacaoParcial.infoConta && normalizarTexto(transacaoParcial.infoConta.tipo) === "cartao de credito") {
    isCreditCardTransaction = true;
    dataVencimentoFinal = calcularVencimentoCartao(transacaoParcial.infoConta, new Date(transacaoParcial.data), dadosContas);
  }
  transacaoParcial.dataVencimento = dataVencimentoFinal;
  transacaoParcial.isCreditCardTransaction = isCreditCardTransaction;
  transacaoParcial.finalId = Utilities.getUuid();

  if (transacaoParcial.parcelasTotais > 1) {
    return prepararConfirmacaoParcelada(transacaoParcial, chatId);
  } else {
    return prepararConfirmacaoSimples(transacaoParcial, chatId);
  }
}

/**
 * NOVO: Solicita a subcategoria ao usuário quando há múltiplas opções.
 * @param {Object} transacaoParcial O objeto de transação com os dados já coletados.
 * @param {Array<string>} subcategorias A lista de subcategorias disponíveis.
 * @param {string} chatId O ID do chat do Telegram.
 */
function solicitarSubcategoria(transacaoParcial, subcategorias, chatId) {
  let mensagem = `Para a categoria *${escapeMarkdown(transacaoParcial.categoria)}*, qual subcategoria você gostaria de usar?`;
  let teclado = { inline_keyboard: [] };
  
  subcategorias.forEach((sub, index) => {
    const button = { text: sub, callback_data: `complete_subcategoria_${transacaoParcial.id}_${index}` };
    if (index % 2 === 0) {
      teclado.inline_keyboard.push([button]);
    } else {
      teclado.inline_keyboard[teclado.inline_keyboard.length - 1].push(button);
    }
  });

  transacaoParcial.assistantOptions = subcategorias;
  setAssistantState(chatId, transacaoParcial);

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
  logToSheet(`Assistente solicitando 'subcategoria' para transação ID ${transacaoParcial.id}`, "INFO");
  return { status: "PENDING_ASSISTANT_ACTION", transactionId: transacaoParcial.id };
}


/**
 * CORRIGIDO: Detecta o tipo de transação e a palavra-chave que o acionou.
 * @param {string} mensagemCompleta O texto da mensagem normalizada.
 * @param {Array<Array<any>>} dadosPalavras Os dados da aba "PalavrasChave".
 * @returns {Object|null} Um objeto {tipo, keyword} ou null se não for detectado.
 */
function detectarTipoTransacao(mensagemCompleta, dadosPalavras) {
  logToSheet(`[detectarTipoTransacao] Mensagem Completa: "${mensagemCompleta}"`, "DEBUG");

  const palavrasReceitaFixas = ['recebi', 'salario', 'rendeu', 'pix recebido', 'transferencia recebida', 'deposito', 'entrada', 'renda', 'pagamento recebido', 'reembolso', 'cashback'];
  const palavrasDespesaFixas = ['gastei', 'paguei', 'comprei', 'saida', 'débito', 'debito'];
  const palavrasTransferenciaFixas = ['transferi', 'transferir']; // CORREÇÃO: Adicionado "transferir"

  for (let palavra of palavrasTransferenciaFixas) {
    if (mensagemCompleta.includes(palavra)) {
      logToSheet(`[detectarTipoTransacao] Transferência detectada pela palavra fixa: "${palavra}"`, "DEBUG");
      return { tipo: "Transferência", keyword: palavra };
    }
  }

  for (let palavraRec of palavrasReceitaFixas) {
    if (mensagemCompleta.includes(palavraRec)) {
      logToSheet(`[detectarTipoTransacao] Receita detectada pela palavra fixa: "${palavraRec}"`, "DEBUG");
      return { tipo: "Receita", keyword: palavraRec };
    }
  }

  for (let palavraDes of palavrasDespesaFixas) {
    if (mensagemCompleta.includes(palavraDes)) {
      logToSheet(`[detectarTipoTransacao] Despesa detectada pela palavra fixa: "${palavraDes}"`, "DEBUG");
      return { tipo: "Despesa", keyword: palavraDes };
    }
  }

  for (let i = 1; i < dadosPalavras.length; i++) {
    const tipoPalavra = (dadosPalavras[i][0] || "").toString().trim().toLowerCase();
    const chave = normalizarTexto(dadosPalavras[i][1] || "");
    const valorInterpretado = (dadosPalavras[i][2] || "").toString().trim();

    if (tipoPalavra === "tipo_transacao" && chave) {
      const regex = new RegExp(`\\b${chave}\\b`);
      if (regex.test(mensagemCompleta)) {
        logToSheet(`[detectarTipoTransacao] Tipo detectado da planilha: "${valorInterpretado}" pela palavra: "${chave}"`, "DEBUG");
        return { tipo: valorInterpretado, keyword: chave };
      }
    }
  }

  logToSheet("[detectarTipoTransacao] Nenhum tipo especifico detectado. Retornando null.", "WARN");
  return null;
}

/**
 * ATUALIZADO: Extrai o valor numérico da mensagem.
 * Agora, a função `parseBrazilianFloat` lida corretamente com as diversas formatações.
 * @param {string} textoNormalizado O texto da mensagem normalizado.
 * @returns {number} O valor numérico extraído, ou NaN.
 */
function extrairValor(textoNormalizado) {
  // A regex agora é mais simples, apenas para encontrar um bloco que parece um número.
  // A complexidade da formatação é delegada para a função parseBrazilianFloat.
  const regex = /(\d[\d\.,]*)/; 
  const match = textoNormalizado.match(regex);
  if (match && match[1]) {
    return parseBrazilianFloat(match[1]); 
  }
  return NaN;
}

/**
 * ATUALIZADO: Extrai a conta, método de pagamento e as palavras-chave correspondentes.
 * @param {string} textoNormalizado O texto da mensagem normalizado.
 * @param {Array<Array<any>>} dadosContas Os dados da aba 'Contas'.
 * @param {Array<Array<any>>} dadosPalavras Os dados da aba 'PalavrasChave'.
 * @returns {Object} Objeto com conta, infoConta, metodoPagamento, keywordConta e keywordMetodo.
 */
function extrairContaMetodoPagamento(textoNormalizado, dadosContas, dadosPalavras) {
  let contaEncontrada = "Não Identificada";
  let metodoPagamentoEncontrado = "Não Identificado";
  let melhorInfoConta = null;
  let maiorSimilaridadeConta = 0;
  let melhorPalavraChaveConta = "";
  let melhorPalavraChaveMetodo = "";

  // 1. Encontrar a melhor conta/cartão
  for (let i = 1; i < dadosContas.length; i++) {
    const nomeContaPlanilha = (dadosContas[i][0] || "").toString().trim();
    const nomeContaNormalizado = normalizarTexto(nomeContaPlanilha);
    const palavrasChaveConta = (dadosContas[i][3] || "").toString().trim().split(',').map(s => normalizarTexto(s.trim()));
    palavrasChaveConta.push(nomeContaNormalizado);

    for (const palavraChave of palavrasChaveConta) {
        if (!palavraChave) continue;
        if (textoNormalizado.includes(palavraChave)) {
            const similarity = calculateSimilarity(textoNormalizado, palavraChave);
            const currentSimilarity = (palavraChave === nomeContaNormalizado) ? similarity * 1.5 : similarity; 
            if (currentSimilarity > maiorSimilaridadeConta) {
                maiorSimilaridadeConta = currentSimilarity;
                contaEncontrada = nomeContaPlanilha;
                melhorInfoConta = obterInformacoesDaConta(nomeContaPlanilha, dadosContas);
                melhorPalavraChaveConta = palavraChave;
            }
        }
    }
  }

  // 2. Extrair Método de Pagamento
  let maiorSimilaridadeMetodo = 0;
  for (let i = 1; i < dadosPalavras.length; i++) {
    const tipo = (dadosPalavras[i][0] || "").toString().trim().toLowerCase();
    const palavraChave = (dadosPalavras[i][1] || "").toString().trim().toLowerCase();
    const valorInterpretado = (dadosPalavras[i][2] || "").toString().trim();

    if (tipo === "meio_pagamento" && palavraChave && textoNormalizado.includes(palavraChave)) {
        const similarity = calculateSimilarity(textoNormalizado, palavraChave);
        if (similarity > maiorSimilaridadeMetodo) {
          maiorSimilaridadeMetodo = similarity;
          metodoPagamentoEncontrado = valorInterpretado;
          melhorPalavraChaveMetodo = palavraChave;
        }
    }
  }

  // 3. Lógica de fallback para método de pagamento
  if (melhorInfoConta && normalizarTexto(melhorInfoConta.tipo) === "cartao de credito") {
    if (normalizarTexto(metodoPagamentoEncontrado) === "nao identificado" || normalizarTexto(metodoPagamentoEncontrado) === "debito") {
      metodoPagamentoEncontrado = "Crédito";
      logToSheet(`[ExtrairContaMetodo] Conta e cartao de credito, metodo de pagamento ajustado para "Credito".`, "DEBUG");
    }
  }
  
  return { 
      conta: contaEncontrada, 
      infoConta: melhorInfoConta, 
      metodoPagamento: metodoPagamentoEncontrado,
      keywordConta: melhorPalavraChaveConta,
      keywordMetodo: melhorPalavraChaveMetodo
  };
}


/**
 * ATUALIZADO: Extrai categoria e subcategoria, priorizando o conhecimento aprendido.
 * @param {string} textoNormalizado O texto da mensagem normalizado.
 * @param {string} tipoTransacao O tipo de transação (Despesa, Receita).
 * @param {Array<Array<any>>} dadosPalavras Os dados da aba 'PalavrasChave'.
 * @returns {Object} Objeto com categoria, subcategoria e keywordCategoria.
 */
function extrairCategoriaSubcategoria(textoNormalizado, tipoTransacao, dadosPalavras) {
  // --- PASSO 1: Tenta encontrar uma categoria aprendida de alta confiança ---
  const learned = findLearnedCategory(textoNormalizado);
  if (learned) {
      return {
          categoria: learned.categoria,
          subcategoria: learned.subcategoria,
          keywordCategoria: learned.keyword // Usa a keyword que ativou a regra
      };
  }

  // --- PASSO 2: Se não encontrou, usa a lógica original com PalavrasChave ---
  let categoriaEncontrada = "Não Identificada";
  let subcategoriaEncontrada = "Não Identificada";
  let melhorScoreSubcategoria = -1;
  let melhorPalavraChaveCategoria = "";

  for (let i = 1; i < dadosPalavras.length; i++) {
    const tipoPalavraChave = (dadosPalavras[i][0] || "").toString().trim().toLowerCase();
    const palavraChave = (dadosPalavras[i][1] || "").toString().trim().toLowerCase();
    const valorInterpretado = (dadosPalavras[i][2] || "").toString().trim();

    if (tipoPalavraChave === "subcategoria" && palavraChave) {
        const regex = new RegExp(`\\b${palavraChave}\\b`, 'i');
        if (regex.test(textoNormalizado)) {
            const similarity = calculateSimilarity(textoNormalizado, palavraChave); 
            if (similarity > melhorScoreSubcategoria) { 
              if (valorInterpretado.includes(">")) {
                const partes = valorInterpretado.split(">");
                const categoria = partes[0].trim();
                const subcategoria = partes[1].trim();
                const tipoCategoria = (dadosPalavras[i][3] || "").toString().trim().toLowerCase();
                
                if (!tipoCategoria || normalizarTexto(tipoCategoria) === normalizarTexto(tipoTransacao)) {
                  categoriaEncontrada = categoria;
                  subcategoriaEncontrada = subcategoria;
                  melhorScoreSubcategoria = similarity;
                  melhorPalavraChaveCategoria = palavraChave;
                }
              }
            }
        }
    }
  }
  return { 
      categoria: categoriaEncontrada, 
      subcategoria: subcategoriaEncontrada,
      keywordCategoria: melhorPalavraChaveCategoria
  };
}



/**
 * **CORRIGIDO:** Extrai a descrição final da transação de forma mais robusta.
 * Remove proativamente palavras-chave, valor e frases de parcelamento para isolar a descrição.
 * @param {string} textoNormalizado O texto normalizado da mensagem do usuário.
 * @param {string} valor O valor extraído (como string).
 * @param {Array<string>} keywordsToRemove As palavras-chave a serem removidas.
 * @returns {string} A descrição limpa.
 */
function extrairDescricao(textoNormalizado, valor, keywordsToRemove) {
  let descricao = ` ${textoNormalizado} `; // Adiciona espaços para facilitar a substituição de palavras inteiras

  // 1. Remove o valor
  descricao = descricao.replace(` ${valor.replace('.', ',')} `, ' ');
  descricao = descricao.replace(` ${valor.replace(',', '.')} `, ' ');

  // 2. Remove frases de parcelamento de forma mais segura
  descricao = descricao.replace(/\s+em\s+\d+\s*x\s+/gi, " ");
  descricao = descricao.replace(/\s+\d+\s*x\s+/gi, " ");
  descricao = descricao.replace(/\s+\d+\s*vezes\s+/gi, " ");

  // 3. Remove outras palavras-chave (tipo, conta, método de pagamento)
  keywordsToRemove.forEach(keyword => {
    if (keyword) {
      const keywordNorm = normalizarTexto(keyword);
      // Usa regex com \b para garantir que está removendo a palavra inteira
      descricao = descricao.replace(new RegExp(`\\b${keywordNorm}\\b`, "gi"), '');
    }
  });
  

  // 4. NOVO: Remove variações de moeda
  const currencyWords = ['reais', 'real', 'r'];
  currencyWords.forEach(word => {
    descricao = descricao.replace(new RegExp(`\\s+${word}\\s+`, 'gi'), " ");
  });

  // 5. Limpa preposições comuns que podem sobrar
  const preposicoes = ['de', 'da', 'do', 'dos', 'das', 'e', 'ou', 'a', 'o', 'no', 'na', 'nos', 'nas', 'com', 'em', 'para', 'por', 'pelo', 'pela', 'via'];
  preposicoes.forEach(prep => {
    descricao = descricao.replace(new RegExp(`\\s+${prep}\\s+`, 'gi'), " ");
  });

  // 6. Limpa espaços extras e retorna
  descricao = descricao.replace(/\s+/g, " ").trim();
  
  if (descricao.length < 2) {
    return "Lançamento Geral";
  }

  return capitalize(descricao);
}

/**
 * Extrai o número total de parcelas da mensagem.
 * @param {string} textoNormalizado O texto da mensagem normalizado.
 * @returns {number} O número de parcelas (padrão 1 se não for encontrado).
 */
function extrairParcelas(textoNormalizado) {
  const regex = /(\d+)\s*(?:x|vezes)/;
  const match = textoNormalizado.match(regex);
  return match ? parseInt(match[1], 10) : 1;
}

/**
 * Prepara e envia uma mensagem de confirmação para transações simples (não parceladas).
 * Armazena os dados da transação em cache.
 * **AGORA INCLUI A LÓGICA DE NUDGE COMPORTAMENTAL.**
 * @param {Object} transacaoData Os dados da transação.
 * @param {string} chatId O ID do chat do Telegram.
 * @returns {Object} Status de confirmação pendente.
 */
function prepararConfirmacaoSimples(transacaoData, chatId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `${CACHE_KEY_PENDING_TRANSACTIONS}_${chatId}_${transacaoData.finalId}`;
  cache.put(cacheKey, JSON.stringify(transacaoData), CACHE_EXPIRATION_PENDING_TRANSACTION_SECONDS);

  let mensagem = `✅ Entendido! Registado. Por favor, confirme se está tudo certo:\n\n`;
  mensagem += `*Tipo:* ${escapeMarkdown(transacaoData.tipo)}\n`;
  mensagem += `*Descricao:* ${escapeMarkdown(transacaoData.descricao)}\n`;
  mensagem += `*Valor:* ${formatCurrency(transacaoData.valor)}\n`;
  mensagem += `*Conta:* ${escapeMarkdown(transacaoData.conta)}\n`;
  mensagem += `*Metodo:* ${escapeMarkdown(transacaoData.metodoPagamento)}\n`;
  mensagem += `*Categoria:* ${escapeMarkdown(transacaoData.categoria)}\n`;
  mensagem += `*Subcategoria:* ${escapeMarkdown(transacaoData.subcategoria)}\n`;
  
  // Adiciona a mensagem de nudge no topo, se ela existir
  if (transacaoData.nudge) {
    mensagem = `*${transacaoData.nudge}*\n\n` + mensagem;
  }

  const teclado = {
    inline_keyboard: [
      [{ text: "✅ Confirmar", callback_data: `confirm_${transacaoData.finalId}` }],
      [{ text: "❌ Cancelar", callback_data: `cancel_${transacaoData.finalId}` }]
    ]
  };

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
  return { status: "PENDING_CONFIRMATION", transactionId: transacaoData.finalId };
}

/**
 * **FUNÇÃO ATUALIZADA**
 * Analisa uma transação pendente e o perfil do usuário para gerar uma mensagem
 * de "nudge" (empurrãozinho) comportamental, se aplicável.
 * @param {string} chatId O ID do chat do usuário.
 * @param {Object} transacao Os dados da transação pendente.
 * @returns {string|null} A mensagem de nudge ou null se nenhuma for aplicável.
 */
function getNudgeMessage(chatId, transacao) {
  // ATUALIZADO: Agora chama a função correta
  const perfilUsuario = getFinancialProfile(chatId); 
  
  if (!perfilUsuario || perfilUsuario.inProgress || transacao.tipo !== 'Despesa') {
    return null; // Só aplica nudges para despesas e se o perfil for conhecido e o quiz não estiver em andamento
  }

  const perfil = perfilUsuario.perfil;
  let nudge = null;
  let nudgeReason = '';
  // Lógica de gatilho para o perfil "Despreocupado"
  if (perfil === 'Despreocupado') {
    // Gatilho 1: Gasto alto em categorias de "Desejos"
    const categoriasDesejo = ['lazer e entretenimento', 'despesas pessoais'];
    if (categoriasDesejo.includes(normalizarTexto(transacao.categoria)) && transacao.valor > 150) {
      nudge = "⚠️ Atenção, Despreocupado! Este é um gasto por impulso ou está planeado? Lembre-se do seu desafio de criar o hábito de poupar.";
    }
  }

  // Lógica de gatilho para o perfil "Sonhador"
  if (perfil === 'Sonhador') {
    const totalMetas = getTotalGoalsSaved(); // Função de apoio do Quiz.gs
    // Gatilho 1: Gasto não essencial quando as metas estão com poucos aportes
    if (totalMetas < 500 && transacao.valor > 100 && normalizarTexto(transacao.categoria) !== 'moradia') {
       nudge = `🤔 Olá, Sonhador! Este gasto de ${formatCurrency(transacao.valor)} te aproxima ou te afasta dos seus grandes sonhos?`;
    }
  }
  
  // Lógica de gatilho para o perfil "Construtor"
  if (perfil === 'Construtor') {
      // Gatilho 1: Muitos gastos pequenos não categorizados podem indicar falta de atenção ao "micro"
      if (normalizarTexto(transacao.categoria) === 'não identificada' && transacao.valor < 50) {
          nudge = "🔎 Olá, Construtor! Notamos um gasto não categorizado. Lembre-se que cuidar dos pequenos detalhes também otimiza o seu património.";
      }
  }

  if (nudge) {
    logToSheet(`[Nudge Gerado] Perfil: ${perfil}, Categoria: ${transacao.categoria}, Valor: ${transacao.valor}. Mensagem: ${nudge}`, "INFO");
  }
  if (nudgeReason) {
    const usuario = getUsuarioPorChatId(chatId, getSheetDataWithCache(SHEET_CONFIGURACOES, CACHE_KEY_CONFIG));
    const nomeCurto = usuario.split(' ')[0];
    nudge = `${nomeCurto}, sou eu, o Zac. Reparei neste gasto. Como seu agente, queria apenas perguntar: ${nudgeReason} Ele está alinhado com a transformação que estamos a construir juntos?`;
    logToSheet(`[Nudge Gerado] Perfil: ${perfil}. Mensagem: ${nudge}`, "INFO");
  }
  
  return nudge;
}




/**
 * Prepara e envia uma mensagem de confirmação para transações parceladas.
 * Armazena os dados da transação em cache.
 * @param {Object} transacaoData Os dados da transação.
 * @param {string} chatId O ID do chat do Telegram.
 * @returns {Object} Status de confirmação pendente.
 */
function prepararConfirmacaoParcelada(transacaoData, chatId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `${CACHE_KEY_PENDING_TRANSACTIONS}_${chatId}_${transacaoData.finalId}`;
  cache.put(cacheKey, JSON.stringify(transacaoData), CACHE_EXPIRATION_PENDING_TRANSACTION_SECONDS);

  let mensagem = `✅ Confirme seu Lançamento Parcelado:\n\n`;
  mensagem += `*Tipo:* ${escapeMarkdown(transacaoData.tipo)}\n`;
  mensagem += `*Descricao:* ${escapeMarkdown(transacaoData.descricao)}\n`;
  mensagem += `*Valor Total:* ${formatCurrency(transacaoData.valor)}\n`;
  mensagem += `*Parcelas:* ${transacaoData.parcelasTotais}x de ${formatCurrency(transacaoData.valor / transacaoData.parcelasTotais)}\n`;
  mensagem += `*Conta:* ${escapeMarkdown(transacaoData.conta)}\n`;
  mensagem += `*Metodo:* ${escapeMarkdown(transacaoData.metodoPagamento)}\n`;
  mensagem += `*Categoria:* ${escapeMarkdown(transacaoData.categoria)}\n`;
  mensagem += `*Subcategoria:* ${escapeMarkdown(transacaoData.subcategoria)}\n`;
  mensagem += `*Primeiro Vencimento:* ${Utilities.formatDate(transacaoData.dataVencimento, Session.getScriptTimeZone(), "dd/MM/yyyy")}\n`;


  const teclado = {
    inline_keyboard: [
      [{ text: "✅ Confirmar Parcelamento", callback_data: `confirm_${transacaoData.finalId}` }],
      [{ text: "❌ Cancelar", callback_data: `cancel_${transacaoData.finalId}` }]
    ]
  };

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
  return { status: "PENDING_CONFIRMATION", transactionId: transacaoData.finalId };
}

/**
 * OTIMIZADO: Registra a transação confirmada na planilha e ajusta os saldos.
 * AGORA TAMBÉM ACIONA O MECANISMO DE APRENDIZADO.
 * @param {Object} transacaoData Os dados da transação.
 * @param {string} usuario O nome do usuário que confirmou.
 * @param {string} chatId O ID do chat do Telegram.
 */
function registrarTransacaoConfirmada(transacaoData, usuario, chatId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
    const contasSheet = ss.getSheetByName(SHEET_CONTAS);

    if (!transacoesSheet || !contasSheet) {
      enviarMensagemTelegram(chatId, "❌ Erro: Aba 'Transacoes' ou 'Contas' não encontrada para registrar.");
      return;
    }
    
    const rowsToAdd = [];
    const timezone = ss.getSpreadsheetTimeZone();

    if (transacaoData.tipo === "Transferência") {
        const dataFormatada = `'${Utilities.formatDate(new Date(transacaoData.data), timezone, "dd/MM/yyyy")}`;
        const dataRegistroFormatada = `'${Utilities.formatDate(new Date(), timezone, "dd/MM/yyyy HH:mm:ss")}`;

        // Transação 1: Saída da conta de origem
        rowsToAdd.push([
            dataFormatada,
            `Transferência para ${transacaoData.contaDestino}`,
            "🔄 Transferências",
            "Entre Contas",
            "Despesa",
            transacaoData.valor,
            "Transferência",
            transacaoData.contaOrigem,
            1, 1, dataFormatada,
            usuario,
            "Ativo", `${transacaoData.finalId}-1`, dataRegistroFormatada
        ]);

        // Transação 2: Entrada na conta de destino
        rowsToAdd.push([
            dataFormatada,
            `Transferência de ${transacaoData.contaOrigem}`,
            "🔄 Transferências",
            "Entre Contas",
            "Receita",
            transacaoData.valor,
            "Transferência",
            transacaoData.contaDestino,
            1, 1, dataFormatada,
            usuario,
            "Ativo", `${transacaoData.finalId}-2`, dataRegistroFormatada
        ]);
        
        if (rowsToAdd.length > 0) {
            transacoesSheet.getRange(transacoesSheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
        }
        
        // --- INÍCIO DA MELHORIA: Lógica de Transferência para Cartão ---
        const contasSheetData = contasSheet.getDataRange().getValues();
        const infoContaDestino = obterInformacoesDaConta(transacaoData.contaDestino, contasSheetData);
        
        let valorAjusteDestino = transacaoData.valor;
        // Se a conta de destino for um cartão de crédito, o valor da transferência (pagamento) deve REDUZIR a dívida.
        if (infoContaDestino && normalizarTexto(infoContaDestino.tipo) === "cartao de credito") {
            valorAjusteDestino = -transacaoData.valor;
            logToSheet(`Ajuste de transferência para cartão de crédito detectado. Valor de ajuste para destino: ${valorAjusteDestino}`, "INFO");
        }
        
        ajustarSaldoIncrementalmente(contasSheet, transacaoData.contaOrigem, -transacaoData.valor, transacaoData.contaDestino, valorAjusteDestino);
        // --- FIM DA MELHORIA ---
        enviarMensagemTelegram(chatId, `✅ Transferência de *${formatCurrency(transacaoData.valor)}* registrada com sucesso!`);

    } else { // Despesa ou Receita
        const valorParcela = transacaoData.valor / transacaoData.parcelasTotais;
        const dataVencimentoBase = new Date(transacaoData.dataVencimento);
        const dataTransacaoBase = new Date(transacaoData.data);
        const dataRegistroBase = new Date(transacaoData.dataRegistro);
        const dataTransacaoFormatada = `'${Utilities.formatDate(dataTransacaoBase, timezone, "dd/MM/yyyy")}`;
        const dataRegistroFormatada = `'${Utilities.formatDate(dataRegistroBase, timezone, "dd/MM/yyyy HH:mm:ss")}`;

        for (let i = 0; i < transacaoData.parcelasTotais; i++) {
            let dataVencimentoParcela = new Date(dataVencimentoBase);
            dataVencimentoParcela.setMonth(dataVencimentoBase.getMonth() + i);
            if (dataVencimentoParcela.getDate() !== dataVencimentoBase.getDate()) {
                dataVencimentoParcela = new Date(dataVencimentoParcela.getFullYear(), dataVencimentoParcela.getMonth() + 1, 0);
            }
            rowsToAdd.push([
              dataTransacaoFormatada, transacaoData.descricao, transacaoData.categoria, transacaoData.subcategoria,
              transacaoData.tipo, valorParcela, transacaoData.metodoPagamento, transacaoData.conta,
              transacaoData.parcelasTotais, i + 1, `'${Utilities.formatDate(dataVencimentoParcela, timezone, "dd/MM/yyyy")}`, usuario, "Ativo", `${transacaoData.finalId}-${i + 1}`, dataRegistroFormatada
            ]);
        }
        
        if (rowsToAdd.length > 0) {
            transacoesSheet.getRange(transacoesSheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
        }

        // --- INÍCIO DA MELHORIA: Lógica de Ajuste Incremental Unificada ---
        const infoConta = obterInformacoesDaConta(transacaoData.conta, contasSheet.getDataRange().getValues());
        if (infoConta) {
          let valorAjuste;
          
          if (normalizarTexto(infoConta.tipo) === "cartao de credito") {
            // Para cartões de crédito, despesas aumentam a dívida (saldo pendente).
            // Receitas (estornos) diminuem a dívida.
            valorAjuste = transacaoData.tipo === 'Receita' ? -transacaoData.valor : transacaoData.valor;
            
            // A dívida total é adicionada de uma vez, independentemente das parcelas.
            ajustarSaldoIncrementalmente(contasSheet, transacaoData.conta, valorAjuste);
            logToSheet(`Ajuste incremental de DÍVIDA aplicado para '${transacaoData.conta}'. Valor: ${valorAjuste}`, "INFO");
          
          } else {
            // Para contas normais (débito, dinheiro), o valor total da transação impacta o saldo.
            valorAjuste = transacaoData.tipo === 'Receita' ? transacaoData.valor : -transacaoData.valor;
            ajustarSaldoIncrementalmente(contasSheet, transacaoData.conta, valorAjuste);
            logToSheet(`Ajuste incremental de SALDO aplicado para '${transacaoData.conta}'. Valor: ${valorAjuste}`, "INFO");
          }
        } else {
            logToSheet(`AVISO: Conta '${transacaoData.conta}' não encontrada para ajuste de saldo incremental.`, "WARN");
        }
        // --- FIM DA MELHORIA ---
        
        enviarMensagemTelegram(chatId, `✅ Lançamento de *${formatCurrency(transacaoData.valor)}* (${transacaoData.parcelasTotais}x) registrado com sucesso!`);
        
        // =================================================================
        // ### INÍCIO DA NOVA LÓGICA DE APRENDIZADO ###
        // =================================================================
        // Se a transação foi originalmente classificada como "Não Identificada",
        // significa que o usuário a corrigiu através do assistente.
        if (transacaoData.originalCategory === "Não Identificada") {
            logToSheet(`[Learning] Gatilho de aprendizado acionado a partir do assistente para a descrição: "${transacaoData.descricao}"`, "INFO");
            learnFromCorrection(transacaoData.descricao, transacaoData.categoria, transacaoData.subcategoria);
        }
        // =================================================================
        // ### FIM DA NOVA LÓGICA DE APRENDIZADO ###
        // =================================================================
    }
    
    updateBudgetSpentValues();

  } catch (e) {
    handleError(e, "registrarTransacaoConfirmada", chatId);
  } finally {
    lock.releaseLock();
  }
}


/**
 * Cancela uma transação pendente.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} transactionId O ID da transação pendente.
 */
function cancelarTransacaoPendente(chatId, transactionId) {
  enviarMensagemTelegram(chatId, "❌ Lançamento cancelado.");
  logToSheet(`Transacao ${transactionId} cancelada por ${chatId}.`, "INFO");
}


/**
 * ATUALIZADO: Calcula a data de vencimento da fatura do cartão de crédito para uma transação.
 * @param {Object} infoConta O objeto de informações da conta (do 'Contas.gs').
 * @param {Date} transactionDate A data da transacao a ser usada como referencia.
 * @param {Array<Array<any>>} dadosContas Os dados da aba 'Contas'.
 * @returns {Date} A data de vencimento calculada.
 */
function calcularVencimentoCartao(infoConta, transactionDate, dadosContas) {
    const diaTransacao = transactionDate.getDate();
    const mesTransacao = transactionDate.getMonth();
    const anoTransacao = transactionDate.getFullYear();

    const diaFechamento = infoConta.diaFechamento;
    const diaVencimento = infoConta.vencimento;
    const tipoFechamento = infoConta.tipoFechamento || "padrao";

    logToSheet(`[CalcVencimento] Calculando vencimento para ${infoConta.nomeOriginal}. Transacao em: ${transactionDate.toLocaleDateString()}, Dia Fechamento: ${diaFechamento}, Dia Vencimento: ${diaVencimento}, Tipo Fechamento: ${tipoFechamento}`, "DEBUG");

    let mesFechamento;
    let anoFechamento;

    if (tipoFechamento === "padrao" || tipoFechamento === "fechamento-mes") {
        if (diaTransacao <= diaFechamento) {
            mesFechamento = mesTransacao;
            anoFechamento = anoTransacao;
        } else {
            mesFechamento = mesTransacao + 1;
            anoFechamento = anoTransacao;
        }
    } else if (tipoFechamento === "fechamento-anterior") {
        mesFechamento = mesTransacao;
        anoFechamento = anoTransacao;
    } else {
        logToSheet(`[CalcVencimento] Tipo de fechamento desconhecido: ${tipoFechamento}. Assumindo padrao.`, "WARN");
        if (diaTransacao <= diaFechamento) {
            mesFechamento = mesTransacao;
            anoFechamento = anoTransacao;
        } else {
            mesFechamento = mesTransacao + 1;
            anoFechamento = anoTransacao;
        }
    }

    let vencimentoAno = anoFechamento;
    let vencimentoMes = mesFechamento + 1;

    if (vencimentoMes > 11) {
        vencimentoMes -= 12;
        vencimentoAno++;
    }

    let dataVencimento = new Date(vencimentoAno, vencimentoMes, diaVencimento);

    if (dataVencimento.getMonth() !== vencimentoMes) {
        dataVencimento = new Date(vencimentoAno, vencimentoMes + 1, 0);
    }
    
    logToSheet(`[CalcVencimento] Data de Vencimento Final Calculada: ${dataVencimento.toLocaleDateString()}`, "DEBUG");
    return dataVencimento;
}

/**
 * NOVO: Calcula a data de vencimento da fatura do cartão de crédito para uma PARCELA específica.
 * Essencial para garantir que cada parcela tenha a data de vencimento correta.
 * @param {Object} infoConta O objeto de informações da conta (do 'Contas.gs').
 * @param {Date} dataPrimeiraParcelaVencimento A data de vencimento da primeira parcela (já calculada por calcularVencimentoCartao).
 * @param {number} numeroParcela O número da parcela atual (1, 2, 3...).
 * @param {number} totalParcelas O número total de parcelas.
 * @param {Array<Array<any>>} dadosContas Os dados da aba 'Contas'.
 * @returns {Date} A data de vencimento calculada para a parcela.
 */
function calcularVencimentoCartaoParaParcela(infoConta, dataPrimeiraParcelaVencimento, numeroParcela, totalParcelas, dadosContas) {
    if (numeroParcela === 1) {
        return dataPrimeiraParcelaVencimento;
    }

    // Começa com a data de vencimento da primeira parcela
    let dataVencimentoParcela = new Date(dataPrimeiraParcelaVencimento);

    // Adiciona o número de meses correspondente à parcela
    dataVencimentoParcela.setMonth(dataVencimentoParcela.getMonth() + (numeroParcela - 1));

    // Ajuste para garantir que o dia do vencimento não "pule" para o mês seguinte
    if (dataVencimentoParcela.getDate() !== dataPrimeiraParcelaVencimento.getDate()) {
        const lastDayOfMonth = new Date(dataVencimentoParcela.getFullYear(), dataVencimentoParcela.getMonth() + 1, 0).getDate();
        dataVencimentoParcela.setDate(Math.min(dataVencimentoParcela.getDate(), lastDayOfMonth));
    }
    logToSheet(`[CalcVencimentoParcela] Calculado vencimento para parcela ${numeroParcela} de ${infoConta.nomeOriginal}: ${dataVencimentoParcela.toLocaleDateString()}`, "DEBUG");
    return dataVencimentoParcela;
}

/**
 * ATUALIZADO: Atualiza os saldos de todas as contas na planilha 'Contas'
 * e os armazena na variável global `globalThis.saldosCalculados`.
 */
function atualizarSaldosDasContas() {
  // Adiciona um bloqueio para garantir que apenas uma instância desta função seja executada de cada vez.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Espera até 30 segundos pelo acesso exclusivo.

  try {
    logToSheet("Iniciando atualizacao de saldos das contas.", "INFO");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contasSheet = ss.getSheetByName(SHEET_CONTAS);
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
    
    if (!contasSheet || !transacoesSheet) {
      throw new Error("Aba 'Contas' ou 'Transacoes' não encontrada para atualização de saldos.");
    }

    const dadosContas = contasSheet.getDataRange().getValues();
    const dadosTransacoes = transacoesSheet.getDataRange().getValues();
    
    globalThis.saldosCalculados = {}; // Limpa os saldos anteriores

    // --- PASSO 1: Inicializa todas as contas ---
    for (let i = 1; i < dadosContas.length; i++) {
      const linha = dadosContas[i];
      const nomeOriginal = (linha[0] || "").toString().trim();
      if (!nomeOriginal) continue;

      const nomeNormalizado = normalizarTexto(nomeOriginal);
      globalThis.saldosCalculados[nomeNormalizado] = {
        nomeOriginal: nomeOriginal,
        nomeNormalizado: nomeNormalizado,
        tipo: (linha[1] || "").toString().toLowerCase().trim(),
        saldo: parseBrazilianFloat(String(linha[3] || '0')), // Saldo Inicial
        limite: parseBrazilianFloat(String(linha[5] || '0')),
        vencimento: parseInt(linha[6]) || null,
        diaFechamento: parseInt(linha[9]) || null,
        tipoFechamento: (linha[10] || "").toString().trim(),
        contaPaiAgrupador: normalizarTexto((linha[12] || "").toString().trim()),
        faturaAtual: 0, 
        saldoTotalPendente: 0
      };
    }
    logToSheet("[AtualizarSaldos] Passo 1/4: Contas inicializadas.", "DEBUG");


    // --- PASSO 2: Processa transações para calcular saldos individuais ---
    const today = new Date();
    let nextCalendarMonth = today.getMonth() + 1;
    let nextCalendarYear = today.getFullYear();
    if (nextCalendarMonth > 11) {
        nextCalendarMonth = 0;
        nextCalendarYear++;
    }

    for (let i = 1; i < dadosTransacoes.length; i++) {
      const linha = dadosTransacoes[i];
      const tipoTransacao = (linha[4] || "").toString().toLowerCase().trim();
      const valor = parseBrazilianFloat(String(linha[5] || '0'));
      const contaNormalizada = normalizarTexto(linha[7] || "");
      const categoria = normalizarTexto(linha[2] || "");
      const subcategoria = normalizarTexto(linha[3] || "");
      const dataVencimento = parseData(linha[10]);

      if (!globalThis.saldosCalculados[contaNormalizada]) continue;

      const infoConta = globalThis.saldosCalculados[contaNormalizada];

      if (infoConta.tipo === "conta corrente" || infoConta.tipo === "dinheiro físico") {
        if (tipoTransacao === "receita") infoConta.saldo += valor;
        else if (tipoTransacao === "despesa") infoConta.saldo -= valor;
      } else if (infoConta.tipo === "cartão de crédito") {
        const isPayment = (categoria === "contas a pagar" && subcategoria === "pagamento de fatura");
        if (isPayment) {
          infoConta.saldoTotalPendente -= valor;
        } else if (tipoTransacao === "despesa") {
          infoConta.saldoTotalPendente += valor;
          if (dataVencimento && dataVencimento.getMonth() === nextCalendarMonth && dataVencimento.getFullYear() === nextCalendarYear) {
            infoConta.faturaAtual += valor;
          }
        }
      }
    }
    logToSheet("[AtualizarSaldos] Passo 2/4: Saldos individuais calculados.", "DEBUG");


    // --- PASSO 3: Consolida saldos de cartões em 'Faturas Consolidadas' ---
    for (const nomeNormalizado in globalThis.saldosCalculados) {
      const infoConta = globalThis.saldosCalculados[nomeNormalizado];
      if (infoConta.tipo === "cartão de crédito" && infoConta.contaPaiAgrupador) {
        const agrupadorNormalizado = infoConta.contaPaiAgrupador;
        if (globalThis.saldosCalculados[agrupadorNormalizado] && globalThis.saldosCalculados[agrupadorNormalizado].tipo === "fatura consolidada") {
          const agrupador = globalThis.saldosCalculados[agrupadorNormalizado];
          agrupador.saldoTotalPendente += infoConta.saldoTotalPendente;
          agrupador.faturaAtual += infoConta.faturaAtual;
        }
      }
    }
    logToSheet("[AtualizarSaldos] Passo 3/4: Saldos consolidados.", "DEBUG");


    // --- PASSO 4: Atualiza a planilha 'Contas' com os novos saldos ---
    const saldosParaPlanilha = [];
    for (let i = 1; i < dadosContas.length; i++) {
      const nomeOriginal = (dadosContas[i][0] || "").toString().trim();
      const nomeNormalizado = normalizarTexto(nomeOriginal);
      if (globalThis.saldosCalculados[nomeNormalizado]) {
        const infoConta = globalThis.saldosCalculados[nomeNormalizado];
        let saldoFinal;
        if (infoConta.tipo === "fatura consolidada" || infoConta.tipo === "cartão de crédito") {
          saldoFinal = infoConta.saldoTotalPendente;
        } else {
          saldoFinal = infoConta.saldo;
        }
        saldosParaPlanilha.push([round(saldoFinal, 2)]);
      } else {
        saldosParaPlanilha.push([dadosContas[i][4]]); // Mantém o valor antigo se a conta não foi encontrada
      }
    }

    if (saldosParaPlanilha.length > 0) {
      // Coluna E (índice 4) é a 'Saldo Atualizado'
      contasSheet.getRange(2, 5, saldosParaPlanilha.length, 1).setValues(saldosParaPlanilha);
    }
    logToSheet("[AtualizarSaldos] Passo 4/4: Planilha 'Contas' atualizada.", "INFO");

  } catch (e) {
    handleError(e, "atualizarSaldosDasContas");
  } finally {
    lock.releaseLock(); // Libera o bloqueio para que outras operações possam ser executadas.
  }
}


/**
 * NOVO: Gera as contas recorrentes para o próximo mês com base na aba 'Contas_a_Pagar'.
 */
function generateRecurringBillsForNextMonth() {
    logToSheet("Iniciando geracao de contas recorrentes para o proximo mes.", "INFO");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);
    
    if (!contasAPagarSheet) {
        logToSheet("Erro: Aba 'Contas_a_Pagar' nao encontrada para gerar contas recorrentes.", "ERROR");
        throw new Error("Aba 'Contas_a_Pagar' não encontrada.");
    }

    const dadosContasAPagar = contasAPagarSheet.getDataRange().getValues();
    const headers = dadosContasAPagar[0];

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

    if ([colID, colDescricao, colCategoria, colValor, colDataVencimento, colStatus, colRecorrente, colContaSugeria, colObservacoes, colIDTransacaoVinculada].some(idx => idx === -1)) {
        logToSheet("Erro: Colunas essenciais faltando na aba 'Contas_a_Pagar' para geracao de contas recorrentes.", "ERROR");
        throw new Error("Colunas essenciais faltando na aba 'Contas_a_Pagar'.");
    }

    const today = new Date();
    const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    const nextMonthNum = nextMonth.getMonth(); // 0-indexed
    const nextYearNum = nextMonth.getFullYear();

    logToSheet(`Gerando contas recorrentes para: ${getNomeMes(nextMonthNum)}/${nextYearNum}`, "DEBUG");

    const newBills = [];
    const existingBillsInNextMonth = new Set();

    for (let i = 1; i < dadosContasAPagar.length; i++) {
        const row = dadosContasAPagar[i];
        const dataVencimentoExistente = parseData(row[colDataVencimento]);
        if (dataVencimentoExistente &&
            dataVencimentoExistente.getMonth() === nextMonthNum &&
            dataVencimentoExistente.getFullYear() === nextYearNum) {
            existingBillsInNextMonth.add(normalizarTexto(row[colDescricao] + row[colValor] + row[colCategoria]));
        }
    }
    logToSheet(`Contas existentes no proximo mes: ${existingBillsInNextMonth.size}`, "DEBUG");


    for (let i = 1; i < dadosContasAPagar.length; i++) {
        const row = dadosContasAPagar[i];
        const recorrente = (row[colRecorrente] || "").toString().trim().toLowerCase();
        
        if (recorrente === "verdadeiro") {
            const currentDescricao = (row[colDescricao] || "").toString().trim();
            const currentValor = parseBrazilianFloat(String(row[colValor]));
            const currentCategoria = (row[colCategoria] || "").toString().trim();
            const currentDataVencimento = parseData(row[colDataVencimento]);
            const currentContaSugeria = (row[colContaSugeria] || "").toString().trim();
            const currentObservacoes = (row[colObservacoes] || "").toString().trim();
            
            const billKey = normalizarTexto(currentDescricao + currentValor + currentCategoria);

            if (existingBillsInNextMonth.has(billKey)) {
                logToSheet(`Conta recorrente "${currentDescricao}" ja existe para ${getNomeMes(nextMonthNum)}/${nextYearNum}. Pulando.`, "DEBUG");
                continue;
            }

            if (currentDataVencimento) {
                let newDueDate = new Date(currentDataVencimento);
                newDueDate.setMonth(newDueDate.getMonth() + 1);

                if (newDueDate.getDate() !== currentDataVencimento.getDate()) {
                    newDueDate = new Date(newDueDate.getFullYear(), newDueDate.getMonth() + 1, 0);
                }

                const newRow = [
                    Utilities.getUuid(),
                    currentDescricao,
                    currentCategoria,
                    currentValor,
                    Utilities.formatDate(newDueDate, Session.getScriptTimeZone(), "dd/MM/yyyy"),
                    "Pendente",
                    "Verdadeiro",
                    currentContaSugeria,
                    currentObservacoes,
                    ""
                ];
                newBills.push(newRow);
                logToSheet(`Conta recorrente "${currentDescricao}" gerada para ${getNomeMes(newDueDate.getMonth())}/${newDueDate.getFullYear()}.`, "INFO");
            }
        }
    }

    if (newBills.length > 0) {
        contasAPagarSheet.getRange(contasAPagarSheet.getLastRow() + 1, 1, newBills.length, newBills[0].length).setValues(newBills);
        logToSheet(`Total de ${newBills.length} contas recorrentes adicionadas.`, "INFO");
    } else {
        logToSheet("Nenhuma nova conta recorrente para adicionar para o proximo mes.", "INFO");
    }
}

/**
 * NOVO: Processa o comando /marcar_pago vindo do Telegram.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} textoRecebido O texto completo do comando.
 * @param {string} usuario O nome do usuário.
 */
function processarMarcarPago(chatId, textoRecebido, usuario) {
  const idContaAPagar = textoRecebido.substring("/marcar_pago_".length);
  logToSheet(`[MarcarPago] Processando marcar pago para ID: ${idContaAPagar}`, "INFO");

  const contaAPagarInfo = obterInformacoesDaContaAPagar(idContaAPagar);

  if (!contaAPagarInfo) {
    enviarMensagemTelegram(chatId, `❌ Conta a Pagar com ID *${escapeMarkdown(idContaAPagar)}* não encontrada.`);
    return;
  }

  if (normalizarTexto(contaAPagarInfo.status) === "pago") {
    enviarMensagemTelegram(chatId, `ℹ️ A conta *${escapeMarkdown(contaAPagarInfo.descricao)}* já está paga.`);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
  const dadosTransacoes = transacoesSheet.getDataRange().getValues();

  let transacaoVinculada = null;
  const hoje = new Date();
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();

  for (let i = 1; i < dadosTransacoes.length; i++) {
    const linha = dadosTransacoes[i];
    const dataTransacao = parseData(linha[0]);
    const descricaoTransacao = normalizarTexto(linha[1]);
    const valorTransacao = parseBrazilianFloat(String(linha[5]));
    const idTransacao = linha[13];

    if (dataTransacao && dataTransacao.getMonth() === mesAtual && dataTransacao.getFullYear() === anoAtual &&
        normalizarTexto(linha[4]) === "despesa" &&
        calculateSimilarity(descricaoTransacao, normalizarTexto(contaAPagarInfo.descricao)) > SIMILARITY_THRESHOLD &&
        Math.abs(valorTransacao - contaAPagarInfo.valor) < 0.01) {
        transacaoVinculada = idTransacao;
        break;
    }
  }

  if (transacaoVinculada) {
    vincularTransacaoAContaAPagar(chatId, idContaAPagar, transacaoVinculada);
  } else {
    const mensagem = `A conta *${escapeMarkdown(contaAPagarInfo.descricao)}* (R$ ${contaAPagarInfo.valor.toFixed(2).replace('.', ',')}) será marcada como paga.`;
    const teclado = {
      inline_keyboard: [
        [{ text: "✅ Marcar como Pago (sem registrar transação)", callback_data: `confirm_marcar_pago_sem_transacao_${idContaAPagar}` }],
        [{ text: "📝 Registrar e Marcar como Pago", callback_data: `confirm_marcar_pago_e_registrar_${idContaAPagar}` }],
        [{ text: "❌ Cancelar", callback_data: `cancel_${idContaAPagar}` }]
      ]
    };
    enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
  }
}

/**
 * NOVO: Função para lidar com a confirmação de marcar conta a pagar.
 * @param {string} chatId O ID do chat.
 * @param {string} action A ação a ser tomada.
 * @param {string} idContaAPagar O ID da conta.
 * @param {string} usuario O nome do usuário.
 */
function handleMarcarPagoConfirmation(chatId, action, idContaAPagar, usuario) {
  logToSheet(`[MarcarPagoConfirm] Acão: ${action}, ID Conta: ${idContaAPagar}, Usuario: ${usuario}`, "INFO");

  const contaAPagarInfo = obterInformacoesDaContaAPagar(idContaAPagar);

  if (!contaAPagarInfo) {
    enviarMensagemTelegram(chatId, `❌ Conta a Pagar com ID *${escapeMarkdown(idContaAPagar)}* não encontrada.`);
    return;
  }

  if (normalizarTexto(contaAPagarInfo.status) === "pago") {
    enviarMensagemTelegram(chatId, `ℹ️ A conta *${escapeMarkdown(contaAPagarInfo.descricao)}* já está paga.`);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);
  const colStatus = contaAPagarInfo.headers.indexOf('Status') + 1;
  const colIDTransacaoVinculada = contaAPagarInfo.headers.indexOf('ID Transacao Vinculada') + 1;

  if (action === "sem_transacao") {
    try {
      contasAPagarSheet.getRange(contaAPagarInfo.linha, colStatus).setValue("Pago");
      contasAPagarSheet.getRange(contaAPagarInfo.linha, colIDTransacaoVinculada).setValue("MARCADO_MANUALMENTE");
      enviarMensagemTelegram(chatId, `✅ Conta *${escapeMarkdown(contaAPagarInfo.descricao)}* marcada como paga (sem registro de transação).`);
      atualizarSaldosDasContas();
    } catch (e) {
      enviarMensagemTelegram(chatId, `❌ Erro ao marcar conta como paga: ${e.message}`);
    }
  } else if (action === "e_registrar") {
    try {
      const transacaoData = {
        id: Utilities.getUuid(),
        data: new Date(),
        descricao: `Pagamento de ${contaAPagarInfo.descricao}`,
        categoria: contaAPagarInfo.categoria,
        subcategoria: "Pagamento de Fatura" || "",
        tipo: "Despesa",
        valor: contaAPagarInfo.valor,
        metodoPagamento: contaAPagarInfo.contaDePagamentoSugeria || "Débito",
        conta: contaAPagarInfo.contaDePagamentoSugeria || "Não Identificada",
        parcelasTotais: 1,
        parcelaAtual: 1,
        dataVencimento: contaAPagarInfo.dataVencimento,
        usuario: usuario,
        status: "Ativo",
        dataRegistro: new Date()
      };
      
      registrarTransacaoConfirmada(transacaoData, usuario, chatId);
      
      contasAPagarSheet.getRange(contaAPagarInfo.linha, colStatus).setValue("Pago");
      contasAPagarSheet.getRange(contaAPagarInfo.linha, colIDTransacaoVinculada).setValue(transacaoData.id);
      enviarMensagemTelegram(chatId, `✅ Transação de *${formatCurrency(transacaoData.valor)}* para *${escapeMarkdown(contaAPagarInfo.descricao)}* registrada e conta marcada como paga!`);
      atualizarSaldosDasContas();
    } catch (e) {
      enviarMensagemTelegram(chatId, `❌ Erro ao registrar e marcar conta como paga: ${e.message}`);
    }
  }
}

/**
 * NOVO: Extrai as contas de origem e destino de uma mensagem de transferência.
 * @param {string} textoNormalizado O texto normalizado.
 * @param {Array<Array<any>>} dadosContas Os dados das contas.
 * @param {Array<Array<any>>} dadosPalavras Os dados das palavras-chave.
 * @returns {Object} Um objeto com as contas de origem e destino.
 */
function extrairContasTransferencia(textoNormalizado, dadosContas, dadosPalavras) {
    let contaOrigem = "Não Identificada";
    let contaDestino = "Não Identificada";

    const matchOrigem = textoNormalizado.match(/(?:de|do)\s(.*?)(?=\s(?:para|pra)|$)/);
    const matchDestino = textoNormalizado.match(/(?:para|pra)\s(.+)/);

    if (matchOrigem && matchOrigem[1]) {
        const { conta } = extrairContaMetodoPagamento(matchOrigem[1].trim(), dadosContas, dadosPalavras);
        contaOrigem = conta;
    }

    if (matchDestino && matchDestino[1]) {
        const { conta } = extrairContaMetodoPagamento(matchDestino[1].trim(), dadosContas, dadosPalavras);
        contaDestino = conta;
    }

    return { contaOrigem, contaDestino };
}


/**
 * NOVO: Prepara e envia uma mensagem de confirmação para transferências.
 * @param {Object} transacaoData Os dados da transferência.
 * @param {string} chatId O ID do chat.
 * @returns {Object} O status de confirmação pendente.
 */
function prepararConfirmacaoTransferencia(transacaoData, chatId) {
    const transactionId = Utilities.getUuid();
    transacaoData.finalId = transactionId;
    transacaoData.data = new Date();

    const cache = CacheService.getScriptCache();
    const cacheKey = `${CACHE_KEY_PENDING_TRANSACTIONS}_${chatId}_${transactionId}`;
    cache.put(cacheKey, JSON.stringify(transacaoData), CACHE_EXPIRATION_PENDING_TRANSACTION_SECONDS);

    let mensagem = `✅ Confirme sua Transferência:\n\n`;
    mensagem += `*Valor:* ${formatCurrency(transacaoData.valor)}\n`;
    mensagem += `*De:* ${escapeMarkdown(transacaoData.contaOrigem)}\n`;
    mensagem += `*Para:* ${escapeMarkdown(transacaoData.contaDestino)}\n`;

    const teclado = {
        inline_keyboard: [
            [{ text: "✅ Confirmar", callback_data: `confirm_${transactionId}` }],
            [{ text: "❌ Cancelar", callback_data: `cancel_${transactionId}` }]
        ]
    };

    enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
    return { status: "PENDING_CONFIRMATION", transactionId: transactionId };
}

/**
 * NOVO E OTIMIZADO: Ajusta o saldo de uma ou duas contas de forma incremental.
 * Lê o saldo atual, soma/subtrai o valor da nova transação e escreve o resultado.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} contasSheet A aba 'Contas' já aberta.
 * @param {string} nomeConta O nome da conta principal a ser ajustada.
 * @param {number} valor O valor da transação (positivo para receitas, negativo para despesas).
 * @param {string} [nomeContaSecundaria] O nome da conta secundária (para transferências).
 * @param {number} [valorSecundario] O valor para a conta secundária (geralmente o oposto do principal).
 */
function ajustarSaldoIncrementalmente(contasSheet, nomeConta, valor, nomeContaSecundaria = null, valorSecundario = 0) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const dadosContas = contasSheet.getDataRange().getValues();
    const headers = dadosContas[0];
    const colNome = headers.indexOf('Nome da Conta');
    const colSaldoAtual = headers.indexOf('Saldo Atual');

    if (colNome === -1 || colSaldoAtual === -1) {
      logToSheet("ERRO: Colunas 'Nome da Conta' ou 'Saldo Atual' não encontradas para ajuste incremental.", "ERROR");
      return;
    }

    // Mapeia todas as contas para encontrar as linhas corretas
    const contasParaAtualizar = {};
    contasParaAtualizar[normalizarTexto(nomeConta)] = { valorAjuste: valor, rowIndex: -1 };
    if (nomeContaSecundaria) {
      contasParaAtualizar[normalizarTexto(nomeContaSecundaria)] = { valorAjuste: valorSecundario, rowIndex: -1 };
    }

    for (let i = 1; i < dadosContas.length; i++) {
      const nomePlanilha = normalizarTexto(dadosContas[i][colNome]);
      if (contasParaAtualizar[nomePlanilha]) {
        contasParaAtualizar[nomePlanilha].rowIndex = i + 1; // Linha base 1
      }
    }

    // Itera sobre as contas encontradas e atualiza o saldo
    for (const nomeNormalizado in contasParaAtualizar) {
      const info = contasParaAtualizar[nomeNormalizado];
      if (info.rowIndex !== -1) {
        const saldoAtualRange = contasSheet.getRange(info.rowIndex, colSaldoAtual + 1);
        const saldoAtual = parseFloat(saldoAtualRange.getValue()) || 0;
        const novoSaldo = saldoAtual + info.valorAjuste;
        saldoAtualRange.setValue(novoSaldo);
        logToSheet(`Saldo da conta '${nomeNormalizado}' ajustado incrementalmente. Novo saldo: ${novoSaldo}`, "INFO");
      } else {
        logToSheet(`AVISO: Conta '${nomeNormalizado}' não encontrada para ajuste de saldo incremental.`, "WARN");
      }
    }

  } catch (e) {
    logToSheet(`ERRO em ajustarSaldoIncrementalmente: ${e.message}`, "ERROR");
  } finally {
    lock.releaseLock();
  }
}

/**
 * CORRIGIDO: Processa a resposta digitada pelo usuário quando o assistente está ativo.
 * Agora usa uma busca direta de contas (obterInformacoesDaConta) em vez do parser de frase completa.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 * @param {string} textoRecebido O texto digitado pelo usuário.
 * @param {Object} assistantState O estado atual do assistente.
 */
function processarRespostaDoAssistente(chatId, usuario, textoRecebido, assistantState) {
  const campoEsperado = assistantState.waitingFor;
  let transacaoParcial = assistantState; // O estado é a própria transação parcial
  let valorValido = false;

  logToSheet(`[Assistente] Processando resposta digitada para '${campoEsperado}': "${textoRecebido}"`, "INFO");

  const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);
  const dadosPalavras = getSheetDataWithCache(SHEET_PALAVRAS_CHAVE, CACHE_KEY_PALAVRAS);

  switch (campoEsperado) {
    case 'valor':
      const valorExtraido = extrairValor(textoRecebido);
      if (!isNaN(valorExtraido) && valorExtraido > 0) {
        transacaoParcial.valor = valorExtraido;
        valorValido = true;
      }
      break;
      
    case 'conta':
    case 'conta_origem':
    case 'conta_destino':
      const infoContaEncontrada = obterInformacoesDaConta(textoRecebido, dadosContas);
      
      if (infoContaEncontrada) {
        const nomeConta = infoContaEncontrada.nomeOriginal;
        
        if (campoEsperado === 'conta') {
          transacaoParcial.conta = nomeConta;
          transacaoParcial.infoConta = infoContaEncontrada;
        } else {
          transacaoParcial[campoEsperado] = nomeConta;
        }
        valorValido = true;
        logToSheet(`[Assistente] Conta "${nomeConta}" reconhecida com sucesso a partir da resposta digitada.`, "INFO");
      }
      break;

    case 'metodo':
      // ### INÍCIO DA LÓGICA CORRIGIDA ###
      const textoNormalizado = normalizarTexto(textoRecebido);
      
      // Procura por uma correspondência nas palavras-chave de método de pagamento
      for (let i = 1; i < dadosPalavras.length; i++) {
        const tipoPalavra = (dadosPalavras[i][0] || "").toLowerCase().trim();
        const palavraChave = normalizarTexto(dadosPalavras[i][1] || "");
        
        if (tipoPalavra === 'meio_pagamento' && textoNormalizado.includes(palavraChave)) {
          const valorInterpretado = (dadosPalavras[i][2] || "").trim();
          transacaoParcial.metodoPagamento = valorInterpretado;
          valorValido = true;
          logToSheet(`[Assistente] Método de pagamento "${valorInterpretado}" reconhecido com sucesso a partir da resposta digitada "${textoRecebido}".`, "INFO");
          break; // Encontrou, pode parar de procurar
        }
      }
      // ### FIM DA LÓGICA CORRIGIDA ###
      break;
      
    // Adicione mais cases para 'categoria', 'subcategoria', etc., aqui se desejar no futuro.
  }

  if (valorValido) {
    // Limpa o estado 'waitingFor' e continua o fluxo normal
    delete transacaoParcial.waitingFor;
    clearActiveAssistantState(chatId); // Limpa o ponteiro e o estado antigo
    processAssistantCompletion(transacaoParcial, chatId, usuario);
  } else {
    // Se o valor digitado for inválido, pede novamente
    enviarMensagemTelegram(chatId, `Peço desculpa, sou o Zaq. Não consegui reconhecer "${escapeMarkdown(textoRecebido)}" como uma resposta válida. Por favor, tente novamente ou escolha uma das opções.`);
  }
}

/**
 * NOVO: Reverte (subtrai) o valor de um aporte na aba 'Metas'.
 * @param {string} nomeMeta O nome da meta a ser revertida.
 * @param {number} valorReverter O valor a ser subtraído do 'Valor Salvo'.
 */
function reverterAporteMeta(nomeMeta, valorReverter) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const metasSheet = ss.getSheetByName(SHEET_METAS);
    if (!metasSheet) {
      logToSheet(`[reverterAporteMeta] Aba 'Metas' não encontrada. Reversão falhou.`, "WARN");
      return;
    }

    const dadosMetas = metasSheet.getDataRange().getValues();
    const headers = dadosMetas[0];
    const colMap = getColumnMap(headers);

    if (colMap['Nome da Meta'] === undefined || colMap['Valor Salvo'] === undefined) {
      logToSheet(`[reverterAporteMeta] Colunas 'Nome da Meta' ou 'Valor Salvo' não encontradas.`, "WARN");
      return;
    }

    const nomeMetaNormalizado = normalizarTexto(nomeMeta);
    let rowIndex = -1;

    for (let i = 1; i < dadosMetas.length; i++) {
      if (normalizarTexto(dadosMetas[i][colMap['Nome da Meta']]) === nomeMetaNormalizado) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex !== -1) {
      const valorSalvoAtual = parseBrazilianFloat(String(dadosMetas[rowIndex - 1][colMap['Valor Salvo']] || '0'));
      const novoValorSalvo = valorSalvoAtual - valorReverter;
      metasSheet.getRange(rowIndex, colMap['Valor Salvo'] + 1).setValue(novoValorSalvo);
      logToSheet(`Reversão de aporte na meta '${nomeMeta}'. Valor salvo atualizado de ${valorSalvoAtual} para ${novoValorSalvo}.`, "INFO");
    } else {
      logToSheet(`[reverterAporteMeta] Meta '${nomeMeta}' não encontrada para reverter o valor.`, "WARN");
    }
  } catch (e) {
    handleError(e, "reverterAporteMeta");
  }
}
/**
 * Reverte o status de uma conta na aba 'Contas_a_Pagar' para 'Pendente'
 * se ela estiver vinculada a um ID de transação que foi excluído.
 * @param {string} idLancamento O ID da transação que foi excluída.
 */
function reverterStatusContaAPagarSeVinculado(idLancamento) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);
    if (!contasAPagarSheet) return;

    const dadosContasAPagar = contasAPagarSheet.getDataRange().getValues();
    const headers = dadosContasAPagar[0];
    const colStatus = headers.indexOf('Status');
    const colIDTransacao = headers.indexOf('ID Transacao Vinculada');

    if (colStatus === -1 || colIDTransacao === -1) return;

    for (let i = 1; i < dadosContasAPagar.length; i++) {
      if (dadosContasAPagar[i][colIDTransacao] === idLancamento) {
        const linhaParaAtualizar = i + 1;
        contasAPagarSheet.getRange(linhaParaAtualizar, colStatus + 1).setValue("Pendente");
        contasAPagarSheet.getRange(linhaParaAtualizar, colIDTransacao + 1).setValue("");
        logToSheet(`Status da conta a pagar (linha ${linhaParaAtualizar}) revertido para 'Pendente'.`, "INFO");
        break;
      }
    }
  } catch (e) {
    handleError(e, "reverterStatusContaAPagarSeVinculado");
  }
}

/**
 * Calcula e retorna um resumo da saúde financeira do usuário.
 * @param {string} chatId O ID do chat do usuário.
 * @returns {string} Uma mensagem formatada com o resumo da saúde financeira.
 */
function getSaudeFinanceira(chatId) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
    const dadosTransacoes = transacoesSheet.getDataRange().getValues();
    
    const hoje = new Date();
    const primeiroDiaDoMes = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
    const ultimoDiaDoMes = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);

    let totalReceitas = 0;
    let totalDespesas = 0;
    let despesasEssenciais = 0;
    let despesasNaoEssenciais = 0;

    const categoriasMap = getCategoriesMap();

    for (let i = 1; i < dadosTransacoes.length; i++) {
        const row = dadosTransacoes[i];
        const dataTransacao = parseData(row[0]);

        if (dataTransacao >= primeiroDiaDoMes && dataTransacao <= ultimoDiaDoMes) {
            const tipo = row[4];
            const valor = parseBrazilianFloat(String(row[5]));
            const categoria = normalizarTexto(row[2]);
            const subcategoria = normalizarTexto(row[3]);

            const isIgnored = (categoria === "contas a pagar" && subcategoria === "pagamento de fatura") ||
                              (categoria === "transferencias" && subcategoria === "entre contas");

            if (!isIgnored) {
                if (tipo === 'Receita') {
                    totalReceitas += valor;
                } else if (tipo === 'Despesa') {
                    totalDespesas += valor;
                    const categoriaInfo = categoriasMap[categoria];
                    if (categoriaInfo && categoriaInfo.tipoGasto === 'necessidade') {
                        despesasEssenciais += valor;
                    } else {
                        despesasNaoEssenciais += valor;
                    }
                }
            }
        }
    }

    const saldoMes = totalReceitas - totalDespesas;
    const taxaDePoupanca = totalReceitas > 0 ? (saldoMes / totalReceitas) * 100 : 0;
    const rendimentoComprometido = totalReceitas > 0 ? (totalDespesas / totalReceitas) * 100 : 0;
    const diasNoMes = hoje.getDate();
    const gastoDiarioMedio = diasNoMes > 0 ? totalDespesas / diasNoMes : 0;

    let mensagem = `? *Resumo da sua Saúde Financeira - ${getNomeMes(hoje.getMonth())}* 💚\n\n`;
    mensagem += `📊 *Visão Geral do Mês:*\n`;
    mensagem += `   - *Receitas:* ${formatarMoeda(totalReceitas)}\n`;
    mensagem += `   - *Despesas:* ${formatarMoeda(totalDespesas)}\n`;
    mensagem += `   - *Saldo:* ${formatarMoeda(saldoMes)}\n\n`;

    mensagem += `💡 *Indicadores Chave:*\n`;
    mensagem += `   - *Taxa de Poupança:* \`${taxaDePoupanca.toFixed(2)}%\`\n`;
    mensagem += `     _(O ideal é acima de 10-20%. Quanto maior, melhor!)_\n`;
    mensagem += `   - *Rendimento Comprometido:* \`${rendimentoComprometido.toFixed(2)}%\`\n`;
    mensagem += `     _(Mostra quanto da sua renda já foi gasta. O ideal é mantê-lo baixo.)_\n`;
    mensagem += `   - *Gasto Diário Médio:* \`${formatarMoeda(gastoDiarioMedio)}\`\n\n`;

    mensagem += `⚖️ *Balanço (Necessidades vs. Desejos):*\n`;
    mensagem += `   - *Gastos Essenciais:* ${formatarMoeda(despesasEssenciais)}\n`;
    mensagem += `   - *Gastos Não Essenciais:* ${formatarMoeda(despesasNaoEssenciais)}\n\n`;

    if (taxaDePoupanca < 10 && totalReceitas > 0) {
        mensagem += `⚠️ *Ponto de Atenção:* Sua taxa de poupança está um pouco baixa. Tente rever os gastos não essenciais para impulsionar suas economias!`;
    } else if (taxaDePoupanca > 20) {
        mensagem += `🎉 *Parabéns!* Você está com uma excelente taxa de poupança. Continue assim!`;
    }

    return mensagem;
}


/**
 * @private
 * Procura por uma categoria aprendida com base na descrição da transação.
 * Esta função consulta a "memória" do bot na aba 'LearnedCategories'.
 * @param {string} textoNormalizado A descrição normalizada da transação.
 * @returns {Object|null} Objeto com { categoria, subcategoria } se encontrar uma correspondência de alta confiança, ou null.
 */
function findLearnedCategory(textoNormalizado) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const learnedSheet = ss.getSheetByName(SHEET_LEARNED_CATEGORIES);
    if (!learnedSheet) return null; // Se a aba não existe, não há nada a aprender

    const data = learnedSheet.getDataRange().getValues();
    if (data.length < 2) return null; // Se a aba está vazia (só cabeçalho)

    const headers = data[0];
    const colMap = getColumnMap(headers);

    let bestMatch = null;
    // A confiança mínima para aplicar a regra automaticamente
    let highestScore = MIN_CONFIDENCE_TO_APPLY - 1; 

    // Itera por todas as regras aprendidas
    for (let i = 1; i < data.length; i++) {
        const keyword = data[i][colMap['Keyword']];
        const score = parseInt(data[i][colMap['ConfidenceScore']]) || 0;

        // Se a descrição contém a palavra-chave e a confiança é alta o suficiente
        if (textoNormalizado.includes(keyword) && score > highestScore) {
            highestScore = score;
            bestMatch = {
                categoria: data[i][colMap['Categoria']],
                subcategoria: data[i][colMap['Subcategoria']],
                keyword: keyword // Retorna a keyword que deu o match
            };
        }
    }
    
    if (bestMatch) {
        logToSheet(`[Learning] Categoria aprendida encontrada via keyword '${bestMatch.keyword}': ${bestMatch.categoria} > ${bestMatch.subcategoria}`, "INFO");
    }
    return bestMatch;
}

