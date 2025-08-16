/**
 * @file Management.gs
 * @description Este arquivo contém funções para gerenciar o estado do usuário,
 * incluindo o fluxo do tutorial, e comandos de gerenciamento de contas/categorias.
 */

/**
 * Define o estado do tutorial de um usuário no cache.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {number} step O número do passo do tutorial.
 * @param {string} [expectedAction] A ação esperada do usuário para este passo (ex: TUTORIAL_STATE_WAITING_DESPESA).
 * @param {number} [messageId] O ID da mensagem do tutorial para possível edição.
 */
function setTutorialState(chatId, step, expectedAction = "", messageId = null) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `${CACHE_KEY_TUTORIAL_STATE}_${chatId}`;
  const state = { currentStep: step, expectedAction: expectedAction, messageId: messageId };
  cache.put(cacheKey, JSON.stringify(state), CACHE_EXPIRATION_TUTORIAL_STATE_SECONDS);
  logToSheet(`Estado do tutorial para ${chatId} salvo: Passo ${step}, Acao Esperada: ${expectedAction}, Message ID: ${messageId}`, "DEBUG");
}

/**
 * Obtém o estado do tutorial de um usuário do cache.
 * @param {string} chatId O ID do chat do Telegram.
 * @returns {Object|null} O objeto de estado do tutorial (currentStep, expectedAction, messageId) ou null.
 */
function getTutorialState(chatId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `${CACHE_KEY_TUTORIAL_STATE}_${chatId}`;
  const cachedState = cache.get(cacheKey);
  if (cachedState) {
    const state = JSON.parse(cachedState);
    logToSheet(`Estado do tutorial para ${chatId} recuperado: Passo ${state.currentStep}, Acao Esperada: ${state.expectedAction}, Message ID: ${state.messageId}`, "DEBUG");
    return state;
  }
  logToSheet(`Nenhum estado de tutorial encontrado para ${chatId}.`, "DEBUG");
  return null;
}

/**
 * Limpa o estado do tutorial de um usuário do cache.
 * @param {string} chatId O ID do chat do Telegram.
 */
function clearTutorialState(chatId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `${CACHE_KEY_TUTORIAL_STATE}_${chatId}`;
  cache.remove(cacheKey);
  logToSheet(`Estado do tutorial para ${chatId} limpo.`, "INFO");
}

/**
 * NOVO: Define o estado de edição de um usuário no cache.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {Object} stateData Os dados do estado de edição (ex: { transactionId: '...', fieldToEdit: '...' }).
 */
function setEditState(chatId, stateData) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `${CACHE_KEY_EDIT_STATE}_${chatId}`;
  const state = {
    action: "awaiting_edit_input",
    ...stateData // Inclui transactionId e fieldToEdit
  };
  // Cache por 15 minutos (900 segundos)
  cache.put(cacheKey, JSON.stringify(state), CACHE_EXPIRATION_EDIT_STATE_SECONDS); // Usar a constante
  logToSheet(`Estado de EDIÇÃO para ${chatId} salvo: ${JSON.stringify(state)}`, "DEBUG");
}

/**
 * NOVO: Obtém o estado de edição de um usuário do cache.
 * @param {string} chatId O ID do chat do Telegram.
 * @returns {Object|null} O objeto de estado de edição ou null.
 */
function getEditState(chatId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `${CACHE_KEY_EDIT_STATE}_${chatId}`;
  const cachedState = cache.get(cacheKey);
  if (cachedState) {
    const state = JSON.parse(cachedState);
    logToSheet(`Estado de EDIÇÃO para ${chatId} recuperado: ${JSON.stringify(state)}`, "DEBUG");
    return state;
  }
  return null;
}

/**
 * NOVO: Limpa o estado de edição de um usuário do cache.
 * @param {string} chatId O ID do chat do Telegram.
 */
function clearEditState(chatId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `${CACHE_KEY_EDIT_STATE}_${chatId}`;
  cache.remove(cacheKey);
  logToSheet(`Estado de EDIÇÃO para ${chatId} limpo.`, "INFO");
}


/**
 * NOVO: Obtém o estado do assistente de um usuário do cache.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} transactionId O ID da transação parcial.
 * @returns {Object|null} O objeto de estado do assistente ou null.
 */
function getAssistantState(chatId, transactionId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `${CACHE_KEY_ASSISTANT_STATE}_${chatId}_${transactionId}`;
  const cachedState = cache.get(cacheKey);
  if (cachedState) {
    const state = JSON.parse(cachedState);
    logToSheet(`Estado do ASSISTENTE para ${chatId} (Transação ID ${transactionId}) recuperado.`, "DEBUG");
    return state;
  }
  return null;
}

/**
 * MODIFICADO: Define o estado do assistente e um ponteiro de estado ativo.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {Object} stateData Os dados da transação parcial.
 */
function setAssistantState(chatId, stateData) {
  const cache = CacheService.getScriptCache();
  const transactionId = stateData.id;
  
  // Salva o estado completo da transação
  const stateKey = `${CACHE_KEY_ASSISTANT_STATE}_${chatId}_${transactionId}`;
  cache.put(stateKey, JSON.stringify(stateData), CACHE_EXPIRATION_PENDING_TRANSACTION_SECONDS);

  // Salva um ponteiro simples indicando qual transação está ativa para este chat
  const pointerKey = `${CACHE_KEY_ASSISTANT_STATE}_active_${chatId}`;
  cache.put(pointerKey, stateKey, CACHE_EXPIRATION_PENDING_TRANSACTION_SECONDS);

  logToSheet(`Estado do ASSISTENTE para ${chatId} (Transação ID ${transactionId}) salvo. Ponteiro ativo definido.`, "DEBUG");
}

/**
 * NOVO: Obtém o estado do assistente ativo para um determinado chat.
 * @param {string} chatId O ID do chat do Telegram.
 * @returns {Object|null} O objeto de estado do assistente ou null se não houver estado ativo.
 */
function getActiveAssistantState(chatId) {
  const cache = CacheService.getScriptCache();
  const pointerKey = `${CACHE_KEY_ASSISTANT_STATE}_active_${chatId}`;
  
  // 1. Tenta encontrar o ponteiro para o estado ativo
  const activeStateKey = cache.get(pointerKey);
  if (!activeStateKey) {
    return null; // Nenhum assistente ativo para este chat
  }

  // 2. Usa a chave do ponteiro para obter o estado completo
  const cachedState = cache.get(activeStateKey);
  if (cachedState) {
    const state = JSON.parse(cachedState);
    logToSheet(`Estado ATIVO do ASSISTENTE para ${chatId} recuperado da chave: ${activeStateKey}.`, "DEBUG");
    return state;
  }
  
  return null;
}

/**
 * NOVO: Limpa o estado ativo do assistente e o ponteiro.
 * @param {string} chatId O ID do chat do Telegram.
 */
function clearActiveAssistantState(chatId) {
  const cache = CacheService.getScriptCache();
  const pointerKey = `${CACHE_KEY_ASSISTANT_STATE}_active_${chatId}`;
  
  // Pega a chave do estado completo antes de apagar o ponteiro
  const activeStateKey = cache.get(pointerKey);
  
  if (activeStateKey) {
    cache.remove(activeStateKey); // Remove o estado completo
  }
  cache.remove(pointerKey); // Remove o ponteiro

  logToSheet(`Estado ATIVO do ASSISTENTE e ponteiro para ${chatId} limpos.`, "INFO");
}

/**
 * Lida com o fluxo do tutorial, enviando mensagens e gerenciando o estado.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 * @param {number} step O passo atual do tutorial.
 */
function handleTutorialStep(chatId, usuario, step) {
  let message = "";
  let keyboardButtons = []; // Array para os botões inline
  let expectedAction = ""; // Ação esperada para o próximo input do usuário
  let messageResponse = null; // Para armazenar o objeto de resposta da mensagem enviada

  logToSheet(`[Tutorial] HandleTutorialStep chamado para passo: ${step}`, "INFO");

  switch (step) {
    case 1:
      message = `👋 Olá, ${escapeMarkdown(usuario)}! Bem-vindo ao tutorial do Gasto Certo!\n\n` +
                `Vou te mostrar como usar o bot para controlar suas finanças de forma simples e rápida.\n\n` +
                `Pronto para começar?`;
      keyboardButtons.push([{ text: "✅ Sim, vamos lá!", callback_data: "/tutorial_next" }]);
      keyboardButtons.push([{ text: "⏭️ Pular Tutorial", callback_data: "/tutorial_skip" }]);
      break;

    case 2:
      message = `*Passo 1: Registrando Despesas* 💸\n\n` +
                `Para registrar um gasto, basta digitar a descrição, o valor, o método de pagamento e a conta/cartão.\n\n` +
                `*Tente registrar sua primeira despesa agora!* \n\n` +
                `*Exemplo:* \`gastei 50 no mercado com Cartao Nubank Breno\`\n` +
                `*Exemplo:* \`paguei 30 de uber no debito do Santander\`\n\n` +
                `_Você pode usar os botões abaixo para navegar ou pular o tutorial._`;
      expectedAction = TUTORIAL_STATE_WAITING_DESPESA; // Define a ação esperada
      break;

    case 3:
      message = `*Passo 2: Registrando Receitas* 💰\n\n` +
                `Registrar uma receita é tão fácil quanto uma despesa! Basta indicar o valor e onde você recebeu.\n\n` +
                `*Tente registrar sua primeira receita agora!* \n\n` +
                `*Exemplo:* \`recebi 3000 de salario no Itau via PIX\`\n` +
                `*Exemplo:* \`ganhei 100 de comissao no PicPay\`\n\n` +
                `_Você pode usar os botões abaixo para navegar ou pular o tutorial._`;
      expectedAction = TUTORIAL_STATE_WAITING_RECEITA; // Define a ação esperada
      break;

    case 4:
      message = `*Passo 3: Visualizando seu Saldo e Extrato* 📊\n\n` +
                `Para ver a situação das suas contas, use os comandos:\n\n` +
                `* /saldo - Mostra o saldo de todas as suas contas e cartões.\n` +
                `* /extrato - Permite ver transações detalhadas por conta ou tipo.\n\n` +
                `*Tente digitar* \`/saldo\` *agora para ver como funciona!* \n\n` +
                `_Você pode usar os botões abaixo para navegar ou pular o tutorial._`;
      expectedAction = TUTORIAL_STATE_WAITING_SALDO; // Define a ação esperada
      break;

    case 5:
      message = `*Passo 4: Gerenciando Contas a Pagar e Metas* 🎯\n\n` +
                `O Gasto Certo também te ajuda a não esquecer suas contas e a acompanhar suas metas!\n\n` +
                `* /contasapagar - Lista suas contas fixas e faturas pendentes.\n` +
                `* /metas - Mostra seu progresso em relação às metas financeiras.\n\n` +
                `*Tente digitar* \`/contasapagar\` *ou* \`/metas\` *para explorar!* \n\n` +
                `_Você pode usar os botões abaixo para navegar ou pular o tutorial._`;
      expectedAction = TUTORIAL_STATE_WAITING_CONTAS_A_PAGAR; // Define a ação esperada para este passo (pode ser /contasapagar ou /metas)
      break;

    case 6:
      message = `🎉 *Parabéns! Você concluiu o tutorial do Gasto Certo!* 🎉\n\n` +
                `Agora você já sabe o básico para controlar suas finanças. Lembre-se:\n\n` +
                `- Use a linguagem natural para registrar transações.\n` +
                `- Explore os comandos:\n` +
                `  - \`/resumo\`\n` +
                `  - \`/saldo\`\n` +
                `  - \`/extrato\`\n` +
                `- E acesse o Dashboard Web para uma visão completa!\n\n` +
                `Se precisar de mais ajuda, digite \`/ajuda\` a qualquer momento.`;
      clearTutorialState(chatId); // Limpa o estado do tutorial ao finalizar
      break;

    default:
      logToSheet(`[Tutorial] Passo ${step} desconhecido ou fora do fluxo. Limpando estado do tutorial.`, "WARN");
      clearTutorialState(chatId);
      enviarMensagemTelegram(chatId, "🤔 Ops! O tutorial foi reiniciado. Digite /tutorial para começar novamente.");
      return; // Sai da função para evitar enviar mensagem duplicada
  }

  // Adiciona botões de navegação se não for o primeiro ou último passo
  if (step > 1 && step < 6) {
    keyboardButtons.push([
      { text: "⬅️ Passo Anterior", callback_data: "/tutorial_prev" },
      { text: "Próximo Passo ➡️", callback_data: "/tutorial_next" }
    ]);
  }
  
  // Adiciona o botão de pular tutorial para todos os passos interativos
  if (step > 1 && step < 6) {
    keyboardButtons.push([{ text: "⏭️ Pular Tutorial", callback_data: "/tutorial_skip" }]);
  }

  // Envia a mensagem com os botões
  const inlineKeyboard = keyboardButtons.length > 0 ? { inline_keyboard: keyboardButtons } : null;
  messageResponse = enviarMensagemTelegram(chatId, message, { reply_markup: inlineKeyboard });
  
  // Salva o estado do tutorial com a ação esperada e o ID da mensagem
  setTutorialState(chatId, step, expectedAction, messageResponse ? messageResponse.message_id : null);
  logToSheet(`Tutorial enviado para ${chatId}, passo: ${step}`, "INFO");
}

/**
 * Processa a entrada do usuário quando o tutorial está ativo.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 * @param {string} textoRecebido O texto da mensagem do usuário.
 * @param {Object} tutorialState O estado atual do tutorial (currentStep, expectedAction, messageId).
 * @returns {boolean} True se a entrada foi tratada pelo tutorial, false caso contrário (para que doPost continue o processamento normal).
 */
function processTutorialInput(chatId, usuario, textoRecebido, tutorialState) {
  // Normaliza o texto recebido para garantir comparações case-insensitive e sem acentos
  const normalizedTextoRecebido = normalizarTexto(textoRecebido); 
  // Extrai o comando base da entrada normalizada (ex: "saldo" de "/saldo com detalhes")
  // Remove a barra inicial se existir, pois normalizedTextoRecebido já pode ter removido.
  const inputComandoBase = normalizedTextoRecebido.startsWith('/') ? normalizedTextoRecebido.substring(1).split(/\s+/)[0] : normalizedTextoRecebido.split(/\s+/)[0];

  logToSheet(`[Tutorial] Processando input para tutorial. Passo: ${tutorialState.currentStep}, Acao Esperada: ${tutorialState.expectedAction}, Texto Original: "${textoRecebido}", Texto Normalizado: "${normalizedTextoRecebido}", Input Comando Base: "${inputComandoBase}"`, "INFO");

  let handledByTutorial = false;
  const dadosPalavras = getSheetDataWithCache(SHEET_PALAVRAS_CHAVE, CACHE_KEY_PALAVRAS);

  switch (tutorialState.expectedAction) {
    case TUTORIAL_STATE_WAITING_DESPESA:
      // Tenta interpretar a mensagem como despesa
      const tipoDetectadoDespesa = detectarTipoTransacao(normalizedTextoRecebido, dadosPalavras);
      if (tipoDetectadoDespesa && tipoDetectadoDespesa.tipo === "Despesa" && extrairValor(normalizedTextoRecebido) > 0) {
        // Se for uma despesa válida, NÃO avança o tutorial AQUI.
        // O avanço será feito em registrarTransacaoConfirmada após a confirmação.
        logToSheet(`[Tutorial] Despesa valida detectada. Aguardando confirmacao para avancar tutorial.`, "DEBUG");
        handledByTutorial = false; // Não marca como handled para que a transação seja processada normalmente
      } else {
        enviarMensagemTelegram(chatId, "🤔 Não consegui identificar uma despesa. Lembre-se, use algo como: `gastei 50 no mercado com Cartao Nubank`.");
        handledByTutorial = true; // Mantém no tutorial, apenas dá um hint
      }
      break;

    case TUTORIAL_STATE_WAITING_RECEITA:
      // Tenta interpretar a mensagem como receita
      const tipoDetectadoReceita = detectarTipoTransacao(normalizedTextoRecebido, dadosPalavras);
      if (tipoDetectadoReceita && tipoDetectadoReceita.tipo === "Receita" && extrairValor(normalizedTextoRecebido) > 0) {
        logToSheet(`[Tutorial] Receita valida detectada. Aguardando confirmacao para avancar tutorial.`, "DEBUG");
        handledByTutorial = false; // Não marca como handled para que a transação seja processada normalmente
      } else {
        enviarMensagemTelegram(chatId, "🤔 Não consegui identificar uma receita. Tente algo como: `recebi 3000 de salario no Itau via PIX`.");
        handledByTutorial = true;
      }
      break;

    case TUTORIAL_STATE_WAITING_SALDO:
      logToSheet(`[Tutorial DEBUG] Comparando Input Comando Base: "${inputComandoBase}" com "saldo" e "extrato"`, "DEBUG");
      // Verifica se o comando base é "saldo" ou "extrato" (sem a barra)
      if (inputComandoBase === "saldo" || inputComandoBase === "extrato") {
        logToSheet(`[Tutorial] Comando de saldo/extrato DETECTADO. Avancando tutorial.`, "DEBUG");
        enviarMensagemTelegram(chatId, "👍 Perfeito! Você usou um comando. O Dashboard Web é o próximo passo.");
        handleTutorialStep(chatId, usuario, tutorialState.currentStep + 1);
        handledByTutorial = true;
      } else {
        logToSheet(`[Tutorial] Comando de saldo/extrato NAO DETECTADO. Enviando hint.`, "DEBUG");
        enviarMensagemTelegram(chatId, "🤔 Tente usar um comando como `/saldo` ou `/extrato` para ver como funciona.");
        handledByTutorial = true; // Mantém o usuário neste passo, fornece dica.
      }
      break;

    case TUTORIAL_STATE_WAITING_CONTAS_A_PAGAR:
      logToSheet(`[Tutorial DEBUG] Comparando Input Comando Base: "${inputComandoBase}" com "contasapagar" e "metas"`, "DEBUG");
      // Verifica se o comando base é "contasapagar" ou "metas" (sem a barra)
      if (inputComandoBase === "contasapagar" || inputComandoBase === "metas") {
        logToSheet(`[Tutorial] Comando de contas a pagar/metas DETECTADO. Avancando tutorial.`, "DEBUG");
        enviarMensagemTelegram(chatId, "✨ Excelente! Você está explorando as ferramentas de planejamento. O tutorial está quase no fim!");
        handleTutorialStep(chatId, usuario, tutorialState.currentStep + 1);
        handledByTutorial = true;
      } else {
        logToSheet(`[Tutorial] Comando de contas a pagar/metas NAO DETECTADO. Enviando hint.`, "DEBUG");
        enviarMensagemTelegram(chatId, "🤔 Tente usar um comando como `/contasapagar` ou `/metas` para explorar.");
        handledByTutorial = true; // Mantém o usuário neste passo, fornece dica.
      }
      break;

    default:
      // Se não houver ação esperada específica para este passo, o tutorial não trata a entrada.
      logToSheet(`[Tutorial] Nenhuma acao esperada especifica para o passo ${tutorialState.currentStep}.`, "DEBUG");
      handledByTutorial = false; 
      break;
  }

  // Se a mensagem original do tutorial tinha um ID e a ação foi tratada pelo tutorial, edita para remover o teclado inline
  // Isso é feito APENAS se a mensagem foi tratada E o tutorial AVANÇOU (ou seja, handledByTutorial é true).
  // Se handledByTutorial for false, significa que a transação ainda está pendente de confirmação.
  // A edição da mensagem do tutorial para remover os botões deve ocorrer quando o tutorial AVANÇA.
  // Para transações (despesa/receita), o avanço ocorre em `registrarTransacaoConfirmada`.
  // Para comandos (/saldo, /contasapagar), o avanço ocorre aqui.
  if (handledByTutorial && tutorialState.messageId) {
      editMessageReplyMarkup(chatId, tutorialState.messageId, null); // Remove os botões
      logToSheet(`[Tutorial] Botões da mensagem de tutorial ${tutorialState.messageId} removidos.`, "DEBUG");
  }

  return handledByTutorial;
}


// --- Funções de Gerenciamento de Contas e Categorias ---

/**
 * Adiciona uma nova conta (ou cartão) à planilha 'Contas'.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 * @param {string} complemento O texto com o nome e tipo da conta (ex: "Nubank Credito", "Itau Corrente").
 */
function adicionarNovaConta(chatId, usuario, complemento) {
  logToSheet(`Adicionar nova conta: ${complemento} por ${usuario}`, "INFO");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_CONTAS);
  if (!sheet) {
    enviarMensagemTelegram(chatId, "❌ Erro: Aba 'Contas' não encontrada.");
    return;
  }

  const partes = complemento.split(/\s+/);
  if (partes.length < 2) {
    enviarMensagemTelegram(chatId, "❌ Formato inválido. Use: `/adicionar_conta <Nome da Conta> <Tipo (Crédito/Corrente/Dinheiro)>`");
    return;
  }

  const tipoContaRaw = partes[partes.length - 1];
  const nomeConta = partes.slice(0, partes.length - 1).join(" ");

  const tipoContaNormalizado = normalizarTexto(tipoContaRaw);
  let tipoValidado = "";

  if (tipoContaNormalizado.includes("credito") || tipoContaNormalizado.includes("crédito")) {
    tipoValidado = "Cartão de Crédito";
  } else if (tipoContaNormalizado.includes("corrente")) {
    tipoValidado = "Conta Corrente";
  } else if (tipoContaNormalizado.includes("dinheiro")) {
    tipoValidado = "Dinheiro Físico";
  } else {
    enviarMensagemTelegram(chatId, "❌ Tipo de conta inválido. Use 'Crédito', 'Corrente' ou 'Dinheiro'.");
    return;
  }

  const dadosContas = sheet.getDataRange().getValues();
  for (let i = 1; i < dadosContas.length; i++) {
    if (normalizarTexto(dadosContas[i][0]) === normalizarTexto(nomeConta)) {
      enviarMensagemTelegram(chatId, `⚠️ A conta *${escapeMarkdown(nomeConta)}* já existe.`);
      return;
    }
  }

  const newRow = [nomeConta, tipoValidado, '', 0, 0, (tipoValidado === "Cartão de Crédito" ? 1000 : ''), (tipoValidado === "Cartão de Crédito" ? 10 : ''), 'Ativo', '', (tipoValidado === "Cartão de Crédito" ? 1 : ''), 'Fechamento-no-mês', 5, '', usuario];
  sheet.appendRow(newRow);
  enviarMensagemTelegram(chatId, `✅ Conta *${escapeMarkdown(nomeConta)}* (${tipoValidado}) adicionada com sucesso!`);
  logToSheet(`Conta ${nomeConta} (${tipoValidado}) adicionada por ${usuario}.`, "INFO");

  // Limpa o cache de contas para que a nova conta seja carregada na próxima vez
  CacheService.getScriptCache().remove(CACHE_KEY_CONTAS);
  atualizarSaldosDasContas(); // Recalcula saldos para incluir a nova conta
}

/**
 * Lista todas as contas (e cartões) cadastradas na planilha 'Contas'.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 */
function listarContas(chatId, usuario) {
  logToSheet(`Listando contas para ${usuario}`, "INFO");
  const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);

  let mensagem = `🏦 *Suas Contas e Cartões:*\n\n`;
  let temContas = false;

  for (let i = 1; i < dadosContas.length; i++) {
    const nomeConta = dadosContas[i][0];
    const tipoConta = dadosContas[i][1];
    const saldoAtualizado = dadosContas[i][4]; // Coluna E (índice 4)
    const limite = dadosContas[i][5]; // Coluna F (índice 5)

    if (nomeConta) {
      mensagem += `*${escapeMarkdown(nomeConta)}* (${tipoConta}): `;
      if (tipoConta === "Cartão de Crédito") {
        mensagem += `Fatura: ${formatCurrency(saldoAtualizado)} / Limite: ${formatCurrency(limite)}\n`;
      } else {
        mensagem += `Saldo: ${formatCurrency(saldoAtualizado)}\n`;
      }
      temContas = true;
    }
  }

  if (!temContas) {
    mensagem = "Você ainda não tem contas cadastradas. Use `/adicionar_conta <Nome> <Tipo>` para começar.";
  }

  enviarMensagemTelegram(chatId, mensagem);
}

/**
 * Adiciona uma nova categoria e subcategoria à planilha 'PalavrasChave'.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 * @param {string} complemento O texto com a categoria, subcategoria e palavra-chave (ex: "Alimentacao Supermercado mercado").
 */
function adicionarNovaCategoria(chatId, usuario, complemento) {
  logToSheet(`Adicionar nova categoria: ${complemento} por ${usuario}`, "INFO");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PALAVRAS_CHAVE);
  if (!sheet) {
    enviarMensagemTelegram(chatId, "❌ Erro: Aba 'PalavrasChave' não encontrada.");
    return;
  }

  const partes = complemento.split(/\s+/);
  if (partes.length < 3) {
    enviarMensagemTelegram(chatId, "❌ Formato inválido. Use: `/adicionar_categoria <Categoria> <Subcategoria> <palavra-chave>`");
    return;
  }

  const categoria = capitalize(partes[0]);
  const subcategoria = capitalize(partes[1]);
  const palavraChave = partes.slice(2).join(" ").toLowerCase();
  const valorInterpretado = `${categoria} > ${subcategoria}`;

  const dadosPalavras = sheet.getDataRange().getValues();
  for (let i = 1; i < dadosPalavras.length; i++) {
    if (normalizarTexto(dadosPalavras[i][1]) === normalizarTexto(palavraChave) && normalizarTexto(dadosPalavras[i][0]) === "subcategoria") {
      enviarMensagemTelegram(chatId, `⚠️ A palavra-chave *${escapeMarkdown(palavraChave)}* já está associada a uma subcategoria.`);
      return;
    }
  }

  const newRow = ["subcategoria", palavraChave, valorInterpretado, ""];
  sheet.appendRow(newRow);
  enviarMensagemTelegram(chatId, `✅ Categoria *${escapeMarkdown(categoria)}* > *${escapeMarkdown(subcategoria)}* adicionada com a palavra-chave *${escapeMarkdown(palavraChave)}*!`);
  logToSheet(`Categoria ${categoria} > ${subcategoria} (palavra-chave: ${palavraChave}) adicionada por ${usuario}.`, "INFO");

  // Limpa o cache de palavras-chave
  CacheService.getScriptCache().remove(CACHE_KEY_PALAVRAS);
}

/**
 * Lista todas as categorias e subcategorias cadastradas na planilha 'PalavrasChave'.
 * @param {string} chatId O ID do chat do Telegram.
 */
function listarCategorias(chatId) {
  logToSheet("Listando categorias.", "INFO");
  const dadosPalavras = getSheetDataWithCache(SHEET_PALAVRAS_CHAVE, CACHE_KEY_PALAVRAS);

  const categorias = {};
  for (let i = 1; i < dadosPalavras.length; i++) {
    if (dadosPalavras[i][0] === "subcategoria") {
      const valorInterpretado = dadosPalavras[i][2]; // Ex: "Alimentação > Supermercado"
      const chave = dadosPalavras[i][1]; // Ex: "mercado"

      if (valorInterpretado && valorInterpretado.includes(">")) {
        const partes = valorInterpretado.split(">");
        const categoria = partes[0].trim();
        const subcategoria = partes[1].trim();

        if (!categorias[categoria]) {
          categorias[categoria] = new Set();
        }
        categorias[categoria].add(`${subcategoria} (palavra: ${chave})`);
      }
    }
  }

  let message = `📚 *Suas Categorias e Subcategorias:*\n\n`;
  let temCategorias = false;

  for (const cat in categorias) {
    message += `*${escapeMarkdown(cat)}*\n`;
    const subcategoriasOrdenadas = Array.from(categorias[cat]).sort();
    subcategoriasOrdenadas.forEach(sub => {
      message += `  • ${escapeMarkdown(sub)}\n`;
    });
    message += "\n";
    temCategorias = true;
  }

  if (!temCategorias) {
    message = "Você ainda não tem categorias cadastradas. Use `/adicionar_categoria <Categoria> <Subcategoria> <palavra-chave>` para começar.";
  }

  // CORREÇÃO: Usar enviarMensagemLongaTelegram para lidar com mensagens muito grandes
  enviarMensagemLongaTelegram(chatId, message); 
}

/**
 * Lista as subcategorias para uma categoria principal específica.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} categoriaPrincipal A categoria principal para filtrar as subcategorias.
 */
function listarSubcategorias(chatId, categoriaPrincipal) {
  logToSheet(`Listando subcategorias para: ${categoriaPrincipal}`, "INFO");
  const dadosPalavras = getSheetDataWithCache(SHEET_PALAVRAS_CHAVE, CACHE_KEY_PALAVRAS);
  const categoriaNormalizada = normalizarTexto(categoriaPrincipal);

  const subcategoriasEncontradas = [];
  for (let i = 1; i < dadosPalavras.length; i++) {
    if (dadosPalavras[i][0] === "subcategoria") {
      const valorInterpretado = dadosPalavras[i][2]; // Ex: "Alimentação > Supermercado"
      const chave = dadosPalavras[i][1]; // Ex: "mercado"

      if (valorInterpretado && valorInterpretado.includes(">")) {
        const partes = valorInterpretado.split(">");
        const categoria = normalizarTexto(partes[0].trim());
        const subcategoria = partes[1].trim();

        if (categoria === categoriaNormalizada) {
          subcategoriasEncontradas.push(`${subcategoria} (palavra: ${chave})`);
        }
      }
    }
  }

  let message = `📚 *Subcategorias para ${escapeMarkdown(capitalize(categoriaPrincipal))}:*\n\n`;
  if (subcategoriasEncontradas.length > 0) {
    subcategoriasEncontradas.sort().forEach(sub => {
      message += `  • ${escapeMarkdown(sub)}\n`;
    });
  } else {
    message += `Nenhuma subcategoria encontrada para *${escapeMarkdown(capitalize(categoriaPrincipal))}*.`;
  }

  enviarMensagemTelegram(chatId, message)

}  

/**
 * NOVO: Salva as configurações do Telegram e executa a configuração inicial do sistema.
 * Chamado pela interface de configuração (SetupDialog.html).
 * @param {Object} config Um objeto com as chaves 'token' e 'chatId'.
 * @returns {string} Uma mensagem de sucesso.
 */
function saveAndSetup(config) {
  try {
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties({
      [TELEGRAM_TOKEN_PROPERTY_KEY]: config.token,
      [ADMIN_CHAT_ID_PROPERTY_KEY]: config.chatId,
      [WEB_APP_URL_PROPERTY_KEY]: config.webAppUrl // Salva a URL do Web App
    });
    logToSheet(`Configurações de Token, Chat ID e URL do Web App salvas com sucesso.`, "INFO");

    const webhookResponse = setupWebhook(); // Não precisa mais passar o token, é lido das propriedades
    if (!webhookResponse.ok) {
      throw new Error(`Falha ao configurar o Webhook: ${webhookResponse.description}`);
    }
    logToSheet(`Webhook configurado com sucesso. Resposta: ${JSON.stringify(webhookResponse)}`, "INFO");

    initializeSheets();
    logToSheet(`Verificação e criação de abas concluída.`, "INFO");
    
    updateAdminConfig(config.chatId);
    logToSheet(`Configuração do administrador atualizada na aba 'Configuracoes'.`, "INFO");

    return "Configuração concluída com sucesso! O sistema está pronto para ser usado.";

  } catch (e) {
    logToSheet(`ERRO durante o setup inicial: ${e.message}`, "ERROR");
    throw e; // Lança o erro para ser capturado pelo withFailureHandler no cliente.
  }
}

/**
 * Função para configurar o webhook do Telegram.
 * Agora lê a URL do Web App diretamente das Propriedades do Script, que é mais confiável.
 * @returns {Object} Um objeto com o resultado da API do Telegram.
 */
function setupWebhook() {
  try {
    const token = getTelegramBotToken();
    // A URL é lida das propriedades, onde foi salva pela caixa de diálogo.
    const webhookUrl = PropertiesService.getScriptProperties().getProperty(WEB_APP_URL_PROPERTY_KEY);

    if (!webhookUrl) {
      const errorMessage = "URL do Web App não encontrada nas Propriedades do Script. Execute a 'Configuração Inicial' e forneça a URL correta.";
      throw new Error(errorMessage);
    }

    const url = `https://api.telegram.org/bot${token}/setWebhook?url=${webhookUrl}`;
    
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const responseText = response.getContentText();
    logToSheet(`Resposta da configuração do webhook: ${responseText}`, "INFO");
    return JSON.parse(responseText);

  } catch (e) {
    logToSheet(`Erro ao configurar o webhook: ${e.message}`, "ERROR");
    return { ok: false, description: e.message };
  }
}


/**
 * NOVO: Adiciona ou atualiza a configuração do usuário administrador na aba 'Configuracoes'.
 * @param {string} adminChatId O Chat ID do administrador.
  * @param {string} adminChatId O Chat ID do administrador.
 */
function updateAdminConfig(adminChatId) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(SHEET_CONFIGURACOES);
    const data = configSheet.getDataRange().getValues();
    let adminRowFound = false;

    // Procura por uma linha de admin existente para atualizar
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === 'chatId') {
            configSheet.getRange(i + 1, 2).setValue(adminChatId);
            configSheet.getRange(i + 1, 3).setValue('Admin'); // Define um nome padrão
            adminRowFound = true;
            break;
        }
    }
    
    // Se não encontrou, adiciona uma nova linha
    if (!adminRowFound) {
        configSheet.appendRow(['chatId', adminChatId, 'Admin', 'Default']);
    }
}


/**
 * Adiciona um novo usuário ao sistema.
 * @param {string} chatId O ID do chat do novo usuário.
 * @param {string} userName O nome do usuário.
 */
function addNewUser(chatId, userName) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEETS.USERS);
  if (sheet) {
    // Verifica se o usuário já existe
    const existingUser = findRowByValue(SHEETS.USERS, 1, chatId);
    if (!existingUser) {
      sheet.appendRow([chatId, userName, new Date()]);
      Logger.log(`Novo usuário adicionado: ${userName} (${chatId})`);
    }
  }
}

/**
 * Inicializa todas as abas necessárias da planilha com base no objeto HEADERS.
 * Garante que o ambiente do usuário seja criado corretamente.
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Itera sobre o objeto HEADERS para criar cada aba com seus respectivos cabeçalhos.
  for (const sheetName in HEADERS) {
    if (Object.prototype.hasOwnProperty.call(HEADERS, sheetName)) {
      if (!ss.getSheetByName(sheetName)) {
        const sheet = ss.insertSheet(sheetName);
        const headers = HEADERS[sheetName];
        if (headers && headers.length > 0) {
          sheet.appendRow(headers);
          logToSheet(`Aba '${sheetName}' criada com sucesso.`, "INFO");
        }
      }
    }
  }
  
  // Garante que a aba de logs também seja criada.
  if (!ss.getSheetByName(SHEET_LOGS_SISTEMA)) {
      const logSheet = ss.insertSheet(SHEET_LOGS_SISTEMA);
      logSheet.appendRow(["timestamp", "level", "message"]);
      logToSheet(`Aba de sistema '${SHEET_LOGS_SISTEMA}' criada com sucesso.`, "INFO");
  }
}
