/**
 * @file Onboarding.gs
 * @description Contém a lógica para a configuração guiada via Telegram.
 */

const SETUP_STATE_KEY = 'guided_setup_state';

// Estados possíveis da conversa
const SETUP_STEPS = {
  PENDING_START: 'pending_start', // NOVO ESTADO
  AWAITING_ACCOUNT_NAME: 'awaiting_account_name',
  AWAITING_ACCOUNT_TYPE: 'awaiting_account_type',
  AWAITING_KEYWORD_FOR_CATEGORY: 'awaiting_keyword_for_category',
  FINISHED: 'finished'
};

/**
 * **FUNÇÃO ATUALIZADA COM A PERSONA ZAQ**
 * Agora apenas envia a mensagem de boas-vindas e atualiza o estado.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do utilizador.
 */
function startGuidedSetup(chatId, usuario) {
  const message = `👋 Olá, ${escapeMarkdown(usuario)}! Sou o Zaq, seu novo agente financeiro.\n\n` +
                  `A minha missão é ajudá-lo a transformar a sua relação com o dinheiro, começando com clareza e simplicidade.\n\n` +
                  `Vamos configurar a sua primeira conta? *Qual o nome da sua principal conta bancária?* (ex: Itaú, Nubank, Carteira)`;
  
  // Atualiza o estado para indicar que agora estamos à espera do nome da conta
  const state = {
    step: SETUP_STEPS.AWAITING_ACCOUNT_NAME,
    data: {}
  };
  setGuidedSetupState(chatId, state);
  
  enviarMensagemTelegram(chatId, message);
  logToSheet(`[Onboarding] Mensagem de início da configuração guiada (Zaq) enviada para ${usuario} (${chatId}).`, "INFO");
}

/**
 * **FUNÇÃO ATUALIZADA**
 * Processa a resposta do utilizador durante a configuração guiada.
 * Agora ignora a primeira mensagem se for /start.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do utilizador.
 * @param {string} text A resposta do utilizador.
 * @param {Object} state O estado atual da configuração.
 */
function handleGuidedSetupInput(chatId, usuario, text, state) {
  // Se for um callback de um botão, limpa o prefixo
  if (text.startsWith('setup_type_')) {
    text = text.replace('setup_type_', '');
  }

  // Se o estado for PENDING_START, ignora. O fluxo será iniciado pelo /start no doPost.
  if (state.step === SETUP_STEPS.PENDING_START) {
    logToSheet(`[Onboarding] Estado PENDING_START detetado. A aguardar /start do utilizador.`, "DEBUG");
    return;
  }

  switch (state.step) {
    case SETUP_STEPS.AWAITING_ACCOUNT_NAME:
      processAccountName(chatId, text, state);
      break;
    case SETUP_STEPS.AWAITING_ACCOUNT_TYPE:
      processAccountType(chatId, text, state);
      break;
    case SETUP_STEPS.AWAITING_KEYWORD_FOR_CATEGORY:
      processKeyword(chatId, text, state);
      break;
    default:
      clearGuidedSetupState(chatId);
      enviarMensagemTelegram(chatId, "Parece que a configuração já foi concluída. Se precisar de ajuda, digite /ajuda.");
      break;
  }
}

/** Processa o nome da conta e pergunta o tipo. */
function processAccountName(chatId, accountName, state) {
  state.data.accountName = accountName;
  state.step = SETUP_STEPS.AWAITING_ACCOUNT_TYPE;
  setGuidedSetupState(chatId, state);

  const message = `✅ Ótimo! E a conta *${escapeMarkdown(accountName)}* é de que tipo?`;
  const teclado = {
    inline_keyboard: [
      [{ text: "Conta Corrente", callback_data: "setup_type_Conta Corrente" }],
      [{ text: "Cartão de Crédito", callback_data: "setup_type_Cartão de Crédito" }],
      [{ text: "Dinheiro Físico", callback_data: "setup_type_Dinheiro Físico" }]
    ]
  };
  enviarMensagemTelegram(chatId, message, { reply_markup: teclado });
}

/** Processa o tipo da conta, regista-a e pergunta sobre a palavra-chave. */
function processAccountType(chatId, accountType, state) {
  state.data.accountType = accountType;
  
  // Adiciona a conta à planilha
  const result = addAccountToSheet(state.data.accountName, state.data.accountType);
  
  if (result.success) {
    enviarMensagemTelegram(chatId, `👍 Conta *${escapeMarkdown(state.data.accountName)}* adicionada com sucesso!`);
    
    state.step = SETUP_STEPS.AWAITING_KEYWORD_FOR_CATEGORY;
    setGuidedSetupState(chatId, state);

    const message = `Excelente! Agora, vamos tornar o bot mais inteligente. O sistema aprende com palavras-chave.\n\n` +
                    `*Diga-me uma palavra-chave que usa para gastos com 'Alimentação'* (ex: mercado, ifood, restaurante).`;
    enviarMensagemTelegram(chatId, message);

  } else {
    enviarMensagemTelegram(chatId, `❌ Ocorreu um erro: ${result.message}. Por favor, tente novamente.`);
    clearGuidedSetupState(chatId);
  }
}

/** Processa a palavra-chave, regista-a e finaliza a configuração. */
function processKeyword(chatId, keyword, state) {
  // Adiciona a palavra-chave à planilha
  const result = addKeywordToSheet('categoria', keyword, 'Alimentação', 'Outros');

  if (result.success) {
    state.step = SETUP_STEPS.FINISHED;
    clearGuidedSetupState(chatId); // Limpa o estado ao finalizar

    const message = `🎉 *Perfeito! Configuração inicial concluída!* 🎉\n\n` +
                    `Você já pode começar a registar as suas finanças. Tente enviar:\n` +
                    `\`gastei 50 com ${escapeMarkdown(keyword)} no ${escapeMarkdown(state.data.accountName)}\`\n\n` +
                    `Para ver todos os comandos, digite \`/ajuda\`.`;
    enviarMensagemTelegram(chatId, message);
    logToSheet(`[Onboarding] Configuração guiada concluída para ${chatId}.`, "INFO");
  } else {
    enviarMensagemTelegram(chatId, `❌ Ocorreu um erro: ${result.message}. Por favor, tente novamente.`);
    clearGuidedSetupState(chatId);
  }
}

// Funções de gestão de estado
function setGuidedSetupState(chatId, state) {
  const cache = CacheService.getScriptCache();
  cache.put(`${SETUP_STATE_KEY}_${chatId}`, JSON.stringify(state), 900); // 15 minutos de validade
}

function getGuidedSetupState(chatId) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(`${SETUP_STATE_KEY}_${chatId}`);
  return cached ? JSON.parse(cached) : null;
}

function clearGuidedSetupState(chatId) {
  const cache = CacheService.getScriptCache();
  cache.remove(`${SETUP_STATE_KEY}_${chatId}`);
}
