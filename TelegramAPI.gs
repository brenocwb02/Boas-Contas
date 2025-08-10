/**
 * @file TelegramAPI.gs
 * @description Este arquivo contém funções para interagir diretamente com a API do Telegram.
 * Inclui envio de mensagens e reconhecimento de callbacks.
 */

/**
 * Envia uma mensagem de texto para um chat específico no Telegram.
 * @param {string|number} chatId O ID do chat para enviar a mensagem.
 * @param {string} text O texto da mensagem. Suporta Markdown.
 * @param {Object} [options={}] Opções adicionais, como 'reply_markup' para teclados inline.
 * @returns {Object|null} O objeto de resultado da API do Telegram ou null em caso de erro.
 */
function enviarMensagemTelegram(chatId, text, options = {}) {
  try {
    const token = getTelegramBotToken(); // Chama a função centralizada
    const url = `${URL_BASE_TELEGRAM}${token}/sendMessage`;

    const payload = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        chat_id: String(chatId),
        text: text,
        parse_mode: "Markdown",
        ...options
      }),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, payload);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode === 200) {
      logToSheet(`Mensagem enviada para ${chatId}: (trecho) "${text.substring(0, 50)}..."`, "INFO");
      return JSON.parse(responseText).result;
    } else {
      logToSheet(`Falha ao enviar mensagem para ${chatId}. Código: ${responseCode}. Resposta: ${responseText}`, "ERROR");
      return null;
    }
  } catch (e) {
    logToSheet(`ERRO FATAL em enviarMensagemTelegram: ${e.message}`, "ERROR");
    return null;
  }
}

/**
 * NOVO E CENTRALIZADO: Obtém o token do bot do Telegram das Propriedades do Script.
 * Esta função agora é a única fonte para obter o token.
 * @returns {string} O token do bot.
 * @throws {Error} Se o token não estiver configurado.
 */
function getTelegramBotToken() {
  // Usa a constante definida em Constants.gs para consistência
  const token = PropertiesService.getScriptProperties().getProperty(TELEGRAM_TOKEN_PROPERTY_KEY);
  if (!token) {
    const errorMessage = "Token do Telegram não encontrado nas Propriedades do Script. Execute a 'Configuração Inicial' no menu da planilha.";
    logToSheet(`ERRO: ${errorMessage}`, "ERROR");
    throw new Error(errorMessage);
  }
  return token;
}


/**
 * NOVO: Função para reconhecer uma callback query do Telegram.
 * Isso impede que o Telegram reenvie a mesma query várias vezes.
 * @param {string} callbackQueryId O ID da callback query a ser respondida.
 * @param {string} [text] Texto opcional para um pop-up temporário no Telegram.
 * @param {boolean} [showAlert] Se deve mostrar um alerta ao usuário.
 */
function answerCallbackQuery(callbackQueryId, text = "", showAlert = false) {
  let token;
  try {
    // Obtém o token do Telegram das propriedades do script, que é mais seguro.
    token = getTelegramBotToken();
  } catch (e) {
    logToSheet(`ERRO CRITICO: Nao foi possivel obter o token do Telegram para responder callback. ${e.message}`, "ERROR");
    Logger.log(`ERRO CRITICO: Nao foi possivel obter o token do Telegram para responder callback. ${e.message}`);
    return;
  }

  const url = `${URL_BASE_TELEGRAM}${token}/answerCallbackQuery`;
  const payload = {
    callback_query_id: callbackQueryId,
    text: text, 
    show_alert: showAlert 
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    UrlFetchApp.fetch(url, options);
    logToSheet(`CallbackQuery ${callbackQueryId} respondida com sucesso.`, "DEBUG");
  } catch (e) {
    logToSheet(`Erro ao responder CallbackQuery ${callbackQueryId}: ${e.message}`, "ERROR");
  }
}
/**
 * Edita uma mensagem enviada anteriormente pelo bot.
 * Usado para remover botões de confirmação após a ação do usuário.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {number} messageId O ID da mensagem a ser editada.
 * @param {Object} [replyMarkup] O novo reply_markup (null para remover botões).
 */
function editMessageReplyMarkup(chatId, messageId, replyMarkup = null) {
  let token;
  try {
    token = getTelegramBotToken();
  } catch (e) {
    logToSheet(`ERRO CRITICO: Nao foi possivel obter o token do Telegram para editar mensagem. ${e.message}`, "ERROR");
    Logger.log(`ERRO CRITICO: Nao foi possivel obter o token do Telegram para editar mensagem. ${e.message}`);
    return;
  }

  const url = `${URL_BASE_TELEGRAM}${token}/editMessageReplyMarkup`;
  const payload = {
    chat_id: chatId,
    message_id: messageId,
    reply_markup: replyMarkup ? JSON.stringify(replyMarkup) : JSON.stringify({}) // Envia objeto vazio para remover
  };

  const fetchOptions = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, fetchOptions);
    const responseData = JSON.parse(response.getContentText());
    if (!responseData.ok) {
      logToSheet(`Erro ao editar reply_markup da mensagem ${messageId}: ${responseData.description}`, "ERROR");
    } else {
      logToSheet(`Reply_markup da mensagem ${messageId} editado com sucesso.`, "DEBUG");
    }
  } catch (e) {
    logToSheet(`Exceção ao editar reply_markup da mensagem ${messageId}: ${e.message}`, "ERROR");
  }
}

/**
 * Envia uma mensagem longa para o Telegram, dividindo-a em partes se exceder o limite.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} text O texto da mensagem a ser enviada.
 * @param {Object} [replyMarkup] O objeto reply_markup para botões inline.
 */
function enviarMensagemLongaTelegram(chatId, text, replyMarkup = null) {
  const MAX_MESSAGE_LENGTH = 4096; // Limite de caracteres do Telegram
  let currentPos = 0;

  while (currentPos < text.length) {
    let part = text.substring(currentPos, currentPos + MAX_MESSAGE_LENGTH);
    let lastNewline = part.lastIndexOf('\n');

    // Tenta cortar no último newline para evitar quebrar palavras ou formatação
    if (lastNewline !== -1 && currentPos + MAX_MESSAGE_LENGTH < text.length) {
      part = part.substring(0, lastNewline);
      currentPos += part.length + 1; // Pula o newline também
    } else {
      currentPos += part.length;
    }

    // Envia a parte da mensagem. Apenas a última parte terá o replyMarkup.
    enviarMensagemTelegram(chatId, part, (currentPos >= text.length) ? replyMarkup : null);
    Utilities.sleep(500); // Pequena pausa para evitar limites de taxa do Telegram
  }
}



/**
 * Baixa um arquivo do Telegram usando o file_id.
 * @param {string} fileId O ID do arquivo a ser baixado.
 * @returns {Blob|null} O conteúdo do arquivo como um Blob, ou null em caso de falha.
 */
function getTelegramFile(fileId) {
  const token = getTelegramBotToken();
  
  // Primeiro, obtemos o file_path
  const getFileUrl = `${URL_BASE_TELEGRAM}${token}/getFile?file_id=${fileId}`;
  const fileInfoResponse = UrlFetchApp.fetch(getFileUrl, { muteHttpExceptions: true });
  const fileInfo = JSON.parse(fileInfoResponse.getContentText());

  if (!fileInfo.ok) {
    logToSheet(`Erro ao obter informações do arquivo ${fileId}: ${fileInfo.description}`, "ERROR");
    return null;
  }

  const filePath = fileInfo.result.file_path;
  const fileUrl = `https://api.telegram.org/file/bot${token}/${filePath}`;
  
  // Agora, baixamos o arquivo
  const fileResponse = UrlFetchApp.fetch(fileUrl, { muteHttpExceptions: true });
  if (fileResponse.getResponseCode() === 200) {
    return fileResponse.getBlob();
  }
  
  logToSheet(`Falha ao baixar o arquivo de ${fileUrl}. Código: ${fileResponse.getResponseCode()}`, "ERROR");
  return null;
}
