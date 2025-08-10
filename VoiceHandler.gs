/**
 * @file VoiceHandler.gs
 * @description Lida com o processamento de mensagens de voz recebidas do Telegram,
 * orquestrando o download do áudio, a transcrição e o processamento do texto resultante.
 */

/**
 * Processa uma mensagem de voz recebida do Telegram.
 * @param {Object} message O objeto de mensagem do Telegram contendo os dados da voz.
 */
function handleVoiceMessage(message) {
  try {
    const chatId = message.chat.id;
    const fileId = message.voice.file_id;

    if (!fileId) {
      enviarMensagemTelegram(chatId, "Não consegui processar o áudio. Tente novamente.");
      logToSheet("Mensagem de voz recebida sem file_id.", "WARN");
      return;
    }

    enviarMensagemTelegram(chatId, "🎙️ A processar a sua mensagem de voz, um momento...");

    // Passo 1: Baixar o arquivo de áudio do Telegram
    const audioBlob = getTelegramFile(fileId);
    if (!audioBlob) {
      enviarMensagemTelegram(chatId, "❌ Falha ao baixar o arquivo de áudio. Por favor, tente novamente.");
      return;
    }

    // **CORREÇÃO:** O nome da função foi corrigido de `transcribeAudio` para `transcreverAudio`.
    const textoTranscrito = transcreverAudio(audioBlob);
    if (!textoTranscrito) {
      enviarMensagemTelegram(chatId, "❌ Desculpe, não consegui entender o que você disse. Pode tentar digitar?");
      return;
    }

    // Passo 3: Processar o texto transcrito como se fosse uma mensagem de texto normal
    enviarMensagemTelegram(chatId, `Você disse: "_${escapeMarkdown(textoTranscrito)}_"`, { parse_mode: 'Markdown' });
    
    const configData = getSheetDataWithCache(SHEET_CONFIGURACOES, CACHE_KEY_CONFIG);
    const usuario = getUsuarioPorChatId(chatId, configData);

    const resultado = interpretarMensagemTelegram(textoTranscrito, usuario, chatId);

    if (resultado && resultado.errorMessage) {
      enviarMensagemTelegram(chatId, `❌ ${resultado.errorMessage}`);
    } else if (!resultado || (!resultado.status && !resultado.message)) {
      enviarMensagemTelegram(chatId, "Não entendi o seu lançamento a partir do áudio. Tente ser mais claro ou digite /ajuda.");
    }
  } catch (e) {
    const chatId = message?.chat?.id;
    logToSheet(`Erro ao processar mensagem de voz: ${e.message}`, "ERROR");
    if (chatId) {
      enviarMensagemTelegram(chatId, "❌ Ocorreu um erro inesperado ao processar sua mensagem de voz. O administrador foi notificado.");
    }
  }
}
