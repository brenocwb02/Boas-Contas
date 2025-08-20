/**
 * @file SpeechToTextAPI.gs
 * @description Funções para interagir com a API de conversão de voz para texto (Speech-to-Text).
 */

/**
 * **ATUALIZADO**
 * Obtém a chave da API do Google Cloud Speech-to-Text a partir das Propriedades do Script.
 * @returns {string|null} A chave da API ou null se não for encontrada.
 */
function getSpeechApiKey() {
  // Agora lê a chave do local seguro (Propriedades do Script)
  const apiKey = PropertiesService.getScriptProperties().getProperty(SPEECH_API_KEY_PROPERTY_KEY);
  if (!apiKey) {
    logToSheet("AVISO: Chave da API de Speech-to-Text não configurada nas Propriedades do Script.", "WARN");
    return null;
  }
  return apiKey;
}

/**
 * Envia um arquivo de áudio para a API de Speech-to-Text e retorna a transcrição.
 * @param {Blob} audioBlob O arquivo de áudio no formato esperado pela API (ex: ogg).
 * @returns {string|null} O texto transcrito ou null em caso de erro.
 */
function transcreverAudio(audioBlob) {
  const API_KEY = getSpeechApiKey();

  if (!API_KEY) {
    // A função getSpeechApiKey já regista um log, não é necessário duplicar.
    return null;
  }

  const API_URL = 'https://speech.googleapis.com/v1/speech:recognize?key=' + API_KEY;

  const audioBytes = audioBlob.getBytes();
  const audioBase64 = Utilities.base64Encode(audioBytes);

  const requestBody = {
    config: {
      encoding: 'OGG_OPUS',
      sampleRateHertz: 48000,
      languageCode: 'pt-BR',
      enableAutomaticPunctuation: true,
      model: 'default'
    },
    audio: {
      content: audioBase64,
    },
  };

  try {
    const response = UrlFetchApp.fetch(API_URL, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true,
    });

    const responseText = response.getContentText();
    const result = JSON.parse(responseText);

    if (result.error) {
      const errorMessage = result.error.message || "Erro desconhecido da API.";
      const errorCode = result.error.code || "N/A";
      logToSheet(`Erro da API Speech-to-Text (Código: ${errorCode}): ${errorMessage}`, "ERROR");
      
      const adminChatId = getAdminChatIdFromProperties();
      if (adminChatId) {
        enviarMensagemTelegram(adminChatId, `⚠️ Alerta de Sistema: A API de transcrição de voz falhou com o erro: ${errorMessage}`);
      }
      return null;
    }

    if (result.results && result.results.length > 0 && result.results[0].alternatives.length > 0) {
      const transcript = result.results[0].alternatives[0].transcript;
      logToSheet(`Áudio transcrito com sucesso: "${transcript}"`, "INFO");
      return transcript;
    } else {
      logToSheet(`Falha na transcrição de áudio (sem resultados). Resposta da API: ${responseText}`, "WARN");
      return null;
    }

  } catch (e) {
    logToSheet(`ERRO FATAL ao contactar a API Speech-to-Text: ${e.message}`, "ERROR");
    return null;
  }
}
