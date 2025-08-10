/**
 * @file SpeechToTextAPI.gs
 * @description Funções para interagir com a API de conversão de voz para texto (Speech-to-Text).
 */

/**
 * **NOVO E CORRIGIDO**
 * Obtém a chave da API do Google Cloud Speech-to-Text a partir da aba 'Configuracoes'.
 * @returns {string|null} A chave da API ou null se não for encontrada.
 */
function getSpeechApiKey() {
  const configData = getSheetDataWithCache(SHEET_CONFIGURACOES, CACHE_KEY_CONFIG);
  for (let i = 0; i < configData.length; i++) {
    // Procura pela chave 'SPEECH_API_KEY' na primeira coluna
    if (configData[i][0] === 'SPEECH_API_KEY') {
      const apiKey = configData[i][1];
      // Verifica se a chave não é o valor padrão
      if (apiKey && apiKey !== "COLE_AQUI_SUA_CHAVE_API_DO_GOOGLE_CLOUD") {
        return apiKey;
      }
    }
  }
  return null;
}

/**
 * **FUNÇÃO ATUALIZADA E FUNCIONAL**
 * Envia um arquivo de áudio para a API de Speech-to-Text e retorna a transcrição.
 * @param {Blob} audioBlob O arquivo de áudio no formato esperado pela API (ex: ogg).
 * @returns {string|null} O texto transcrito ou null em caso de erro.
 */
function transcreverAudio(audioBlob) {
  const API_KEY = getSpeechApiKey();

  if (!API_KEY) {
    logToSheet("Chave da API Speech-to-Text não configurada na aba 'Configuracoes'.", "ERROR");
    return null;
  }

  const API_URL = 'https://speech.googleapis.com/v1/speech:recognize?key=' + API_KEY;

  // O Telegram envia áudio no formato OGG com codec Opus, que precisa de ser enviado como base64
  const audioBytes = audioBlob.getBytes();
  const audioBase64 = Utilities.base64Encode(audioBytes);

  const requestBody = {
    config: {
      encoding: 'OGG_OPUS',
      sampleRateHertz: 48000, // O Telegram usa esta taxa de amostragem
      languageCode: 'pt-BR',
      enableAutomaticPunctuation: true,
      model: 'default'
    },
    audio: {
      content: audioBase64,
    },
  };

  const response = UrlFetchApp.fetch(API_URL, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true,
  });

  const result = JSON.parse(response.getContentText());

  if (result.results && result.results.length > 0 && result.results[0].alternatives.length > 0) {
    const transcript = result.results[0].alternatives[0].transcript;
    logToSheet(`Áudio transcrito com sucesso: "${transcript}"`, "INFO");
    return transcript;
  } else {
    logToSheet(`Falha na transcrição de áudio. Resposta da API: ${JSON.stringify(result)}`, "WARN");
    return null;
  }
}
