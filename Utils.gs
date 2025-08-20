/**
 * @file Utils.gs
 * @description Contém funções utilitárias genéricas que podem ser usadas em diversas partes do código.
 * Inclui manipulação de strings, datas, números e interação com a planilha.
 */

// ===================================================================================
// NOVA SEÇÃO: GESTÃO DE ERROS CENTRALIZADA
// ===================================================================================

/**
 * Função central para gerir e registar erros em todo o sistema.
 * @param {Error} error O objeto de erro capturado no bloco catch.
 * @param {string} context Uma string que descreve onde o erro ocorreu (ex: "doPost", "interpretarMensagem").
 * @param {string|number|null} chatId O ID do chat do utilizador, se disponível, para enviar uma mensagem de feedback.
 */
function handleError(error, context, chatId = null) {
  const errorMessage = `ERRO em ${context}: ${error.message}. Stack: ${error.stack}`;
  
  // Passo 1: Registar sempre o erro detalhado para o administrador.
  logToSheet(errorMessage, "ERROR");

  // Passo 2: Se um chatId for fornecido, enviar uma mensagem amigável ao utilizador.
  if (chatId) {
    const userMessage = "❌ Ocorreu um erro inesperado. A equipa de suporte já foi notificada. Por favor, tente novamente mais tarde.";
    enviarMensagemTelegram(chatId, userMessage);
  }

  // Passo 3 (Opcional): Notificar o administrador sobre erros críticos.
  // Pode ativar esta notificação para erros que não sejam de API (que já notificam).
  const adminChatId = getAdminChatIdFromProperties();
  if (adminChatId && !error.message.includes("API")) { // Exemplo: não notificar para erros de API já tratados
     // enviarMensagemTelegram(adminChatId, `⚠️ Alerta de Erro Crítico no Sistema:\n\nContexto: ${context}\nMensagem: ${error.message}`);
  }
}

// ===================================================================================
// FUNÇÕES DE MANIPULAÇÃO DE STRINGS
// ===================================================================================

/**
 * Normaliza texto: remove acentos, converte para minúsculas, remove caracteres especiais (exceto números e espaços)
 * e substitui múltiplos espaços por um único. Útil para comparações de strings.
 * @param {string} txt O texto a ser normalizado.
 * @returns {string} O texto normalizado.
 */
function normalizarTexto(txt) {
  if (!txt) return "";
  return txt
    .toString()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9\s]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Escapa caracteres especiais do Markdown para evitar erros de parsing no Telegram.
 * @param {string} text O texto a ser escapado.
 * @returns {string} O texto com caracteres especiais escapados.
 */
function escapeMarkdown(text) {
  if (!text) return "";
  return text.replace(/_/g, '\\_')
             .replace(/\*/g, '\\*')
             .replace(/\[/g, '\\[')
             .replace(/\]/g, '\\]')
             .replace(/\(/g, '\\(')
             .replace(/\)/g, '\\)')
             .replace(/~/g, '\\~')
             .replace(/`/g, '\\`')
             .replace(/>/g, '\\>')
             .replace(/#/g, '\\#')
             .replace(/\+/g, '\\+')
             .replace(/-/g, '\\-')
             .replace(/=/g, '\\=')
             .replace(/\|/g, '\\|')
             .replace(/{/g, '\\{')
             .replace(/}/g, '\\}')
             .replace(/\./g, '\\.')
             .replace(/!/g, '\\!');
}

/**
 * Capitaliza a primeira letra de cada palavra em uma string, exceto para preposições e artigos comuns.
 * @param {string} texto O texto a ser capitalizado.
 * @returns {string} O texto com as primeiras letras capitalizadas onde apropriado.
 */
function capitalize(texto) {
  if (!texto) return "";
  const preposicoes = new Set(["de", "da", "do", "dos", "das", "e", "ou", "a", "o", "no", "na", "nos", "nas"]);
  return texto.split(' ').map((word, index) => {
    if (index > 0 && preposicoes.has(word.toLowerCase())) {
      return word.toLowerCase();
    }
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join(' ');
}

// ===================================================================================
// FUNÇÕES DE MANIPULAÇÃO DE NÚMEROS E MOEDA
// ===================================================================================

/**
 * Função robusta para parsear strings de valores monetários no formato brasileiro (ex: "3.810,77") ou internacional.
 * @param {string|number} valueString A string ou número a ser parseado.
 * @returns {number} O valor numérico parseado, ou 0 se não for um número válido.
 */
function parseBrazilianFloat(valueString) {
  if (typeof valueString === 'number') return valueString;
  if (typeof valueString !== 'string') return 0;

  let cleanValue = valueString.replace('R$', '').trim();
  const lastCommaIndex = cleanValue.lastIndexOf(',');
  const lastDotIndex = cleanValue.lastIndexOf('.');

  if (lastCommaIndex > lastDotIndex) {
    cleanValue = cleanValue.replace(/\./g, '').replace(',', '.');
  } else {
    cleanValue = cleanValue.replace(/,/g, '');
  }
  return parseFloat(cleanValue) || 0;
}

/**
 * Formata um valor numérico como uma string de moeda brasileira (BRL).
 * @param {number} value O valor a ser formatado.
 * @returns {string} A string formatada, ex: "R$ 1.234,56".
 */
function formatCurrency(value) {
  if (typeof value !== 'number') {
    const numericValue = parseFloat(value);
    if (isNaN(numericValue)) return "R$ 0,00";
    value = numericValue;
  }
  return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value);
}

/**
 * Arredonda um número para um número específico de casas decimais.
 * @param {number} value O número a ser arredondado.
 * @param {number} decimals O número de casas decimais.
 * @returns {number} O número arredondado.
 */
function round(value, decimals) {
  return Number(Math.round(value + 'e' + decimals) + 'e-' + decimals);
}

// ===================================================================================
// FUNÇÕES DE MANIPULAÇÃO DE DATAS
// ===================================================================================

/**
 * Converte um valor de data (string ou Date) para um objeto Date.
 * Suporta formatos "DD/MM/YYYY" e "YYYY-MM-DD".
 * @param {any} valor O valor a ser convertido.
 * @returns {Date|null} Um objeto Date ou null se a conversão falhar.
 */
function parseData(valor) {
  if (!valor || (typeof valor === 'string' && valor.trim() === '')) return null;
  if (valor instanceof Date) return valor;
  if (typeof valor !== "string") return null;

  // Tenta formato DD/MM/YYYY
  if (valor.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
    try {
      const parts = valor.split("/");
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10);
      const year = parseInt(parts[2], 10);
      if (month < 1 || month > 12 || day < 1 || day > 31) throw new Error("Data inválida.");
      const dateObject = new Date(year, month - 1, day);
      if (dateObject.getFullYear() !== year || dateObject.getMonth() !== month - 1 || dateObject.getDate() !== day) {
        throw new Error("Data inválida, valores foram ajustados.");
      }
      return dateObject;
    } catch (e) { /* Ignora e tenta o próximo formato */ }
  }

  // Tenta formato YYYY-MM-DD
  if (valor.match(/^\d{4}-\d{2}-\d{2}$/)) {
    try {
      const parts = valor.split('-');
      return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    } catch (e) { /* Ignora */ }
  }
  
  return null;
}

/**
 * Obtém o nome do mês em português a partir do índice (0-11).
 * @param {number} mes O índice do mês (0 para Janeiro, 11 para Dezembro).
 * @returns {string} O nome do mês, ou uma string vazia se o índice for inválido.
 */
function getNomeMes(mes) {
  const meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
  return meses[mes] || "";
}

/**
 * Analisa uma string para extrair mês e ano.
 * @param {string} inputString A string contendo mês e/ou ano (ex: "junho 2024", "julho", "06 24").
 * @returns {Object} Um objeto com `month` (1-12) and `year`.
 */
function parseMonthAndYear(inputString) {
  const today = new Date();
  let month = today.getMonth() + 1;
  let year = today.getFullYear();

  if (!inputString) return { month, year };

  const normalizedInput = normalizarTexto(inputString);
  const parts = normalizedInput.split(/\s+/);
  const monthNames = {"janeiro": 1, "jan": 1, "fevereiro": 2, "fev": 2, "marco": 3, "mar": 3, "abril": 4, "abr": 4, "maio": 5, "mai": 5, "junho": 6, "jun": 6, "julho": 7, "jul": 7, "agosto": 8, "ago": 8, "setembro": 9, "set": 9, "outubro": 10, "out": 10, "novembro": 11, "nov": 11, "dezembro": 12, "dez": 12};

  parts.forEach(part => {
    if (monthNames[part]) month = monthNames[part];
    else if (/^\d{1,2}$/.test(part) && parseInt(part, 10) >= 1 && parseInt(part, 10) <= 12) month = parseInt(part, 10);
    else if (/^\d{4}$/.test(part)) year = parseInt(part, 10);
    else if (/^\d{2}$/.test(part)) year = 2000 + parseInt(part, 10);
  });

  return { month, year };
}

// ===================================================================================
// FUNÇÕES DE COMPARAÇÃO DE STRINGS
// ===================================================================================

/**
 * Calcula a distância de Levenshtein entre duas strings.
 * @param {string} s1 A primeira string.
 * @param {string} s2 A segunda string.
 * @returns {number} A distância de Levenshtein.
 */
function levenshteinDistance(s1, s2) {
  s1 = s1.toLowerCase();
  s2 = s2.toLowerCase();
  const costs = [];
  for (let i = 0; i <= s1.length; i++) {
    let lastValue = i;
    for (let j = 0; j <= s2.length; j++) {
      if (i === 0) costs[j] = j;
      else if (j > 0) {
        let newValue = costs[j - 1];
        if (s1.charAt(i - 1) !== s2.charAt(j - 1))
          newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
        costs[j - 1] = lastValue;
        lastValue = newValue;
      }
    }
    if (i > 0) costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}

/**
 * Calcula a similaridade entre duas strings com base na distância de Levenshtein.
 * @param {string} s1 A primeira string.
 * @param {string} s2 A segunda string.
 * @returns {number} O coeficiente de similaridade (0 a 1).
 */
function calculateSimilarity(s1, s2) {
  const maxLength = Math.max(s1.length, s2.length);
  if (maxLength === 0) return 1.0;
  return (maxLength - levenshteinDistance(s1, s2)) / maxLength;
}

// ===================================================================================
// FUNÇÕES DE UTILIDADE DA PLANILHA
// ===================================================================================

/**
 * Cria um mapa de nomes de cabeçalho para seus índices de coluna.
 * @param {Array<string>} headers A linha de cabeçalho.
 * @returns {Object} Um objeto mapeando nomes de cabeçalho para seus índices base 0.
 */
function getColumnMap(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[header.trim()] = index;
  });
  return map;
}

// ===================================================================================
// FUNÇÕES DE UTILIDADE DO PROJETO
// ===================================================================================

/**
 * Obtém uma lista de todos os nomes de usuários configurados.
 * @param {Array<Array<any>>} configData Os dados da aba "Configuracoes".
 * @returns {Array<string>} Uma lista com os nomes dos usuários.
 */
function getAllUserNames(configData) {
  const userNames = new Set();
  for (let i = 1; i < configData.length; i++) {
    const nome = configData[i][2];
    if (nome) userNames.add(nome.trim());
  }
  return Array.from(userNames);
}

/**
 * Procura por um nome de usuário conhecido dentro de uma string de texto.
 * @param {string} text O texto onde procurar.
 * @param {Array<string>} userNames A lista de nomes de usuários conhecidos.
 * @returns {string|null} O nome do usuário encontrado ou null.
 */
function findUserNameInText(text, userNames) {
  if (!text) return null;
  const normalizedText = normalizarTexto(text);
  for (const userName of userNames) {
    if (normalizedText.includes(normalizarTexto(userName))) {
      return userName;
    }
  }
  return null;
}

/**
 * Inclui o conteúdo de um arquivo HTML dentro de outro.
 * @param {string} filename O nome do arquivo HTML a ser incluído (sem a extensão .html).
 * @return {string} O conteúdo do arquivo.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
