/**
 * @file Utils.gs
 * @description Este arquivo contém funções utilitárias genéricas que podem ser usadas em diversas partes do código.
 * Inclui manipulação de strings, datas e números.
 */

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
 * NOVO: Escapa caracteres especiais do Markdown para evitar erros de parsing no Telegram.
 * Esta função é mais abrangente para MarkdownV2, mas segura para Markdown padrão.
 * @param {string} text O texto a ser escapado.
 * @returns {string} O texto com caracteres especiais escapados.
 */
function escapeMarkdown(text) {
  if (!text) return "";
  // Caracteres a escapar para MarkdownV2 (mais seguro para evitar problemas de parsing)
  // _ * [ ] ( ) ~ ` > # + - = | { } . !
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
             .replace(/\./g, '\\.') // Escapar ponto também, pois pode ser problemático em alguns contextos
             .replace(/!/g, '\\!');
}

/**
 * CORRIGIDO: Capitaliza a primeira letra de cada palavra em uma string, exceto para preposições e artigos comuns.
 * Garante que preposições sejam sempre minúsculas.
 * @param {string} texto O texto a ser capitalizado.
 * @returns {string} O texto com as primeiras letras capitalizadas onde apropriado.
 */
function capitalize(texto) {
  if (!texto) return "";
  const preposicoes = new Set(["de", "da", "do", "dos", "das", "e", "ou", "a", "o", "no", "na", "nos", "nas"]);
  return texto.split(' ').map((word, index) => {
    if (index > 0 && preposicoes.has(word.toLowerCase())) {
      return word.toLowerCase(); // Preposições e artigos em minúsculas
    }
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join(' ');
}

/**
 * NOVO: Função robusta para parsear strings de valores monetários no formato brasileiro (ex: "3.810,77").
 * Remove separadores de milhar e converte o separador decimal (vírgula) para ponto.
 * @param {string|number} valueString A string ou número a ser parseado.
 * @returns {number} O valor numérico parseado, ou 0 se não for um número válido.
 */
function parseBrazilianFloat(valueString) {
  if (typeof valueString === 'number') {
    return valueString; // Já é um número, retorna como está
  }
  if (typeof valueString !== 'string') {
    return 0; // Não é string nem número, retorna 0
  }

  let cleanValue = valueString.replace('R$', '').trim();

  const lastCommaIndex = cleanValue.lastIndexOf(',');
  const lastDotIndex = cleanValue.lastIndexOf('.');

  if (lastCommaIndex > lastDotIndex) { // Formato brasileiro: 1.234,56 ou 1.234
    cleanValue = cleanValue.replace(/\./g, ''); // Remove separadores de milhares (pontos)
    cleanValue = cleanValue.replace(',', '.');  // Substitui a vírgula decimal por ponto
  } else if (lastDotIndex > lastCommaIndex) { // Formato internacional: 1,234.56
    cleanValue = cleanValue.replace(/,/g, ''); // Remove separadores de milhares (vírgulas)
    // O ponto decimal já está correto
  }
  // Se não houver vírgula nem ponto, parseFloat lidará com isso (ex: "123")

  return parseFloat(cleanValue) || 0;
}

/**
 * NOVO: Formata um valor numérico como uma string de moeda brasileira (BRL).
 * @param {number} value O valor a ser formatado.
 * @returns {string} A string formatada, ex: "R$ 1.234,56".
 */
function formatCurrency(value) {
  if (typeof value !== 'number') {
    const numericValue = parseFloat(value);
    if (isNaN(numericValue)) {
      return "R$ 0,00";
    }
    value = numericValue;
  }
  return new Intl.NumberFormat('pt-BR', {
    style: 'currency',
    currency: 'BRL'
  }).format(value);
}

/**
 * CORRIGIDO E APRIMORADO: Converte um valor de data (string ou Date) para um objeto Date,
 * usando Utilities.parseDate para robustez com fusos horários.
 * Suporta formatos "DD/MM/YYYY" e "YYYY-MM-DD".
 * @param {any} valor O valor a ser convertido.
 * @returns {Date|null} Um objeto Date ou null se a conversão falhar.
 */
function parseData(valor) {
  if (valor instanceof Date) return valor;

  if (typeof valor !== "string") {
    logToSheet(`[parseData] Valor nao e string nem Date: "${valor}" (tipo: ${typeof valor}). Retornando null.`, "DEBUG");
    return null;
  }

  const timezone = Session.getScriptTimeZone(); // Obtém o fuso horário do script
  let parsedDate = null;

  // Tenta formato DD/MM/YYYY
  if (valor.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
    try {
      // Utilities.parseDate funciona melhor com MM/DD/YYYY ou YYYY-MM-DD
      // Vamos reformatar para MM/DD/YYYY antes de passar para Utilities.parseDate
      const parts = valor.split("/");
      const reformattedDate = `${parts[1]}/${parts[0]}/${parts[2]}`; // MM/DD/YYYY
      parsedDate = Utilities.parseDate(reformattedDate, timezone, "MM/dd/yyyy");
      logToSheet(`[parseData] Parsed DD/MM/YYYY "${valor}" para ${parsedDate.toLocaleDateString()} (via reformat).`, "DEBUG");
      return parsedDate;
    } catch (e) {
      logToSheet(`[parseData] Erro ao tentar parsear DD/MM/YYYY "${valor}" (reformat): ${e.message}`, "DEBUG");
    }
  }

  // Tenta formato YYYY-MM-DD
  if (valor.match(/^\d{4}-\d{2}-\d{2}$/)) {
    try {
      parsedDate = Utilities.parseDate(valor, timezone, "yyyy-MM-dd");
      logToSheet(`[parseData] Parsed YYYY-MM-DD "${valor}" para ${parsedDate.toLocaleDateString()}.`, "DEBUG");
      return parsedDate;
    } catch (e) {
      logToSheet(`[parseData] Erro ao tentar parsear YYYY-MM-DD "${valor}": ${e.message}`, "DEBUG");
    }
  }
  
  logToSheet(`[parseData] Falha ao parsear data "${valor}" em qualquer formato reconhecido. Retornando null.`, "WARN");
  return null;
}

/**
 * Obtém o nome do mês em português a partir do índice (0-11).
 * @param {number} mes O índice do mês (0 para Janeiro, 11 para Dezembro).
 * @returns {string} O nome do mês, ou uma string vazia se o índice for inválido.
 */
function getNomeMes(mes) {
  const meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];
  return meses[mes] || "";
}

/**
 * Arredonda um número para um número específico de casas decimais.
 * Evita problemas de precisão de ponto flutuante em JavaScript.
 * @param {number} value O número a ser arredondado.
 * @param {number} decimals O número de casas decimais.
 * @returns {number} O número arredondado.
 */
function round(value, decimals) {
  return Number(Math.round(value + 'e' + decimals) + 'e-' + decimals);
}

/**
 * Helper function to parse month and year from a string.
 * @param {string} inputString The string containing month and/or year (e.g., "junho 2024", "julho", "06 24").
 * @returns {Object} An object with `month` (1-12) and `year`. Defaults to current month/year if not found.
 */
function parseMonthAndYear(inputString) {
  const today = new Date();
  let month = today.getMonth() + 1; // 1-indexed
  let year = today.getFullYear();

  if (!inputString) {
    return { month: month, year: year };
  }

  const normalizedInput = normalizarTexto(inputString);
  const parts = normalizedInput.split(/\s+/);

  const monthNames = {
    "janeiro": 1, "jan": 1,
    "fevereiro": 2, "fev": 2,
    "marco": 3, "mar": 3,
    "abril": 4, "abr": 4,
    "maio": 5, "mai": 5,
    "junho": 6, "jun": 6,
    "julho": 7, "jul": 7,
    "agosto": 8, "ago": 8,
    "setembro": 9, "set": 9,
    "outubro": 10, "out": 10,
    "novembro": 11, "nov": 11,
    "dezembro": 12, "dez": 12
  };

  let potentialMonth = null;
  let potentialYear = null;

  for (const part of parts) {
    // Try to parse as month name
    if (monthNames[part]) {
      potentialMonth = monthNames[part];
    }
    // Try to parse as numeric month (e.g., "06", "6")
    else if (/^\d{1,2}$/.test(part) && parseInt(part, 10) >= 1 && parseInt(part, 10) <= 12) {
      potentialMonth = parseInt(part, 10);
    }
    // Try to parse as year (4-digit or 2-digit)
    else if (/^\d{4}$/.test(part)) {
      potentialYear = parseInt(part, 10);
    }
    else if (/^\d{2}$/.test(part)) {
      // Assume 2-digit years are in the 21st century (e.g., 24 -> 2024)
      potentialYear = 2000 + parseInt(part, 10);
    }
  }

  if (potentialMonth !== null) {
    month = potentialMonth;
  }
  if (potentialYear !== null) {
    year = potentialYear;
  }

  return { month: month, year: year };
}

/**
 * Calcula a distância de Levenshtein entre duas strings.
 * Usada para "fuzzy matching" (correspondência aproximada).
 * @param {string} s1 A primeira string.
 * @param {string} s2 A segunda string.
 * @returns {number} A distância de Levenshtein (número de edições para transformar s1 em s2).
 */
function levenshteinDistance(s1, s2) {
  s1 = s1.toLowerCase();
  s2 = s2.toLowerCase();

  const costs = new Array();
  for (let i = 0; i <= s1.length; i++) {
    let lastValue = i;
    for (let j = 0; j <= s2.length; j++) {
      if (i === 0)
        costs[j] = j;
      else {
        if (j > 0) {
          let newValue = costs[j - 1];
          if (s1.charAt(i - 1) !== s2.charAt(j - 1))
            newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
          costs[j - 1] = lastValue;
          lastValue = newValue;
        }
      }
    }
    if (i > 0)
      costs[s2.length] = lastValue;
  }
  return costs[s2.length];
}

/**
 * Calcula a similaridade entre duas strings com base na distância de Levenshtein.
 * Retorna um valor entre 0 e 1, onde 1 é idêntico e 0 é completamente diferente.
 * @param {string} s1 A primeira string.
 * @param {string} s2 A segunda string.
 * @returns {number} O coeficiente de similaridade.
 */
function calculateSimilarity(s1, s2) {
  const maxLength = Math.max(s1.length, s2.length);
  if (maxLength === 0) return 1.0; // Ambas vazias, consideradas idênticas.
  return (maxLength - levenshteinDistance(s1, s2)) / maxLength;
}

/**
 * NOVO: Obtém uma lista de todos os nomes de usuários configurados.
 * @param {Array<Array<any>>} configData Os dados da aba "Configuracoes".
 * @returns {Array<string>} Uma lista com os nomes dos usuários.
 */
function getAllUserNames(configData) {
  const userNames = new Set();
  for (let i = 1; i < configData.length; i++) {
    const nome = configData[i][2]; // Coluna 'nomeUsuario'
    if (nome) {
      userNames.add(nome.trim());
    }
  }
  return Array.from(userNames);
}

/**
 * NOVO: Procura por um nome de usuário conhecido dentro de uma string de texto.
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
