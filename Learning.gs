/**
 * @file Learning.gs
 * @description Contém a lógica de aprendizado do sistema a partir das interações do usuário.
 */

// A categoria aprendida só será usada automaticamente após ser corrigida 2 vezes.
const MIN_CONFIDENCE_TO_APPLY = 2;

/**
 * Extrai a palavra-chave mais relevante de uma descrição para fins de aprendizado.
 * @param {string} description A descrição da transação.
 * @returns {string|null} A palavra-chave extraída ou null se nenhuma for relevante.
 */
function extractKeywordForLearning(description) {
    // Palavras comuns a serem ignoradas
    const commonWords = new Set(['de', 'da', 'do', 'dos', 'das', 'e', 'ou', 'a', 'o', 'no', 'na', 'nos', 'nas', 'com', 'em', 'para', 'por', 'pelo', 'pela', 'via', 'compra', 'pagamento', 'lançamento']);
    const words = normalizarTexto(description).split(' ');
    
    // Encontra a primeira palavra que não seja comum e tenha mais de 2 letras
    const keyword = words.find(word => !commonWords.has(word) && word.length > 2);
    
    return keyword || null;
}

/**
 * Registra ou reforça uma associação entre uma palavra-chave e uma categoria/subcategoria.
 * Esta função é o "cérebro" do aprendizado, sendo chamada após uma correção do usuário.
 * @param {string} description A descrição original da transação da qual aprender.
 * @param {string} correctedCategory A categoria correta definida pelo usuário.
 * @param {string} correctedSubcategory A subcategoria correta definida pelo usuário.
 */
function learnFromCorrection(description, correctedCategory, correctedSubcategory) {
    const keyword = extractKeywordForLearning(description);
    
    // Validação para garantir que temos todos os dados necessários para aprender
    if (!keyword || !correctedCategory || !correctedSubcategory) {
        logToSheet(`[Learning] Não foi possível aprender. Faltam dados. Keyword: ${keyword}, Categoria: ${correctedCategory}, Subcategoria: ${correctedSubcategory}`, "WARN");
        return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let learnedSheet = ss.getSheetByName(SHEET_LEARNED_CATEGORIES);

    // Se a aba de aprendizado não existir, cria-a
    if (!learnedSheet) {
        learnedSheet = ss.insertSheet(SHEET_LEARNED_CATEGORIES);
        learnedSheet.appendRow(HEADERS[SHEET_LEARNED_CATEGORIES]);
        logToSheet(`[Learning] Aba '${SHEET_LEARNED_CATEGORIES}' criada.`, "INFO");
    }

    const data = learnedSheet.getDataRange().getValues();
    const headers = data[0];
    const colMap = getColumnMap(headers);

    let rowIndex = -1;
    // Procura se já existe uma regra de aprendizado para esta combinação
    for (let i = 1; i < data.length; i++) {
        if (data[i][colMap['Keyword']] === keyword && data[i][colMap['Categoria']] === correctedCategory && data[i][colMap['Subcategoria']] === correctedSubcategory) {
            rowIndex = i + 1; // Linha real na planilha
            break;
        }
    }

    if (rowIndex !== -1) {
        // Se a regra já existe, reforça sua confiança
        const currentScore = parseInt(data[rowIndex - 1][colMap['ConfidenceScore']]) || 0;
        learnedSheet.getRange(rowIndex, colMap['ConfidenceScore'] + 1).setValue(currentScore + 1);
        learnedSheet.getRange(rowIndex, colMap['LastUpdated'] + 1).setValue(new Date());
        logToSheet(`[Learning] Confiança para '${keyword}' -> '${correctedCategory} > ${correctedSubcategory}' aumentada para ${currentScore + 1}.`, "INFO");
    } else {
        // Se for uma nova regra, adiciona-a com confiança inicial 1
        learnedSheet.appendRow([keyword, correctedCategory, correctedSubcategory, 1, new Date()]);
        logToSheet(`[Learning] Nova associação aprendida: '${keyword}' -> '${correctedCategory} > ${correctedSubcategory}'.`, "INFO");
    }
}
