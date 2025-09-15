/**
 * @file Code_Dashboard.gs
 * @description Funções do lado do servidor para o Dashboard Financeiro,
 * incluindo a coleta de dados, e operações de CRUD para transações via Web App.
 * @version com getDashboardData refatorada para maior clareza e manutenibilidade.
 */

// ===================================================================================
// SEÇÃO DE CONCILIAÇÃO BANCÁRIA (CORRIGIDA)
// ===================================================================================

/**
 * Recebe o conteúdo de um extrato do frontend, chama o processador e retorna o resultado.
 * @param {string} fileContent O conteúdo do arquivo como uma string de texto.
 * @param {string} fileName O nome do arquivo.
 * @returns {object} O resultado da conciliação.
 */
function processBankStatement(fileContent, fileName) {
  // CORREÇÃO: Chama a função _processBankStatement do arquivo Importador.gs
  return _processBankStatement(fileContent, fileName);
}

/**
 * Recebe as transações selecionadas do frontend e as importa para a planilha.
 * @param {Array<object>} transactionsToImport As transações a serem adicionadas.
 * @returns {object} Um objeto de sucesso com os dados do dashboard atualizados.
 */
function importNewTransactions(transactionsToImport) {
  // CORREÇÃO: Chama a função _importNewTransactions do arquivo Importador.gs
  return _importNewTransactions(transactionsToImport);
}

/**
 * Serve o arquivo HTML do dashboard como um Web App com verificação de segurança.
 * @param {Object} e O objeto de evento do Apps Script.
 * @returns {HtmlOutput} O conteúdo HTML do dashboard ou uma página de erro.
 */
function doGet(e) {
  try {
    if (!isLicenseValid()) {
      logToSheet("[Dashboard Access] Acesso bloqueado: Licença do produto inválida.", "ERROR");
      return HtmlService.createHtmlOutputFromFile('AcessoNegadoLicensa.html').setTitle("Licença Inválida");
    }

    const SCRIPT_VERSION = "1.3"; 
    const token = e.parameter.token;
    const cache = CacheService.getScriptCache();
    const cacheKey = `${CACHE_KEY_DASHBOARD_TOKEN}_${token}`;
    
    if (!token) {
      logToSheet("[Dashboard Access] Acesso negado: nenhum token fornecido.", "WARN");
      return HtmlService.createHtmlOutput('<h1><i class="fas fa-lock"></i> Acesso Negado</h1><p>Este link não é válido. Para acessar o dashboard, por favor, solicite um novo link de acesso através do bot no Telegram com o comando <code>/dashboard</code>.</p><style>body{font-family: sans-serif; text-align: center; padding-top: 50px; color: #333;} i{color: #d9534f;}</style>').setTitle("Acesso Negado");
    }

    const expectedChatId = cache.get(cacheKey);
    if (!expectedChatId) {
      logToSheet(`[Dashboard Access] Acesso negado: token inválido ou expirado ('${token}').`, "WARN");
      return HtmlService.createHtmlOutput('<h1><i class="fas fa-clock"></i> Link Inválido ou Expirado</h1><p>Este link de acesso não é mais válido. Ele pode ter expirado ou já ter sido utilizado. Por favor, solicite um novo com o comando <code>/dashboard</code> no Telegram.</p><style>body{font-family: sans-serif; text-align: center; padding-top: 50px; color: #f0ad4e;}</style>').setTitle("Link Expirado");
    }

    logToSheet(`[Dashboard Access] Acesso concedido para o chatId ${expectedChatId} com o token '${token}'.`, "INFO");

    const template = HtmlService.createTemplateFromFile('Dashboard');
    template.chatId = expectedChatId; 
    template.version = SCRIPT_VERSION;
    return template.evaluate()
        .setTitle('Dashboard Financeiro')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

  } catch (error) {
    logToSheet(`[Dashboard Access] Erro crítico na função doGet: ${error.message}`, "ERROR");
    return HtmlService.createHtmlOutput('<h1><i class="fas fa-server"></i> Erro Interno</h1><p>Ocorreu um erro inesperado ao tentar carregar o dashboard. O administrador foi notificado.</p><style>body{font-family: sans-serif; text-align: center; padding-top: 50px; color: #c9302c;}</style>').setTitle("Erro Interno");
  }
}

/**
 * Calcula os gastos de cartão de crédito para um mês e ano específicos.
 * @param {Array<Array<any>>} dadosTransacoes Dados da aba 'Transacoes'.
 * @param {Array<Array<any>>} dadosContas Dados da aba 'Contas'.
 * @param {number} mes O mês para o cálculo (1-12).
 * @param {number} ano O ano para o cálculo.
 * @returns {Array<Object>} Um array com os gastos mensais por cartão.
 */
function getCreditCardSpendingForMonth(dadosTransacoes, dadosContas, mes, ano) {
    const spendingMap = {};
    const targetMonth = mes - 1;

    dadosContas.slice(1).forEach(row => {
        const nomeOriginal = (row[0] || "").trim();
        const tipo = (row[1] || "").toLowerCase().trim();
        if (tipo === 'cartão de crédito' && nomeOriginal) {
            const limite = parseBrazilianFloat(String(row[5] || '0'));
            spendingMap[nomeOriginal] = { gasto: 0, limite: limite };
        }
    });

    dadosTransacoes.slice(1).forEach(row => {
        const dataVencimento = parseData(row[10]);
        const tipoTransacao = (row[4] || "").toLowerCase().trim();
        const conta = (row[7] || "").trim();
        const categoria = normalizarTexto(row[2]);
        const subcategoria = normalizarTexto(row[3]);

        if (
            dataVencimento &&
            dataVencimento.getMonth() === targetMonth &&
            dataVencimento.getFullYear() === ano &&
            tipoTransacao === 'despesa' &&
            spendingMap[conta] &&
            !(categoria === "contas a pagar" && subcategoria === "pagamento de fatura")
        ) {
            const valor = parseBrazilianFloat(String(row[5]));
            spendingMap[conta].gasto += valor;
        }
    });

    return Object.entries(spendingMap).map(([nome, data]) => ({
        nomeOriginal: nome,
        faturaAtual: round(data.gasto, 2),
        limite: round(data.limite, 2)
    }));
}

/**
 * Calcula o saldo das contas correntes e dinheiro até o final de um determinado mês/ano.
 * @param {Array<Array<any>>} dadosTransacoes Dados da aba 'Transacoes'.
 * @param {Array<Array<any>>} dadosContas Dados da aba 'Contas'.
 * @param {number} mes O mês para o cálculo (1-12).
 * @param {number} ano O ano para o cálculo.
 * @returns {Array<Object>} Um array com os saldos das contas no final do período.
 */
function getAccountBalancesForMonth(dadosTransacoes, dadosContas, mes, ano) {
    const balancesMap = {};
    const endDate = new Date(ano, mes, 0, 23, 59, 59);

    dadosContas.slice(1).forEach(row => {
        const nomeOriginal = (row[0] || "").trim();
        const tipo = (row[1] || "").toLowerCase().trim();
        const banco = (row[2] || "").trim();
        const saldoInicial = parseBrazilianFloat(String(row[3] || '0'));

        if ((tipo === 'conta corrente' || tipo === 'dinheiro físico') && nomeOriginal) {
            balancesMap[nomeOriginal] = {
                saldo: saldoInicial,
                tipo: tipo,
                banco: banco,
                saldoInicial: saldoInicial
            };
        }
    });

    dadosTransacoes.slice(1).forEach(row => {
        const dataTransacao = parseData(row[0]);
        const tipoTransacao = (row[4] || "").toLowerCase().trim();
        const conta = (row[7] || "").trim();
        const valor = parseBrazilianFloat(String(row[5]));

        if (dataTransacao && dataTransacao <= endDate && balancesMap.hasOwnProperty(conta)) {
            if (tipoTransacao === 'receita') {
                balancesMap[conta].saldo += valor;
            } else if (tipoTransacao === 'despesa') {
                balancesMap[conta].saldo -= valor;
            }
        }
    });

    return Object.entries(balancesMap).map(([nome, data]) => ({
        nomeOriginal: nome,
        saldo: round(data.saldo, 2),
        tipo: data.tipo,
        banco: data.banco,
        saldoInicial: data.saldoInicial
    }));
}


// ===================================================================================
// SEÇÃO DE COLETA DE DADOS PARA O DASHBOARD (REFATORADA)
// ===================================================================================

/**
 * @private
 * Função auxiliar para extrair o ícone de uma string de categoria.
 */
function _extractIconAndCleanCategory(categoryString) {
    const str = String(categoryString || "");
    if (!str) return { cleanCategory: "", icon: "" };
    const match = str.match(/^(\p{Emoji}|\p{Emoji_Modifier_Base}|\p{Emoji_Component}|\p{Emoji_Modifier}|\p{Emoji_Presentation})\s*(.*)/u);
    if (match) return { cleanCategory: match[2].trim(), icon: match[1] };
    return { cleanCategory: str.trim(), icon: "" };
}

/**
 * @private
 * Calcula o resumo mensal (receitas, despesas, saldo).
 */
function _getDashboardSummary(dadosTransacoes, currentMonth, currentYear) {
  let totalReceitasMes = 0;
  let totalDespesasMes = 0;

  for (let i = 1; i < dadosTransacoes.length; i++) {
    const row = dadosTransacoes[i];
    const data = parseData(row[0]);
    if (!data || data.getMonth() !== currentMonth || data.getFullYear() !== currentYear) continue;

    const tipo = row[4];
    const valor = parseBrazilianFloat(String(row[5])) || 0;
    const categoria = normalizarTexto(row[2]);
    const subcategoria = normalizarTexto(row[3]);

    const isIgnored = (categoria === "contas a pagar" && subcategoria === "pagamento de fatura") ||
                      (categoria === "transferencias" && subcategoria === "entre contas") ||
                      (categoria === "pagamentos recebidos" && subcategoria === "pagamento de fatura");

    if (!isIgnored) {
      if (tipo === "Receita") totalReceitasMes += valor;
      else if (tipo === "Despesa") totalDespesasMes += valor;
    }
  }

  return {
    totalReceitas: round(totalReceitasMes, 2),
    totalDespesas: round(totalDespesasMes, 2),
    saldoLiquidoMes: round(totalReceitasMes - totalDespesasMes, 2)
  };
}

/**
 * @private
 * Calcula a análise de Necessidades vs. Desejos.
 */
function _getNeedsWantsSummary(dadosTransacoes, categoriasMap, currentMonth, currentYear) {
  let gastoNecessidades = 0;
  let gastoDesejos = 0;
  let despesasNaoClassificadas = 0;

  for (let i = 1; i < dadosTransacoes.length; i++) {
    const row = dadosTransacoes[i];
    const data = parseData(row[0]);
    if (!data || data.getMonth() !== currentMonth || data.getFullYear() !== currentYear) continue;

    if (row[4] === "Despesa") {
      const valor = parseBrazilianFloat(String(row[5])) || 0;
      const categoria = normalizarTexto(row[2]);
      const subcategoria = normalizarTexto(row[3]);
      
      const isIgnored = (categoria === "contas a pagar" && subcategoria === "pagamento de fatura") ||
                        (categoria === "transferencias" && subcategoria === "entre contas");

      if (!isIgnored) {
        const categoriaInfo = categoriasMap[categoria];
        if (categoriaInfo && categoriaInfo.tipoGasto === 'necessidade') gastoNecessidades += valor;
        else if (categoriaInfo && categoriaInfo.tipoGasto === 'desejo') gastoDesejos += valor;
        else despesasNaoClassificadas += valor;
      }
    }
  }

  return {
    necessidades: round(gastoNecessidades, 2),
    desejos: round(gastoDesejos, 2),
    naoClassificado: round(despesasNaoClassificadas, 2)
  };
}

/**
 * @private
 * Obtém as contas a pagar do mês.
 */
function _getBillsToPay(dadosContasAPagar, currentMonth, currentYear) {
  const billsToPay = [];
  if (dadosContasAPagar.length > 1) {
    const headers = dadosContasAPagar[0];
    const colMap = getColumnMap(headers);

    if (colMap['Descricao'] !== undefined && colMap['Valor'] !== undefined && colMap['Data de Vencimento'] !== undefined) {
      for (let i = 1; i < dadosContasAPagar.length; i++) {
        const row = dadosContasAPagar[i];
        const dataVencimento = parseData(row[colMap['Data de Vencimento']]);
        if (dataVencimento && dataVencimento.getMonth() === currentMonth && dataVencimento.getFullYear() === currentYear && normalizarTexto(row[colMap['Recorrente']]) === "verdadeiro") {
          billsToPay.push({
            descricao: (row[colMap['Descricao']] || "").toString().trim(),
            valor: round(parseBrazilianFloat(String(row[colMap['Valor']])), 2),
            dataVencimento: Utilities.formatDate(dataVencimento, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            status: (row[colMap['Status']] || "").toString().trim()
          });
        }
      }
      billsToPay.sort((a, b) => parseData(a.dataVencimento).getTime() - parseData(b.dataVencimento).getTime());
    }
  }
  return billsToPay;
}

/**
 * @private
 * Obtém os últimos 10 lançamentos do mês.
 */
function _getRecentTransactions(dadosTransacoes, currentMonth, currentYear) {
    const recentTransactions = [];
    if (dadosTransacoes.length > 1) {
        for (let i = dadosTransacoes.length - 1; i > 0 && recentTransactions.length < 10; i--) {
            const linha = dadosTransacoes[i];
            const dataObj = parseData(linha[0]);
            if (dataObj && dataObj.getMonth() === currentMonth && dataObj.getFullYear() === currentYear) {
                recentTransactions.push({
                    id: linha[13],
                    data: Utilities.formatDate(dataObj, Session.getScriptTimeZone(), "dd/MM/yyyy"),
                    descricao: linha[1],
                    categoria: linha[2],
                    subcategoria: linha[3],
                    tipo: linha[4],
                    valor: round(parseBrazilianFloat(String(linha[5])), 2),
                    metodoPagamento: linha[6],
                    conta: linha[7],
                    usuario: linha[11]
                });
            }
        }
    }
    return recentTransactions;
}

/**
 * @private
 * Calcula o progresso das metas financeiras (OBJETIVOS DE POUPANÇA) a partir da nova aba "Metas".
 * Esta função substitui a lógica anterior que lia a previsão de despesas.
 * VERSÃO CORRIGIDA E MAIS ROBUSTA.
 */
function _getGoalsProgress(dadosMetas, categoryIconsMap) {
    const goalsProgress = [];
    // Verifica se a aba Metas existe e tem pelo menos uma meta (além do cabeçalho)
    if (!dadosMetas || dadosMetas.length < 2) {
        console.log("AVISO: A aba 'Metas' está vazia ou não foi encontrada. Nenhuma meta será exibida.");
        return goalsProgress; // Retorna um array vazio, não causa erro.
    }

    // Pega o cabeçalho para encontrar as colunas pelo nome
    const headers = dadosMetas[0];
    const colMap = getColumnMap(headers);

    // Verifica se as colunas essenciais existem na nova aba Metas
    if (colMap['Nome da Meta'] === undefined || colMap['Valor Objetivo'] === undefined || colMap['Valor Salvo'] === undefined) {
        console.log("AVISO: A aba 'Metas' não tem a estrutura esperada (Nome da Meta, Valor Objetivo, Valor Salvo).");
        return goalsProgress; // Retorna um array vazio para não quebrar o dashboard.
    }

    // Itera por cada linha de meta (começando da linha 2, pulando o cabeçalho)
    for (let i = 1; i < dadosMetas.length; i++) {
        const row = dadosMetas[i];

        const nomeMeta = row[colMap['Nome da Meta']];
        // Converte para número, tratando tanto o formato "R$ 1.234,56" quanto "1234.56"
        const valorObjetivo = parseBrazilianFloat(String(row[colMap['Valor Objetivo']] || '0'));
        const valorSalvo = parseBrazilianFloat(String(row[colMap['Valor Salvo']] || '0'));

        // Adiciona a meta à lista apenas se ela tiver um nome e um objetivo definido
        if (nomeMeta && valorObjetivo > 0) {
            const percentage = valorObjetivo > 0 ? round((valorSalvo / valorObjetivo) * 100, 2) : 0;

            goalsProgress.push({
                // A estrutura de dados enviada para o frontend agora é mais simples
                categoria: nomeMeta, // Usamos o campo 'categoria' para o nome da meta
                subcategoria: '', // Não temos mais subcategoria aqui
                meta: valorObjetivo,
                gasto: valorSalvo, // O campo 'gasto' agora representa 'salvo'
                percentage: percentage,
                icon: '🎯' // Um ícone padrão para metas
            });
        }
    }

    return goalsProgress;
}

/**
 * @private
 * Calcula o progresso do orçamento para o mês.
 */
function _getBudgetProgress(dadosOrcamento, dadosTransacoes, currentMonth, currentYear, categoryIconsMap) {
    const budgetProgress = [];
    if (dadosOrcamento.length <= 1) return budgetProgress;

    const headers = dadosOrcamento[0];
    const colMap = getColumnMap(headers);
    if ([colMap['Categoria'], colMap['Valor Orcado'], colMap['Mes referencia']].some(c => c === undefined)) return budgetProgress;

    let orcamentoMap = {};
    for (let i = 1; i < dadosOrcamento.length; i++) {
        const row = dadosOrcamento[i];
        const dataReferencia = parseData(row[colMap['Mes referencia']]);
        if (dataReferencia && dataReferencia.getMonth() === currentMonth && dataReferencia.getFullYear() === currentYear) {
            const { cleanCategory: categoria, icon } = _extractIconAndCleanCategory(row[colMap['Categoria']]);
            const valorOrcado = parseBrazilianFloat(String(row[colMap['Valor Orcado']]));
            if (categoria && valorOrcado > 0) {
                const key = normalizarTexto(categoria);
                orcamentoMap[key] = {
                    categoria: categoria, orcado: valorOrcado, gasto: 0,
                    icon: icon || categoryIconsMap[key] || ''
                };
            }
        }
    }

    for (let i = 1; i < dadosTransacoes.length; i++) {
        const row = dadosTransacoes[i];
        const data = parseData(row[10]); // Data de Vencimento
        if (data && data.getMonth() === currentMonth && data.getFullYear() === currentYear && row[4] === "Despesa") {
            const { cleanCategory: categoria } = _extractIconAndCleanCategory(row[2]);
            const key = normalizarTexto(categoria);
            if (orcamentoMap[key]) {
                 const subcategoria = normalizarTexto(row[3]);
                 if (!(key === "contas a pagar" && subcategoria === "pagamento de fatura") && !(key === "transferencias" && subcategoria === "entre contas")) {
                    orcamentoMap[key].gasto += parseBrazilianFloat(String(row[5]));
                 }
            }
        }
    }

    for (const key in orcamentoMap) {
        const item = orcamentoMap[key];
        const percentage = item.orcado > 0 ? round((item.gasto / item.orcado) * 100, 2) : 0;
        budgetProgress.push({ ...item, gasto: round(item.gasto, 2), percentage });
    }
    return budgetProgress;
}

/**
 * @private
 * Prepara os dados para o gráfico de despesas por categoria.
 */
function _getExpensesByCategoryChartData(dadosTransacoes, currentMonth, currentYear, categoryIconsMap) {
    const tempExpensesMap = {};
    for (let i = 1; i < dadosTransacoes.length; i++) {
        const row = dadosTransacoes[i];
        const data = parseData(row[10]); // Data de Vencimento
        if (data && data.getMonth() === currentMonth && data.getFullYear() === currentYear && row[4] === "Despesa") {
            const { cleanCategory, icon } = _extractIconAndCleanCategory(row[2]);
            const categoriaNormalizada = normalizarTexto(cleanCategory);
            const subcategoriaNormalizada = normalizarTexto(row[3]);

            if (!(categoriaNormalizada === "contas a pagar" && subcategoriaNormalizada === "pagamento de fatura") && !(categoriaNormalizada === "transferencias" && subcategoriaNormalizada === "entre contas")) {
                if (!tempExpensesMap[categoriaNormalizada]) {
                    tempExpensesMap[categoriaNormalizada] = {
                        categoriaOriginal: cleanCategory, total: 0,
                        icon: icon || categoryIconsMap[categoriaNormalizada] || ''
                    };
                }
                tempExpensesMap[categoriaNormalizada].total += parseBrazilianFloat(String(row[5]));
            }
        }
    }

    return Object.values(tempExpensesMap).map(item => ({
        category: item.categoriaOriginal,
        value: round(item.total, 2),
        icon: item.icon
    }));
}

/**
 * @private
 * NOVO: Busca os dados históricos do valor da carteira de investimentos.
 * @returns {Array<Object>} Um array com objetos { data, valor }.
 */
function _getPortfolioHistoryData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName("PortfolioHistory"); // Usa a constante
  if (!historySheet || historySheet.getLastRow() < 2) {
    return [];
  }
  const data = historySheet.getRange("A2:B" + historySheet.getLastRow()).getValues();
  return data.map(row => ({
    date: new Date(row[0]).toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' }),
    value: parseFloat(row[1]) || 0
  }));
}

/**
 * **VERSÃO REDESENHADA E ORGANIZADA**
 * Coleta todos os dados necessários para o dashboard orquestrando chamadas a funções auxiliares.
 */
function getDashboardData(mes, ano) {
  logToSheet(`Iniciando coleta de dados para o Dashboard para Mes: ${mes}, Ano: ${ano}`, "INFO");
  
  // 1. Carregar todas as fontes de dados de uma vez
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
  const contasSheet = ss.getSheetByName(SHEET_CONTAS);
  const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);
  const metasSheet = ss.getSheetByName(SHEET_METAS);
  const orcamentoSheet = ss.getSheetByName(SHEET_ORCAMENTO);
  const categoriesSheet = ss.getSheetByName(SHEET_CATEGORIAS);

  if (!transacoesSheet || !contasSheet || !contasAPagarSheet || !metasSheet || !orcamentoSheet || !categoriesSheet) {
    throw new Error("Uma ou mais abas essenciais para o dashboard não foram encontradas.");
  }

  const dadosTransacoes = transacoesSheet.getDataRange().getValues();
  const dadosContas = contasSheet.getDataRange().getValues();
  const dadosContasAPagar = contasAPagarSheet.getDataRange().getValues();
  const dadosMetas = metasSheet.getDataRange().getValues();
  const dadosOrcamento = orcamentoSheet.getDataRange().getValues();
  const dadosCategorias = categoriesSheet.getDataRange().getValues();
  
  const currentMonth = mes - 1;
  const currentYear = ano;
  const nomeMesAtual = getNomeMes(currentMonth);
  const categoryIconsMap = { "vida espiritual": "🕊️", "moradia": "🏠", "despesas fixas / contas": "🧾", "alimentacao": "🛒", "familia / filhos": "👨‍👩‍👧‍👦", "educacao e desenvolvimento": "🎓", "transporte": "🚗", "saude": "💊", "despesas pessoais": "👔", "impostos e taxas": "📊", "lazer e entretenimento": "🎉", "relacionamentos": "❤️", "reserva / prevencao": "🛡️", "investimentos / futuro": "📈", "receitas de trabalho": "💼", "apoio / ajuda externa": "🤝", "outros ganhos": "🎁", "renda extra e investimentos": "💸", "artigos residenciais": "🛋️", "pag. de terceiros": "👥", "conta bancaria": "🏦", "transferencias": "🔄" };

  // 2. Chamar funções auxiliares para processar cada parte dos dados
  const dashboardSummary = _getDashboardSummary(dadosTransacoes, currentMonth, currentYear);
  const accountBalances = getAccountBalancesForMonth(dadosTransacoes, dadosContas, mes, ano);
  const creditCardSummaries = getCreditCardSpendingForMonth(dadosTransacoes, dadosContas, mes, ano);
  const billsToPay = _getBillsToPay(dadosContasAPagar, currentMonth, currentYear);
  const recentTransactions = _getRecentTransactions(dadosTransacoes, currentMonth, currentYear);
  const goalsProgress = _getGoalsProgress(dadosMetas, categoryIconsMap);
  const budgetProgress = _getBudgetProgress(dadosOrcamento, dadosTransacoes, currentMonth, currentYear, categoryIconsMap);
  const expensesByCategoryArray = _getExpensesByCategoryChartData(dadosTransacoes, currentMonth, currentYear, categoryIconsMap);
  const netWorthData = calculateNetWorth(); // <-- ADICIONADO
  const investmentData = _getInvestmentsData(); // <-- ADICIONADO
  const portfolioHistory = _getPortfolioHistoryData(); // <-- ADICIONADO
  
  const categoriasMap = getCategoriesMap();
  const needsWantsSummary = _getNeedsWantsSummary(dadosTransacoes, categoriasMap, currentMonth, currentYear);

  // 3. Montar e retornar o objeto final
  logToSheet("Coleta de dados para o Dashboard concluida.", "INFO");
  return {
    summary: dashboardSummary,
    accountBalances: accountBalances,
    creditCardSummaries: creditCardSummaries,
    billsToPay: billsToPay,
    recentTransactions: recentTransactions,
    goalsProgress: goalsProgress,
    budgetProgress: budgetProgress,
    expensesByCategory: expensesByCategoryArray,
    needsWantsSummary: needsWantsSummary,
    netWorth: netWorthData,
    investments: investmentData, // <-- ADICIONADO
    portfolioHistory: portfolioHistory, // <-- ADICIONADO
    accounts: getAccountsForDropdown(dadosContas),
    categories: getCategoriesForDropdown(dadosCategorias),
    paymentMethods: getPaymentMethodsForDropdown()
  };
}


/**
 * Busca e retorna todas as transações de uma categoria específica para um determinado mês e ano.
 * @param {string} categoryName O nome da categoria a ser filtrada.
 * @param {number} month O mês (1-12).
 * @param {number} year O ano.
 * @returns {Array<Object>} Uma lista de objetos de transação.
 */
function getTransactionsByCategory(categoryName, month, year) {
  try {
    logToSheet(`[Dashboard] Buscando transações para categoria '${categoryName}', Mês: ${month}, Ano: ${year}`, "INFO");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
    const dadosTransacoes = transacoesSheet.getDataRange().getValues();
    const targetMonth = month - 1;

    const transactions = [];

    for (let i = 1; i < dadosTransacoes.length; i++) {
      const row = dadosTransacoes[i];
      const data = parseData(row[10]);
      const { cleanCategory: transCategory } = _extractIconAndCleanCategory(row[2]);

      if (data && data.getMonth() === targetMonth && data.getFullYear() === year && normalizarTexto(transCategory) === normalizarTexto(categoryName)) {
        transactions.push({
          data: Utilities.formatDate(parseData(row[0]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
          descricao: row[1],
          categoria: row[2],
          subcategoria: row[3],
          tipo: row[4],
          valor: parseBrazilianFloat(String(row[5])),
          conta: row[7]
        });
      }
    }
    logToSheet(`[Dashboard] Encontradas ${transactions.length} transações para a categoria '${categoryName}'.`, "DEBUG");
    return transactions.sort((a, b) => new Date(a.data.split('/').reverse().join('-')) - new Date(b.data.split('/').reverse().join('-')));
  } catch (e) {
    handleError(e, `getTransactionsByCategory para ${categoryName}`, null);
    throw new Error(`Erro ao buscar transações: ${e.message}`);
  }
}

/**
 * **ATUALIZADO COM LOCKSERVICE E RETORNO DINÂMICO**
 * Adiciona uma nova transação à planilha "Transacoes" a partir do formulário web.
 * @param {Object} transactionData Objeto contendo os dados da transação.
 * @returns {Object} Um objeto indicando sucesso ou falha e os dados atualizados do dashboard.
 */
function addTransactionFromWeb(transactionData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Transacoes');
    if (!sheet) throw new Error("Planilha 'Transacoes' não encontrada.");

    const newTransactionId = Utilities.getUuid();
    
    let formattedDueDate = '';
    if (transactionData.dueDate) {
      const dateObject = new Date(transactionData.dueDate + 'T00:00:00');
      formattedDueDate = Utilities.formatDate(dateObject, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }

    let formattedTransactionDate = '';
    if (transactionData.date) {
        const dateObject = new Date(transactionData.date + 'T00:00:00');
        formattedTransactionDate = Utilities.formatDate(dateObject, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }

    const newRow = [
      formattedTransactionDate,
      transactionData.description,
      transactionData.category || '',
      transactionData.subcategory || '',
      transactionData.type,
      transactionData.value,
      transactionData.paymentMethod || '',
      transactionData.account,
      transactionData.installments,
      1,
      formattedDueDate,
      '', // Usuário (será preenchido se necessário no futuro)
      'Ativo',
      newTransactionId,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss")
    ];

    sheet.appendRow(newRow);
    atualizarSaldosDasContas(); 
    
    const today = new Date(transactionData.date);
    const dashboardData = getDashboardData(today.getMonth() + 1, today.getFullYear());
    
    return { 
      success: true, 
      message: 'Transação adicionada com sucesso.',
      newTransactionId: newTransactionId,
      dashboardData: dashboardData 
    };
  } catch (e) {
    handleError(e, "addTransactionFromWeb");
    return { success: false, message: 'Erro ao adicionar transação: ' + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * **ATUALIZADO COM LOCKSERVICE E RETORNO DINÂMICO**
 * Deleta uma transação da planilha 'Transacoes' a partir do Dashboard.
 * @param {string} transactionId O ID único da transação a ser deletada.
 * @returns {object} Um objeto com status de sucesso ou erro e os dados atualizados do dashboard.
 */
function deleteTransactionFromWeb(transactionId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
    if (!transacoesSheet) throw new Error(`Planilha "${SHEET_TRANSACOES}" não encontrada.`);

    const data = transacoesSheet.getDataRange().getValues();
    const headers = data[0];
    const idColumnIndex = headers.indexOf('ID Transacao');
    const descColumnIndex = headers.indexOf('Descricao');
    const valueColumnIndex = headers.indexOf('Valor');
    const dateColumnIndex = headers.indexOf('Data');
    
    if (idColumnIndex === -1 || descColumnIndex === -1 || valueColumnIndex === -1 || dateColumnIndex === -1) {
      throw new Error("Colunas essenciais não encontradas.");
    }

    const rowIndexToDelete = data.slice(1).findIndex(row => row[idColumnIndex] == transactionId);

    if (rowIndexToDelete !== -1) {
      const rowData = data[rowIndexToDelete + 1];
      const transactionDate = parseData(rowData[dateColumnIndex]);
      const description = rowData[descColumnIndex];
      const value = parseBrazilianFloat(String(rowData[valueColumnIndex]));

      if (description.startsWith("Aporte Meta:")) {
        const metaName = description.substring("Aporte Meta:".length).trim();
        reverterAporteMeta(metaName, value);
      }
      
      reverterStatusContaAPagarSeVinculado(transactionId);

      transacoesSheet.deleteRow(rowIndexToDelete + 2);
      logToSheet(`Transação com ID ${transactionId} deletada via Dashboard.`, "INFO");
      
      atualizarSaldosDasContas();
      
      const dashboardData = getDashboardData(transactionDate.getMonth() + 1, transactionDate.getFullYear());
      return { success: true, message: `Transação excluída com sucesso.`, dashboardData: dashboardData };
    } else {
      return { success: false, message: `Transação com ID ${transactionId} não encontrada.` };
    }
  } catch (e) {
    handleError(e, "deleteTransactionFromWeb");
    return { success: false, message: `Erro ao excluir transação: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

/**
 * **ATUALIZADO COM LOCKSERVICE E RETORNO DINÂMICO**
 * Atualiza uma transação existente.
 * @param {Object} transactionData Objeto com os dados da transação.
 * @returns {Object} Objeto indicando sucesso ou falha e os dados atualizados do dashboard.
 */
function updateTransactionFromWeb(transactionData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TRANSACOES);
    if (!sheet) throw new Error(`Planilha '${SHEET_TRANSACOES}' não encontrada.`);

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const colMap = getColumnMap(headers);

    const idColumn = colMap["ID Transacao"];
    const rowIndexToUpdate = data.slice(1).findIndex(row => row[idColumn] === transactionData.id);

    if (rowIndexToUpdate !== -1) {
      const rowIndex = rowIndexToUpdate + 2;
      sheet.getRange(rowIndex, colMap["Data"] + 1).setValue(new Date(transactionData.date + 'T00:00:00'));
      sheet.getRange(rowIndex, colMap["Descricao"] + 1).setValue(transactionData.description);
      sheet.getRange(rowIndex, colMap["Categoria"] + 1).setValue(transactionData.category);
      sheet.getRange(rowIndex, colMap["Subcategoria"] + 1).setValue(transactionData.subcategory);
      sheet.getRange(rowIndex, colMap["Tipo"] + 1).setValue(transactionData.type);
      sheet.getRange(rowIndex, colMap["Valor"] + 1).setValue(parseBrazilianFloat(String(transactionData.value)));
      sheet.getRange(rowIndex, colMap["Metodo de Pagamento"] + 1).setValue(transactionData.paymentMethod);
      sheet.getRange(rowIndex, colMap["Conta/Cartão"] + 1).setValue(transactionData.account);
      sheet.getRange(rowIndex, colMap["Parcelas Totais"] + 1).setValue(parseInt(transactionData.installments));
      sheet.getRange(rowIndex, colMap["Data de Vencimento"] + 1).setValue(new Date((transactionData.dueDate || transactionData.date) + 'T00:00:00'));
      
      atualizarSaldosDasContas();
      
      const transactionDate = new Date(transactionData.date);
      const dashboardData = getDashboardData(transactionDate.getMonth() + 1, transactionDate.getFullYear());
      
      return { success: true, message: 'Transação atualizada com sucesso.', dashboardData: dashboardData };
    }
    throw new Error("Transação não encontrada para atualização.");
  } catch (e) {
    handleError(e, "updateTransactionFromWeb");
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}


/**
 * Retorna uma lista de contas e cartões para popular um dropdown.
 */
function getAccountsForDropdown(accountsData) {
  return accountsData.slice(1).map(row => ({
    nomeOriginal: row[0],
    tipo: row[1]
  }));
}

/**
 * Retorna uma estrutura de categorias e subcategorias para popular dropdowns.
 */
function getCategoriesForDropdown(categoriesData) {
  const categories = {};
  categoriesData.slice(1).forEach(row => {
    const categoryName = row[0];
    const subcategoryName = row[1];
    const type = row[2];
    if (!categories[categoryName]) {
      categories[categoryName] = { type: type, subcategories: [] };
    }
    if (subcategoryName && !categories[categoryName].subcategories.includes(subcategoryName)) {
      categories[categoryName].subcategories.push(subcategoryName);
    }
  });
  return categories;
}

/**
 * Retorna uma lista de métodos de pagamento.
 */
function getPaymentMethodsForDropdown() {
  return ["Débito", "Crédito", "Dinheiro", "Pix", "Boleto", "Transferência Bancária"];
}

/**
 * Função chamada pelo menu do Add-on para abrir o dashboard numa janela modal.
 */
function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
      .setWidth(1200)
      .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Boas Contas Dashboard');
}


// ### INÍCIO DA ATUALIZAÇÃO ###
// NOVAS FUNÇÕES PARA GERIR CONTAS A PARTIR DO DASHBOARD
// ### FIM DA ATUALIZAÇÃO ###

/**
 * **ATUALIZADO COM RETORNO DINÂMICO**
 * Adiciona ou atualiza uma conta na aba 'Contas'.
 * @param {Object} accountData - Dados da conta a serem salvos.
 * @returns {Object} Objeto com status de sucesso ou erro e os dados atualizados do dashboard.
 */
function addOrUpdateAccountFromWeb(accountData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_CONTAS);
    if (!sheet) throw new Error(`Aba "${SHEET_CONTAS}" não encontrada.`);

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameColIndex = headers.indexOf('Nome da Conta');
    if (nameColIndex === -1) throw new Error("Coluna 'Nome da Conta' não encontrada.");

    const isEditing = !!accountData.originalName;
    let rowIndexToUpdate = -1;

    const newNameExists = data.slice(1).some((row, index) => {
        if (isEditing && row[nameColIndex] === accountData.originalName) {
            rowIndexToUpdate = index + 2;
            return false;
        }
        return row[nameColIndex] === accountData.name;
    });

    if (newNameExists) {
        throw new Error(`O nome de conta "${accountData.name}" já existe.`);
    }
    
    if (isEditing && rowIndexToUpdate === -1) {
        rowIndexToUpdate = data.slice(1).findIndex(row => row[nameColIndex] === accountData.originalName) + 2;
        if(rowIndexToUpdate < 2) throw new Error(`Conta original "${accountData.originalName}" não encontrada para edição.`);
    }

    const typeColIndex = headers.indexOf('Tipo');
    const bankColIndex = headers.indexOf('Banco');
    const initialBalanceColIndex = headers.indexOf('Saldo Inicial');

    if (isEditing) {
      sheet.getRange(rowIndexToUpdate, nameColIndex + 1).setValue(accountData.name);
      if (typeColIndex > -1) sheet.getRange(rowIndexToUpdate, typeColIndex + 1).setValue(accountData.type);
      if (bankColIndex > -1) sheet.getRange(rowIndexToUpdate, bankColIndex + 1).setValue(accountData.bank);
      if (initialBalanceColIndex > -1) sheet.getRange(rowIndexToUpdate, initialBalanceColIndex + 1).setValue(accountData.initialBalance);
      logToSheet(`Conta '${accountData.originalName}' atualizada para '${accountData.name}'.`, "INFO");

    } else {
      const newRow = Array(headers.length).fill('');
      newRow[nameColIndex] = accountData.name;
      newRow[typeColIndex] = accountData.type;
      newRow[bankColIndex] = accountData.bank;
      newRow[initialBalanceColIndex] = accountData.initialBalance;
      newRow[headers.indexOf('Status')] = 'Ativo';
      sheet.appendRow(newRow);
      logToSheet(`Nova conta '${accountData.name}' adicionada.`, "INFO");
    }

    atualizarSaldosDasContas();
    
    const today = new Date();
    const dashboardData = getDashboardData(today.getMonth() + 1, today.getFullYear());
    
    return { 
      success: true, 
      message: `Conta ${isEditing ? 'atualizada' : 'salva'} com sucesso.`,
      dashboardData: dashboardData
    };

  } catch (e) {
    handleError(e, "addOrUpdateAccountFromWeb");
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * **ATUALIZADO COM RETORNO DINÂMICO**
 * Exclui uma conta da aba 'Contas'.
 * @param {string} accountName - O nome da conta a ser excluída.
 * @returns {Object} Objeto com status de sucesso ou erro e os dados atualizados do dashboard.
 */
function deleteAccountFromWeb(accountName) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_CONTAS);
    if (!sheet) throw new Error(`Aba "${SHEET_CONTAS}" não encontrada.`);
    
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
    const dadosTransacoes = transacoesSheet.getDataRange().getValues();
    const contaColIndexTrans = dadosTransacoes[0].indexOf('Conta/Cartão');
    const hasTransactions = dadosTransacoes.slice(1).some(row => row[contaColIndexTrans] === accountName);

    if (hasTransactions) {
      throw new Error(`Não é possível excluir a conta "${accountName}", pois ela possui transações vinculadas.`);
    }

    const data = sheet.getDataRange().getValues();
    const nameColIndex = data[0].indexOf('Nome da Conta');
    const rowIndexToDelete = data.slice(1).findIndex(row => row[nameColIndex] === accountName);

    if (rowIndexToDelete !== -1) {
      sheet.deleteRow(rowIndexToDelete + 2);
      logToSheet(`Conta '${accountName}' excluída.`, "INFO");
      
      const today = new Date();
      const dashboardData = getDashboardData(today.getMonth() + 1, today.getFullYear());
      
      return { 
        success: true, 
        message: 'Conta excluída com sucesso.',
        dashboardData: dashboardData
      };
    } else {
      throw new Error(`Conta "${accountName}" não encontrada para exclusão.`);
    }
  } catch (e) {
    handleError(e, "deleteAccountFromWeb");
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}


// ===================================================================================
// ##      INÍCIO DO NOVO CÓDIGO DE PERSONALIZAÇÃO DO DASHBOARD (BACKEND)          ##
// ===================================================================================

/**
 * Salva as preferências de layout do dashboard para o usuário atual.
 * @param {string} layoutConfigJSON Uma string JSON contendo a ordem e a visibilidade dos cards.
 * @returns {Object} Um objeto com o status da operação.
 */
function saveDashboardLayout(layoutConfigJSON) {
  try {
    // Usamos 'user properties' para salvar dados específicos de cada usuário.
    PropertiesService.getUserProperties().setProperty('DASHBOARD_LAYOUT', layoutConfigJSON);
    Logger.log('Layout salvo para o usuário: ' + layoutConfigJSON);
    return { status: 'success', message: 'Layout salvo com sucesso!' };
  } catch (e) {
    Logger.log('Erro ao salvar o layout: ' + e.toString());
    return { status: 'error', message: 'Erro ao salvar layout: ' + e.toString() };
  }
}

/**
 * Recupera as preferências de layout do dashboard para o usuário atual.
 * @returns {string} Uma string JSON com as configurações de layout salvas, ou um layout padrão.
 */
function getDashboardLayout() {
  try {
    const savedLayout = PropertiesService.getUserProperties().getProperty('DASHBOARD_LAYOUT');
    if (savedLayout) {
      Logger.log('Layout encontrado para o usuário: ' + savedLayout);
      return savedLayout;
    } else {
      // Se o usuário nunca salvou um layout, retorna uma configuração padrão.
      // Os IDs aqui devem corresponder aos IDs dos cards no HTML.
      const defaultLayout = {
        hiddenCards: [] // Agora focado em cards individuais
      };
      Logger.log('Nenhum layout encontrado. Usando o padrão.');
      return JSON.stringify(defaultLayout);
    }
  } catch (e) {
    Logger.log('Erro ao recuperar o layout: ' + e.toString());
    // Em caso de erro, retorna o layout padrão para não quebrar a interface.
    const defaultLayout = {
        hiddenCards: []
    };
    return JSON.stringify(defaultLayout);
  }
}

// ===================================================================================
// ##       FIM DO NOVO CÓDIGO DE PERSONALIZAÇÃO DO DASHBOARD (BACKEND)            ##
// ===================================================================================
