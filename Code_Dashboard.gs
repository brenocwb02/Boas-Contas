/**
 * @file Code_Dashboard.gs
 * @description Funções do lado do servidor para o Dashboard Financeiro,
 * incluindo a coleta de dados, e operações de CRUD para transações via Web App.
 */

/**
 * ATUALIZADO: Serve o arquivo HTML do dashboard como um Web App com verificação de segurança.
 * Esta função agora inclui a verificação de licença como primeiro passo.
 * @param {Object} e O objeto de evento do Apps Script, contendo os parâmetros da URL.
 * @returns {HtmlOutput} O conteúdo HTML do dashboard ou uma página de erro.
 */
function doGet(e) {
  try {
    // PASSO 1: VERIFICAÇÃO DA LICENÇA DO PRODUTO
    if (!isLicenseValid()) {
      logToSheet("[Dashboard Access] Acesso bloqueado: Licença do produto inválida.", "ERROR");
      // Retorna uma página de erro de licença (você precisará criar este arquivo HTML)
      return HtmlService.createHtmlOutputFromFile('AcessoNegadoLicenca.html').setTitle("Licença Inválida");
    }

    // PASSO 2: VERIFICAÇÃO DO TOKEN DE ACESSO TEMPORÁRIO (lógica existente)
    const SCRIPT_VERSION = "1.3"; 
    const token = e.parameter.token;
    const cache = CacheService.getScriptCache();
    const cacheKey = `${CACHE_KEY_DASHBOARD_TOKEN}_${token}`;
    
    // 2.1. Verifica se um token foi fornecido na URL.
    if (!token) {
      logToSheet("[Dashboard Access] Acesso negado: nenhum token fornecido.", "WARN");
      return HtmlService.createHtmlOutput(
        '<h1><i class="fas fa-lock"></i> Acesso Negado</h1>' +
        '<p>Este link não é válido. Para acessar o dashboard, por favor, solicite um novo link de acesso através do bot no Telegram com o comando <code>/dashboard</code>.</p>' +
        '<style>body{font-family: sans-serif; text-align: center; padding-top: 50px; color: #333;} i{color: #d9534f;}</style>'
      ).setTitle("Acesso Negado");
    }

    // 2.2. Verifica se o token existe e é válido no cache.
    const expectedChatId = cache.get(cacheKey);
    if (!expectedChatId) {
      logToSheet(`[Dashboard Access] Acesso negado: token inválido ou expirado ('${token}').`, "WARN");
      return HtmlService.createHtmlOutput(
        '<h1><i class="fas fa-clock"></i> Link Inválido ou Expirado</h1>' +
        '<p>Este link de acesso não é mais válido. Ele pode ter expirado ou já ter sido utilizado. Por favor, solicite um novo com o comando <code>/dashboard</code> no Telegram.</p>' +
        '<style>body{font-family: sans-serif; text-align: center; padding-top: 50px; color: #f0ad4e;}</style>'
      ).setTitle("Link Expirado");
    }

    logToSheet(`[Dashboard Access] Acesso concedido para o chatId ${expectedChatId} com o token '${token}'.`, "INFO");

    // 4. Serve a página do dashboard.
    const template = HtmlService.createTemplateFromFile('Dashboard');
    template.chatId = expectedChatId; 
    template.version = SCRIPT_VERSION; // Passa a versão para o template HTML
    return template.evaluate()
        .setTitle('Dashboard Financeiro')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');

  } catch (error) {
    logToSheet(`[Dashboard Access] Erro crítico na função doGet: ${error.message}`, "ERROR");
    return HtmlService.createHtmlOutput(
        '<h1><i class="fas fa-server"></i> Erro Interno</h1>' +
        '<p>Ocorreu um erro inesperado ao tentar carregar o dashboard. O administrador foi notificado.</p>' +
        '<style>body{font-family: sans-serif; text-align: center; padding-top: 50px; color: #c9302c;}</style>'
    ).setTitle("Erro Interno");
  }
}



/**
 * Funcao auxiliar para obter mapeamento de meses (movida para o escopo global).
 * @returns {Object} Um mapa de nomes de meses (normalizados) para seus números (1-12).
 */
function getNomeMesesMap() {
  return {
    "janeiro": 1, "fevereiro": 2, "março": 3, "abril": 4, "maio": 5, "junho": 6,
    "julho": 7, "agosto": 8, "setembro": 9, "outubro": 10, "novembro": 11, "dezembro": 12
  };
}

/**
 * NOVO: Calcula os gastos de cartão de crédito para um mês e ano específicos.
 * Esta função filtra as transações pela data de vencimento para montar a "fatura" do mês selecionado.
 * @param {Array<Array<any>>} dadosTransacoes Dados da aba 'Transacoes'.
 * @param {Array<Array<any>>} dadosContas Dados da aba 'Contas'.
 * @param {number} mes O mês para o cálculo (1-12).
 * @param {number} ano O ano para o cálculo.
 * @returns {Array<Object>} Um array com os gastos mensais por cartão.
 */
function getCreditCardSpendingForMonth(dadosTransacoes, dadosContas, mes, ano) {
    const spendingMap = {}; // Estrutura: { nomeOriginal: { gasto: 0, limite: 0 } }
    const targetMonth = mes - 1; // Mês em JS é 0-indexado

    // 1. Inicializa o mapa com todos os cartões de crédito e seus limites
    dadosContas.slice(1).forEach(row => {
        const nomeOriginal = (row[0] || "").trim();
        const tipo = (row[1] || "").toLowerCase().trim();
        if (tipo === 'cartão de crédito' && nomeOriginal) {
            const limite = parseCurrencyValue(String(row[5] || '0'));
            spendingMap[nomeOriginal] = { gasto: 0, limite: limite };
        }
    });

    // 2. Itera sobre as transações para calcular os gastos do mês
    dadosTransacoes.slice(1).forEach(row => {
        const dataVencimento = parseData(row[10]); // Coluna K: Data de Vencimento
        const tipoTransacao = (row[4] || "").toLowerCase().trim();
        const conta = (row[7] || "").trim();
        const categoria = normalizarTexto(row[2]);
        const subcategoria = normalizarTexto(row[3]);

        // Condições para somar o gasto:
        // - A data de vencimento corresponde ao mês/ano do filtro
        // - É uma despesa
        // - A conta é um cartão de crédito que existe no nosso mapa
        // - Não é um pagamento de fatura (para evitar contagem dupla)
        if (
            dataVencimento &&
            dataVencimento.getMonth() === targetMonth &&
            dataVencimento.getFullYear() === ano &&
            tipoTransacao === 'despesa' &&
            spendingMap[conta] &&
            !(categoria === "contas a pagar" && subcategoria === "pagamento de fatura")
        ) {
            const valor = parseCurrencyValue(row[5]);
            spendingMap[conta].gasto += valor;
        }
    });

    // 3. Converte o mapa para um array de objetos para o frontend
    return Object.entries(spendingMap).map(([nome, data]) => ({
        nomeOriginal: nome,
        faturaAtual: round(data.gasto, 2), // Propriedade 'faturaAtual' para manter compatibilidade com o HTML
        limite: round(data.limite, 2)
    }));
}

/**
 * NOVO: Calcula o saldo das contas correntes e dinheiro até o final de um determinado mês/ano.
 * @param {Array<Array<any>>} dadosTransacoes Dados da aba 'Transacoes'.
 * @param {Array<Array<any>>} dadosContas Dados da aba 'Contas'.
 * @param {number} mes O mês para o cálculo (1-12).
 * @param {number} ano O ano para o cálculo.
 * @returns {Array<Object>} Um array com os saldos das contas no final do período.
 */
function getAccountBalancesForMonth(dadosTransacoes, dadosContas, mes, ano) {
    const balancesMap = {};
    const endDate = new Date(ano, mes, 0, 23, 59, 59); // Último dia do mês selecionado

    // 1. Inicializa o mapa com o saldo inicial de todas as contas correntes e dinheiro
    dadosContas.slice(1).forEach(row => {
        const nomeOriginal = (row[0] || "").trim();
        const tipo = (row[1] || "").toLowerCase().trim();
        if ((tipo === 'conta corrente' || tipo === 'dinheiro físico') && nomeOriginal) {
            const saldoInicial = parseCurrencyValue(String(row[3] || '0'));
            balancesMap[nomeOriginal] = saldoInicial;
        }
    });

    // 2. Itera sobre TODAS as transações ATÉ a data final para calcular o saldo
    dadosTransacoes.slice(1).forEach(row => {
        const dataTransacao = parseData(row[0]);
        const tipoTransacao = (row[4] || "").toLowerCase().trim();
        const conta = (row[7] || "").trim();
        const valor = parseCurrencyValue(row[5]);

        // Apenas processa transações que ocorreram ATÉ a data final e pertencem a uma conta relevante
        if (dataTransacao && dataTransacao <= endDate && balancesMap.hasOwnProperty(conta)) {
            if (tipoTransacao === 'receita') {
                balancesMap[conta] += valor;
            } else if (tipoTransacao === 'despesa') {
                balancesMap[conta] -= valor;
            }
        }
    });

    // 3. Converte o mapa para um array de objetos para o frontend
    return Object.entries(balancesMap).map(([nome, saldo]) => ({
        nomeOriginal: nome,
        saldo: round(saldo, 2)
    }));
}





/**
 * Coleta todos os dados necessários para o dashboard em uma única chamada.
 * Isso minimiza as chamadas de Apps Script do lado do cliente (HTML).
 * @param {number} mes O mês para filtrar os dados (1-12).
 * @param {number} ano O ano para filtrar os dados.
 * @returns {Object} Um objeto contendo os dados do resumo, saldos de contas, resumos de cartões, contas a pagar, transações recentes, metas e orçamento.
 */
function getDashboardData(mes, ano) {
  logToSheet(`Iniciando coleta de dados para o Dashboard para Mes: ${mes}, Ano: ${ano}`, "INFO");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
  const contasSheet = ss.getSheetByName(SHEET_CONTAS);
  const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);
  const metasSheet = ss.getSheetByName(SHEET_METAS); // Nova adição
  const orcamentoSheet = ss.getSheetByName(SHEET_ORCAMENTO); // Nova adição
  const categoriesSheet = ss.getSheetByName(SHEET_CATEGORIAS); // Adicionado para dropdown de categorias
  const invoicesSheet = ss.getSheetByName(SHEET_FATURAS); // Adicionado para resumos de cartões

  if (!transacoesSheet || !contasSheet || !contasAPagarSheet || !metasSheet || !orcamentoSheet || !categoriesSheet || !invoicesSheet) {
    logToSheet("Erro: Uma ou mais abas essenciais (Transacoes, Contas, Contas_a_Pagar, Metas, Orcamento, Categorias, Faturas) não foram encontradas para o dashboard.", "ERROR");
    throw new Error("Abas essenciais para o dashboard não encontradas. Verifique os nomes das abas.");
  }

  const dadosTransacoes = transacoesSheet.getDataRange().getValues();
  const dadosContas = contasSheet.getDataRange().getValues();
  const dadosContasAPagar = contasAPagarSheet.getDataRange().getValues();
  const dadosMetas = metasSheet.getDataRange().getValues();
  const dadosOrcamento = orcamentoSheet.getDataRange().getValues();
  const dadosCategorias = categoriesSheet.getDataRange().getValues(); // Dados para dropdown de categorias
  const dadosFaturas = invoicesSheet.getDataRange().getValues(); // Dados para resumos de cartões

  // Usa o mês e ano passados como parâmetro
  const currentMonth = mes - 1; // 0-indexed para operações com Date
  const currentYear = ano;
  const nomeMesAtual = getNomeMes(currentMonth); // Ex: "julho"

  // Mapeamento de ícones para categorias.
  // Será usado como fallback se a categoria da planilha não tiver um ícone.
  const categoryIconsMap = {
    "vida espiritual": "?", // Changed to a more common emoji
    "moradia": "🏠",
    "despesas fixas / contas": "🧾",
    "alimentacao": "🛒",
    "familia / filhos": "👨‍👩‍👧‍👦",
    "educacao e desenvolvimento": "🎓",
    "transporte": "🚗",
    "saude": "💊",
    "despesas pessoais": "👔",
    "impostos e taxas": "📊", // Changed to a more common emoji
    "lazer e entretenimento": "",
    "relacionamentos": "❤️", // Changed to a more common emoji
    "reserva / prevencao": "🛡️", // Changed to a more common emoji
    "investimentos / futuro": "📈",
    "receitas de trabalho": "💼",
    "apoio / ajuda externa": "🤝", // Changed to a more common emoji
    "outros ganhos": "🎁",
    "renda extra e investimentos": "💸",
    "artigos residenciais": "🛋️",
    "pag. de terceiros": "👥", // Changed to a more common emoji
    "conta bancaria": "🏦", // Changed to a more common emoji
    "transferencias": "🔄",
    // Adicione outras categorias conforme necessário
  };

  /**
   * CORRIGIDO: Garante que a entrada seja uma string antes de tentar usar .match().
   * Função auxiliar para extrair o ícone de uma string de categoria
   * Retorna o nome da categoria limpo e o ícone (se encontrado)
   */
  function extractIconAndCleanCategory(categoryString) {
    // CORREÇÃO CRÍTICA AQUI: Converte para string antes de qualquer operação.
    const str = String(categoryString || "");
    if (!str) return { cleanCategory: "", icon: "" };
    // Regex para capturar um emoji no início da string, seguido por espaço e o resto da string
    const match = str.match(/^(\p{Emoji}|\p{Emoji_Modifier_Base}|\p{Emoji_Component}|\p{Emoji_Modifier}|\p{Emoji_Presentation})\s*(.*)/u);
    if (match) {
      return { cleanCategory: match[2].trim(), icon: match[1] };
    }
    return { cleanCategory: str.trim(), icon: "" };
  }

  // --- 1. Resumo Mensal (Receitas, Despesas, Saldo Líquido) ---
  let totalReceitasMes = 0;
  let totalDespesasMesExcluindoPagamentosETransferencias = 0;
  
  // CORRIGIDO: Agora usa a data da transação para o resumo do mês
  for (let i = 1; i < dadosTransacoes.length; i++) {
    const dataRaw = dadosTransacoes[i][0];
    const data = parseData(dataRaw);
    const tipo = dadosTransacoes[i][4];
    // Usando a nova função parseCurrencyValue para garantir a correta interpretação do valor
    const valor = parseCurrencyValue(dadosTransacoes[i][5]) || 0; 
    const categoria = dadosTransacoes[i][2];
    const subcategoria = dadosTransacoes[i][3];

    if (!data || data.getMonth() !== currentMonth || data.getFullYear() !== currentYear) {
        continue;
    }

    if (tipo === "Receita") {
        const categoriaNormalizada = normalizarTexto(categoria);
        const subcategoriaNormalizada = normalizarTexto(subcategoria);
        // EXCLUSÃO ADICIONADA: Exclui transferências e recebimentos de pagamento de fatura das receitas totais
        if (
            !(categoriaNormalizada === "transferencias" && subcategoriaNormalizada === "entre contas") &&
            !(categoriaNormalizada === "pagamentos recebidos" && subcategoriaNormalizada === "pagamento de fatura")
        ) {
            totalReceitasMes += valor;
        }
    } else if (tipo === "Despesa") {
        const categoriaNormalizada = normalizarTexto(categoria);
        const subcategoriaNormalizada = normalizarTexto(subcategoria);
        if (
            !(categoriaNormalizada === "contas a pagar" && subcategoriaNormalizada === "pagamento de fatura") &&
            !(categoriaNormalizada === "transferencias" && subcategoriaNormalizada === "entre contas")
        ) {
            totalDespesasMesExcluindoPagamentosETransferencias += valor;
        }
    }
  }
  const saldoLiquidoMes = totalReceitasMes - totalDespesasMesExcluindoPagamentosETransferencias;

  const dashboardSummary = {
    totalReceitas: round(totalReceitasMes, 2),
    totalDespesas: round(totalDespesasMesExcluindoPagamentosETransferencias, 2),
    saldoLiquidoMes: round(saldoLiquidoMes, 2)
  };
  logToSheet(`Dashboard Summary: ${JSON.stringify(dashboardSummary)}`, "DEBUG");


  // --- 2. Saldos de Contas Correntes e Dinheiro (CORRIGIDO) ---
  const accountBalances = getAccountBalancesForMonth(dadosTransacoes, dadosContas, mes, ano);
  logToSheet(`Account Balances for end of ${mes}/${ano}: ${JSON.stringify(accountBalances)}`, "DEBUG");


  // --- 3. Resumo de Cartões de Crédito (CORRIGIDO) ---
  // A lógica agora calcula os gastos do mês selecionado, em vez de usar os saldos globais.
  const creditCardSummaries = getCreditCardSpendingForMonth(dadosTransacoes, dadosContas, mes, ano);
  logToSheet(`Credit Card Summaries for ${mes}/${ano}: ${JSON.stringify(creditCardSummaries)}`, "DEBUG");


  // --- 4. Contas a Pagar (para o mês atual) ---
  const billsToPay = [];
  // Verifica se dadosContasAPagar tem pelo menos uma linha (cabeçalhos)
  if (dadosContasAPagar.length > 0) {
    const contasAPagarHeaders = dadosContasAPagar[0]; // Pega os cabeçalhos da primeira linha
    const colDescricao = contasAPagarHeaders.indexOf('Descricao');
    const colValor = contasAPagarHeaders.indexOf('Valor');
    const colDataVencimento = contasAPagarHeaders.indexOf('Data de Vencimento');
    const colStatus = contasAPagarHeaders.indexOf('Status');
    const colRecorrente = contasAPagarHeaders.indexOf('Recorrente');

    // Verifica se as colunas essenciais foram encontradas
    if (colDescricao !== -1 && colValor !== -1 && colDataVencimento !== -1 && colStatus !== -1 && colRecorrente !== -1) {
      for (let i = 1; i < dadosContasAPagar.length; i++) { // Começa da linha 2 (índice 1)
        const row = dadosContasAPagar[i];
        const dataVencimentoRaw = row[colDataVencimento];
        const dataVencimento = parseData(dataVencimentoRaw);
        
        // Check if the bill is for the current month/year and is recurrent
        if (dataVencimento && dataVencimento.getMonth() === currentMonth && dataVencimento.getFullYear() === currentYear && normalizarTexto(row[colRecorrente]) === "verdadeiro") {
          // Usando a nova função parseCurrencyValue
          const valor = parseCurrencyValue(row[colValor]) || 0;

          billsToPay.push({
            descricao: (row[colDescricao] || "").toString().trim(),
            valor: round(valor, 2),
            dataVencimento: Utilities.formatDate(dataVencimento, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            status: (row[colStatus] || "").toString().trim()
          });
        }
      }
      // Sort bills by due date
      billsToPay.sort((a, b) => {
        const dateA = parseData(a.dataVencimento);
        const dateB = parseData(b.dataVencimento);
        return dateA.getTime() - dateB.getTime();
      });
    } else {
      logToSheet("AVISO: Colunas essenciais para 'Contas a Pagar' não encontradas nos cabeçalhos. Verifique a aba 'Contas_a_Pagar'.", "WARN");
    }
  } else {
    logToSheet("AVISO: Aba 'Contas_a_Pagar' está vazia ou contém apenas cabeçalhos. Nenhuma conta a pagar será processada.", "WARN");
  }
  logToSheet(`Bills To Pay: ${JSON.stringify(billsToPay)}`, "DEBUG");


  // --- 5. Últimos Lançamentos (ex: 10 mais recentes) ---
  const recentTransactions = [];
  const numTransactions = 10; // Quantidade de transações recentes a exibir.

  // Verifica se dadosTransacoes tem pelo menos uma linha (cabeçalhos)
  if (dadosTransacoes.length > 0) {
    for (let i = dadosTransacoes.length - 1; i > 0 && recentTransactions.length < numTransactions; i--) {
      const linha = dadosTransacoes[i];
      const dataObj = parseData(linha[0]);
      // Filtra transações recentes pelo mês e ano selecionados
      if (dataObj && dataObj.getMonth() === currentMonth && dataObj.getFullYear() === currentYear) {
        recentTransactions.push({
          // **CORREÇÃO AQUI**: Adicionada a propriedade 'id' que estava faltando.
          id: linha[13], 
          data: Utilities.formatDate(dataObj, Session.getScriptTimeZone(), "dd/MM/yyyy"),
          descricao: linha[1],
          categoria: linha[2], // Categoria com ícone (se presente na planilha)
          subcategoria: linha[3],
          tipo: linha[4],
          // Usando a nova função parseCurrencyValue
          valor: round(parseCurrencyValue(linha[5]), 2),
          metodoPagamento: linha[6],
          conta: linha[7],
          usuario: linha[11]
        });
      }
    }
  } else {
    logToSheet("AVISO: Aba 'Transacoes' está vazia ou contém apenas cabeçalhos. Nenhuma lançamento recente será processado.", "WARN");
  }
  logToSheet(`Recent Transactions: ${JSON.stringify(recentTransactions)}`, "DEBUG");

  // --- 6. Progresso das Metas (para o mês atual) ---
  const goalsProgress = [];
  // Verifica se dadosMetas tem pelo menos 3 linhas (cabeçalhos e linha de meses)
  if (dadosMetas.length > 2) {
    const cabecalhoMetas = dadosMetas[2]; // Linha 3 (índice 2) do cabeçalho da aba "Metas".
    let colMetaMes = -1;

    // Encontra a coluna correspondente ao mês e ano atual no cabeçalho das metas.
    for (let i = 2; i < cabecalhoMetas.length; i++) {
      // Apenas verifica se a string do cabeçalho contém o nome do mês atual (ex: "julho/2025")
      if (String(cabecalhoMetas[i]).toLowerCase().includes(`${nomeMesAtual.toLowerCase()}/${currentYear}`)) {
        colMetaMes = i;
        break;
      }
    }

    if (colMetaMes !== -1) {
      let metasMap = {}; // { 'categoria_subcategoria': { meta: X, gasto: Y } }

      // Inicializa metas do mês
      for (let i = 3; i < dadosMetas.length; i++) { // Começa da linha 4 (índice 3)
        const categoriaMetaWithIcon = (dadosMetas[i][0] || "").toString().trim();
        const { cleanCategory: categoriaMeta, icon: planilhaIconMeta } = extractIconAndCleanCategory(categoriaMetaWithIcon);
        const subcategoriaMeta = (dadosMetas[i][1] || "").toString().trim();
        const valorMetaTexto = dadosMetas[i][colMetaMes];

        if (categoriaMeta && subcategoriaMeta && valorMetaTexto) {
          // Usando a nova função parseCurrencyValue
          const meta = parseCurrencyValue(valorMetaTexto);
          if (!isNaN(meta) && meta > 0) {
            const key = normalizarTexto(`${categoriaMeta}_${subcategoriaMeta}`);
            metasMap[key] = {
              categoria: categoriaMeta,
              subcategoria: subcategoriaMeta,
              meta: meta,
              gasto: 0,
              percentage: 0,
              icon: planilhaIconMeta || categoryIconsMap[normalizarTexto(categoriaMeta)] || '' // Prioriza ícone da planilha, senão do mapa
            };
          }
        }
      }

      // Acumula gastos para as metas
      // CORRIGIDO: Usa a data de VENCIMENTO da transação para o cálculo das metas.
      for (let i = 1; i < dadosTransacoes.length; i++) {
        const data = parseData(dadosTransacoes[i][10]); // Data de Vencimento
        const tipo = dadosTransacoes[i][4];
        const categoriaTransacaoWithIcon = dadosTransacoes[i][2];
        const { cleanCategory: categoriaTransacao } = extractIconAndCleanCategory(categoriaTransacaoWithIcon);
        const subcategoriaTransacao = dadosTransacoes[i][3];
        // Usando a nova função parseCurrencyValue
        const valor = parseCurrencyValue(dadosTransacoes[i][5]);

        if (data && data.getMonth() === currentMonth && data.getFullYear() === currentYear && tipo === "Despesa") {
          const key = normalizarTexto(`${categoriaTransacao}_${subcategoriaTransacao}`);
          if (metasMap[key]) {
            metasMap[key].gasto = round(metasMap[key].gasto + valor, 2);
          }
        }
      }

      // Formata o resultado para o dashboard
      for (const key in metasMap) {
        const item = metasMap[key];
        // ALTERAÇÃO AQUI: Apenas adiciona a meta à lista se houver algum gasto nela.
        if (item.gasto > 0) {
            const percentage = item.meta > 0 ? round((item.gasto / item.meta) * 100, 2) : 0;
            goalsProgress.push({
              categoria: item.categoria,
              subcategoria: item.subcategoria,
              meta: item.meta,
              gasto: item.gasto,
              percentage: percentage,
              icon: item.icon // Mantém o ícone
            });
        }
      }
    } else {
      logToSheet("AVISO: Coluna de metas para o mês/ano atual não encontrada na aba 'Metas'. Verifique os cabeçalhos.", "WARN");
    }
  } else {
    logToSheet("AVISO: Aba 'Metas' está vazia ou não contém dados suficientes. Nenhuma meta será processada.", "WARN");
  }
  logToSheet(`Goals Progress: ${JSON.stringify(goalsProgress)}`, "DEBUG");

  // --- 7. Progresso do Orçamento (para o mês atual) ---
  const budgetProgress = [];
  
  logToSheet(`[Orcamento Debug] dadosOrcamento.length: ${dadosOrcamento.length}`, "DEBUG");
  if (dadosOrcamento.length > 0) {
      logToSheet(`[Orcamento Debug] Conteudo da primeira linha (dadosOrcamento[0]): ${JSON.stringify(dadosOrcamento[0])}`, "DEBUG");
  }

  // Verifica se dadosOrcamento tem pelo menos uma linha (cabeçalhos)
  if (dadosOrcamento.length > 0) {
    const orcamentoHeaders = dadosOrcamento[0]; // Assume que a primeira linha é o cabeçalho

    // Encontra os índices das colunas fixas
    const colOrcamentoCategoria = orcamentoHeaders.indexOf('Categoria');
    const colOrcamentoValorOrcado = orcamentoHeaders.indexOf('Valor Orcado');
    const colOrcamentoValorGasto = orcamentoHeaders.indexOf('Valor Gasto'); // Este valor será sobrescrito pelo cálculo
    const colOrcamentoMesReferencia = orcamentoHeaders.indexOf('Mes referencia'); 

    logToSheet(`[Orcamento Debug] indexOf('Categoria'): ${colOrcamentoCategoria}`, "DEBUG");
    logToSheet(`[Orcamento Debug] indexOf('Valor Orcado'): ${colOrcamentoValorOrcado}`, "DEBUG");
    logToSheet(`[Orcamento Debug] indexOf('Valor Gasto'): ${colOrcamentoValorGasto}`, "DEBUG");
    logToSheet(`[Orcamento Debug] indexOf('Mes referencia'): ${colOrcamentoMesReferencia}`, "DEBUG");


    // Verifica se todas as colunas essenciais foram encontradas
    if (colOrcamentoCategoria !== -1 && colOrcamentoValorOrcado !== -1 && colOrcamentoValorGasto !== -1 && colOrcamentoMesReferencia !== -1) {
      let orcamentoMap = {}; // { 'categoria': { orcado: X, gasto: Y } }

      // Inicializa orçamento do mês com base na aba Orcamento
      for (let i = 1; i < dadosOrcamento.length; i++) { // Começa da linha 2 (índice 1)
        const row = dadosOrcamento[i];
        const mesReferenciaRaw = row[colOrcamentoMesReferencia]; // Não converta para string ainda
        const categoriaOrcamentoWithIcon = (row[colOrcamentoCategoria] || "").toString().trim();
        const { cleanCategory: categoriaOrcamento, icon: planilhaIconOrcamento } = extractIconAndCleanCategory(categoriaOrcamentoWithIcon);

        // CORREÇÃO AQUI: Parse a data de referência da planilha
        const dataReferenciaOrcamento = parseData(mesReferenciaRaw); 
        
        // ADIÇÃO DE LOGS PARA DEPURAR CATEGORIA E DATA
        logToSheet(`[Orcamento Debug] Linha ${i+1}: Mes Ref Raw: "${mesReferenciaRaw}", Categoria Raw: "${categoriaOrcamentoWithIcon}", Categoria Limpa: "${categoriaOrcamento}"`, "DEBUG");
        logToSheet(`[Orcamento Debug] Data Referencia Orcamento (Parsed): ${dataReferenciaOrcamento ? dataReferenciaOrcamento.toLocaleDateString() : 'N/A'}`, "DEBUG");

        let mesRefNum = -1;
        let anoRefNum = -1;

        if (dataReferenciaOrcamento) {
          mesRefNum = dataReferenciaOrcamento.getMonth(); // 0-indexed
          anoRefNum = dataReferenciaOrcamento.getFullYear();
        }

        logToSheet(`[Orcamento Debug] Mes Ref Num (0-indexed): ${mesRefNum}, Ano Ref Num: ${anoRefNum}`, "DEBUG");
        logToSheet(`[Orcamento Debug] Current Month (0-indexed): ${currentMonth}, Current Year: ${currentYear}`, "DEBUG");
        logToSheet(`[Orcamento Debug] Comparacao Mes: ${mesRefNum === currentMonth}, Comparacao Ano: ${anoRefNum === currentYear}`, "DEBUG");


        // Filtra para o mês e ano de referência atual
        if (categoriaOrcamento && mesRefNum === currentMonth && anoRefNum === currentYear) {
          const valorOrcado = parseCurrencyValue(row[colOrcamentoValorOrcado] || '0');
          // O valor gasto da planilha é apenas um ponto de partida, o real será recalculado
          const valorGastoInicial = parseCurrencyValue(row[colOrcamentoValorGasto] || '0'); 

          const key = normalizarTexto(categoriaOrcamento);
          orcamentoMap[key] = {
            categoria: categoriaOrcamento,
            orcado: valorOrcado,
            gasto: 0, // Inicia com 0, será preenchido pelas transações
            icon: planilhaIconOrcamento || categoryIconsMap[normalizarTexto(categoriaOrcamento)] || '' // Prioriza ícone da planilha, senão do mapa
          };
          logToSheet(`[Orcamento Debug] Orcamento inicializado para categoria "${categoriaOrcamento}": Orcado ${valorOrcado}`, "DEBUG");
        } else {
          // ADIÇÃO DE LOGS DETALHADOS PARA QUANDO A LINHA É IGNORADA
          logToSheet(`[Orcamento Debug] Linha ${i+1} ignorada (nao corresponde ao mes/ano atual ou categoria vazia). Condicao: Categoria(${!!categoriaOrcamento}) && Mes(${mesRefNum === currentMonth}) && Ano(${anoRefNum === currentYear}).`, "DEBUG");
        }
      }

      // Recalcula/confirma gastos para o orçamento com base nas transações
      // CORRIGIDO: Usa a data de VENCIMENTO da transação para o cálculo do orçamento.
      for (let i = 1; i < dadosTransacoes.length; i++) {
        const data = parseData(dadosTransacoes[i][10]); // Data de Vencimento
        const tipo = dadosTransacoes[i][4];
        const categoriaTransacaoWithIcon = dadosTransacoes[i][2];
        const { cleanCategory: categoriaTransacao } = extractIconAndCleanCategory(categoriaTransacaoWithIcon);
        const subcategoria = dadosTransacoes[i][3]; // Necessário para excluir pagamentos de fatura
        const valor = parseCurrencyValue(dadosTransacoes[i][5]);

        if (data && data.getMonth() === currentMonth && data.getFullYear() === currentYear && tipo === "Despesa") {
          const categoriaNormalizada = normalizarTexto(categoriaTransacao);
          const subcategoriaNormalizada = normalizarTexto(subcategoria);

          // Exclui pagamentos de fatura e transferências do cálculo do orçamento
          if (!(categoriaNormalizada === "contas a pagar" && subcategoriaNormalizada === "pagamento de fatura") &&
              !(categoriaNormalizada === "transferencias" && subcategoriaNormalizada === "entre contas")) {
              
              const key = normalizarTexto(categoriaTransacao);
              if (orcamentoMap[key]) {
                orcamentoMap[key].gasto = round(orcamentoMap[key].gasto + valor, 2);
                logToSheet(`[Orcamento Debug] Gasto de ${valor} adicionado para categoria "${categoriaTransacao}". Gasto atual: ${orcamentoMap[key].gasto}`, "DEBUG");
              } else {
                logToSheet(`[Orcamento Debug] Transacao de despesa para categoria "${categoriaTransacao}" nao encontrada no orcamentoMap.`, "DEBUG");
              }
          } else {
            logToSheet(`[Orcamento Debug] Transacao de despesa para categoria "${categoriaTransacao}" (${subcategoria}) excluida do calculo do orcamento.`, "DEBUG");
          }
        } else {
          logToSheet(`[Orcamento Debug] Transacao ${i+1} ignorada para calculo do orcamento (data ou tipo).`, "DEBUG");
        }
      }

      // Formata o resultado para o dashboard
      for (const key in orcamentoMap) {
        const item = orcamentoMap[key];
        const percentage = item.orcado > 0 ? round((item.gasto / item.orcado) * 100, 2) : 0;
        budgetProgress.push({
          categoria: item.categoria,
          orcado: item.orcado,
          gasto: item.gasto,
          percentage: percentage,
          icon: item.icon // Mantém o ícone
        });
      }
    } else {
      logToSheet("ERRO: Colunas 'Categoria', 'Valor Orcado', 'Valor Gasto' ou 'Mes referencia' não encontradas na aba 'Orcamento'. Verifique os cabeçalhos.", "ERROR");
      // NOVO LOG: Indica qual coluna pode estar faltando
      logToSheet(`[Orcamento Debug] Status de colunas: Categoria: ${colOrcamentoCategoria !== -1}, Valor Orcado: ${colOrcamentoValorOrcado !== -1}, Valor Gasto: ${colOrcamentoValorGasto !== -1}, Mes referencia: ${colOrcamentoMesReferencia !== -1}`, "ERROR");
    }
  } else {
    logToSheet("AVISO: Aba 'Orcamento' está vazia ou contém apenas cabeçalhos. Nenhum orçamento será processado.", "WARN");
  }
  // Adiciona um log para depuração: se o array budgetProgress estiver vazio, indica que nenhum dado foi processado.
  if (budgetProgress.length === 0) {
      logToSheet("AVISO: 'budgetProgress' esta vazio. Verifique se a aba 'Orcamento' tem dados para o mes/ano atual e se os cabecalhos estao corretos.", "WARN");
  }
  logToSheet(`Budget Progress: ${JSON.stringify(budgetProgress)}`, "DEBUG");


  // --- 8. Despesas por Categoria para o Gráfico ---
  // A estrutura de expensesByCategory será um array de objetos para facilitar o uso dos ícones no gráfico
  const expensesByCategoryArray = [];
  const tempExpensesMap = {}; // Para somar valores por categoria

  // Verifica se dadosTransacoes tem pelo menos uma linha (cabeçalhos)
  if (dadosTransacoes.length > 0) {
    // CORRIGIDO: Usa a data de VENCIMENTO da transação para o cálculo do gráfico de despesas.
    for (let i = 1; i < dadosTransacoes.length; i++) {
      const dataRaw = dadosTransacoes[i][10]; // Data de Vencimento
      const data = parseData(dataRaw);
      const categoriaTransacaoWithIcon = dadosTransacoes[i][2];
      const { cleanCategory: categoriaTransacao, icon: planilhaIconTransacao } = extractIconAndCleanCategory(categoriaTransacaoWithIcon);
      const tipo = dadosTransacoes[i][4];
      // Usando a nova função parseCurrencyValue
      const valor = parseCurrencyValue(dadosTransacoes[i][5]);

      if (!data || data.getMonth() !== currentMonth || data.getFullYear() !== currentYear) {
          continue;
      }

      if (tipo === "Despesa") {
          const categoriaNormalizada = normalizarTexto(categoriaTransacao);
          const subcategoriaNormalizada = normalizarTexto(dadosTransacoes[i][3]); // Subcategoria
          // Excluir pagamentos de fatura e transferências para o gráfico de despesas "reais"
          if (!(categoriaNormalizada === "contas a pagar" && subcategoriaNormalizada === "pagamento de fatura") &&
              !(categoriaNormalizada === "transferencias" && subcategoriaNormalizada === "entre contas")) {
              
              if (!tempExpensesMap[categoriaNormalizada]) {
                  tempExpensesMap[categoriaNormalizada] = {
                      categoriaOriginal: categoriaTransacao,
                      total: 0,
                      icon: planilhaIconTransacao || categoryIconsMap[categoriaNormalizada] || '' // Prioriza ícone da planilha, senão do mapa
                  };
              }
              tempExpensesMap[categoriaNormalizada].total += valor;
          }
      }
    }

    // Converte o mapa temporário para o array final
    for (const key in tempExpensesMap) {
        expensesByCategoryArray.push({
            category: tempExpensesMap[key].categoriaOriginal,
            value: round(tempExpensesMap[key].total, 2),
            icon: tempExpensesMap[key].icon
        });
    }
  } else {
    logToSheet("AVISO: Aba 'Transacoes' está vazia ou contém apenas cabeçalhos. Nenhuma despesa por categoria será processada.", "WARN");
  }

  logToSheet(`Expenses By Category Array: ${JSON.stringify(expensesByCategoryArray)}`, "DEBUG");


  logToSheet("Coleta de dados para o Dashboard concluida.", "INFO");
  return {
    summary: dashboardSummary,
    accountBalances: accountBalances,
    creditCardSummaries: creditCardSummaries,
    billsToPay: billsToPay,
    recentTransactions: recentTransactions,
    goalsProgress: goalsProgress, // Adiciona metas
    budgetProgress: budgetProgress, // Adiciona orçamento
    expensesByCategory: expensesByCategoryArray, // Adiciona dados para o gráfico
    accounts: getAccountsForDropdown(dadosContas), // Adicionado para o dropdown
    categories: getCategoriesForDropdown(dadosCategorias), // Adicionado para o dropdown
    paymentMethods: getPaymentMethodsForDropdown() // Adicionado para o dropdown
  };
}

/**
 * NOVO: Busca e retorna todas as transações de uma categoria específica para um determinado mês e ano.
 * Chamada pelo gráfico clicável no dashboard.
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
    const targetMonth = month - 1; // JS month is 0-indexed

    const transactions = [];

    // Função auxiliar para remover emojis e normalizar, se necessário.
    function extractIconAndCleanCategory(categoryString) {
      // CORREÇÃO CRÍTICA AQUI: Converte para string antes de qualquer operação.
      const str = String(categoryString || "");
      if (!str) return { cleanCategory: "", icon: "" };
      const match = str.match(/^(\p{Emoji}|\p{Emoji_Modifier_Base}|\p{Emoji_Component}|\p{Emoji_Modifier}|\p{Emoji_Presentation})\s*(.*)/u);
      if (match) {
        return { cleanCategory: match[2].trim(), icon: match[1] };
      }
      return { cleanCategory: str.trim(), icon: "" };
    }

    for (let i = 1; i < dadosTransacoes.length; i++) {
      const row = dadosTransacoes[i];
      const data = parseData(row[0]);
      const { cleanCategory: transCategory } = extractIconAndCleanCategory(row[2]);

      if (data && data.getMonth() === targetMonth && data.getFullYear() === year && normalizarTexto(transCategory) === normalizarTexto(categoryName)) {
        transactions.push({
          data: Utilities.formatDate(data, Session.getScriptTimeZone(), "dd/MM/yyyy"),
          descricao: row[1],
          categoria: row[2],
          subcategoria: row[3],
          tipo: row[4],
          valor: parseCurrencyValue(row[5]),
          conta: row[7]
        });
      }
    }
    logToSheet(`[Dashboard] Encontradas ${transactions.length} transações para a categoria '${categoryName}'.`, "DEBUG");
    // Ordena as transações pela data, da mais recente para a mais antiga
    return transactions.sort((a, b) => {
        const dateA = new Date(a.data.split('/').reverse().join('-'));
        const dateB = new Date(b.data.split('/').reverse().join('-'));
        return dateB - dateA;
    });
  } catch (e) {
    logToSheet(`[Dashboard] ERRO em getTransactionsByCategory: ${e.message}`, "ERROR");
    // Re-lança o erro para que o handler onFailure do lado do cliente seja acionado
    throw new Error(`Erro ao buscar transações: ${e.message}`);
  }
}


/**
 * NOVO: Função robusta para parsear valores monetários em formato brasileiro ou internacional.
 * Lida com "R$", separadores de milhares (ponto ou vírgula) e separadores decimais (vírgula ou ponto).
 * @param {any} valueString O valor a ser parseado (pode ser string ou number).
 * @returns {number} O valor numérico parseado.
 */
function parseCurrencyValue(valueString) {
  if (typeof valueString === 'number') {
    return valueString;
  }
  let cleaned = String(valueString).replace("R$", "").trim();

  const lastCommaIndex = cleaned.lastIndexOf(',');
  const lastDotIndex = cleaned.lastIndexOf('.');

  if (lastCommaIndex > lastDotIndex) { // Formato brasileiro: 1.234,56
    cleaned = cleaned.replace(/\./g, ''); // Remove separadores de milhares (pontos)
    cleaned = cleaned.replace(',', '.');  // Substitui a vírgula decimal por ponto
  } else if (lastDotIndex > lastCommaIndex) { // Formato internacional: 1,234.56 ou 1234.56
    cleaned = cleaned.replace(/,/g, ''); // Remove separadores de milhares (vírgulas)
    // O ponto decimal já está correto
  }
  // Se não houver vírgula nem ponto, parseFloat lidará com isso (ex: "123")

  return parseFloat(cleaned) || 0; // Garante que retorne 0 se o parse falhar
}

/**
 * Filtra as transações pelo mês e ano.
 * @param {Array<Array>} data Array de transações.
 * @param {number} month Mês (1-12).
 * @param {number} year Ano.
 * @returns {Array<Object>} Transações filtradas como objetos.
 */
function filterTransactionsByMonthYear(data, month, year) {
  return data.map(row => ({
    data: new Date(row[0]),
    descricao: row[1],
    tipo: row[2],
    valor: parseFloat(row[3]),
    conta: row[4],
    categoria: row[5],
    subcategoria: row[6],
    metodoPagamento: row[7],
    parcelas: parseInt(row[8]),
    pago: row[9] === 'Sim' // Converte para booleano
  })).filter(transaction => {
    const transactionDate = transaction.data;
    return transactionDate.getMonth() + 1 === month && transactionDate.getFullYear() === year;
  });
}

/**
 * Calcula o resumo financeiro (receitas, despesas, saldo).
 * @param {Array<Object>} transactions Transações filtradas.
 * @returns {Object} Resumo financeiro.
 */
function calculateSummary(transactions) {
  let totalReceitas = 0;
  let totalDespesas = 0;

  transactions.forEach(t => {
    if (t.tipo === 'Receita') {
      totalReceitas += t.valor;
    } else if (t.tipo === 'Despesa') {
      totalDespesas += t.valor;
    }
  });

  const saldoLiquidoMes = totalReceitas - totalDespesas;

  return {
    totalReceitas: totalReceitas,
    totalDespesas: totalDespesas,
    saldoLiquidoMes: saldoLiquidoMes
  };
}

// As funções calculateAccountBalances, calculateCreditCardSummaries, getRecentTransactions,
// getBillsToPay, calculateGoalsProgress, calculateBudgetProgress, calculateExpensesByCategory
// foram removidas daqui pois a lógica principal de cálculo de saldos e resumos de cartões
// é feita por 'atualizarSaldosDasContas()' e 'globalThis.saldosCalculados',
// e as outras funções de cálculo de dashboard já existiam e foram mantidas.

/**
 * Adiciona uma nova transação à planilha "Transacoes" a partir do formulário web.
 * @param {Object} transactionData Objeto contendo os dados da transação.
 * Esperado: date, description, type, value, account, category, subcategory,
 * paymentMethod, installments, currentInstallment (opcional),
 * dueDate (opcional), user (opcional), status (opcional).
 * @returns {Object} Um objeto indicando sucesso ou falha.
 */
function addTransactionFromWeb(transactionData) {
  Logger.log('[addTransactionFromWeb] Preparando para adicionar transação na planilha \'Transacoes\'.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Transacoes'); // Certifique-se de que o nome da sua aba está correto

  if (!sheet) {
    Logger.log('Erro: Planilha \'Transacoes\' não encontrada.');
    return { success: false, message: 'Planilha \'Transacoes\' não encontrada.' };
  }

  // Obter a última linha com conteúdo para adicionar a nova transação abaixo
  // Se a planilha estiver vazia (além do cabeçalho), getLastRow() pode retornar 0 ou 1.
  // appendRow adicionará à próxima linha vazia, o que geralmente funciona bem.
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  Logger.log('[addTransactionFromWeb] Última linha detectada (getLastRow()): ' + lastRow);
  Logger.log('[addTransactionFromWeb] Última coluna detectada (getLastColumn()): ' + lastColumn);

  // --- TRATAMENTO DA DATA DE VENCIMENTO (DUPLICADO DO CÓDIGO ANTERIOR PARA GARANTIR) ---
  let formattedDueDate = '';
  if (transactionData.dueDate) {
    try {
      // Divide a string da data (YYYY-MM-DD) em partes
      const parts = transactionData.dueDate.split('-');
      const year = parseInt(parts[0]);
      const month = parseInt(parts[1]) - 1; // Mês é 0-indexed no JavaScript (e no Apps Script Date object)
      const day = parseInt(parts[2]);

      // Cria um objeto Date no fuso horário do script.
      // Isso é crucial para evitar que a data seja interpretada como UTC e ajuste para o dia anterior.
      const dateObject = new Date(year, month, day);

      // Formata a data para o formato desejado na planilha (ex: DD/MM/AAAA)
      // Usar Session.getScriptTimeZone() garante que o fuso horário configurado para o seu script
      // seja respeitado, evitando o problema de -1 dia.
      formattedDueDate = Utilities.formatDate(dateObject, Session.getScriptTimeZone(), "dd/MM/yyyy");
      Logger.log('[addTransactionFromWeb] Data de vencimento formatada: ' + formattedDueDate);
    } catch (e) {
      Logger.log('Erro ao formatar data de vencimento: ' + e.message);
      // Em caso de erro, use a string original ou deixe vazio
      formattedDueDate = transactionData.dueDate;
    }
  }

  // --- TRATAMENTO DA DATA DA TRANSAÇÃO (ASSUMINDO QUE JÁ ESTÁ EM YYYY-MM-DD) ---
  let formattedTransactionDate = '';
  if (transactionData.date) {
    try {
      const parts = transactionData.date.split('-');
      const year = parseInt(parts[0]);
      const month = parseInt(parts[1]) - 1;
      const day = parseInt(parts[2]);
      const dateObject = new Date(year, month, day);
      formattedTransactionDate = Utilities.formatDate(dateObject, Session.getScriptTimeZone(), "dd/MM/yyyy");
      Logger.log('[addTransactionFromWeb] Data da transação formatada: ' + formattedTransactionDate);
    } catch (e) {
      Logger.log('Erro ao formatar data da transação: ' + e.message);
      formattedTransactionDate = transactionData.date;
    }
  }


  // Mapear os dados para a ordem das colunas na sua planilha
  // Ajuste a ordem e o número de colunas conforme a sua planilha 'Transacoes'
  const newRow = [
    formattedTransactionDate, // Coluna A: Data
    transactionData.description, // Coluna B: Descrição
    transactionData.category || '', // Coluna C: Categoria (vazio se não aplicável)
    transactionData.subcategory || '', // Coluna D: Subcategoria (vazio se não aplicável)
    transactionData.type, // Coluna E: Tipo (Despesa, Receita, Transferência)
    transactionData.value, // Coluna F: Valor
    transactionData.paymentMethod || '', // Coluna G: Método de Pagamento
    transactionData.account, // Coluna H: Conta/Cartão
    transactionData.installments, // Coluna I: Parcelas
    1, // Coluna J: Parcela Atual (assumindo 1 para nova transação, ajuste se tiver lógica de parcelamento)
    formattedDueDate, // Coluna K: Data de Vencimento
    '', // Coluna L: Observações (vazio)
    'Ativo', // Coluna M: Status (ex: Ativo)
    Utilities.getUuid(), // Coluna N: ID Único
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss") // Coluna O: Data de Registro
  ];

  Logger.log('[addTransactionFromWeb] Nova linha a ser adicionada: ' + JSON.stringify(newRow));

  try {
    // Adiciona a nova linha à planilha
    sheet.appendRow(newRow);
    Logger.log('[addTransactionFromWeb] Transação adicionada com sucesso.');
    return { success: true, message: 'Transação adicionada com sucesso.' };
  } catch (e) {
    Logger.log('Erro ao adicionar transação à planilha: ' + e.message);
    return { success: false, message: 'Erro ao adicionar transação: ' + e.message };
  }
}

/**
 * Retorna uma lista de contas e cartões para popular um dropdown.
 * @param {Array<Array>} accountsData Dados brutos da planilha "Contas".
 * @returns {Array<Object>} Lista de objetos { nomeOriginal: string, tipo: string }.
 */
function getAccountsForDropdown(accountsData) {
  // Ignora o cabeçalho
  const dataWithoutHeader = accountsData.slice(1); 
  return dataWithoutHeader.map(row => ({
    nomeOriginal: row[0], // Nome da Conta (Coluna A)
    tipo: row[1] // Tipo de Conta (e.g., 'conta corrente', 'cartao de credito', 'dinheiro fisico') (Coluna B)
  }));
}

/**
 * Retorna uma estrutura de categorias e subcategorias para popular dropdowns.
 * @param {Array<Array>} categoriesData Dados brutos da planilha "Categorias".
 * @returns {Object} Objeto com categorias principais e suas subcategorias.
 * Ex: { "Categoria Principal": { type: "Despesa", subcategories: ["Subcategoria1", "Subcategoria2"] }, ... }
 */
function getCategoriesForDropdown(categoriesData) {
  const categories = {};
  // Ignora o cabeçalho
  const dataWithoutHeader = categoriesData.slice(1); 
  dataWithoutHeader.forEach(row => {
    const categoryName = row[0]; // Categoria Principal
    const subcategoryName = row[1]; // Subcategoria
    const type = row[2]; // Tipo (Receita/Despesa)

    if (!categories[categoryName]) {
      categories[categoryName] = {
        type: type,
        subcategories: []
      };
    }
    if (subcategoryName && !categories[categoryName].subcategories.includes(subcategoryName)) {
      categories[categoryName].subcategories.push(subcategoryName);
    }
  });
  return categories;
}

/**
 * Retorna uma lista de métodos de pagamento.
 * Pode ser de uma planilha ou uma lista fixa.
 * @returns {Array<string>} Lista de métodos de pagamento.
 */
function getPaymentMethodsForDropdown() {
  // Exemplo: lista fixa. Se tiver uma planilha "Metodos de Pagamento", buscar de lá.
  return ["Débito", "Crédito", "Dinheiro", "Pix", "Boleto", "Transferência Bancária"];
}

// Função para incluir arquivos HTML/CSS/JS (se você tiver múltiplos arquivos)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

/**
 * **FUNÇÃO CORRIGIDA**
 * Deleta uma transação da planilha 'Transacoes' e atualiza os saldos.
 * Esta função é chamada pelo Dashboard HTML.
 * @param {string} transactionId O ID único da transação a ser deletada.
 * @returns {object} Um objeto com status de sucesso ou erro.
 */
function deleteTransactionFromWeb(transactionId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // CORREÇÃO 1: Usar a constante correta 'SHEET_TRANSACOES'
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);

    if (!transacoesSheet) {
      throw new Error(`Planilha "${SHEET_TRANSACOES}" não encontrada.`);
    }

    const data = transacoesSheet.getDataRange().getValues();
    const headers = data[0];
    const idColumnIndex = headers.indexOf('ID Transacao');

    if (idColumnIndex === -1) {
      throw new Error("Coluna 'ID Transacao' não encontrada na planilha.");
    }

    let rowIndexToDelete = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColumnIndex] == transactionId) {
        rowIndexToDelete = i + 1;
        break;
      }
      // Se a transação excluída tiver parcelas, exclui todas as parcelas relacionadas
      // A implementação de parcelas requer um ID de "transação pai" ou uma forma de agrupar.
      // Assumindo que transactionId pode ser o "id da transação principal" para exclusão.
      const baseTransactionId = transactionId.split('-')[0];
      if (String(data[i][idColumnIndex]).startsWith(baseTransactionId + '-')) {
          // Se encontrou uma parcela relacionada, não exclui aqui, apenas informa.
          // A exclusão de parcelas é um cenário complexo e deve ser tratada com cuidado.
          // Por simplicidade, este código exclui apenas a linha que corresponde EXATAMENTE ao ID.
          // Se você quiser excluir TODAS as parcelas de uma transação principal,
          // a lógica aqui precisaria ser mais elaborada (por exemplo, coletar todos os `rowIndexToDelete`
          // e depois excluir as linhas em lote, de baixo para cima).
      }
    }

    if (rowIndexToDelete !== -1) {
      transacoesSheet.deleteRow(rowIndexToDelete);
      logToSheet(`Transação com ID ${transactionId} deletada da linha ${rowIndexToDelete}.`, "INFO");

      // CORREÇÃO 2: Chamar a função correta para atualizar os saldos
      atualizarSaldosDasContas();

      return { success: true, message: `Transação ${transactionId} excluída com sucesso.` };
    } else {
      return { success: false, message: `Transação com ID ${transactionId} não encontrada.` };
    }
  } catch (e) {
    logToSheet(`Erro ao deletar transação: ${e.message}`, "ERROR");
    return { success: false, message: `Erro ao excluir transação: ${e.message}` };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Atualiza uma transação existente com verificação de cabeçalhos.
 * @param {Object} transactionData Objeto com os dados da transação.
 * @returns {Object} Objeto indicando sucesso ou falha.
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

    // CORREÇÃO: Ajustar os nomes das colunas para corresponder aos cabeçalhos da planilha
    const requiredColumns = ["Data", "Descricao", "Categoria", "Subcategoria", "Tipo", "Valor", "Metodo de Pagamento", "Conta/Cartão", "Parcelas Totais", "Data de Vencimento", "ID Transacao"];
    const missingColumns = requiredColumns.filter(col => colMap[col.trim()] === undefined);

    if (missingColumns.length > 0) {
      const errorMessage = `As seguintes colunas não foram encontradas na aba '${SHEET_TRANSACOES}': ${missingColumns.join(', ')}. Por favor, verifique se os nomes dos cabeçalhos na sua planilha estão corretos e sem espaços extras. Cabeçalhos encontrados: [${headers.join(' | ')}]`;
      throw new Error(errorMessage);
    }

    const idColumn = colMap["ID Transacao"];
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColumn] === transactionData.id) {
        const rowIndex = i + 1;
        
        // CORREÇÃO DE FUSO HORÁRIO: Adiciona 'T00:00:00' para garantir que a data seja interpretada no fuso horário local do script.
        sheet.getRange(rowIndex, colMap["Data"] + 1).setValue(new Date(transactionData.date + 'T00:00:00'));
        sheet.getRange(rowIndex, colMap["Descricao"] + 1).setValue(transactionData.description);
        sheet.getRange(rowIndex, colMap["Categoria"] + 1).setValue(transactionData.category);
        sheet.getRange(rowIndex, colMap["Subcategoria"] + 1).setValue(transactionData.subcategory);
        sheet.getRange(rowIndex, colMap["Tipo"] + 1).setValue(transactionData.type);
        sheet.getRange(rowIndex, colMap["Valor"] + 1).setValue(parseCurrencyValue(String(transactionData.value)));
        sheet.getRange(rowIndex, colMap["Metodo de Pagamento"] + 1).setValue(transactionData.paymentMethod); // Ajustado
        sheet.getRange(rowIndex, colMap["Conta/Cartão"] + 1).setValue(transactionData.account); // Ajustado
        sheet.getRange(rowIndex, colMap["Parcelas Totais"] + 1).setValue(parseInt(transactionData.installments));
        sheet.getRange(rowIndex, colMap["Data de Vencimento"] + 1).setValue(new Date((transactionData.dueDate || transactionData.date) + 'T00:00:00'));
        
        atualizarSaldosDasContas();
        return { success: true, message: 'Transação atualizada com sucesso.' };
      }
    }
    throw new Error("Transação não encontrada para atualização.");
  } catch (e) {
    logToSheet(`Erro em updateTransactionFromWeb: ${e.message}`, "ERROR");
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}


// --- Funções Auxiliares (mantidas do seu código original) ---

/**
 * CORREÇÃO: Adicionada a função getColumnMap que estava em falta.
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
