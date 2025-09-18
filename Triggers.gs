/**
 * @file Triggers.gs
 * @description Contém a lógica para o acionador mestre que gere todas as tarefas automáticas.
 */

const SHEET_TRANSACOES_ARQUIVO = "Transacoes_Arquivo";

/**
 * Função mestra para ser chamada pelo único acionador de tempo do sistema.
 * Ela decide qual tarefa automática deve ser executada com base na hora e no dia.
 */
function masterTriggerHandler() {
  const agora = new Date();
  const hora = agora.getHours();
  const diaDoMes = agora.getDate();

  logToSheet("[Master Trigger] Acionador mestre executado.", "INFO");

  // 1. Executa a verificação de notificações a cada hora.
  try {
    checkAndSendNotifications();
  } catch (e) {
    handleError(e, "masterTriggerHandler -> checkAndSendNotifications");
  }

  // 2. Gera contas recorrentes uma vez por dia, no final do mês (dia 25).
  if (hora === 2 && diaDoMes === 25) {
    logToSheet("[Master Trigger] Hora de gerar contas recorrentes para o próximo mês.", "INFO");
    try {
      generateRecurringBillsForNextMonth();
    } catch (e) {
      handleError(e, "masterTriggerHandler -> generateRecurringBillsForNextMonth");
    }
  }
  
  // 3. Analisa novas assinaturas uma vez por mês (no dia 2)
  if (hora === 4 && diaDoMes === 2) {
    logToSheet("[Master Trigger] Hora de analisar novas assinaturas.", "INFO");
    try {
      analisarNovasAssinaturas();
    } catch (e) {
      handleError(e, "masterTriggerHandler -> analisarNovasAssinaturas");
    }
  }

  // 4. Analisa o ritmo de gastos do orçamento a meio do mês (dia 15)
  if (hora === 5 && diaDoMes === 15) {
    logToSheet("[Master Trigger] Hora de analisar o ritmo de gastos do orçamento.", "INFO");
    try {
      analisarRitmoDeGastos();
    } catch (e) {
      handleError(e, "masterTriggerHandler -> analisarRitmoDeGastos");
    }
  }

  // 5. Arquiva transações antigas uma vez por mês (no primeiro dia do mês, de madrugada)
  if (hora === 3 && diaDoMes === 1) {
    logToSheet("[Master Trigger] Hora de arquivar transações antigas.", "INFO");
    try {
      arquivarTransacoesAntigas();
    } catch (e) {
      handleError(e, "masterTriggerHandler -> arquivarTransacoesAntigas");
    }
  }


  // 6. Regista o valor diário do portfólio de investimentos (ex: no final do dia).
  if (hora === 23) {
    logToSheet("[Master Trigger] Hora de registar o valor do portfólio de investimentos.", "INFO");
    try {
      logDailyPortfolioValue();
    } catch (e) {
      handleError(e, "masterTriggerHandler -> logDailyPortfolioValue");
    }
  }
}

/**
 * NOVO: Analisa o ritmo de gastos do mês atual e envia um alerta se a projeção
 * ultrapassar o valor orçamentado para alguma categoria.
 */
function analisarRitmoDeGastos() {
  const adminChatId = getAdminChatIdFromProperties();
  if (!adminChatId) return;

  logToSheet("[AnalisarRitmoGastos] Iniciando análise de ritmo de gastos.", "INFO");
  const hoje = new Date();
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();
  const diaAtual = hoje.getDate();
  const diasNoMes = new Date(anoAtual, mesAtual + 1, 0).getDate();

  const orcamentoDoMes = getBudgetProgressForTelegram(mesAtual + 1, anoAtual);
  if (!orcamentoDoMes || orcamentoDoMes.length === 0) {
    logToSheet("[AnalisarRitmoGastos] Nenhum orçamento encontrado para o mês atual.", "INFO");
    return;
  }

  let alertas = [];

  orcamentoDoMes.forEach(item => {
    if (item.gasto > 0) {
      const gastoProjetado = (item.gasto / diaAtual) * diasNoMes;
      // Alerta se a projeção for pelo menos 10% superior ao orçamentado
      if (gastoProjetado > item.orcado * 1.10) { 
        alertas.push({
          categoria: item.categoria,
          gastoAtual: item.gasto,
          orcado: item.orcado,
          projetado: gastoProjetado
        });
      }
    }
  });

  if (alertas.length > 0) {
    const usuario = getUsuarioPorChatId(adminChatId, getSheetDataWithCache(SHEET_CONFIGURACOES, CACHE_KEY_CONFIG));
    const nomeCurto = usuario.split(' ')[0];
    let mensagem = `Olá, ${escapeMarkdown(nomeCurto)}! Zaq aqui com uma dica sobre o seu orçamento. 👀\n\n` +
                   `Estamos a meio do mês e, se continuar no ritmo atual, os seus gastos em algumas categorias podem ultrapassar o que foi planeado:\n\n`;
    
    alertas.forEach(alerta => {
      mensagem += `*${escapeMarkdown(alerta.categoria)}:*\n` +
                  ` • Orçado: ${formatCurrency(alerta.orcado)}\n` +
                  ` • Gasto até agora: ${formatCurrency(alerta.gasto)}\n` +
                  ` • Projeção para o final do mês: *${formatCurrency(alerta.projetado)}* forecasted\n\n`;
    });
    
    mensagem += `Ainda há tempo para ajustar! Manter o foco agora pode fazer toda a diferença no final do mês. 😉`;
    
    enviarMensagemTelegram(adminChatId, mensagem);
    logToSheet(`[AnalisarRitmoGastos] Alerta de "furo" no orçamento enviado para ${adminChatId}.`, "INFO");
  } else {
    logToSheet("[AnalisarRitmoGastos] Ritmo de gastos dentro do esperado. Nenhum alerta enviado.", "INFO");
  }
}

/**
 * NOVO: Analisa os gastos dos últimos 60 dias para detetar novas despesas recorrentes (assinaturas).
 */
function analisarNovasAssinaturas() {
  const adminChatId = getAdminChatIdFromProperties();
  if (!adminChatId) return;

  logToSheet("[AnalisarAssinaturas] Iniciando análise de novas despesas recorrentes.", "INFO");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
  if (!transacoesSheet) {
    logToSheet("[AnalisarAssinaturas] Aba 'Transacoes' não encontrada.", "ERROR");
    return;
  }
  
  const dadosTransacoes = transacoesSheet.getDataRange().getValues();
  const hoje = new Date();
  const inicioPeriodoRecente = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - 60); // Últimos 60 dias

  const gastos = {};

  for (let i = 1; i < dadosTransacoes.length; i++) {
    const linha = dadosTransacoes[i];
    const dataTransacao = parseData(linha[0]);
    
    if (dataTransacao && dataTransacao >= inicioPeriodoRecente) {
      const tipo = (linha[4] || "").toLowerCase();
      if (tipo === "despesa") {
        const descricao = normalizarTexto(linha[1]);
        const mes = dataTransacao.getMonth();

        if (!gastos[descricao]) {
          gastos[descricao] = { count: 0, months: new Set(), originalDesc: linha[1] };
        }
        gastos[descricao].count++;
        gastos[descricao].months.add(mes);
      }
    }
  }

  const novasAssinaturas = [];
  for (const desc in gastos) {
    // Considera uma nova assinatura se apareceu pelo menos 2 vezes nos últimos 2 meses
    if (gastos[desc].months.size >= 2) {
      novasAssinaturas.push(gastos[desc].originalDesc);
    }
  }

  if (novasAssinaturas.length > 0) {
    const usuario = getUsuarioPorChatId(adminChatId, getSheetDataWithCache(SHEET_CONFIGURACOES, CACHE_KEY_CONFIG));
    const nomeCurto = usuario.split(' ')[0];
    let mensagem = `Olá, ${escapeMarkdown(nomeCurto)}! Sou eu, o Zaq. 👋\n\n` +
                   `Fiz uma análise dos seus gastos e reparei em algumas despesas que se repetiram nos últimos meses:\n\n`;
    
    novasAssinaturas.forEach(desc => {
      mensagem += `• _${escapeMarkdown(desc)}_\n`;
    });
    
    mensagem += `\nIsto podem ser novas subscrições. Está tudo correto, ou é uma oportunidade para economizar? Fica a dica! 😉`;
    
    enviarMensagemTelegram(adminChatId, mensagem);
    logToSheet(`[AnalisarAssinaturas] Sugestão de economia sobre novas assinaturas enviada para ${adminChatId}.`, "INFO");
  } else {
    logToSheet("[AnalisarAssinaturas] Nenhuma nova despesa recorrente detetada.", "INFO");
  }
}



/**
 * Move transações com mais de 2 anos da aba principal para uma aba de arquivo.
 * Esta função de manutenção ajuda a manter a performance do sistema.
 */
function arquivarTransacoesAntigas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
  if (!transacoesSheet) {
    logToSheet("Aba 'Transacoes' não encontrada para arquivamento.", "WARN");
    return;
  }

  // Define o ponto de corte: transações com mais de 2 anos serão arquivadas.
  const cutoffDate = new Date();
  cutoffDate.setFullYear(cutoffDate.getFullYear() - 2);
  logToSheet(`[Arquivamento] Iniciando arquivamento de transações anteriores a ${cutoffDate.toLocaleDateString()}.`, "INFO");

  let arquivoSheet = ss.getSheetByName(SHEET_TRANSACOES_ARQUIVO);
  if (!arquivoSheet) {
    arquivoSheet = ss.insertSheet(SHEET_TRANSACOES_ARQUIVO);
    const headers = transacoesSheet.getRange(1, 1, 1, transacoesSheet.getLastColumn()).getValues();
    arquivoSheet.appendRow(headers[0]);
    logToSheet(`Aba '${SHEET_TRANSACOES_ARQUIVO}' criada.`, "INFO");
  }

  const data = transacoesSheet.getDataRange().getValues();
  const headers = data[0];
  const colDate = headers.indexOf('Data');

  if (colDate === -1) {
    logToSheet("Coluna 'Data' não encontrada na aba 'Transacoes'. Arquivamento cancelado.", "ERROR");
    return;
  }

  const rowsToArchive = [];
  const rowIndexesToDelete = [];

  // Itera a partir da segunda linha para ignorar o cabeçalho
  for (let i = 1; i < data.length; i++) {
    const transactionDate = parseData(data[i][colDate]);
    if (transactionDate && transactionDate < cutoffDate) {
      rowsToArchive.push(data[i]);
      rowIndexesToDelete.push(i + 1); // Guarda o índice baseado em 1
    }
  }

  if (rowsToArchive.length > 0) {
    // Adiciona as linhas arquivadas à aba de arquivo
    arquivoSheet.getRange(arquivoSheet.getLastRow() + 1, 1, rowsToArchive.length, headers.length).setValues(rowsToArchive);

    // Exclui as linhas da aba principal, de trás para a frente para evitar problemas de índice
    for (let i = rowIndexesToDelete.length - 1; i >= 0; i--) {
      transacoesSheet.deleteRow(rowIndexesToDelete[i]);
    }

    logToSheet(`[Arquivamento] ${rowsToArchive.length} transações foram arquivadas com sucesso.`, "INFO");
    const adminChatId = getAdminChatIdFromProperties();
    if (adminChatId) {
      enviarMensagemTelegram(adminChatId, `✅ Manutenção Automática: ${rowsToArchive.length} transações antigas foram arquivadas para manter o sistema rápido.`);
    }
  } else {
    logToSheet("[Arquivamento] Nenhuma transação antiga encontrada para arquivar.", "INFO");
  }
}


const SHEET_PORTFOLIO_HISTORY = "PortfolioHistory";

/**
 * Calcula o valor total atual da carteira de investimentos e regista-o numa aba de histórico.
 * Deve ser executada por um acionador diário para criar o histórico para o gráfico de evolução.
 */
function logDailyPortfolioValue() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let historySheet = ss.getSheetByName(SHEET_PORTFOLIO_HISTORY);

    // Se a aba de histórico não existir, cria-a
    if (!historySheet) {
      historySheet = ss.insertSheet(SHEET_PORTFOLIO_HISTORY);
      historySheet.appendRow(["Data", "ValorTotal"]);
      historySheet.hideSheet(); // Esconde a aba do utilizador
      logToSheet(`Aba de histórico '${SHEET_PORTFOLIO_HISTORY}' criada.`, "INFO");
    }

    // Calcula o valor total atual da carteira
    const totalValue = getTotalInvestmentsValue(); // Função de Investimentos.gs

    if (totalValue > 0) {
      const today = new Date();
      historySheet.appendRow([today, totalValue]);
      logToSheet(`Valor do portfólio (R$ ${totalValue}) registado para ${today.toLocaleDateString()}.`, "INFO");
    } else {
      logToSheet("Valor do portfólio é zero. Nenhum registo de histórico adicionado.", "INFO");
    }
  } catch (e) {
    handleError(e, "logDailyPortfolioValue");
  }
}

