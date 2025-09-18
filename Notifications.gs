/**
 * @file Notifications.gs
 * @description Este arquivo contém funções para gerar e enviar notificações proativas via Telegram.
 * Inclui alertas de orçamento, lembretes de contas a pagar e resumos de gastos.
 */

/**
 * Função principal para verificar e enviar todas as notificações configuradas.
 * Esta função será chamada por um gatilho de tempo.
 */
function checkAndSendNotifications() {
  logToSheet("Iniciando verificação e envio de notificações proativas.", "INFO");
  const notificationConfig = getNotificationConfig(); 

  if (!notificationConfig) {
    logToSheet("Configurações de notificações não encontradas. Nenhuma notificação será enviada.", "WARN");
    return;
  }

  for (const chatId in notificationConfig) {
    const userConfig = notificationConfig[chatId];
    logToSheet(`Verificando notificações para Chat ID: ${chatId} (Usuário: ${userConfig.usuario})`, "DEBUG");

    if (userConfig.enableBudgetAlerts) sendBudgetAlerts(chatId, userConfig.usuario);
    if (userConfig.enableBillReminders) sendUpcomingBillReminders(chatId, userConfig.usuario);
    if (userConfig.enableDailySummary && isTimeForDailySummary(userConfig.dailySummaryTime)) sendDailySummary(chatId, userConfig.usuario);
    if (userConfig.enableWeeklySummary && isTimeForWeeklySummary(userConfig.weeklySummaryDay, userConfig.weeklySummaryTime)) {
      generateAndSendWeeklyInsight(chatId, userConfig.usuario);
    }
    // ### INÍCIO DA ATUALIZAÇÃO ###
    // Adiciona a verificação de alertas de fatura, se estiver ativada na configuração do utilizador
    if (userConfig.alertasDeFatura) {
      sendCreditCardBillAlerts(chatId, userConfig.usuario);
    }
    // ### FIM DA ATUALIZAÇÃO ###
    // VERIFICAÇÃO DE METAS ATINGIDAS ACONTECE AQUI
    verificarMetasAtingidas(chatId, userConfig.usuario);
  }
  logToSheet("Verificação e envio de notificações concluídos.", "INFO");
}

/**
 * NOVO: Verifica se alguma meta de poupança foi atingida e notifica o utilizador.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do utilizador.
 */
function verificarMetasAtingidas(chatId, usuario) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const metasSheet = ss.getSheetByName(SHEET_METAS);
  const alertasSheet = ss.getSheetByName(SHEET_ALERTAS_ENVIADOS);
  if (!metasSheet || !alertasSheet) return;

  const dadosMetas = metasSheet.getDataRange().getValues();
  const alertasEnviados = alertasSheet.getDataRange().getValues();
  const headers = dadosMetas[0];
  const colMap = getColumnMap(headers);

  for (let i = 1; i < dadosMetas.length; i++) {
    const linha = i + 1;
    const nomeMeta = dadosMetas[i][colMap['Nome da Meta']];
    const valorObjetivo = parseBrazilianFloat(String(dadosMetas[i][colMap['Valor Objetivo']] || '0'));
    const valorSalvo = parseBrazilianFloat(String(dadosMetas[i][colMap['Valor Salvo']] || '0'));
    const status = (dadosMetas[i][colMap['Status']] || "").toLowerCase();

    if (valorSalvo >= valorObjetivo && status === 'em andamento') {
      const alertKey = `${chatId}|meta_atingida|${nomeMeta}`;
      if (!_jaEnviado(alertKey, alertasEnviados)) {
        const mensagem = `🎉 Parabéns, ${escapeMarkdown(usuario.split(' ')[0])}!! 🎉\n\n` +
                         `Você atingiu a sua meta de *${escapeMarkdown(nomeMeta)}*!\n\n` +
                         `Objetivo: ${formatCurrency(valorObjetivo)}\n` +
                         `Salvo: ${formatCurrency(valorSalvo)}\n\n` +
                         `Estou a celebrar esta sua grande conquista consigo! 🥳`;
        
        enviarMensagemTelegram(chatId, mensagem);
        
        // Atualiza o status na planilha e regista o alerta
        metasSheet.getRange(linha, colMap['Status'] + 1).setValue('Atingida');
        alertasSheet.appendRow([new Date(), chatId, alertKey]);
        logToSheet(`[Metas] Alerta de meta atingida enviado para '${nomeMeta}'.`, "INFO");
      }
    }
  }
}

/**
 * **FUNÇÃO TOTALMENTE REESTRUTURADA**
 * Gera múltiplos insights sobre os gastos da última semana e envia os mais relevantes.
 * @param {string} chatId O ID do chat para enviar o insight.
 * @param {string} usuario O nome do utilizador para filtrar as transações.
 */
function generateAndSendWeeklyInsight(chatId, usuario) {
  logToSheet(`[Insight Semanal] Iniciando geração para ${usuario} (${chatId}).`, "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
  if (!transacoesSheet) {
    logToSheet("[Insight Semanal] Aba 'Transacoes' não encontrada.", "ERROR");
    return;
  }
  const transacoes = transacoesSheet.getDataRange().getValues();

  // 1. Gera uma lista de possíveis insights
  const potentialInsights = [];
  potentialInsights.push(_findTopSpendingCategory(usuario, transacoes));
  potentialInsights.push(_findBiggestSpendingIncrease(usuario, transacoes));
  potentialInsights.push(_detectNewRecurringExpenses(usuario, transacoes));

  // 2. Filtra os insights que foram gerados com sucesso (não nulos) e ordena por prioridade
  const validInsights = potentialInsights.filter(Boolean).sort((a, b) => (a.priority || 99) - (b.priority || 99));

  if (validInsights.length === 0) {
    logToSheet(`[Insight Semanal] Nenhum insight relevante gerado para ${usuario}.`, "INFO");
    return;
  }

  // 3. Constrói a mensagem final com os insights mais importantes
  const nomeFormatado = escapeMarkdown(usuario.split(' ')[0]);
  let mensagem = `💡 *Seu Insight Semanal do Gasto Certo*\n\n` +
                 `Olá, ${nomeFormatado}! Aqui está a sua análise da semana que passou:\n\n`;

  // Adiciona o primeiro (e mais importante) insight
  mensagem += `*${validInsights[0].title}*\n${validInsights[0].text}\n\n`;

  // Se houver um segundo insight relevante, adiciona-o também
  if (validInsights.length > 1) {
    mensagem += `*${validInsights[1].title}*\n${validInsights[1].text}\n\n`;
  }
  
  mensagem += `_Continue a registar para receber mais insights!_ 🚀`;

  enviarMensagemTelegram(chatId, mensagem);
  logToSheet(`[Insight Semanal] Mensagem com ${validInsights.length} insights enviada para ${usuario}.`, "INFO");
}

/**
 * Verifica se é hora de enviar o resumo diário com base na hora configurada.
 * @param {string} timeString A hora configurada no formato "HH:mm".
 * @returns {boolean} True se for a hora de enviar, false caso contrário.
 */
function isTimeForDailySummary(timeString) {
  if (!timeString) return false;
  const now = new Date();
  const [configHour, configMinute] = timeString.split(':').map(Number);
  
  const currentHour = now.getHours();
  const currentMinute = now.getMinutes();

  return currentHour === configHour && currentMinute >= configMinute && currentMinute < configMinute + 5; // 5 minutos de janela
}

/**
 * Verifica se é hora de enviar o resumo semanal com base no dia da semana e hora configurados.
 * @param {number} dayOfWeek O dia da semana configurado (0=Domingo, 6=Sábado).
 * @param {string} timeString A hora configurada no formato "HH:mm".
 * @returns {boolean} True se for a hora de enviar, false caso contrário.
 */
function isTimeForWeeklySummary(dayOfWeek, timeString) {
  if (dayOfWeek === null || dayOfWeek === undefined || !timeString) return false;
  const now = new Date();
  const [configHour, configMinute] = timeString.split(':').map(Number);

  const currentDay = now.getDay();
  const currentHour = now.getHours();
  const currentMinute = now.getMinutes();

  return currentDay === dayOfWeek && currentHour === configHour && currentMinute >= configMinute && currentMinute < configMinute + 5;
}

/**
 * Verifica e envia alertas sobre fechamento e vencimento de faturas de cartão de crédito.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 */
function sendCreditCardBillAlerts(chatId, usuario) {
  try {
    logToSheet(`[AlertasCartao] Iniciando verificação para ${usuario} (${chatId})`, "INFO");
    atualizarSaldosDasContas();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contasSheet = ss.getSheetByName(SHEET_CONTAS);
    const alertasSheet = ss.getSheetByName(SHEET_ALERTAS_ENVIADOS);
    if (!contasSheet || !alertasSheet) {
      logToSheet("[AlertasCartao] Aba 'Contas' ou 'AlertasEnviados' não encontrada.", "ERROR");
      return;
    }
    
    const dadosContas = contasSheet.getDataRange().getValues();
    const alertasEnviados = alertasSheet.getDataRange().getValues();
    const today = new Date();
    const hoje = today.getDate();
    const mesAtual = today.getMonth();
    const anoAtual = today.getFullYear();

    for (let i = 1; i < dadosContas.length; i++) {
      const infoConta = obterInformacoesDaConta(dadosContas[i][0], dadosContas);
      
      if (infoConta && infoConta.tipo === 'cartão de crédito') {
        const faturaAtual = globalThis.saldosCalculados[infoConta.nomeNormalizado]?.faturaAtual || 0;

        // 1. Alerta de Fechamento da Fatura
        if (infoConta.diaFechamento === hoje) {
          const alertKey = `${chatId}|${infoConta.nomeNormalizado}|fechamento|${anoAtual}-${mesAtual}`;
          if (faturaAtual > 0 && !_jaEnviado(alertKey, alertasEnviados)) {
            const mensagem = `💳 *Alerta de Fatura Fechada*\n\nA fatura do seu cartão *${escapeMarkdown(infoConta.nomeOriginal)}* fechou hoje com o valor de *${formatCurrency(faturaAtual)}*.\n\nO vencimento é no dia *${infoConta.vencimento}*.`;
            enviarMensagemTelegram(chatId, mensagem);
            alertasSheet.appendRow([new Date(), chatId, alertKey]);
          }
        }

        // 2. Lembrete de Vencimento da Fatura
        const dataVencimento = new Date(anoAtual, mesAtual, infoConta.vencimento);
        const diffTime = dataVencimento.getTime() - today.getTime();
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

        if (diffDays >= 0 && diffDays <= BILL_REMINDER_DAYS_BEFORE) {
           const alertKey = `${chatId}|${infoConta.nomeNormalizado}|vencimento|${anoAtual}-${mesAtual}`;
           if (faturaAtual > 0 && !_jaEnviado(alertKey, alertasEnviados)) {
             const mensagem = `🔔 *Lembrete de Vencimento*\n\nA fatura do seu cartão *${escapeMarkdown(infoConta.nomeOriginal)}* no valor de *${formatCurrency(faturaAtual)}* vence em *${diffDays} dia(s)*.`;
             enviarMensagemTelegram(chatId, mensagem);
             alertasSheet.appendRow([new Date(), chatId, alertKey]);
           }
        }
      }
    }
  } catch (e) {
    handleError(e, `sendCreditCardBillAlerts para ${usuario}`, chatId);
  }
}

/**
 * @private
 * Verifica se um alerta com uma chave específica já foi enviado.
 * @param {string} key A chave única do alerta.
 * @param {Array<Array<any>>} sentAlertsData Os dados da aba 'AlertasEnviados'.
 * @returns {boolean} True se o alerta já foi enviado, false caso contrário.
 */
function _jaEnviado(key, sentAlertsData) {
  for (let i = 1; i < sentAlertsData.length; i++) {
    if (sentAlertsData[i][2] === key) {
      return true;
    }
  }
  return false;
}



/**
 * Envia alertas de orçamento excedido para o usuário.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 */
function sendBudgetAlerts(chatId, usuario) {
  logToSheet(`Verificando alertas de orçamento para ${usuario} (${chatId}).`, "INFO");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orcamentoSheet = ss.getSheetByName(SHEET_ORCAMENTO);
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);

  if (!orcamentoSheet || !transacoesSheet) {
    logToSheet("Aba 'Orcamento' ou 'Transacoes' não encontrada para alertas de orçamento.", "ERROR");
    return;
  }

  const orcamentoData = orcamentoSheet.getDataRange().getValues();
  const transacoesData = transacoesSheet.getDataRange().getValues();

  const today = new Date();
  const currentMonth = today.getMonth() + 1; // 1-indexed
  const currentYear = today.getFullYear();

  const userBudgets = {};
  const userSpendings = {};

  // Coleta orçamentos por categoria/subcategoria para o usuário no mês/ano atual
  for (let i = 1; i < orcamentoData.length; i++) {
    const row = orcamentoData[i];
    const orcamentoUsuario = (row[0] || "").toString().trim();
    const orcamentoAno = parseInt(row[1]);
    const orcamentoMes = parseInt(row[2]);
    const categoria = (row[3] || "").toString().trim();
    const subcategoria = (row[4] || "").toString().trim();
    // NOVO: Usar parseBrazilianFloat
    const valorOrcado = parseBrazilianFloat(String(row[5]));

    if (normalizarTexto(orcamentoUsuario) === normalizarTexto(usuario) &&
        orcamentoAno === currentYear && orcamentoMes === currentMonth &&
        valorOrcado > 0) {
      const key = `${categoria}>${subcategoria}`;
      userBudgets[key] = valorOrcado;
      userSpendings[key] = 0; // Inicializa gasto para esta categoria
    }
  }

  // Calcula gastos para as categorias orçadas do usuário no mês/ano atual
  for (let i = 1; i < transacoesData.length; i++) {
    const row = transacoesData[i];
    const dataTransacao = parseData(row[0]);
    const tipoTransacao = (row[4] || "").toString().trim();
    // NOVO: Usar parseBrazilianFloat
    const valorTransacao = parseBrazilianFloat(String(row[5]));
    const categoriaTransacao = (row[2] || "").toString().trim();
    const subcategoriaTransacao = (row[3] || "").toString().trim();
    const usuarioTransacao = (row[11] || "").toString().trim();

    if (dataTransacao && dataTransacao.getMonth() + 1 === currentMonth &&
        dataTransacao.getFullYear() === currentYear &&
        normalizarTexto(usuarioTransacao) === normalizarTexto(usuario) &&
        tipoTransacao === "Despesa") {
      const key = `${categoriaTransacao}>${subcategoriaTransacao}`;
      if (userSpendings.hasOwnProperty(key)) {
        userSpendings[key] += valorTransacao;
      }
    }
  }

  let alertsSent = false;
  let alertMessage = `⚠️ *Alerta de Orçamento - ${getNomeMes(currentMonth - 1)}/${currentYear}* ⚠️\n\n`;
  let hasAlerts = false;

  for (const key in userBudgets) {
    const orcado = userBudgets[key];
    const gasto = userSpendings[key];
    const percentage = (gasto / orcado) * 100;

    if (percentage >= BUDGET_ALERT_THRESHOLD_PERCENT) {
      const [categoria, subcategoria] = key.split('>');
      // NOVO: Usar escapeMarkdown
      alertMessage += `*${escapeMarkdown(capitalize(categoria))} > ${escapeMarkdown(capitalize(subcategoria))}*\n`;
      alertMessage += `  Gasto: ${formatCurrency(gasto)} (Orçado: ${formatCurrency(orcado)})\n`;
      alertMessage += `  Progresso: ${percentage.toFixed(1)}% (${percentage >= 100 ? 'EXCEDIDO!' : 'próximo ao limite!'})\n\n`;
      hasAlerts = true;
    }
  }

  if (hasAlerts) {
    enviarMensagemTelegram(chatId, alertMessage);
    logToSheet(`Alerta de orçamento enviado para ${usuario} (${chatId}).`, "INFO");
    alertsSent = true;
  } else {
    logToSheet(`Nenhum alerta de orçamento para ${usuario} (${chatId}).`, "DEBUG");
  }

  return alertsSent;
}

/**
 * Envia lembretes de contas a pagar próximas ao vencimento para o usuário.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 */
function sendUpcomingBillReminders(chatId, usuario) {
  logToSheet(`Verificando lembretes de contas a pagar para ${usuario} (${chatId}).`, "INFO");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);

  if (!contasAPagarSheet) {
    logToSheet("Aba 'Contas_a_Pagar' não encontrada para lembretes.", "ERROR");
    return;
  }

  const contasAPagarData = contasAPagarSheet.getDataRange().getValues();
  const headers = contasAPagarData[0];
  const colStatus = headers.indexOf('Status');
  const colDataVencimento = headers.indexOf('Data de Vencimento');
  const colDescricao = headers.indexOf('Descricao');
  const colValor = headers.indexOf('Valor');

  if (colStatus === -1 || colDataVencimento === -1 || colDescricao === -1 || colValor === -1) {
    logToSheet("Colunas essenciais (Status, Data de Vencimento, Descricao, Valor) não encontradas na aba 'Contas_a_Pagar'.", "ERROR");
    return;
  }

  const today = new Date();
  let remindersSent = false;
  let reminderMessage = `🔔 *Lembrete de Contas a Pagar* 🔔\n\n`;
  let hasReminders = false;

  for (let i = 1; i < contasAPagarData.length; i++) {
    const row = contasAPagarData[i];
    const status = (row[colStatus] || "").toString().trim();
    const dataVencimento = parseData(row[colDataVencimento]);
    const descricao = (row[colDescricao] || "").toString().trim();
    // NOVO: Usar parseBrazilianFloat
    const valor = parseBrazilianFloat(String(row[colValor]));

    if (status.toLowerCase() === "pendente" && dataVencimento) {
      const diffTime = dataVencimento.getTime() - today.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      if (diffDays >= 0 && diffDays <= BILL_REMINDER_DAYS_BEFORE) {
        // NOVO: Usar escapeMarkdown
        reminderMessage += `*${escapeMarkdown(capitalize(descricao))}*\n`;
        reminderMessage += `  Valor: ${formatCurrency(valor)}\n`;
        reminderMessage += `  Vencimento: ${Utilities.formatDate(dataVencimento, Session.getScriptTimeZone(), "dd/MM/yyyy")}\n`;
        reminderMessage += `  Faltam: ${diffDays} dias\n\n`;
        hasReminders = true;
      }
    }
  }

  if (hasReminders) {
    enviarMensagemTelegram(chatId, reminderMessage);
    logToSheet(`Lembrete de contas a pagar enviado para ${usuario} (${chatId}).`, "INFO");
    remindersSent = true;
  } else {
    logToSheet(`Nenhum lembrete de contas a pagar para ${usuario} (${chatId}).`, "DEBUG");
  }

  return remindersSent;
}

/**
 * Envia um resumo diário de gastos e receitas para o usuário.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 */
function sendDailySummary(chatId, usuario) {
  logToSheet(`Gerando resumo diário para ${usuario} (${chatId}).`, "INFO");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);

  if (!transacoesSheet) {
    logToSheet("Aba 'Transacoes' não encontrada para resumo diário.", "ERROR");
    return;
  }

  const transacoesData = transacoesSheet.getDataRange().getValues();
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1); // Resumo é para o dia anterior

  let dailyReceitas = 0;
  let dailyDespesas = 0;

  for (let i = 1; i < transacoesData.length; i++) {
    const row = transacoesData[i];
    const dataTransacao = parseData(row[0]);
    const tipoTransacao = (row[4] || "").toString().trim();
    // NOVO: Usar parseBrazilianFloat
    const valorTransacao = parseBrazilianFloat(String(row[5]));
    const usuarioTransacao = (row[11] || "").toString().trim();

    if (dataTransacao && dataTransacao.getDate() === yesterday.getDate() &&
        dataTransacao.getMonth() === yesterday.getMonth() &&
        dataTransacao.getFullYear() === yesterday.getFullYear() &&
        normalizarTexto(usuarioTransacao) === normalizarTexto(usuario)) {
      if (tipoTransacao === "Receita") {
        dailyReceitas += valorTransacao;
      } else if (tipoTransacao === "Despesa") {
        dailyDespesas += valorTransacao;
      }
    }
  }

  let summaryMessage = `📊 *Resumo Diário - ${Utilities.formatDate(yesterday, Session.getScriptTimeZone(), "dd/MM/yyyy")}* 📊\n\n`;
  summaryMessage += `💰 Receitas: ${formatCurrency(dailyReceitas)}\n`;
  summaryMessage += `💸 Despesas: ${formatCurrency(dailyDespesas)}\n`;
  summaryMessage += `✨ Saldo do Dia: ${formatCurrency(dailyReceitas - dailyDespesas)}\n\n`;
  summaryMessage += "Mantenha o controle! 💪";

  enviarMensagemTelegram(chatId, summaryMessage);
  logToSheet(`Resumo diário enviado para ${usuario} (${chatId}).`, "INFO");
}

/**
 * Envia um resumo semanal de gastos e receitas para o usuário.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 */
function sendWeeklySummary(chatId, usuario) {
  logToSheet(`Gerando resumo semanal para ${usuario} (${chatId}).`, "INFO");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);

  if (!transacoesSheet) {
    logToSheet("Aba 'Transacoes' não encontrada para resumo semanal.", "ERROR");
    return;
  }

  const transacoesData = transacoesSheet.getDataRange().getValues();
  const today = new Date();
  const startOfWeek = new Date(today);
  startOfWeek.setDate(today.getDate() - today.getDay()); // Início da semana (Domingo)
  startOfWeek.setHours(0, 0, 0, 0);

  const endOfWeek = new Date(startOfWeek);
  endOfWeek.setDate(startOfWeek.getDate() + 6); // Fim da semana (Sábado)
  endOfWeek.setHours(23, 59, 59, 999);

  let weeklyReceitas = 0;
  let weeklyDespesas = 0;
  const expensesByCategory = {};

  for (let i = 1; i < transacoesData.length; i++) {
    const row = transacoesData[i];
    const dataTransacao = parseData(row[0]);
    const tipoTransacao = (row[4] || "").toString().trim();
    // NOVO: Usar parseBrazilianFloat
    const valorTransacao = parseBrazilianFloat(String(row[5]));
    const categoriaTransacao = (row[2] || "").toString().trim();
    const usuarioTransacao = (row[11] || "").toString().trim();

    if (dataTransacao && dataTransacao >= startOfWeek && dataTransacao <= endOfWeek &&
        normalizarTexto(usuarioTransacao) === normalizarTexto(usuario)) {
      if (tipoTransacao === "Receita") {
        weeklyReceitas += valorTransacao;
      } else if (tipoTransacao === "Despesa") {
        weeklyDespesas += valorTransacao;
        expensesByCategory[categoriaTransacao] = (expensesByCategory[categoriaTransacao] || 0) + valorTransacao;
      }
    }
  }

  let summaryMessage = `📈 *Resumo Semanal - ${Utilities.formatDate(startOfWeek, Session.getScriptTimeZone(), "dd/MM/yyyy")} a ${Utilities.formatDate(endOfWeek, Session.getScriptTimeZone(), "dd/MM/yyyy")}* 📉\n\n`;
  summaryMessage += `💰 Receitas: ${formatCurrency(weeklyReceitas)}\n`;
  summaryMessage += `💸 Despesas: ${formatCurrency(weeklyDespesas)}\n`;
  summaryMessage += `✨ Saldo da Semana: ${formatCurrency(weeklyReceitas - weeklyDespesas)}\n\n`;

  summaryMessage += "*Principais Despesas por Categoria:*\n";
  const sortedExpenses = Object.entries(expensesByCategory).sort(([, a], [, b]) => b - a);
  if (sortedExpenses.length > 0) {
    sortedExpenses.slice(0, 5).forEach(([category, amount]) => { // Top 5 categorias
      // NOVO: Usar escapeMarkdown
      summaryMessage += `  • ${escapeMarkdown(capitalize(category))}: ${formatCurrency(amount)}\n`;
    });
  } else {
    summaryMessage += "  _Nenhuma despesa registrada nesta semana._\n";
  }
  summaryMessage += "\nContinue acompanhando suas finanças! 🚀";

  enviarMensagemTelegram(chatId, summaryMessage);
  logToSheet(`Resumo semanal enviado para ${usuario} (${chatId}).`, "INFO");
}

/**
 * Obtém as configurações de notificação da aba 'Notificacoes_Config'.
 * @returns {Object} Um objeto onde a chave é o Chat ID e o valor são as configurações do usuário.
 */
function getNotificationConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(SHEET_NOTIFICACOES_CONFIG); 

  if (!configSheet) {
    logToSheet("Aba 'Notificacoes_Config' não encontrada. Nenhuma configuracao de notificacao sera lida.", "ERROR");
    return null;
  }

  const data = configSheet.getDataRange().getValues();
  const headers = data[0];

  const colChatId = headers.indexOf('Chat ID');
  const colUsuario = headers.indexOf('Usuário');
  const colEnableBudgetAlerts = headers.indexOf('Alertas Orçamento');
  const colEnableBillReminders = headers.indexOf('Lembretes Contas a Pagar');
  const colEnableDailySummary = headers.indexOf('Resumo Diário');
  const colDailySummaryTime = headers.indexOf('Hora Resumo Diário (HH:mm)');
  const colEnableWeeklySummary = headers.indexOf('Resumo Semanal');
  const colWeeklySummaryDay = headers.indexOf('Dia Resumo Semanal (0-6)');
  const colWeeklySummaryTime = headers.indexOf('Hora Resumo Semanal (HH:mm)');
  // ### INÍCIO DA ATUALIZAÇÃO ###
  const colAlertasFatura = headers.indexOf('Alertas de Fatura');
  // ### FIM DA ATUALIZAÇÃO ###


  if ([colChatId, colUsuario].some(idx => idx === -1)) {
    logToSheet("Colunas essenciais ('Chat ID', 'Usuário') para 'Notificacoes_Config' ausentes.", "ERROR");
    return null;
  }

  const notificationConfig = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const chatId = (row[colChatId] || "").toString().trim();
    if (chatId) {
      notificationConfig[chatId] = {
        usuario: (row[colUsuario] || "").toString().trim(),
        enableBudgetAlerts: colEnableBudgetAlerts > -1 && (row[colEnableBudgetAlerts] || "").toString().toLowerCase().trim() === 'sim',
        enableBillReminders: colEnableBillReminders > -1 && (row[colEnableBillReminders] || "").toString().toLowerCase().trim() === 'sim',
        enableDailySummary: colEnableDailySummary > -1 && (row[colEnableDailySummary] || "").toString().toLowerCase().trim() === 'sim',
        dailySummaryTime: colDailySummaryTime > -1 ? (row[colDailySummaryTime] || "").toString().trim() : '',
        enableWeeklySummary: colEnableWeeklySummary > -1 && (row[colEnableWeeklySummary] || "").toString().toLowerCase().trim() === 'sim',
        weeklySummaryDay: colWeeklySummaryDay > -1 ? parseInt(row[colWeeklySummaryDay]) : 1, // Padrão é Segunda-feira
        weeklySummaryTime: colWeeklySummaryTime > -1 ? (row[colWeeklySummaryTime] || "").toString().trim() : '',
        // ### INÍCIO DA ATUALIZAÇÃO ###
        // Lê a nova coluna. Se não existir, o valor será `false` por defeito.
        alertasDeFatura: colAlertasFatura > -1 && (row[colAlertasFatura] || "").toString().toLowerCase().trim() === 'sim'
        // ### FIM DA ATUALIZAÇÃO ###
      };
    }
  }
  return notificationConfig;
}




/**
 * Função auxiliar que gera e envia o insight para um único usuário.
 * @param {object} usuario Objeto com {chatId, nome}.
 * @param {Array<Array<any>>} transacoes Todos os dados da aba de transações.
 */
function gerarEEnviarInsightParaUsuario(usuario, transacoes) {
  const { chatId, nome } = usuario;

  // 1. Definir o período da semana passada (Domingo a Sábado)
  const hoje = new Date();
  const diaDaSemana = hoje.getDay(); // 0=Domingo, 6=Sábado
  const fimDaSemana = new Date(hoje);
  fimDaSemana.setDate(hoje.getDate() - diaDaSemana - 1); // Fim no último sábado
  fimDaSemana.setHours(23, 59, 59, 999);
  
  const inicioDaSemana = new Date(fimDaSemana);
  inicioDaSemana.setDate(fimDaSemana.getDate() - 6); // Início no último domingo
  inicioDaSemana.setHours(0, 0, 0, 0);

  // 2. Calcular gastos da semana por categoria
  const gastosSemana = {};
  let totalGastoSemana = 0;

  for (let i = 1; i < transacoes.length; i++) {
    const linha = transacoes[i];
    const dataTransacao = parseData(linha[0]);
    
    if (dataTransacao >= inicioDaSemana && dataTransacao <= fimDaSemana) {
      const tipo = (linha[4] || "").toLowerCase();
      const categoria = linha[2];
      const valor = parseBrazilianFloat(String(linha[5]));

      if (tipo === "despesa" && categoria && categoria.trim() !== "🔄 Transferências") {
        gastosSemana[categoria] = (gastosSemana[categoria] || 0) + valor;
        totalGastoSemana += valor;
      }
    }
  }

  if (totalGastoSemana === 0) {
    logToSheet(`Nenhum gasto encontrado na última semana para ${nome}. Insight não enviado.`, "INFO");
    return;
  }

  // 3. Encontrar a categoria com o maior gasto absoluto
  const categoriaMaiorGasto = Object.keys(gastosSemana).reduce((a, b) => gastosSemana[a] > gastosSemana[b] ? a : b);
  const valorMaiorGasto = gastosSemana[categoriaMaiorGasto];

  // 4. Calcular a média histórica e a variação para TODAS as categorias da semana
  const inicioHistorico = new Date(inicioDaSemana);
  inicioHistorico.setDate(inicioDaSemana.getDate() - (8 * 7)); // 8 semanas atrás
  
  const analisesCategorias = {};

  for (const categoriaDaSemana in gastosSemana) {
    let gastoHistorico = 0;
    let semanasComGasto = new Set();

    for (let i = 1; i < transacoes.length; i++) {
      const linha = transacoes[i];
      const dataTransacao = parseData(linha[0]);

      if (dataTransacao >= inicioHistorico && dataTransacao < inicioDaSemana) {
        if ((linha[4] || "").toLowerCase() === "despesa" && linha[2] === categoriaDaSemana) {
          gastoHistorico += parseBrazilianFloat(String(linha[5]));
          const semanaDoAno = Utilities.formatDate(dataTransacao, Session.getScriptTimeZone(), "w");
          semanasComGasto.add(semanaDoAno);
        }
      }
    }
    
    const numSemanas = semanasComGasto.size > 0 ? semanasComGasto.size : 1;
    const mediaSemanalHistorica = gastoHistorico / numSemanas;
    
    if (mediaSemanalHistorica > 0) {
      const diferencaPercentual = ((gastosSemana[categoriaDaSemana] - mediaSemanalHistorica) / mediaSemanalHistorica) * 100;
      analisesCategorias[categoriaDaSemana] = {
        percentual: diferencaPercentual,
        media: mediaSemanalHistorica
      };
    }
  }

  // 5. Encontrar a categoria com a maior VARIAÇÃO (aumento)
  let categoriaMaiorVariacao = null;
  let maiorVariacao = -Infinity; // Inicia com valor muito baixo para encontrar a maior variação

  for (const categoria in analisesCategorias) {
    if (analisesCategorias[categoria].percentual > maiorVariacao) {
      maiorVariacao = analisesCategorias[categoria].percentual;
      categoriaMaiorVariacao = categoria;
    }
  }

  // 6. Gerar o insight e formatar a mensagem
  const nomeFormatado = escapeMarkdown(nome.split(' ')[0]);
  let mensagem = `💡 *Seu Insight Semanal do Gasto Certo*\n\n` +
                 `Olá, ${nomeFormatado}! Aqui está a sua análise da semana que passou:\n\n` +
                 `🥇 *Maior Gasto:*\n` +
                 `Sua maior despesa foi com *${escapeMarkdown(categoriaMaiorGasto)}*, totalizando *${formatCurrency(valorMaiorGasto)}*.\n\n`;
  
  let analise = "";
  // Adiciona a análise da MAIOR VARIAÇÃO, se for interessante
  if (categoriaMaiorVariacao && maiorVariacao > 20) { // Limite de 20% para ser considerado um "destaque"
      const media = analisesCategorias[categoriaMaiorVariacao].media;
      analise = `👀 *Destaque da Semana:*\n` +
                `Notamos uma mudança nos seus hábitos! Seus gastos com *${escapeMarkdown(categoriaMaiorVariacao)}* tiveram um aumento de *${maiorVariacao.toFixed(0)}%* em relação à sua média semanal de ${formatCurrency(media)}.`;
  } 
  // Se não houver variação notável, analisa a categoria principal
  else if (analisesCategorias[categoriaMaiorGasto]) {
    const { percentual, media } = analisesCategorias[categoriaMaiorGasto];
    if (percentual > 15) {
      analise = `👀 *Análise do Maior Gasto:*\n` +
                `Este valor é *${percentual.toFixed(0)}% superior* à sua média semanal de ${formatCurrency(media)} para esta categoria.`;
    } else {
      analise = `👀 *Análise do Maior Gasto:*\n` +
                `O seu gasto nesta categoria está *dentro da sua média semanal* de ${formatCurrency(media)}.`;
    }
  }

  if (analise) {
    mensagem += `${analise}\n\n`;
  }

  mensagem += `_Continue a registar para receber mais insights!_`;

  enviarMensagemTelegram(chatId, mensagem);
  logToSheet(`Insight Semanal enviado com sucesso para ${nome} (${chatId}).`, "INFO");
}

// ===================================================================================
// ##      NOVAS FUNÇÕES AUXILIARES PARA GERAR INSIGHTS                            ##
// ===================================================================================

/**
 * @private
 * Encontra a categoria com o maior gasto absoluto na última semana.
 * @returns {object|null} Um objeto de insight ou null.
 */
function _findTopSpendingCategory(usuario, transacoes) {
  const { inicioDaSemana, fimDaSemana } = _getLastWeekDateRange();
  const gastosSemana = {};

  for (let i = 1; i < transacoes.length; i++) {
    const linha = transacoes[i];
    const dataTransacao = parseData(linha[0]);
    if (dataTransacao >= inicioDaSemana && dataTransacao <= fimDaSemana && normalizarTexto(linha[11]) === normalizarTexto(usuario)) {
      const tipo = (linha[4] || "").toLowerCase();
      const categoria = linha[2];
      const valor = parseBrazilianFloat(String(linha[5]));
      if (tipo === "despesa" && categoria && categoria.trim() !== "🔄 Transferências") {
        gastosSemana[categoria] = (gastosSemana[categoria] || 0) + valor;
      }
    }
  }

  if (Object.keys(gastosSemana).length === 0) return null;

  const categoriaMaiorGasto = Object.keys(gastosSemana).reduce((a, b) => gastosSemana[a] > gastosSemana[b] ? a : b);
  const valorMaiorGasto = gastosSemana[categoriaMaiorGasto];

  return {
    priority: 2, // Prioridade média
    title: "🥇 Seu Maior Gasto Semanal",
    text: `A sua maior despesa foi com *${escapeMarkdown(categoriaMaiorGasto)}*, totalizando *${formatCurrency(valorMaiorGasto)}*.`
  };
}

/**
 * @private
 * Encontra a categoria que teve o maior aumento percentual em relação à média histórica.
 * @returns {object|null} Um objeto de insight ou null.
 */
function _findBiggestSpendingIncrease(usuario, transacoes) {
  const { inicioDaSemana, fimDaSemana } = _getLastWeekDateRange();
  const gastosSemana = {};

  // Calcula gastos da última semana por categoria
  for (let i = 1; i < transacoes.length; i++) {
      const linha = transacoes[i];
      const dataTransacao = parseData(linha[0]);
      if (dataTransacao >= inicioDaSemana && dataTransacao <= fimDaSemana && normalizarTexto(linha[11]) === normalizarTexto(usuario) && (linha[4] || "").toLowerCase() === "despesa") {
          const categoria = linha[2];
          const valor = parseBrazilianFloat(String(linha[5]));
          if(categoria && categoria.trim() !== "🔄 Transferências") gastosSemana[categoria] = (gastosSemana[categoria] || 0) + valor;
      }
  }

  const inicioHistorico = new Date(inicioDaSemana);
  inicioHistorico.setDate(inicioDaSemana.getDate() - (8 * 7)); // 8 semanas de histórico

  let categoriaMaiorVariacao = null;
  let maiorVariacao = 25; // Apenas considera variações acima de 25%

  for (const categoria in gastosSemana) {
      let gastoHistorico = 0;
      let semanasComGasto = new Set();
      for (let i = 1; i < transacoes.length; i++) {
          const linha = transacoes[i];
          const dataTransacao = parseData(linha[0]);
          if (dataTransacao >= inicioHistorico && dataTransacao < inicioDaSemana && linha[2] === categoria && normalizarTexto(linha[11]) === normalizarTexto(usuario) && (linha[4] || "").toLowerCase() === "despesa") {
              gastoHistorico += parseBrazilianFloat(String(linha[5]));
              semanasComGasto.add(Utilities.formatDate(dataTransacao, Session.getScriptTimeZone(), "w-YYYY"));
          }
      }

      if(gastoHistorico > 0) {
        const mediaSemanal = gastoHistorico / (semanasComGasto.size || 1);
        const variacao = ((gastosSemana[categoria] - mediaSemanal) / mediaSemanal) * 100;
        if (variacao > maiorVariacao) {
          maiorVariacao = variacao;
          categoriaMaiorVariacao = {
            nome: categoria,
            variacao: variacao,
            gastoAtual: gastosSemana[categoria],
            media: mediaSemanal
          };
        }
      }
  }

  if (categoriaMaiorVariacao) {
    return {
      priority: 1, // Prioridade alta
      title: "👀 Destaque da Semana",
      text: `Os seus gastos com *${escapeMarkdown(categoriaMaiorVariacao.nome)}* aumentaram *${categoriaMaiorVariacao.variacao.toFixed(0)}%* em relação à sua média semanal (de ${formatCurrency(categoriaMaiorVariacao.media)} para ${formatCurrency(categoriaMaiorVariacao.gastoAtual)}).`
    };
  }

  return null;
}


/**
 * @private
 * Deteta novas despesas recorrentes (potenciais assinaturas) nos últimos 30 dias.
 * @returns {object|null} Um objeto de insight ou null.
 */
function _detectNewRecurringExpenses(usuario, transacoes) {
  const hoje = new Date();
  const inicioPeriodoRecente = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - 30);
  const inicioPeriodoAnterior = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - 60);

  const gastosRecentes = {};
  const gastosAnteriores = {};

  for (let i = 1; i < transacoes.length; i++) {
    const linha = transacoes[i];
    const dataTransacao = parseData(linha[0]);
    if (dataTransacao >= inicioPeriodoAnterior && normalizarTexto(linha[11]) === normalizarTexto(usuario) && (linha[4] || "").toLowerCase() === "despesa") {
      const descricao = normalizarTexto(linha[1]);
      if (dataTransacao >= inicioPeriodoRecente) {
        gastosRecentes[descricao] = (gastosRecentes[descricao] || 0) + 1;
      } else {
        gastosAnteriores[descricao] = (gastosAnteriores[descricao] || 0) + 1;
      }
    }
  }

  const novasAssinaturas = [];
  for (const desc in gastosRecentes) {
    if (gastosRecentes[desc] >= 2 && !gastosAnteriores[desc]) {
      novasAssinaturas.push(capitalize(desc));
    }
  }

  if (novasAssinaturas.length > 0) {
    return {
      priority: 0, // Prioridade máxima
      title: "🧐 Nova Despesa Recorrente?",
      text: `Notámos novos gastos recorrentes com: *${escapeMarkdown(novasAssinaturas.join(', '))}*. Trata-se de uma nova assinatura?`
    };
  }
  
  return null;
}


/**
 * @private
 * Retorna o intervalo de datas para a última semana completa (Domingo a Sábado).
 * @returns {object} Um objeto com as datas { inicioDaSemana, fimDaSemana }.
 */
function _getLastWeekDateRange() {
  const hoje = new Date();
  const diaDaSemana = hoje.getDay(); // 0=Dom, 6=Sáb
  const fimDaSemana = new Date(hoje);
  fimDaSemana.setDate(hoje.getDate() - diaDaSemana - 1); // Último sábado
  fimDaSemana.setHours(23, 59, 59, 999);
  
  const inicioDaSemana = new Date(fimDaSemana);
  inicioDaSemana.setDate(fimDaSemana.getDate() - 6); // Último domingo
  inicioDaSemana.setHours(0, 0, 0, 0);

  return { inicioDaSemana, fimDaSemana };
}


/**
 * @private
 * Calcula o valor total de uma fatura de cartão para um mês/ano de vencimento específico.
 * @param {string} cardName O nome do cartão.
 * @param {number} targetMonth O mês de vencimento (0-11).
 * @param {number} targetYear O ano de vencimento.
 * @param {Array<Array<any>>} dadosTransacoes Os dados da aba de transações.
 * @returns {number} O valor total da fatura.
 */
function _calculateBillForCard(cardName, targetMonth, targetYear, dadosTransacoes) {
  let totalFatura = 0;
  const cardNameNormalized = normalizarTexto(cardName);

  for (let i = 1; i < dadosTransacoes.length; i++) {
    const row = dadosTransacoes[i];
    const tipo = (row[4] || "").toLowerCase();
    const conta = normalizarTexto(row[7]);
    const dataVencimento = parseData(row[10]);

    if (
      tipo === "despesa" &&
      conta === cardNameNormalized &&
      dataVencimento &&
      dataVencimento.getMonth() === targetMonth &&
      dataVencimento.getFullYear() === targetYear
    ) {
      // Exclui pagamentos de fatura para não abater no valor
      const categoria = normalizarTexto(row[2]);
      const subcategoria = normalizarTexto(row[3]);
      if (!(categoria === "contas a pagar" && subcategoria === "pagamento de fatura")) {
        totalFatura += parseBrazilianFloat(String(row[5]));
      }
    }
  }
  return round(totalFatura, 2);
}
