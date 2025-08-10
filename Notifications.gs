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

  // A função getNotificationConfig carrega as configurações da SHEET_NOTIFICACOES_CONFIG
  const notificationConfig = getNotificationConfig(); 

  if (!notificationConfig) {
    logToSheet("Configurações de notificações não encontradas. Nenhuma notificação será enviada.", "WARN");
    return;
  }

  // Envia notificações para cada usuário/grupo configurado
  for (const chatId in notificationConfig) {
    const userConfig = notificationConfig[chatId];
    logToSheet(`Verificando configurações de notificação para Chat ID: ${chatId} (Usuário: ${userConfig.usuario})`, "DEBUG");

    if (userConfig.enableBudgetAlerts) {
      sendBudgetAlerts(chatId, userConfig.usuario);
    }
    if (userConfig.enableBillReminders) {
      sendUpcomingBillReminders(chatId, userConfig.usuario);
    }
    if (userConfig.enableDailySummary && isTimeForDailySummary(userConfig.dailySummaryTime)) {
      sendDailySummary(chatId, userConfig.usuario);
    }
    if (userConfig.enableWeeklySummary && isTimeForWeeklySummary(userConfig.weeklySummaryDay, userConfig.weeklySummaryTime)) {
      sendWeeklySummary(chatId, userConfig.usuario);
    }
  }

  logToSheet("Verificação e envio de notificações concluídos.", "INFO");
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
  
  // Verifica se a hora atual está dentro de um pequeno intervalo da hora configurada.
  // Isso é importante porque gatilhos de tempo não são executados no milissegundo exato.
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

  const currentDay = now.getDay(); // 0 for Sunday, 6 for Saturday
  const currentHour = now.getHours();
  const currentMinute = now.getMinutes();

  return currentDay === dayOfWeek && currentHour === configHour && currentMinute >= configMinute && currentMinute < configMinute + 5;
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
  // Agora lê da aba Notificacoes_Config, conforme o plano original
  const configSheet = ss.getSheetByName(SHEET_NOTIFICACOES_CONFIG); 

  if (!configSheet) {
    logToSheet("Aba 'Notificacoes_Config' não encontrada. Nenhuma configuracao de notificacao sera lida.", "ERROR");
    return null;
  }

  const data = configSheet.getDataRange().getValues();
  const headers = data[0]; // Primeira linha são os cabeçalhos

  // Mapeia índices de coluna
  const colChatId = headers.indexOf('Chat ID');
  const colUsuario = headers.indexOf('Usuário');
  const colEnableBudgetAlerts = headers.indexOf('Alertas Orçamento');
  const colEnableBillReminders = headers.indexOf('Lembretes Contas a Pagar');
  const colEnableDailySummary = headers.indexOf('Resumo Diário');
  const colDailySummaryTime = headers.indexOf('Hora Resumo Diário (HH:mm)');
  const colEnableWeeklySummary = headers.indexOf('Resumo Semanal');
  const colWeeklySummaryDay = headers.indexOf('Dia Resumo Semanal (0-6)');
  const colWeeklySummaryTime = headers.indexOf('Hora Resumo Semanal (HH:mm)');

  // Verifica se as colunas essenciais para as notificações existem
  if ([colChatId, colUsuario, colEnableBudgetAlerts, colEnableBillReminders,
       colEnableDailySummary, colDailySummaryTime, colEnableWeeklySummary,
       colWeeklySummaryDay, colWeeklySummaryTime].some(idx => idx === -1)) {
    logToSheet("Colunas essenciais para 'Notificacoes_Config' ausentes. Verifique os cabeçalhos.", "ERROR");
    return null;
  }

  const notificationConfig = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const chatId = (row[colChatId] || "").toString().trim();
    if (chatId) { // Apenas processa se o Chat ID não estiver vazio
      notificationConfig[chatId] = {
        usuario: (row[colUsuario] || "").toString().trim(),
        enableBudgetAlerts: (row[colEnableBudgetAlerts] || "").toString().toLowerCase().trim() === 'sim',
        enableBillReminders: (row[colEnableBillReminders] || "").toString().toLowerCase().trim() === 'sim',
        enableDailySummary: (row[colEnableDailySummary] || "").toString().toLowerCase().trim() === 'sim',
        dailySummaryTime: (row[colDailySummaryTime] || "").toString().trim(),
        enableWeeklySummary: (row[colEnableWeeklySummary] || "").toString().toLowerCase().trim() === 'sim',
        weeklySummaryDay: parseInt(row[colWeeklySummaryDay]), // 0=Domingo, 6=Sábado
        weeklySummaryTime: (row[colWeeklySummaryTime] || "").toString().trim()
      };
    }
  }
  return notificationConfig;
}
