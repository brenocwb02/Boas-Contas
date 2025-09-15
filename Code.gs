/**
 * @file Code.gs
 * @description Este arquivo contém a função principal `doPost` que atua como o webhook do Telegram,
 * processando as mensagens e callbacks recebidas.
 */

// Variável global para armazenar os saldos calculados.
// Usar `globalThis` é uma boa prática para garantir que ela seja acessível em diferentes arquivos .gs.
// É populada pela função `atualizarSaldosDasContas` em FinancialLogic.gs.
globalThis.saldosCalculados = {};

/**
 * **FUNÇÃO CORRIGIDA E FINALIZADA**
 * Processa uma consulta em linguagem natural do usuário.
 * Ex: "quanto gastei com ifood este mês?", "listar despesas com transporte em junho"
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 * @param {string} textoConsulta A pergunta completa do usuário.
 */
function processarConsultaLinguagemNatural(chatId, usuario, textoConsulta) {
  logToSheet(`[ConsultaLN] Iniciando processamento para: "${textoConsulta}"`, "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
  if (!transacoesSheet) {
    enviarMensagemTelegram(chatId, "❌ Erro: Aba 'Transacoes' não encontrada para realizar a consulta.");
    return;
  }
  const transacoes = transacoesSheet.getDataRange().getValues();
  const consultaNormalizada = normalizarTexto(textoConsulta);

  // --- 1. Determinar o Período de Tempo ---
  const hoje = new Date();
  let dataInicio = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
  let dataFim = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0, 23, 59, 59);
  let periodoTexto = "este mês";
  let periodoDetectado = "este mes";

  const meses = { "janeiro": 0, "fevereiro": 1, "marco": 2, "abril": 3, "maio": 4, "junho": 5, "julho": 6, "agosto": 7, "setembro": 8, "outubro": 9, "novembro": 10, "dezembro": 11 };
  for (const nomeMes in meses) {
    if (consultaNormalizada.includes(nomeMes)) {
      const mesIndex = meses[nomeMes];
      let ano = hoje.getFullYear();
      if (mesIndex > hoje.getMonth() && !/\d{4}/.test(consultaNormalizada)) {
        ano--;
      }
      dataInicio = new Date(ano, mesIndex, 1);
      dataFim = new Date(ano, mesIndex + 1, 0, 23, 59, 59);
      periodoTexto = `em ${capitalize(nomeMes)}`;
      periodoDetectado = nomeMes;
      break;
    }
  }

  if (consultaNormalizada.includes("mes passado")) {
    dataInicio = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
    dataFim = new Date(hoje.getFullYear(), hoje.getMonth(), 0, 23, 59, 59);
    periodoTexto = "no mês passado";
    periodoDetectado = "mes passado";
  } else if (consultaNormalizada.includes("hoje")) {
    dataInicio = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
    dataFim = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate(), 23, 59, 59);
    periodoTexto = "hoje";
    periodoDetectado = "hoje";
  } else if (consultaNormalizada.includes("ontem")) {
    const ontem = new Date(hoje);
    ontem.setDate(hoje.getDate() - 1);
    dataInicio = new Date(ontem.getFullYear(), ontem.getMonth(), ontem.getDate());
    dataFim = new Date(ontem.getFullYear(), ontem.getMonth(), ontem.getDate(), 23, 59, 59);
    periodoTexto = "ontem";
    periodoDetectado = "ontem";
  }

  logToSheet(`[ConsultaLN] Período de tempo determinado: ${dataInicio.toLocaleDateString()} a ${dataFim.toLocaleDateString()} (${periodoTexto})`, "DEBUG");

  // --- 2. Determinar o Tipo de Consulta e Filtros ---
  const tipoConsulta = consultaNormalizada.includes("listar") || consultaNormalizada.includes("quais") ? "LISTAR" : "SOMAR";
  let tipoTransacaoFiltro = null;
  if (consultaNormalizada.includes("despesa")) tipoTransacaoFiltro = "Despesa";
  if (consultaNormalizada.includes("receita")) tipoTransacaoFiltro = "Receita";
  
  // --- INÍCIO DA LÓGICA CORRIGIDA PARA EXTRAIR O FILTRO DE TEXTO ---
  let consultaSemPeriodo = consultaNormalizada.replace(periodoDetectado, ' ').trim();
  let filtroTexto = "";
  const regexFiltro = /(?:com|de|sobre)\s+(.+?)(?=$)/; // Simplificado para pegar tudo após a preposição
  const matchFiltro = consultaSemPeriodo.match(regexFiltro);
  
  if (matchFiltro) {
    filtroTexto = matchFiltro[1].trim();
  } else {
    let tempConsulta = ' ' + consultaSemPeriodo + ' ';
    const palavrasParaRemover = [
      "quanto gastei", "listar despesas", "total de", "quanto recebi", "listar receitas",
      "quanto", "qual", "quais", "listar", "mostrar", "total", "despesas", "receitas", "despesa", "receita",
      "meu", "minha", "meus", "minhas", "de", "do", "da", "em", "no", "na"
    ];
    palavrasParaRemover.sort((a,b) => b.length - a.length).forEach(palavra => {
        tempConsulta = tempConsulta.replace(new RegExp(`\\s${palavra}\\s`, 'gi'), ' ');
    });
    filtroTexto = tempConsulta.trim();
  }
  // --- FIM DA LÓGICA CORRIGIDA ---

  logToSheet(`[ConsultaLN] Tipo: ${tipoConsulta}, Filtro de Tipo: ${tipoTransacaoFiltro || 'Nenhum'}, Filtro de Texto: "${filtroTexto}"`, "DEBUG");

  // --- 3. Executar a Consulta ---
  let totalSoma = 0;
  let transacoesEncontradas = [];
  
  for (let i = 1; i < transacoes.length; i++) {
    const linha = transacoes[i];
    const dataTransacao = parseData(linha[0]);
    const descricao = linha[1];
    const categoria = linha[2];
    const subcategoria = linha[3];
    const tipo = linha[4];
    const valor = parseBrazilianFloat(String(linha[5]));
    const conta = linha[7];
    const id = linha[13];

    if (!dataTransacao || dataTransacao < dataInicio || dataTransacao > dataFim) continue;
    if (tipoTransacaoFiltro && normalizarTexto(tipo) !== normalizarTexto(tipoTransacaoFiltro)) continue;

    if (filtroTexto) {
        const relevanteParaFiltro = (
            normalizarTexto(descricao).includes(normalizarTexto(filtroTexto)) ||
            normalizarTexto(categoria).includes(normalizarTexto(filtroTexto)) ||
            normalizarTexto(subcategoria).includes(normalizarTexto(filtroTexto)) ||
            normalizarTexto(conta).includes(normalizarTexto(filtroTexto))
        );
        if (!relevanteParaFiltro) continue;
    }

    if (tipo === "Despesa" && (normalizarTexto(categoria) === "contas a pagar" && normalizarTexto(subcategoria) === "pagamento de fatura" || normalizarTexto(categoria) === "transferencias")) continue;
    if (tipo === "Receita" && normalizarTexto(categoria) === "transferencias") continue;

    if (tipoConsulta === "SOMAR") {
      totalSoma += valor;
    } else {
      transacoesEncontradas.push({ data: Utilities.formatDate(dataTransacao, Session.getScriptTimeZone(), "dd/MM/yyyy"), descricao, categoria, subcategoria, tipo, valor, conta, id });
    }
  }

  let mensagemResposta = "";
  if (tipoConsulta === "SOMAR") {
    let prefixoTipo = tipoTransacaoFiltro === "Receita" ? "Receitas" : "Gastos";
    if (!tipoTransacaoFiltro && filtroTexto) prefixoTipo = "Gastos";
    mensagemResposta = `O *total de ${prefixoTipo}* ${filtroTexto ? `com "${escapeMarkdown(filtroTexto)}"` : ""} ${periodoTexto} é de: *${formatCurrency(totalSoma)}*.`;
  } else {
    if (transacoesEncontradas.length > 0) {
      mensagemResposta = `*Lançamentos ${filtroTexto ? `de "${escapeMarkdown(filtroTexto)}"` : ""} ${tipoTransacaoFiltro ? `(${escapeMarkdown(tipoTransacaoFiltro)})` : ''} ${periodoTexto}:*\n\n`;
      transacoesEncontradas.sort((a, b) => parseData(b.data).getTime() - parseData(a.data).getTime());
      transacoesEncontradas.slice(0, 10).forEach(t => {
        const valorFormatado = formatCurrency(t.valor);
        const tipoIcon = t.tipo === "Receita" ? "💰" : "💸";
        mensagemResposta += `${tipoIcon} ${escapeMarkdown(t.descricao)} (${escapeMarkdown(t.categoria)} > ${escapeMarkdown(t.subcategoria)}) - *${valorFormatado}*\n`;
      });
      if (transacoesEncontradas.length > 10) {
        mensagemResposta += `\n...e mais ${transacoesEncontradas.length - 10} lançamentos.`;
      }
    } else {
      mensagemResposta = `Nenhum lançamento ${filtroTexto ? `com "${escapeMarkdown(filtroTexto)}"` : ""} encontrado ${periodoTexto}.`;
    }
  }

  enviarMensagemTelegram(chatId, mensagemResposta);
  logToSheet(`[ConsultaLN] Resposta enviada para ${chatId}: "${mensagemResposta.substring(0, 100)}..."`, "INFO");
}

/**
 * NOVO: Gera e envia um link de acesso seguro e temporário para o Dashboard.
 * Esta função foi movida para este arquivo para corrigir o erro 'not defined'.
 * @param {string} chatId O ID do chat do Telegram.
 */
function enviarLinkDashboard(chatId) {
  logToSheet(`[Dashboard] Gerando link de acesso seguro para o chatId: ${chatId}`, "INFO");
  const cache = CacheService.getScriptCache();
  
  // 1. Gera um token único e aleatório.
  const token = Utilities.getUuid();

  // 2. Armazena o token no cache, associando-o ao chatId do usuário.
  // A chave é o token, o valor é o chatId. A validade é definida na constante.
  const cacheKey = `${CACHE_KEY_DASHBOARD_TOKEN}_${token}`;
  cache.put(cacheKey, chatId.toString(), CACHE_EXPIRATION_DASHBOARD_TOKEN_SECONDS);
  logToSheet(`[Dashboard] Token '${token}' armazenado no cache para o chatId '${chatId}' por ${CACHE_EXPIRATION_DASHBOARD_TOKEN_SECONDS} segundos.`, "DEBUG");

  // 3. Obtém a URL do Web App.
  const webAppUrl = ScriptApp.getService().getUrl();

  // 4. Constrói a URL segura com o token como parâmetro.
  const secureUrl = `${webAppUrl}?token=${token}`;

  // 5. Envia a mensagem para o usuário.
  const mensagem = `Aqui está o seu link de acesso seguro ao Dashboard Financeiro. \n\n` +
                   `*Atenção:* Este link é de uso único e expira em ${CACHE_EXPIRATION_DASHBOARD_TOKEN_SECONDS / 60} minutos.\n\n` +
                   `[Clique aqui para abrir o Dashboard](${secureUrl})`;
  
  // Envia com parse_mode 'Markdown' para garantir a formatação do link.
  enviarMensagemTelegram(chatId, mensagem, { parse_mode: 'Markdown' });
}

/**
 * Função executada quando o Add-on é instalado pela primeira vez pelo utilizador.
 * @param {Object} e O objeto de evento de instalação.
 */
function onInstall(e) {
  onOpen(e);
}


/**
 * **FUNÇÃO ATUALIZADA**
 * Função principal que é acionada pelo webhook do Telegram.
 * O bloco try...catch agora usa a nova função centralizada handleError e o `finally` para descarregar os logs.
 * @param {Object} e O objeto de evento do webhook.
 */
function doPost(e) {
  let chatIdForError = null;
  try {
    if (!isLicenseValid()) {
      logToSheet("Acesso bloqueado via doPost: Licença inválida ou ausente.", "ERROR");
      return; 
    }
   
    if (!e || !e.postData || !e.postData.contents) {
      logToSheet("doPost recebido com dados vazios ou invalidos.", "WARN");
      return;
    }

    currentLogLevel = getLogLevelConfig();
    const data = JSON.parse(e.postData.contents || '{}');
    const chatId = data.message?.chat?.id || data.callback_query?.message?.chat?.id;
    chatIdForError = chatId; // Armazena o chatId para usar no bloco catch
    let textoRecebido = (data.message?.text || data.callback_query?.data || "").trim();

    if (!chatId) {
        logToSheet("Requisição recebida sem chatId. Ignorando.", "WARN");
        return;
    }

    // --- Carrega dados essenciais no início ---
    const configData = getSheetDataWithCache(SHEET_CONFIGURACOES, CACHE_KEY_CONFIG);
    const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);
    const usuario = getUsuarioPorChatId(chatId, configData);
    
    // **CORREÇÃO CRÍTICA:** Lógica para iniciar e continuar a configuração guiada.
    const guidedSetupState = getGuidedSetupState(chatId);

    // Se o comando for /start E o utilizador estiver no primeiro passo da configuração, inicia a conversa.
    if (textoRecebido.toLowerCase() === '/start' && guidedSetupState && guidedSetupState.step === SETUP_STEPS.PENDING_START) {
        logToSheet(`[doPost] Comando /start recebido. A iniciar configuração guiada para ${chatId}.`, "INFO");
        startGuidedSetup(chatId, usuario);
        return; // Interrompe a execução para não cair no /start padrão
    }
    
    // Se já estiver noutro passo da configuração, processa a resposta.
    if (guidedSetupState && guidedSetupState.step !== SETUP_STEPS.PENDING_START) {
      logToSheet(`[doPost] Utilizador ${chatId} está na configuração guiada. A processar...`, "INFO");
      handleGuidedSetupInput(chatId, usuario, textoRecebido, guidedSetupState);
      if (data.callback_query) answerCallbackQuery(data.callback_query.id);
      return;
    }

    // O resto da sua função original `doPost` continua a partir daqui...
    if (data.message && data.message.voice) {
      logToSheet("Mensagem de voz recebida.", "INFO");
      handleVoiceMessage(data.message); // Chama o novo handler do VoiceHandler.gs
      return; // Encerra a execução para esta requisição, pois já foi tratada.
    }
    
    // Verifica se existe um estado de edição ou assistente ativo.
    const editState = getEditState(chatId);
    const assistantState = getActiveAssistantState(chatId); // USA A NOVA FUNÇÃO

    const updateId = data.update_id;
    if (updateId) {
      const cache = CacheService.getScriptCache();
      const cachedUpdate = cache.get(updateId.toString());
      if (cachedUpdate) {
        logToSheet(`Update ID ${updateId} ja processado. Ignorando execucao duplicada.`, "WARN");
        return;
      }
      cache.put(updateId.toString(), "processed", 60);
    }

    let comandoBase;
    let complemento;

    if (data.callback_query) {
      answerCallbackQuery(data.callback_query.id);
      logToSheet(`Callback query ID ${data.callback_query.id} reconhecida.`, "DEBUG");

      if (textoRecebido.startsWith('confirm_')) {
        comandoBase = "/confirm";
        complemento = textoRecebido.substring('confirm_'.length);
      } else if (textoRecebido.startsWith('cancel_')) {
        comandoBase = "/cancel";
        complemento = textoRecebido.substring('cancel_'.length);
      }
      // --- NOVO: Lógica para tratar respostas do Quiz ---
      else if (textoRecebido.startsWith('quiz_')) {
        const parts = textoRecebido.split('_');
        const questionIndex = parseInt(parts[1], 10);
        const optionIndex = parseInt(parts[2], 10);
        handleQuizAnswer(chatId, questionIndex, optionIndex);
        return; // Finaliza a execução aqui
      }
      // --- FIM DA NOVA LÓGICA ---
      else if (textoRecebido.startsWith('complete_')) {
        comandoBase = "/complete_assistant_action";
        complemento = textoRecebido.substring('complete_'.length);
      }
      // --- TAREFAS: Lógica para os botões inline (callbacks) ---
      else if (textoRecebido.startsWith('add_agenda_')) {
        const idTarefa = textoRecebido.substring('add_agenda_'.length);
        adicionarTarefaNaAgenda(chatId, idTarefa);
        return; // Finaliza a execução
      }
      else if (textoRecebido.startsWith('concluir_tarefa_')) {
        const idTarefa = textoRecebido.substring('concluir_tarefa_'.length);
        concluirTarefa(chatId, idTarefa);
        return; // Finaliza a execução
      }
      // --- TAREFAS: NOVO CALLBACK PARA EXCLUIR TAREFA ---
      else if (textoRecebido.startsWith('excluir_tarefa_')) {
        const idTarefa = textoRecebido.substring('excluir_tarefa_'.length);
        excluirTarefa(chatId, idTarefa);
        return; // Finaliza a execução
      }
      // --- FIM TAREFAS ---
      else if (textoRecebido.startsWith('/tutorial_')) {
        comandoBase = textoRecebido;
        complemento = "";
      }
      else if (textoRecebido.startsWith("/marcar_pago_")) {
        comandoBase = "/marcar_pago";
        complemento = textoRecebido.substring("/marcar_pago_".length);
      } else if (textoRecebido.startsWith("/excluir_")) {
        comandoBase = "/excluir";
        complemento = textoRecebido.substring("/excluir_".length);
      } else if (textoRecebido.startsWith("/extrato_usuario_")) {
        comandoBase = "/extrato_usuario";
        complemento = textoRecebido.substring("/extrato_usuario_".length);
      }
      else if (textoRecebido.startsWith('edit_')) {
          comandoBase = "/editar_campo_callback";
          complemento = textoRecebido.substring("edit_".length);
      }
      else if (textoRecebido === "cancelar_edicao") {
          comandoBase = "/cancelar_edicao";
          complemento = "";
      }
      else {
        comandoBase = textoRecebido.startsWith("/") ? textoRecebido : "/" + textoRecebido;
        complemento = "";
      }
    }
    else if (data.message) {
      const textoLimpo = textoRecebido.trim();
      const partesTexto = textoLimpo.split(/\s+/);
      const primeiraPalavra = partesTexto[0].toLowerCase();
      
      const comandoNormalizado = primeiraPalavra.startsWith('/') ? primeiraPalavra.substring(1) : primeiraPalavra;

      // --- TAREFAS: Adicionar novos comandos conhecidos ---
      const comandosConhecidosSemBarra = [
          "start", "dashboard", "resumo", "saldo", "extrato", "proximasfaturas", 
          "contasapagar", "ajuda", "editar", "vincular_conta", "tutorial", 
          "adicionar_conta", "listar_contas", "adicionar_categoria", "listar_categorias", 
          "listar_subcategorias", "tarefa", "lembrete", "tarefas", "agenda", "concluir", 
          "excluir_tarefa", "saude",
          "orcamento", "metas", "novameta", "aportarmeta", "patrimonio", "meuperfil",
          "comprar_ativo", "vender_ativo", "importar", "arquivar_agora"  // <-- ADICIONADOS AQUI
      ];

      if (comandosConhecidosSemBarra.includes(comandoNormalizado)) {
          comandoBase = `/${comandoNormalizado}`;
          complemento = partesTexto.slice(1).join(" ");
      } 
      // ### LÓGICA DE INTERCEPTAÇÃO ATUALIZADA ###
      else if (assistantState && !textoLimpo.startsWith('/')) {
          logToSheet(`[doPost] Estado do assistente detectado para ${chatId}. Processando resposta digitada.`, "INFO");
          processarRespostaDoAssistente(chatId, usuario, textoRecebido, assistantState);
          return; // Finaliza a execução
      }
      else if (editState && !textoLimpo.startsWith('/')) {
          logToSheet(`[doPost] Estado de edição detectado para ${chatId}. Processando entrada de edição.`, "INFO");
          processarEdicaoFinal(chatId, usuario, textoRecebido, editState, dadosContas);
          return; // Finaliza a execução
      }
      // ### FIM DA LÓGICA DE INTERCEPTAÇÃO ###
      else {
          comandoBase = "/lancamento";
          complemento = textoLimpo;
      }
    } else {
      logToSheet("Webhook recebido, mas sem mensagem ou callback query reconhecida.", "INFO");
      return;
    }

    logToSheet(`doPost - Chat ID: ${chatId}, Texto Recebido: "${textoRecebido}", Comando Base: "${comandoBase}", Complemento: "${complemento}"`, "INFO");

    const debugTutorialState = getTutorialState(chatId);
    logToSheet(`[DEBUG doPost Start] ChatID: ${chatId}, ComandoBase: "${comandoBase}", TextoRecebido: "${textoRecebido}", TutorialState: ${JSON.stringify(debugTutorialState)}`, "DEBUG");

    if (usuario === "Desconhecido") {
      enviarMensagemTelegram(chatId, "❌ Voce não está autorizado a usar este bot.");
      logToSheet(`Usuario ${chatId} não autorizado.`, "WARN");
      return;
    }

    const { month: targetMonth, year: targetYear } = parseMonthAndYear(complemento);
    logToSheet(`doPost - Mes Alvo: ${targetMonth}, Ano Alvo: ${targetYear}`, "DEBUG");

    if (debugTutorialState && debugTutorialState.currentStep > 0 &&
        !comandoBase.startsWith("/tutorial_") &&
        comandoBase !== "/confirm" &&
        comandoBase !== "/cancel") {

      logToSheet(`[doPost] Usuario ${chatId} esta no tutorial (Passo ${debugTutorialState.currentStep}, Acao Esperada: ${debugTutorialState.expectedAction}). Tentando processar input pelo tutorial.`, "INFO");
      const handledByTutorial = processTutorialInput(chatId, usuario, textoRecebido, debugTutorialState);
      if (handledByTutorial) {
        logToSheet(`[doPost] Mensagem tratada pelo tutorial.`, "INFO");
        return;
      }
    }


    // --- Processamento dos comandos ---
    switch (comandoBase) {
      case "/start":
          const startMessage = `Olá ${escapeMarkdown(usuario)}! Bem-vindo ao Boas Contas. 👋 Sou o Zaq, o seu agente financeiro pessoal.\n\n` +
                               `Estou a postos para o ajudar a registar gastos, receitas, ver o seu saldo e muito mais. A minha missão é tornar o controlo das suas finanças tão simples como ter uma conversa.\n\n` +
                               `*Para começar, tente algo como:*\n` +
                               `• \`gastei 50 no mercado com o Nubank\`\n` +
                               `• \`recebi 3000 de salário no Itaú\`\n` +
                               `• \`transferi 100 do Itaú para o PicPay\`\n\n` +
                               `Para ver tudo o que posso fazer, use o comando /ajuda. ✨`;
          enviarMensagemTelegram(chatId, startMessage, { parse_mode: 'Markdown' });
          return;

      // --- NOVO CASE PARA O QUIZ ---
      case "/meuperfil":
          logToSheet(`Comando /meuperfil detectado.`, "INFO");
          handleMeuPerfilCommand(chatId, usuario);
          return;
      // --- FIM DO NOVO CASE ---
      
      // --- NOVO COMANDO PARA PROVENTOS ---
      case "/provento":
          logToSheet(`Comando /provento detectado. Complemento: "${complemento}"`, "INFO");
          handleProventoCommand(chatId, complemento, usuario);
          return;
      // --- FIM DO NOVO COMANDO ---

      case "/confirm":
        logToSheet(`Comando /confirm detectado para transacao ID: ${complemento}`, "INFO");
        const cacheConfirm = CacheService.getScriptCache();
        const cacheKeyConfirm = `${CACHE_KEY_PENDING_TRANSACTIONS}_${chatId}_${complemento}`;
        const cachedTransactionDataConfirm = cacheConfirm.get(cacheKeyConfirm);

        if (cachedTransactionDataConfirm) {
          const transacaoData = JSON.parse(cachedTransactionDataConfirm);
          registrarTransacaoConfirmada(transacaoData, usuario, chatId);
          cacheConfirm.remove(cacheKeyConfirm);
        } else {
          enviarMensagemTelegram(chatId, "⚠️ Esta confirmação expirou ou já foi processada.");
          logToSheet(`CallbackQuery para transacao ID ${complemento} recebida, mas dados nao encontrados no cache (confirm).`, "WARN");
        }
        return;

      case "/cancel":
        logToSheet(`Comando /cancel detectado para transacao ID: ${complemento}`, "INFO");
        const cacheCancel = CacheService.getScriptCache();
        const cacheKeyCancel = `${CACHE_KEY_PENDING_TRANSACTIONS}_${chatId}_${complemento}`;
        const cachedTransactionDataCancel = cacheCancel.get(cacheKeyCancel);

        if (cachedTransactionDataCancel) {
          cancelarTransacaoPendente(chatId, complemento);
          cacheCancel.remove(cacheKeyCancel);
        } else {
          enviarMensagemTelegram(chatId, "⚠️ Este cancelamento expirou ou já foi processada.");
          logToSheet(`CallbackQuery para transacao ID ${complemento} recebida, mas dados nao encontrados no cache (cancel).`, "WARN");
        }
        return;
      
      case "/complete_assistant_action":
        const partesComplemento = complemento.split('_');
        const indiceSelecionado = parseInt(partesComplemento.pop(), 10);
        const transacaoId = partesComplemento.pop();
        const campo = partesComplemento.join('_');

        const transacaoParcial = getAssistantState(chatId, transacaoId);

        if (transacaoParcial && transacaoParcial.assistantOptions && transacaoParcial.assistantOptions[indiceSelecionado] !== undefined) {
          
          const valorRealSelecionado = transacaoParcial.assistantOptions[indiceSelecionado];
          
          transacaoParcial.assistantOptions = null;

          if (campo === 'conta' || campo === 'conta_origem' || campo === 'conta_destino') {
            const infoConta = obterInformacoesDaConta(valorRealSelecionado, dadosContas);
            if(campo === 'conta') {
                transacaoParcial.conta = valorRealSelecionado;
                transacaoParcial.infoConta = infoConta;
                if (infoConta && normalizarTexto(infoConta.tipo) === "cartao de credito") {
                    transacaoParcial.metodoPagamento = "Crédito";
                }
            } else if (campo === 'conta_origem') {
                transacaoParcial.contaOrigem = valorRealSelecionado;
            } else if (campo === 'conta_destino') {
                transacaoParcial.contaDestino = valorRealSelecionado;
            }
          } else if (campo === 'categoria') {
            transacaoParcial.categoria = valorRealSelecionado;
            const dadosPalavras = getSheetDataWithCache(SHEET_PALAVRAS_CHAVE, CACHE_KEY_PALAVRAS);
            const { subcategoria } = extrairCategoriaSubcategoria(transacaoParcial.descricao, transacaoParcial.tipo, dadosPalavras);
            if (subcategoria !== "Não Identificada") {
                transacaoParcial.subcategoria = subcategoria;
            }
          } else if (campo === 'subcategoria') {
            transacaoParcial.subcategoria = valorRealSelecionado;
          } else if (campo === 'metodo') {
            transacaoParcial.metodoPagamento = valorRealSelecionado;
          }

          // ### INÍCIO DA CORREÇÃO ###
          // Usa a nova função para limpar o estado e o ponteiro
          clearActiveAssistantState(chatId);
          // ### FIM DA CORREÇÃO ###
          
          processAssistantCompletion(transacaoParcial, chatId, usuario);

        } else {
          enviarMensagemTelegram(chatId, "⚠️ Esta ação expirou ou é inválida. Por favor, envie o lançamento novamente.");
          if(transacaoParcial) {
            // ### INÍCIO DA CORREÇÃO ###
            clearActiveAssistantState(chatId);
            // ### FIM DA CORREÇÃO ###
          }
        }
        return;

      case "/dashboard":
        logToSheet(`Comando /dashboard detectado.`, "INFO");
        enviarLinkDashboard(chatId);
        return;

      // --- TAREFAS: NOVOS COMANDOS PARA O MÓDULO DE TAREFAS ---
      case "/tarefa":
      case "/lembrete":
          logToSheet(`Comando /tarefa ou /lembrete detectado. Complemento: "${complemento}"`, "INFO");
          if (complemento) {
            criarNovaTarefa(chatId, complemento);
          } else {
            enviarMensagemTelegram(chatId, "❌ Por favor, forneça a descrição da tarefa. Ex: `/tarefa Reunião com a equipe amanhã às 10h`");
          }
          return;

      case "/tarefas":
      case "/agenda":
          logToSheet(`Comando /tarefas ou /agenda detectado.`, "INFO");
          listarTarefasPendentes(chatId);
          return;

      case "/concluir":
          logToSheet(`Comando /concluir tarefa detectado. ID: "${complemento}"`, "INFO");
          if (complemento) {
            concluirTarefa(chatId, complemento);
          } else {
            enviarMensagemTelegram(chatId, "❌ Por favor, forneça o ID da tarefa que deseja concluir. Ex: `/concluir abc123`");
          }
          return;
      
      case "/excluir_tarefa":
          logToSheet(`Comando /excluir_tarefa detectado. ID: "${complemento}"`, "INFO");
           if (complemento) {
            excluirTarefa(chatId, complemento);
          } else {
            enviarMensagemTelegram(chatId, "❌ Por favor, forneça o ID da tarefa que deseja excluir. Ex: `/excluir_tarefa abc123`");
          }
          return;
      // --- FIM TAREFAS ---
      
      // ### INÍCIO DOS NOVOS CASES DE INVESTIMENTO ###
      case "/comprar_ativo":
          logToSheet(`Comando /comprar_ativo detectado. Complemento: "${complemento}"`, "INFO");
          const compraMatch = complemento.match(/([a-zA-Z0-9]+)\s+([\d.,]+)\s+([\d.,]+)\s+(?:de|do|da)\s+(.+)/i);
          if (compraMatch) {
            const [, ticker, quantidade, preco, corretora] = compraMatch;
            handleComprarAtivo(chatId, ticker, parseFloat(quantidade.replace(',', '.')), parseBrazilianFloat(preco), corretora, usuario);
          } else {
            enviarMensagemTelegram(chatId, "❌ Formato inválido. Use: `/comprar_ativo TICKER QTD PREÇO de CONTA`\nEx: `/comprar_ativo ITSA4 100 8.50 do NuInvest`");
          }
          return;

      case "/vender_ativo":
          logToSheet(`Comando /vender_ativo detectado. Complemento: "${complemento}"`, "INFO");
          const vendaMatch = complemento.match(/([a-zA-Z0-9]+)\s+([\d.,]+)\s+([\d.,]+)\s+(?:para|pra)\s+(.+)/i);
           if (vendaMatch) {
            const [, ticker, quantidade, preco, contaDestino] = vendaMatch;
            handleVenderAtivo(chatId, ticker, parseFloat(quantidade.replace(',', '.')), parseBrazilianFloat(preco), contaDestino, usuario);
          } else {
            enviarMensagemTelegram(chatId, "❌ Formato inválido. Use: `/vender_ativo TICKER QTD PREÇO para CONTA`\nEx: `/vender_ativo ITSA4 50 10.00 para NuInvest`");
          }
          return;
      // ### FIM DOS NOVOS CASES DE INVESTIMENTO ###

      case "/adicionar_conta":
          logToSheet(`Comando /adicionar_conta detectado. Complemento: "${complemento}"`, "INFO");
          adicionarNovaConta(chatId, usuario, complemento);
          return;
      case "/listar_contas":
          logToSheet(`Comando /listar_contas detectado.`, "INFO");
          listarContas(chatId, usuario);
          return;
      case "/adicionar_categoria":
          logToSheet(`Comando /adicionar_categoria detectado. Complemento: "${complemento}"`, "INFO");
          adicionarNovaCategoria(chatId, usuario, complemento);
          return;
      case "/listar_categorias":
          logToSheet(`Comando /listar_categorias detectado.`, "INFO");
          listarCategorias(chatId);
          return;
      case "/listar_subcategorias":
          logToSheet(`Comando /listar_subcategorias detectado. Complemento: "${complemento}"`, "INFO");
          if (complemento) {
            listarSubcategorias(chatId, complemento);
          } else {
            enviarMensagemTelegram(chatId, "❌ Por favor, forneça o nome da categoria principal. Ex: `/listar_subcategorias Alimentação`");
          }
          return;
      case "/tutorial":
      case "/tutorial_start":
          logToSheet(`Comando /tutorial ou /tutorial_start detectado.`, "INFO");
          clearTutorialState(chatId);
          handleTutorialStep(chatId, usuario, 1);
          return;
      case "/tutorial_next":
          logToSheet(`Comando /tutorial_next detectado.`, "INFO");
          let tutorialStateNext = getTutorialState(chatId);
          if (tutorialStateNext && tutorialStateNext.currentStep > 0 && tutorialStateNext.currentStep < 6) {
            handleTutorialStep(chatId, usuario, tutorialStateNext.currentStep + 1);
          } else if (tutorialStateNext && tutorialStateNext.currentStep === 6) {
            handleTutorialStep(chatId, usuario, 6);
          } else {
            enviarMensagemTelegram(chatId, "🤔 Não há tutorial em andamento. Digite /tutorial para começar!");
            clearTutorialState(chatId);
          }
          return;
      case "/tutorial_prev":
          logToSheet(`Comando /tutorial_prev detectado.`, "INFO");
          let tutorialStatePrev = getTutorialState(chatId);
          if (tutorialStatePrev && tutorialStatePrev.currentStep > 1) {
            handleTutorialStep(chatId, usuario, tutorialStatePrev.currentStep - 1);
          } else {
            enviarMensagemTelegram(chatId, "Você já está no início do tutorial. Digite /tutorial para reiniciar.");
            clearTutorialState(chatId);
          }
          return;
      case "/tutorial_skip":
          logToSheet(`Comando /tutorial_skip detectado.`, "INFO");
          enviarMensagemTelegram(chatId, "Tutorial pulado. Se precisar de ajuda, digite /ajuda a qualquer momento.");
          clearTutorialState(chatId);
          return;
      case "/editar":
            if(normalizarTexto(complemento) === 'ultimo' || normalizarTexto(complemento) === 'último'){
                iniciarEdicaoUltimo(chatId, usuario);
            } else {
                enviarMensagemTelegram(chatId, "Comando de edição inválido. Use `/editar ultimo`.");
            }
            return;
      case "/editar_campo_callback":
            const campoParaEditar = complemento;
            solicitarNovoValorParaEdicao(chatId, campoParaEditar);
            return;
      case "/cancelar_edicao":
            clearEditState(chatId);
            enviarMensagemTelegram(chatId, "✅ Edição finalizada.");
            return;

      case "/extrato":
          logToSheet(`Comando /extrato detectado. Complemento: "${complemento}"`, "INFO");
          if (!complemento) {
            mostrarMenuExtrato(chatId);
          } else {
            enviarExtrato(chatId, usuario, complemento);
          }
          return;
      case "/extrato_tudo":
          logToSheet(`Comando /extrato_tudo detectado.`, "INFO");
          enviarExtrato(chatId, usuario, "tudo");
          return;
      case "/extrato_receitas":
          logToSheet(`Comando /extrato_receitas detectado.`, "INFO");
          enviarExtrato(chatId, usuario, "receitas");
          return;
      case "/extrato_despesas":
          logToSheet(`Comando /extrato_despesas detectado.`, "INFO");
          enviarExtrato(chatId, usuario, "despesas");
          return;
      case "/extrato_pessoa":
          logToSheet(`Comando /extrato_pessoa detectado.`, "INFO");
          mostrarMenuPorPessoa(chatId, configData);
          return;
      case "/resumo":
          const allUserNames = getAllUserNames(configData);
          const targetUser = findUserNameInText(complemento, allUserNames);
          const { month: targetMonthResumo, year: targetYearResumo } = parseMonthAndYear(complemento);

          if (targetUser) {
            logToSheet(`Comando /resumo por pessoa detectado para ${targetUser}.`, "INFO");
            enviarResumoPorPessoa(chatId, usuario, targetUser, targetMonthResumo, targetYearResumo);
          } else {
            logToSheet(`Comando /resumo geral detectado.`, "INFO");
            enviarResumo(chatId, usuario, targetMonthResumo, targetYearResumo);
          }
          return;
      case "/saldo":
          logToSheet(`Comando /saldo detectado.`, "INFO");
          enviarSaldo(chatId, usuario);
          return;
      case "/patrimonio": // <-- NOVO CASE ADICIONADO
          logToSheet(`Comando /patrimonio detectado.`, "INFO");
          enviarPatrimonioLiquido(chatId);
          return;
      case "/proximasfaturas":
          logToSheet(`Comando /proximasfaturas detectado.`, "INFO");
          enviarFaturasFuturas(chatId, usuario);
          return;
      case "/contasapagar":
          logToSheet(`Comando /contasapagar detectado. Mes: ${targetMonth}, Ano: ${targetYear}`, "INFO");
          enviarContasAPagar(chatId, usuario, targetMonth, targetYear);
          return;
      case "/marcar_pago":
          logToSheet(`Comando /marcar_pago detectado. ID da Conta: "${complemento}"`, "INFO");
          processarMarcarPago(chatId, textoRecebido, usuario);
          return;
      case "/excluir":
          logToSheet(`Comando /excluir detectado para ID: ${complemento}`, "INFO");
          excluirLancamentoPorId(complemento, chatId);
          return;
      case "/extrato_usuario":
          logToSheet(`Comando /extrato_usuario detectado para usuario: ${complemento}`, "INFO");
          enviarExtrato(chatId, usuario, complemento);
          return;
      case "/vincular_conta":
          logToSheet(`Comando /vincular_conta detectado. Complemento: "${complemento}"`, "INFO");
          const lastSpaceIndex = complemento.lastIndexOf(' ');
          if (lastSpaceIndex !== -1) {
            const idContaAPagar = complemento.substring(0, lastSpaceIndex).trim();
            const idTransacao = complemento.substring(lastSpaceIndex + 1).trim();
            if (idContaAPagar && idTransacao) {
              vincularTransacaoAContaAPagar(chatId, idContaAPagar, idTransacao);
            } else {
              enviarMensagemTelegram(chatId, "❌ Formato invalido para vincular. Use: `/vincular_conta <ID_CONTA_A_PAGAR> <ID_TRANSACAO>`");
            }
          } else {
            enviarMensagemTelegram(chatId, "❌ Formato invalido para vincular. Use: `/vincular_conta <ID_CONTA_A_PAGAR> <ID_TRANSACAO>`");
          }
          return;

      // --- NOVO CASE PARA O CHECK-UP FINANCEIRO ---
      case "/saude":
          logToSheet(`Comando /saude detectado.`, "INFO");
          enviarSaudeFinanceira(chatId, usuario);
          return;
      // --- FIM DO NOVO CASE ---
      // --- INÍCIO DA CORREÇÃO ---
      // Adicionados os cases para os novos comandos
      case "/importar":
          logToSheet(`Comando /importar detectado.`, "INFO");
          handleImportarCommand(chatId);
          return;

      case "/arquivar_agora":
          logToSheet(`Comando /arquivar_agora detectado.`, "INFO");
          handleArquivarAgoraCommand(chatId);
          return;
      // --- FIM DA CORREÇÃO ---

      case "/ajuda":
          logToSheet(`Comando /ajuda detectado.`, "INFO");
          enviarAjuda(chatId);
          return;
      // --- NOVOS COMANDOS E CORREÇÕES ---
      case "/orcamento":
          logToSheet(`Comando /orcamento detectado. Mes: ${targetMonth}, Ano: ${targetYear}`, "INFO");
          handleOrcamentoCommand(chatId, usuario, targetMonth, targetYear);
          return;

      case "/metas":
          logToSheet(`Comando /metas detectado.`, "INFO");
          handleMetasCommand(chatId);
          return;

      // --- NOVOS COMANDOS PARA GESTÃO DE METAS ---
      case "/novameta":
          logToSheet(`Comando /novameta detectado. Complemento: "${complemento}"`, "INFO");
          handleNovaMetaCommand(chatId, complemento);
          return;
          
      case "/aportarmeta":
          logToSheet(`Comando /aportarmeta detectado. Complemento: "${complemento}"`, "INFO");
          handleAportarMetaCommand(chatId, complemento, usuario);
          return;
      // --- FIM DOS NOVOS COMANDOS ---
      // --- NOVO COMANDO PARA PROVENTOS ---
      case "/provento":
          logToSheet(`Comando /provento detectado. Complemento: "${complemento}"`, "INFO");
          handleProventoCommand(chatId, complemento, usuario);
          return;
      // --- FIM DO NOVO COMANDO ---


      default:
        const palavrasConsulta = ["quanto", "qual", "quais", "listar", "mostrar", "total"];
        const primeiraPalavraConsulta = textoRecebido.toLowerCase().split(' ')[0];

        if (palavrasConsulta.includes(primeiraPalavraConsulta)) {
            logToSheet(`Consulta em linguagem natural detectada: "${textoRecebido}".`, "INFO");
            processarConsultaLinguagemNatural(chatId, usuario, textoRecebido);
            return;
        }

        logToSheet(`Comando '${comandoBase}' não reconhecido como comando direto. Tentando interpretar como lançamento.`, "INFO");
        const resultadoLancamento = interpretarMensagemTelegram(textoRecebido, usuario, chatId);

        if (resultadoLancamento && resultadoLancamento.handled) {
          logToSheet("Mensagem ja tratada pela lógica de investimentos.", "INFO");
        } else if (resultadoLancamento && resultadoLancamento.errorMessage) {
          enviarMensagemTelegram(chatId, `❌ ${resultadoLancamento.errorMessage}`);
          logToSheet(`Erro na interpretação da mensagem: ${resultadoLancamento.errorMessage}`, "WARN");
        } else if (resultadoLancamento && resultadoLancamento.status === 'PENDING_CONFIRMATION') {
          logToSheet(`Confirmacao de transacao pendente para ID: ${resultadoLancamento.transactionId}`, "INFO");
        } else if (resultadoLancamento && resultadoLancamento.status === 'PENDING_ASSISTANT_ACTION') {
          logToSheet(`Ação do assistente pendente para ID: ${resultadoLancamento.transactionId}`, "INFO");
        } else {
          enviarMensagemTelegram(chatId, "❌ Não entendi seu comando ou lançamento. Digite /ajuda para ver o que posso fazer.");
        }
        return;
    }
  } catch (err) {
    const chatIdForError = e?.postData?.contents ? JSON.parse(e.postData.contents)?.message?.chat?.id || JSON.parse(e.postData.contents)?.callback_query?.message?.chat?.id : null;
    logToSheet(`ERRO FATAL E INESPERADO EM doPost: ${err.message}. Stack: ${err.stack}`, "ERROR");
    if (chatIdForError) {
        enviarMensagemTelegram(chatIdForError, "❌ Ocorreu um erro crítico no sistema. O administrador foi notificado. Por favor, tente novamente mais tarde.");
    }
  } finally {
    // --- MELHORIA DE PERFORMANCE DE LOGS ---
    // Esta linha garante que todos os logs acumulados durante a execução
    // sejam escritos na planilha de uma só vez, no final.
    flushLogs();
    // --- FIM DA MELHORIA ---
  }
}

// ===================================================================================
// ### INÍCIO DAS NOVAS FUNÇÕES DE GESTÃO DE INVESTIMENTOS (PROVENTOS) ###
// ===================================================================================

/**
 * Lida com o comando /provento para registrar dividendos ou outros rendimentos de ativos.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} complemento O texto que segue o comando.
 * @param {string} usuario O nome do usuário que registrou o provento.
 */
function handleProventoCommand(chatId, complemento, usuario) {
  try {
    // Formato esperado: /provento TICKER VALOR para CONTA_DESTINO
    const match = complemento.match(/([a-zA-Z0-9]+)\s+([\d.,]+)\s+(?:para|em|na)\s+(.+)/i);
    if (!match) {
      enviarMensagemTelegram(chatId, "❌ Formato inválido. Use: `/provento TICKER VALOR para CONTA`\nEx: `/provento ITSA4 50 para NuInvest`");
      return;
    }

    const [, ticker, valor, contaDestino] = match;
    const valorProvento = parseBrazilianFloat(valor);

    // Chama a função de lógica principal no arquivo de Investimentos
    registrarProvento(chatId, ticker, valorProvento, contaDestino, usuario);

  } catch (e) {
    handleError(e, "handleProventoCommand", chatId);
  }
}

// ===================================================================================
// ### FIM DAS NOVAS FUNÇÕES DE GESTÃO DE INVESTIMENTOS (PROVENTOS) ###
// ===================================================================================


/**
 * Função executada quando a planilha é aberta. Cria o menu do Add-on.
 * @param {Object} e O objeto de evento de abertura.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  try {
    const props = PropertiesService.getScriptProperties();
    const systemStatus = props.getProperty('SYSTEM_STATUS');

    const menu = ui.createMenu('Boas Contas'); // Nome do Add-on no menu

    if (!isLicenseValid()) {
      menu.addItem('⚠️ Ativar Produto', 'activateProduct');
    } else if (systemStatus !== 'INITIALIZED') {
      menu.addItem('🚀 Inicializar Sistema', 'initializeSystem');
    } else {
      // Menu completo para sistema ativado e inicializado
      menu.addItem('📊 Abrir Dashboard', 'showDashboard'); // ITEM ADICIONADO
      menu.addItem('⚙️ Configurações', 'showConfigurationSidebar');
      menu.addItem('🤖 Configuração do Bot (Telegram)', 'showSetupUI');
      menu.addSeparator();
      menu.addItem('🔄 Atualizar Menu do Bot', 'setTelegramMenu');
      menu.addSeparator();
      menu.addItem('📖 Guia de Comandos', 'showCommandsGuide');
      menu.addSeparator();
      menu.addItem('✅ Verificação do Sistema', 'runSystemDiagnostics');
      menu.addItem('💰 Atualizar Orçamento', 'updateBudgetSpentValues');
      menu.addSeparator();
      menu.addItem('🔧 Re-sincronizar Saldos', 'reconciliarSaldosManualmente');
      menu.addItem('🗓️ Gerar Contas Recorrentes', 'triggerGenerateRecurringBills');
    }
    menu.addToUi();
  } catch (e) {
    handleError(e, "onOpen");
    ui.createMenu('Boas Contas')
      .addItem('⚠️ Ativar Produto', 'activateProduct')
      .addToUi();
  }
}

/**
 * Mostra o guia de comandos interativo em um diálogo modal.
 */
function showCommandsGuide() {
  // Usa createTemplateFromFile para processar o include do CSS
  const template = HtmlService.createTemplateFromFile('Infográfico Interativo Gasto Certo');
  const html = template.evaluate()
      .setWidth(900)  // Largura maior para melhor visualização
      .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Guia de Comandos do Gasto Certo');
}

/**
 * Retorna o conteúdo HTML do Guia de Comandos para ser exibido no modal do dashboard.
 */
function getCommandsGuideHtml() {
  // Usa createTemplateFromFile para processar o include do CSS
  const template = HtmlService.createTemplateFromFile('Infográfico Interativo Gasto Certo');
  return template.evaluate().getContent();
}

/**
 * Busca as tarefas da planilha e as formata para o calendário.
 * VERSÃO MELHORADA: Garante que as datas são válidas e adiciona logs de erro.
 */
function getCalendarTasks() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tarefas');
    if (!sheet) {
      Logger.log('Aba "Tarefas" não encontrada.');
      return { success: false, message: 'Aba "Tarefas" não encontrada.' };
    }
    
    const data = sheet.getDataRange().getValues();
    const tasks = [];
    
    // Começa em 1 para ignorar o cabeçalho
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // Colunas: ID, Descrição, Data, Status
      const id = row[0];
      const title = row[1];
      const dateValue = row[2];
      const status = row[3];

      // Validação para garantir que a data não está vazia e é válida
      if (title && dateValue && !isNaN(new Date(dateValue).getTime())) {
        const task = {
          id: id,
          title: title,
          start: new Date(dateValue).toISOString().split('T')[0], // Formato YYYY-MM-DD
          allDay: true,
          backgroundColor: status === 'Concluída' ? 'var(--accent-green)' : 'var(--accent-purple)',
          borderColor: status === 'Concluída' ? 'var(--accent-green)' : 'var(--accent-purple)'
        };
        tasks.push(task);
      } else {
        Logger.log(`Linha ${i + 1} ignorada: Título ou data inválida. Título: "${title}", Data: "${dateValue}"`);
      }
    }
    
    Logger.log(`Encontradas ${tasks.length} tarefas válidas.`);
    return { success: true, tasks: tasks };
    
  } catch (e) {
    Logger.log('Erro ao buscar tarefas: ' + e.stack);
    return { success: false, message: 'Erro ao buscar tarefas: ' + e.message };
  }
}


function showSetupUI() {
  // 1. Crie um template a partir do arquivo
  const template = HtmlService.createTemplateFromFile('SetupDialog');
  
  // 2. Avalie o template para processar os comandos internos (como o 'include')
  const html = template.evaluate()
      .setWidth(600) // Aumentei a largura para o novo layout
      .setHeight(680); // Ajustei a altura

  // 3. Mostre o resultado no diálogo modal
  SpreadsheetApp.getUi().showModalDialog(html, 'Configuração do Bot');
}

/**
 * **FUNÇÃO ATUALIZADA**
 * Agora também salva a chave da API de Speech-to-Text.
 */
function saveCredentialsAndSetupWebhook(token, chatId, webAppUrl, speechApiKey) { // <-- PARÂMETRO ADICIONADO
  try {
    if (!token || !chatId || !webAppUrl) {
      throw new Error("O Token, o ID do Chat e a URL do App da Web são obrigatórios.");
    }

    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperties({
      [TELEGRAM_TOKEN_PROPERTY_KEY]: token,
      [ADMIN_CHAT_ID_PROPERTY_KEY]: chatId,
      [WEB_APP_URL_PROPERTY_KEY]: webAppUrl,
      [SPEECH_API_KEY_PROPERTY_KEY]: speechApiKey || "" // <-- ADICIONADO (guarda vazio se não for fornecido)
    });
    
    logToSheet("Configurações de Token, Chat ID, URL e Chave de API salvas com sucesso.", "INFO");

    updateAdminConfig(chatId);

    const webhookResult = setupWebhook();

    if (webhookResult && webhookResult.ok) {
        logToSheet("Configuração do Webhook concluída com sucesso.", "INFO");
        
        const state = { step: SETUP_STEPS.PENDING_START, data: {} };
        setGuidedSetupState(chatId, state);
        logToSheet(`[Onboarding] Estado inicial PENDING_START definido para ${chatId}.`, "INFO");

        return { success: true, message: "Conectado com sucesso! Agora, vá ao seu bot no Telegram e envie a mensagem /start para iniciar a configuração." };
    } else {
        const errorDescription = webhookResult ? webhookResult.description : "Resposta inválida da API do Telegram.";
        throw new Error(`Falha ao configurar o webhook: ${errorDescription}`);
    }

  } catch (e) {
    logToSheet(`Erro durante a configuração: ${e.message}`, "ERROR");
    return { success: false, message: e.message };
  }
}



/**
 * Função para configurar o webhook do Telegram.
 * Agora lê a URL do Web App diretamente das Propriedades do Script, que é mais confiável.
 * @returns {Object} Um objeto com o resultado da API do Telegram.
 */
function setupWebhook() {
  try {
    const token = getTelegramBotToken();
    // A URL é lida das propriedades, onde foi salva pela caixa de diálogo.
    const webhookUrl = PropertiesService.getScriptProperties().getProperty(WEB_APP_URL_PROPERTY_KEY);

    if (!webhookUrl) {
      const errorMessage = "URL do Web App não encontrada nas Propriedades do Script. Execute a 'Configuração Inicial' e forneça a URL correta.";
      throw new Error(errorMessage);
    }

    const url = `https://api.telegram.org/bot${token}/setWebhook?url=${webhookUrl}`;
    
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const responseText = response.getContentText();
    logToSheet(`Resposta da configuração do webhook: ${responseText}`, "INFO");
    return JSON.parse(responseText);

  } catch (e) {
    logToSheet(`Erro ao configurar o webhook: ${e.message}`, "ERROR");
    return { ok: false, description: e.message };
  }
}


/**
 * **NOVA FUNÇÃO**
 * Adiciona ou atualiza a configuração do utilizador administrador na aba 'Configuracoes'.
 * @param {string} adminChatId O Chat ID do administrador.
 */
function updateAdminConfig(adminChatId) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("Configuracoes");
    if (!configSheet) {
        logToSheet("Aba 'Configuracoes' não encontrada para atualizar o admin.", "ERROR");
        return;
    }
    const data = configSheet.getDataRange().getValues();
    let adminRowFound = false;

    // Procura pela primeira linha com um chatId de exemplo para atualizar
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === 'chatId' && (data[i][1] === '123456789' || data[i][2] === 'Pessoa 1')) {
            configSheet.getRange(i + 1, 2).setValue(adminChatId);
            configSheet.getRange(i + 1, 3).setValue('Admin'); // Define o nome como 'Admin'
            adminRowFound = true;
            logToSheet(`Chat ID do admin atualizado para ${adminChatId} na linha ${i + 1}.`, "INFO");
            break;
        }
    }
    
    // Se não encontrou, adiciona uma nova linha (fallback)
    if (!adminRowFound) {
        configSheet.appendRow(['chatId', adminChatId, 'Admin', 'Família']);
        logToSheet(`Nenhum admin de exemplo encontrado. Nova linha de admin adicionada para ${adminChatId}.`, "INFO");
    }
    // Limpa o cache para garantir que os dados atualizados sejam lidos
    CacheService.getScriptCache().remove(CACHE_KEY_CONFIG);
}


/**
 * Adiciona um novo usuário ao sistema.
 * @param {string} chatId O ID do chat do novo usuário.
 * @param {string} userName O nome do usuário.
 */
function addNewUser(chatId, userName) {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEETS.USERS);
  if (sheet) {
    // Verifica se o usuário já existe
    const existingUser = findRowByValue(SHEETS.USERS, 1, chatId);
    if (!existingUser) {
      sheet.appendRow([chatId, userName, new Date()]);
      Logger.log(`Novo usuário adicionado: ${userName} (${chatId})`);
    }
  }
}

/**
 * Inicializa todas as abas necessárias da planilha com base no objeto HEADERS.
 * Garante que o ambiente do usuário seja criado corretamente.
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Itera sobre o objeto HEADERS para criar cada aba com seus respectivos cabeçalhos.
  for (const sheetName in HEADERS) {
    if (Object.prototype.hasOwnProperty.call(HEADERS, sheetName)) {
      if (!ss.getSheetByName(sheetName)) {
        const sheet = ss.insertSheet(sheetName);
        const headers = HEADERS[sheetName];
        if (headers && headers.length > 0) {
          sheet.appendRow(headers);
          logToSheet(`Aba '${sheetName}' criada com sucesso.`, "INFO");
        }
      }
    }
  }
  
  // Garante que a aba de logs também seja criada.
  if (!ss.getSheetByName(SHEET_LOGS_SISTEMA)) {
      const logSheet = ss.insertSheet(SHEET_LOGS_SISTEMA);
      logSheet.appendRow(["timestamp", "level", "message"]);
      logToSheet(`Aba de sistema '${SHEET_LOGS_SISTEMA}' criada com sucesso.`, "INFO");
  }
}

function reconciliarSaldosManualmente() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Iniciando re-sincronização completa dos saldos... Isso pode levar um momento.', 'Manutenção', 30);
  atualizarSaldosDasContas(); // Chama a função original e completa
  SpreadsheetApp.getUi().alert('Sucesso!', 'Os saldos de todas as contas foram recalculados e sincronizados com sucesso.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// Em Code.gs
function exibirDashboardModerno() {
  var html = HtmlService.createHtmlOutputFromFile('Dashboard_Moderno')
      .setTitle('Boas Contas - Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showSidebar(html);
}


