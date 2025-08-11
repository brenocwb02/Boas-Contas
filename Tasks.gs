/**
 * @file Tasks.gs
 * @description Módulo para gerenciamento de tarefas com integração ao Google Agenda.
 */

const SHEET_TAREFAS = "Tarefas";
const TASK_STATUS_PENDING = "Pendente";
const TASK_STATUS_COMPLETED = "Concluída";

// ===================================================================================
// FUNÇÕES PRINCIPAIS (CHAMADAS PELO doPost)
// ===================================================================================

/**
 * Ponto de entrada para criar uma nova tarefa a partir de uma mensagem do Telegram.
 * Interpreta a linguagem natural para extrair descrição e data.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} textoCompleto O texto enviado pelo usuário (ex: "Reunião com cliente amanhã às 15h").
 */
function criarNovaTarefa(chatId, textoCompleto) {
  // 1. Tenta extrair data e hora do texto
  const { data, textoRestante } = extrairDataEHora(textoCompleto);

  if (!textoRestante) {
    enviarMensagemTelegram(chatId, "❌ Não consegui entender a descrição da tarefa. Tente novamente, por exemplo: `/tarefa Comprar pão amanhã de manhã`");
    return;
  }

  const descricao = textoRestante;
  const dataConclusao = data;

  // 2. Adiciona a tarefa na planilha
  const novaTarefa = adicionarTarefaNaPlanilha(chatId, descricao, dataConclusao);

  // 3. Envia mensagem de confirmação com opção de adicionar à agenda
  let mensagem = `✅ Tarefa criada com sucesso!\n\n` +
                 `*ID:* \`${novaTarefa.id}\`\n` +
                 `*Descrição:* ${escapeMarkdown(novaTarefa.descricao)}\n` +
                 `*Prazo:* ${novaTarefa.dataConclusao ? Utilities.formatDate(novaTarefa.dataConclusao, Session.getScriptTimeZone(), "dd/MM/yyyy 'às' HH:mm") : "Sem prazo definido"}`;

  const teclado = {
    inline_keyboard: [
      [{ text: "🗓️ Adicionar ao Google Agenda", callback_data: `add_agenda_${novaTarefa.id}` }],
      [{ text: "✅ Concluir Tarefa", callback_data: `concluir_tarefa_${novaTarefa.id}` }]
    ]
  };

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
  logToSheet(`Tarefa criada para ${chatId}: ${descricao}`, "INFO");
}

/**
 * Lista todas as tarefas pendentes do usuário de forma organizada e visual.
 * @param {string} chatId O ID do chat do Telegram.
 */
function listarTarefasPendentes(chatId) {
  const todasAsTarefas = getTarefasDoUsuario(chatId, TASK_STATUS_PENDING);

  if (todasAsTarefas.length === 0) {
    enviarMensagemTelegram(chatId, "🎉 Você não tem nenhuma tarefa pendente!");
    return;
  }

  // Ordena as tarefas: primeiro as com data (mais próximas primeiro), depois as sem data
  todasAsTarefas.sort((a, b) => {
    if (a.dataConclusao && b.dataConclusao) {
      return new Date(a.dataConclusao) - new Date(b.dataConclusao);
    }
    if (a.dataConclusao) return -1; // 'a' tem data, 'b' não tem, 'a' vem primeiro
    if (b.dataConclusao) return 1;  // 'b' tem data, 'a' não tem, 'b' vem primeiro
    return 0; // Ambas sem data, mantém a ordem
  });

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  const tarefasVencidas = [];
  const tarefasHoje = [];
  const tarefasAmanha = [];
  const tarefasProximos7Dias = [];
  const tarefasFuturas = [];
  const tarefasSemPrazo = [];

  todasAsTarefas.forEach(tarefa => {
    if (!tarefa.dataConclusao) {
      tarefasSemPrazo.push(tarefa);
      return;
    }
    
    const dataTarefa = new Date(tarefa.dataConclusao);
    dataTarefa.setHours(0, 0, 0, 0);

    const diffTime = dataTarefa.getTime() - hoje.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    if (diffDays < 0) {
      tarefasVencidas.push(tarefa);
    } else if (diffDays === 0) {
      tarefasHoje.push(tarefa);
    } else if (diffDays === 1) {
      tarefasAmanha.push(tarefa);
    } else if (diffDays <= 7) {
      tarefasProximos7Dias.push(tarefa);
    } else {
      tarefasFuturas.push(tarefa);
    }
  });

  let mensagem = "📝 *Suas tarefas pendentes:*\n";
  const teclado = { inline_keyboard: [] };
  
  // Agrupa as tarefas mais urgentes para criar botões para elas (limite de 7)
  const tarefasParaAcoes = [...tarefasVencidas, ...tarefasHoje, ...tarefasAmanha].slice(0, 7);

  // Função auxiliar para formatar uma secção de tarefas (APENAS TEXTO)
  const formatarSecao = (titulo, tarefas, emoji) => {
    if (tarefas.length === 0) return "";
    let secaoStr = `\n*${titulo}*\n`;
    tarefas.forEach(t => {
      const prazo = t.dataConclusao ? Utilities.formatDate(new Date(t.dataConclusao), Session.getScriptTimeZone(), "dd/MM") : "";
      secaoStr += `${emoji} \`${t.id}\` - ${escapeMarkdown(t.descricao)} ${prazo ? `_(${prazo})_` : ''}\n`;
    });
    return secaoStr;
  };

  mensagem += formatarSecao("🔴 Vencidas", tarefasVencidas, "❗️");
  mensagem += formatarSecao("🟢 Para Hoje", tarefasHoje, "➡️");
  mensagem += formatarSecao("🔵 Para Amanhã", tarefasAmanha, "▶️");
  mensagem += formatarSecao("🗓️ Próximos 7 Dias", tarefasProximos7Dias, "▪️");
  mensagem += formatarSecao("📅 Futuras", tarefasFuturas, "▪️");
  mensagem += formatarSecao("🗂️ Sem Prazo Definido", tarefasSemPrazo, "▪️");
  
  // Adiciona os botões de ação rápida ao teclado
  if (tarefasParaAcoes.length > 0) {
    mensagem += "\n*👇 Ações Rápidas:*";
    tarefasParaAcoes.forEach(t => {
      const descAbreviada = t.descricao.length > 25 ? t.descricao.substring(0, 22) + '...' : t.descricao;
      teclado.inline_keyboard.push([
        { text: `✅ ${descAbreviada}`, callback_data: `concluir_tarefa_${t.id}` },
        { text: `🗑️`, callback_data: `excluir_tarefa_${t.id}` }
      ]);
    });
  } else {
    mensagem += "\n_Use `/concluir <ID>` para marcar uma tarefa como concluída._";
  }

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
}


/**
 * Marca uma tarefa como concluída.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} idTarefa O ID da tarefa a ser concluída.
 */
function concluirTarefa(chatId, idTarefa) {
  const tarefa = encontrarTarefaPorId(idTarefa);

  if (!tarefa || tarefa.chatId !== chatId.toString()) {
    enviarMensagemTelegram(chatId, `❌ Tarefa com ID \`${escapeMarkdown(idTarefa)}\` não encontrada ou não pertence a você.`);
    return;
  }

  if (tarefa.status === TASK_STATUS_COMPLETED) {
    enviarMensagemTelegram(chatId, `ℹ️ A tarefa "${escapeMarkdown(tarefa.descricao)}" já estava concluída.`);
    return;
  }

  // Se houver um evento na agenda, remove
  if (tarefa.idEventoAgenda) {
    removerEventoDaAgenda(tarefa.idEventoAgenda);
  }

  // Atualiza o status na planilha
  atualizarStatusDaTarefa(tarefa.linha, TASK_STATUS_COMPLETED);

  enviarMensagemTelegram(chatId, `✅ Tarefa "${escapeMarkdown(tarefa.descricao)}" marcada como concluída!`);
  logToSheet(`Tarefa ${idTarefa} concluída por ${chatId}.`, "INFO");
}

/**
 * Exclui uma tarefa permanentemente.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} idTarefa O ID da tarefa a ser excluída.
 */
function excluirTarefa(chatId, idTarefa) {
  const tarefa = encontrarTarefaPorId(idTarefa);

  if (!tarefa || tarefa.chatId !== chatId.toString()) {
    enviarMensagemTelegram(chatId, `❌ Tarefa com ID \`${escapeMarkdown(idTarefa)}\` não encontrada ou não pertence a você.`);
    return;
  }

  // Se houver um evento na agenda, remove
  if (tarefa.idEventoAgenda) {
    removerEventoDaAgenda(tarefa.idEventoAgenda);
  }

  // Exclui a linha da planilha
  excluirLinhaDaTarefa(tarefa.linha);

  enviarMensagemTelegram(chatId, `🗑️ Tarefa "${escapeMarkdown(tarefa.descricao)}" excluída com sucesso.`);
  logToSheet(`Tarefa ${idTarefa} excluída por ${chatId}.`, "INFO");
}


// ===================================================================================
// INTEGRAÇÃO COM GOOGLE AGENDA
// ===================================================================================

/**
 * Adiciona uma tarefa ao Google Agenda do usuário.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} idTarefa O ID da tarefa.
 */
function adicionarTarefaNaAgenda(chatId, idTarefa) {
  const tarefa = encontrarTarefaPorId(idTarefa);

  if (!tarefa || tarefa.chatId !== chatId.toString()) {
    enviarMensagemTelegram(chatId, `❌ Tarefa com ID \`${escapeMarkdown(idTarefa)}\` não encontrada.`);
    return;
  }

  if (tarefa.idEventoAgenda) {
    enviarMensagemTelegram(chatId, "ℹ️ Esta tarefa já está na sua agenda.");
    return;
  }

  if (!tarefa.dataConclusao) {
    enviarMensagemTelegram(chatId, "❌ Não é possível adicionar à agenda uma tarefa sem prazo definido.");
    return;
  }

  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const evento = calendar.createEvent(
      tarefa.descricao,
      new Date(tarefa.dataConclusao),
      new Date(new Date(tarefa.dataConclusao).getTime() + 60 * 60 * 1000) // Duração de 1 hora
    );
    
    const idEvento = evento.getId();
    atualizarIdEventoDaTarefa(tarefa.linha, idEvento);

    enviarMensagemTelegram(chatId, `✅ Tarefa "${escapeMarkdown(tarefa.descricao)}" adicionada à sua Google Agenda!`);
    logToSheet(`Evento ${idEvento} criado para a tarefa ${idTarefa}.`, "INFO");
  } catch (e) {
    enviarMensagemTelegram(chatId, "❌ Ocorreu um erro ao adicionar à agenda. Você precisa autorizar o script a acessar sua agenda. Tente novamente.");
    logToSheet(`Erro ao criar evento na agenda: ${e.message}`, "ERROR");
  }
}

/**
 * Remove um evento do Google Agenda.
 * @param {string} idEventoAgenda O ID do evento a ser removido.
 */
function removerEventoDaAgenda(idEventoAgenda) {
  try {
    const evento = CalendarApp.getEventById(idEventoAgenda);
    if (evento) {
      evento.deleteEvent();
      logToSheet(`Evento ${idEventoAgenda} removido da agenda.`, "INFO");
    }
  } catch (e) {
    logToSheet(`Erro ao remover evento ${idEventoAgenda} da agenda: ${e.message}`, "WARN");
  }
}

// ===================================================================================
// FUNÇÕES DE ALERTA (PARA SEREM CHAMADAS POR UM GATILHO/TRIGGER)
// ===================================================================================

/**
 * Verifica tarefas que vencem em breve e envia lembretes.
 * Deve ser configurado para rodar diariamente através de um gatilho de tempo.
 */
function enviarLembretesDeTarefas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TAREFAS);
  if (!sheet) return; // Se a aba não existir, não faz nada
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colMap = getColumnMap(headers);

  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const status = row[colMap["Status"]];
    const dataConclusaoRaw = row[colMap["DataConclusao"]];
    
    if (status === TASK_STATUS_PENDING && dataConclusaoRaw) {
      const dataConclusao = new Date(dataConclusaoRaw);
      dataConclusao.setHours(0, 0, 0, 0);

      const diffTime = dataConclusao.getTime() - hoje.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      // Envia lembrete para tarefas que vencem amanhã (diffDays === 1)
      if (diffDays === 1) {
        const chatId = row[colMap["ChatIDUsuario"]];
        const descricao = row[colMap["Descricao"]];
        const mensagem = `🔔 *Lembrete de Tarefa para Amanhã:*\n\n${escapeMarkdown(descricao)}`;
        enviarMensagemTelegram(chatId, mensagem);
        logToSheet(`Lembrete de tarefa enviado para ${chatId} para a tarefa: ${descricao}`, "INFO");
      }
    }
  }
}


// ===================================================================================
// FUNÇÕES AUXILIARES (INTERAÇÃO COM A PLANILHA E PARSING)
// ===================================================================================

/**
 * Adiciona uma nova linha na aba "Tarefas".
 * @returns {object} O objeto da tarefa criada.
 */
function adicionarTarefaNaPlanilha(chatId, descricao, dataConclusao) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TAREFAS);
  const id = Utilities.getUuid().substring(0, 6);
  const dataCriacao = new Date();

  const newRow = [
    id,
    descricao,
    dataCriacao,
    dataConclusao || "", // Deixa em branco se não houver data
    TASK_STATUS_PENDING,
    "", // ID Evento Agenda
    chatId
  ];
  sheet.appendRow(newRow);

  // Força a escrita dos dados na planilha imediatamente
  SpreadsheetApp.flush();

  return {
    id: id,
    descricao: descricao,
    dataConclusao: dataConclusao
  };
}

/**
 * Encontra uma tarefa na planilha pelo seu ID.
 * @returns {object|null} O objeto da tarefa ou null se não encontrada.
 */
function encontrarTarefaPorId(idTarefa) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TAREFAS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colMap = getColumnMap(headers);

  for (let i = 1; i < data.length; i++) {
    // Compara como string para evitar problemas de tipo
    if (data[i][colMap["ID"]].toString() == idTarefa.toString()) { 
      return {
        linha: i + 1,
        id: data[i][colMap["ID"]],
        descricao: data[i][colMap["Descricao"]],
        dataConclusao: data[i][colMap["DataConclusao"]] ? new Date(data[i][colMap["DataConclusao"]]) : null,
        status: data[i][colMap["Status"]],
        idEventoAgenda: data[i][colMap["IDEventoAgenda"]],
        chatId: data[i][colMap["ChatIDUsuario"]].toString()
      };
    }
  }
  return null;
}

/**
 * Obtém todas as tarefas de um usuário com um determinado status.
 * @returns {Array<object>} Um array de objetos de tarefa.
 */
function getTarefasDoUsuario(chatId, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TAREFAS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colMap = getColumnMap(headers);
  const tarefas = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][colMap["ChatIDUsuario"]] == chatId && data[i][colMap["Status"]] === status) {
      tarefas.push({
        id: data[i][colMap["ID"]],
        descricao: data[i][colMap["Descricao"]],
        dataConclusao: data[i][colMap["DataConclusao"]]
      });
    }
  }
  return tarefas;
}

function atualizarStatusDaTarefa(linha, novoStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TAREFAS);
  const colIndex = getColumnMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0])["Status"] + 1;
  sheet.getRange(linha, colIndex).setValue(novoStatus);
}

function atualizarIdEventoDaTarefa(linha, idEvento) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TAREFAS);
  const colIndex = getColumnMap(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0])["IDEventoAgenda"] + 1;
  sheet.getRange(linha, colIndex).setValue(idEvento);
}

function excluirLinhaDaTarefa(linha) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TAREFAS);
  sheet.deleteRow(linha);
}

/**
 * Tenta extrair data e hora de uma string de texto.
 * @param {string} texto O texto de onde extrair a data.
 * @returns {object} Um objeto com a data encontrada e o texto restante.
 */
function extrairDataEHora(texto) {
    let data = null;
    let textoRestante = texto;
    let matchEncontrado = '';

    const hoje = new Date();
    const amanha = new Date();
    amanha.setDate(hoje.getDate() + 1);

    // Expressões regulares ordenadas da mais específica para a mais genérica
    const patterns = [
        // Padrão 1: Data completa com hora (ex: 14/08/2025 às 15:30)
        { regex: /(\d{1,2}\/\d{1,2}\/\d{4})\s*(?:em|às|as)?\s*(\d{1,2}):(\d{2})h?/i, handler: (match) => {
            const parts = match[1].split('/');
            return new Date(parts[2], parts[1] - 1, parts[0], match[2], match[3] || '0', 0);
        }},
        // Padrão 2: Data completa sem hora (ex: em 14/08/2025)
        { regex: /(?:em\s+)?(\d{1,2}\/\d{1,2}\/\d{4})/i, handler: (match) => {
            const parts = match[1].split('/');
            return new Date(parts[2], parts[1] - 1, parts[0], 9, 0, 0); // Predefinição para as 9h
        }},
        // Padrão 3: "Amanhã" com hora (ex: amanhã às 15h, amanha as 10:30)
        { regex: /amanh[aã]\s*(?:[àa]s)?\s*(\d{1,2})(?::(\d{2}))?h?/i, handler: (match) => {
            return new Date(amanha.getFullYear(), amanha.getMonth(), amanha.getDate(), match[1], match[2] || '0', 0);
        }},
        // Padrão 4: "Hoje" com hora (ex: hoje às 15h, hoje as 10:30)
        { regex: /hoje\s*(?:[àa]s)?\s*(\d{1,2})(?::(\d{2}))?h?/i, handler: (match) => {
            return new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate(), match[1], match[2] || '0', 0);
        }},
        // Padrão 5: Dia da semana COM HORA (ex: na sexta-feira às 11h)
        { regex: /(?:na|pr[oó]xima)?\s*(domingo|segunda|ter[çc]a|quarta|quinta|sexta|s[áa]bado)(?:-feira)?\s*(?:[àa]s)?\s*(\d{1,2})(?::(\d{2}))?h?/i, handler: (match) => {
            const dias = {domingo:0, segunda:1, terca:2, terça:2, quarta:3, quinta:4, sexta:5, sabado:6, sábado:6};
            const diaAlvo = dias[match[1].toLowerCase().replace('ç', 'c')];
            if (diaAlvo === undefined) return null;
            const diaAtual = hoje.getDay();
            let diasAAdicionar = diaAlvo - diaAtual;
            if (diasAAdicionar <= 0) {
                diasAAdicionar += 7;
            }
            const dataAlvo = new Date();
            dataAlvo.setDate(hoje.getDate() + diasAAdicionar);
            // Define a hora e o minuto a partir do match
            dataAlvo.setHours(match[2] || 9, match[3] || 0, 0, 0);
            return dataAlvo;
        }},
        // Padrão 6: Dia da semana SEM HORA (ex: na sexta-feira) - MENOS PRIORIDADE
        { regex: /(?:na|pr[oó]xima)?\s*(domingo|segunda|ter[çc]a|quarta|quinta|sexta|s[áa]bado)(?:-feira)?/i, handler: (match) => {
            const dias = {domingo:0, segunda:1, terca:2, terça:2, quarta:3, quinta:4, sexta:5, sabado:6, sábado:6};
            const diaAlvo = dias[match[1].toLowerCase().replace('ç', 'c')];
            if (diaAlvo === undefined) return null;
            const diaAtual = hoje.getDay();
            let diasAAdicionar = diaAlvo - diaAtual;
            if (diasAAdicionar <= 0) {
                diasAAdicionar += 7;
            }
            const dataAlvo = new Date();
            dataAlvo.setDate(hoje.getDate() + diasAAdicionar);
            dataAlvo.setHours(9, 0, 0, 0); // Predefinição para as 9h
            return dataAlvo;
        }},
        // Padrão 7: Períodos do dia (ex: hoje à noite, amanhã de manhã)
        { regex: /(hoje|amanh[aã])\s*(?:[àa])?\s*(manh[aã]|tarde|noite)/i, handler: (match) => {
            const baseDate = match[1].startsWith('amanh') ? amanha : hoje;
            let hora = 9; // manhã
            if (match[2] === 'tarde') hora = 14;
            if (match[2] === 'noite') hora = 20;
            return new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate(), hora, 0, 0);
        }}
    ];

    for (const pattern of patterns) {
        const match = texto.match(pattern.regex);
        if (match) {
            data = pattern.handler(match);
            matchEncontrado = match[0];
            break;
        }
    }

    if (matchEncontrado) {
        textoRestante = texto.replace(matchEncontrado, "").trim();
    }

    // Limpa preposições e artigos comuns que podem sobrar no início ou fim
    textoRestante = textoRestante.replace(/^(de|da|do|para|com|em)\s+/i, '').trim();
    textoRestante = textoRestante.replace(/\s+(de|da|do|para|com|em)$/i, '').trim();

    // Capitaliza a primeira letra
    if (textoRestante) {
        textoRestante = textoRestante.charAt(0).toUpperCase() + textoRestante.slice(1);
    }

    return { data: data, textoRestante: textoRestante };
}
