/**
 * @file Commands.gs
 * @description Este arquivo contém as implementações de todos os comandos do bot do Telegram.
 * Cada função aqui corresponde a um comando específico (/resumo, /saldo, etc.).
 * VERSÃO COM A FUNÇÃO enviarSaldo REFATORADA.
 */

// Variável global para armazenar os saldos calculados.
// Usar `globalThis` é uma boa prática para garantir que ela seja acessível em diferentes arquivos .gs.
// É populada pela função `atualizarSaldosDasContas` em FinancialLogic.gs.
globalThis.saldosCalculados = {};

/**
 * Gera uma mensagem de resumo financeiro mensal, incluindo receitas, despesas, saldo e gastos por categoria/cartão.
 * Inclui também o progresso das metas.
 * @param {number} mes O mês para o resumo (1-12).
 * @param {number} ano O ano para o resumo.
 * @returns {string} A mensagem formatada de resumo financeiro.
 */
function gerarResumoMensal(mes, ano) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoes = ss.getSheetByName(SHEET_TRANSACOES).getDataRange().getValues();
  const metasSheet = ss.getSheetByName(SHEET_METAS).getDataRange().getValues();
  const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);

  logToSheet(`Inicio de gerarResumoMensal para ${mes}/${ano}`, "INFO");

  const mesIndex = mes - 1;
  const nomeMes = getNomeMes(mesIndex);

  let resumoCategorias = {};
  let resumoCartoes = {};
  let metasPorCategoria = {};
  let totalReceitasMes = 0;
  let totalDespesasMesExcluindoPagamentosETransferencias = 0;

  // --- Processamento de Metas ---
  const cabecalhoMetas = metasSheet[2];
  let colMetaMes = -1;

  for (let i = 2; i < cabecalhoMetas.length; i++) {
    const headerValue = String(cabecalhoMetas[i]).toLowerCase();
    const targetHeader = `${nomeMes.toLowerCase()}/${ano}`;
    if (headerValue.includes(targetHeader)) {
      colMetaMes = i;
      break;
    }
  }

  if (colMetaMes !== -1) {
    for (let i = 3; i < metasSheet.length; i++) {
      const categoriaMeta = (metasSheet[i][0] || "").toString().trim();
      const subcategoriaMeta = (metasSheet[i][1] || "").toString().trim();
      const valorMetaTexto = metasSheet[i][colMetaMes];

      if (categoriaMeta && subcategoriaMeta && valorMetaTexto) {
        const meta = parseBrazilianFloat(String(valorMetaTexto));
        if (!isNaN(meta) && meta > 0) {
          const chaveCategoria = normalizarTexto(categoriaMeta);
          const chaveSubcategoria = normalizarTexto(`${categoriaMeta} ${subcategoriaMeta}`);
          if (!metasPorCategoria[chaveCategoria]) {
            metasPorCategoria[chaveCategoria] = { totalMeta: 0, totalGasto: 0, subcategories: {} };
          }
          metasPorCategoria[chaveCategoria].subcategories[chaveSubcategoria] = { meta: meta, gasto: 0 };
          metasPorCategoria[chaveCategoria].totalMeta += meta;
        }
      }
    }
  }

  // --- LÓGICA UNIFICADA: Calcular Despesas, Receitas e Gastos de Metas ---
  for (let i = 1; i < transacoes.length; i++) {
    const dataRaw = transacoes[i][0];
    const tipo = transacoes[i][4];
    let valor = parseBrazilianFloat(String(transacoes[i][5]));
    const categoria = transacoes[i][2];
    const subcategoria = transacoes[i][3];
    const conta = transacoes[i][7];
    const dataVencimentoRaw = transacoes[i][10];

    // Se for RECEITA, filtra pela DATA DA TRANSAÇÃO
    if (tipo === "Receita") {
      const data = parseData(dataRaw);
      if (data && data.getMonth() === mesIndex && data.getFullYear() === ano) {
        const categoriaNormalizada = normalizarTexto(categoria);
        const subcategoriaNormalizada = normalizarTexto(subcategoria);
        if (!(categoriaNormalizada === "transferencias" && subcategoriaNormalizada === "entre contas")) {
            totalReceitasMes += valor;
        }
      }
    }
    // Se for DESPESA, filtra pela DATA DE VENCIMENTO
    else if (tipo === "Despesa") {
      const dataVencimento = parseData(dataVencimentoRaw);
      if (dataVencimento && dataVencimento.getMonth() === mesIndex && dataVencimento.getFullYear() === ano) {
        const categoriaNormalizada = normalizarTexto(categoria);
        const subcategoriaNormalizada = normalizarTexto(subcategoria);

        // Lógica para Fluxo de Caixa e Resumo de Categorias
        if (
            !(categoriaNormalizada === "contas a pagar" && subcategoriaNormalizada === "pagamento de fatura") &&
            !(categoriaNormalizada === "transferencias" && subcategoriaNormalizada === "entre contas")
        ) {
            totalDespesasMesExcluindoPagamentosETransferencias += valor;

            if (!resumoCategorias[categoria]) {
              resumoCategorias[categoria] = { total: 0, subcategories: {} };
            }
            resumoCategorias[categoria].total += valor;
            if (!resumoCategorias[categoria].subcategories[subcategoria]) {
              resumoCategorias[categoria].subcategories[subcategoria] = 0;
            }
            resumoCategorias[categoria].subcategories[subcategoria] += valor;
            
            // Soma nos gastos da meta (APENAS UMA VEZ)
            const chaveCategoriaMeta = normalizarTexto(categoria);
            const chaveSubcategoriaMeta = normalizarTexto(`${categoria} ${subcategoria}`);
            if (metasPorCategoria[chaveCategoriaMeta] && metasPorCategoria[chaveCategoriaMeta].subcategories[chaveSubcategoriaMeta]) {
              metasPorCategoria[chaveCategoriaMeta].subcategories[chaveSubcategoriaMeta].gasto += valor;
              metasPorCategoria[chaveCategoriaMeta].totalGasto += valor;
            }
        }
        
        // Lógica para Fatura Bruta do Cartão
        const infoConta = obterInformacoesDaConta(conta, dadosContas);
        if (infoConta && normalizarTexto(infoConta.tipo) === "cartao de credito") {
            const nomeCartaoResumoNormalizado = infoConta.contaPaiAgrupador || infoConta.nomeNormalizado; 
            if (!resumoCartoes[nomeCartaoResumoNormalizado]) {
                resumoCartoes[nomeCartaoResumoNormalizado] = { 
                    faturaBrutaMes: 0,
                    vencimento: infoConta.vencimento, 
                    limite: infoConta.limite,
                    nomeOriginalParaExibicao: infoConta.nomeOriginal 
                };
            }
            resumoCartoes[nomeCartaoResumoNormalizado].faturaBrutaMes += valor;
        }
      }
    }
  }

  // --- Construção da Mensagem de Resumo ---
  let mensagemResumo = `📊 *Resumo Financeiro de ${nomeMes}/${ano}*\n\n`;

  mensagemResumo += `*💰 Fluxo de Caixa do Mes*\n`;
  mensagemResumo += `• *Receitas Totais:* R$ ${totalReceitasMes.toFixed(2).replace('.', ',')}\n`;
  mensagemResumo += `• *Despesas Totais (excluindo pagamentos de fatura e transferencias):* R$ ${totalDespesasMesExcluindoPagamentosETransferencias.toFixed(2).replace('.', ',')}\n`;
  const saldoLiquidoMes = totalReceitasMes - totalDespesasMesExcluindoPagamentosETransferencias;
  let emojiSaldo = saldoLiquidoMes > 0 ? "✅" : (saldoLiquidoMes < 0 ? "❌" : "⚖️");
  mensagemResumo += `• *Saldo Liquido do Mes:* ${emojiSaldo} R$ ${saldoLiquidoMes.toFixed(2).replace('.', ',')}\n\n`;

  mensagemResumo += `*💸 Despesas Detalhadas por Categoria*\n`;
  const categoriasOrdenadas = Object.keys(resumoCategorias).sort((a, b) => resumoCategorias[b].total - resumoCategorias[a].total);

  if (categoriasOrdenadas.length === 0) {
      mensagemResumo += "Nenhuma despesa detalhada registrada para este mes.\n";
  } else {
      categoriasOrdenadas.forEach(categoria => {
        const totalCategoria = resumoCategorias[categoria].total;
        const metaInfo = metasPorCategoria[normalizarTexto(categoria)] || { totalMeta: 0, totalGasto: 0 };
        
        mensagemResumo += `\n*${escapeMarkdown(capitalize(categoria))}:* R$ ${totalCategoria.toFixed(2).replace('.', ',')}`;
        
        if (metaInfo.totalMeta > 0) {
          const percMeta = (metaInfo.totalGasto / metaInfo.totalMeta) * 100;
          let emojiMeta = percMeta >= 100 ? "⛔" : (percMeta >= 80 ? "⚠️" : "✅");
          mensagemResumo += ` ${emojiMeta} (${percMeta.toFixed(0)}% da meta de R$ ${metaInfo.totalMeta.toFixed(2).replace('.', ',')})`;
        }
        mensagemResumo += `\n`;

        const subcategoriasOrdenadas = Object.keys(resumoCategorias[categoria].subcategories).sort((a, b) => resumoCategorias[categoria].subcategories[b] - resumoCategorias[categoria].subcategories[a]);
        subcategoriasOrdenadas.forEach(sub => {
          const gastoSub = resumoCategorias[categoria].subcategories[sub];
          const chaveSubcategoriaMeta = normalizarTexto(`${categoria} ${sub}`);
          const subMetaInfo = metasPorCategoria[normalizarTexto(categoria)]?.subcategories[chaveSubcategoriaMeta];

          let subLine = `  • ${escapeMarkdown(capitalize(sub))}: R$ ${gastoSub.toFixed(2).replace('.', ',')}`;
          if (subMetaInfo && subMetaInfo.meta > 0) {
            let subPerc = (subMetaInfo.gasto / subMetaInfo.meta) * 100;
            let subEmoji = subPerc >= 100 ? "⛔" : (subPerc >= 80 ? "⚠️" : "✅");
            subLine += ` / R$ ${subMetaInfo.meta.toFixed(2).replace('.', ',')} ${subEmoji} ${subPerc.toFixed(0)}%`;
          }
          mensagemResumo += `${subLine}\n`;
        });
      });
  }

  mensagemResumo += `\n*💳 Gastos de Cartão de Crédito no Mês*\n`;
  const cartoesOrdenados = Object.keys(resumoCartoes).sort((a, b) => resumoCartoes[b].faturaBrutaMes - resumoCartoes[a].faturaBrutaMes);
  if (cartoesOrdenados.length === 0) {
    mensagemResumo += "Nenhum gasto em cartão de crédito registrado neste mês.\n";
  } else {
    cartoesOrdenados.forEach(cartaoNormalizadoKey => { 
      const infoCartao = resumoCartoes[cartaoNormalizadoKey];
      if (infoCartao.faturaBrutaMes > 0) { 
          const vencimentoTexto = infoCartao.vencimento ? ` (Venc: Dia ${infoCartao.vencimento})` : "";
          const limiteTexto = infoCartao.limite > 0 ? ` / Limite: R$ ${infoCartao.limite.toFixed(2).replace('.', ',')}` : "";
          const displayName = escapeMarkdown(infoCartao.nomeOriginalParaExibicao || capitalize(cartaoNormalizadoKey)); 
          mensagemResumo += `• *${displayName}*: R$ ${infoCartao.faturaBrutaMes.toFixed(2).replace('.', ',')}${vencimentoTexto}${limiteTexto}\n`;
      }
    });
  }
  
  return mensagemResumo;
}


/**
 * Envia o resumo financeiro do mês atual para o chat do Telegram.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário que solicitou o resumo.
 * @param {number} mes O mês para o resumo (1-12).
 * @param {number} ano O ano para o resumo.
 */
function enviarResumo(chatId, usuario, mes, ano) {
  const targetMes = mes;
  const targetAno = ano;

  const mensagemResumo = gerarResumoMensal(targetMes, targetAno);
  enviarMensagemTelegram(chatId, mensagemResumo);
  logToSheet(`Resumo mensal enviado para ${chatId}.`, "INFO");
}


// ===================================================================================
// SEÇÃO DE SALDO - REFATORADA
// ===================================================================================

/**
 * @private
 * Formata uma seção da mensagem de saldo, lidando com grupos e filhos.
 * @param {string} titulo O título da seção.
 * @param {Object} dadosAgrupados Os dados calculados e agrupados.
 * @param {string} mensagemSemDados A mensagem a ser exibida se não houver dados.
 * @returns {string} A string formatada para a seção.
 */
function _formatarSecaoMensagem(titulo, dadosAgrupados, mensagemSemDados) {
  let secaoTexto = `\n*${titulo}*\n`;
  if (Object.keys(dadosAgrupados).length === 0) {
    secaoTexto += `_${mensagemSemDados}_\n`;
    return secaoTexto;
  }

  Object.keys(dadosAgrupados).forEach(pai => {
    const grupo = dadosAgrupados[pai];
    secaoTexto += `• *${escapeMarkdown(capitalize(pai))}: ${formatCurrency(grupo.total)}*\n`;
    if (grupo.filhos.length > 1) {
      grupo.filhos.forEach(filho => {
        secaoTexto += `  - _${escapeMarkdown(filho.nome.split(' ').pop())}: ${formatCurrency(filho.valor)}_\n`;
      });
    }
  });
  return secaoTexto;
}

/**
 * @private
 * Calcula e agrupa os saldos das contas correntes e dinheiro.
 * @param {Object} saldosCalculados O objeto global com os saldos pré-calculados.
 * @returns {{total: number, contas: Array<Object>}} Objeto com o total e a lista de contas.
 */
function _calcularSaldosContasCorrentes(saldosCalculados) {
  let total = 0;
  const contas = [];
  for (const nomeNormalizado in saldosCalculados) {
    const infoConta = saldosCalculados[nomeNormalizado];
    if (infoConta.tipo === "conta corrente" || infoConta.tipo === "dinheiro físico") {
      total += infoConta.saldo;
      contas.push(infoConta);
    }
  }
  return { total, contas };
}

/**
 * @private
 * Calcula e agrupa as faturas com vencimento no próximo ciclo.
 * @param {Object} saldosCalculados O objeto global com os saldos pré-calculados.
 * @returns {Object} Um objeto com os dados das faturas agrupados por conta pai.
 */
function _calcularFaturasProximoMes(saldosCalculados) {
  const agrupadores = {};
  for (const nomeNormalizado in saldosCalculados) {
    const infoConta = saldosCalculados[nomeNormalizado];
    if (infoConta.tipo === "cartão de crédito" && infoConta.faturaAtual > 0) {
      const pai = infoConta.contaPaiAgrupador || infoConta.nomeNormalizado;
      if (!agrupadores[pai]) {
        agrupadores[pai] = { total: 0, filhos: [] };
      }
      agrupadores[pai].total += infoConta.faturaAtual;
      agrupadores[pai].filhos.push({ nome: infoConta.nomeOriginal, valor: infoConta.faturaAtual });
    }
  }
  return agrupadores;
}

/**
 * @private
 * Calcula e agrupa as faturas com vencimento no mês atual.
 * @param {Array<Array<any>>} dadosTransacoes Dados da aba 'Transacoes'.
 * @param {Array<Array<any>>} dadosContas Dados da aba 'Contas'.
 * @param {number} mes O mês atual (0-11).
 * @param {number} ano O ano atual.
 * @returns {Object} Um objeto com os dados das faturas do mês atual agrupados.
 */
function _calcularFaturasMesAtual(dadosTransacoes, dadosContas, mes, ano) {
  const faturasDoMes = {};
  for (let i = 1; i < dadosTransacoes.length; i++) {
    const linha = dadosTransacoes[i];
    const dataVencimento = parseData(linha[10]);
    if (dataVencimento && dataVencimento.getMonth() === mes && dataVencimento.getFullYear() === ano) {
      const conta = linha[7];
      const infoConta = obterInformacoesDaConta(conta, dadosContas);
      if (infoConta && normalizarTexto(infoConta.tipo) === "cartao de credito") {
        const pai = infoConta.contaPaiAgrupador || infoConta.nomeNormalizado;
        const valor = parseBrazilianFloat(String(linha[5]));
        if (!faturasDoMes[pai]) {
          faturasDoMes[pai] = { total: 0, filhos: {} };
        }
        faturasDoMes[pai].total += valor;
        faturasDoMes[pai].filhos[infoConta.nomeOriginal] = (faturasDoMes[pai].filhos[infoConta.nomeOriginal] || 0) + valor;
      }
    }
  }
  // Formata a saída para ser compatível com _formatarSecaoMensagem
  const resultadoFormatado = {};
  for(const pai in faturasDoMes) {
    resultadoFormatado[pai] = {
      total: faturasDoMes[pai].total,
      filhos: Object.entries(faturasDoMes[pai].filhos).map(([nome, valor]) => ({nome, valor}))
    };
  }
  return resultadoFormatado;
}

/**
 * @private
 * Calcula e agrupa a dívida total pendente em todos os cartões.
 * @param {Object} saldosCalculados O objeto global com os saldos pré-calculados.
 * @returns {{total: number, devedores: Object}} Objeto com o total e os dados agrupados.
 */
function _calcularDividaTotalCartoes(saldosCalculados) {
  const devedores = {};
  let total = 0;
  for (const nomeNormalizado in saldosCalculados) {
    const infoConta = saldosCalculados[nomeNormalizado];
    if (infoConta.tipo === "cartão de crédito" && infoConta.saldoTotalPendente > 0.01) {
      const pai = infoConta.contaPaiAgrupador || infoConta.nomeNormalizado;
      if (!devedores[pai]) {
        devedores[pai] = { total: 0, filhos: [] };
      }
      devedores[pai].total += infoConta.saldoTotalPendente;
      devedores[pai].filhos.push({ nome: infoConta.nomeOriginal, valor: infoConta.saldoTotalPendente });
      total += infoConta.saldoTotalPendente;
    }
  }
  return { total, devedores };
}

/**
 * **REFATORADO:** Envia o saldo atual das contas e faturas de cartão de crédito.
 * A lógica de cálculo foi movida para funções auxiliares para maior clareza.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário solicitante.
 */
function enviarSaldo(chatId, usuario) {
  logToSheet(`Iniciando enviarSaldo para chatId: ${chatId}, usuario: ${usuario}`, "INFO");

  try {
    // 1. Carregar Dados
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configData = getSheetDataWithCache(SHEET_CONFIGURACOES, CACHE_KEY_CONFIG);
    const transacoesData = getSheetDataWithCache(SHEET_TRANSACOES, 'transacoes_cache');
    const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);
    const grupoUsuarioChat = getGrupoPorChatId(chatId, configData);
    const today = new Date();

    // 2. Garantir que os saldos globais estão atualizados
    atualizarSaldosDasContas(); 
    logToSheet(`[enviarSaldo] Saldos globais atualizados.`, "DEBUG");

    // 3. Chamar funções de cálculo especializadas
    const { total: totalContasCorrentes, contas: contasCorrentes } = _calcularSaldosContasCorrentes(globalThis.saldosCalculados);
    const faturasProximoMes = _calcularFaturasProximoMes(globalThis.saldosCalculados);
    const faturasMesAtual = _calcularFaturasMesAtual(transacoesData, dadosContas, today.getMonth(), today.getFullYear());
    const { total: totalDividaCartoes, devedores: devedoresCartoes } = _calcularDividaTotalCartoes(globalThis.saldosCalculados);

    // 4. Montar a Mensagem
    let mensagemSaldo = `💰 *Saldos Atuais - ${escapeMarkdown(grupoUsuarioChat || 'Pessoal')}*\n`;
    
    // Seção Contas Correntes
    mensagemSaldo += `──────────────\n*💵 Contas e Dinheiro*\n`;
    contasCorrentes.sort((a, b) => a.nomeOriginal.localeCompare(b.nomeOriginal)).forEach(conta => {
        mensagemSaldo += `• ${escapeMarkdown(capitalize(conta.nomeOriginal))}: *${formatCurrency(conta.saldo)}*\n`;
    });

    // Seções de Faturas e Dívidas
    mensagemSaldo += _formatarSecaoMensagem("🗓️ Faturas (Próximo Vencimento)", faturasProximoMes, "Nenhuma fatura para o próximo ciclo.");
    mensagemSaldo += _formatarSecaoMensagem(`🧾 Faturas a Vencer em ${getNomeMes(today.getMonth())}`, faturasMesAtual, "Nenhuma fatura de cartão a vencer neste mês.");
    mensagemSaldo += _formatarSecaoMensagem("💳 Dívida Total dos Cartões", devedoresCartoes, "Nenhum saldo devedor encontrado.");

    // Seção de Resumo Geral
    mensagemSaldo += `──────────────\n*🏦 Resumo Geral*\n`;
    mensagemSaldo += `*Total Disponível:* ${formatCurrency(totalContasCorrentes)}\n`;
    mensagemSaldo += `*Saldo Líquido (Disponível - Dívida):* ${formatCurrency(totalContasCorrentes - totalDividaCartoes)}\n`;

    // 5. Enviar Mensagem
    enviarMensagemTelegram(chatId, mensagemSaldo);
    logToSheet(`Saldo enviado para chatId: ${chatId}.`, "INFO");

  } catch (e) {
    // Chamamos a função centralizada para registar o erro e notificar o utilizador.
    handleError(e, `enviarSaldo para ${usuario}`, chatId);
    // **CORREÇÃO CRÍTICA ADICIONADA AQUI**
    // Adicionamos 'return' para parar a execução da função imediatamente.
    return;
  }
}

// ===================================================================================
// FIM DA SEÇÃO DE SALDO
// ===================================================================================


/**
 * Envia uma mensagem de ajuda com exemplos de comandos para o chat do Telegram.
 * @param {string} chatId O ID do chat do Telegram.
 */
function enviarAjuda(chatId) {
  const mensagem = `
👋 *Bem-vindo ao Boas Contas!*

Aqui está um guia completo das minhas funcionalidades. Use os botões para acesso rápido aos comandos mais comuns.

---
*💸 LANÇAMENTOS (LINGUAGEM NATURAL)*
---
Para registar, basta enviar uma mensagem como se estivesse a conversar.

*Gastos:* \`gastei 50 no mercado com Nubank\`
*Receitas:* \`recebi 3000 de salario no Itau\`
*Transferências:* \`transferi 200 do Itau para o PicPay\`
*Parcelamentos:* \`comprei um telemóvel de 1200 em 10x\`
*Dividir Despesas:* \`dividi 100 do jantar com a Ana pelo Itau\`
*Empréstimos:* \`emprestei 50 para o João com o Nubank\`

---
*📊 ANÁLISES E RELATÓRIOS*
---
• \`/resumo\` – Visão geral financeira do mês.
• \`/saldo\` – Saldos de todas as contas e faturas.
• \`/extrato\` – Suas últimas transações.
• \`/saude\` – Um check-up da sua saúde financeira.
• \`/dashboard\` – Aceder ao dashboard web completo.

---
*📈 INVESTIMENTOS*
---
• \`comprei 10 ITSA4 a 8.50 pela NuInvest\`
• \`vendi 5 PETR4 a 32.00 para a NuInvest\`
• \`/provento ITSA4 50 para NuInvest\` – Regista dividendos.
• \`/patrimonio\` – Mostra o seu património líquido total.

---
*🎯 PLANEAMENTO*
---
• \`/orcamento\` – Acompanhe seu orçamento de gastos.
• \`/metas\` – Veja o progresso das suas metas de poupança.
• \`/novameta NOME VALOR\` – Cria uma nova meta.
• \`/aportarmeta META VALOR de CONTA\` – Adiciona dinheiro a uma meta.
• \`/contasapagar\` – Status das suas contas fixas.

---
*🗓️ PRODUTIVIDADE*
---
• \`/tarefa DESCRIÇÃO\` – Cria uma nova tarefa.
• \`/tarefas\` – Lista suas tarefas pendentes.

---
*⚙️ OUTROS COMANDOS*
---
• \`/meuperfil\` – Descubra o seu perfil financeiro.
• \`/editar ultimo\` – Corrigir o último lançamento.
• \`/ajuda\` – Ver esta mensagem novamente.
  `;

  // --- INÍCIO DA CORREÇÃO ---
  // Reorganiza o teclado para incluir o botão /extrato e manter a estrutura de 3x3
  const teclado = {
    inline_keyboard: [
      [
        { text: "📊 Resumo", callback_data: "/resumo" },
        { text: "💰 Saldo", callback_data: "/saldo" },
        { text: "📄 Extrato", callback_data: "/extrato" }
      ],
      [
        { text: "🩺 Saúde", callback_data: "/saude" },
        { text: "📈 Património", callback_data: "/patrimonio" },
        { text: "🧾 Orçamento", callback_data: "/orcamento" }
      ],
      [
        { text: "🎯 Metas", callback_data: "/metas" },
        { text: "📝 Tarefas", callback_data: "/tarefas" },
        { text: "🌐 Dashboard", callback_data: "/dashboard" }
      ]
    ]
  };
  // --- FIM DA CORREÇÃO ---

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
}





/**
 * Verifica as metas financeiras e envia alertas para o Telegram se os limites forem atingidos.
 * Esta função é geralmente executada por um gatilho de tempo (trigger).
 */
function verificarAlertas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoes = ss.getSheetByName(SHEET_TRANSACOES).getDataRange().getValues();
  const metasSheet = ss.getSheetByName(SHEET_METAS).getDataRange().getValues();
  const alertasSheet = ss.getSheetByName(SHEET_ALERTAS_ENVIADOS);
  const alertas = alertasSheet.getDataRange().getValues();
  const config = ss.getSheetByName(SHEET_CONFIGURACOES).getDataRange().getValues();

  const hoje = new Date();
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();
  const nomeMes = getNomeMes(mesAtual);

  const cabecalho = metasSheet[2];
  let colMetaMes = -1;

  for (let i = 2; i < cabecalho.length; i++) {
    if (String(cabecalho[i]).toLowerCase().includes(nomeMes.toLowerCase())) {
      colMetaMes = i;
      break;
    }
  }
  if (colMetaMes === -1) {
    logToSheet(`[VerificarAlertas] Coluna do mes para ${nomeMes}/${anoAtual} não encontrada na aba 'Metas'.`, "INFO");
    return;
  }

  let metas = {};

  logToSheet("[VerificarAlertas] --- Inicio do Processamento de Metas (verificarAlertas) ---", "DEBUG");
  for (let i = 3; i < metasSheet.length; i++) {
    const categoriaOriginal = (metasSheet[i][0] || "").toString().trim();
    const subcategoriaOriginal = (metasSheet[i][1] || "").toString().trim();
    const valorMetaTexto = metasSheet[i][colMetaMes];

    if (!categoriaOriginal || !subcategoriaOriginal || !valorMetaTexto) continue;

    const chave = normalizarTexto(`${categoriaOriginal} ${subcategoriaOriginal}`);

    let meta = parseBrazilianFloat(String(valorMetaTexto));

    if (isNaN(meta) || meta <= 0) continue;

    metas[chave] = {
      categoria: categoriaOriginal,
      subcategoria: subcategoriaOriginal,
      meta: meta,
      gastoPorUsuario: {}
    };
  }
  logToSheet("[VerificarAlertes] --- Fim do Processamento de Metas (verificarAlertas) ---", "DEBUG");


  logToSheet("[VerificarAlertas] --- Inicio do Processamento de Transacoes (verificarAlertas) ---", "DEBUG");
  for (let i = 1; i < transacoes.length; i++) {
    const dataVencimento = parseData(transacoes[i][10]); // Use Data de Vencimento
    const tipo = transacoes[i][4];
    const categoria = transacoes[i][2];
    const subcategoria = transacoes[i][3];
    const rawValor = transacoes[i][5];
    const usuario = transacoes[i][11];

    if (
      !dataVencimento || dataVencimento.getMonth() !== mesAtual || dataVencimento.getFullYear() !== anoAtual || // Filter by DUE DATE
      tipo !== "Despesa"
    ) continue;

    const chave = normalizarTexto(`${categoria} ${subcategoria}`);
    if (!metas[chave]) continue;

    let valor = parseBrazilianFloat(String(rawValor));

    if (!isNaN(valor)) {
      if (!metas[chave].gastoPorUsuario[usuario]) {
        metas[chave].gastoPorUsuario[usuario] = 0;
      }
      metas[chave].gastoPorUsuario[usuario] += valor;
    }
  }
  logToSheet("[VerificarAlertas] --- Fim do Processamento de Transacoes (verificarAlertas) ---", "DEBUG");


  const jaEnviados = alertas.map(row => `${row[1]}|${row[2]}|${row[3]}|${row[4]}`);

  for (const chave in metas) {
    const metaObj = metas[chave];
    for (const usuario in metaObj.gastoPorUsuario) {
      const gasto = metaObj.gastoPorUsuario[usuario];
      const perc = (gasto / metaObj.meta) * 100;

      for (const tipoAlerta of [80, 100]) {
        if (perc >= tipoAlerta) {
          const codigo = `${usuario}|${metaObj.categoria}|${metaObj.subcategoria}|${tipoAlerta}%`;
          if (!jaEnviados.includes(codigo)) {
            const mensagem = tipoAlerta === 80
              ? `⚠️ *Atencao!* "${escapeMarkdown(metaObj.subcategoria)}" ja atingiu *${Math.round(perc)}%* da meta de ${nomeMes}.\nMeta: R$ ${metaObj.meta.toFixed(2).replace('.', ',')} • Gasto: R$ ${gasto.toFixed(2).replace('.', ',')}`
              : `⛔ *Meta ultrapassada!* "${escapeMarkdown(metaObj.subcategoria)}" ja passou *100%* da meta de ${nomeMes}.\nMeta: R$ ${metaObj.meta.toFixed(2).replace('.', ',')} • Gasto: R$ ${gasto.toFixed(2).replace('.', ',')}`;

            const chatId = getChatId(config, usuario);
            if (chatId) {
              enviarMensagemTelegram(chatId, mensagem);

              alertasSheet.appendRow([
                Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
                usuario,
                metaObj.categoria,
                metaObj.subcategoria,
                `${tipoAlerta}%`
              ]);
              logToSheet(`Alerta de meta enviado para ${usuario} para ${metaObj.subcategoria} (${tipoAlerta}%).`, "INFO");
            } else {
              logToSheet(`[VerificarAlertas] Não foi possivel encontrar chatId para o usuario ${usuario} para enviar alerta de meta para ${metaObj.subcategoria}.`, "WARN");
            }
          } else {
            logToSheet(`[VerificarAlertas] Alerta para ${usuario} para ${metaObj.subcategoria} (${tipoAlerta}%) ja foi enviado. Pulando.`, "DEBUG");
          }
        }
      }
    }
  }
  logToSheet("[VerificarAlertas] Verificacao de alertas concluida.", "INFO");
}

/**
 * Envia o extrato das últimas transações para o chat do Telegram.
 * Permite filtrar por tipo (receita/despesa) ou por usuário.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário que solicitou o extrato.
 * @param {string} [complemento=""] Um complemento de filtro (ex: "receitas", "despesas", nome de usuário, "tudo").
 */
function enviarExtrato(chatId, usuario, complemento = "") {
  logToSheet(`[Extrato] Iniciando enviarExtrato para chatId: ${chatId}, usuario: ${usuario}, complemento: "${complemento}"`, "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoes = ss.getSheetByName(SHEET_TRANSACOES).getDataRange().getValues();
  const config = ss.getSheetByName(SHEET_CONFIGURACOES).getDataRange().getValues();
  const grupoLinha = getGrupoPorChatId(chatId, config);

  const complementoNormalizado = normalizarTexto(complemento);
  logToSheet(`[Extrato] Complemento normalizado: "${complementoNormalizado}"`, "DEBUG");

  const { month: targetMonth, year: targetYear } = parseMonthAndYear(complemento);
  const targetMesIndex = targetMonth - 1;
  const nomeMes = getNomeMes(targetMesIndex);
  logToSheet(`[Extrato] Mes Alvo: ${nomeMes}/${targetYear}`, "DEBUG");


  let tipoFiltro = null;
  let usuarioAlvo = null;

  if (complementoNormalizado.includes("receitas")) {
    tipoFiltro = "Receita";
    logToSheet(`[Extrato] Filtro de tipo: Receita`, "DEBUG");
  }
  else if (complementoNormalizado.includes("despesas")) {
    tipoFiltro = "Despesa";
    logToSheet(`[Extrato] Filtro de tipo: Despesa`, "DEBUG");
  }

  for (let i = 1; i < config.length; i++) {
    const nomeConfig = config[i][2];
    if (!nomeConfig || normalizarTexto(nomeConfig) === "nomeusuario") continue;
    const nomeNormalizadoConfig = normalizarTexto(nomeConfig);
    if (complementoNormalizado.includes(nomeNormalizadoConfig)) {
      usuarioAlvo = nomeConfig;
      logToSheet(`[Extrato] Usuario Alvo detectado no complemento: ${usuarioAlvo}`, "DEBUG");
      break;
    }
  }

  let ultimas = [];

  for (let i = transacoes.length - 1; i > 0; i--) {
    const linha = transacoes[i];
    const data = parseData(linha[0]);
    const desc = linha[1];
    const categoria = linha[2];
    const subcategoria = linha[3];
    const tipo = linha[4];
    const valor = linha[5];
    const metodo = linha[6];
    const conta = linha[7];
    const usuarioLinha = linha[11];
    const id = linha[13];
    const grupoTransacao = getGrupoPorChatIdByUsuario(usuarioLinha, config);

    logToSheet(`[Extrato] Processando transacao ID: ${id || 'N/A'}, Data: ${data ? data.toLocaleDateString() : 'N/A'}, Usuario Linha: "${usuarioLinha}", Tipo: ${tipo}`, "DEBUG");

    let isIncluded = false;
    if (!data || data.getMonth() !== targetMesIndex || data.getFullYear() !== targetYear) {
      logToSheet(`[Extrato] Transacao ID: ${id} ignorada: Data (${data ? data.toLocaleDateString() : 'N/A'}) fora do mes/ano alvo.`, "DEBUG");
      continue;
    }

    if (complementoNormalizado.includes("tudo")) {
      const isOwnerOrAdmin = (normalizarTexto(usuario) === normalizarTexto(getUsuarioPorChatId(chatId, config)));
      logToSheet(`[Extrato] Modo 'tudo'. Usuario solicitante: "${usuario}", isOwnerOrAdmin: ${isOwnerOrAdmin}`, "DEBUG");

      if (isOwnerOrAdmin) {
          isIncluded = (grupoTransacao === grupoLinha);
          logToSheet(`[Extrato] Admin/Owner. Grupo Transacao: ${grupoTransacao}, Grupo Chat: ${grupoLinha}. Includo: ${isIncluded}`, "DEBUG");
      } else {
          isIncluded = (normalizarTexto(usuarioLinha) === normalizarTexto(usuario));
          logToSheet(`[Extrato] Nao Admin/Owner. Usuario Linha: "${usuarioLinha}", Usuario Solicitante: "${usuario}". Includo: ${isIncluded}`, "DEBUG");
      }
    } else if (usuarioAlvo) {
      isIncluded = (normalizarTexto(usuarioLinha) === normalizarTexto(usuarioAlvo));
      logToSheet(`[Extrato] Filtro por usuario alvo: "${usuarioAlvo}". Usuario Linha: "${usuarioLinha}". Includo: ${isIncluded}`, "DEBUG");
    } else {
      isIncluded = (normalizarTexto(usuarioLinha) === normalizarTexto(usuario));
      logToSheet(`[Extrato] Filtro padrao (proprio usuario). Usuario Linha: "${usuarioLinha}", Usuario Solicitante: "${usuario}". Includo: ${isIncluded}`, "DEBUG");
    }

    if (isIncluded && (!tipoFiltro || normalizarTexto(tipo) === normalizarTexto(tipoFiltro))) {
      ultimas.push({
        data: linha[0],
        descricao: desc,
        categoria,
        subcategoria,
        tipo,
        valor,
        metodo,
        conta,
        usuario: usuarioLinha,
        id: linha[13]
      });
      logToSheet(`[Extrato] Transacao ID: ${id} adicionada ao extrato.`, "DEBUG");
    } else {
      logToSheet(`[Extrato] Transacao ID: ${id} ignorada por filtros (isIncluded: ${isIncluded}, tipoFiltro: ${tipoFiltro}, tipoTransacao: ${tipo}).`, "DEBUG");
    }

    if (ultimas.length >= 5 && !complementoNormalizado.includes("tudo")) {
      logToSheet(`[Extrato] Limite de 5 transacoes atingido (nao 'tudo').`, "DEBUG");
      break;
    }
    if (ultimas.length >= 10 && complementoNormalizado.includes("tudo")) {
      logToSheet(`[Extrato] Limite de 10 transacoes atingido ('tudo').`, "DEBUG");
      break;
    }
  }

  ultimas.reverse();
  logToSheet(`[Extrato] Total de transacoes apos filtros e ordenacao: ${ultimas.length}`, "INFO");


  if (ultimas.length === 0) {
    enviarMensagemTelegram(chatId, `📄 Nenhum lancamento ${tipoFiltro || ""} encontrado em ${nomeMes}/${targetYear}${usuarioAlvo ? ' para ' + escapeMarkdown(usuarioAlvo) : ''}.`);
    logToSheet(`[Extrato] Nenhuma transacao encontrada para os filtros.`, "INFO");
    return;
  }

  let mensagemInicial = `? *Ultimos lancamentos ${tipoFiltro ? "(" + tipoFiltro + ")" : ""} – ${nomeMes}/${targetYear}*`;

  if (usuarioAlvo) mensagemInicial += `\n👤 Usuario: ${escapeMarkdown(capitalize(usuarioAlvo))}`;
  else mensagemInicial += `\n👥 Grupo: ${escapeMarkdown(grupoLinha)}`;

  mensagemInicial += "\n\n";

  enviarMensagemTelegram(chatId, mensagemInicial);
  logToSheet(`[Extrato] Mensagem inicial enviada.`, "DEBUG");

  ultimas.forEach((t) => {
    const dataObj = parseData(t.data);
    const dataFormatada = dataObj
      ? Utilities.formatDate(dataObj, Session.getScriptTimeZone(), "dd/MM/yyyy")
      : "Data invalida";

    const meio = t.metodo ? `💳 ${escapeMarkdown(t.metodo)} | ` : "";

    let valorNumerico = parseBrazilianFloat(String(t.valor));

    const textoTransacao = `📌 *${escapeMarkdown(t.descricao)}*\n🗓 ${dataFormatada} – ${escapeMarkdown(t.categoria)} > ${escapeMarkdown(t.subcategoria)}\n${meio}R$ ${valorNumerico.toFixed(2).replace('.', ',')} – ${escapeMarkdown(t.tipo)}`;

    const tecladoTransacao = {
      inline_keyboard: [[{
        text: "🗑 Excluir lancamento",
        callback_data: `/excluir_${t.id}`
      }]]
    };

    enviarMensagemTelegram(chatId, textoTransacao, { reply_markup: tecladoTransacao });
    logToSheet(`[Extrato] Transacao ID: ${t.id} enviada com botao de exclusao.`, "DEBUG");
  });
  logToSheet(`[Extrato] Envio de extrato concluido.`, "INFO");
}

/**
 * Mostra um menu inline no Telegram para opções de extrato.
 * @param {string} chatId O ID do chat do Telegram.
 */
function mostrarMenuExtrato(chatId) {
  const mensagem = "? O que voce deseja ver?";

  const teclado = {
    inline_keyboard: [
      [
        { text: "🔍 Tudo", callback_data: "/extrato_tudo" },
        { text: "💰 Receitas", callback_data: "/extrato_receitas" },
        { text: "💸 Despesas", callback_data: "/extrato_despesas" }
      ],
      [
        { text: "👤 Por Pessoa", callback_data: "/extrato_pessoa" }
      ]
    ]
  };

  const config = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(SHEET_CONFIGURACOES)
    .getDataRange()
    .getValues();

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
}

/**
 * Mostra um menu inline no Telegram para selecionar um usuário para visualizar o extrato.
 * @param {Array<Array<any>>} config Os dados da aba "Configuracoes".
 * @param {string} chatId O ID do chat do Telegram.
 */
function mostrarMenuPorPessoa(chatId, config) {
  const nomes = [];
  for (let i = 1; i < config.length; i++) {
    const chave = config[i][0];
    const nome = config[i][2];
    if (chave === "chatId" && nome && !nomes.includes(nome)) {
      nomes.push(nome);
    }
  }

  const linhas = nomes.map((nome) => {
    return [{ text: nome, callback_data: `/extrato_usuario_${normalizarTexto(nome)}` }];
  });

  linhas.push([{ text: "↩️ Voltar", callback_data: "/extrato" }]);

  const teclado = { inline_keyboard: linhas };

  const mensagem = "👤 Escolha uma pessoa:";

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
}

/**
 * ATUALIZADA: Exclui um lançamento da aba "Transacoes" pelo seu ID único.
 * Se o lançamento for um "Aporte Meta", o valor é revertido na aba "Metas".
 * Se estiver vinculado a uma conta a pagar, o status é revertido.
 * @param {string} idLancamento O ID único do lançamento a ser excluído.
 * @param {string} chatId O ID do chat do Telegram para enviar feedback.
 */
function excluirLancamentoPorId(idLancamento, chatId) {
  logToSheet(`Iniciando exclusao de lancamento para ID: ${idLancamento}`, "INFO");
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
    if (!transacoesSheet) throw new Error("Aba 'Transacoes' não encontrada.");

    const dadosTransacoes = transacoesSheet.getDataRange().getValues();
    const headersTransacoes = dadosTransacoes[0];
    const colMap = getColumnMap(headersTransacoes);
    const colIdTransacao = colMap['ID Transacao'];
    const colDescricao = colMap['Descricao'];
    const colValor = colMap['Valor'];
    const colTipo = colMap['Tipo'];
    const colConta = colMap['Conta / Cartão'];


    if (colIdTransacao === -1 || colDescricao === -1 || colValor === -1 || colTipo === -1 || colConta === -1) {
      throw new Error("Colunas essenciais não encontradas na aba 'Transacoes'.");
    }

    let linhaParaExcluir = -1;
    let lancamentoParaExcluir = null;

    for (let i = 1; i < dadosTransacoes.length; i++) {
      if (dadosTransacoes[i][colIdTransacao] === idLancamento) {
        linhaParaExcluir = i + 1;
        lancamentoParaExcluir = dadosTransacoes[i];
        break;
      }
    }

    if (linhaParaExcluir !== -1) {
      const descricaoLancamento = lancamentoParaExcluir[colDescricao];
      const valorLancamento = parseBrazilianFloat(String(lancamentoParaExcluir[colValor]));

      // ### INÍCIO DA NOVA LÓGICA ###
      // Verifica se é um aporte de meta e reverte o valor
      if (descricaoLancamento.startsWith("Aporte Meta:")) {
        const nomeMetaExtraido = descricaoLancamento.substring("Aporte Meta:".length).trim();
        reverterAporteMeta(nomeMetaExtraido, valorLancamento);
      }
      // ### FIM DA NOVA LÓGICA ###

      // Lógica existente para reverter contas a pagar
      reverterStatusContaAPagarSeVinculado(idLancamento);

      // --- INÍCIO DA MELHORIA: Ajuste Incremental Reverso ---
      const tipo = lancamentoParaExcluir[colTipo];
      const conta = lancamentoParaExcluir[colConta];
      const contasSheet = ss.getSheetByName(SHEET_CONTAS);
      const dadosContas = contasSheet.getDataRange().getValues();
      const infoConta = obterInformacoesDaConta(conta, dadosContas);

      let valorReversao;
      if (infoConta && normalizarTexto(infoConta.tipo) === "cartao de credito") {
        // Reverter um gasto em cartão de crédito significa DIMINUIR a dívida.
        // Reverter uma receita (estorno) significa AUMENTAR a dívida.
        valorReversao = tipo === 'Receita' ? valorLancamento : -valorLancamento;
      } else {
        // Reverter um gasto em conta normal significa AUMENTAR o saldo.
        // Reverter uma receita significa DIMINUIR o saldo.
        valorReversao = tipo === 'Receita' ? -valorLancamento : valorLancamento;
      }

      // Exclui a linha e DEPOIS aplica o ajuste
      transacoesSheet.deleteRow(linhaParaExcluir);
      logToSheet(`Lancamento '${descricaoLancamento}' (ID: ${idLancamento}) excluído da aba 'Transacoes'.`, "INFO");
      
      ajustarSaldoIncrementalmente(contasSheet, conta, valorReversao);
      logToSheet(`Ajuste incremental de REVERSÃO aplicado para '${conta}'. Valor: ${valorReversao}`, "INFO");
      // --- FIM DA MELHORIA ---
      
      enviarMensagemTelegram(chatId, `✅ Lançamento '${escapeMarkdown(descricaoLancamento)}' excluído com sucesso! Saldo e metas atualizados.`);
    } else {
      enviarMensagemTelegram(chatId, `❌ Lançamento com ID *${escapeMarkdown(idLancamento)}* não encontrado.`);
      logToSheet(`Erro: Lancamento ID ${idLancamento} nao encontrado para exclusao.`, "WARN");
    }
  } catch (e) {
    handleError(e, `excluirLancamentoPorId para ${idLancamento}`, chatId);
  } finally {
    lock.releaseLock();
  }
}

/**
 * NOVO: Envia um resumo das faturas futuras de cartões de crédito.
 * Calcula o total de despesas por cartão e por mês de vencimento futuro.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário que solicitou.
 */
function enviarFaturasFuturas(chatId, usuario) {
  logToSheet(`Iniciando enviarFaturasFuturas para chatId: ${chatId}, usuario: ${usuario}`, "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaTransacoes = ss.getSheetByName(SHEET_TRANSACOES);
  // Carrega a aba Contas usando o cache
  const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS); 

  if (!abaTransacoes || !dadosContas) { // Verifica dadosContas
    enviarMensagemTelegram(chatId, "❌ Erro: As abas 'Transacoes' ou 'Contas' não foram encontradas. Verifique os nomes das abas.");
    logToSheet("Erro: Abas Transacoes ou Contas não encontradas.", "ERROR");
    return;
  }

  const dadosTransacoes = abaTransacoes.getDataRange().getValues();

  const hoje = new Date();
  const currentMonth = hoje.getMonth();
  const currentYear = hoje.getFullYear();

  let faturasFuturas = {};

  const infoCartoesMap = {};
  for (let i = 1; i < dadosContas.length; i++) {
    const nomeConta = (dadosContas[i][0] || "").toString().trim();
    const nomeContaNormalizada = normalizarTexto(nomeConta);
    const tipoConta = (dadosContas[i][1] || "").toString().toLowerCase().trim();
    if (normalizarTexto(tipoConta) === "cartao de credito") {
      infoCartoesMap[nomeContaNormalizada] = obterInformacoesDaConta(nomeConta, dadosContas); // Passa dadosContas
    }
  }

  for (let i = 1; i < dadosTransacoes.length; i++) {
    const linhaTransacao = dadosTransacoes[i];
    const tipoTransacao = (linhaTransacao[4] || "").toString().toLowerCase().trim();
    const contaAssociada = (linhaTransacao[7] || "").toString().trim();
    const contaAssociadaNormalizada = normalizarTexto(contaAssociada);
    const categoria = (linhaTransacao[2] || "").toString().trim();
    const subcategoria = (linhaTransacao[3] || "").toString().trim();
    
    let valor = parseBrazilianFloat(String(linhaTransacao[5]));

    if (tipoTransacao === "despesa" && 
        infoCartoesMap[contaAssociadaNormalizada] &&
        !(normalizarTexto(categoria) === "contas a pagar" && normalizarTexto(subcategoria) === "pagamento de fatura")) {

      const infoCartao = infoCartoesMap[contaAssociadaNormalizada];
      const dataVencimentoDaTransacao = parseData(linhaTransacao[10]); 

      if (dataVencimentoDaTransacao) {
        const vencimentoMes = dataVencimentoDaTransacao.getMonth();
        const vencimentoAno = dataVencimentoDaTransacao.getFullYear();

        const isTrulyFuture = (vencimentoAno > currentYear) || (vencimentoAno === currentYear && vencimentoMes > currentMonth);

        if (isTrulyFuture) {
          const chaveFatura = `${infoCartao.nomeOriginal}|${vencimentoAno}-${vencimentoMes}`;
          if (!faturasFuturas[chaveFatura]) {
            faturasFuturas[chaveFatura] = {
              cartaoOriginal: infoCartao.nomeOriginal,
              mesVencimento: vencimentoMes,
              anoVencimento: vencimentoAno,
              total: 0
            };
          }
          faturasFuturas[chaveFatura].total = round(faturasFuturas[chaveFatura].total + valor, 2);
          logToSheet(`Transacao '${linhaTransacao[1]}' (ID: ${linhaTransacao[13]}) INCLUIDA em faturas futuras. Vencimento: ${dataVencimentoDaTransacao.toLocaleDateString()}. Fatura futura atual: ${faturasFuturas[chaveFatura].total}`, "DEBUG");
        } else {
          logToSheet(`Transacao '${linhaTransacao[1]}' (ID: ${linhaTransacao[13]}) IGNORADA para faturas futuras. Vencimento (${dataVencimentoDaTransacao.toLocaleDateString()}) não é considerado futuro.`, "DEBUG");
        }
      } else {
        logToSheet(`Vencimento para transacao '${linhaTransacao[1]}' (ID: ${linhaTransacao[13]}) e NULO. Ignorando.`, "WARN");
      }
    }
  }

  let mensagem = `🧾 *Faturas Futuras de Cartao de Credito*\n\n`;
  let temFaturas = false;

  const faturasOrdenadas = Object.values(faturasFuturas).sort((a, b) => {
    if (a.cartaoOriginal !== b.cartaoOriginal) {
      return a.cartaoOriginal.localeCompare(b.cartaoOriginal);
    }
    if (a.anoVencimento !== b.anoVencimento) {
      return a.anoVencimento - b.anoVencimento;
    }
    return a.mesVencimento - b.mesVencimento;
  });

  if (faturasOrdenadas.length === 0) {
    mensagem += "Nenhuma fatura futura lancada alem do proximo ciclo de vencimento.\n";
  } else {
    let currentCard = "";
    faturasOrdenadas.forEach(fatura => {
      if (fatura.total === 0) return;

      temFaturas = true;
      if (fatura.cartaoOriginal !== currentCard) {
        mensagem += `\n*${escapeMarkdown(capitalize(fatura.cartaoOriginal))}:*\n`;
        currentCard = fatura.cartaoOriginal;
      }
      mensagem += `  • ${getNomeMes(fatura.mesVencimento)}/${fatura.anoVencimento}: R$ ${fatura.total.toFixed(2).replace('.', ',')}\n`;
    });
  }

  if (!temFaturas && faturasOrdenadas.length > 0) {
      mensagem = `? *Faturas Futuras de Cartao de Credito*\n\nNenhuma fatura futura lancada alem do proximo ciclo de vencimento com valor positivo.\n`;
  } else if (!temFaturas && faturasOrdenadas.length === 0) {
      mensagem = `🧾 *Faturas Futuras de Cartao de Credito*\n\nNenhuma fatura futura lancada alem do proximo ciclo de vencimento.\n`;
  }


  enviarMensagemTelegram(chatId, mensagem);
  logToSheet(`Faturas futuras enviadas para chatId: ${chatId}.`, "INFO");
}

/**
 * NOVO: Envia o status das contas fixas (Contas_a_Pagar) para o chat do Telegram.
 * Verifica quais contas recorrentes foram pagas no mês e quais estão pendentes.
 * Agora, inclui botões inline para marcar contas como pagas.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário que solicitou.
 * @param {number} mes O mês para verificar (1-12).
 * @param {number} ano O ano para verificar.
 */
function enviarContasAPagar(chatId, usuario, mes, ano) {
  logToSheet(`[ContasAPagar] Iniciando enviarContasAPagar para chatId: ${chatId}, usuario: ${usuario}, Mes: ${mes}, Ano: ${ano}`, "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaContasAPagar = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);
  const abaTransacoes = ss.getSheetByName(SHEET_TRANSACOES);

  if (!abaContasAPagar || !abaTransacoes) {
    enviarMensagemTelegram(chatId, "❌ Erro: As abas 'Contas_a_Pagar' ou 'Transacoes' não foram encontradas. Verifique os nomes das abas.");
    logToSheet("Erro: Abas Contas_a_Pagar ou Transacoes não encontradas.", "ERROR");
    return;
  }

  const dadosContasAPagar = abaContasAPagar.getDataRange().getValues();
  const dadosTransacoes = abaTransacoes.getDataRange().getValues();

  // Obter cabeçalhos das abas para acesso dinâmico às colunas
  const headersContasAPagar = dadosContasAPagar[0];
  const headersTransacoes = dadosTransacoes[0];

  // Mapeamento de índices de coluna para Contas_a_Pagar
  const colID = headersContasAPagar.indexOf('ID');
  const colDescricao = headersContasAPagar.indexOf('Descricao');
  const colCategoria = headersContasAPagar.indexOf('Categoria');
  const colValor = headersContasAPagar.indexOf('Valor');
  const colDataVencimento = headersContasAPagar.indexOf('Data de Vencimento');
  const colStatus = headersContasAPagar.indexOf('Status');
  const colRecorrente = headersContasAPagar.indexOf('Recorrente');
  const colContaSugeria = headersContasAPagar.indexOf('Conta de Pagamento Sugerida');
  const colObservacoes = headersContasAPagar.indexOf('Observacoes');
  const colIDTransacaoVinculada = headersContasAPagar.indexOf('ID Transacao Vinculada');

  // Verificar se todas as colunas essenciais foram encontradas
  if ([colID, colDescricao, colCategoria, colValor, colDataVencimento, colStatus, colRecorrente, colContaSugeria, colObservacoes, colIDTransacaoVinculada].some(idx => idx === -1)) {
    const missingCols = [];
    if (colID === -1) missingCols.push('ID');
    if (colDescricao === -1) missingCols.push('Descricao');
    if (colCategoria === -1) missingCols.push('Categoria');
    if (colValor === -1) missingCols.push('Valor');
    if (colDataVencimento === -1) missingCols.push('Data de Vencimento');
    if (colStatus === -1) missingCols.push('Status');
    if (colRecorrente === -1) missingCols.push('Recorrente');
    if (colContaSugeria === -1) missingCols.push('Conta de Pagamento Sugerida');
    if (colObservacoes === -1) missingCols.push('Observacoes');
    if (colIDTransacaoVinculada === -1) missingCols.push('ID Transacao Vinculada');
    
    enviarMensagemTelegram(chatId, `❌ Erro: Colunas essenciais faltando na aba 'Contas_a_Pagar': ${missingCols.join(', ')}. Verifique os cabeçalhos.`);
    logToSheet(`Erro: Colunas essenciais faltando na aba 'Contas_a_Pagar': ${missingCols.join(', ')}`, "ERROR");
    return;
  }

  const colTransacaoData = headersTransacoes.indexOf('Data');
  const colTransacaoDescricao = headersTransacoes.indexOf('Descricao');
  const colTransacaoTipo = headersTransacoes.indexOf('Tipo');
  const colTransacaoValor = headersTransacoes.indexOf('Valor');
  const colTransacaoCategoria = headersTransacoes.indexOf('Categoria');
  const colTransacaoID = headersTransacoes.indexOf('ID Transacao');

  if ([colTransacaoData, colTransacaoDescricao, colTransacaoTipo, colTransacaoValor, colTransacaoCategoria, colTransacaoID].some(idx => idx === -1)) {
    const missingCols = [];
    if (colTransacaoData === -1) missingCols.push('Data');
    if (colTransacaoDescricao === -1) missingCols.push('Descricao');
    if (colTransacaoTipo === -1) missingCols.push('Tipo');
    if (colTransacaoValor === -1) missingCols.push('Valor');
    if (colTransacaoCategoria === -1) missingCols.push('Categoria');
    if (colTransacaoID === -1) missingCols.push('ID Transacao');

    enviarMensagemTelegram(chatId, `❌ Erro: Colunas essenciais faltando na aba 'Transacoes': ${missingCols.join(', ')}. Verifique os cabeçalhos.`);
    logToSheet(`Erro: Colunas essenciais faltando na aba 'Transacoes': ${missingCols.join(', ')}`, "ERROR");
    return;
  }


  const targetMesIndex = mes - 1;
  const nomeMes = getNomeMes(targetMesIndex);

  let contasFixas = [];
  let contasPagasIds = new Set(); // Para rastrear IDs de contas pagas

  // 1. Carregar contas fixas do mês alvo
  for (let i = 1; i < dadosContasAPagar.length; i++) {
    const linha = dadosContasAPagar[i];
    const dataVencimentoConta = parseData(linha[colDataVencimento]);

    if (!dataVencimentoConta || dataVencimentoConta.getMonth() !== targetMesIndex || dataVencimentoConta.getFullYear() !== ano) {
      continue; // Ignora contas fora do mês/ano alvo
    }

    const idConta = linha[colID];
    const descricao = (linha[colDescricao] || "").toString().trim();
    let valor = parseBrazilianFloat(String(linha[colValor]));
    const recorrente = (linha[colRecorrente] || "").toString().trim().toLowerCase();
    const idTransacaoVinculada = (linha[colIDTransacaoVinculada] || "").toString().trim();
    const statusConta = (linha[colStatus] || "").toString().trim().toLowerCase();

    if (recorrente === "verdadeiro" && idConta && valor > 0) {
      const isPaid = (statusConta === "pago");
      contasFixas.push({
        id: idConta,
        descricao: descricao,
        valor: valor,
        categoria: (linha[colCategoria] || "").toString().trim(),
        paga: isPaid,
        rowIndex: i + 1, // Linha base 1 na planilha
        idTransacaoVinculada: idTransacaoVinculada // Mantém o ID vinculado
      });
      if (isPaid) {
        contasPagasIds.add(idConta);
      }
    }
  }
  logToSheet(`[ContasAPagar] Contas fixas carregadas para ${nomeMes}/${ano}: ${JSON.stringify(contasFixas)}`, "INFO");

  // 2. Tentar vincular transações a contas fixas que ainda não estão pagas
  for (let i = 1; i < dadosTransacoes.length; i++) {
    const linhaTransacao = dadosTransacoes[i];
    const dataTransacao = parseData(linhaTransacao[colTransacaoData]);
    const tipoTransacao = (linhaTransacao[colTransacaoTipo] || "").toString().toLowerCase().trim();
    const descricaoTransacao = (linhaTransacao[colTransacaoDescricao] || "").toString().trim();
    let valorTransacao = parseBrazilianFloat(String(linhaTransacao[colTransacaoValor]));
    const categoriaTransacao = (linhaTransacao[colTransacaoCategoria] || "").toString().trim();
    const idTransacao = (linhaTransacao[colTransacaoID] || "").toString().trim();

    // Filtra transações pelo mês/ano alvo e tipo "despesa"
    if (!dataTransacao || dataTransacao.getMonth() !== targetMesIndex || dataTransacao.getFullYear() !== ano || tipoTransacao !== "despesa") {
      continue;
    }
    logToSheet(`[ContasAPagar] Processando transacao (ID: ${idTransacao}, Desc: "${descricaoTransacao}", Valor: ${valorTransacao.toFixed(2)}) para vinculacao.`, "DEBUG");

    for (let j = 0; j < contasFixas.length; j++) {
      const contaFixa = contasFixas[j];
      if (contaFixa.paga) {
        logToSheet(`[ContasAPagar] Conta fixa "${contaFixa.descricao}" (ID: ${contaFixa.id}) ja esta paga. Pulando.`, "DEBUG");
        continue; // Se já está paga, não precisa tentar vincular novamente
      }

      // Verificação de vínculo manual (se a transação já está vinculada a esta conta)
      if (contaFixa.idTransacaoVinculada === idTransacao) {
        contaFixa.paga = true;
        contasPagasIds.add(contaFixa.id);
        logToSheet(`[ContasAPagar] Conta fixa "${contaFixa.descricao}" (ID: ${contaFixa.id}) marcada como PAGA por vínculo manual com transacao ID: ${idTransacao}.`, "INFO");
        // Atualizar status na planilha
        abaContasAPagar.getRange(contaFixa.rowIndex, colStatus + 1).setValue("Pago");
        // Não precisa atualizar colIDTransacaoVinculada, já está lá
        break; // Encontrou e vinculou, passa para a próxima transação
      }

      // Lógica de auto-vinculação por similaridade
      const descNormalizadaContaFixa = normalizarTexto(contaFixa.descricao);
      const descNormalizadaTransacao = normalizarTexto(descricaoTransacao);
      const categoriaNormalizadaContaFixa = normalizarTexto(contaFixa.categoria);
      const categoriaNormalizadaTransacao = normalizarTexto(categoriaTransacao);

      const similarityScore = calculateSimilarity(descNormalizadaTransacao, descNormalizadaContaFixa);
      const isCategoryMatch = categoriaNormalizadaContaFixa.startsWith(categoriaNormalizadaContaFixa);
      const isValueMatch = Math.abs(valorTransacao - contaFixa.valor) < 0.01; // Tolerância de 1 centavo

      logToSheet(`[ContasAPagar Debug] Comparando Transacao (Desc: "${descricaoTransacao}", Cat: "${categoriaTransacao}", Valor: ${valorTransacao.toFixed(2)}) com Conta Fixa (Desc: "${contaFixa.descricao}", Cat: "${contaFixa.categoria}", Valor: ${contaFixa.valor.toFixed(2)}).`, "DEBUG");
      logToSheet(`[ContasAPagar Debug] Similaridade Descricao: ${similarityScore.toFixed(2)} (Limite: ${SIMILARITY_THRESHOLD}), Categoria Match: ${isCategoryMatch}, Valor Match: ${isValueMatch}.`, "DEBUG");

      if (
        similarityScore >= SIMILARITY_THRESHOLD &&
        isCategoryMatch &&
        isValueMatch
      ) {
        contaFixa.paga = true;
        contasPagasIds.add(contaFixa.id);
        logToSheet(`[ContasAPagar] Conta fixa "${contaFixa.descricao}" (ID: ${contaFixa.id}) marcada como PAGA pela transacao "${descricaoTransacao}" (Valor: R$ ${valorTransacao.toFixed(2)}).`, "INFO");
        
        // Atualiza o status e o ID da transação vinculada na planilha
        abaContasAPagar.getRange(contaFixa.rowIndex, colStatus + 1).setValue("Pago");
        abaContasAPagar.getRange(contaFixa.rowIndex, colIDTransacaoVinculada + 1).setValue(idTransacao);
        logToSheet(`[ContasAPagar] Planilha atualizada para conta fixa ID: ${contaFixa.id}. Status: Pago, ID Transacao Vinculada: ${idTransacao}.`, "INFO");
        break; // Encontrou e vinculou, passa para a próxima transação
      }
    }
  }

  // 3. Construir a mensagem e os botões
  let mensagem = `🧾 *Contas Fixas de ${nomeMes}/${ano}*\n\n`;
  let contasPendentesLista = [];
  let contasPagasLista = [];
  let keyboardButtons = [];

  contasFixas.forEach(conta => {
    if (conta.paga) {
      contasPagasLista.push(`✅ ${escapeMarkdown(capitalize(conta.descricao))}: R$ ${conta.valor.toFixed(2).replace('.', ',')}`);
    } else {
      contasPendentesLista.push(`❌ ${escapeMarkdown(capitalize(conta.descricao))}: R$ ${conta.valor.toFixed(2).replace('.', ',')}`);
      keyboardButtons.push([{
        text: `✅ Marcar '${capitalize(conta.descricao)}' como Pago`,
        callback_data: `/marcar_pago_${conta.id}`
      }]);
    }
  });

  if (contasPagasLista.length > 0) {
    mensagem += `*Contas Pagas:*\n${contasPagasLista.join('\n')}\n\n`;
  } else {
    mensagem += `Nenhuma conta fixa paga encontrada para este mes.\n\n`;
  }

  if (contasPendentesLista.length > 0) {
    mensagem += `*Contas Pendentes:*\n${contasPendentesLista.join('\n')}\n\n`;
  } else {
    mensagem += `Todas as contas fixas foram pagas para este mes! 🎉\n\n`;
  }

  const teclado = { inline_keyboard: keyboardButtons };

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });

  logToSheet(`[ContasAPagar] Status das contas a pagar enviado para chatId: ${chatId}.`, "INFO");
}



/**
 * MODIFICADO: Inicia o processo de edição da última transação do usuário.
 * Agora apenas encontra a transação e chama a função de envio da mensagem.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 */
function iniciarEdicaoUltimo(chatId, usuario) {
  logToSheet(`[Edicao] Iniciando edicao da ultima transacao para ${usuario} (${chatId}).`, "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
  const configSheet = ss.getSheetByName(SHEET_CONFIGURACOES);
  
  if (!transacoesSheet || !configSheet) {
    enviarMensagemTelegram(chatId, "❌ Erro: Abas essenciais não encontradas para edição.");
    return;
  }

  const dadosTransacoes = transacoesSheet.getDataRange().getValues();
  const dadosConfig = configSheet.getDataRange().getValues();

  let ultimaTransacao = null;
  const usuarioNormalizado = normalizarTexto(usuario);
  const grupoUsuarioChat = getGrupoPorChatId(chatId, dadosConfig);

  for (let i = dadosTransacoes.length - 1; i > 0; i--) {
    const linha = dadosTransacoes[i];
    const usuarioLinha = normalizarTexto(linha[11]);
    const grupoTransacao = getGrupoPorChatIdByUsuario(linha[11], dadosConfig);

    if (usuarioLinha === usuarioNormalizado || grupoTransacao === grupoUsuarioChat) {
      ultimaTransacao = {
        linha: i + 1,
        id: linha[13],
        data: Utilities.formatDate(parseData(linha[0]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        descricao: linha[1],
        categoria: linha[2],
        subcategoria: linha[3],
        tipo: linha[4],
        valor: parseBrazilianFloat(String(linha[5])),
        metodoPagamento: linha[6],
        conta: linha[7],
        parcelasTotais: linha[8],
        parcelaAtual: linha[9],
        dataVencimento: Utilities.formatDate(parseData(linha[10]), Session.getScriptTimeZone(), "dd/MM/yyyy"), 
        usuario: linha[11],
        status: linha[12]
      };
      break;
    }
  }

  if (!ultimaTransacao) {
    enviarMensagemTelegram(chatId, "⚠️ Nenhuma transação recente encontrada para você ou seu grupo para editar.");
    return;
  }

  // Chama a nova função reutilizável
  enviarMensagemDeEdicao(chatId, ultimaTransacao);
}


/**
 * NOVO: Solicita ao usuário o novo valor para o campo que ele deseja editar.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} campo O nome do campo a ser editado.
 */
function solicitarNovoValorParaEdicao(chatId, campo) {
  logToSheet(`[Edicao] Solicitando novo valor para campo '${campo}' para ${chatId}.`, "INFO");

  const editState = getEditState(chatId);
  if (!editState || !editState.transactionId) { // Verifica se transactionId existe no estado
    enviarMensagemTelegram(chatId, "⚠️ Sua sessão de edição expirou ou é inválida. Por favor, inicie uma nova edição com `/editar ultimo`.");
    return;
  }

  // Atualiza o estado de edição com o campo a ser editado
  editState.fieldToEdit = campo;
  setEditState(chatId, editState); // Salva o estado atualizado no cache

  let mensagemCampo = "";
  switch (campo) {
    case "data":
      mensagemCampo = "Por favor, envie a *nova data* para o lançamento (formato DD/MM/AAAA):";
      break;
    case "descricao":
      mensagemCampo = "Por favor, envie a *nova descrição* para o lançamento:";
      break;
    case "valor":
      mensagemCampo = "Por favor, envie o *novo valor* para o lançamento (ex: 123.45 ou 123,45):";
      break;
    case "tipo":
      mensagemCampo = "Por favor, envie o *novo tipo* (Despesa, Receita, Transferência):";
      break;
    case "conta":
      mensagemCampo = "Por favor, envie a *nova conta/cartão* para o lançamento:";
      break;
    case "categoria":
      mensagemCampo = "Por favor, envie a *nova categoria* para o lançamento:";
      break;
    case "subcategoria":
      mensagemCampo = "Por favor, envie a *nova subcategoria* para o lançamento:";
      break;
    case "metodoPagamento":
      mensagemCampo = "Por favor, envie o *novo método de pagamento* (ex: Pix, Débito, Crédito):";
      break;
    case "dataVencimento":
        mensagemCampo = "Por favor, envie a *nova data de vencimento* (formato DD/MM/AAAA):";
        break;
    default:
      mensagemCampo = "Campo inválido para edição. Por favor, tente novamente.";
      logToSheet(`[Edicao] Campo '${campo}' inválido solicitado para edição.`, "WARN");
      clearEditState(chatId);
      return;
  }
  
  // Teclado para cancelar edição
  const teclado = {
    inline_keyboard: [
      [{ text: "❌ Cancelar Edição", callback_data: `cancelar_edicao` }]
    ]
  };

  enviarMensagemTelegram(chatId, mensagemCampo, { reply_markup: teclado });
}

/**
 * MODIFICADO: Processa a entrada do usuário para a edição de um campo específico.
 * Após a edição, busca a transação atualizada e pergunta se o usuário quer editar mais algo.
 * AGORA INCLUI A LÓGICA DE APRENDIZADO.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 * @param {string} novoValor O novo valor enviado pelo usuário.
 * @param {Object} editState O estado atual da edição.
 * @param {Array<Array<any>>} dadosContas Dados da aba 'Contas'.
 */
function processarEdicaoFinal(chatId, usuario, novoValor, editState, dadosContas) {
  logToSheet(`[Edicao] Processando edicao final. Transacao ID: ${editState.transactionId}, Campo: ${editState.fieldToEdit}, Novo Valor: "${novoValor}"`, "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
  const dadosPalavras = getSheetDataWithCache(SHEET_PALAVRAS_CHAVE, CACHE_KEY_PALAVRAS);

  if (!transacoesSheet) {
    enviarMensagemTelegram(chatId, "❌ Erro: Aba 'Transacoes' não encontrada para edição.");
    clearEditState(chatId);
    return;
  }

  const allTransactionsData = transacoesSheet.getDataRange().getValues();
  const headers = allTransactionsData[0];
  const colMap = getColumnMap(headers);

  let rowIndex = -1;
  for (let i = 1; i < allTransactionsData.length; i++) {
    if (allTransactionsData[i][colMap["ID Transacao"]] === editState.transactionId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) {
    enviarMensagemTelegram(chatId, "❌ Transação não encontrada ou já excluída.");
    clearEditState(chatId);
    return;
  }

  let colIndex = -1;
  let valorParaSet = novoValor;
  let mensagemSucesso = "";
  let erroValidacao = false;

  const originalDescription = allTransactionsData[rowIndex - 1][colMap["Descricao"]];
  const originalCategory = allTransactionsData[rowIndex - 1][colMap["Categoria"]];
  const originalSubcategory = allTransactionsData[rowIndex - 1][colMap["Subcategoria"]];

  switch (editState.fieldToEdit) {
    case "data":
      colIndex = colMap["Data"];
      const parsedDate = parseData(novoValor);
      if (!parsedDate) {
        mensagemSucesso = "❌ Data inválida. Use o formato DD/MM/AAAA.";
        erroValidacao = true;
      } else {
        valorParaSet = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        mensagemSucesso = "Data atualizada!";
      }
      break;
    case "descricao":
      colIndex = colMap["Descricao"];
      valorParaSet = capitalize(novoValor);
      mensagemSucesso = "Descrição atualizada!";
      break;
    case "valor":
      colIndex = colMap["Valor"];
      const parsedValue = parseBrazilianFloat(novoValor);
      if (isNaN(parsedValue) || parsedValue <= 0) {
        mensagemSucesso = "❌ Valor inválido. Por favor, digite um número positivo (ex: 123.45 ou 123,45).";
        erroValidacao = true;
      } else {
        valorParaSet = parsedValue;
        mensagemSucesso = "Valor atualizado!";
      }
      break;
    case "tipo":
      colIndex = colMap["Tipo"];
      const tipoNormalizado = normalizarTexto(novoValor);
      if (["despesa", "receita", "transferencia"].includes(tipoNormalizado)) {
        valorParaSet = capitalize(tipoNormalizado);
        mensagemSucesso = "Tipo atualizado!";
      } else {
        mensagemSucesso = "❌ Tipo inválido. Use 'Despesa', 'Receita' ou 'Transferência'.";
        erroValidacao = true;
      }
      break;
    case "conta":
      colIndex = colMap["Conta/Cartão"];
      const { conta: detectedAccount } = extrairContaMetodoPagamento(novoValor, dadosContas, dadosPalavras);
      if (detectedAccount && detectedAccount !== "Não Identificada") {
          valorParaSet = detectedAccount;
          mensagemSucesso = "Conta/Cartão atualizado!";
      } else {
          mensagemSucesso = "❌ Conta/Cartão não reconhecido. Por favor, use o nome exato da conta ou um apelido configurado.";
          erroValidacao = true;
      }
      break;
    case "categoria":
      colIndex = colMap["Categoria"];
      const dadosCategorias = getSheetDataWithCache(SHEET_CATEGORIAS, CACHE_KEY_CATEGORIAS);
      const categoriaNormalizadaInput = normalizarTexto(novoValor);
      
      const matchExatoCategoria = dadosCategorias.slice(1).find(row => {
          const { cleanCategory } = _extractIconAndCleanCategory(row[0]);
          return normalizarTexto(cleanCategory) === categoriaNormalizadaInput;
      });

      if (matchExatoCategoria) {
          valorParaSet = matchExatoCategoria[0].trim();
          mensagemSucesso = "Categoria atualizada!";
          // Gatilho de aprendizado
          learnFromCorrection(originalDescription, valorParaSet, originalSubcategory);
      } else {
          const { categoria: detectedCategory } = extrairCategoriaSubcategoria(novoValor, allTransactionsData[rowIndex-1][colMap["Tipo"]], dadosPalavras);
          if (detectedCategory && detectedCategory !== "Não Identificada") {
              valorParaSet = detectedCategory;
              mensagemSucesso = "Categoria atualizada!";
              // Gatilho de aprendizado
              learnFromCorrection(originalDescription, valorParaSet, originalSubcategory);
          } else {
              mensagemSucesso = "❌ Categoria não reconhecida. Por favor, digite um nome de categoria existente ou uma palavra-chave válida.";
              erroValidacao = true;
          }
      }
      break;
    case "subcategoria":
      colIndex = colMap["Subcategoria"];
      const tipoTransacaoOriginal = allTransactionsData[rowIndex-1][colMap["Tipo"]];
      const { categoria: catOriginal, subcategoria: detectedSubcategory } = extrairCategoriaSubcategoria(novoValor, tipoTransacaoOriginal, dadosPalavras);
      if (detectedSubcategory && detectedSubcategory !== "Não Identificada") {
          const currentCategory = allTransactionsData[rowIndex-1][colMap["Categoria"]];
          if (catOriginal && normalizarTexto(catOriginal) !== normalizarTexto(currentCategory)) {
              transacoesSheet.getRange(rowIndex, colMap["Categoria"] + 1).setValue(catOriginal);
              logToSheet(`[Edicao] Categoria atualizada de '${currentCategory}' para '${catOriginal}' ao editar subcategoria.`, "DEBUG");
          }
          valorParaSet = detectedSubcategory;
          mensagemSucesso = "Subcategoria atualizada!";
          // Gatilho de aprendizado
          learnFromCorrection(originalDescription, catOriginal || originalCategory, valorParaSet);
      } else {
          mensagemSucesso = "❌ Subcategoria não reconhecida. Por favor, verifique as palavras-chave da subcategoria.";
          erroValidacao = true;
      }
      break;
    case "metodoPagamento":
      colIndex = colMap["Metodo de Pagamento"];
      const metodoNormalizado = normalizarTexto(novoValor);
      const metodosValidos = ["credito", "debito", "dinheiro", "pix", "boleto", "transferencia bancaria"];
      if (metodosValidos.includes(metodoNormalizado)) {
        valorParaSet = capitalize(metodoNormalizado);
        mensagemSucesso = "Método de pagamento atualizado!";
      } else {
        mensagemSucesso = "❌ Método de pagamento inválido. Use 'Débito', 'Crédito', 'Dinheiro', 'Pix', 'Boleto' ou 'Transferência Bancária'.";
        erroValidacao = true;
      }
      break;
    case "dataVencimento":
      colIndex = colMap["Data de Vencimento"];
      const parsedDueDate = parseData(novoValor);
      if (!parsedDueDate) {
        mensagemSucesso = "❌ Data de vencimento inválida. Use o formato DD/MM/AAAA.";
        erroValidacao = true;
      } else {
        valorParaSet = Utilities.formatDate(parsedDueDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
        mensagemSucesso = "Data de vencimento atualizada!";
      }
      break;
    default:
      mensagemSucesso = "❌ Campo de edição desconhecido.";
      erroValidacao = true;
      break;
  }

  if (erroValidacao) {
    enviarMensagemTelegram(chatId, mensagemSucesso);
    logToSheet(`[Edicao] Erro de validacao para campo '${editState.fieldToEdit}': ${mensagemSucesso}`, "WARN");
    return;
  }

  let lock; 
  try {
    lock = LockService.getScriptLock();
    lock.waitLock(30000);
    
    transacoesSheet.getRange(rowIndex, colIndex + 1).setValue(valorParaSet);
    logToSheet(`[Edicao] Transacao ID ${editState.transactionId} - Campo '${editState.fieldToEdit}' atualizado para: "${valorParaSet}".`, "INFO");
    
    atualizarSaldosDasContas(); 

    const dadosTransacoesAtualizados = transacoesSheet.getDataRange().getValues();
    let transacaoAtualizada = null;
    for (let i = 1; i < dadosTransacoesAtualizados.length; i++) {
        if (dadosTransacoesAtualizados[i][colMap["ID Transacao"]] === editState.transactionId) {
            const linha = dadosTransacoesAtualizados[i];
            transacaoAtualizada = {
                linha: i + 1, id: linha[13], data: Utilities.formatDate(parseData(linha[0]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
                descricao: linha[1], categoria: linha[2], subcategoria: linha[3], tipo: linha[4],
                valor: parseBrazilianFloat(String(linha[5])), metodoPagamento: linha[6], conta: linha[7],
                parcelasTotais: linha[8], parcelaAtual: linha[9], dataVencimento: Utilities.formatDate(parseData(linha[10]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
                usuario: linha[11], status: linha[12]
            };
            break;
        }
    }

    if (transacaoAtualizada) {
        enviarMensagemDeEdicao(chatId, transacaoAtualizada);
    } else {
        enviarMensagemTelegram(chatId, "✅ Alteração salva! Edição finalizada.");
        clearEditState(chatId);
    }

  } catch (e) {
    logToSheet(`ERRO ao atualizar transacao ID ${editState.transactionId}: ${e.message}`, "ERROR");
    enviarMensagemTelegram(chatId, `❌ Houve um erro ao atualizar o lançamento: ${e.message}`);
  } finally {
    if (lock) {
      lock.releaseLock();
    }
  }
}


/**
 * NOVO: Envia um resumo financeiro do mês para um usuário específico.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} solicitante O nome do usuário que solicitou o resumo (pode ser diferente do alvo).
 * @param {string} usuarioAlvo O nome do usuário para quem o resumo é.
 * @param {number} mes O mês para o resumo (1-12).
 * @param {number} ano O ano para o resumo.
 */
function enviarResumoPorPessoa(chatId, solicitante, usuarioAlvo, mes, ano) {
  logToSheet(`[ResumoPessoa] Iniciando resumo para ${usuarioAlvo} (solicitado por ${solicitante}) para ${mes}/${ano}`, "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoes = ss.getSheetByName(SHEET_TRANSACOES).getDataRange().getValues();
  const metasSheet = ss.getSheetByName(SHEET_METAS).getDataRange().getValues();
  const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);

  const mesIndex = mes - 1;
  const nomeMes = getNomeMes(mesIndex);

  let resumoCategorias = {};
  let metasPorCategoria = {};
  let totalReceitasMes = 0;
  let totalDespesasMesExcluindoPagamentosETransferencias = 0;

  // Processamento de Metas (Filtrado por usuário, se a meta for por usuário, o que não parece ser o caso agora)
  // Atualmente, as metas são "familiares". Se quiser metas por pessoa, a aba 'Metas' precisaria de uma coluna 'Usuário'.
  const cabecalhoMetas = metasSheet[2];
  let colMetaMes = -1;
  for (let i = 2; i < cabecalhoMetas.length; i++) {
    const headerValue = String(cabecalhoMetas[i]).toLowerCase();
    const targetHeader = `${nomeMes.toLowerCase()}/${ano}`;
    if (headerValue.includes(targetHeader)) {
      colMetaMes = i;
      break;
    }
  }

  if (colMetaMes !== -1) {
    for (let i = 3; i < metasSheet.length; i++) {
      const categoriaOriginal = (metasSheet[i][0] || "").toString().trim();
      const subcategoriaOriginal = (metasSheet[i][1] || "").toString().trim();
      const valorMetaTexto = metasSheet[i][colMetaMes];

      if (categoriaOriginal && subcategoriaOriginal && valorMetaTexto) {
        const meta = parseBrazilianFloat(String(valorMetaTexto));
        if (!isNaN(meta) && meta > 0) {
          const chaveCategoria = normalizarTexto(categoriaOriginal);
          const chaveSubcategoria = normalizarTexto(`${categoriaOriginal} ${subcategoriaOriginal}`);
          if (!metasPorCategoria[chaveCategoria]) {
            metasPorCategoria[chaveCategoria] = { totalMeta: 0, totalGasto: 0, subcategories: {} };
          }
          metasPorCategoria[chaveCategoria].subcategories[chaveSubcategoria] = { meta: meta, gasto: 0 };
          metasPorCategoria[chaveCategoria].totalMeta += meta;
        }
      }
    }
  }

  // Processamento de Transações (Filtrado por usuário alvo)
  for (let i = 1; i < transacoes.length; i++) {
    const dataRaw = transacoes[i][0];
    const data = parseData(dataRaw);
    const tipo = transacoes[i][4];
    let valor = parseBrazilianFloat(String(transacoes[i][5]));
    const categoria = transacoes[i][2];
    const subcategoria = transacoes[i][3];
    const usuarioTransacao = transacoes[i][11];

    if (!data || data.getMonth() !== mesIndex || data.getFullYear() !== ano || normalizarTexto(usuarioTransacao) !== normalizarTexto(usuarioAlvo)) {
      continue;
    }

    // Mesma lógica de fluxo de caixa que em gerarResumoMensal
    if (tipo === "Receita") {
        const categoriaNormalizada = normalizarTexto(categoria);
        const subcategoriaNormalizada = normalizarTexto(subcategoria);
        if (!(categoriaNormalizada === "transferencias" && subcategoriaNormalizada === "entre contas") &&
            !(categoriaNormalizada === "pagamentos recebidos" && subcategoriaNormalizada === "pagamento de fatura")) {
            totalReceitasMes += valor;
        }
    } else if (tipo === "Despesa") {
        const categoriaNormalizada = normalizarTexto(categoria);
        const subcategoriaNormalizada = normalizarTexto(subcategoria);
        if (!(categoriaNormalizada === "contas a pagar" && subcategoriaNormalizada === "pagamento de fatura") &&
            !(categoriaNormalizada === "transferencias" && subcategoriaNormalizada === "entre contas")) {
            totalDespesasMesExcluindoPagamentosETransferencias += valor;
            // Para metas e detalhe de categoria, usar Data de Vencimento
            const dataVencimentoRaw = transacoes[i][10]; 
            const dataVencimento = parseData(dataVencimentoRaw);

            if (dataVencimento && dataVencimento.getMonth() === mesIndex && dataVencimento.getFullYear() === ano) {
              if (!resumoCategorias[categoria]) {
                resumoCategorias[categoria] = { total: 0, subcategories: {} };
              }
              resumoCategorias[categoria].total += valor;
              if (!resumoCategorias[categoria].subcategories[subcategoria]) {
                resumoCategorias[categoria].subcategories[subcategoria] = 0;
              }
              resumoCategorias[categoria].subcategories[subcategoria] += valor;

              const chaveCategoriaMeta = normalizarTexto(categoria);
              const chaveSubcategoriaMeta = normalizarTexto(`${categoria} ${subcategoria}`);
              if (metasPorCategoria[chaveCategoriaMeta] && metasPorCategoria[chaveCategoriaMeta].subcategories[chaveSubcategoriaMeta]) {
                metasPorCategoria[chaveCategoriaMeta].subcategories[chaveSubcategoriaMeta].gasto += valor;
                metasPorCategoria[chaveCategoriaMeta].totalGasto += valor;
              }
            }
        }
    }
  }

  let mensagemResumo = `📊 *Resumo Financeiro de ${nomeMes}/${ano} - ${escapeMarkdown(capitalize(usuarioAlvo))}*\n\n`;

  mensagemResumo += `*💰 Fluxo de Caixa do Mes*\n`;
  mensagemResumo += `• *Receitas Totais:* R$ ${totalReceitasMes.toFixed(2).replace('.', ',')}\n`;
  mensagemResumo += `• *Despesas Totais (excluindo pagamentos de fatura e transferencias):* R$ ${totalDespesasMesExcluindoPagamentosETransferencias.toFixed(2).replace('.', ',')}\n`;
  const saldoLiquidoMes = totalReceitasMes - totalDespesasMesExcluindoPagamentosETransferencias;
  let emojiSaldo = "⚖️";
  if (saldoLiquidoMes > 0) emojiSaldo = "✅";
  else if (saldoLiquidoMes < 0) emojiSaldo = "❌";
  mensagemResumo += `• *Saldo Liquido do Mes:* ${emojiSaldo} R$ ${saldoLiquidoMes.toFixed(2).replace('.', ',')}\n\n`;

  mensagemResumo += `*💸 Despesas Detalhadas por Categoria*\n`;
  const categoriasOrdenadas = Object.keys(resumoCategorias).sort((a, b) => resumoCategorias[b].total - resumoCategorias[a].total);

  if (categoriasOrdenadas.length === 0) {
      mensagemResumo += "Nenhuma despesa detalhada registrada para este usuario neste mes.\n";
  } else {
      categoriasOrdenadas.forEach(categoria => {
          const totalCategoria = resumoCategorias[categoria].total;
          const metaInfo = metasPorCategoria[normalizarTexto(categoria)] || { totalMeta: 0, totalGasto: 0, subcategories: {} };
          
          mensagemResumo += `\n*${escapeMarkdown(capitalize(categoria))}:* R$ ${totalCategoria.toFixed(2).replace('.', ',')}`;
          if (metaInfo.totalMeta > 0) {
            const percMeta = metaInfo.totalMeta > 0 ? (metaInfo.gasto / metaInfo.meta) * 100 : 0;
            let emojiMeta = "";
            if (percMeta >= 100) emojiMeta = "⛔";
            else if (percMeta >= 80) emojiMeta = "⚠️";
            else emojiMeta = "✅";
            mensagemResumo += ` ${emojiMeta} (${percMeta.toFixed(0)}% da meta de R$ ${metaInfo.totalMeta.toFixed(2).replace('.', ',')})`;
          }
          mensagemResumo += `\n`;

          const subcategoriasOrdenadas = Object.keys(resumoCategorias[categoria].subcategories).sort((a, b) => resumoCategorias[categoria].subcategories[b] - resumoCategorias[categoria].subcategories[a]);
          subcategoriasOrdenadas.forEach(sub => {
            const gastoSub = resumoCategorias[categoria].subcategories[sub];
            const chaveSubcategoriaMeta = normalizarTexto(`${categoria} ${sub}`);
            const subMetaInfo = metasPorCategoria[normalizarTexto(categoria)]?.subcategories[chaveSubcategoriaMeta];

            let subLine = `  • ${escapeMarkdown(capitalize(sub))}: R$ ${gastoSub.toFixed(2).replace('.', ',')}`;
            if (subMetaInfo && subMetaInfo.meta > 0) {
              let subEmoji = "";
              let subPerc = (subMetaInfo.gasto / subMetaInfo.meta) * 100;
              if (subPerc >= 100) subEmoji = "⛔";
              else if (subPerc >= 80) subEmoji = "⚠️";
              else subEmoji = "✅";
              subLine += ` / R$ ${subMetaInfo.meta.toFixed(2).replace('.', ',')} ${subEmoji} ${subPerc.toFixed(0)}%`;
            }
            mensagemResumo += `${subLine}\n`;
          });
      });
  }

  enviarMensagemTelegram(chatId, mensagemResumo);
  logToSheet(`Resumo por pessoa enviado para ${chatId} para o usuário ${usuarioAlvo}.`, "INFO");
}


// ===================================================================================
// SEÇÃO DE CHECK-UP FINANCEIRO (/saude)
// ===================================================================================

/**
 * @private
 * Calcula o fluxo de caixa do mês (receitas, despesas, necessidades, desejos).
 * @param {Array<Array<any>>} transacoes - Dados da aba de transações.
 * @param {Object} categoriasMap - Mapa de categorias gerado por getCategoriesMap.
 * @param {number} mesAtual - O mês atual (0-11).
 * @param {number} anoAtual - O ano atual.
 * @returns {Object} Um objeto com os totais calculados.
 */
function _calculateMonthlyFlow(transacoes, categoriasMap, mesAtual, anoAtual) {
  let receitasMes = 0, despesasMes = 0, gastoNecessidades = 0, gastoDesejos = 0;

  for (let i = 1; i < transacoes.length; i++) {
    const linha = transacoes[i];
    const dataTransacao = parseData(linha[0]);
    
    if (dataTransacao && dataTransacao.getMonth() === mesAtual && dataTransacao.getFullYear() === anoAtual) {
      const tipo = (linha[4] || "").toLowerCase();
      const categoria = linha[2];
      const subcategoria = linha[3];
      const valor = parseBrazilianFloat(String(linha[5]));

      const isIgnored = (normalizarTexto(categoria) === "transferencias" && normalizarTexto(subcategoria) === "entre contas") ||
                        (normalizarTexto(categoria) === "contas a pagar" && normalizarTexto(subcategoria) === "pagamento de fatura");

      if (!isIgnored) {
        if (tipo === "receita") {
          receitasMes += valor;
        } else if (tipo === "despesa") {
          despesasMes += valor;
          const categoriaInfo = categoriasMap[normalizarTexto(categoria)];
          if (categoriaInfo && categoriaInfo.tipoGasto === 'necessidade') gastoNecessidades += valor;
          else if (categoriaInfo && categoriaInfo.tipoGasto === 'desejo') gastoDesejos += valor;
        }
      }
    }
  }
  return { receitasMes, despesasMes, gastoNecessidades, gastoDesejos };
}

/**
 * @private
 * Calcula o total de contas a pagar recorrentes do mês.
 * @param {Array<Array<any>>} contasAPagar - Dados da aba Contas_a_Pagar.
 * @param {number} mesAtual - O mês atual (0-11).
 * @param {number} anoAtual - O ano atual.
 * @returns {number} O valor total das contas a pagar.
 */
function _calculateCommittedBills(contasAPagar, mesAtual, anoAtual) {
  let totalContasAPagarMes = 0;
  for (let i = 1; i < contasAPagar.length; i++) {
    const linha = contasAPagar[i];
    const dataVencimento = parseData(linha[4]);
    if (dataVencimento && dataVencimento.getMonth() === mesAtual && dataVencimento.getFullYear() === anoAtual) {
      totalContasAPagarMes += parseBrazilianFloat(String(linha[3]));
    }
  }
  return totalContasAPagarMes;
}

/**
 * @private
 * Obtém o valor total das faturas de cartão de crédito do ciclo atual.
 * @returns {number} O valor total das faturas.
 */
function _getCreditCardBills() {
  atualizarSaldosDasContas();
  let totalFaturasPagar = 0;
  for (const nomeNormalizado in globalThis.saldosCalculados) {
      const infoConta = globalThis.saldosCalculados[nomeNormalizado];
      if (infoConta.tipo === "cartão de crédito" || infoConta.tipo === "fatura consolidada") {
        totalFaturasPagar += infoConta.faturaAtual;
      }
  }
  return totalFaturasPagar;
}

/**
 * @private
 * Formata a mensagem final do check-up financeiro.
 * @param {Object} data - Um objeto com todos os dados calculados.
 * @returns {string} A mensagem formatada para o Telegram.
 */
function _formatSaudeFinanceiraMessage(data) {
  const { nomeMes, anoAtual, receitasMes, despesasMes, gastoNecessidades, gastoDesejos, rendimentoComprometidoTotal } = data;

  const saldoLiquido = receitasMes - despesasMes;
  const taxaDePoupanca = receitasMes > 0 ? (saldoLiquido / receitasMes) * 100 : 0;
  const percNecessidades = receitasMes > 0 ? (gastoNecessidades / receitasMes) * 100 : 0;
  const percDesejos = receitasMes > 0 ? (gastoDesejos / receitasMes) * 100 : 0;
  const percRendimentoComprometido = receitasMes > 0 ? (rendimentoComprometidoTotal / receitasMes) * 100 : 0;
  const gastoDiarioMedio = despesasMes / (new Date()).getDate();

  let mensagem = `🩺 *Check-up Financeiro de ${nomeMes} de ${anoAtual}*\n\n`;

  mensagem += `*📊 Análise de Gastos (vs. Rendimento)*\n`;
  mensagem += `● *Necessidades:* ${percNecessidades.toFixed(0)}% (${formatCurrency(gastoNecessidades)})\n`;
  mensagem += `● *Desejos:* ${percDesejos.toFixed(0)}% (${formatCurrency(gastoDesejos)})\n`;
  mensagem += `● *Poupança:* ${taxaDePoupanca.toFixed(0)}% (${formatCurrency(saldoLiquido)})\n`;
  mensagem += `_Meta ideal: 50% Necessidades, 30% Desejos, 20% Poupança._\n\n`;

  let emojiPoupanca = "🟠";
  if (taxaDePoupanca >= 20) emojiPoupanca = "🟢";
  else if (taxaDePoupanca < 0) emojiPoupanca = "🔴";
  mensagem += `● *Taxa de Poupança:* ${emojiPoupanca} ${taxaDePoupanca.toFixed(0)}%\n_${escapeMarkdown(`Você está a poupar ${taxaDePoupanca.toFixed(0)}% do que ganha.`)}_\n\n`;

  let emojiComprometido = "🟢";
  if (percRendimentoComprometido > 50) emojiComprometido = "🔴";
  else if (percRendimentoComprometido > 30) emojiComprometido = "🟠";
  mensagem += `● *Rendimento Comprometido:* ${emojiComprometido} ${percRendimentoComprometido.toFixed(0)}%\n_${escapeMarkdown(`${percRendimentoComprometido.toFixed(0)}% do seu rendimento está alocado a faturas e contas a pagar.`)}_\n\n`;

  mensagem += `● *Gasto Diário Médio:* 💸 ${formatCurrency(gastoDiarioMedio)}\n_Até agora, este é o seu gasto médio por dia neste mês._`;

  return mensagem;
}

/**
 * **FUNÇÃO ATUALIZADA E ORGANIZADA**
 * CALCULA E ENVIA UM CHECK-UP DA SAÚDE FINANCEIRA DO MÊS ATUAL.
 * Agora inclui a análise de Necessidades vs. Desejos.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 */
function enviarSaudeFinanceira(chatId, usuario) {
  logToSheet(`Iniciando Check-up Financeiro para ${usuario} (${chatId}).`, "INFO");

  try {
    // 1. Obter Dados e Variáveis de Tempo
    const hoje = new Date();
    const mesAtual = hoje.getMonth();
    const anoAtual = hoje.getFullYear();
    const transacoes = getSheetDataWithCache(SHEET_TRANSACOES, CACHE_KEY_TRANSACOES, 60);
    const contasAPagar = getSheetDataWithCache(SHEET_CONTAS_A_PAGAR, CACHE_KEY_CONTAS_A_PAGAR, 300);
    const categoriasMap = getCategoriesMap();

    // 2. Calcular os Componentes Principais usando Funções Auxiliares
    const { receitasMes, despesasMes, gastoNecessidades, gastoDesejos } = _calculateMonthlyFlow(transacoes, categoriasMap, mesAtual, anoAtual);
    const totalContasAPagarMes = _calculateCommittedBills(contasAPagar, mesAtual, anoAtual);
    const totalFaturasPagar = _getCreditCardBills();
    const rendimentoComprometidoTotal = totalContasAPagarMes + totalFaturasPagar;

    // 3. Montar o Pacote de Dados para a Mensagem
    const messageData = {
      nomeMes: getNomeMes(mesAtual),
      anoAtual: anoAtual,
      receitasMes,
      despesasMes,
      gastoNecessidades,
      gastoDesejos,
      rendimentoComprometidoTotal
    };

    // 4. Formatar e Enviar a Mensagem
    const mensagem = _formatSaudeFinanceiraMessage(messageData);
    enviarMensagemTelegram(chatId, mensagem);

  } catch (e) {
    handleError(e, `enviarSaudeFinanceira para ${usuario}`, chatId);
  }
}

/**
 * NOVA FUNÇÃO AUXILIAR
 * Cria um mapa de categorias para fácil acesso ao tipo e tipo de gasto.
 * @returns {Object} Um mapa no formato { 'nome_categoria_normalizado': { tipo: 'Despesa', tipoGasto: 'Necessidade' } }.
 */
function getCategoriesMap() {
  const categoriasData = getSheetDataWithCache(SHEET_CATEGORIAS, CACHE_KEY_CATEGORIAS);
  const map = {};
  // Assumindo que o cabeçalho está na primeira linha
  for (let i = 1; i < categoriasData.length; i++) {
    const row = categoriasData[i];
    const categoria = (row[0] || "").trim();
    const tipo = (row[2] || "").toLowerCase().trim();
    const tipoGasto = (row[3] || "").toLowerCase().trim(); // Nova coluna D
    
    if (categoria) {
      const categoriaNormalizada = normalizarTexto(categoria);
      if (!map[categoriaNormalizada]) {
        map[categoriaNormalizada] = {
          tipo: tipo,
          tipoGasto: tipoGasto // 'necessidade' ou 'desejo'
        };
      }
    }
  }
  return map;
}



/**
 * NOVO: Envia a mensagem de edição para uma transação específica.
 * Reutilizável para iniciar a edição e para continuar após uma alteração.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {Object} transacao O objeto completo da transação a ser editada.
 */
function enviarMensagemDeEdicao(chatId, transacao) {
  // Armazena o estado da edição no cache
  setEditState(chatId, {
    transactionId: transacao.id,
    rowIndex: transacao.linha,
    originalData: transacao // Armazena a transação completa
  });

  const mensagem = `✏️ *Editando o lançamento* (ID: \`${escapeMarkdown(transacao.id)}\`):\n\n` +
                   `*Data:* ${transacao.data}\n` +
                   `*Descricao:* ${escapeMarkdown(transacao.descricao)}\n` +
                   `*Valor:* ${formatCurrency(transacao.valor)}\n` +
                   `*Tipo:* ${transacao.tipo}\n` +
                   `*Conta:* ${escapeMarkdown(transacao.conta)}\n` +
                   `*Categoria:* ${escapeMarkdown(transacao.categoria)}\n` +
                   `*Subcategoria:* ${escapeMarkdown(transacao.subcategoria)}\n` +
                   `*Metodo:* ${escapeMarkdown(transacao.metodoPagamento)}\n` +
                   `*Vencimento:* ${transacao.dataVencimento}\n\n` +
                   `Qual campo deseja editar? Ou clique em 'Finalizar'.`;

  const teclado = {
    inline_keyboard: [
      [{ text: "Data", callback_data: `edit_data` },
       { text: "Descrição", callback_data: `edit_descricao` }],
      [{ text: "Valor", callback_data: `edit_valor` },
       { text: "Tipo", callback_data: `edit_tipo` }],
      [{ text: "Conta/Cartão", callback_data: `edit_conta` },
       { text: "Categoria", callback_data: `edit_categoria` }],
      [{ text: "Subcategoria", callback_data: `edit_subcategoria` },
       { text: "Método Pgto", callback_data: `edit_metodoPagamento` }],
      [{ text: "Data Vencimento", callback_data: `edit_dataVencimento` }],
      [{ text: "✅ Finalizar Edição", callback_data: `cancelar_edicao` }] // Reutilizamos o callback de cancelar para finalizar
    ]
  };

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
}

/**
 * NOVO: Lida com o comando /orcamento, buscando e formatando o progresso do orçamento de despesas do mês.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 * @param {number} mes O mês para o resumo (1-12).
 * @param {number} ano O ano para o resumo.
 */
function handleOrcamentoCommand(chatId, usuario, mes, ano) {
  try {
    // Esta função busca os dados do orçamento. Ela será adicionada em Budget.gs
    const progresso = getBudgetProgressForTelegram(mes, ano); 
    
    if (!progresso || progresso.length === 0) {
      return enviarMensagemTelegram(chatId, `📊 Nenhum orçamento de despesas definido para ${getNomeMes(mes - 1)}/${ano}.`);
    }

    let message = `📊 *Progresso do Orçamento de ${getNomeMes(mes - 1)}/${ano}*\n\n`;
    progresso.forEach(item => {
      const emojiStatus = item.percentage > 100 ? '❗️' : (item.percentage > 85 ? '⚠️' : '✅');
      const gastoFormatado = formatCurrency(item.gasto);
      const orcadoFormatado = formatCurrency(item.orcado);
      const barra = criarBarraDeProgresso(item.percentage);

      message += `${item.icon || '🔹'} *${item.categoria}*\n`;
      message += `${barra} ${item.percentage.toFixed(1)}%\n`;
      message += `_${gastoFormatado} de ${orcadoFormatado}_ ${emojiStatus}\n\n`;
    });

    enviarMensagemTelegram(chatId, message);
  } catch (e) {
    handleError(e, "handleOrcamentoCommand", chatId);
  }
}

/**
 * NOVO: Lida com o comando /statusmetas, buscando e formatando o progresso das metas de poupança.
 * @param {string} chatId O ID do chat do Telegram.
 */
function handleMetasCommand(chatId) {
    try {
        // Esta nova função busca os dados da sua nova aba "Metas". Será adicionada em Budget.gs
        const statusMetas = getGoalsStatusForTelegram(); 
        
        if (!statusMetas || statusMetas.length === 0) {
            return enviarMensagemTelegram(chatId, "🎯 Nenhuma meta de poupança definida. Crie uma na sua planilha na aba 'Metas'!");
        }

        let message = "🎯 *Progresso das Suas Metas de Poupança*\n\n";
        statusMetas.forEach(meta => {
            const salvoFormatado = formatCurrency(meta.salvo);
            const objetivoFormatado = formatCurrency(meta.objetivo);
            const barra = criarBarraDeProgresso(meta.percentage);

            message += `*${meta.nome}*\n`;
            message += `${barra} ${meta.percentage.toFixed(1)}%\n`;
            message += `_${salvoFormatado} de ${objetivoFormatado}_\n\n`;
        });

        enviarMensagemTelegram(chatId, message);
    } catch (e) {
        handleError(e, "handleStatusMetasCommand", chatId);
    }
}

/**
 * NOVO (ou para garantir que existe): Cria uma barra de progresso com emojis.
 * @param {number} percentage A percentagem de progresso.
 * @returns {string} A barra de progresso formatada.
 */
function criarBarraDeProgresso(percentage) {
    const totalBlocks = 10;
    const filledBlocks = Math.round(Math.min(percentage, 100) / 100 * totalBlocks);
    const emptyBlocks = totalBlocks - filledBlocks;
    return '▓'.repeat(filledBlocks) + '░'.repeat(emptyBlocks);
}

// ===================================================================================
// ### INÍCIO DA ATUALIZAÇÃO: NOVAS FUNÇÕES DE GESTÃO DE METAS ###
// ===================================================================================

/**
 * Lida com o comando /novameta para criar um novo objetivo de poupança.
 * Formato esperado: /novameta [Nome da Meta] [Valor Objetivo]
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} complemento O texto que segue o comando.
 */
function handleNovaMetaCommand(chatId, complemento) {
  try {
    // Extrai o nome e o valor do complemento
    const match = complemento.match(/(.+)\s+([\d.,]+)$/);
    if (!match) {
      enviarMensagemTelegram(chatId, "❌ Formato inválido. Use: `/novameta NOME DA META VALOR`\nExemplo: `/novameta Viagem ao Japão 15000`");
      return;
    }

    const nomeMeta = match[1].trim();
    const valorObjetivo = parseBrazilianFloat(match[2]);

    if (!nomeMeta || isNaN(valorObjetivo) || valorObjetivo <= 0) {
      enviarMensagemTelegram(chatId, "❌ Dados inválidos. Certifique-se de que o nome não está vazio e o valor é um número positivo.");
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const metasSheet = ss.getSheetByName(SHEET_METAS);
    if (!metasSheet) {
      throw new Error("Aba 'Metas' não encontrada.");
    }

    // Adiciona a nova meta na planilha
    metasSheet.appendRow([
      nomeMeta,
      valorObjetivo,
      0, // Valor Salvo inicial
      '', // Data Alvo (opcional)
      'Em Andamento' // Status
    ]);

    logToSheet(`Nova meta criada: '${nomeMeta}' com objetivo de R$ ${valorObjetivo.toFixed(2)}`, "INFO");
    enviarMensagemTelegram(chatId, `✅ Nova meta criada com sucesso!\n\n🎯 *${escapeMarkdown(nomeMeta)}*\n*Objetivo:* ${formatCurrency(valorObjetivo)}`);

  } catch (e) {
    handleError(e, "handleNovaMetaCommand", chatId);
  }
}

/**
 * ATUALIZADA: Lida com o comando /aportarmeta para adicionar valor a uma meta de poupança.
 * Corrige o erro onde o nome do usuário era gravado incorretamente na aba 'Transacoes'.
 * Formato esperado: /aportarmeta [Palavra-chave da Meta] [Valor] de [Conta de Origem]
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} complemento O texto que segue o comando.
 * @param {string} usuario O nome do usuário que fez o aporte.
 */
function handleAportarMetaCommand(chatId, complemento, usuario) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Extrai a palavra-chave da meta, o valor e a conta de origem
    const match = complemento.match(/(.+?)\s+([\d.,]+)\s+(?:de|do|da)\s+(.+)/i);
    if (!match) {
      enviarMensagemTelegram(chatId, "❌ Formato inválido. Use: `/aportarmeta META VALOR de CONTA`\nExemplo: `/aportarmeta Japão 500 do Nubank`");
      return;
    }

    const metaKeyword = normalizarTexto(match[1].trim());
    const valorAporte = parseBrazilianFloat(match[2]);
    const nomeContaOrigem = match[3].trim();

    if (!metaKeyword || isNaN(valorAporte) || valorAporte <= 0 || !nomeContaOrigem) {
      enviarMensagemTelegram(chatId, "❌ Dados inválidos. Verifique a palavra-chave da meta, o valor e o nome da conta.");
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const metasSheet = ss.getSheetByName(SHEET_METAS);
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
    const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);

    // Encontra a conta de origem
    const contaOrigemInfo = obterInformacoesDaConta(nomeContaOrigem, dadosContas);
    if (!contaOrigemInfo) {
      enviarMensagemTelegram(chatId, `❌ Conta de origem "${escapeMarkdown(nomeContaOrigem)}" não encontrada.`);
      return;
    }
    const nomeRealConta = contaOrigemInfo.nomeOriginal;

    // Encontra a meta na planilha
    const dadosMetas = metasSheet.getDataRange().getValues();
    const headers = dadosMetas[0];
    const colMap = getColumnMap(headers);
    
    let rowIndex = -1;
    let nomeRealMeta = "";
    for (let i = 1; i < dadosMetas.length; i++) {
      const nomeMetaAtual = dadosMetas[i][colMap['Nome da Meta']];
      if (normalizarTexto(nomeMetaAtual).includes(metaKeyword)) {
        rowIndex = i + 1;
        nomeRealMeta = nomeMetaAtual;
        break;
      }
    }

    if (rowIndex === -1) {
      enviarMensagemTelegram(chatId, `❌ Nenhuma meta encontrada com a palavra-chave "${escapeMarkdown(metaKeyword)}".`);
      return;
    }

    // Atualiza o valor salvo na aba Metas
    const valorSalvoAtual = parseBrazilianFloat(String(dadosMetas[rowIndex - 1][colMap['Valor Salvo']] || '0'));
    const novoValorSalvo = valorSalvoAtual + valorAporte;
    metasSheet.getRange(rowIndex, colMap['Valor Salvo'] + 1).setValue(novoValorSalvo);

    // ### INÍCIO DA CORREÇÃO ###
    // Regista a transação com o 'usuario' na coluna correta (índice 11)
    const idTransacao = Utilities.getUuid();
    const hoje = new Date();
    transacoesSheet.appendRow([
      hoje,                             // Data (A)
      `Aporte Meta: ${nomeRealMeta}`,  // Descricao (B)
      '📈 Investimentos / Futuro',     // Categoria (C)
      'Aporte em Meta',                 // Subcategoria (D)
      'Despesa',                        // Tipo (E)
      valorAporte,                      // Valor (F)
      'Transferência',                  // Metodo de Pagamento (G)
      nomeRealConta,                    // Conta/Cartão (H)
      1,                                // Parcelas Totais (I)
      1,                                // Parcela Atual (J)
      hoje,                             // Data de Vencimento (K)
      usuario,                          // Usuario (L) <-- CORRIGIDO
      'Ativo',                          // Status (M)
      idTransacao,                      // ID Transacao (N)
      hoje                              // Data de Registro (O)
    ]);
    // ### FIM DA CORRECIÇÃO ###
    
    atualizarSaldosDasContas();

    logToSheet(`Aporte de R$ ${valorAporte.toFixed(2)} realizado na meta '${nomeRealMeta}' a partir da conta '${nomeRealConta}'.`, "INFO");
    
    const valorObjetivo = parseBrazilianFloat(String(dadosMetas[rowIndex - 1][colMap['Valor Objetivo']] || '0'));
    const percentualConcluido = valorObjetivo > 0 ? (novoValorSalvo / valorObjetivo) * 100 : 0;

    enviarMensagemTelegram(chatId, `✅ Aporte de ${formatCurrency(valorAporte)} registado com sucesso!\n\n` +
                                   `🎯 *${escapeMarkdown(nomeRealMeta)}*\n` +
                                   `${criarBarraDeProgresso(percentualConcluido)} ${percentualConcluido.toFixed(1)}%\n` +
                                   `*Salvo:* ${formatCurrency(novoValorSalvo)} / ${formatCurrency(valorObjetivo)}`);

  } catch (e) {
    handleError(e, "handleAportarMetaCommand", chatId);
  } finally {
    lock.releaseLock();
  }
}

// ===================================================================================
// ### FIM DA ATUALIZAÇÃO ###
// ===================================================================================


/**
 * **ATUALIZADO COM LOCKSERVICE**
 * Processa o comando /gasto, /compra, etc., para adicionar uma nova despesa.
 * @param {string} chatId O ID do chat do usuário.
 * @param {string} text O texto da mensagem.
 * @param {string} command O comando que iniciou a chamada (ex: "gasto").
 */
function processarComandoGasto(chatId, text, command) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // ... (o código interno da função permanece o mesmo)
    const dados = extrairDadosGasto(text, command);
    if (!dados) {
      sendMessage(chatId, `Formato inválido. Use: \`/${command} <valor> <descrição>\``);
      return;
    }
    adicionarTransacao(chatId, dados.valor, dados.descricao, "Despesa");
    atualizarSaldosDasContas();
  } catch (e) {
    handleError(e, "processarComandoGasto", chatId);
  } finally {
    lock.releaseLock();
  }
}

/**
 * **ATUALIZADO COM LOCKSERVICE**
 * Processa o comando /receita, /ganhei, etc., para adicionar uma nova receita.
 * @param {string} chatId O ID do chat do usuário.
 * @param {string} text O texto da mensagem.
 * @param {string} command O comando que iniciou a chamada (ex: "receita").
 */
function processarComandoReceita(chatId, text, command) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // ... (o código interno da função permanece o mesmo)
    const dados = extrairDadosGasto(text, command);
    if (!dados) {
      sendMessage(chatId, `Formato inválido. Use: \`/${command} <valor> <descrição>\``);
      return;
    }
    adicionarTransacao(chatId, dados.valor, dados.descricao, "Receita");
    atualizarSaldosDasContas();
  } catch (e) {
    handleError(e, "processarComandoReceita", chatId);
  } finally {
    lock.releaseLock();
  }
}

/**
 * **ATUALIZADO COM LOCKSERVICE**
 * Processa o comando /excluir para remover a última transação ou uma transação específica.
 * @param {string} chatId O ID do chat do usuário.
 * @param {string} text O texto completo da mensagem do usuário.
 */
function processarComandoExcluir(chatId, text) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_TRANSACOES);
    if (!sheet) {
      sendMessage(chatId, "❌ A aba de transações não foi encontrada.");
      return;
    }

    const range = sheet.getDataRange();
    const values = range.getValues();
    
    if (values.length < 2) {
      sendMessage(chatId, "ℹ️ Não há nenhuma transação para excluir.");
      return;
    }

    const lastRowIndex = values.length;
    const lastRowData = values[lastRowIndex - 1];
    
    const descricao = lastRowData[1];
    const valor = formatarMoeda(lastRowData[5]);
    const data = lastRowData[0];

    if (descricao.startsWith("Aporte Meta:")) {
      const metaName = descricao.substring("Aporte Meta:".length).trim();
      reverterAporteMeta(metaName, parseBrazilianFloat(String(lastRowData[5])));
    }

    const transactionId = lastRowData[13];
    if (transactionId) {
      reverterStatusContaAPagarSeVinculado(transactionId);
    }

    sheet.deleteRow(lastRowIndex);
    
    atualizarSaldosDasContas();

    const mensagem = `✅ Transação excluída com sucesso:\n\n*Descrição:* ${descricao}\n*Valor:* ${valor}\n*Data:* ${data}`;
    sendMessage(chatId, mensagem);
    logToSheet(`Transação excluída para ${chatId}: ${descricao}`, "INFO");

  } catch (e) {
    handleError(e, "processarComandoExcluir", chatId);
  } finally {
    lock.releaseLock();
  }
}

/**
 * **ATUALIZADO COM LOCKSERVICE**
 * Processa o comando /transferir para mover valores entre contas.
 * @param {string} chatId O ID do chat do usuário.
 * @param {string} text O texto da mensagem.
 */
function processarComandoTransferencia(chatId, text) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // ... (o código interno da função permanece o mesmo)
    const regex = /transferir\s+([\d,.]+)\s+de\s+(.+?)\s+para\s+(.+)/i;
    const match = text.match(regex);

    if (!match) {
        sendMessage(chatId, "Formato inválido. Use: `/transferir <valor> de <conta_origem> para <conta_destino>`");
        return;
    }

    const valor = parseBrazilianFloat(match[1]);
    const contaOrigem = match[2].trim();
    const contaDestino = match[3].trim();

    adicionarTransferencia(chatId, valor, contaOrigem, contaDestino);
    atualizarSaldosDasContas();
  } catch (e) {
    handleError(e, "processarComandoTransferencia", chatId);
  } finally {
    lock.releaseLock();
  }
}

/**
 * NOVO: Envia o Patrimônio Líquido para o chat do Telegram.
 * @param {string} chatId O ID do chat do Telegram.
 */
function enviarPatrimonioLiquido(chatId) {
  try {
    const netWorthData = calculateNetWorth(); // Chama a função do Patrimonio.gs
    
    const mensagem = `
🏛️ *Visão Geral do seu Patrimônio*

*Ativos:* ${formatCurrency(netWorthData.assets)}
(Seus bens e investimentos)

*Passivos:* ${formatCurrency(netWorthData.liabilities)}
(Suas dívidas e financiamentos)

------------------------------------
*Patrimônio Líquido:* *${formatCurrency(netWorthData.netWorth)}*
`;

    enviarMensagemTelegram(chatId, mensagem);
    logToSheet(`Patrimônio Líquido enviado para ${chatId}.`, "INFO");

  } catch (e) {
    handleError(e, "enviarPatrimonioLiquido", chatId);
  }
}

/**
 * NOVO: Lida com o comando /meuperfil para iniciar o Quiz Financeiro.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 */
function handleMeuPerfilCommand(chatId, usuario) {
  // Verifica se um quiz já não está em andamento para evitar confusão
  const existingState = getQuizState(chatId);
  if (existingState && existingState.currentQuestion < QUIZ_QUESTIONS.length) {
    enviarMensagemTelegram(chatId, "Parece que você já tem um quiz em andamento. Vamos continuar de onde parou.");
    sendQuizQuestion(chatId, existingState);
  } else {
    enviarMensagemTelegram(chatId, "Olá! Vamos descobrir qual é o seu perfil financeiro. Responda às próximas 5 perguntas.");
    startFinancialQuiz(chatId, usuario);
  }
}

// ===================================================================================
// ### INÍCIO DAS NOVAS FUNÇÕES DE GESTÃO DE INVESTIMENTOS ###
// ===================================================================================

/**
 * Lida com o comando para registrar a compra de um ativo.
 * É chamado tanto por linguagem natural (de FinancialLogic.gs) quanto por comando direto.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} ticker O código do ativo (ex: ITSA4).
 * @param {number} quantidade A quantidade de ativos comprados.
 * @param {number} preco O preço unitário de compra.
 * @param {string} nomeCorretora O nome da conta/corretora de onde o dinheiro saiu.
 * @param {string} usuario O nome do usuário que fez a compra.
 */
function handleComprarAtivo(chatId, ticker, quantidade, preco, nomeCorretora, usuario) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Validação dos dados de entrada
    if (!ticker || isNaN(quantidade) || quantidade <= 0 || isNaN(preco) || preco < 0 || !nomeCorretora) {
      enviarMensagemTelegram(chatId, "❌ Dados inválidos para a compra. Verifique o ticker, quantidade, preço e corretora.");
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const investimentosSheet = ss.getSheetByName(SHEET_INVESTIMENTOS);
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
    const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);

    // Verifica se a conta/corretora de origem existe
    const corretoraInfo = obterInformacoesDaConta(nomeCorretora, dadosContas);
    if (!corretoraInfo) {
      enviarMensagemTelegram(chatId, `❌ Conta/Corretora "${escapeMarkdown(nomeCorretora)}" não encontrada.`);
      return;
    }

    const valorTotalCompra = quantidade * preco;

    // Procura pelo ativo na aba 'Investimentos'
    const dadosInvestimentos = investimentosSheet.getDataRange().getValues();
    let ativoRowIndex = -1;
    for (let i = 1; i < dadosInvestimentos.length; i++) {
      if (dadosInvestimentos[i][0].toUpperCase() === ticker.toUpperCase()) {
        ativoRowIndex = i + 1;
        break;
      }
    }

    if (ativoRowIndex !== -1) {
      // Ativo já existe, atualiza a posição (preço médio)
      const qtdAtual = parseFloat(dadosInvestimentos[ativoRowIndex - 1][2]) || 0;
      const valorInvestidoAtual = parseBrazilianFloat(String(dadosInvestimentos[ativoRowIndex - 1][4]));
      
      const novaQtd = qtdAtual + quantidade;
      const novoValorInvestido = valorInvestidoAtual + valorTotalCompra;
      const novoPrecoMedio = novoValorInvestido / novaQtd;

      investimentosSheet.getRange(ativoRowIndex, 3).setValue(novaQtd); // Coluna C: Quantidade
      investimentosSheet.getRange(ativoRowIndex, 4).setValue(novoPrecoMedio); // Coluna D: Preço Médio
    } else {
      // Novo ativo, adiciona uma nova linha com as fórmulas
      const proximaLinha = investimentosSheet.getLastRow() + 1;
      investimentosSheet.appendRow([
        ticker.toUpperCase(),
        "Ação/FII", // Tipo genérico
        quantidade,
        preco,
        `=C${proximaLinha}*D${proximaLinha}`, // Fórmula do Valor Investido
        `=GOOGLEFINANCE("${ticker}")`, // Fórmula do Preço Atual
        `=C${proximaLinha}*F${proximaLinha}`, // Fórmula do Valor Atual
        `=G${proximaLinha}-E${proximaLinha}`, // Fórmula do Lucro/Prejuízo
        "Aberta"
      ]);
    }

    // Registra a transação de DESPESA correspondente
    transacoesSheet.appendRow([
      new Date(),
      `Compra de ${quantidade} ${ticker.toUpperCase()}`,
      "📈 Investimentos / Futuro",
      "Compra de Ativos",
      "Despesa",
      valorTotalCompra,
      "Transferência", // Método de pagamento
      corretoraInfo.nomeOriginal,
      1, 1, new Date(),
      usuario,
      "Ativo",
      Utilities.getUuid(),
      new Date()
    ]);
    
    // Ajuste incremental do saldo da conta de origem
    const contasSheet = ss.getSheetByName(SHEET_CONTAS);
    ajustarSaldoIncrementalmente(contasSheet, corretoraInfo.nomeOriginal, -valorTotalCompra);

    // Envia a mensagem de sucesso
    enviarMensagemTelegram(chatId, `✅ Compra de *${quantidade} ${escapeMarkdown(ticker.toUpperCase())}* registada com sucesso! O valor de ${formatCurrency(valorTotalCompra)} foi debitado de ${corretoraInfo.nomeOriginal}.`);
  
  } catch (e) {
    handleError(e, "handleComprarAtivo", chatId);
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

/**
 * Lida com o comando para registrar a venda de um ativo.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} ticker O código do ativo (ex: ITSA4).
 * @param {number} quantidade A quantidade de ativos vendidos.
 * @param {number} preco O preço unitário de venda.
 * @param {string} nomeContaDestino O nome da conta/corretora para onde o dinheiro foi.
 * @param {string} usuario O nome do usuário que fez a venda.
 */
function handleVenderAtivo(chatId, ticker, quantidade, preco, nomeContaDestino, usuario) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Validação dos dados
    if (!ticker || isNaN(quantidade) || quantidade <= 0 || isNaN(preco) || preco <= 0 || !nomeContaDestino) {
      enviarMensagemTelegram(chatId, "❌ Dados inválidos. Verifique o ticker, quantidade, preço e conta de destino.");
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const investimentosSheet = ss.getSheetByName(SHEET_INVESTIMENTOS);
    const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
    const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);

    // Verifica se a conta de destino existe
    const contaDestinoInfo = obterInformacoesDaConta(nomeContaDestino, dadosContas);
    if (!contaDestinoInfo) {
      enviarMensagemTelegram(chatId, `❌ Conta de destino "${escapeMarkdown(nomeContaDestino)}" não encontrada.`);
      return;
    }

    const valorTotalVenda = quantidade * preco;

    // Encontra o ativo na planilha
    const dadosInvestimentos = investimentosSheet.getDataRange().getValues();
    let ativoRowIndex = -1;
    for (let i = 1; i < dadosInvestimentos.length; i++) {
      if (dadosInvestimentos[i][0].toUpperCase() === ticker.toUpperCase()) {
        ativoRowIndex = i + 1;
        break;
      }
    }

    if (ativoRowIndex === -1) {
      enviarMensagemTelegram(chatId, `❌ Ativo *${escapeMarkdown(ticker.toUpperCase())}* não encontrado na sua carteira.`);
      return;
    }

    const qtdAtual = parseFloat(dadosInvestimentos[ativoRowIndex - 1][2]) || 0;
    if (quantidade > qtdAtual) {
      enviarMensagemTelegram(chatId, `❌ Você não pode vender ${quantidade} de ${escapeMarkdown(ticker.toUpperCase())}. Você possui apenas ${qtdAtual}.`);
      return;
    }

    // Atualiza a quantidade ou o status na aba 'Investimentos'
    const novaQtd = qtdAtual - quantidade;
    investimentosSheet.getRange(ativoRowIndex, 3).setValue(novaQtd); // Coluna C: Quantidade

    if (novaQtd === 0) {
      investimentosSheet.getRange(ativoRowIndex, 9).setValue("Fechada"); // Coluna I: Status
    }

    // Registra a transação de RECEITA correspondente
    transacoesSheet.appendRow([
      new Date(),
      `Venda de ${quantidade} ${ticker.toUpperCase()}`,
      "📈 Investimentos / Futuro",
      "Venda de Ativos",
      "Receita",
      valorTotalVenda,
      "Transferência", // Método de pagamento
      contaDestinoInfo.nomeOriginal,
      1, 1, new Date(),
      usuario,
      "Ativo",
      Utilities.getUuid(),
      new Date()
    ]);
    
    // OTIMIZAÇÃO: Ajusta o saldo incrementalmente
    const contasSheet = ss.getSheetByName(SHEET_CONTAS);
    ajustarSaldoIncrementalmente(contasSheet, contaDestinoInfo.nomeOriginal, valorTotalVenda);

    enviarMensagemTelegram(chatId, `✅ Venda de *${quantidade} ${escapeMarkdown(ticker.toUpperCase())}* registada com sucesso! O valor de ${formatCurrency(valorTotalVenda)} foi creditado em ${contaDestinoInfo.nomeOriginal}.`);

  } catch (e) {
    handleError(e, "handleVenderAtivo", chatId);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// ===================================================================================
// ### FIM DAS NOVAS FUNÇÕES DE GESTÃO DE INVESTIMENTOS ###
// ===================================================================================

/// ===================================================================================
// ##      NOVAS FUNÇÕES PARA DIVISÃO DE DESPESAS E EMPRÉSTIMOS (VERSÃO ATUALIZADA)   ##
// ===================================================================================

/**
 * Lida com o comando /dividir, registando uma despesa e criando uma conta a receber pela metade do valor.
 * @param {string} chatId - O ID do chat do Telegram.
 * @param {string} usuario - O nome do utilizador que está a dividir a despesa.
 * @param {number} valorTotal - O valor total da despesa.
 * @param {string} descricao - A descrição da despesa.
 * @param {string} pessoa - O nome da pessoa com quem a despesa foi dividida.
 * @param {string} restoDaFrase - O resto da frase que pode conter a conta de pagamento.
 */
function handleDividirDespesa(chatId, usuario, valorTotal, descricao, pessoa, restoDaFrase) {
  try {
    const valorAReceber = valorTotal / 2;
    const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);
    const dadosPalavras = getSheetDataWithCache(SHEET_PALAVRAS_CHAVE, CACHE_KEY_PALAVRAS);

    // ATUALIZADO: Lógica mais robusta para separar a pessoa da conta
    const separadoresConta = ['pelo', 'pela', 'no', 'na', 'usando', 'com'];
    let nomePessoa = '';
    let stringDaConta = '';
    let separadorEncontrado = false;

    const palavras = restoDaFrase.split(' ');
    for (let i = 0; i < palavras.length; i++) {
        if (separadoresConta.includes(palavras[i].toLowerCase()) && i > 0) {
            nomePessoa = palavras.slice(0, i).join(' ');
            stringDaConta = palavras.slice(i + 1).join(' ');
            separadorEncontrado = true;
            break;
        }
    }

    if (!separadorEncontrado) {
        nomePessoa = restoDaFrase; // Assume que a frase toda é a pessoa se não encontrar separador
        stringDaConta = restoDaFrase; // Tenta encontrar a conta na frase toda como fallback
    }
    
    // Extrai a conta de pagamento da string isolada
    const { conta, metodoPagamento } = extrairContaMetodoPagamento(stringDaConta, dadosContas, dadosPalavras);
    
    if (conta === "Não Identificada") {
      enviarMensagemTelegram(chatId, `❌ Não consegui identificar a conta de pagamento. Tente novamente, por exemplo: \`/dividir 100 do jantar com Gisa pelo Itau\``);
      return;
    }

    // 1. Registar a despesa completa
    const despesaId = Utilities.getUuid();
    registrarTransacaoNaPlanilha(new Date(), `(Dividido) ${capitalize(descricao)}`, 'Lazer', 'Saídas', 'Despesa', valorTotal, metodoPagamento, conta, 1, 1, new Date(), usuario, 'Ativo', despesaId, new Date());
    logToSheet(`[Dividir] Despesa total de ${valorTotal} registada para ${usuario} na conta ${conta}.`, "INFO");

    // 2. Criar a conta a receber
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);
    const hoje = new Date();
    const dataVencimento = new Date(hoje.getFullYear(), hoje.getMonth() + 1, hoje.getDate()); // Vencimento em 30 dias

    const newRow = [
      Utilities.getUuid(),
      `Valor a receber de ${capitalize(nomePessoa)} (${capitalize(descricao)})`, // Descrição melhorada
      'Contas a Receber',
      valorAReceber,
      Utilities.formatDate(dataVencimento, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      'Pendente',
      'Falso', // Não é recorrente
      conta,   // Sugestão de conta para receber
      `Referente à despesa dividida ID: ${despesaId}`,
      ''
    ];
    contasAPagarSheet.appendRow(newRow);
    logToSheet(`[Dividir] Conta a receber de ${valorAReceber} criada para ${nomePessoa}.`, "INFO");

    // 3. Enviar confirmação
    const mensagem = `✅ Despesa dividida com sucesso!\n\n` +
                     `- Registado um gasto de *${formatCurrency(valorTotal)}* na sua conta *${escapeMarkdown(conta)}*.\n` +
                     `- Criado um lembrete para receber *${formatCurrency(valorAReceber)}* de *${escapeMarkdown(capitalize(nomePessoa))}*.`;
    enviarMensagemTelegram(chatId, mensagem);
    
    atualizarSaldosDasContas();

  } catch (e) {
    handleError(e, "handleDividirDespesa", chatId);
  }
}

/**
 * Lida com o comando /emprestei, registando uma despesa e criando uma conta a receber pelo valor total.
 * @param {string} chatId - O ID do chat do Telegram.
 * @param {string} usuario - O nome do utilizador que está a emprestar.
 * @param {number} valor - O valor emprestado.
 * @param {string} pessoaEDesc - A string que contém o nome da pessoa e a descrição.
 * @param {string} conta - A conta de onde o dinheiro saiu.
 */
function handleEmprestarValor(chatId, usuario, valor, pessoaEDesc, conta) {
  try {
    if (conta === "Não Identificada") {
      enviarMensagemTelegram(chatId, `❌ Não consegui identificar a conta de onde o dinheiro saiu. Tente novamente, por exemplo: \`/emprestei 150 para a mamã com o Nubank\``);
      return;
    }
    
    // 1. Registar a despesa (saída do dinheiro)
    const despesaId = Utilities.getUuid();
    registrarTransacaoNaPlanilha(new Date(), `Empréstimo para ${capitalize(pessoaEDesc)}`, 'Outros', 'Empréstimos Concedidos', 'Despesa', valor, 'Transferência', conta, 1, 1, new Date(), usuario, 'Ativo', despesaId, new Date());
    logToSheet(`[Emprestei] Saída de ${valor} registada para ${usuario} da conta ${conta}.`, "INFO");
    
    // 2. Criar a conta a receber pelo valor total
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);
    const hoje = new Date();
    const dataVencimento = new Date(hoje.getFullYear(), hoje.getMonth() + 1, hoje.getDate());

    const newRow = [
      Utilities.getUuid(),
      `Valor a receber de ${capitalize(pessoaEDesc)}`,
      'Contas a Receber',
      valor,
      Utilities.formatDate(dataVencimento, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      'Pendente',
      'Falso',
      conta,
      `Referente ao empréstimo ID: ${despesaId}`,
      ''
    ];
    contasAPagarSheet.appendRow(newRow);
    logToSheet(`[Emprestei] Conta a receber de ${valor} criada para ${pessoaEDesc}.`, "INFO");

    // 3. Enviar confirmação
    const mensagem = `✅ Empréstimo registado com sucesso!\n\n` +
                     `- Registada uma saída de *${formatCurrency(valor)}* da sua conta *${escapeMarkdown(conta)}*.\n` +
                     `- Criado um lembrete para receber *${formatCurrency(valor)}* de *${escapeMarkdown(capitalize(pessoaEDesc))}*.`;
    enviarMensagemTelegram(chatId, mensagem);
    
    atualizarSaldosDasContas();

  } catch (e) {
    handleError(e, "handleEmprestarValor", chatId);
  }
}

/**
 * NOVO: Inicia o fluxo para importação de faturas/extratos.
 * @param {string} chatId O ID do chat do Telegram.
 */
function handleImportarCommand(chatId) {
  const mensagem = `🧾 *Importação de Fatura/Extrato*\n\n` +
                   `Para importar os seus lançamentos de uma só vez, utilize a ferramenta de **Conciliação Bancária** no seu dashboard web.\n\n` +
                   `Ela aceita ficheiros nos formatos \`.CSV\` e \`.OFX\`, que pode descarregar do site do seu banco.\n\n` +
                   `Vou gerar um link de acesso seguro para si. Clique nele para abrir o dashboard e começar.`;

  enviarMensagemTelegram(chatId, mensagem);
  enviarLinkDashboard(chatId); // Reutiliza a função que gera o link seguro
}

/**
 * NOVO: Lida com o comando para forçar o arquivamento de transações.
 * Apenas o administrador do sistema pode executar este comando.
 * @param {string} chatId O ID do chat do Telegram.
 */
function handleArquivarAgoraCommand(chatId) {
  logToSheet(`Comando /arquivar_agora recebido pelo chatId ${chatId}.`, "INFO");
  const adminChatId = getAdminChatIdFromProperties();

  if (chatId.toString() !== adminChatId.toString()) {
    enviarMensagemTelegram(chatId, "❌ Apenas o administrador do sistema pode executar este comando.");
    logToSheet(`Tentativa não autorizada de executar /arquivar_agora pelo chatId ${chatId}.`, "WARN");
    return;
  }

  try {
    enviarMensagemTelegram(chatId, "🔐 Ação de administrador confirmada. Iniciando o arquivamento manual de transações antigas... Isto pode demorar um momento.");
    // A função arquivarTransacoesAntigas já envia uma mensagem de conclusão.
    arquivarTransacoesAntigas();
  } catch (e) {
    handleError(e, "handleArquivarAgoraCommand", chatId);
  }
}
