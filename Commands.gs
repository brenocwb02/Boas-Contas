/**
 * @file Commands.gs
 * @description Este arquivo contém as implementações de todos os comandos do bot do Telegram.
 * Cada função aqui corresponde a um comando específico (/resumo, /saldo, etc.).
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


/**
 * ATUALIZADA: Envia o saldo atual das contas e faturas de cartão de crédito.
 * A lógica agora confia nos valores pré-calculados e consolidados por `atualizarSaldosDasContas`.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário solicitante.
 */
function enviarSaldo(chatId, usuario) {
  logToSheet(`Iniciando enviarSaldo para chatId: ${chatId}, usuario: ${usuario}`, "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configData = getSheetDataWithCache(SHEET_CONFIGURACOES, CACHE_KEY_CONFIG);
  const transacoesData = getSheetDataWithCache(SHEET_TRANSACOES, 'transacoes_cache');
  const dadosContas = getSheetDataWithCache(SHEET_CONTAS, CACHE_KEY_CONTAS);
  const grupoUsuarioChat = getGrupoPorChatId(chatId, configData);

  atualizarSaldosDasContas(); 
  logToSheet(`[enviarSaldo] Saldos atualizados. Iniciando construção da mensagem.`, "DEBUG");

  let mensagemSaldo = `💰 *Saldos Atuais - ${escapeMarkdown(grupoUsuarioChat || 'Pessoal')}*\n`;
  let totalContasCorrentes = 0;
  
  try {
    // --- 1. Exibir Saldos de Contas Correntes e Dinheiro ---
    mensagemSaldo += `──────────────\n*💵 Contas e Dinheiro*\n`;
    const contasCorrentes = [];
    for (const nomeNormalizado in globalThis.saldosCalculados) {
      const infoConta = globalThis.saldosCalculados[nomeNormalizado];
      if (infoConta.tipo === "conta corrente" || infoConta.tipo === "dinheiro físico") {
        totalContasCorrentes += infoConta.saldo;
        contasCorrentes.push(infoConta);
      }
    }
    contasCorrentes.sort((a, b) => a.nomeOriginal.localeCompare(b.nomeOriginal)).forEach(conta => {
        mensagemSaldo += `• ${escapeMarkdown(capitalize(conta.nomeOriginal))}: *${formatCurrency(conta.saldo)}*\n`;
    });

    // --- ESTRUTURA PARA AGRUPAR CARTÕES ---
    const agrupadores = {};

    // --- 2. Processar e Agrupar Faturas do Próximo Mês ---
    for (const nomeNormalizado in globalThis.saldosCalculados) {
      const infoConta = globalThis.saldosCalculados[nomeNormalizado];
      if (infoConta.tipo === "cartão de crédito" && infoConta.faturaAtual > 0) {
        const pai = infoConta.contaPaiAgrupador || infoConta.nomeNormalizado;
        if (!agrupadores[pai]) {
          agrupadores[pai] = { faturaProximoMes: 0, filhos: [] };
        }
        agrupadores[pai].faturaProximoMes += infoConta.faturaAtual;
        agrupadores[pai].filhos.push({ nome: infoConta.nomeOriginal, faturaProximoMes: infoConta.faturaAtual });
      }
    }

    mensagemSaldo += `\n*🗓️ Faturas (Próximo Vencimento)*\n`;
    let temFaturaProximoMes = false;
    Object.keys(agrupadores).forEach(pai => {
        const grupo = agrupadores[pai];
        if (grupo.faturaProximoMes > 0) {
            temFaturaProximoMes = true;
            mensagemSaldo += `• *${escapeMarkdown(capitalize(pai))}: ${formatCurrency(grupo.faturaProximoMes)}*\n`;
            if (grupo.filhos.length > 1) {
                grupo.filhos.forEach(filho => {
                    mensagemSaldo += `  - _${escapeMarkdown(filho.nome.split(' ').pop())}: ${formatCurrency(filho.faturaProximoMes)}_\n`;
                });
            }
        }
    });
    if (!temFaturaProximoMes) {
      mensagemSaldo += "_Nenhuma fatura para o próximo ciclo._\n";
    }

    // --- 3. Exibir Faturas a Vencer no Mês Atual (APENAS CARTÕES) ---
    const today = new Date();
    const currentMonth = today.getMonth();
    const currentYear = today.getFullYear();
    mensagemSaldo += `\n*🧾 Faturas a Vencer em ${getNomeMes(currentMonth)}*\n`;
    
    const faturasDoMesAtual = {};
    for (let i = 1; i < transacoesData.length; i++) {
        const linha = transacoesData[i];
        const dataVencimento = parseData(linha[10]);
        if (dataVencimento && dataVencimento.getMonth() === currentMonth && dataVencimento.getFullYear() === currentYear) {
            const conta = linha[7];
            const infoConta = obterInformacoesDaConta(conta, dadosContas);
            if (infoConta && normalizarTexto(infoConta.tipo) === "cartao de credito") {
                const pai = infoConta.contaPaiAgrupador || infoConta.nomeNormalizado;
                const valor = parseBrazilianFloat(String(linha[5]));
                if (!faturasDoMesAtual[pai]) {
                    faturasDoMesAtual[pai] = { total: 0, filhos: {} };
                }
                faturasDoMesAtual[pai].total += valor;
                faturasDoMesAtual[pai].filhos[infoConta.nomeOriginal] = (faturasDoMesAtual[pai].filhos[infoConta.nomeOriginal] || 0) + valor;
            }
        }
    }

    let temFaturaEsteMes = false;
    Object.keys(faturasDoMesAtual).forEach(pai => {
        const fatura = faturasDoMesAtual[pai];
        if (fatura.total > 0) {
            temFaturaEsteMes = true;
            mensagemSaldo += `• *${escapeMarkdown(capitalize(pai))}: ${formatCurrency(fatura.total)}*\n`;
            if (Object.keys(fatura.filhos).length > 1) {
                Object.keys(fatura.filhos).forEach(filho => {
                    mensagemSaldo += `  - _${escapeMarkdown(filho.split(' ').pop())}: ${formatCurrency(fatura.filhos[filho])}_\n`;
                });
            }
        }
    });

    if (!temFaturaEsteMes) {
        mensagemSaldo += "_Nenhuma fatura de cartão a vencer neste mês._\n";
    }

    // --- 4. Exibir Saldo Devedor Total (Consolidado) ---
    mensagemSaldo += `\n*💳 Dívida Total dos Cartões*\n`;
    let totalFaturasPagar = 0;
    const devedores = {};
    for (const nomeNormalizado in globalThis.saldosCalculados) {
      const infoConta = globalThis.saldosCalculados[nomeNormalizado];
      if (infoConta.tipo === "cartão de crédito" && infoConta.saldoTotalPendente > 0.01) {
        const pai = infoConta.contaPaiAgrupador || infoConta.nomeNormalizado;
        if (!devedores[pai]) {
          devedores[pai] = { total: 0, filhos: [] };
        }
        devedores[pai].total += infoConta.saldoTotalPendente;
        devedores[pai].filhos.push({ nome: infoConta.nomeOriginal, valor: infoConta.saldoTotalPendente });
        totalFaturasPagar += infoConta.saldoTotalPendente;
      }
    }
    
    let temSaldoDevedor = false;
    Object.keys(devedores).forEach(pai => {
        temSaldoDevedor = true;
        const grupo = devedores[pai];
        mensagemSaldo += `• *${escapeMarkdown(capitalize(pai))}: ${formatCurrency(grupo.total)}*\n`;
        if (grupo.filhos.length > 1) {
            grupo.filhos.forEach(filho => {
                mensagemSaldo += `  - _${escapeMarkdown(filho.nome.split(' ').pop())}: ${formatCurrency(filho.valor)}_\n`;
            });
        }
    });

    if (!temSaldoDevedor) {
      mensagemSaldo += "_Nenhum saldo devedor encontrado._\n";
    }

    // --- 5. Totais Gerais ---
    mensagemSaldo += `──────────────\n*🏦 Resumo Geral*\n`;
    mensagemSaldo += `*Total Disponível:* ${formatCurrency(totalContasCorrentes)}\n`;
    mensagemSaldo += `*Saldo Líquido (Disponível - Dívida):* ${formatCurrency(totalContasCorrentes - totalFaturasPagar)}\n`;

    enviarMensagemTelegram(chatId, mensagemSaldo);
    logToSheet(`Saldo enviado para chatId: ${chatId}.`, "INFO");

  } catch (e) {
    logToSheet(`ERRO FATAL ao construir ou enviar mensagem de saldo: ${e.message} na linha ${e.lineNumber}. Stack: ${e.stack}`, "ERROR");
    enviarMensagemTelegram(chatId, "❌ Houve um erro inesperado ao gerar seu saldo. Por favor, tente novamente mais tarde.");
  }
}


/**
 * Envia uma mensagem de ajuda com exemplos de comandos para o chat do Telegram.
 * @param {string} chatId O ID do chat do Telegram.
 */
function enviarAjuda(chatId) {
  const mensagem = `
📌*Como usar o Bot:*

Para registrar transacoes, envie uma mensagem no formato livre, incluindo valor, descricao, metodo de pagamento e conta/cartao. Quanto mais detalhes, melhor!

*💸 Para Gastos (Despesas):*
Use palavras como _gastei_, _paguei_, _comprei_.
• Ex: \`gastei 50 no mercado com Cartao Nubank Breno\`
• Ex: \`paguei 50 de uber no debito do Santander\`
• Ex: \`comprei 30 de pão com PIX do Itau\`
• Ex: \`paguei 2200 da fatura do Cartao Nubank com Itau\` (Para pagar fatura!)

*💰 Para Receitas:*
Use palavras como _recebi_, _ganhei_.
• Ex: \`recebi 3000 de salario no Itau via PIX\`
• Ex: \`recebi 500 de freelance no Nubank por transferencia\`
• Ex: \`recebi 200 de presente na Carteira (dinheiro fisico)\`

*? Dica para Receitas/Despesas:* Para que o bot entenda melhor, inclua sempre a *conta/cartao* e, se possivel, o *metodo de pagamento* (ex: PIX, debito, transferencia) na sua frase. Garanta que as palavras-chave para suas contas e metodos estao configuradas na aba \`PalavrasChave\`.

*🔄 Para Transferencias entre Contas:*
Use _transferi_ ou _enviei_.
• Ex: \`transferi 200 do Itau para o Mercado Pago\`

*🔢 Para Parcelamentos:*
Apenas adicione "em X vezes" ou "X vezes" ao final da frase.
• Ex: \`gastei 600 em roupas no Cartao Nubank Breno em 3x\`

*📊 Comandos de Consulta:*
• \`/resumo\` – Resumo financeiro do mes atual (ou use \`/resumo <mes> <ano>\` para meses anteriores, ex: \`/resumo junho 2024\`)
• \`/saldo\` – Saldo de todas as contas e faturas (sempre o saldo atual)
• \`/extrato\` – Ver suas ultimas transacoes (com filtro opcional, ex: \`/extrato despesas julho 2024\`)
  • Ex: \`/extrato receitas\` (ver so receitas)
  • Ex: \`/extrato despesas\` (ver so despesas)
  • Ex: \`/extrato Gisele\` (ver extrato de uma pessoa)
  • Ex: \`/extrato tudo\` (ver todas as transacoes do mes atual)
  • Ex: \`/extrato julho 2024\` (ver todas as transacoes de julho de 2024)
• \`/proximasfaturas\` – Veja o total de gastos ja lancados para faturas futuras
• \`/contasapagar\` – Verifique o status das suas contas fixas do mes (ou use \`/contasapagar <mes> <ano>\`)
• \`/vincular_conta <ID_CONTA_A_PAGAR> <ID_TRANSACAO>\` – Vincula manualmente uma transacao a uma conta a pagar fixa.
• \`/metas\` – Acompanhe suas metas financeiras (do mes atual, ou use \`/metas <mes> <ano>\`)
• \`/ajuda\` – Exibe esta lista de comandos

*? Dica:* Use sempre os *nomes exatos* das suas Contas e Cartões (ex: "Cartao Nubank Breno", "Itau", "Mercado Pago"). Se o bot não entender sua mensagem, tente reformular de forma mais simples e direta.
  `;

  const teclado = {
    inline_keyboard: [
      [
        { text: "📊 Resumo", callback_data: "/resumo" },
        { text: "💰 Saldo", callback_data: "/saldo" }
      ],
      [
        { text: "📄 Extrato", callback_data: "/extrato" },
        { text: "🎯 Metas", callback_data: "/metas" }
      ],
      [
        { text: "🧾 Proximas Faturas", callback_data: "/proximasfaturas" },
        { text: "🗓️ Contas a Pagar", callback_data: "/contasapagar" }
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
 * Envia o progresso das metas financeiras para o chat do Telegram.
 * Soma os gastos por categoria e subcategoria e compara com as metas definidas na planilha.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário que solicitou as metas.
 * @param {number} mes O mês para as metas (1-12).
 * @param {number} ano O ano para as metas.
 */
function enviarMetas(chatId, usuario, mes, ano) {
  logToSheet(`[Metas] Iniciando enviarMetas para usuario: ${usuario}, Mes: ${mes}, Ano: ${ano}`, "INFO");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoes = ss.getSheetByName(SHEET_TRANSACOES).getDataRange().getValues();
  const metasSheet = ss.getSheetByName(SHEET_METAS).getDataRange().getValues();

  const targetMesIndex = mes - 1;
  const targetAno = ano;
  const nomeMes = getNomeMes(targetMesIndex);

  const cabecalho = metasSheet[2];
  let colMetaMes = -1;

  for (let i = 2; i < cabecalho.length; i++) {
    if (String(cabecalho[i]).toLowerCase().includes(nomeMes.toLowerCase())) {
      colMetaMes = i;
      break;
    }
  }

  if (colMetaMes === -1) {
    logToSheet(`[Metas] Coluna do mes para ${nomeMes}/${targetAno} não encontrada na aba 'Metas'.`, "ERROR");
    enviarMensagemTelegram(chatId, "❌ Não foi possivel carregar as metas para este mes. Verifique a aba 'Metas' para o mes de " + nomeMes + "/" + targetAno + ".");
    return;
  }
  logToSheet(`[Metas] Coluna de meta encontrada para ${nomeMes}/${targetAno} no indice: ${colMetaMes}`, "DEBUG");

  let metas = {};
  let totaisPorCategoria = {};

  logToSheet("[Metas] --- Inicio do Processamento de Metas (enviarMetas) ---", "DEBUG");
  for (let i = 3; i < metasSheet.length; i++) {
    const categoriaOriginal = (metasSheet[i][0] || "").toString().trim();
    const subcategoriaOriginal = (metasSheet[i][1] || "").toString().trim();
    const valorMetaTexto = metasSheet[i][colMetaMes];

    if (!categoriaOriginal || !subcategoriaOriginal || !valorMetaTexto) {
        logToSheet(`[Metas] Linha ${i + 1} da aba Metas ignorada (categoria, subcategoria ou valor vazios).`, "DEBUG");
        continue;
    }

    const chaveCategoria = normalizarTexto(categoriaOriginal);
    const chaveSubcategoria = normalizarTexto(`${categoriaOriginal} ${subcategoriaOriginal}`);

    let meta = parseBrazilianFloat(String(valorMetaTexto));

    if (isNaN(meta) || meta <= 0) {
        logToSheet(`[Metas] Meta invalida para "${categoriaOriginal} > ${subcategoriaOriginal}" (R$ ${valorMetaTexto}). Pulando.`, "DEBUG");
        continue;
    }

    metas[chaveSubcategoria] = {
      categoriaOriginal: categoriaOriginal,
      subcategoriaOriginal: subcategoriaOriginal,
      categoriaNormalizada: chaveCategoria,
      meta: meta,
      gasto: 0
    };

    if (!totaisPorCategoria[chaveCategoria]) {
      totaisPorCategoria[chaveCategoria] = { meta: 0, gasto: 0, subcategories: [], nomeOriginal: categoriaOriginal };
      logToSheet(`[Metas] Criando entrada para categoria total: ${chaveCategoria}`, "DEBUG");
    }

    totaisPorCategoria[chaveCategoria].meta += meta;
    totaisPorCategoria[chaveCategoria].subcategories.push(chaveSubcategoria);
    logToSheet(`[Metas] Meta Processada: Original="${categoriaOriginal} > ${subcategoriaOriginal}", Normalizada (chaveSubcategoria)="${chaveSubcategoria}", Meta: R$ ${meta.toFixed(2)}`, "DEBUG");
  }
  logToSheet("[Metas] --- Fim do Processamento de Metas (enviarMetas) ---", "DEBUG");
  logToSheet(`[Metas] Metas carregadas: ${JSON.stringify(metas)}`, "DEBUG");

  logToSheet("[Metas] --- Inicio do Processamento de Transacoes (enviarMetas) ---", "DEBUG");
  for (let i = 1; i < transacoes.length; i++) {
    const dataVencimento = parseData(transacoes[i][10]); // Use Data de Vencimento
    const tipo = transacoes[i][4];
    const categoriaTransacao = (transacoes[i][2] || "").toString().trim(); 
    const subcategoriaTransacao = (transacoes[i][3] || "").toString().trim(); 
    const rawValor = transacoes[i][5];
    const usuarioLinha = (transacoes[i][11] || "").toString().trim();

    logToSheet(`[Metas] Transacao ${i + 1} (ID: ${transacoes[i][13] || 'N/A'}): Data Vencimento: ${dataVencimento ? dataVencimento.toLocaleDateString() : 'Invalida'}, Tipo: ${tipo}, Categoria: ${categoriaTransacao}, Subcategoria: ${subcategoriaTransacao}, Valor: ${rawValor}, Usuário da Linha: "${usuarioLinha}"`, "DEBUG");

    if (
      !dataVencimento || dataVencimento.getMonth() !== targetMesIndex || dataVencimento.getFullYear() !== targetAno || // Filter by DUE DATE
      tipo !== "Despesa"
    ) {
        logToSheet(`[Metas] Transacao ${i + 1} ignorada: Data de Vencimento (${dataVencimento ? dataVencimento.toLocaleDateString() : 'N/A'}) fora do mes/ano alvo ou nao e despesa.`, "DEBUG");
        continue;
    }

    const chaveTransacaoNormalizada = normalizarTexto(`${categoriaTransacao} ${subcategoriaTransacao}`);
    logToSheet(`[Metas] Transacao ${i + 1} - Chave normalizada: "${chaveTransacaoNormalizada}"`, "DEBUG");

    if (metas[chaveTransacaoNormalizada]) {
      const metaEntry = metas[chaveTransacaoNormalizada];
      const targetCategoryNormalizada = metaEntry.categoriaNormalizada;

      let valor = parseBrazilianFloat(String(rawValor));

      if (!isNaN(valor)) {
        metaEntry.gasto += valor;
        logToSheet(`[Metas] Gasto de R$ ${valor.toFixed(2)} adicionado para meta "${chaveTransacaoNormalizada}". Gasto atual na meta: R$ ${metaEntry.gasto.toFixed(2)}`, "DEBUG");

        if (targetCategoryNormalizada && totaisPorCategoria[targetCategoryNormalizada]) {
          totaisPorCategoria[targetCategoryNormalizada].gasto += valor;
          logToSheet(`[Metas] Gasto de R$ ${valor.toFixed(2)} adicionado para total da categoria "${targetCategoryNormalizada}". Gasto atual total: R$ ${totaisPorCategoria[targetCategoryNormalizada].gasto.toFixed(2)}`, "DEBUG");
        } else {
          logToSheet(`[Metas] ERRO: Categoria normalizada "${targetCategoryNormalizada}" não encontrada em 'totaisPorCategoria' para meta "${chaveTransacaoNormalizada}".`, "ERROR");
        }
      } else {
          logToSheet(`[Metas] Valor invalido na transacao ${i + 1} para meta "${chaveTransacaoNormalizada}": ${rawValor}`, "DEBUG");
      }
    } else {
        logToSheet(`[Metas] Transacao ${i + 1} ("${chaveTransacaoNormalizada}") não encontrou meta correspondente.`, "DEBUG");
    }
  }
  logToSheet("[Metas] --- Fim do Processamento de Transacoes (enviarMetas) ---", "DEBUG");

  logToSheet(`[Metas] Estado final de 'metas': ${JSON.stringify(metas)}`, "DEBUG");
  logToSheet(`[Metas] Estado final de 'totaisPorCategoria': ${JSON.stringify(totaisPorCategoria)}`, "DEBUG");


  let mensagem = `🎯 *Metas de ${nomeMes}/${targetAno} (Visão Familiar)*\n`;
  let totalGeral = 0;
  let temMetasParaExibir = false;

  const categoriasOrdenadas = Object.keys(totaisPorCategoria).sort((a, b) => {
    const nomeOriginalA = totaisPorCategoria[a].nomeOriginal;
    const nomeOriginalB = totaisPorCategoria[b].nomeOriginal;
    return nomeOriginalA.localeCompare(nomeOriginalB);
  });

  for (const categoriaNormalizada of categoriasOrdenadas) { 
    const bloco = totaisPorCategoria[categoriaNormalizada];
    const percCat = bloco.meta > 0 ? (bloco.gasto / bloco.meta) * 100 : 0;

    const linhasSub = [];
    const subcategoriasOrdenadas = bloco.subcategories.sort((a, b) => {
      const itemA = metas[a];
      const itemB = metas[b];
      return itemA.subcategoriaOriginal.localeCompare(itemB.subcategoriaOriginal);
    });

    for (const chaveSubcategoria of subcategoriasOrdenadas) { // CORREÇÃO: Iterar sobre 'subcategoriasOrdenadas'
      const item = metas[chaveSubcategoria];
      if (item.gasto > 0 || item.meta > 0) {
        temMetasParaExibir = true; 
        const perc = item.meta > 0 ? (item.gasto / item.meta) * 100 : 0;
        let emoji = "";
        if (perc >= 100 && item.meta > 0) emoji = "⛔";
        else if (perc >= 80 && item.meta > 0) emoji = "⚠️";
        else if (item.meta > 0) emoji = "✅";
        else emoji = "ℹ️";

        const nome = escapeMarkdown(item.subcategoriaOriginal).padEnd(20, ".");
        const linha = `• ${nome} R$ ${item.gasto.toFixed(2).padStart(7).replace('.', ',')} / R$ ${item.meta.toFixed(2).padEnd(7).replace('.', ',')} ${emoji} ${perc.toFixed(0)}%`;
        linhasSub.push(linha);
      }
    }

    if (linhasSub.length > 0) {
      mensagem += `\n──────────────\n*${escapeMarkdown(capitalize(bloco.nomeOriginal))}* — ${percCat.toFixed(0)}% da meta (R$ ${bloco.gasto.toFixed(2).replace('.', ',')} / R$ ${bloco.meta.toFixed(2).replace('.', ',')})\n`;
      mensagem += linhasSub.join("\n");
      totalGeral += bloco.gasto;
    }
  }

  logToSheet(`[Metas] Valor final de 'temMetasParaExibir': ${temMetasParaExibir}`, "DEBUG");

  if (!temMetasParaExibir) {
    mensagem = `🎯 Nenhuma meta configurada ou atingida para ${nomeMes}/${targetAno} (Visão Familiar).`;
  } else {
     mensagem += `\n\n💵 *Total Gasto Geral:* R$ ${totalGeral.toFixed(2).replace('.', ',')}`;
  }

  enviarMensagemTelegram(chatId, mensagem);
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
 * Se o lançamento excluído estiver vinculado a uma conta na aba "Contas_a_Pagar",
 * o status dessa conta será revertido para "Pendente" e o vínculo será removido.
 * @param {string} idLancamento O ID único do lançamento a ser excluído.
 * @param {string} chatId O ID do chat do Telegram para enviar feedback.
 */
function excluirLancamentoPorId(idLancamento, chatId) {
  logToSheet(`Iniciando exclusao de lancamento para ID: ${idLancamento}`, "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
  const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);

  if (!transacoesSheet) {
    enviarMensagemTelegram(chatId, "❌ Erro: Aba 'Transacoes' não encontrada.");
    logToSheet("Erro: Aba 'Transacoes' não encontrada para exclusao.", "ERROR");
    return;
  }

  const dadosTransacoes = transacoesSheet.getDataRange().getValues();
  const headersTransacoes = transacoesSheet.getRange(1, 1, 1, transacoesSheet.getLastColumn()).getValues()[0];
  const colIdTransacao = headersTransacoes.indexOf('ID Transacao');

  if (colIdTransacao === -1) {
    enviarMensagemTelegram(chatId, "❌ Erro: Coluna 'ID Transacao' não encontrada na aba 'Transacoes'.");
    logToSheet("Erro: Coluna 'ID Transacao' ausente na aba 'Transacoes' para exclusao.", "ERROR");
    return;
  }

  let linhaParaExcluir = -1;
  let descricaoLancamento = "";

  for (let i = 1; i < dadosTransacoes.length; i++) {
    if (dadosTransacoes[i][colIdTransacao] === idLancamento) {
      linhaParaExcluir = i + 1;
      descricaoLancamento = dadosTransacoes[i][1];
      break;
    }
  }

  if (linhaParaExcluir !== -1) {
    transacoesSheet.deleteRow(linhaParaExcluir);
    logToSheet(`Lancamento '${descricaoLancamento}' (ID: ${idLancamento}) excluído da aba 'Transacoes'.`, "INFO");
    
    if (contasAPagarSheet) {
      const dadosContasAPagar = contasAPagarSheet.getDataRange().getValues();
      const headersContasAPagar = contasAPagarSheet.getRange(1, 1, 1, contasAPagarSheet.getLastColumn()).getValues()[0];
      const colStatusContasAPagar = headersContasAPagar.indexOf('Status');
      const colIDTransacaoVinculada = headersContasAPagar.indexOf('ID Transacao Vinculada');

      if (colStatusContasAPagar !== -1 && colIDTransacaoVinculada !== -1) {
        let contaAPagarAtualizada = false;
        for (let i = 1; i < dadosContasAPagar.length; i++) {
          if (dadosContasAPagar[i][colIDTransacaoVinculada] === idLancamento) {
            const linhaContaAPagar = i + 1;
            const descricaoContaAPagar = dadosContasAPagar[i][1];
            
            contasAPagarSheet.getRange(linhaContaAPagar, colStatusContasAPagar + 1).setValue("Pendente");
            contasAPagarSheet.getRange(linhaContaAPagar, colIDTransacaoVinculada + 1).setValue("");
            logToSheet(`Conta a Pagar '${descricaoContaAPagar}' (ID: ${dadosContasAPagar[i][0]}) revertida para 'Pendente' apos exclusao de transacao vinculada.`, "INFO");
            contaAPagarAtualizada = true;
            break;
          }
        }
        if (!contaAPagarAtualizada) {
          logToSheet(`Nenhuma conta a pagar vinculada ao ID de transacao '${idLancamento}' foi encontrada para reverter status.`, "DEBUG");
        }
      } else {
        logToSheet("Colunas 'Status' ou 'ID Transacao Vinculada' ausentes na aba 'Contas_a_Pagar'. Nao foi possivel reverter status.", "WARN");
      }
    } else {
      logToSheet("Aba 'Contas_a_Pagar' nao encontrada. Nao foi possivel reverter status de contas vinculadas.", "WARN");
    }

    atualizarSaldosDasContas();

    enviarMensagemTelegram(chatId, `✅ Lançamento '${escapeMarkdown(descricaoLancamento)}' (ID: ${escapeMarkdown(idLancamento)}) excluído com sucesso! Saldo atualizado.`);
  } else {
    enviarMensagemTelegram(chatId, `❌ Lançamento com ID *${escapeMarkdown(idLancamento)}* não encontrado.`);
    logToSheet(`Erro: Lancamento ID ${idLancamento} nao encontrado para exclusao.`, "WARN");
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
 * **FUNÇÃO CORRIGIDA**
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
      break;
    }
  }

  if (consultaNormalizada.includes("mes passado")) {
    dataInicio = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);
    dataFim = new Date(hoje.getFullYear(), hoje.getMonth(), 0, 23, 59, 59);
    periodoTexto = "no mês passado";
  } else if (consultaNormalizada.includes("hoje")) {
    dataInicio = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
    dataFim = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate(), 23, 59, 59);
    periodoTexto = "hoje";
  } else if (consultaNormalizada.includes("ontem")) {
    const ontem = new Date(hoje);
    ontem.setDate(hoje.getDate() - 1);
    dataInicio = new Date(ontem.getFullYear(), ontem.getMonth(), ontem.getDate());
    dataFim = new Date(ontem.getFullYear(), ontem.getMonth(), ontem.getDate(), 23, 59, 59);
    periodoTexto = "ontem";
  }

  logToSheet(`[ConsultaLN] Período de tempo determinado: ${dataInicio.toLocaleDateString()} a ${dataFim.toLocaleDateString()} (${periodoTexto})`, "DEBUG");

  // --- 2. Determinar o Tipo de Consulta e Filtros ---
  const tipoConsulta = consultaNormalizada.includes("listar") || consultaNormalizada.includes("quais") ? "LISTAR" : "SOMAR";
  let tipoTransacaoFiltro = null;
  if (consultaNormalizada.includes("despesa")) tipoTransacaoFiltro = "Despesa";
  if (consultaNormalizada.includes("receita")) tipoTransacaoFiltro = "Receita";
  
  const regexFiltro = /(?:com|de|sobre)\s+(.+?)(?=\s+em\s+[a-z]+|\s+este\s+mes|\s+mes\s+passado|\s+hoje|\s+ontem|$)/;
  const matchFiltro = consultaNormalizada.match(regexFiltro);
  
  let filtroTexto = "";
  if (matchFiltro) {
    filtroTexto = matchFiltro[1].trim();
  } else {
    let tempConsulta = ' ' + consultaNormalizada + ' ';
    const palavrasParaRemover = [
      "quanto gastei", "listar despesas", "total de", "quanto recebi", "listar receitas",
      "este mes", "mes passado", "hoje", "ontem", "do mes", "no mes",
      "janeiro", "fevereiro", "marco", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
      "quanto", "qual", "quais", "listar", "mostrar", "total", "despesas", "receitas", "despesa", "receita",
      "meu", "minha", "meus", "minhas"
    ];
    palavrasParaRemover.sort((a,b) => b.length - a.length).forEach(palavra => {
        tempConsulta = tempConsulta.replace(new RegExp(`\\s${palavra}\\s`, 'gi'), ' ');
    });
    filtroTexto = tempConsulta.trim();
  }

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

    // Filtro por período
    if (!dataTransacao || dataTransacao < dataInicio || dataTransacao > dataFim) {
      continue;
    }

    // Filtro por tipo de transação (se especificado)
    if (tipoTransacaoFiltro && normalizarTexto(tipo) === normalizarTexto(tipoTransacaoFiltro)) {
      continue;
    }

    // Filtro por texto na descrição, categoria, subcategoria ou conta
    const relevanteParaFiltro = (
      normalizarTexto(descricao).includes(normalizarTexto(filtroTexto)) ||
      normalizarTexto(categoria).includes(normalizarTexto(filtroTexto)) ||
      normalizarTexto(subcategoria).includes(normalizarTexto(filtroTexto)) ||
      normalizarTexto(conta).includes(normalizarTexto(filtroTexto))
    );

    if (filtroTexto && !relevanteParaFiltro) {
        continue;
    }
    // Exclui pagamentos de fatura e transferências para evitar dupla contagem em consultas de "gastos" totais
    if (tipo === "Despesa" && (normalizarTexto(categoria) === "contas a pagar" && normalizarTexto(subcategoria) === "pagamento de fatura" || normalizarTexto(categoria) === "transferencias")) {
        logToSheet(`[ConsultaLN] Transacao ID ${id} (${categoria} > ${subcategoria}) excluida da soma/listagem (pagamento de fatura/transferencia).`, "DEBUG");
        continue;
    }
    if (tipo === "Receita" && normalizarTexto(categoria) === "transferencias") {
        logToSheet(`[ConsultaLN] Transacao ID ${id} (${categoria} > ${subcategoria}) excluida da soma/listagem (transferencia).`, "DEBUG");
        continue;
    }

    if (tipoConsulta === "SOMAR") {
      totalSoma += valor;
    } else { // LISTAR
      transacoesEncontradas.push({
        data: Utilities.formatDate(dataTransacao, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        descricao: descricao,
        categoria: categoria,
        subcategoria: subcategoria,
        tipo: tipo,
        valor: valor,
        conta: conta,
        id: id // Inclui ID para possível exclusão
      });
    }
  }

  let mensagemResposta = "";
  if (tipoConsulta === "SOMAR") {
    let prefixoTipo = tipoTransacaoFiltro === "Receita" ? "Receita" : "Gasto";
    mensagemResposta = `O *total de ${prefixoTipo}* ${filtroTexto ? `com "${escapeMarkdown(filtroTexto)}"` : ""} ${periodoTexto} é de: *${formatCurrency(totalSoma)}*.`;
  } else { // LISTAR
    if (transacoesEncontradas.length > 0) {
      mensagemResposta = `*Lancamentos ${filtroTexto ? `com "${escapeMarkdown(filtroTexto)}"` : ""} ${periodoTexto}:*\n\n`;
      transacoesEncontradas.sort((a, b) => parseData(b.data).getTime() - parseData(a.data).getTime()); // Mais recente primeiro
      transacoesEncontradas.slice(0, 10).forEach(t => { // Limita a 10 para não sobrecarregar
        const valorFormatado = formatCurrency(t.valor);
        const tipoIcon = t.tipo === "Receita" ? "💰" : "💸";
        mensagemResposta += `${tipoIcon} ${escapeMarkdown(t.descricao)} (${escapeMarkdown(t.categoria)} > ${escapeMarkdown(t.subcategoria)}) - *${valorFormatado}*\n`;
      });
      if (transacoesEncontradas.length > 10) {
        mensagemResposta += `\n...e mais ${transacoesEncontradas.length - 10} lancamentos.`;
      }
    } else {
      mensagemResposta = `Nenhum lançamento ${filtroTexto ? `com "${escapeMarkdown(filtroTexto)}"` : ""} encontrado ${periodoTexto}.`;
    }
  }

  enviarMensagemTelegram(chatId, mensagemResposta);
  logToSheet(`[ConsultaLN] Resposta enviada para ${chatId}: "${mensagemResposta.substring(0, 100)}..."`, "INFO");
}

/**
 * NOVO: Inicia o processo de edição da última transação do usuário.
 * Armazena o estado de edição no cache.
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

  // Busca a última transação do usuário ou do grupo
  for (let i = dadosTransacoes.length - 1; i > 0; i--) {
    const linha = dadosTransacoes[i];
    const usuarioLinha = normalizarTexto(linha[11]);
    const grupoTransacao = getGrupoPorChatIdByUsuario(linha[11], dadosConfig);

    if (usuarioLinha === usuarioNormalizado || grupoTransacao === grupoUsuarioChat) {
      ultimaTransacao = {
        linha: i + 1, // Linha da planilha (base 1)
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
      logToSheet(`[Edicao] Ultima transacao encontrada: ID ${ultimaTransacao.id}, Descricao: "${ultimaTransacao.descricao}"`, "DEBUG");
      break;
    }
  }

  if (!ultimaTransacao) {
    enviarMensagemTelegram(chatId, "⚠️ Nenhuma transação recente encontrada para você ou seu grupo para editar.");
    logToSheet(`[Edicao] Nenhuma transacao encontrada para edicao para ${usuario}.`, "INFO");
    return;
  }

  // Armazena o estado da edição no cache
  setEditState(chatId, {
    transactionId: ultimaTransacao.id,
    rowIndex: ultimaTransacao.linha,
    originalData: ultimaTransacao // Armazena a transação completa original
  });

  const mensagem = `✏️ *Editando seu último lançamento* (ID: \`${escapeMarkdown(ultimaTransacao.id)}\`):\n\n` +
                   `*Data:* ${ultimaTransacao.data}\n` +
                   `*Descricao:* ${escapeMarkdown(ultimaTransacao.descricao)}\n` +
                   `*Valor:* ${formatCurrency(ultimaTransacao.valor)}\n` +
                   `*Tipo:* ${ultimaTransacao.tipo}\n` +
                   `*Conta:* ${escapeMarkdown(ultimaTransacao.conta)}\n` +
                   `*Categoria:* ${escapeMarkdown(ultimaTransacao.categoria)}\n` +
                   `*Subcategoria:* ${escapeMarkdown(ultimaTransacao.subcategoria)}\n` +
                   `*Metodo:* ${escapeMarkdown(ultimaTransacao.metodoPagamento)}\n` +
                   `*Vencimento:* ${ultimaTransacao.dataVencimento}\n\n` +
                   `Qual campo deseja editar?`;

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
      [{ text: "❌ Cancelar Edição", callback_data: `cancelar_edicao` }]
    ]
  };

  enviarMensagemTelegram(chatId, mensagem, { reply_markup: teclado });
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
 * NOVO: Processa a entrada do usuário para a edição de um campo específico.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 * @param {string} novoValor O novo valor enviado pelo usuário.
 * @param {Object} editState O estado atual da edição (contém transactionId e fieldToEdit).
 * @param {Array<Array<any>>} dadosContas Dados da aba 'Contas' para validação de conta/cartão.
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

  const headers = transacoesSheet.getDataRange().getValues()[0];
  const colMap = getColumnMap(headers);

  // Busca a linha da transação novamente (garante que não foi excluída etc.)
  const colIdTransacao = colMap["ID Transacao"];
  let rowIndex = -1;
  const allTransactionsData = transacoesSheet.getDataRange().getValues();
  for (let i = 1; i < allTransactionsData.length; i++) {
    if (allTransactionsData[i][colIdTransacao] === editState.transactionId) {
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
      const { categoria: detectedCategory } = extrairCategoriaSubcategoria(novoValor, allTransactionsData[rowIndex-1][colMap["Tipo"]], dadosPalavras); // Passa o tipo original da transação
      if (detectedCategory && detectedCategory !== "Não Identificada") {
          valorParaSet = detectedCategory;
          mensagemSucesso = "Categoria atualizada!";
          // Se a categoria mudar, a subcategoria pode precisar ser reavaliada
          // ou pode ser um bom momento para pedir a subcategoria novamente.
          // Por simplicidade, não vamos pedir a subcategoria aqui, mas é um ponto de melhoria.
      } else {
          mensagemSucesso = "❌ Categoria não reconhecida. Por favor, verifique as palavras-chave da categoria.";
          erroValidacao = true;
      }
      break;
    case "subcategoria":
      colIndex = colMap["Subcategoria"];
      const tipoTransacaoOriginal = allTransactionsData[rowIndex-1][colMap["Tipo"]]; // Obtém o tipo da transação original
      const { categoria: catOriginal, subcategoria: detectedSubcategory } = extrairCategoriaSubcategoria(novoValor, tipoTransacaoOriginal, dadosPalavras);
      if (detectedSubcategory && detectedSubcategory !== "Não Identificada") {
          // Também tenta atualizar a categoria se a subcategoria for mais específica
          const currentCategory = allTransactionsData[rowIndex-1][colMap["Categoria"]];
          if (catOriginal && normalizarTexto(catOriginal) !== normalizarTexto(currentCategory)) {
              // Se a nova subcategoria veio de uma categoria diferente, atualiza a categoria também
              transacoesSheet.getRange(rowIndex, colMap["Categoria"] + 1).setValue(catOriginal);
              logToSheet(`[Edicao] Categoria atualizada de '${currentCategory}' para '${catOriginal}' ao editar subcategoria.`, "DEBUG");
          }
          valorParaSet = detectedSubcategory;
          mensagemSucesso = "Subcategoria atualizada!";
      } else {
          mensagemSucesso = "❌ Subcategoria não reconhecida. Por favor, verifique as palavras-chave da subcategoria.";
          erroValidacao = true;
      }
      break;
    case "metodoPagamento":
      colIndex = colMap["Metodo de Pagamento"]; // CORREÇÃO: Nome da coluna ajustado
      const metodoNormalizado = normalizarTexto(novoValor);
      const metodosValidos = ["credito", "debito", "dinheiro", "pix", "boleto", "transferencia bancaria"]; // Adicionar mais se necessário
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
    // Não limpa o estado de edição para permitir que o usuário tente novamente
    // Ou pode adicionar um botão para "Cancelar Edição" aqui
    logToSheet(`[Edicao] Erro de validacao para campo '${editState.fieldToEdit}': ${mensagemSucesso}`, "WARN");
    return;
  }

  // CORREÇÃO: Mover a declaração de 'lock' para fora do try
  let lock; 
  try {
    lock = LockService.getScriptLock();
    lock.waitLock(30000);
    transacoesSheet.getRange(rowIndex, colIndex + 1).setValue(valorParaSet);
    logToSheet(`[Edicao] Transacao ID ${editState.transactionId} - Campo '${editState.fieldToEdit}' atualizado para: "${valorParaSet}".`, "INFO");
    enviarMensagemTelegram(chatId, `✅ ${mensagemSucesso}`);
    atualizarSaldosDasContas(); // Recalcula saldos após a atualização
    clearEditState(chatId); // Limpa o estado de edição após o sucesso
  } catch (e) {
    logToSheet(`ERRO ao atualizar transacao ID ${editState.transactionId}: ${e.message}`, "ERROR");
    enviarMensagemTelegram(chatId, `❌ Houve um erro ao atualizar o lançamento: ${e.message}`);
  } finally {
    if (lock) { // Verifica se lock foi definido antes de tentar liberar
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
