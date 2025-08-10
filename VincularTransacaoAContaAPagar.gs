/**
 * @file VincularTransacaoAContaAPagar.gs
 * @description Este arquivo contém a função para vincular manualmente uma transação a uma conta a pagar fixa.
 * Esta função foi referenciada no `doPost` mas não estava definida no código original.
 */

/**
 * Placeholder para a função vincularTransacaoAContaAPagar.
 * Esta função não foi definida no código original, mas é referenciada em doPost.
 * Você precisará implementá-la.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} idContaAPagar O ID da conta a pagar.
 * @param {string} idTransacao O ID da transação a ser vinculada.
 */
function vincularTransacaoAContaAPagar(chatId, idContaAPagar, idTransacao) {
  logToSheet(`[VincularConta] Chamada para vincular conta a pagar ${idContaAPagar} com transacao ${idTransacao}.`, "INFO");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const contasAPagarSheet = ss.getSheetByName(SHEET_CONTAS_A_PAGAR);

  if (!contasAPagarSheet) {
    enviarMensagemTelegram(chatId, "❌ Erro: Aba 'Contas_a_Pagar' não encontrada para vincular transação.");
    logToSheet("Erro: Aba 'Contas_a_Pagar' ausente para vincular transação.", "ERROR");
    return;
  }

  const contaAPagarInfo = obterInformacoesDaContaAPagar(idContaAPagar);

  if (!contaAPagarInfo) {
    enviarMensagemTelegram(chatId, `❌ Conta a Pagar com ID *${escapeMarkdown(idContaAPagar)}* não encontrada.`);
    logToSheet(`Erro: Conta a Pagar ID ${idContaAPagar} não encontrada para vincular.`, "WARN");
    return;
  }

  // Verifica se a conta a pagar já está vinculada ou paga
  if (normalizarTexto(contaAPagarInfo.status) === "pago" || contaAPagarInfo.idTransacaoVinculada) {
    enviarMensagemTelegram(chatId, `ℹ️ A conta *${escapeMarkdown(contaAPagarInfo.descricao)}* já está paga ou vinculada.`);
    logToSheet(`Conta a Pagar ID ${idContaAPagar} já está paga ou vinculada.`, "INFO");
    return;
  }

  // Encontra a transação na aba Transacoes para confirmar que ela existe
  const transacoesSheet = ss.getSheetByName(SHEET_TRANSACOES);
  if (!transacoesSheet) {
    enviarMensagemTelegram(chatId, "❌ Erro: Aba 'Transacoes' não encontrada para verificar a transação.");
    logToSheet("Erro: Aba 'Transacoes' ausente para verificar transação.", "ERROR");
    return;
  }
  const dadosTransacoes = transacoesSheet.getDataRange().getValues();
  const headersTransacoes = transacoesSheet.getRange(1, 1, 1, transacoesSheet.getLastColumn()).getValues()[0];
  const colIdTransacao = headersTransacoes.indexOf('ID Transacao');

  let transacaoEncontrada = false;
  for (let i = 1; i < dadosTransacoes.length; i++) {
    if (dadosTransacoes[i][colIdTransacao] === idTransacao) {
      transacaoEncontrada = true;
      break;
    }
  }

  if (!transacaoEncontrada) {
    enviarMensagemTelegram(chatId, `❌ Transação com ID *${escapeMarkdown(idTransacao)}* não encontrada na aba 'Transacoes'.`);
    logToSheet(`Erro: Transacao ID ${idTransacao} não encontrada para vincular.`, "WARN");
    return;
  }

  try {
    const linhaReal = contaAPagarInfo.linha;
    const colStatus = contaAPagarInfo.headers.indexOf('Status') + 1;
    const colIDTransacaoVinculada = contaAPagarInfo.headers.indexOf('ID Transacao Vinculada') + 1;

    contasAPagarSheet.getRange(linhaReal, colStatus).setValue("Pago");
    contasAPagarSheet.getRange(linhaReal, colIDTransacaoVinculada).setValue(idTransacao);
    logToSheet(`Conta a Pagar '${contaAPagarInfo.descricao}' (ID: ${idContaAPagar}) vinculada a transacao ID: ${idTransacao} e marcada como PAGA.`, "INFO");
    enviarMensagemTelegram(chatId, `✅ Conta *${escapeMarkdown(contaAPagarInfo.descricao)}* vinculada com sucesso à transação *${escapeMarkdown(idTransacao)}* e marcada como paga!`);
    
    atualizarSaldosDasContas(); // Atualiza os saldos após a vinculação
  } catch (e) {
    logToSheet(`ERRO ao vincular transacao ID ${idTransacao} a conta a pagar ID ${idContaAPagar}: ${e.message} na linha ${e.lineNumber}. Stack: ${e.stack}`, "ERROR");
    enviarMensagemTelegram(chatId, `❌ Houve um erro ao vincular a transação. Por favor, tente novamente mais tarde. (Erro: ${e.message})`);
  }
}
