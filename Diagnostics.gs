/**
 * @file Diagnostics.gs
 * @description Contém funções para verificar a saúde e a configuração do sistema.
 */

/**
 * Executa uma série de verificações nos pontos críticos da configuração do sistema
 * e exibe um relatório para o utilizador.
 */
function runSystemDiagnostics() {
  const ui = SpreadsheetApp.getUi();
  SpreadsheetApp.getActiveSpreadsheet().toast('A executar diagnóstico...', 'Verificação do Sistema', 5);

  let report = "🔍 <strong>Relatório de Diagnóstico do Sistema</strong><br><br>";
  let allOk = true;

  try {
    // 1. Verificar Token do Telegram
    const token = PropertiesService.getScriptProperties().getProperty(TELEGRAM_TOKEN_PROPERTY_KEY);
    if (token && token.length > 20) {
      report += "🔑 Token do Telegram: OK ✅<br>";
    } else {
      report += "🔑 Token do Telegram: ❌ FALHA (Não encontrado. Configure em 'Configuração do Bot').<br>";
      allOk = false;
    }

    // 2. Verificar Chat ID do Admin
    const chatId = PropertiesService.getScriptProperties().getProperty(ADMIN_CHAT_ID_PROPERTY_KEY);
    if (chatId) {
      report += "👤 Chat ID do Admin: OK ✅<br>";
    } else {
      report += "👤 Chat ID do Admin: ❌ FALHA (Não encontrado. Configure em 'Configuração do Bot').<br>";
      allOk = false;
    }

    // 3. Verificar URL do Web App
    const webAppUrl = PropertiesService.getScriptProperties().getProperty(WEB_APP_URL_PROPERTY_KEY);
    if (webAppUrl && webAppUrl.startsWith("https://script.google.com/macros/s/")) {
      report += "🌐 URL do Web App: OK ✅<br>";
    } else {
      report += "🌐 URL do Web App: ❌ FALHA (URL inválida ou não encontrada. Configure em 'Configuração do Bot').<br>";
      allOk = false;
    }

    // 4. Verificar Webhook do Telegram (só se o token e a URL estiverem OK)
    if (token && webAppUrl) {
      const webhookInfoUrl = `https://api.telegram.org/bot${token}/getWebhookInfo`;
      const response = UrlFetchApp.fetch(webhookInfoUrl, { muteHttpExceptions: true });
      const result = JSON.parse(response.getContentText());

      if (result.ok && result.result.url === webAppUrl) {
        report += "🔗 Webhook do Telegram: OK ✅ (Conectado corretamente a esta planilha).<br>";
      } else if (result.ok && result.result.url) {
        report += `🔗 Webhook do Telegram: ❌ FALHA (O bot está conectado a outra URL).<br>`;
        report += "   ↳ <strong>Solução:</strong> Vá a 'Configuração do Bot' e clique em 'Salvar' novamente para reconfigurar.<br>";
        allOk = false;
      } else {
        report += `🔗 Webhook do Telegram: ❌ FALHA (Não foi possível verificar. Erro: ${result.description || 'Desconhecido'}).<br>`;
        allOk = false;
      }
    } else {
      report += "🔗 Webhook do Telegram: ⚠️ PENDENTE (Token ou URL do Web App em falta).<br>";
    }

    report += "<br>------------------------------------<br>";
    if (allOk) {
      report += "🎉 <strong>O seu sistema parece estar configurado corretamente!</strong>";
    } else {
      report += "⚠️ <strong>Foram encontrados problemas na configuração. Por favor, siga as sugestões acima para os corrigir.</strong>";
    }

    // Usar a função de formatação para exibir o relatório
    showFormattedAlert("Diagnóstico do Sistema", report);

  } catch (e) {
    logToSheet(`Erro durante o diagnóstico do sistema: ${e.message}`, "ERROR");
    ui.alert("Erro", `Ocorreu um erro inesperado ao executar o diagnóstico: ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Exibe um alerta formatado em HTML para uma melhor visualização.
 * @param {string} title O título da caixa de diálogo.
 * @param {string} htmlMessage A mensagem a ser exibida em formato HTML.
 */
function showFormattedAlert(title, htmlMessage) {
  const htmlOutput = HtmlService.createHtmlOutput(
    `<div style="font-family: sans-serif; font-size: 14px; line-height: 1.6;">${htmlMessage}</div>`
  )
  .setWidth(450)
  .setHeight(300);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}
