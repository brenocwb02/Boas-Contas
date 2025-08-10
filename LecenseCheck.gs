const LICENSE_SERVER_URL = "https://script.google.com/macros/s/AKfycbzxbwGNWISM_fhByxxMSUMjYW2fil83p42VRpHP9poFC06VgKGh0WMqtz2kaVGV_xKpbw/exec";

function activateProduct() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Ativação do Produto', 'Insira sua chave de licença:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const licenseKey = response.getResponseText().trim();
    if (validateLicense(licenseKey)) {
      ui.alert('Sucesso!', 'Produto ativado. Recarregue a planilha.', ui.ButtonSet.OK);
    }
  }
}

function validateLicense(licenseKey) {
  const userEmail = Session.getActiveUser().getEmail();
  const url = `${LICENSE_SERVER_URL}?key=${encodeURIComponent(licenseKey)}&email=${encodeURIComponent(userEmail)}`;

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const result = JSON.parse(response.getContentText());

    if (result.success) {
      PropertiesService.getScriptProperties().setProperties({
        'LICENSE_KEY': licenseKey,
        'LICENSE_STATUS': 'VALID',
        'LICENSED_USER': userEmail,
        'SYSTEM_STATUS': 'ACTIVATED'
      });
      return true;
    } else {
      SpreadsheetApp.getUi().alert('Erro', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
      return false;
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro de Conexão', 'Falha ao conectar com o servidor de licenciamento.', SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}

/**
 * **FUNÇÃO CORRIGIDA**
 * Verifica se a licença do produto é válida.
 * Esta nova versão funciona tanto para ações do utilizador na planilha quanto para o webhook do Telegram.
 */
function isLicenseValid() {
  const props = PropertiesService.getScriptProperties();
  const status = props.getProperty('LICENSE_STATUS');
  
  // A verificação principal e mais segura é simplesmente se o status da licença é 'VALID'.
  // O status só é definido como 'VALID' após uma chamada bem-sucedida ao servidor de licenciamento,
  // que já valida o e-mail do utilizador e a chave. Esta verificação funciona em todos os contextos.
  return status === 'VALID';
}
