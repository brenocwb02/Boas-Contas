const LICENSE_SERVER_URL = "https://script.google.com/macros/s/AKfycbzxbwGNWISM_fhByxxMSUMjYW2fil83p42VRpHP9poFC06VgKGh0WMqtz2kaVGV_xKpbw/exec";

function activateProduct() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Ativação do Produto', 'Insira sua chave de licença:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const licenseKey = response.getResponseText().trim();
    if (validateLicense(licenseKey)) {
      ui.alert('Sucesso!', 'Produto ativado com sucesso. Por favor, recarregue a planilha para que todas as funcionalidades sejam carregadas.', ui.ButtonSet.OK);
    }
  }
}

function validateLicense(licenseKey) {
  const userEmail = Session.getActiveUser().getEmail();
  // NOVO: Obtém o ID único desta planilha
  const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  // Envia o ID da planilha para o servidor de licenças para validação
  const url = `${LICENSE_SERVER_URL}?key=${encodeURIComponent(licenseKey)}&email=${encodeURIComponent(userEmail)}&sheetId=${encodeURIComponent(spreadsheetId)}`;

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const result = JSON.parse(response.getContentText());

    if (result.success) {
      // Salva todas as informações da licença, incluindo o ID da planilha
      PropertiesService.getScriptProperties().setProperties({
        'LICENSE_KEY': licenseKey,
        'LICENSE_STATUS': 'VALID',
        'LICENSED_USER': userEmail,
        'LICENSED_SHEET_ID': spreadsheetId, // Armazena o ID da planilha licenciada
        'SYSTEM_STATUS': 'ACTIVATED'
      });
      return true;
    } else {
      SpreadsheetApp.getUi().alert('Erro de Ativação', result.message, SpreadsheetApp.getUi().ButtonSet.OK);
      return false;
    }
  } catch (e) {
    SpreadsheetApp.getUi().alert('Erro de Conexão', 'Falha ao conectar com o servidor de licenciamento. Verifique a sua conexão à internet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}

/**
 * **FUNÇÃO ATUALIZADA E MAIS ROBUSTA**
 * Verifica se a licença do produto é válida para ESTA planilha específica.
 */
function isLicenseValid() {
  const props = PropertiesService.getScriptProperties();
  const status = props.getProperty('LICENSE_STATUS');
  const licensedSheetId = props.getProperty('LICENSED_SHEET_ID');
  const currentSheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  // A verificação agora é dupla: o status deve ser 'VALID' E o ID da planilha
  // atual deve ser o mesmo que foi guardado durante a ativação.
  return status === 'VALID' && licensedSheetId === currentSheetId;
}
