function verificarAtivacao() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Configuracoes');
  const data = sheet.getDataRange().getValues();
  const chaveDigitada = sheet.getRange('B9').getValue();
  const emailAtual = Session.getActiveUser().getEmail();
  const CHAVE_CORRETA = '12345-ABCDE'; // Substitua pela sua chave real

  let autorizado = false;

  for (let i = 1; i < data.length; i++) {
    const email = data[i][0];
    const status = data[i][1];
    if ((email === emailAtual && status === 'ATIVA') || chaveDigitada === CHAVE_CORRETA) {
      autorizado = true;
      break;
    }
  }

  if (!autorizado) {
    SpreadsheetApp.getUi().alert("❌ Esta cópia não está ativada. O acesso está bloqueado.");

    const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const name = sheets[i].getName();
      if (name !== 'Configuracoes') {
        sheets[i].hideSheet();
      }
    }

    throw new Error("⛔ Acesso negado.");
  }
}



function mostrarTodasAbas() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  sheets.forEach(s => s.showSheet());
}
