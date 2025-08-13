/**
 * @file Setup.gs
 * @description Contém a lógica para a inicialização e configuração do sistema pelo cliente,
 * agora usando um método de cópia de uma planilha modelo.
 */

// **IMPORTANTE:** Cole aqui o ID da sua Planilha 'Modelo' Mestra criada no Passo 1.
const MODEL_SPREADSHEET_ID = "1_8LsJNV89HEaMZzYjlfgOM4V-eJ63GeeTQ9h61C4wvs";

/**
 * **FUNÇÃO ATUALIZADA**
 * Agora, em vez de criar cada aba manualmente, esta função copia todas as abas
 * de uma planilha modelo pré-formatada para a planilha do cliente.
 */
function initializeSystem() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
        'Confirmação de Inicialização',
        'Este processo irá configurar a sua planilha com todas as abas e formatações padrão. Deseja continuar?',
        ui.ButtonSet.YES_NO
    );

    if (response == ui.Button.YES) {
        try {
            const destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
            destinationSpreadsheet.toast('A preparar a sua planilha... Este processo pode demorar um momento.', 'Inicialização', -1);

            // Abre a planilha modelo de onde vamos copiar as abas
            const modelSpreadsheet = SpreadsheetApp.openById(MODEL_SPREADSHEET_ID);
            const modelSheets = modelSpreadsheet.getSheets();

            // Copia cada aba do modelo para a planilha atual do cliente
            modelSheets.forEach(sheet => {
                const sheetName = sheet.getName();
                // Verifica se uma aba com o mesmo nome já não existe (caso o utilizador execute duas vezes)
                if (!destinationSpreadsheet.getSheetByName(sheetName)) {
                    sheet.copyTo(destinationSpreadsheet).setName(sheetName);
                    logToSheet(`Aba '${sheetName}' copiada com sucesso do modelo.`, "INFO");
                }
            });

            // Apaga a aba "Página1" ou "Sheet1" que é criada por defeito
            const defaultSheet = destinationSpreadsheet.getSheetByName('Página1') || destinationSpreadsheet.getSheetByName('Sheet1');
            if (defaultSheet) {
                destinationSpreadsheet.deleteSheet(defaultSheet);
            }

            // Esconde a aba de boas-vindas
            const welcomeSheet = destinationSpreadsheet.getSheetByName('✅ Bem-vindo');
            if (welcomeSheet) {
                welcomeSheet.hideSheet();
            }

            PropertiesService.getScriptProperties().setProperty('SYSTEM_STATUS', 'INITIALIZED');
            
            destinationSpreadsheet.toast('Sistema inicializado com sucesso! Por favor, recarregue a página.', 'Sucesso!', 5);
            ui.alert('Sucesso!', 'O sistema foi inicializado. O próximo passo é configurar o seu bot no menu "Gasto Certo > Configuração do Bot (Telegram)".', ui.ButtonSet.OK);

        } catch (e) {
            logToSheet(`ERRO CRÍTICO durante a inicialização por cópia: ${e.message}`, "ERROR");
            ui.alert('Erro na Inicialização', `Ocorreu um erro ao configurar a sua planilha: ${e.message}. Verifique se o ID da planilha modelo está correto e se você tem permissão para a aceder.`, ui.ButtonSet.OK);
        }
    }
}


/**
 * Obtém a URL para o editor de scripts do projeto atual.
 * O usuário será levado diretamente para a página correta.
 */
function getScriptEditorUrl() {
  const scriptId = ScriptApp.getScriptId();
  return `https://script.google.com/d/${scriptId}/edit`;
}

// O resto das funções (showConfigurationSidebar, getSidebarData, etc.) permanecem as mesmas
// pois elas apenas leem e escrevem dados, independentemente de como as abas foram criadas.

function showConfigurationSidebar() {
    const template = HtmlService.createTemplateFromFile('Configuration');
    const html = template.evaluate()
        .setTitle('Configurações do Sistema');
    SpreadsheetApp.getUi().showSidebar(html);
}

function getSidebarData() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const accountsSheet = ss.getSheetByName("Contas");
        const accountsData = accountsSheet.getRange("A2:A" + accountsSheet.getLastRow()).getValues();
        const accountNames = accountsData.map(row => row[0]).filter(Boolean);
        const categoriesSheet = ss.getSheetByName("Categorias");
        const categoriesData = categoriesSheet.getRange("A2:A" + categoriesSheet.getLastRow()).getValues();
        const mainCategories = [...new Set(categoriesData.map(row => row[0]).filter(Boolean))];
        return {
            accounts: accountNames,
            categories: mainCategories
        };
    } catch (e) {
        return { error: e.message };
    }
}

function addAccountToSheet(accountName, accountType) {
    try {
        if (!accountName || !accountType) {
            throw new Error("O nome e o tipo da conta são obrigatórios.");
        }
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("Contas");
        if (!sheet) throw new Error("Aba 'Contas' não encontrada.");
        const data = sheet.getDataRange().getValues();
        const accountExists = data.some(row => row[0].toLowerCase() === accountName.toLowerCase());
        if (accountExists) {
            throw new Error(`A conta "${accountName}" já existe.`);
        }
        sheet.appendRow([accountName, accountType, "", 0, 0, "", "", "Ativo", "", "", "", "", "", ""]);
        CacheService.getScriptCache().remove('contas_cache');
        return { success: true, message: `Conta "${accountName}" adicionada com sucesso!` };
    } catch (e) {
        return { success: false, message: e.message };
    }
}

function addKeywordToSheet(keywordType, keyword, value1, value2) {
    try {
        if (!keywordType || !keyword || !value1) {
            throw new Error("Todos os campos são obrigatórios.");
        }
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("PalavrasChave");
        if (!sheet) throw new Error("Aba 'PalavrasChave' não encontrada.");
        let rowData;
        if (keywordType === 'conta') {
            rowData = ["conta", keyword.toLowerCase(), value1, ""];
        } else if (keywordType === 'categoria') {
            if (!value2) throw new Error("A subcategoria é obrigatória.");
            const valorInterpretado = `${value1} > ${value2}`;
            rowData = ["subcategoria", keyword.toLowerCase(), valorInterpretado, "Despesa"];
        } else {
            throw new Error("Tipo de palavra-chave inválido.");
        }
        sheet.appendRow(rowData);
        CacheService.getScriptCache().remove('palavras_chave_cache');
        return { success: true, message: `Palavra-chave "${keyword}" adicionada com sucesso!` };
    } catch (e) {
        return { success: false, message: e.message };
    }
}
