<!DOCTYPE html>
<html>
<head>
    <base target="_top">
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="p-4 bg-gray-50 text-gray-800">
    <h3 class="text-lg font-bold mb-4 text-gray-900">Configurações Rápidas</h3>
    
    <!-- Seção para adicionar contas -->
    <div class="mb-6 p-4 bg-white rounded-lg shadow-sm border border-gray-200">
        <h4 class="font-semibold mb-3 text-gray-700">1. Adicionar Conta / Cartão</h4>
        <p class="text-sm text-gray-500 mb-3">Primeiro, adicione as suas contas bancárias e cartões de crédito aqui.</p>
        <input id="accountName" type="text" placeholder="Nome (ex: Nubank, Itaú)" class="w-full p-2 border rounded mb-2">
        <select id="accountType" class="w-full p-2 border rounded mb-2 text-gray-500">
            <option value="Conta Corrente">Conta Corrente</option>
            <option value="Cartão de Crédito">Cartão de Crédito</option>
            <option value="Dinheiro Físico">Dinheiro Físico</option>
            <option value="Fatura Consolidada">Fatura Consolidada</option>
        </select>
        <button onclick="addAccount()" class="w-full bg-blue-500 text-white p-2 rounded hover:bg-blue-600 transition-colors">Adicionar Conta</button>

        <div class="mt-4 p-3 bg-blue-50 border border-blue-200 rounded-lg">
            <h5 class="font-semibold text-blue-800 mb-2">💡 Dica sobre Cartões de Crédito</h5>
            <p class="text-sm text-blue-700">
                Após adicionar um cartão, vá à aba <strong>`Contas`</strong> para configurar o dia de fechamento e vencimento. O bot usa essa informação para calcular as faturas corretamente.
            </p>
            <ul class="text-sm text-blue-700 list-disc list-inside mt-2 space-y-1">
                <li><strong>fechamento-anterior:</strong> Use quando o fechamento e o vencimento ocorrem no mesmo mês (ex: fecha dia 4, vence dia 11).</li>
                <li><strong>fechamento-mes:</strong> Use quando o fechamento ocorre num mês e o vencimento no mês seguinte (ex: fecha dia 29/Jul, vence dia 10/Ago).</li>
            </ul>
        </div>
    </div>

    <!-- Seção para adicionar palavras-chave -->
    <div class="p-4 bg-white rounded-lg shadow-sm border border-gray-200">
        <h4 class="font-semibold mb-3 text-gray-700">2. Ensinar o Bot (Palavras-Chave)</h4>
        <p class="text-sm text-gray-500 mb-3">Agora, ensine o bot a reconhecer os seus lançamentos. Selecione o que quer ensinar:</p>
        
        <select id="keywordTypeSelect" onchange="updateFormFields()" class="w-full p-2 border rounded mb-3 text-gray-500">
            <option value="categoria">Categorizar um Gasto (ex: ifood)</option>
            <option value="conta">Dar um Apelido a uma Conta (ex: nu)</option>
        </select>

        <!-- Campos para Categoria/Subcategoria -->
        <div id="categoriaFields">
            <input id="categoryKeyword" type="text" placeholder="Palavra-Chave (ex: ifood, uber)" class="w-full p-2 border rounded mb-2">
            <!-- **CAMPO ATUALIZADO** -->
            <select id="mainCategory" class="w-full p-2 border rounded mb-2 text-gray-500">
                <option value="">Selecione a Categoria Principal</option>
            </select>
            <input id="subcategory" type="text" placeholder="Subcategoria (ex: Delivery)" class="w-full p-2 border rounded mb-2">
        </div>

        <!-- Campos para Apelido de Conta -->
        <div id="contaFields" class="hidden">
            <input id="accountNickname" type="text" placeholder="Apelido (ex: nu, roxinho)" class="w-full p-2 border rounded mb-2">
            <!-- **CAMPO ATUALIZADO** -->
            <select id="accountRealName" class="w-full p-2 border rounded mb-2 text-gray-500">
                <option value="">Selecione a Conta Real</option>
            </select>
        </div>

        <button onclick="addKeyword()" class="w-full bg-green-500 text-white p-2 rounded hover:bg-green-600 transition-colors mt-2">Adicionar Palavra-Chave</button>
    </div>
    
    <div id="status" class="mt-4 text-center font-medium"></div>

    <script>
        // **NOVA LÓGICA** para popular os menus suspensos
        document.addEventListener("DOMContentLoaded", function() {
            showStatus("A carregar dados...", false);
            google.script.run
                .withSuccessHandler(populateDropdowns)
                .withFailureHandler(error => showStatus("Erro ao carregar dados: " + error.message, true))
                .getSidebarData();
            updateFormFields();
        });

        function populateDropdowns(data) {
            if (data.error) {
                showStatus("Erro: " + data.error, true);
                return;
            }

            const accountSelect = document.getElementById('accountRealName');
            data.accounts.forEach(account => {
                const option = document.createElement('option');
                option.value = account;
                option.textContent = account;
                accountSelect.appendChild(option);
            });

            const categorySelect = document.getElementById('mainCategory');
            data.categories.forEach(category => {
                const option = document.createElement('option');
                option.value = category;
                option.textContent = category;
                categorySelect.appendChild(option);
            });
            showStatus(""); // Limpa a mensagem de "a carregar"
        }

        function showStatus(message, isError = false) {
            const statusDiv = document.getElementById('status');
            statusDiv.textContent = message;
            statusDiv.className = 'mt-4 text-center font-medium ' + (isError ? 'text-red-500' : 'text-green-500');
        }

        function updateFormFields() {
            const selectedType = document.getElementById('keywordTypeSelect').value;
            document.getElementById('categoriaFields').classList.toggle('hidden', selectedType !== 'categoria');
            document.getElementById('contaFields').classList.toggle('hidden', selectedType !== 'conta');
        }

        function addAccount() {
            const name = document.getElementById('accountName').value;
            const type = document.getElementById('accountType').value;
            google.script.run
                .withSuccessHandler(response => {
                    showStatus(response.message, !response.success);
                    if(response.success) {
                        // Recarrega os dados da barra lateral para incluir a nova conta
                        google.script.run.withSuccessHandler(populateDropdowns).getSidebarData();
                    }
                })
                .withFailureHandler(error => showStatus(error.message, true))
                .addAccountToSheet(name, type);
        }

        function addKeyword() {
            const selectedType = document.getElementById('keywordTypeSelect').value;
            
            if (selectedType === 'categoria') {
                const keyword = document.getElementById('categoryKeyword').value;
                const mainCat = document.getElementById('mainCategory').value;
                const subCat = document.getElementById('subcategory').value;
                google.script.run
                    .withSuccessHandler(response => showStatus(response.message, !response.success))
                    .withFailureHandler(error => showStatus(error.message, true))
                    .addKeywordToSheet('categoria', keyword, mainCat, subCat);
            } else if (selectedType === 'conta') {
                const nickname = document.getElementById('accountNickname').value;
                const realName = document.getElementById('accountRealName').value;
                google.script.run
                    .withSuccessHandler(response => showStatus(response.message, !response.success))
                    .withFailureHandler(error => showStatus(error.message, true))
                    .addKeywordToSheet('conta', nickname, realName, null);
            }
        }
    </script>
</body>
</html>
