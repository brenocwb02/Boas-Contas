/**
 * @file Constants.gs
 * @description Este arquivo centraliza todas as constantes e configurações globais do projeto,
 * como nomes de planilhas, chaves de cache e URLs de API.
 */
// --- ATUALIZADO: Chaves para PropertiesService ---
// Usado para armazenar o token do bot e o ID do chat do admin de forma segura e persistente.
const TELEGRAM_TOKEN_PROPERTY_KEY = 'TELEGRAM_TOKEN';
const ADMIN_CHAT_ID_PROPERTY_KEY = 'ADMIN_CHAT_ID';
const WEB_APP_URL_PROPERTY_KEY = 'WEB_APP_URL'; // NOVO: Chave para a URL do Web App
const SPEECH_API_KEY_PROPERTY_KEY = 'SPEECH_API_KEY'; // <-- ADICIONADO


/**
 * NOVO: Obtém o Chat ID do administrador das Propriedades do Script.
 * @returns {string} O Chat ID do administrador.
 */
function getAdminChatIdFromProperties() {
    const chatId = PropertiesService.getScriptProperties().getProperty(ADMIN_CHAT_ID_PROPERTY_KEY);
    return chatId;
}

const URL_BASE_TELEGRAM = "https://api.telegram.org/bot";

// --- Nomes das Abas da Planilha ---
const SHEET_TRANSACOES = "Transacoes";
const SHEET_CONTAS = "Contas";
const SHEET_PALAVRAS_CHAVE = "PalavrasChave";
const SHEET_USUARIOS = "Usuarios"; // Mantido para consistência, embora a lógica principal use SHEET_CONFIGURACOES
const SHEET_CONFIGURACOES = "Configuracoes";
const SHEET_LOGS_SISTEMA = "Logs_Sistema";
const SHEET_CATEGORIAS = "Categorias";
const SHEET_METAS = "Metas";
const SHEET_FATURAS = "Faturas"; // Mantido para consistência, pode ser usado em futuras implementações
const SHEET_ORCAMENTO = "Orcamento"; // Mantido para consistência
const SHEET_ALERTAS_ENVIADOS = "AlertasEnviados"; // Adicionado para rastrear alertas
const SHEET_CONTAS_A_PAGAR = "Contas_a_Pagar"; // Adicionado para contas fixas
const SHEET_NOTIFICACOES_CONFIG = 'Notificacoes_Config'; // Adicionado para configuracoes de notificacoes
const SHEET_LEARNED_CATEGORIES = "LearnedCategories";

// --- Constantes de Cache ---
const CACHE_KEY_PALAVRAS = 'palavras_chave_cache';
const CACHE_KEY_CONTAS = 'contas_cache';
const CACHE_KEY_CONFIG = 'config_cache';
const CACHE_KEY_TUTORIAL_STATE = 'tutorial_state';
const CACHE_KEY_PENDING_TRANSACTIONS = 'pending_transaction';
const CACHE_KEY_EDIT_STATE = 'edit_state';
const CACHE_KEY_ASSISTANT_STATE = 'assistant_state'; // NOVO: Chave para o estado do assistente
const CACHE_KEY_CATEGORIAS = 'categorias_cache';
// *** CORREÇÃO: Constantes em falta adicionadas aqui ***
const CACHE_KEY_TRANSACOES = 'transacoes_cache';
const CACHE_KEY_CONTAS_A_PAGAR = 'contas_a_pagar_cache';
// ******************************************************

// NOVO: Adicionadas constantes para o token de acesso ao dashboard
const CACHE_KEY_DASHBOARD_TOKEN = 'dashboard_access_token';
const CACHE_EXPIRATION_DASHBOARD_TOKEN_SECONDS = 300; // 5 minutos de validade para o token
const CACHE_EXPIRATION_SECONDS = 21600; // 6 horas
const CACHE_EXPIRATION_PENDING_TRANSACTION_SECONDS = 900; // 15 minutos
const CACHE_EXPIRATION_TUTORIAL_STATE_SECONDS = 1800; // 30 minutos
const CACHE_EXPIRATION_EDIT_STATE_SECONDS = 900; // 15 minutos

// --- Constantes de Lógica Financeira ---
const SIMILARITY_THRESHOLD = 0.75; // Limite de similaridade para auto-vinculação (0 a 1)
const BUDGET_ALERT_THRESHOLD_PERCENT = 80; // Percentual para enviar alerta de orçamento
const BILL_REMINDER_DAYS_BEFORE = 3; // Dias antes do vencimento para enviar lembrete

// --- Constantes de Lógica de Tutorial ---
const TUTORIAL_STATE_WAITING_DESPESA = "waiting_for_despesa";
const TUTORIAL_STATE_WAITING_RECEITA = "waiting_for_receita";
const TUTORIAL_STATE_WAITING_SALDO = "waiting_for_saldo";
const TUTORIAL_STATE_WAITING_CONTAS_A_PAGAR = "waiting_for_contas_a_pagar";



// --- Níveis de Log ---
// Mapeia os níveis de log para valores numéricos para facilitar a comparação.
const LOG_LEVEL_MAP = {
  "DEBUG": 1,
  "INFO": 2,
  "WARN": 3,
  "ERROR": 4,
  "NONE": 5
};
// Variável global para armazenar o nível de log atual, inicializada com um padrão seguro.
let currentLogLevel = "INFO";

// --- Regex para Pagamento de Fatura ---
// Captura variações de "paguei [valor] da fatura do [cartão] com [conta]"
const regexPagamentoFatura = /paguei\s+(?:r\$)?\s*([\d.,]+)\s+da\s+fatura\s+(?:do|da|de)?\s*(.+?)\s+com\s+(.+)/i;

/**
 * Define os cabeçalhos esperados para cada aba da planilha.
 * Ajuda a acessar colunas por nome em vez de índice, tornando o código mais legível e robusto.
 * Os nomes devem corresponder exatamente aos cabeçalhos na sua planilha.
 */
const HEADERS = {
  // CORRIGIDO: Os nomes das colunas foram atualizados para corresponder às suas planilhas.
  [SHEET_TRANSACOES]: ["Data", "Descricao", "Categoria", "Subcategoria", "Tipo", "Valor", "Metodo de Pagamento", "Conta / Cartão", "Parcela Total", "Parcela Atual", "Data de Vencimento", "Usuario", "Status", "ID Transacao", "Data de Registro"],
  [SHEET_CONTAS]: ["Nome da Conta", "Tipo", "Banco", "Saldo Inicial", "Saldo Atual", "Limite", "Dia de Vencimento", "Status", "Categoria", "Dia de Fechamento", "Tipo de Fechamento", "Dias Antes Vencimento", "Conta Pai Agrupador", "Pessoa"],
  [SHEET_USUARIOS]: ["ID Usuario Telegram", "Nome", "Permissoes", "Ativo"],
  [SHEET_CONFIGURACOES]: ["chave", "valor", "NomeUsuario", "grupo"],
  [SHEET_CATEGORIAS]: ["Categoria", "Subcategoria", "Tipo", "Tipo de Gasto"], // Seus cabeçalhos podem variar
  [SHEET_METAS]: ["Categoria", "Subcategoria", "janeiro/2025", "Total Geral"], // Exemplo, ajuste conforme sua planilha
  [SHEET_FATURAS]: ["ID Fatura", "Cartao", "Mes Referencia", "Data Fechamento", "Data Vencimento", "Valor Total", "Valor Pago", "Status", "ID Transacao Pagamento"], // Seus cabeçalhos podem variar
  [SHEET_ORCAMENTO]: ["ID Orcamento", "Mes referencia", "Categoria", "Valor Orcado", "Valor Gasto"], // Seus cabeçalhos podem variar
  [SHEET_NOTIFICACOES_CONFIG]: ["Chat ID", "Usuário", "Alertas Orçamento", "Lembretes Contas a Pagar", "Resumo Diário", "Hora Resumo Diário (HH:mm)", "Resumo Semanal", "Dia Resumo Semanal (0-6)", "Hora Resumo Semanal (HH:mm)","Alertas de Fatura"],
  [SHEET_PALAVRAS_CHAVE]: ["Tipo da Palavra-Chave", "Palavra-Chave", "Valor Interpretado"], // Seus cabeçalhos podem variar
  [SHEET_LEARNED_CATEGORIES]: ["Keyword", "Categoria", "Subcategoria", "ConfidenceScore", "LastUpdated"]
};
