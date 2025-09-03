/**
 * @file Quiz.gs
 * @description Contém toda a lógica e os dados para o Quiz de Perfil Financeiro.
 */

// Chave para armazenar o estado do quiz do usuário no cache
const QUIZ_STATE_KEY = 'financial_quiz_state';

// --- BANCO DE PERGUNTAS E RESPOSTAS ---
// Cada opção de resposta atribui pontos a um perfil específico.
const QUIZ_QUESTIONS = [
  {
    question: "1/5: Quando você recebe um dinheiro extra e inesperado, qual é a sua primeira reação?",
    options: [
      { text: "Guardar a maior parte para uma meta futura.", profile: "Planejador", points: 2 },
      { text: "Pensar em como posso investi-lo para que renda mais.", profile: "Construtor", points: 2 },
      { text: "Gastar com algo que eu quero há muito tempo.", profile: "Despreocupado", points: 2 },
      { text: "Usar para realizar um pequeno sonho ou ter uma experiência.", profile: "Sonhador", points: 2 },
      { text: "Depende do meu humor e da situação no momento.", profile: "Camaleão", points: 2 }
    ]
  },
  {
    question: "2/5: Como você descreveria a sua relação com o seu orçamento mensal?",
    options: [
      { text: "Tenho uma planilha ou app e sigo o plano à risca.", profile: "Planejador", points: 2 },
      { text: "Foco mais em aumentar a minha renda do que em controlar gastos pequenos.", profile: "Construtor", points: 2 },
      { text: "Orçamento? Eu apenas tento não gastar mais do que ganho.", profile: "Despreocupado", points: 2 },
      { text: "É difícil seguir um orçamento com tantos sonhos para realizar.", profile: "Sonhador", points: 2 },
      { text: "Às vezes controlo, outras vezes sou mais flexível.", profile: "Camaleão", points: 2 }
    ]
  },
  {
    question: "3/5: Ao pensar em investimentos, o que é mais importante para você?",
    options: [
      { text: "A segurança e a previsibilidade do retorno, mesmo que seja menor.", profile: "Planejador", points: 2 },
      { text: "O potencial de alto crescimento, mesmo que envolva mais risco.", profile: "Construtor", points: 2 },
      { text: "Investir parece muito complicado, prefiro não pensar nisso.", profile: "Despreocupado", points: 2 },
      { text: "O quanto esse investimento me aproxima de um grande objetivo de vida.", profile: "Sonhador", points: 2 },
      { text: "Sigo as dicas de amigos ou o que está em alta no momento.", profile: "Camaleão", points: 2 }
    ]
  },
  {
    question: "4/5: Qual destas frases melhor descreve a sua atitude em relação a dívidas?",
    options: [
      { text: "Evito ao máximo. Só faço se for algo muito planeado, como um imóvel.", profile: "Planejador", points: 2 },
      { text: "Vejo como uma ferramenta. Posso usar para alavancar um negócio ou oportunidade.", profile: "Construtor", points: 2 },
      { text: "Acabo por me endividar com coisas do dia a dia, como o cartão de crédito.", profile: "Despreocupado", points: 2 },
      { text: "Às vezes, faço dívidas para realizar um sonho ou uma viagem.", profile: "Sonhador", points: 2 },
      { text: "A minha situação com dívidas varia muito de mês para mês.", profile: "Camaleão", points: 2 }
    ]
  },
    {
    question: "5/5: O que mais lhe traria satisfação financeira?",
    options: [
      { text: "Saber que tenho um futuro financeiro seguro e bem planeado.", profile: "Planejador", points: 2 },
      { text: "Construir um património sólido e ver os meus ativos a crescer.", profile: "Construtor", points: 2 },
      { text: "Ter dinheiro suficiente para não me preocupar com as contas do dia a dia.", profile: "Despreocupado", points: 2 },
      { text: "Poder realizar os meus maiores sonhos sem me preocupar com o custo.", profile: "Sonhador", points: 2 },
      { text: "Ter flexibilidade para lidar com o que quer que a vida me traga.", profile: "Camaleão", points: 2 }
    ]
  }
];

/**
 * Inicia o Quiz de Perfil Financeiro para um usuário.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {string} usuario O nome do usuário.
 */
function startFinancialQuiz(chatId, usuario) {
  const initialState = {
    currentQuestion: 0,
    scores: {
      "Planejador": 0,
      "Construtor": 0,
      "Despreocupado": 0,
      "Sonhador": 0,
      "Camaleão": 0
    }
  };
  setQuizState(chatId, initialState);
  sendQuizQuestion(chatId, initialState);
}

/**
 * Envia a pergunta atual do quiz para o usuário.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {Object} state O estado atual do quiz.
 */
function sendQuizQuestion(chatId, state) {
  const questionIndex = state.currentQuestion;
  
  if (questionIndex >= QUIZ_QUESTIONS.length) {
    finishQuiz(chatId, state);
    return;
  }

  const questionData = QUIZ_QUESTIONS[questionIndex];
  const keyboard = {
    inline_keyboard: questionData.options.map((option, index) => {
      return [{ text: option.text, callback_data: `quiz_${questionIndex}_${index}` }];
    })
  };

  enviarMensagemTelegram(chatId, questionData.question, { reply_markup: keyboard });
}

/**
 * Processa a resposta do usuário a uma pergunta do quiz.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {number} questionIndex O índice da pergunta respondida.
 * @param {number} optionIndex O índice da opção escolhida.
 */
function handleQuizAnswer(chatId, questionIndex, optionIndex) {
  const state = getQuizState(chatId);
  if (!state || state.currentQuestion !== questionIndex) {
    enviarMensagemTelegram(chatId, "Houve um problema com o quiz. Por favor, comece novamente com /meuperfil.");
    clearQuizState(chatId);
    return;
  }

  const chosenOption = QUIZ_QUESTIONS[questionIndex].options[optionIndex];
  if (chosenOption) {
    state.scores[chosenOption.profile] += chosenOption.points;
  }

  state.currentQuestion++;
  setQuizState(chatId, state);
  sendQuizQuestion(chatId, state);
}

/**
 * Finaliza o quiz, calcula o resultado e envia para o usuário.
 * @param {string} chatId O ID do chat do Telegram.
 * @param {Object} state O estado final do quiz.
 */
function finishQuiz(chatId, state) {
  let finalProfile = "Camaleão"; // Perfil padrão
  let maxScore = 0;

  for (const profile in state.scores) {
    if (state.scores[profile] > maxScore) {
      maxScore = state.scores[profile];
      finalProfile = profile;
    }
  }

  // Salva o perfil do usuário nas propriedades para uso futuro
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty(`PROFILE_${chatId}`, finalProfile);

  const profileDescription = getProfileDescription(finalProfile);
  const message = `🎉 **Quiz Finalizado!** 🎉\n\nO seu perfil financeiro é: *${finalProfile}*\n\n${profileDescription}`;
  
  enviarMensagemTelegram(chatId, message);
  clearQuizState(chatId);
}

/**
 * Obtém a descrição de um perfil financeiro específico.
 * @param {string} profileName O nome do perfil.
 * @returns {string} A descrição do perfil.
 */
function getProfileDescription(profileName) {
  // Simplesmente retorna o conteúdo do arquivo Markdown correspondente.
  // Criaremos um novo arquivo .md para armazenar estas descrições.
  const descriptions = {
    'Construtor': 'Você tem foco em construir património e não tem medo de correr riscos calculados para ver o seu dinheiro render. O seu maior desafio é, por vezes, não dar a devida atenção aos pequenos gastos do dia a dia.',
    'Planejador': 'Organização e segurança são os seus lemas. Você gosta de ter tudo sob controlo, com planilhas e metas bem definidas. O seu desafio é, ocasionalmente, permitir-se alguma flexibilidade e aproveitar oportunidades inesperadas.',
    'Camaleão': 'Você adapta-se facilmente às circunstâncias. Ora está a controlar tudo, ora está mais relaxado. Essa flexibilidade é uma força, mas o seu desafio é manter a consistência nos seus objetivos de longo prazo.',
    'Despreocupado': 'Viver o presente é a sua prioridade. Você lida com o dinheiro conforme as necessidades surgem, sem muito planeamento. O seu maior desafio é criar o hábito de poupar para o futuro e evitar que as dívidas se acumulem.',
    'Sonhador': 'Você é motivado por grandes objetivos e sonhos. O dinheiro é um meio para realizar essas aspirações. O seu desafio é transformar esses grandes sonhos em pequenos passos financeiros, criando um plano concreto para os alcançar.'
  };
  return descriptions[profileName] || "Descrição não encontrada.";
}


// --- Funções de Gestão de Estado (Cache) ---

function setQuizState(chatId, state) {
  CacheService.getScriptCache().put(`${QUIZ_STATE_KEY}_${chatId}`, JSON.stringify(state), 900); // 15 min de validade
}

function getQuizState(chatId) {
  const cached = CacheService.getScriptCache().get(`${QUIZ_STATE_KEY}_${chatId}`);
  return cached ? JSON.parse(cached) : null;
}

function clearQuizState(chatId) {
  CacheService.getScriptCache().remove(`${QUIZ_STATE_KEY}_${chatId}`);
}


/**
 * **FUNÇÃO ATUALIZADA**
 * Agora, além de retornar os dados do quiz, também pode retornar apenas o perfil do utilizador.
 * @param {string} chatId O ID do chat do Telegram.
 * @returns {Object|null} O objeto de estado do quiz (se estiver em andamento) ou um objeto com o perfil finalizado.
 */
function getFinancialProfile(chatId) {
    // 1. Verifica se há um quiz em andamento no cache
    const quizState = getQuizState(chatId);
    if (quizState) {
        return { inProgress: true, state: quizState };
    }

    // 2. Se não houver quiz ativo, busca o perfil já definido nas propriedades do utilizador
    const userProps = PropertiesService.getUserProperties();
    const profile = userProps.getProperty(`PROFILE_${chatId}`);

    if (profile) {
        return { inProgress: false, perfil: profile };
    }

    // 3. Se não encontrou em nenhum dos locais, retorna nulo
    return null;
}
