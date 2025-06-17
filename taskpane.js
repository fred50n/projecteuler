// taskpane.js - Versão que busca problemas individualmente

// Aguarda o ambiente do Office estar pronto
Office.onReady(info => {
    if (info.host === Office.HostType.Excel) {
        initializeApp();
    }
});

let currentProblemId = 1;

async function initializeApp() {
    // Associa nossas funções aos cliques dos botões
    document.getElementById("verify-button").onclick = openProblemOnWeb;
    document.getElementById("next-problem-button").onclick = goToNextProblem;

    await loadProgress();           // 1. Carrega o ID do último problema salvo
    await fetchAndDisplayProblem(); // 2. Busca e exibe o problema atual
}

// Carrega o progresso salvo nas configurações deste documento
function loadProgress() {
    return new Promise(resolve => {
        const savedId = Office.context.document.settings.get('euler_currentProblemId');
        if (savedId) {
            currentProblemId = parseInt(savedId, 10);
        }
        resolve();
    });
}

// Salva o progresso atual no arquivo
function saveProgress() {
    Office.context.document.settings.set('euler_currentProblemId', currentProblemId);
    Office.context.document.settings.saveAsync();
}

/**
 * Esta é a nova função principal. Ela busca um único problema e o exibe na tela.
 */
async function fetchAndDisplayProblem() {
    const titleElement = document.getElementById("problem-title");
    const descriptionElement = document.getElementById("problem-description");
    const loadingElement = document.getElementById("loading-message");

    titleElement.textContent = `Carregando Problema #${currentProblemId}...`;
    descriptionElement.textContent = "";
    loadingElement.textContent = "Buscando na web...";

    try {
        const url = `https://projecteuler.net/minimal=${currentProblemId}`;
        // Continuamos usando um proxy para evitar problemas de segurança do navegador (CORS)
        const proxyUrl = `https://api.allorigins.win/get?url=${encodeURIComponent(url)}`;
        
        const response = await fetch(proxyUrl);
        if (!response.ok) throw new Error("Falha na rede ao buscar o problema.");

        const data = await response.json();
        const problemText = data.contents;

        // O texto vem no formato: "ID. Título\n\nDescrição..."
        // Vamos separar o título da descrição
        const lines = problemText.split('\n');
        const problemTitle = lines[0]; // A primeira linha é sempre o título
        const problemDescription = lines.slice(2).join('\n'); // O resto (após uma linha em branco) é a descrição

        titleElement.textContent = problemTitle;
        descriptionElement.textContent = problemDescription.trim();

    } catch (error) {
        console.error("Erro ao buscar problema:", error);
        titleElement.textContent = "Erro ao Carregar";
        descriptionElement.textContent = `Não foi possível buscar o problema #${currentProblemId}. Verifique sua conexão ou tente novamente.`;
    } finally {
        loadingElement.textContent = "";
    }
}

// Abre a página HTML completa do problema no navegador
function openProblemOnWeb() {
    window.open(`https://projecteuler.net/problem=${currentProblemId}`, '_blank');
}

// Avança para o próximo nível
async function goToNextProblem() {
    currentProblemId++;
    saveProgress();
    await fetchAndDisplayProblem(); // Busca e exibe o novo problema
}
