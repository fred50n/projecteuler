// Aguarda o ambiente do Office estar pronto
Office.onReady(info => {
    if (info.host === Office.HostType.Excel) {
        // Inicializa a aplicação
        initializeApp();
    }
});

let problemsArray = [];
let currentProblemId = 1;

async function initializeApp() {
    // Adiciona os listeners aos botões
    document.getElementById("verify-button").onclick = openProblemOnWeb;
    document.getElementById("next-problem-button").onclick = goToNextProblem;
    document.getElementById("loading").textContent = "Buscando problemas do Project Euler...";

    try {
        await loadProgress(); // Carrega o progresso salvo
        await fetchProblems(); // Busca os problemas da web
        displayCurrentProblem(); // Exibe o problema atual
        document.getElementById("loading").textContent = "";
    } catch (error) {
        console.error(error);
        document.getElementById("problem-title").textContent = "Erro!";
        document.getElementById("problem-description").textContent = "Não foi possível carregar os problemas. Verifique o console para mais detalhes.";
        document.getElementById("loading").textContent = "";
    }
}

// Carrega o progresso salvo nas configurações do documento
function loadProgress() {
    return new Promise((resolve, reject) => {
        Office.context.document.settings.refreshAsync(result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const savedId = Office.context.document.settings.get('currentProblemId');
                if (savedId) {
                    currentProblemId = parseInt(savedId, 10);
                }
                resolve();
            } else {
                reject('Erro ao carregar configurações: ' + result.error.message);
            }
        });
    });
}

// Salva o progresso atual
function saveProgress() {
    Office.context.document.settings.set('currentProblemId', currentProblemId);
    Office.context.document.settings.saveAsync(result => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
            console.error('Erro ao salvar progresso: ' + result.error.message);
        }
    });
}

// Busca os problemas do CSV
async function fetchProblems() {
    const url = "https://projecteuler.net/minimal=problems;csv";
    // Usamos um proxy para evitar problemas de CORS (bloqueio de requisição entre domínios)
    const proxyUrl = `https://api.allorigins.win/get?url=${encodeURIComponent(url)}`;
    
    const response = await fetch(proxyUrl);
    if (!response.ok) {
        throw new Error("Falha na rede ao buscar problemas.");
    }
    const data = await response.json();
    problemsArray = data.contents.split('\n').filter(line => line); // Filtra linhas vazias
}

// Exibe o problema atual na tela
function displayCurrentProblem() {
    if (currentProblemId > problemsArray.length) {
        document.getElementById("problem-title").textContent = "Parabéns!";
        document.getElementById("problem-description").textContent = "Você completou todos os problemas disponíveis.";
        return;
    }

    const problemLine = problemsArray[currentProblemId - 1];
    // O CSV é um pouco inconsistente, então limpamos bem os dados
    const problemData = problemLine.split('","').map(item => item.replace(/"/g, ''));

    const id = problemData[0];
    const title = problemData[1];
    let description = problemData.slice(2).join(',');

    // Limpa tags HTML da descrição para melhor visualização
    description = description
        .replace(/<p>/gi, '\n\n')
        .replace(/<\/p>/gi, '')
        .replace(/<br\s*\/?>/gi, '\n');

    document.getElementById("problem-title").textContent = `Problema ${id}: ${title}`;
    document.getElementById("problem-description").textContent = description.trim();
}

// Abre a página do problema no site oficial
function openProblemOnWeb() {
    window.open(`https://projecteuler.net/problem=${currentProblemId}`, '_blank');
}

// Avança para o próximo problema
function goToNextProblem() {
    currentProblemId++;
    saveProgress();
    displayCurrentProblem();
}