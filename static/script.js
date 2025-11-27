// Obtém referências aos elementos
const messagesTextArea = document.getElementById('messages');
const extractButton = document.getElementById('extractButton');
const pdfFilesInput = document.getElementById('pdfFiles');
const excelFileInput = document.getElementById('excelFile');
const resultsArea = document.getElementById('resultsArea');
const downloadLinksDiv = document.getElementById('downloadLinks');
const progressContainer = document.getElementById('progressContainer');
const progressText = document.getElementById('progressText');
const progressPercentage = document.getElementById('progressPercentage');
const progressFill = document.getElementById('progressFill');
const skipPercentuaisCheckbox = document.getElementById('skipPercentuais');

let eventSource = null;

function updateMessages(message) {
    messagesTextArea.value += message + '\n';
    messagesTextArea.scrollTop = messagesTextArea.scrollHeight; // Scroll automático
}

function setProcessingState(isProcessing) {
    if (isProcessing) {
        extractButton.disabled = true;
        extractButton.textContent = 'Processando...';
        // Limpa resultados anteriores
        resultsArea.style.display = 'none';
        downloadLinksDiv.innerHTML = '';
        // Mostra container de progresso
        progressContainer.style.display = 'block';
        updateProgress(0, 0);
    } else {
        extractButton.disabled = false;
        extractButton.textContent = 'Iniciar Extração';
        // Esconde container de progresso
        progressContainer.style.display = 'none';
    }
}

function updateProgress(current, total) {
    if (total === 0) {
        progressText.textContent = 'Preparando...';
        progressPercentage.textContent = '0%';
        progressFill.style.width = '0%';
        return;
    }
    
    const percentage = Math.round((current / total) * 100);
    progressText.textContent = `Processando: ${current}/${total} PDFs`;
    progressPercentage.textContent = `${percentage}%`;
    progressFill.style.width = `${percentage}%`;
}

function startProgressListener() {
    if (eventSource) {
        eventSource.close();
    }
    
    eventSource = new EventSource('/progress');
    
    eventSource.onmessage = function(event) {
        const data = event.data;
        
        if (data === 'DONE') {
            eventSource.close();
            eventSource = null;
            return;
        }
        
        if (data === 'ping') {
            return;
        }
        
        // Formato esperado: "current/total"
        const match = data.match(/(\d+)\/(\d+)/);
        if (match) {
            const current = parseInt(match[1]);
            const total = parseInt(match[2]);
            updateProgress(current, total);
        }
    };
    
    eventSource.onerror = function(error) {
        console.error('Erro no EventSource:', error);
        if (eventSource) {
            eventSource.close();
            eventSource = null;
        }
    };
}

async function startExtraction() {
    const pdfFiles = pdfFilesInput.files;
    const excelFile = excelFileInput.files[0];
    const skipPercentuais = skipPercentuaisCheckbox ? skipPercentuaisCheckbox.checked : false;

    // Validação de entrada
    if (pdfFiles.length === 0) {
        updateMessages("ERRO: Por favor, selecione os arquivos PDF.");
        return;
    }
    if (!skipPercentuais && !excelFile) {
        updateMessages("ERRO: Por favor, selecione o arquivo Excel de percentuais ou marque 'Extrair sem arquivo de percentuais'.");
        return;
    }

    setProcessingState(true);
    updateMessages("Iniciando upload e extração...");
    
    // Inicia o listener de progresso
    startProgressListener();

    const formData = new FormData();
    for (let i = 0; i < pdfFiles.length; i++) {
        formData.append('pdf_files', pdfFiles[i]);
    }
    if (!skipPercentuais && excelFile) {
        formData.append('excel_file', excelFile);
    } else if (skipPercentuais) {
        // Indica explicitamente ao backend que ele deve pular percentuais
        formData.append('skip_percentuals', '1');
    }

    try {
        // Envia os arquivos para o endpoint /upload_and_extract no backend Flask
        const response = await fetch('/upload_and_extract', {
            method: 'POST',
            body: formData 
        });

        const result = await response.json();

        if (!response.ok) {
            // Se o servidor retornar um erro (ex: 400, 500)
            throw new Error(result.message || `Erro do servidor: ${response.status}`);
        }

        // Se o backend processar com sucesso (status "success")
        updateMessages("Extração concluída com sucesso!");
        updateMessages(result.message);

        // Mostra a área de resultados e cria os links de download
        if (result.download_links) {
            resultsArea.style.display = 'block';
            
            if (result.download_links.excel_report) {
                const link = document.createElement('a');
                link.href = result.download_links.excel_report;
                link.textContent = 'Baixar Relatório Excel (.xlsx)';
                link.target = '_blank'; // Abre em nova aba
                downloadLinksDiv.appendChild(link);
            }
            if (result.download_links.csv_report) {
                const link = document.createElement('a');
                link.href = result.download_links.csv_report;
                link.textContent = 'Baixar Relatório CSV (.csv)';
                link.target = '_blank';
                downloadLinksDiv.appendChild(link);
            }
        }

    } catch (error) {
        // Pega erros de rede ou erros retornados pelo backend
        updateMessages(`FALHA NA OPERAÇÃO: ${error.message}`);
        console.error("Erro completo:", error);
    } finally {
        // Fecha o EventSource se ainda estiver aberto
        if (eventSource) {
            eventSource.close();
            eventSource = null;
        }
        // Reabilita o botão, independentemente do resultado
        setProcessingState(false);
    }
}

// Inicializa a área de mensagens
document.addEventListener('DOMContentLoaded', () => {
    updateMessages("Pronto para começar. Selecione os arquivos e clique em 'Iniciar Extração'.");
    // Caso o checkbox exista, atualiza o estado do input do Excel
    if (skipPercentuaisCheckbox) {
        const toggleExcelState = () => {
            if (skipPercentuaisCheckbox.checked) {
                excelFileInput.disabled = true;
                excelFileInput.parentElement.classList.add('disabled');
            } else {
                excelFileInput.disabled = false;
                excelFileInput.parentElement.classList.remove('disabled');
            }
        };
        skipPercentuaisCheckbox.addEventListener('change', toggleExcelState);
        // Estado inicial
        toggleExcelState();
    }
});