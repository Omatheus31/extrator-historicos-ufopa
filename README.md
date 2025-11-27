# 🚀 Extrator de Históricos Acadêmicos - UFOPA (atualizado)

Aplicação web para automatizar a extração de Componentes Pendentes e Resumo de Carga Horária a partir de históricos em PDF.

> Este README foi atualizado para refletir suporte a arquivos Excel modernos (.xlsx), instruções de execução no Windows (PowerShell) e dicas de depuração.

---

## O que o projeto faz

- Recebe múltiplos PDFs de históricos e um arquivo de controle (planilha de percentuais).
- Extrai componentes curriculares obrigatórios pendentes (ignora ENADE), marca disciplinas "(Matriculado)" quando identificadas e calcula um resumo de carga horária pendente.
- Gera três relatórios em `generated_reports/`:
  - `relatorio_componentes.xlsx` (Excel formatado)
  - `relatorio_final.csv` (compacto, delimitador `;`)
  - `relatorio_historicos.txt` (linhas simples)

---

## Principais melhorias nesta versão

- Suporte para arquivos de percentuais em `.xls` e `.xlsx`. O backend detecta automaticamente a extensão e usa `xlrd` para `.xls` e `openpyxl` para `.xlsx`.
- Frontend envia os campos `pdf_files` (múltiplos) e `excel_file` (um arquivo). Existe também a opção de extrair sem percentuais — marque a caixa "Extrair sem percentuais" na UI para pular o upload do Excel.
- `requirements.txt` foi atualizado para indicar `xlrd==1.2.0` (compatibilidade com arquivos `.xls`).

---

## Requisitos

- Python 3.8+ (recomendado 3.11)
- pip

Dependências (já listadas em `requirements.txt`):

- Flask
- Flask-CORS
- pdfplumber
- xlrd==1.2.0
- openpyxl

---

## Instalação e execução (PowerShell - Windows)

Abra o PowerShell e execute os passos abaixo na pasta do projeto (ou indique a pasta
onde quer que os `uploads/` e `generated_reports/` sejam criados).

1) Navegue até a pasta do repositório clonado (ou escolha a pasta base desejada):

```powershell
# Exemplo: entre na pasta onde você clonou o repositório
cd "C:\caminho\para\extrator_components_UFOPA"

# OU: você pode executar a aplicação a partir de qualquer pasta e passar a pasta base:
# python app.py --base-dir "C:\pasta\onde\serao\dados"

# OU: exporte a variável de ambiente para definir a pasta base (PowerShell):
# $env:EXTRACTION_BASE_DIR = 'C:\pasta\onde\serao\dados'
# python app.py
```

2) Crie e ative um ambiente virtual:

```powershell
python -m venv .\venv
.\venv\Scripts\Activate
```

3) Atualize pip e instale dependências:

```powershell
pip install --upgrade pip
pip install -r .\requirements.txt
```

4) Inicie o servidor Flask:

```powershell
python .\app.py
```

5) Abra no navegador: http://127.0.0.1:5000

---

## Uso da interface

1. Selecione os arquivos PDF (múltiplos) no primeiro campo.
2. (Opcional) Selecione o arquivo de percentuais (`.xls` ou `.xlsx`) no segundo campo. Se não quiser usar percentuais, marque a opção "Extrair sem percentuais".
3. Clique em "Iniciar Extração".
4. Acompanhe as mensagens na área de logs; ao final, os links para download aparecerão.

---

## Contrato da API (para integrações)

- Endpoint: `POST /upload_and_extract`
- Form data:
  - `pdf_files` — arquivos PDF (campo repetível / múltiplo)
  - `excel_file` — arquivo de percentuais (`.xls` ou `.xlsx`). Opcional se enviar `skip_percentuals`.
  - `skip_percentuals` — flag opcional (valor `1`) para indicar que a extração deve prosseguir sem arquivo de percentuais.
- Resposta JSON (success):
  ```json
  {
    "status": "success",
    "message": "Extração e geração de relatórios concluídas com sucesso!",
    "download_links": {
      "excel_report": "/download/relatorio_componentes.xlsx",
      "csv_report": "/download/relatorio_final.csv",
      "txt_report": "/download/relatorio_historicos.txt"
    }
  }
  ```

Os arquivos podem ser baixados via `GET /download/<filename>`.

---

## Estrutura de pastas geradas

- `uploads/` — arquivos enviados temporariamente (limpo a cada execução)
- `generated_reports/` — relatórios gerados (relatório Excel, CSV e TXT)

---

## Solução de problemas

- Erro ao abrir XLS:
  - Verifique se o arquivo é `.xls` e, se for, assegure que `xlrd==1.2.0` esteja instalado (já fixado em `requirements.txt`).
- Planilha com layout diferente:
  - O script atual lê dados a partir da linha 10 e usa Coluna B (matrícula) e Coluna G (percentual). Se seu layout for diferente, posso ajustar o script para corresponder ao seu arquivo.
- Upload não funciona / erro CORS: confirme que `Flask-CORS` está instalado (aplicação já habilita CORS no `app.py`).
- Tempo de processamento / arquivos grandes: aumente `app.config['MAX_CONTENT_LENGTH']` em `app.py` se necessário.
- Permissões: o servidor grava em disco (`uploads/`, `generated_reports/`); verifique permissões de escrita.

---

## Testes sugeridos

1. Teste rápido: coloque 2-3 PDFs em `pdfs/` (ou use a UI) e uma planilha `.xls` ou `.xlsx` com o formato esperado e verifique se os três relatórios são gerados.
2. Teste `.xls` e `.xlsx` para confirmar ambas as rotas de leitura funcionam.

---

## Próximos passos (opcionais)

- Adicionar logging em arquivo (INFO/DEBUG) para facilitar diagnóstico.
- Adicionar testes unitários para `carregar_percentuais()` usando amostras `.xls` e `.xlsx`.
- Tornar o caminho da planilha e offsets configuráveis via variáveis de ambiente ou UI avançada.

---

Se quiser, eu posso adaptar o carregamento dos percentuais ao layout exato do seu arquivo — envie as primeiras 10-15 linhas (CSV exportado) e eu faço a adaptação e um teste rápido.
# 🚀 Extrator de Históricos Acadêmicos - UFOPA

> Uma aplicação web simples para automatizar a extração de dados de Componentes Pendentes e Carga Horária de Históricos Escolares (PDF) da UFOPA.

Este projeto transforma um processo manual de análise de PDFs em uma aplicação web rápida e intuitiva. Usuários podem fazer o upload de múltiplos históricos em PDF, junto com um arquivo de controle (XLS), e receber relatórios consolidados em segundos.

---

## ✨ Funcionalidades

* **Interface Web:** Uma UI limpa e amigável, eliminando a necessidade de rodar scripts manualmente.
* **Upload Múltiplo:** Envie dezenas de arquivos PDF de uma só vez.
* **Upload de Controle:** Envie o arquivo `.xls` que contém os dados de percentual cumprido.
* **Extração Inteligente:** O backend lê os PDFs, identifica tabelas e textos, e extrai:
    * Componentes Curriculares Obrigatórios Pendentes (ignorando ENADE).
    * Disciplinas em que o aluno está "Matriculado".
    * Resumo de Carga Horária (Optativos, Complementares, Total Pendente).
* **Geração de Relatórios:** Cria e disponibiliza para download três arquivos:
    1.  `relatorio_componentes.xlsx` (Relatório completo formatado).
    2.  `relatorio_final.csv` (Relatório compacto em CSV).
    3.  `relatorio_historicos.txt` (Relatório simples em TXT).

---

## 🛠️ Tecnologias Utilizadas

Este projeto é dividido em duas partes principais:

* **Backend (API)**:
    * **Python 3**
    * **Flask** (Para o servidor web e API)
    * **pdfplumber** (Para extração de dados dos PDFs)
    * **xlrd** (Para leitura do arquivo `.xls` de percentuais)
    * **openpyxl** (Para a geração do relatório `.xlsx` final)

* **Frontend (UI)**:
    * **HTML5**
    * **CSS3**
    * **JavaScript (Fetch API)**

---

## ⚙️ Instalação e Configuração

Siga estes passos para rodar o projeto localmente:

1.  **Clone o repositório:**
    ```bash
    git clone [URL_DO_SEU_REPOSITÓRIO_GITHUB_AQUI]
    cd [NOME_DA_SUA_PASTA]
    ```

2.  **Crie e Ative um Ambiente Virtual** (Recomendado):
    ```bash
    # Criar
    python -m venv venv
    
    # Ativar (Windows)
    .\venv\Scripts\activate
    
    # Ativar (macOS/Linux)
    source venv/bin/activate
    ```

3.  **Crie o arquivo `requirements.txt`:**
    Crie um arquivo chamado `requirements.txt` na raiz do projeto e cole o seguinte conteúdo nele:
    ```
    Flask
    Flask-CORS
    pdfplumber
    xlrd
    openpyxl
    ```

4.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```

---

## 🏃 Como Rodar

Com tudo instalado, basta iniciar o servidor Flask:

1.  **Inicie o servidor:**
    ```bash
    python app.py
    ```

2.  **Acesse no Navegador:**
    Abra seu navegador e acesse:
    [**http://127.0.0.1:5000**](http://127.0.0.1:5000)

### Como Usar a Ferramenta

1.  **Passo 1:** Clique em "Escolher arquivos" e selecione todos os PDFs dos históricos que deseja processar.
2.  **Passo 2:** Clique em "Escolher arquivo" e selecione o arquivo `.xls` que contém os dados de percentual cumprido.
3.  **Passo 3:** Clique no botão verde "Iniciar Extração".
4.  **Passo 4:** Aguarde as mensagens de status. Quando a extração terminar, os links para download dos relatórios aparecerão abaixo.

---

## 👨‍💻 Autores

* **Backend (Lógica de Extração e API):** [Harry120705](https://github.com/Harry120705)
* **Frontend (Interface Web):** [Omatheus31](https://github.com/Omatheus31)