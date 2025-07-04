# Automação de Apostilamento SEI

Este projeto contém uma aplicação Python completa com interface gráfica para automatizar o processo de apostilamento de aposentadoria no sistema SEI (Sistema Eletrônico de Informações) do Governo de Goiás, interagindo também com o sistema RHnet.

A aplicação foi projetada para ser portátil e fácil de distribuir, gerenciando suas próprias dependências e arquivos de dados.

## Visão Geral

O objetivo principal desta automação é reduzir drasticamente o tempo e o esforço manual necessários para processar documentos de aposentadoria. A aplicação executa uma sequência de tarefas repetitivas de forma automática, desde a busca de informações até a geração e o envio de documentos, tudo gerenciado por uma interface gráfica intuitiva para o usuário.

## Funcionalidades Principais

-   **Interface Gráfica (GUI):**
    -   Uma tela de login segura para inserir as credenciais dos sistemas SEI e RHnet, evitando que fiquem expostas no código.
    -   Um painel de controle principal que permite **iniciar, pausar, retomar e parar** a automação a qualquer momento.
    -   Um checklist em tempo real que exibe o progresso de cada etapa para o processo atual (Edital, Ficha Financeira, Apostila, etc.).
    -   Um contador de processos analisados e um log detalhado que mostra todas as ações do robô em tempo real.
-   **Navegação Inteligente:**
    -   O robô navega pelos processos no SEI, identifica os que estão marcados para "APOSTILAMENTO" e os processa sequencialmente.
    -   Acessa o sistema RHnet para buscar a ficha financeira e dados cadastrais do servidor.
    -   Lê e interpreta informações de documentos `Despacho`, `Portaria` e `Diário Oficial` para extrair dados cruciais para o apostilamento.
-   **Manipulação de Arquivos e Documentos:**

    -   **Arquivos Temporários:** Todas as descargas de arquivos (Diário Oficial, Ficha Financeira) são feitas em diretórios temporários que são automaticamente criados e destruídos, garantindo que o sistema não acumule lixo.

    -   **Ficha Financeira:** Baixa as páginas da ficha financeira do RHnet, as mescla em um único arquivo PDF e faz o upload para o processo no SEI.

    -   **Edital:** Com base no ano de ingresso do servidor, localiza os editais correspondentes (CAPA e/ou LISTA) a partir de uma base de dados local e os anexa ao processo.

    -   **Apostila e Despacho:** Gera novos documentos dentro do SEI, preenchendo-os dinamicamente com as informações coletadas.

-   **Gerenciamento de Fluxo de Trabalho e Resiliência:**

    -   **Gerenciamento Automático do ChromeDriver:** A aplicação verifica a versão do Google Chrome instalado no computador do usuário e baixa/atualiza o ChromeDriver correspondente automaticamente.

    -   **Arquivos de Log Persistentes:** Salva o histórico de processos bem-sucedidos e com falha em arquivos (`successful_processes.txt`, `failed_processes.txt`) na mesma pasta do executável, permitindo o acompanhamento e evitando reprocessamento.

    -   **Lógica de Retentativas:** Implementa esperas explícitas (WebDriverWait) e lógicas de retentativa para lidar com a latência da rede e o carregamento dinâmico das páginas, tornando a automação mais estável.

## Tecnologias Utilizadas

-   **Python 3.11+**
-   **Tkinter:** Para a construção da interface gráfica do usuário.
-   **Selenium:** Para automação de navegador web.
-   **webdriver-manager:** Para o gerenciamento automático do ChromeDriver.
-   **PyMuPDF (fitz):** Para manipulação e extração de texto de arquivos PDF.
-   **Regex (Módulo `re`):** Para extração de padrões de texto complexos.
-   **PyInstaller:** Para empacotar a aplicação em um executável distribuível.

## Como Usar a Aplicação

Para executar o programa, não é necessário instalar Python ou qualquer dependência, basta ter o **Google Chrome** instalado.

1.  **Baixe e Descompacte:** Faça o download do arquivo `.zip` da aplicação e extraia seu conteúdo para uma pasta de sua preferência (por exemplo, na sua Área de Trabalho).

2.  **Estrutura de Arquivos:**
    -   O ponto de entrada principal da aplicação é o arquivo app.py. Para iniciar, execute o seguinte comando no terminal:
        ```bash
        Pasta_da_Aplicacao/
        |
        |-- Apostilamento.exe       <-- O programa principal
        |
        +-- DIARIOS_E_DITAIS/       <-- A base de dados de Editais (essencial)
        ```
Importante: O arquivo `Apostilamento.exe` e a pasta `DIARIOS_E_DITAIS` devem sempre permanecer juntos no mesmo diretório.

3.  **Execute:** Dê um duplo clique no arquivo Apostilamento.exe para iniciar.

4.  **Primeira Execução:** Na primeira vez que você rodar o programa, ele poderá levar alguns segundos a mais para iniciar, pois fará o download do ChromeDriver compatível com a sua versão do Google Chrome. Isso só acontece uma vez.

5.  **Login:** A tela de login aparecerá. Insira suas credenciais e clique em "Acessar Automação".

6.  **Painel Principal:** O painel de controle será exibido. Clique em "Start" para iniciar o processo de automação.

## Estrutura do Projeto

-   `app.py`: **Ponto de entrada da aplicação.**  Contém a interface gráfica (GUI) e gerencia o ciclo de vida da automação.
-   `utils.py`: Módulo de utilidades que centraliza funções compartilhadas, como a criação de sessões do WebDriver e o gerenciamento de arquivos de log.
-   `Apostilamento.py`: Script principal que orquestra todo o fluxo de trabalho da automação no SEI.
-   `RHnet.py`: Módulo responsável pela automação no sistema RHnet.
-   `Edital.py`: Módulo para a criação e upload dos documentos de Edital.
-   `Apostila.py`: Módulo para a criação do documento Apostila.
-   `Despacho.py`: Módulo para a criação do documento Despacho.
-   `Ficha_Financeira.py`: Módulo para mesclar e fazer upload da Ficha Financeira.
