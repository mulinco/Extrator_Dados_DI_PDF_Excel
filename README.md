# Extrator de Dados de Declarações de Importação (D.I.) para Excel

## 📄 Descrição do Projeto

Este projeto Python automatiza a extração de dados de **Declarações de Importação (D.I.)** em formato PDF para o Excel. Ele se conecta ao Google Drive para acessar os documentos, utiliza parsing de PDF e Expressões Regulares (Regex) para extrair informações-chave como 'Nossa Referência', 'Fatura Comercial (INVOICE)', 'HAWB' e o número da 'Declaração (D.I.)'. Os dados são estruturados e exportados para uma planilha Excel, otimizando o controle financeiro e fiscal de importações.

## 🎯 Problema Resolvido

A entrada manual de dados de Declarações de Importação em planilhas é uma tarefa repetitiva, demorada e suscetível a erros. Este projeto visa eliminar essa ineficiência, transformando um processo que poderia levar horas em segundos, aumentando a precisão e liberando recursos para análises mais estratégicas.

## ✨ Funcionalidades

* **Automação de Download:** Conecta-se à Google Drive API para listar e baixar automaticamente os PDFs das D.I. de uma pasta específica.
* **Extração Inteligente de Dados:** Utiliza Expressões Regulares (Regex) para identificar e extrair precisamente os seguintes campos de cada D.I. (especialmente da Página 1 e Página 2):
    * **D.I. (Número da Declaração)**
    * **Nome do Processo (Nossa Referência)**
    * **INVOICE (Número da Fatura Comercial)**
    * **HAWB (House Air Waybill)**
* **Consolidação para Excel:** Organiza todos os dados extraídos em um DataFrame do Pandas e exporta para uma planilha Excel (`.xlsx`) consolidada.
* **Reautenticação Automática:** Gerencia tokens de autenticação (`refresh_token`) para acesso contínuo ao Google Drive sem intervenção manual após a primeira autenticação.

## 🛠️ Tecnologias Utilizadas

* **Python 3.x**
* **`PyDrive2`**: Para integração com a Google Drive API.
* **`PyMuPDF` (Fitz)**: Para leitura e extração de texto de arquivos PDF.
* **`re` (Regular Expressions)**: Para a definição e aplicação de padrões de busca de dados no texto dos PDFs.
* **`Pandas`**: Para manipulação, organização e exportação de dados para Excel.
* **Git & GitHub**: Para controle de versão e hospedagem do projeto.

## 🚀 Como Rodar o Projeto

Siga os passos abaixo para configurar e executar o script em seu ambiente.

### Pré-requisitos

* Python 3.x instalado.
* `pip` (gerenciador de pacotes do Python).
* Acesso à Internet para download de bibliotecas e conexão com Google Drive.

### Configuração de Ambiente

1.  **Clone o Repositório:**
    ```bash
    git clone [https://github.com/mulinco/Extrator_Dados_DI_PDF_Excel.git](https://github.com/mulinco/Extrator_Dados_DI_PDF_Excel.git)
    cd Extrator_Dados_DI_PDF_Excel
    ```

2.  **Crie e Ative um Ambiente Virtual (Recomendado):**
    ```bash
    python -m venv venv
    # No Windows:
    .\venv\Scripts\activate
    # No macOS/Linux:
    source venv/bin/activate
    ```

3.  **Instale as Dependências:**
    ```bash
    pip install pandas openpyxl pydrive2 pymupdf
    ```

### Configuração da Google Drive API

1.  **Acesse o Google Cloud Console:** Vá para [console.cloud.google.com](https://console.cloud.google.com/).
2.  **Crie um Projeto** (ou selecione um existente).
3.  **Habilite a Google Drive API:** No menu lateral, vá em "APIs e Serviços" > "Biblioteca" e busque por "Google Drive API" para habilitá-la.
4.  **Crie Credenciais OAuth 2.0:**
    * No menu lateral, vá em "APIs e Serviços" > "Credenciais".
    * Clique em "Crie credenciais" > "ID do cliente OAuth".
    * Selecione "Aplicativo de computador".
    * Baixe o arquivo JSON gerado (ex: `client_secret_SEUID.apps.googleusercontent.com.json`).
    * **Renomeie este arquivo para `client_secrets.json`** e coloque-o na pasta raiz do seu projeto (`Extrator_Dados_DI_PDF_Excel/`).
5.  **Configure a Tela de Consentimento OAuth:**
    * No menu lateral, vá em "APIs e Serviços" > "Tela de Consentimento OAuth".
    * Defina o "Tipo de Usuário" como "Externo".
    * Preencha o "Nome do aplicativo", "E-mail de suporte ao usuário" e "Informações de contato do desenvolvedor".
    * Na seção "Escopos", adicione o escopo `https://www.googleapis.com/auth/drive.readonly` (ou `drive` se precisar de mais permissões).
    * Na seção "Usuários de Teste", **adicione o seu próprio e-mail do Google** (o mesmo que você usará para autenticar).

### Executando o Script

1.  **Localize a pasta das D.I. no Google Drive:** No seu Google Drive, encontre a pasta que contém os PDFs das Declarações de Importação. O ID da pasta está na URL (ex: `https://drive.google.com/drive/folders/SEU_ID_DA_PASTA`).
2.  **Atualize o Código:** Abra o arquivo principal do seu projeto (`automatepdf.py` ou o nome que você deu) e **substitua o `folder_id`** pela ID da sua pasta do Google Drive.
    ```python
    folder_id = 'SEU_ID_DA_PASTA_DO_DRIVE_AQUI' # Substitua pelo ID real
    ```
3.  **Execute o Script:**
    No terminal, dentro da pasta do projeto:
    ```bash
    python seu_script.py # Substitua 'seu_script.py' pelo nome do seu arquivo principal
    ```
    * Na primeira execução, uma aba do navegador será aberta para que você autorize o acesso do script à sua conta Google Drive. Prossiga com a autenticação.
    * Um arquivo `mycreds.txt` será gerado na pasta do projeto para futuras autenticações automáticas.

## 🔒 Segurança e Privacidade

* **NUNCA adicione `mycreds.txt` ou `client_secrets.json` ao seu repositório Git.** Esses arquivos contêm suas credenciais e devem ser mantidos privados.
* **Não inclua PDFs com dados reais/sensíveis** diretamente no repositório. Para demonstrações, utilize PDFs com dados fictícios ou anonimizados.
* O arquivo `.gitignore` (presente neste repositório) já está configurado para ignorar esses arquivos e a pasta temporária de PDFs (`temp_declaracoes_importacao/`).

## ✍️ Autora

* **[Maria Rodrigues](https://www.linkedin.com/in/mariaclararodrigues3113/)**
* [LinkedIn Profile](https://www.linkedin.com/in/mariaclararodrigues3113/)
* [GitHub Profile](https://github.com/mulinco)

---
