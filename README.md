# 📊 Power BI Auto-Documentação

Este projeto automatiza a criação de documentação técnica para relatórios do Power BI. Ele extrai metadados de um arquivo Excel (gerado por ferramentas de análise de modelo) e gera automaticamente um **Dicionário de Dados em Excel** e uma **Documentação Técnica em Word** estruturada.

## 🚀 Funcionalidades

- **Automação de Nomes**: Gera arquivos de saída baseados no nome do arquivo de entrada.
- **Dicionário de Medidas**: Extrai Nome, Tabela, Pasta de Exibição e a **Expressão DAX**.
- **Dicionário de Colunas**: Mapeia colunas físicas e calculadas, identificando tipos de dados.
- **Documentação Word Profissional**: Cria automaticamente Capa, Sumário, Seção de Objetivo e placeholders para o Modelo de Dados (Star Schema).
- **Padronização**: Ideal para projetos corporativos que exigem governança de dados.

## 🛠️ Tecnologias Utilizadas

- **Python 3.13+**
- **Pandas**: Processamento e manipulação de dados.
- **Python-docx**: Geração e formatação de documentos Word.
- **Openpyxl**: Manipulação de arquivos Excel (.xlsx).

## 📋 Pré-requisitos

Antes de iniciar, você precisará ter o Python instalado em sua máquina. Para instalar as dependências necessárias, execute:

```Bash

pip install -r requirements.txt

```
📖 Como Usar
Exporte os metadados do seu Power BI para um arquivo Excel chamado resulto.xlsx (recomendado usar ferramentas como DAX Studio, Tabular Editor ou Measure Killer).

Coloque o script main.py na mesma pasta do arquivo.

No código, ajuste os campos manuais (Nome do Projeto, Responsável, etc.).

Execute o script:

```Bash

python main.py

```
Verifique os arquivos gerados: results_Documentacao.xlsx e results_Documentacao.docx.

📂 Estrutura do Documento Word
O documento gerado segue o padrão de mercado:

Capa (Título, Responsável, Versão).

Visão Geral e Objetivo (Contexto de negócio).

Arquitetura de Dados (Espaço para o print do modelo Star Schema).

Tabelas do Modelo.

Dicionário de Medidas (KPIs) com fórmulas DAX.

Dicionário de Colunas.

👤 Autor
Projeto desenvolvido por Adriano Soares, unindo experiência em logística e análise de dados.


---