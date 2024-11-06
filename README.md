# Projeto_RPA_Challenge_Invoice

Este projeto automatiza a extração de informações de faturas em PDF e armazena os dados extraídos em um banco de dados MySQL e em um arquivo Excel. A aplicação usa técnicas de RPA (Robotic Process Automation) para facilitar a extração de dados de documentos não estruturados, tornando o processo de registro de faturas mais eficiente e menos propenso a erros manuais.

## Funcionalidades
* Leitura de arquivos PDF: Extrai o número e a data da fatura da primeira página de cada arquivo PDF.
* Inserção no banco de dados: Armazena os dados extraídos em uma tabela de um banco MySQL.
* Registro em planilha Excel: Armazena os dados em um arquivo Excel para fácil consulta.
* Tratamento de erros: Caso um PDF não tenha as informações esperadas, registra o erro no banco de dados e na planilha Excel.
  
## Tecnologias Utilizadas
* Python: Linguagem principal do projeto.
* pdfplumber: Biblioteca para leitura e extração de texto de arquivos PDF.
* openpyxl: Biblioteca para manipulação de arquivos Excel.
* MySQL Connector: Biblioteca para conectar e executar operações em um banco de dados MySQL.
* Regex (re): Para busca e extração de padrões específicos, como o número da fatura e a data.
  
## Pré-requisitos
* Python 3.x
* Bibliotecas Python: pdfplumber, openpyxl, mysql-connector-python
* MySQL (instalado e configurado)
  
## Para instalar as bibliotecas Python necessárias, execute:
![image](https://github.com/user-attachments/assets/dbbbf22e-04cc-4cbc-9aee-21d571a194b7)
## Estrutura do Código
* Função execute_insert: Realiza a inserção dos dados no banco de dados MySQL.
* Função main: Contém o fluxo principal de processamento:
  * Conecta ao banco de dados.
  * Lê todos os arquivos PDF de um diretório especificado.
  * Extrai o número da fatura e a data da primeira página de cada PDF.
  * Salva as informações no banco de dados e no arquivo Excel.
  * Salva o Excel ao final do processo, usando a data e hora no nome do arquivo.
    
## Tabela no Banco de Dados
A tabela invoice_records deve ter a seguinte estrutura:
![image](https://github.com/user-attachments/assets/30831543-8081-4c32-ae9f-f5f5196c4540)

## Como Executar
### 1. Configurar o Banco de Dados:
* Certifique-se de que o banco de dados MySQL está ativo e contém a tabela invoice_records conforme a estrutura acima.
  
### 2. Organizar Arquivos PDF:
* Coloque todos os arquivos PDF de fatura em uma pasta chamada pdf_invoices, localizada no mesmo diretório do script.
  
### 3.Executar o Script:

* Execute o script usando o comando:
  
![image](https://github.com/user-attachments/assets/d527616a-005c-4cbc-86c7-98aa3a26aab7)

### 4.Saída:

* O script gerará um arquivo Excel com os dados das faturas e registrará cada fatura no banco de dados.
* Caso ocorra algum erro (como dados ausentes no PDF), ele será registrado no Excel e no banco de dados.
  
## Exemplo de Uso
Este projeto é útil para empresas que processam faturas recebidas em formato PDF e precisam extrair informações essenciais para armazenamento e consulta posterior. Automatizando essa tarefa com RPA, o processo fica mais rápido e confiável, reduzindo o trabalho manual.

## Próximos Passos
* Adicionar suporte para múltiplas páginas em PDFs.
* Incluir funcionalidades para processar outros tipos de documentos.
* Desenvolver uma interface gráfica para usuários menos técnicos.
