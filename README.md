# automacao-emails-relatorios
# Automação de E-mails e Relatórios com Python

## Descrição

Este projeto visa automatizar o envio de e-mails e a geração de relatórios com base em dados extraídos de planilhas. A automação de processos como esse pode economizar tempo e reduzir erros em tarefas repetitivas, sendo muito útil para empresas e profissionais que precisam enviar relatórios periódicos.

## Funcionalidades

- **Leitura de Dados**: Extração de dados de planilhas Excel.
- **Análise de Dados**: Processamento e análise de dados usando `pandas`.
- **Geração de Relatórios**: Criação de gráficos e relatórios em formato PDF ou Excel.
- **Envio de E-mails**: Envio automatizado de e-mails com relatórios anexados.
- **Agendamento de Tarefas**: Execução automática das tarefas em horários específicos.

## Tecnologias Utilizadas

- **Linguagem**: Python
- **Bibliotecas**: 
  - `pandas` para manipulação de dados
  - `matplotlib` e `seaborn` para criação de gráficos
  - `xlsxwriter` para geração de arquivos Excel
  - `fpdf` para criação de PDFs
  - `smtplib` para envio de e-mails
  - `schedule` para agendamento de tarefas

## Como Usar

### Pré-requisitos

- Python 3.x instalado
- Biblioteca `virtualenv` para criar um ambiente virtual

### Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/lucasneiva/automacao-emails-relatorios.git
   cd automacao-emails-relatorios
   ``` 

### Como Usar

### Pré-requisitos

- Python 3.x instalado
- Biblioteca `virtualenv` para criar um ambiente virtual

### Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/lucasneiva/automacao-emails-relatorios.git
   cd automacao-emails-relatorios
   ```

2. Crie um ambiente virtual e ative-o:
   ```bash
   python -m venv automacao_venv
   source automacao_venv/bin/activate  # No Windows use: automacao_venv\Scripts\activate
   ```

3. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```

### Exemplo de `requirements.txt`

Certifique-se de criar um arquivo `requirements.txt` na raiz do projeto com o seguinte conteúdo:

```plaintext
pandas
openpyxl
matplotlib
seaborn
xlsxwriter
fpdf
schedule
```

### Uso

1. Coloque seus dados em uma planilha Excel (`dados.xlsx`) na raiz do projeto.
2. Execute o script para ler os dados, gerar o relatório e enviar o e-mail:
   ```bash
   python main.py
   ```