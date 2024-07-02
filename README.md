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

## Cenário: Loja de Eletrônicos "TechMaster"

1. Contexto:
   A TechMaster é uma pequena loja de eletrônicos que vende smartphones, laptops e acessórios. Eles têm três filiais em diferentes bairros da cidade.

2. Necessidade:
   O gerente geral precisa receber relatórios semanais sobre o desempenho de vendas de cada filial.

3. Dados disponíveis:
   Cada filial mantém uma planilha Excel com as vendas diárias, incluindo:
   - Data da venda
   - Produto vendido
   - Quantidade
   - Preço unitário
   - Filial

4. Objetivo da automação:
   Criar um sistema que leia as planilhas de cada filial, consolide os dados, gere um relatório semanal e envie por e-mail para o gerente geral.

5. Etapas do processo:
   a. Leitura de dados: Ler as planilhas de vendas das três filiais.
   b. Consolidação: Juntar os dados em um único DataFrame.
   c. Análise: Calcular métricas como total de vendas por filial, produtos mais vendidos, etc.
   d. Visualização: Criar gráficos mostrando o desempenho de cada filial e produtos.
   e. Geração de relatório: Criar um PDF ou Excel com os resultados e gráficos.
   f. Envio de e-mail: Enviar o relatório automaticamente para o gerente geral.

6. Frequência:
   O processo deve ser executado toda segunda-feira às 8h, analisando os dados da semana anterior.

7. Estrutura da planilha:
   Cada filial terá uma planilha chamada "vendas_filial_X.xlsx" (onde X é o número da filial) com as seguintes colunas:
   - Data
   - Produto
   - Categoria (Smartphone, Laptop, Acessório)
   - Quantidade
   - Preço Unitário
   - Filial

8. Métricas para o relatório:
   - Total de vendas por filial
   - Produto mais vendido em cada filial
   - Comparação de vendas entre as filiais
   - Tendência de vendas ao longo da semana
   - Top 5 produtos mais lucrativos

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

### requirements.txt

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

### Configuração

Crie um arquivo config.py na raiz do projeto com suas credenciais de e-mail:
python

```
EMAIL_REMETENTE = "seu_email@gmail.com"
EMAIL_SENHA = "sua_senha"
EMAIL_DESTINATARIO = "destinatario@exemplo.com"
```

### Uso

1. Execute o script para ler os dados, gerar o relatório e enviar o e-mail:
   ```bash
   python main.py
   ```