import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime, timedelta
import schedule
import time
from config import EMAIL_REMETENTE, EMAIL_SENHA, EMAIL_DESTINATARIO

def ler_dados():
    """Lê os dados das três planilhas e retorna um DataFrame combinado."""
    df1 = pd.read_excel('vendas_filiais/vendas_filial_1.xlsx')
    df2 = pd.read_excel('vendas_filiais/vendas_filial_2.xlsx')
    df3 = pd.read_excel('vendas_filiais/vendas_filial_3.xlsx')
    df_combinado = pd.concat([df1, df2, df3], ignore_index=True)
    df_combinado['Data'] = pd.to_datetime(df_combinado['Data'])
    return df_combinado

def analisar_dados(df):
    """Realiza análises nos dados."""
    # Total de vendas por filial
    vendas_por_filial = df.groupby('Filial')['Quantidade'].sum().sort_values(ascending=False)
    
    # Produto mais vendido em cada filial
    produto_mais_vendido = df.groupby(['Filial', 'Produto'])['Quantidade'].sum().groupby(level=0).idxmax()
    
    # Receita total por categoria
    df['Receita'] = df['Quantidade'] * df['Preço Unitário']
    receita_por_categoria = df.groupby('Categoria')['Receita'].sum().sort_values(ascending=False)
    
    # Tendência de vendas ao longo da semana
    vendas_diarias = df.groupby('Data')['Quantidade'].sum()
    
    # Top 5 produtos mais lucrativos
    produtos_lucrativos = df.groupby('Produto')['Receita'].sum().sort_values(ascending=False).head()
    
    return vendas_por_filial, produto_mais_vendido, receita_por_categoria, vendas_diarias, produtos_lucrativos

def criar_graficos(vendas_por_filial, receita_por_categoria, vendas_diarias, produtos_lucrativos):
    """Cria gráficos baseados nas análises."""
    # Gráfico de barras para vendas por filial
    plt.figure(figsize=(10, 6))
    vendas_por_filial.plot(kind='bar')
    plt.title('Vendas Totais por Filial')
    plt.tight_layout()
    plt.savefig('vendas_por_filial.png')
    plt.close()

    # Gráfico de pizza para receita por categoria
    plt.figure(figsize=(8, 8))
    plt.pie(receita_por_categoria, labels=receita_por_categoria.index, autopct='%1.1f%%')
    plt.title('Distribuição de Receita por Categoria')
    plt.savefig('receita_por_categoria.png')
    plt.close()

    # Gráfico de linha para tendência de vendas
    plt.figure(figsize=(12, 6))
    vendas_diarias.plot(kind='line')
    plt.title('Tendência de Vendas ao Longo da Semana')
    plt.tight_layout()
    plt.savefig('tendencia_vendas.png')
    plt.close()

    # Gráfico de barras horizontais para produtos mais lucrativos
    plt.figure(figsize=(10, 6))
    produtos_lucrativos.plot(kind='barh')
    plt.title('Top 5 Produtos Mais Lucrativos')
    plt.tight_layout()
    plt.savefig('produtos_lucrativos.png')
    plt.close()

def gerar_relatorio_pdf(vendas_por_filial, produto_mais_vendido, receita_por_categoria, vendas_diarias, produtos_lucrativos):
    """Gera um relatório PDF com os resultados das análises e gráficos."""
    class PDF(FPDF):
        def header(self):
            self.set_font('Arial', 'B', 12)
            self.cell(0, 10, 'Relatório Semanal de Vendas - TechMaster', 0, 1, 'C')
        
        def footer(self):
            self.set_y(-15)
            self.set_font('Arial', 'I', 8)
            self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

    pdf = PDF()
    pdf.add_page()
    
    pdf.set_font('Arial', '', 12)
    pdf.cell(0, 10, f'Data do relatório: {datetime.now().strftime("%d/%m/%Y")}', 0, 1)
    pdf.ln(10)

    # Adicionar resultados e gráficos ao PDF
    pdf.cell(0, 10, 'Vendas Totais por Filial', 0, 1)
    pdf.image('vendas_por_filial.png', x=10, w=190)
    pdf.ln(10)

    pdf.cell(0, 10, 'Distribuição de Receita por Categoria', 0, 1)
    pdf.image('receita_por_categoria.png', x=10, w=190)
    pdf.ln(10)

    pdf.cell(0, 10, 'Tendência de Vendas ao Longo da Semana', 0, 1)
    pdf.image('tendencia_vendas.png', x=10, w=190)
    pdf.ln(10)

    pdf.cell(0, 10, 'Top 5 Produtos Mais Lucrativos', 0, 1)
    pdf.image('produtos_lucrativos.png', x=10, w=190)

    pdf.output('relatorio_vendas.pdf')

def enviar_email(destinatario):
    """Envia o relatório por e-mail."""
    remetente = EMAIL_REMETENTE
    senha = EMAIL_SENHA

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = destinatario
    msg['Subject'] = "Relatório Semanal de Vendas - TechMaster"

    corpo = "Segue em anexo o relatório semanal de vendas."
    msg.attach(MIMEText(corpo, 'plain'))

    filename = "relatorio_vendas.pdf"
    attachment = open(filename, "rb")

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {filename}")

    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(remetente, senha)
    text = msg.as_string()
    server.sendmail(remetente, destinatario, text)
    server.quit()

def executar_relatorio():
    """Função principal que executa todo o processo de geração e envio do relatório."""
    df = ler_dados()
    vendas_por_filial, produto_mais_vendido, receita_por_categoria, vendas_diarias, produtos_lucrativos = analisar_dados(df)
    criar_graficos(vendas_por_filial, receita_por_categoria, vendas_diarias, produtos_lucrativos)
    gerar_relatorio_pdf(vendas_por_filial, produto_mais_vendido, receita_por_categoria, vendas_diarias, produtos_lucrativos)
    destinatario = EMAIL_DESTINATARIO
    enviar_email(EMAIL_DESTINATARIO)
    print("Relatório gerado e enviado com sucesso!")

# Agendar a execução
schedule.every().monday.at("08:00").do(executar_relatorio)


if __name__ == "__main__":
#    while True:
#        schedule.run_pending()
#        time.sleep(60)  # Verifica a cada minuto
    executar_relatorio();