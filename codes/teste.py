import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Passo 1: Leitura do Excel
df = pd.read_excel(r'C:\\Users\mateu\OneDrive\Documents\DOCUMENTOS , FAMILIA , CURRICULO\BM Finance Group\Alves&Bigler\Rel de Acordos  Cobrança 12.03.2024 a 20.03.2024 - JUA CONDOMINIO.xlsx')

# Passo 2: Tratamento de dados vazios (se necessário)
df.fillna(value='', inplace=True)

# Passo 3: Geração de relatório em PDF para todos os clientes de uma vez
nome_arquivo = "relatorio_todos_clientes.pdf"
c = canvas.Canvas(nome_arquivo, pagesize=letter)

y = 750  # Posição inicial
for index, row in df.iterrows():
    x = 50  # Posição horizontal
    for coluna in df.columns:
        c.drawString(x, y, str(row[coluna]))
        x += 100  # Espaço entre colunas
    y -= 20  # Espaço entre linhas

c.save()

print(f"Relatório para todos os clientes gerado com sucesso: {nome_arquivo}")
