import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# Passo 1: Leitura do Excel
df = pd.read_excel(r'C:\\Users\mateu\OneDrive\Documents\DOCUMENTOS , FAMILIA , CURRICULO\BM Finance Group\Alves&Bigler\Rel de Acordos  Cobrança 12.03.2024 a 20.03.2024 - JUA CONDOMINIO.xlsx', header=10)


# ('Rel de Acordos  Cobrança 12.03.2024 a 20.03.2024 - JUA CONDOMINIO.xlsx')

# Passo 2: Tratamento de dados vazios (se necessário)
df.fillna(value='', inplace=True)

# Passo 3: Iteração sobre os clientes
clientes = df['NOME'].unique()

for cliente in clientes:
    # Passo 4: Filtragem dos dados por cliente
    dados_cliente = df[df['NOME'] == cliente]

    # Passo 5: Geração de relatório em PDF
    nome_arquivo = f"relatorio_{cliente}.pdf"
    c = canvas.Canvas(nome_arquivo, pagesize=letter)

    # Escreva os dados no PDF
    y = 750  # Posição inicial
    for index, row in dados_cliente.iterrows():
        x = 50  # Posição horizontal
        for coluna in df.columns:
            c.drawString(x, y, str(row[coluna]))
            x += 100  # Espaço entre colunas
        y -= 20  # Espaço entre linhas

    c.save()

    # Passo 6: Salvar o relatório
    print(f"Relatório para {cliente} gerado com sucesso: {nome_arquivo}")
