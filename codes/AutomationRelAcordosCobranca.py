import pandas as pd
from reportlab.lib.pagesizes import landscape, inch
from reportlab.pdfgen import canvas

# Passo 1: Leitura do Excel
df = pd.read_excel(r'C:\\Users\mateu\OneDrive\Documents\DOCUMENTOS , FAMILIA , CURRICULO\BM Finance Group\Alves&Bigler\Rel de Acordos  Cobrança 12.03.2024 a 20.03.2024 - JUA CONDOMINIO.xlsx', header=10)

# Passo 2: Tratamento de dados vazios (se necessário)
df.fillna(value='', inplace=True)

# Obter cabeçalhos do Excel
cabeçalhos_excel = list(df.columns)

# Passo 3: Iteração sobre os clientes
clientes = df['CARTEIRA'].unique()

for cliente in clientes:
    # Passo 4: Filtragem dos dados por cliente
    dados_cliente = df[df['CARTEIRA'] == cliente]

    # Passo 5: Geração de relatório em PDF
    nome_arquivo = f"relatorio_{cliente}.pdf"

    # Definindo o tamanho da página como paisagem e dimensões personalizadas
    width, height = 45*inch, 12*inch  # Largura e altura em polegadas
    pagesize = landscape((width, height))

    # Criando o arquivo PDF com as dimensões personalizadas
    c = canvas.Canvas(nome_arquivo, pagesize=pagesize)

    # Escreva os cabeçalhos no PDF
    y = 700  # Posição inicial (para ajustar conforme necessário)
    x = 50  # Posição horizontal (para ajustar conforme necessário)
    for cabeçalho in cabeçalhos_excel:
        c.drawString(x, y, cabeçalho)
        x += 180  # Espaço entre cabeçalhos (para ajustar conforme necessário)

    # Escreva os dados no PDF
    y -= 20  # Espaço entre cabeçalhos e dados (para ajustar conforme necessário)
    for index, row in dados_cliente.iterrows():
        x = 50  # Reiniciar a posição horizontal para os dados
        for coluna in cabeçalhos_excel:
            c.drawString(x, y, str(row[coluna]))
            x += 180  # Espaço entre colunas (para ajustar conforme necessário)
        y -= 20  # Espaço entre linhas (para ajustar conforme necessário)

    c.save()

    # Passo 6: Salvar o relatório
    print(f"Relatório para {cliente} gerado com sucesso: {nome_arquivo}")
