from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER
import pandas as pd

# Função para criar e salvar o PDF para cada cliente
def criar_pdf_para_cliente(cliente, dados_cliente):
    # Definindo o tamanho da página paisagem com dimensões personalizadas
    width, height = landscape((30*inch, 10*inch))  # Tamanho personalizado
    
    # Ajustando a largura das colunas para se ajustarem ao novo tamanho da página
    num_cols = len(dados_cliente.columns)
    col_width = width / num_cols
    
    # Definindo o estilo da tabela e do texto
    style_sheet = getSampleStyleSheet()
    estilo_tabela = style_sheet['BodyText']
    estilo_tabela.alignment = TA_CENTER
    estilo_cabecalho = style_sheet['Title']
    estilo_cabecalho.alignment = TA_CENTER
    
    # Definindo o estilo da fonte e tamanho
    fonte = 'Helvetica'
    tamanho_fonte = 8  # Reduzindo o tamanho da fonte
    
    # Criando o arquivo PDF
    pdf_output_path = f'relatorio_{cliente}.pdf'
    pdf = SimpleDocTemplate(pdf_output_path, pagesize=landscape((30*inch, 10*inch)))  # Tamanho personalizado

    # Criando uma lista para armazenar os dados da tabela
    table_data = []

    # Adicionando cabeçalhos à lista
    table_data.append(dados_cliente.columns.tolist())

    # Adicionando os dados à lista
    for _, row in dados_cliente.iterrows():
        table_data.append(row.tolist())

    # Criando a tabela com os dados
    table = Table(table_data, colWidths=[col_width] * num_cols)

    # Aplicando estilos à tabela
    table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.black),
                               ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                               ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                               ('FONTNAME', (0, 0), (-1, 0), fonte),
                               ('FONTSIZE', (0, 0), (-1, -1), tamanho_fonte),
                               ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                               ('BACKGROUND', (0, 1), (-1, -1), colors.white),  # Cor de fundo branco para os dados
                               ('GRID', (0, 0), (-1, -1), 1, colors.black)]))  # Adicionando bordas

    # Adicionando a tabela ao PDF
    pdf.build([table])

# Leitura do arquivo Excel original
excel_file_path = r'C:\\Users\\mateu\\OneDrive\\Documents\\DOCUMENTOS , FAMILIA , CURRICULO\\BM Finance Group\\Alves&Bigler\\Rel de Acordos  Cobrança 12.03.2024 a 20.03.2024 - JUA CONDOMINIO.xlsx'
df = pd.read_excel(excel_file_path, header=10)

# Tratamento de dados vazios (se necessário)
df.fillna(value='', inplace=True)

# Obtendo a lista de clientes
clientes = df['CARTEIRA'].unique()

# Salvando os dados de cada cliente em um PDF separado
for cliente in clientes:
    # Filtragem dos dados por cliente
    dados_cliente = df[df['CARTEIRA'] == cliente]
    
    # Criando e salvando o PDF para o cliente atual
    criar_pdf_para_cliente(cliente, dados_cliente)
