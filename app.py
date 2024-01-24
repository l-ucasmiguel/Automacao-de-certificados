# Pegar os dados da planilha como, nome do curso, nome do participante, tipo de participação, data de inicio, data final, carga horaria, e data de emissao.
# Transferir para a imagem do certificado


# Importando a lib openpyxl e seus módulos
import openpyxl
from PIL import Image, ImageDraw, ImageFont


# Abrir a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']


for linha in sheet_alunos.iter_rows(min_row=2):
    # acessar cada célula que contém a info que precisamos
    nome_curso = linha[0].value             # Nome do curso
    nome_participante = linha[1].value      # Nome do participante
    tipo_participacao = linha[2].value      # Tipo da participação
    data_inicio = linha[3].value            # Data inicio
    data_final = linha[4].value             # Data final
    carga_horaria = linha[5].value          # Carga horária
    data_emissao = linha[6].value           # Data emissão


    # Tranferir os dados da planilha para a imagem do certificado
    font_nome = ImageFont.truetype('./tahomabd.ttf')
    font_geral = ImageFont.truetype('./tahoma.ttf')

    Image.open('./certificado_padrao.jpg')
    print('TESTE')