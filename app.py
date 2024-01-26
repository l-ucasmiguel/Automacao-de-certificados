# Precisamos pegar os dados na planilha e transferir para a imagem do certificado
# Necessários lib pillow e openpyxl


# Importando a lib openpyxl e Pillow (PIL) e seus módulos
import openpyxl, os
from PIL import Image, ImageDraw, ImageFont


# Definir o nome da pasta onde os certificados serão armazenados
pasta_certificados = 'certificados'


# Verificar se a pasta "certificados" existe, se não, criar
if not os.path.exists(pasta_certificados):
    os.makedirs(pasta_certificados)


# Carregar a planilha de alunos
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']


# Iterar sobre as linhas da planilha a partir da segunda linha (min_row=2)
for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)):
    # Extrair dados de cada coluna da linha
    nome_curso = linha[0].value             # Nome do curso
    nome_participante = linha[1].value      # Nome do participante
    tipo_participacao = linha[2].value      # Tipo da participação
    data_inicio = linha[3].value            # Data inicio
    data_final = linha[4].value             # Data final
    carga_horaria = linha[5].value          # Carga horária
    data_emissao = linha[6].value           # Data emissão


    # Carregar diferentes fontes para uso posterior
    fonte_nome = ImageFont.truetype('./tahomabd.ttf',90)
    fonte_geral = ImageFont.truetype('./tahoma.ttf',80)
    fonte_data = ImageFont.truetype('./tahoma.ttf',55)


    # Abrir uma imagem de certificado padrão
    image = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(image)


    # Adicionar texto à imagem nas posições desejadas
    desenhar.text((1020,825), nome_participante,fill='black', font=fonte_nome)
    desenhar.text((1070,953), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((1437,1066), tipo_participacao, fill='black', font=fonte_geral)
    desenhar.text((1480,1182), carga_horaria, fill='black', font=fonte_geral)
    desenhar.text((750,1770), data_inicio, fill='blue', font=fonte_data)
    desenhar.text((750,1930), data_final, fill='blue', font=fonte_data)
    desenhar.text((2220,1930), data_emissao, fill='blue', font=fonte_data)


    # Construir o caminho para o certificado na pasta "certificados"
    caminho_certificado = os.path.join(pasta_certificados, f'{indice} - {nome_participante} - Certificado.png')

    # Salvar a imagem no caminho especificado
    image.save(caminho_certificado)