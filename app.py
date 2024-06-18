import openpyxl
from PIL import ImageFont, Image, ImageDraw

workbook = openpyxl.load_workbook("planilha_alunos.xlsx")
planilha = workbook["Planilha1"]

for indice, linha in enumerate(planilha.iter_rows(min_row=2)):
    nome_curso = linha[0].value
    nome_aluno = linha[1].value
    participacao = linha[2].value
    carga_horaria = linha[5].value
    data_inicio = linha[3].value
    data_final = linha[4].value
    data_emissao = linha[6].value
    print(indice, nome_aluno)

    fonte_nome = ImageFont.truetype("./MATURASC.TTF", 90)
    fonte_geral = ImageFont.truetype("./MTCORSVA.TTF", 80)
    fonte_data = ImageFont.truetype("./MTCORSVA.TTF", 65)

    imagem = Image.open("./certificado_padrao.jpg")
    desenhar = ImageDraw.Draw(imagem)

    desenhar.text((1000, 827),nome_aluno,fill="black",font= fonte_nome)
    desenhar.text((1055, 965), nome_curso, fill="black", font= fonte_geral)
    desenhar.text((1420, 1080), participacao, fill="black", font= fonte_geral)
    desenhar.text((1475, 1200), str(carga_horaria), fill="black", font= fonte_geral)
    desenhar.text((740, 1770), data_inicio, fill="black", font= fonte_data)
    desenhar.text((740, 1930), data_final, fill="black", font= fonte_data)
    desenhar.text((2220, 1930), data_emissao, fill="black", font= fonte_data)


    imagem.save(f"./Certificados/{indice} {nome_aluno} certificado.png")
