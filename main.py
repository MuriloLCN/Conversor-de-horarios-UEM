import fitz
import xlsxwriter

materias = {}


def sobrepor_nomes(texto_original: str) -> str:
    global materias

    for item in materias.keys():
        if str(item) in texto_original:
            t = texto_original.replace(str(item), materias[item])
            return t

    return texto_original


def editar_tabela(textinho: str):
    workbook = xlsxwriter.Workbook('Horario.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.set_column('A:G', 20)
    worksheet.default_row_height = 35

    worksheet.center_vertically()
    worksheet.center_horizontally()

    cor_preta = workbook.add_format()
    cor_preta.set_bg_color('black')
    cor_preta.set_bold()
    cor_preta.set_align('center')
    cor_preta.set_align('vcenter')
    cor_preta.set_color('white')

    alinhado = workbook.add_format()
    alinhado.set_align('center')
    alinhado.set_align('vcenter')
    alinhado.set_text_wrap()
    alinhado.set_font_size(8)

    worksheet.set_row(7, 5)
    worksheet.set_row(20, 5)

    worksheet.set_row(1, 20)
    worksheet.set_row(14, 20)

    worksheet.write('A1', 'Horário de aulas', alinhado)
    worksheet.write('B1', '1º Semestre', alinhado)

    worksheet.write('A2', 'Hora aula', cor_preta)
    worksheet.write('B2', 'Segunda', cor_preta)
    worksheet.write('C2', 'Terça', cor_preta)
    worksheet.write('D2', 'Quarta', cor_preta)
    worksheet.write('E2', 'Quinta', cor_preta)
    worksheet.write('F2', 'Sexta', cor_preta)
    worksheet.write('G2', 'Sábado', cor_preta)

    worksheet.write('A3', '7:45 - 8:35', cor_preta)
    worksheet.write('A4', '8:35 - 9:25', cor_preta)
    worksheet.write('A5', '9:40 - 10:30', cor_preta)
    worksheet.write('A6', '10:30 - 11:20', cor_preta)
    worksheet.write('A7', '11:20 - 12:10', cor_preta)

    worksheet.write('A8', '---', cor_preta)

    worksheet.write('A9', '13:30 - 14:20', cor_preta)
    worksheet.write('A10', '14:20 - 15:10', cor_preta)
    worksheet.write('A11', '15:20 - 16:10', cor_preta)
    worksheet.write('A12', '16:10 - 17:00', cor_preta)
    worksheet.write('A13', '17:00 - 17:50', cor_preta)

    worksheet.write('A14', 'Horário de aulas', alinhado)
    worksheet.write('B14', '2º Semestre', alinhado)

    worksheet.write('A15', 'Hora aula', cor_preta)
    worksheet.write('B15', 'Segunda', cor_preta)
    worksheet.write('C15', 'Terça', cor_preta)
    worksheet.write('D15', 'Quarta', cor_preta)
    worksheet.write('E15', 'Quinta', cor_preta)
    worksheet.write('F15', 'Sexta', cor_preta)
    worksheet.write('G15', 'Sábado', cor_preta)

    worksheet.write('A16', '7:45 - 8:35', cor_preta)
    worksheet.write('A17', '8:35 - 9:25', cor_preta)
    worksheet.write('A18', '9:40 - 10:30', cor_preta)
    worksheet.write('A19', '10:30 - 11:20', cor_preta)
    worksheet.write('A20', '11:20 - 12:10', cor_preta)

    worksheet.write('A21', '---', cor_preta)

    worksheet.write('A22', '13:30 - 14:20', cor_preta)
    worksheet.write('A23', '14:20 - 15:10', cor_preta)
    worksheet.write('A24', '15:20 - 16:10', cor_preta)
    worksheet.write('A25', '16:10 - 17:00', cor_preta)
    worksheet.write('A26', '17:00 - 17:50', cor_preta)

    partes = textinho.split('*')

    colunas = ['A', 'B', 'C', 'D', 'E', 'F', 'G']

    # Manhã 1 semestre
    for i in range(5):
        i += 1
        parte_linha = partes[(i+1) * 16]

        segmento = parte_linha.split('|')

        for j in range(6):
            seg = juntar_partes(segmento, j + 3)
            seg = sobrepor_nomes(seg)
            worksheet.write(f'{colunas[j+1]}{i+2}', seg, alinhado)

    # Tarde 1 semestre
    for i in range(5):
        i += 1
        parte_linha = partes[(i+1) * 16]

        segmento = parte_linha.split('|')

        for j in range(6):
            seg = juntar_partes(segmento, j + 12)
            seg = sobrepor_nomes(seg)
            worksheet.write(f'{colunas[j + 1]}{i + 8}', seg, alinhado)

    # Manhã 2 semestre
    for i in range(5):
        i += 1
        parte_linha = partes[(i + 8) * 16 + 4]

        segmento = parte_linha.split('|')

        # Nos horários de quem entrou por cotas aparecem 20 asteriscos a mais
        if segmento == ['---------']:
            parte_linha = partes[(i + 8) * 16 + 24]
            segmento = parte_linha.split('|')

        for j in range(6):
            seg = juntar_partes(segmento, j + 3)
            seg = sobrepor_nomes(seg)
            worksheet.write(f'{colunas[j + 1]}{i + 15}', seg, alinhado)

    # Tarde 2 semestre
    for i in range(5):
        i += 1
        parte_linha = partes[(i + 8) * 16 + 4]

        segmento = parte_linha.split('|')

        if segmento == ['---------']:
            parte_linha = partes[(i + 8) * 16 + 24]
            segmento = parte_linha.split('|')

        for j in range(6):
            seg = juntar_partes(segmento, j + 12)
            seg = sobrepor_nomes(seg)
            worksheet.write(f'{colunas[j + 1]}{i + 21}', seg, alinhado)

    workbook.close()


def pegar_texto() -> str:
    with fitz.Document("horario.pdf") as doc:
        text = ""
        for page in doc:
            text += page.get_text()

    return text


def pegar_materias(textinho: str):
    global materias

    texto_cortado = textinho.split('*')
    # Tem 112 '*' antes da parte onde ficam as matérias
    parte_materias = texto_cortado[112]
    corte = parte_materias.split('\n')
    # Começando no index 4 pra tirar o cabeçalho
    for materia in corte[4:]:
        materia = " ".join(materia.split())
        pts = materia.split(' ')
        materia = " "
        # Até o 5o último index pra tirar os detalhes
        for parte in pts[:-5]:
            materia += parte + ' '
        if materia == ' ':
            continue

        try:
            txt = materia.split(' ')
            id_materia = int(txt[1])
            nome_materia = ''
            # Pulando o número e turma
            for palavra in txt[3:]:
                nome_materia += palavra + ' '
            nova_materia = {id_materia: " ".join(nome_materia.split())}
            materias.update(nova_materia)
        except (ValueError, IndexError):
            continue


def juntar_partes(arr, index):
    try:
        return arr[index].replace('-', ' - ') + ' ' + arr[index + 18]
    except IndexError:
        return ' '


try:
    texto = pegar_texto()
    pegar_materias(texto)
    editar_tabela(texto)
except fitz.FileNotFoundError:
    print("Não há nenhum arquivo com o nome 'horario.pdf' para ser lido")
    exit()
