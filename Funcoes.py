import pdfplumber
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from openpyxl import load_workbook


# Ler os dados da planilha e retornar um dicionario
def ler_dados_excel(arquivo):
    dados = {}
    excel = load_workbook(arquivo)
    prancha = excel.worksheets[0]
    elementos = prancha["A"]
    chave1 = None
    lista_chaves = []
    for n, celula in enumerate(elementos):
        if celula.value is None:
            chave1 = None
            lista_chaves = []
        else:
            if chave1 is None:
                chave1 = celula.value
                linha_chaves = prancha[n + 1]
                for cel1 in linha_chaves[1:]:
                    chave = cel1.value
                    if chave is not None:
                        lista_chaves.append(chave)
                dados[chave1] = {}
            else:
                variavel = elementos[n].value
                if len(lista_chaves) <= 1:
                    dados[chave1][variavel] = prancha[f"B{n + 1}"].value
                else:
                    dic = {}
                    linha_valores = prancha[n + 1][1:]
                    for i, chave3 in enumerate(lista_chaves):
                        valor = linha_valores[i].value
                        if valor is not None:
                            dic[chave3] = valor
                    dados[chave1][variavel] = dic
    return dados


# Converte valores de mm para points
def mm_para_pt(mm):
    return mm * 2.83465


# Converte valores de points para mm
def pt_para_mm(pt):
    return pt / 2.83465


# Criar lista de valores
def criar_lista(dic):
    P_min = int(dic['VALOR MINIMO'])
    P_max = int(dic['VALOR MAXIMO'])
    numero_amostras = int(dic['NUMERO DE AMOSTRAS'])
    amplitude = P_max - P_min
    intervalo = round(amplitude / (numero_amostras - 1), 1)
    lista = []
    for i in range(numero_amostras):
        valor = P_min + (intervalo * i)
        lista.append(valor)
    return lista


# Criar lista de cores RGB primarias
def gerar_cores_rgb(qnt):
    r = int(qnt / 6)
    x = [0.95]
    for i in range(r):
        x.append(x[-1] * (1 / 2))
    y = [0.05]
    lista1 = [[y[0], y[0], y[0]]]
    while len(lista1) <= (qnt + 1):
        for v in lista1:
            list2 = []
            for i, j in enumerate(v):
                d = v.copy()
                if j in y:
                    d[i] = x[0]
                    list2.append(d)

            for j in list2:
                if j not in lista1:
                    lista1.append(j)
            for j in lista1:
                if j[0] == j[1] and j[0] == j[2]:
                    lista1.remove(j)
        y.append(x.pop(0))

    while len(lista1) == qnt:
        lista1.remove(-1)
    for k in lista1:
        k[0] = 1 - k[0]
        k[1] = 1 - k[1]
        k[2] = 1 - k[2]
        k.reverse()
    return lista1


# Criar lista de subtons de uma cor RBG
def gradiente_cores_cmyk(cor0, qnt, mini, maxi):
    r = qnt
    coresg = []
    int = maxi - mini
    for i in range(r, -r, -2):
        n = (1 - (((i * i) ** 0.5) - i) / (2 * r))
        p = (((i * i) ** 0.5) + i) / (2 * r)
        corg = []
        for y in cor0:
            z = mini + int * n * y
            corg.append(round(z, 2))
        corg.append(round(p * maxi, 2))
        coresg.append(corg)
    return coresg


# Converter CMYK para RGB (Escala de 1)
def cmyk1_rgb1(cmyk):
    c, m, y, k = cmyk
    R = round((1 - c) * (1 - k), 4)
    G = round((1 - m) * (1 - k), 4)
    B = round((1 - y) * (1 - k), 4)
    rgb = [R, G, B]
    return rgb


# Converter elementos do arquivo base em dados
def converter_elem_graficos(lista, classe, escala, posicao):
    largura0, altura0 = escala
    x00, y00 = posicao
    lista_saida = []
    if classe == "Retangulos" or classe == "Corte":
        for w in lista:
            coord_p, escala_p = ((w["x0"] - x00) / largura0, (w["y0"] - y00) / altura0), (
                w['width'] / largura0, w['height'] / altura0)
            espessura = w['linewidth'] / altura0
            if classe == "Retangulos":
                cmyk_c, cmyk_i = (w["stroking_color"]), (w['non_stroking_color'])
                lista_saida.append(
                    {"Escala": escala_p, "Posicao": coord_p, "Cor_contorno": cmyk_c, "Cor": cmyk_i, "fill": w["fill"],
                     "Contorno": espessura})
            else:
                lista_saida.append({"Escala": escala_p, "Posicao": coord_p, "fill": False, "Contorno": espessura})

    return lista_saida


# Converte os textos do arquivo base em dados
def convert_carct(lista, escala, posicao):
    xt0, yt0 = posicao
    largt0, altt0 = escala
    lista_plvr = []
    for p in lista:
        posicao = [(p["x0"] - xt0) / largt0, (altt0 - p['doctop'] - 2 * yt0) / (altt0 + yt0)]
        cor_plvr = cmyk1_rgb1(p['non_stroking_color'])
        fonte = (p['height']) / altt0
        lista_plvr.append({"Caract": p["text"], "Fonte": fonte, 'Cor': cor_plvr, "Origem": posicao})
    return lista_plvr


# Ler informações do arquivo base
def extrair_dados(nome_arquivo):
    with pdfplumber.open(nome_arquivo) as pdf:
        for pagina in pdf.pages:
            curvas = pagina.objects['curve']
            caracteres = pagina.objects["char"]
            for i, v in enumerate(curvas):
                if v['stroking_color'] == [1.0]:
                    corte = curvas.pop(i)

    escala0 = [corte['width'], corte['height']]
    posicao0 = [corte["x0"], corte["y0"]]
    lista_elementos = converter_elem_graficos(curvas, "Retangulos", escala0, posicao0)
    lista_corte = converter_elem_graficos([corte], "Corte", escala0, posicao0)
    lista_crt = convert_carct(caracteres, escala0, posicao0)
    dic = {"Corte": lista_corte, "Retangulos": lista_elementos, "Caracteres": lista_crt}
    return dic


# Substituir os # pelos codigos
def substituir_caracteres(par_pla, par_pdf):
    lista_chaves = list(par_pla.keys())
    lista_del = []
    for h, dic_crt in enumerate(par_pdf):
        caracter = dic_crt["Caract"]
        if caracter == "#":
            for chave in lista_chaves:
                tamanho = len(chave)
                palavra = ""
                for v in range(tamanho):
                    palavra += par_pdf[h + v]["Caract"]
                if palavra == chave:
                    lista_chaves.remove(chave)
                    for c in range(tamanho - 1):
                        lista_del.append(h + c + 1)
                    dic_crt["Caract"] = chave
    lista_del.reverse()
    for dl in lista_del:
        par_pdf.pop(dl)
    return par_pdf


# Inserir marcas de registro
def inserir_marcas_registro(diametro_registro, lista, canv):
    for k in range(len(lista)):
        x, y = lista[k]
        canv.setStrokeColorCMYK(1, 1, 1, 1)
        canv.setFillColorCMYK(1, 1, 1, 1)
        canv.circle(x, y, diametro_registro / 2, fill=True)


# Criar retângulos no canva
def criar_retangulo(canv, dic, posicao, escala):
    posi = dic["Posicao"]
    xr, yr = posi
    x0, y0 = posicao
    x, y = x0 + (xr * escala), y0 + (yr * escala)
    esc = dic["Escala"]
    largret, alturaret = esc[0] * escala, esc[1] * escala
    fill = dic["fill"]
    if fill:
        C, M, Y, K = dic["Cor"]
        canv.setFillColorCMYK(C, M, Y, K)
    C, M, Y, K = dic["Cor_contorno"]
    canv.setStrokeColorCMYK(C, M, Y, K)
    canv.setLineWidth(dic["Contorno"] * escala * 1.2)
    canv.rect(x, y, largret, alturaret, fill=dic["fill"])


# Criar caracteres no canva
def criar_caract(canv, dictc, posicao, escala):
    posi = dictc["Origem"]
    xr, yr = posi
    x0, y0 = posicao
    x, y = round(x0 + (xr * escala), 4), round(y0 + (yr * escala), 4)
    fonte = dictc["Fonte"] * escala
    carct = dictc["Caract"]
    R, G, B = dictc["Cor"]
    canv.setFont('Arial-Bold', fonte)
    canv.setFillColorRGB(R, G, B)
    canv.drawString(x, y, carct)


# Criar o contorno de corte no canva
def criar_corte(canv, dic, posicao, escala):
    posi = dic["Posicao"]
    xr, yr = posi
    x0, y0 = posicao
    x, y = x0 + (xr * escala), y0 + (yr * escala)
    esc = dic["Escala"]
    largret, alturaret = esc[0] * escala, esc[1] * escala
    fill = dic["fill"]
    R, G, B = dic["COR"]
    canv.setStrokeColorRGB(R, G, B)
    canv.setLineWidth(dic["Contorno"] * escala)
    canv.rect(x, y, largret, alturaret, fill=dic["fill"])
