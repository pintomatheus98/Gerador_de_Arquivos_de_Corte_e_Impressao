from Funcoes import *

pdfmetrics.registerFont(TTFont('Arial-Bold', 'C:/Windows/Fonts/arialbd.ttf'))
dados = ler_dados_excel("ENTRADA\Planilha Modelo.xlsx")
lista_amostras = []

'''___________________________ CRIAR LISTA DE AMOSTRAS _________________________________'''

#  Definir variaveis
cont = 1  # Contador
crt = dados["ANALISE"]['CODIGO DA AMOSTRA']  # Caractere do código
dic_P1 = dados["PARAMETRO 1"]  # Dicionario do parâmetro 1
dic_P2 = dados["PARAMETRO 2"]  # Dicionario do parâmetro 1
P1 = dic_P1['NOME']  # Parâmetro 1
P2 = dic_P2['NOME']  # Parâmetro 2
lista_P1 = criar_lista(dic_P1)  # Lista de valores do parâmetro 1
lista_P2 = criar_lista(dic_P2)  # Lista de valores do parâmetro 1
lista_rgb = gerar_cores_rgb(len(lista_P1))  # Lista de cores primarias

#  Combinar valores e dados das amostras
for i, valor_P1 in enumerate(lista_P1):
    # Criar um gadiente de uma cor RGB ( os tons são gerados em CMYK )
    lista_cores = gradiente_cores_cmyk(lista_rgb[i], len(lista_P2), 0.05, 0.7)
    for j, valor_P2 in enumerate(lista_P2):
        dic_amostra = {}  # Dicionario da amostra
        uni_P1 = dic_P1['UNIDADE DE MEDIDA']  # Unidade de medida P1
        uni_P2 = dic_P2['UNIDADE DE MEDIDA']  # Unidade de medida P1
        RGB = cmyk1_rgb1(lista_cores[j])  # Cor da Amostra
        cod = crt + str(cont).zfill(2)  # Codigo amostra

        # Adicionar dados as ao dicionario da amostra
        dic_amostra["COD"] = cod
        dic_amostra['P1'] = round(valor_P1, 2)
        dic_amostra["uniP1"] = uni_P1
        dic_amostra['P2'] = valor_P2
        dic_amostra["uniP2"] = uni_P2
        dic_amostra["RGB"] = RGB
        lista_amostras.append(dic_amostra)
        cont += 1

'''____________________________ CRIAR PDF _________________________________'''

#   Dados do material para ordem de serviço
material = dados["ANALISE"]["MATERIAL"]  # Material
espessura = dados["ANALISE"]['ESPESSURA']  # Espessura do material
sup = dados["ANALISE"]['SUPERFICIE']  # Superficie do material

#   Definir variaveis
dic_impressao = dados['DADOS DE IMPRESSÃO']  # Dados de impressão
dic_variaveis = dados['VARIAVEIS DE IMPRESSÃO']  # Variaveis de impressão
arquivo_base = dic_impressao['ARQUIVO BASE']  # Endereço arquivo base
altura = mm_para_pt(dados['ANALISE']['ALTURA DA AMOSTRA'])  # Altura da amostra em points
Nome_teste = dados["ANALISE"]["NOME DA ANALISE"]  # Nome do teste
esp = mm_para_pt(dic_impressao["ESPAÇAMENTO"])  # Espaçamento entre amostras
fol_pag = mm_para_pt(dic_impressao["FOLGA PAGINA"])  # Folga da página
marca_registro = mm_para_pt(dic_impressao["DIAMETRO REGISTRO"])  # Diâmetro da marca de registro
nP1 = dados['PARAMETRO 1']['NUMERO DE AMOSTRAS']  # Numero de valores do parâmetro 1
nP2 = dados['PARAMETRO 2']['NUMERO DE AMOSTRAS']  # Numero de valores do parâmetro 2

#   Retirar elementos do arquivo base
dados_arquivo_base = extrair_dados(arquivo_base)  # ler os dados do pdf base
list_crt = substituir_caracteres(dic_variaveis, dados_arquivo_base["Caracteres"])  # Inclui os códigos no texto

#   Definir dimensões do arquivo
alt_peca = altura + esp  # Altura total peça
esp_reg = fol_pag + (marca_registro / 2)  # Espaçamento para registrp
margem_pagina = fol_pag + (esp / 2) + marca_registro  # Margem da pagina
area_corte = (alt_peca * nP1, alt_peca * nP2)  # Area total de corte
pagina = [(2 * margem_pagina) + area_corte[0], (2 * margem_pagina) + area_corte[1]]  # Dimenssões da página
alt_arq_mm = int(pt_para_mm(pagina[0]))  # Altura do arquivo
lar_arq_mm = int(pt_para_mm(pagina[1]))  # Altura do arquivo

#   Definir nome dos arquivos
Nome_arq_corte = f"SAIDA\CORTE {Nome_teste}_{material} {espessura} {sup} _{alt_arq_mm}x{lar_arq_mm}mm_1x.pdf"
Nome_arq_impr = f"SAIDA\IMPRESSAO {Nome_teste}_{material} {espessura} {sup} _{alt_arq_mm}x{lar_arq_mm}mm_1x.pdf"

#   Criar canvas dos arquivos
canc = canvas.Canvas(Nome_arq_corte, pagesize=pagina)  # Canvas do Corte
cani = canvas.Canvas(Nome_arq_impr, pagesize=pagina)  # Canvas do Impressão

#   Definir coordenadas das marcas de registro
cord_amostras = []
for i in range(nP1):
    for j in range(nP2):
        x = margem_pagina + (esp / 2) + (alt_peca * i)
        y = pagina[1] - margem_pagina - alt_peca - (alt_peca * j) + (esp / 2)
        cord_amostras.append([x, y])

#   Inserir marcas de registro nos arquivos
cord_reg = [(esp_reg, esp_reg), (pagina[0] - esp_reg, esp_reg + (area_corte[1] / 4)),
            (pagina[0] - esp_reg, pagina[1] - esp_reg), (esp_reg, pagina[1] - esp_reg),
            (pagina[0] - esp_reg, esp_reg)]  # Lista de coordenadas das marcas de registro
inserir_marcas_registro(marca_registro, cord_reg, canc)  # Insere marcas de registro no arquivo de corte
inserir_marcas_registro(marca_registro, cord_reg, cani)  # Insere marcas de registro no arquivo de impressão

#   Inserir valores variados nos dicionarios da amostra/ Verifica se possui unidade de medida
for amostra in lista_amostras:
    for variavel, valor in dic_variaveis.items():
        uni = ""
        if valor in amostra.keys():
            k_uni = f'uni{valor}'
            if k_uni in amostra.keys():
                uni = f' {amostra[k_uni]}'
            amostra[variavel] = f"{amostra[valor]}{uni}"
        else:
            for chave1 in amostra.keys():
                if chave1 in valor:
                    k_uni = f'uni{chave1}'
                    if k_uni in amostra.keys():
                        uni = f' {amostra[k_uni]}'
                    valor = valor.replace(chave1, f"{amostra[chave1]}")
            for chave2 in dados['ANALISE'].keys():
                if chave2 in valor:
                    k_uni = f'uni{chave2}'
                    if k_uni in amostra.keys():
                        uni = f' {amostra[k_uni]}'
                    valor = valor.replace(chave2, f"{dados['ANALISE'][chave2]}")
            try:
                amostra[variavel] = str(round(eval(valor), 1)) + uni
            except:
                continue

#   Inserir slementos das amostras nos arquivos
for i, amostra in enumerate(lista_amostras):  # Itera o dicionario de amostras
    coord = cord_amostras[i]  # Seleciona a coordenada

    for v in dados_arquivo_base["Retangulos"]:  # Iterar elementos gráficos
        criar_retangulo(cani, v, coord, altura)  # Inserir no arquivo de impressão

    for w in dados_arquivo_base["Corte"]:  # Iterar contornos de corte
        w["COR"] = amostra["RGB"]  # Definir cor
        criar_corte(canc, w, coord, altura)  # Inserir no arquivo de corte

    for dic_palavra in list_crt:  # Iterar caracteres
        palavra = dic_palavra["Caract"]  # Definir texto
        if palavra in amostra.keys():  # Checa se é variavel
            dic_palavra_n = dic_palavra.copy()  # Busca a Chave
            dic_palavra_n["Caract"] = amostra[palavra]  # Busca o valor
            criar_caract(cani, dic_palavra_n, coord, altura)  # Insere no arquivo de impressão
        else:
            criar_caract(cani, dic_palavra, coord, altura)  # Insere no arquivo de impressão

cani.save()  # Salva o arquivo de corte
canc.save()  # Salva o arquivo de corte
