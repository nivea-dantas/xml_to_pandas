# Imports
from xml.dom import minidom
import pandas as pd
import re
from datetime import datetime

#pega a data corrente para inserir na coluna Data de Lançamento e extrai o mes para tratar e colocar na coluna Mês
today = datetime.now()
today = today.strftime('%Y-%m-%d')
mes = today[5:7]

# Tratamento para a coluna Mês
if mes == '01':
    mes = 'JANEIRO'
elif mes == '02':
    mes = 'FEVEREIRO'
elif mes == '03':
    mes = 'MARÇO'
elif mes == '04':
    mes = 'ABRIL'
elif mes == '05':
    mes = 'MAIO'
elif mes == '06':
    mes = 'JUNHO'
elif mes == '07':
    mes = 'JULHO'
elif mes == '08':
    mes = 'AGOSTO'
elif mes == '09':
    mes = 'SETEMBRO'
elif mes == '10':
    mes = 'OUTUBRO'
elif mes == '11':
    mes = 'NOVEMBRO'
else:
    mes = 'DEZEMBRO'

# Criação de path do arquivo xml e excel para tratamento.
# Incluir o caminho do arquivo. Ex: C:/documents/user - sem a barra final
# incluir o nome do arquivo sem a extensão
path = str(input('coloque o caminho do arquivo XML:'))
path = path.replace('\\', '/')
arquivo = str(input('Digite o nome do arquivo XML:'))
arquivo = path +'/'+ arquivo + '.xml'
path2 = str(input('coloque o caminho do arquivo EXCEL:'))
path2 = path2.replace('\\', '/')
arquivo2 = str(input('Digite o nome do arquivo EXCEL:'))
arquivo2 = path2 +'/'+ arquivo2 + '.xlsx'
aba_planilha = str(input('Digite o nome da Aba da planilha para inserir a informacao:'))

# leitura do arquivo Excel
exel = pd.read_excel(arquivo2, sheet_name= aba_planilha, header=5)

# Leitura do arquivo xml e extração dos dados por meio das tags
with open(arquivo, 'r') as f:
    xml = minidom.parse(f)
    dtEmi = xml.getElementsByTagName('dhEmi')
    RSocial = xml.getElementsByTagName('xNome')
    NFantasia = xml.getElementsByTagName('xFant')
    qtde = xml.getElementsByTagName('qCom')
    peca = xml.getElementsByTagName('xProd')
    vl_tot = xml.getElementsByTagName('vProd')
    nfe = xml.getElementsByTagName('nNF')
    desconto = xml.getElementsByTagName('vDesc')
    vl_final = xml.getElementsByTagName('vNF')

'''Segundo o negócio, foi necessário incluir mais de uma informação em uma mesma coluna, 
pois, estavam inputando informações por numeros de NFE e não por SKU. Assim foi realizado este tratamento personalizado para a necessidade do cliente.'''
peca2 = []
count = 0
if len(peca) > 1:
    for i in peca:
        dado = i.firstChild.data
        peca2.append(dado)
    separador = ' + '
    peca = separador.join(peca2)
    qtde2 = []
    count = 0
    for i in qtde:
        dado = i.firstChild.data
        qtde2.append(float(dado))
    qtde = sum(qtde2)
    vl_tot2 = []
    count = 0
    for i in vl_tot:
        dado = i.firstChild.data
        vl_tot2.append(float(dado))
    vl_tot = sum(vl_tot2)
    desconto2 = []
    count = 0
    for i in desconto:
        dado = i.firstChild.data
        desconto2.append(float(dado))
    desconto = sum(desconto2)
else: 
    peca = peca[0].firstChild.data
    qtde = qtde[0].firstChild.data
    vl_tot = vl_tot[0].firstChild.data
    desconto = desconto[0].firstChild.data


#joga dados em uma lista para extrair as informações
lista = [dtEmi, RSocial, NFantasia, nfe, vl_final]

#checa se existe algum dado vazio, se houver irá colocar o item como sem informação
lista_elementos = []
for i in lista:
    if len(i) > 0:
        dado = i[0].firstChild.data
        lista_elementos.append(dado)
    else:
        dado = 'sem informação'
        lista_elementos.append(dado)

lista_elementos.append(peca)
lista_elementos.append(qtde)
lista_elementos.append(vl_tot)
lista_elementos.append(desconto)
lista_elementos[0] = lista_elementos[0][:10]

# insere uma linha no arquivo e grava as indormações conforme o nome da coluna
exel.iloc[-1] = {'MÊS': mes,
                 'DATA DE LAÇAMENTO': today,
                 'DATA EMISSÃO':lista_elementos[0], 
                 'RAZÃO SOCIAL': lista_elementos[1], 
                 'NOME FANTASIA': lista_elementos[2], 
                 'QTDE': lista_elementos[6], 
                 'PEÇAS': lista_elementos[5], 
                 'VALOR  TOTAL': lista_elementos[7], 
                 'NFE / RECIBO': lista_elementos[3],
                 'DESCONTO': lista_elementos[8], 
                 'VALOR FINAL': lista_elementos[4]}

# Tratamento de Erro e saving das informações no arquivo
with pd.ExcelWriter(arquivo2, mode='a', engine="openpyxl", if_sheet_exists='replace') as writer:
    exel.to_excel(writer, sheet_name = aba_planilha)
    print('Informações adicionadas com sucesso.')