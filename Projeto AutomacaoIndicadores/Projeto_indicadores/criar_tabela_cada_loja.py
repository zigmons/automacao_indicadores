from importacao_basedados import *

vendas = vendas_bd.merge(lojas_bd, on='ID Loja')
# print(vendas)

dicionario_lojas = {}

for loja in lojas_bd['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja,:]
# print(dicionario_lojas['Rio Mar Recife'])
# print(dicionario_lojas['Shopping Vila Velha'])


#definir dia do indicador

dia_indicador = vendas['Data'].max()
# print(dia_indicador)
# print('{}/{}'.format(dia_indicador.day, dia_indicador.month))