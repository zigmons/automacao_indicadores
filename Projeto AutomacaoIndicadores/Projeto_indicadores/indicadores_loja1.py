from salvar_planilha_pastabackup import *

#faturamento
#diversidade de produtos
#ticket medio

loja = 'Norte Shopping'
vendas_loja = dicionario_lojas[loja]
vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador,:]

#faturamento
faturamento_ano = vendas_loja['Valor Final'].sum()
faturamento_dia = vendas_loja_dia['Valor Final'].sum()

#diversidade de produtos
qtde_produtos_ano = len(vendas_loja['Produto'].unique())
qtde_produtos_dia = len(vendas_loja_dia['Produto'].unique())


#ticket medio

valor_venda = vendas_loja.groupby('Código Venda').sum()
ticket_medio_ano = valor_venda['Valor Final'].mean()

valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtde_produtos_dia = 4
meta_qtde_produtos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500