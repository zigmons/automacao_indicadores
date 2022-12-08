from criar_tabela_cada_loja import *
import win32com.client as win32

caminho_backup = pathlib.Path(r'C:\Users\rafae\iCloudDrive\PythonAulas\Projeto 1 - Automações de Processo - Aplicação mercado de trabalho\Projeto AutomacaoIndicadores\Backup Arquivos Lojas')


faturamento_lojas = vendas.groupby('Loja')[('Loja', 'Valor Final')].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r'C:\Users\rafae\iCloudDrive\PythonAulas\Projeto 1 - Automações de Processo - Aplicação mercado de trabalho\Projeto AutomacaoIndicadores\Backup Arquivos Lojas\{}'.format(nome_arquivo))


vendas_dia = vendas.loc[vendas['Data']==dia_indicador,:]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[('Loja', 'Valor Final')].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r'C:\Users\rafae\iCloudDrive\PythonAulas\Projeto 1 - Automações de Processo - Aplicação mercado de trabalho\Projeto AutomacaoIndicadores\Backup Arquivos Lojas\{}'.format(nome_arquivo))

outlook = win32.Dispatch('outlook.application')

nome = emails_bd.loc[emails_bd['Loja'] == loja, 'Gerente'].values[0]
mail = outlook.CreateItem(0)
mail.To = emails_bd.loc[emails_bd['Loja'] == 'Diretoria', 'E-mail'].values[0]
mail.Subject = 'Ranking{}/{}'.format(dia_indicador.day, dia_indicador.month)


mail.Body = f'''
Prezados,

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}

Segue em anexo os rankings do ano e do dia de todas as lojas.

Qualquer dúvida estou à disposição.
   '''

attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment))
attachment = pathlib.Path.cwd() / caminho_backup /  f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()
print(f'E-mail da Diretoria enviado')
