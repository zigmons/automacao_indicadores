import pandas as pd
import pathlib
import win32com.client as win32


emails_bd = pd.read_excel(r'C:\Users\rafae\iCloudDrive\PythonAulas\Projeto 1 - Automações de Processo - Aplicação mercado de trabalho\Projeto AutomacaoIndicadores\Bases de Dados\Emails.xlsx')
lojas_bd = pd.read_csv(r'C:\Users\rafae\iCloudDrive\PythonAulas\Projeto 1 - Automações de Processo - Aplicação mercado de trabalho\Projeto AutomacaoIndicadores\Bases de Dados/Lojas.csv', encoding='latin1', sep=';')
vendas_bd = pd.read_excel(r'C:\Users\rafae\iCloudDrive\PythonAulas\Projeto 1 - Automações de Processo - Aplicação mercado de trabalho\Projeto AutomacaoIndicadores\Bases de Dados\Vendas.xlsx')

vendas = vendas_bd.merge(lojas_bd, on='ID Loja')

dicionario_lojas = {}

for loja in lojas_bd['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja,:]

dia_indicador = vendas['Data'].max()

caminho_backup = pathlib.Path(r'C:\Users\rafae\iCloudDrive\PythonAulas\Projeto 1 - Automações de Processo - Aplicação mercado de trabalho\Projeto AutomacaoIndicadores\Backup Arquivos Lojas')

arquivos_pasta_backup = caminho_backup.iterdir()

lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()

    nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
    local_arquivo = caminho_backup / loja / nome_arquivo
    dicionario_lojas[loja].to_excel(local_arquivo)

meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtde_produtos_dia = 4
meta_qtde_produtos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500


for loja in dicionario_lojas:

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


    outlook = win32.Dispatch('outlook.application')

    nome = emails_bd.loc[emails_bd['Loja']==loja,'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.To = emails_bd.loc[emails_bd['Loja']==loja,'E-mail'].values[0]
    mail.Subject = 'OnePage Dia{}/{} - Loja{}'.format(dia_indicador.day, dia_indicador.month, loja)

    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'
    if qtde_produtos_dia >= meta_qtde_produtos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'
    if qtde_produtos_ano >= meta_qtde_produtos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'
    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'
    mail.HTMLBody = f'''
     <p>Bom dia, {nome}</p>
    
        <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong> foi:</p>
    
        <table>
          <tr>
            <th>Indicador</th>
            <th>Valor Dia</th>
            <th>Meta Dia</th>
            <th>Cenário Dia</th>
          </tr>
          <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${faturamento_dia:.2f}</td>
            <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
            <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
          </tr>
          <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{qtde_produtos_dia}</td>
            <td style="text-align: center">{meta_qtde_produtos_dia}</td>
            <td style="text-align: center"><font color="{cor_qtde_dia}">◙</font></td>
          </tr>
          <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
            <td style="text-align: center">R${meta_ticketmedio_dia:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
          </tr>
        </table>
        <br>
        <table>
          <tr>
            <th>Indicador</th>
            <th>Valor Ano</th>
            <th>Meta Ano</th>
            <th>Cenário Ano</th>
          </tr>
          <tr>
            <td>Faturamento</td>
            <td style="text-align: center">R${faturamento_ano:.2f}</td>
            <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
            <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
          </tr>
          <tr>
            <td>Diversidade de Produtos</td>
            <td style="text-align: center">{qtde_produtos_ano}</td>
            <td style="text-align: center">{meta_qtde_produtos_ano}</td>
            <td style="text-align: center"><font color="{cor_qtde_ano}">◙</font></td>
          </tr>
          <tr>
            <td>Ticket Médio</td>
            <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
            <td style="text-align: center">R${meta_ticketmedio_ano:.2f}</td>
            <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
          </tr>
        </table>
    
        <p>Segue em anexo a planilha com todos os dados para mais detalhes.</p>
    
        <p>Qualquer dúvida estou à disposição.</p>
        <p>Att., Rafael Sousa</p>
    
    
    
    
    '''

    attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment))
    mail.Send()
    print(f'E-mail da Loja {loja} enviado')