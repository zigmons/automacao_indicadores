from indicadores_loja1 import *
import win32com.client as win32




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