import pandas as pd
import pathlib

# importar e tratar as bases de dados
# criar 1 arquivo para cada loja
# salvar backup nas pastas
# calcular os indicadores
# enviar o onepage
# enviar o email para diretoria

emails_bd = pd.read_excel(r'C:\Users\rafae\iCloudDrive\PythonAulas\Projeto 1 - Automações de Processo - Aplicação mercado de trabalho\Projeto AutomacaoIndicadores\Bases de Dados\Emails.xlsx')
lojas_bd = pd.read_csv(r'C:\Users\rafae\iCloudDrive\PythonAulas\Projeto 1 - Automações de Processo - Aplicação mercado de trabalho\Projeto AutomacaoIndicadores\Bases de Dados/Lojas.csv', encoding='latin1', sep=';')
vendas_bd = pd.read_excel(r'C:\Users\rafae\iCloudDrive\PythonAulas\Projeto 1 - Automações de Processo - Aplicação mercado de trabalho\Projeto AutomacaoIndicadores\Bases de Dados\Vendas.xlsx')
# print(emails_bd)
# print(lojas_bd)
# print(vendas_bd)