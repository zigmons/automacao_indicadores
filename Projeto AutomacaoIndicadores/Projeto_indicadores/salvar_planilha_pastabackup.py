import pathlib

from criar_tabela_cada_loja import *

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