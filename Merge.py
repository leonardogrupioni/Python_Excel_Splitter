# Merge.py
import pandas as pd
import openpyxl

def mesclar_arquivos(lista_arquivos, arquivo_destino, callback_progresso=None):
    try:
        # Verificar se a lista de arquivos não está vazia
        if not lista_arquivos:
            raise ValueError("A lista de arquivos está vazia.")

        total_arquivos = len(lista_arquivos)

        # Ler o primeiro arquivo, incluindo o cabeçalho
        df_principal = pd.read_excel(lista_arquivos[0], engine='openpyxl', dtype=str)
        print(f"Lendo o arquivo principal: {lista_arquivos[0]}")

        # Atualizar a barra de progresso
        if callback_progresso:
            progresso = int((1 / total_arquivos) * 100)
            callback_progresso(progresso)

        # Iterar sobre os demais arquivos
        for idx, arquivo in enumerate(lista_arquivos[1:], start=2):
            print(f"Lendo e concatenando o arquivo: {arquivo}")
            # Ler o arquivo, pulando o cabeçalho (primeira linha)
            df_temp = pd.read_excel(arquivo, engine='openpyxl', dtype=str, header=None, skiprows=1)
            # Definir os nomes das colunas como sendo os mesmos do df_principal
            df_temp.columns = df_principal.columns
            # Concatenar ao DataFrame principal
            df_principal = pd.concat([df_principal, df_temp], ignore_index=True)

            # Atualizar a barra de progresso
            if callback_progresso:
                progresso = int((idx / total_arquivos) * 100)
                callback_progresso(progresso)

        # Escrever o DataFrame resultante em um novo arquivo Excel
        df_principal.to_excel(arquivo_destino, index=False, engine='openpyxl')

        # Ajustar o formato das células para texto
        ajustar_formato_texto(arquivo_destino)

        # Finalizar o progresso em 100%
        if callback_progresso:
            callback_progresso(100)

    except Exception as e:
        print(f"Erro ao mesclar os arquivos: {e}")
        raise

def ajustar_formato_texto(arquivo_excel):
    # Abrir o arquivo Excel com openpyxl
    wb = openpyxl.load_workbook(arquivo_excel)
    ws = wb.active

    # Ajustar o formato de todas as células para texto e converter valores para string
    for row in ws.iter_rows(min_row=1):  # Inclui o cabeçalho
        for cell in row:
            cell.number_format = '@'
            if cell.value is not None:
                cell.value = str(cell.value)

    wb.save(arquivo_excel)
