# Splitter.py
import pandas as pd
import os
import openpyxl

def obter_num_linhas(caminho_arquivo):
    try:
        caminho_arquivo = os.path.normpath(caminho_arquivo)
        print(f"Lendo o arquivo em: {caminho_arquivo}")
        # Ler todas as colunas como strings
        df = pd.read_excel(caminho_arquivo, engine='openpyxl', dtype=str)
        num_linhas = len(df.index)
        return num_linhas
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return None

def dividir_arquivo(caminho_arquivo, partes, callback_progresso=None):
    try:
        # Ler todas as colunas como strings
        df = pd.read_excel(caminho_arquivo, engine='openpyxl', dtype=str)
        num_linhas = len(df.index)
        linhas_por_parte = num_linhas // partes
        resto = num_linhas % partes

        inicio = 0
        for i in range(1, partes + 1):
            linhas_extra = 1 if i <= resto else 0
            fim = inicio + linhas_por_parte + linhas_extra
            df_part = df.iloc[inicio:fim]
            nome_arquivo = os.path.join(os.path.dirname(caminho_arquivo), f"Lote_{i}.xlsx")

            # Escrever o DataFrame para Excel
            df_part.to_excel(nome_arquivo, index=False, engine='openpyxl')

            # Abrir o arquivo Excel com openpyxl para ajustar o formato das células
            wb = openpyxl.load_workbook(nome_arquivo)
            ws = wb.active

            # Ajustar o formato de todas as células para texto e converter valores para string
            for row in ws.iter_rows(min_row=1):  # Inclui o cabeçalho
                for cell in row:
                    cell.number_format = '@'
                    if cell.value is not None:
                        cell.value = str(cell.value)

            wb.save(nome_arquivo)
            inicio = fim

            # Atualizar a barra de progresso
            if callback_progresso:
                progresso = int((i / partes) * 100)
                callback_progresso(progresso)

        # Finalizar o progresso em 100%
        if callback_progresso:
            callback_progresso(100)

    except Exception as e:
        print(f"Erro ao dividir o arquivo: {e}")
        raise
