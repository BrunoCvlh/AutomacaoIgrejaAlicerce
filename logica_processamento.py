# logica_processamento.py

import pandas as pd
import os
from pathlib import Path


def processar_e_transformar_planilha(caminho_arquivo):
    """
    Carrega o arquivo, aplica ffill na Coluna A e transforma o formato wide (mensal)
    para o formato long (colunas de identificação + Colunas C, D, e nova coluna Data),
    salvando o resultado na pasta 'Downloads'.

    Args:
        caminho_arquivo (str): O caminho completo para o arquivo de entrada.

    Returns:
        str: Uma mensagem de sucesso ou erro.
    """
    if not caminho_arquivo or not os.path.exists(caminho_arquivo):
        return "Erro: O caminho do arquivo é inválido ou o arquivo não existe."

    try:
        # 1. Definir o caminho de destino na pasta 'Downloads'
        download_path = Path.home() / 'Downloads'
        nome_original = Path(caminho_arquivo).stem
        nome_novo_arquivo = f"{nome_original}_Transformado.xlsx"
        caminho_destino = download_path / nome_novo_arquivo

        # 2. Carregar o arquivo com header=None para capturar todas as linhas de cabeçalho
        if caminho_arquivo.lower().endswith(('.csv')):
            try:
                # Tenta UTF-8 e, se falhar, tenta Latin-1
                df_raw = pd.read_csv(caminho_arquivo, header=None, sep=',', encoding='utf-8')
            except UnicodeDecodeError:
                df_raw = pd.read_csv(caminho_arquivo, header=None, sep=',', encoding='latin1')
        else:
            # Trata como XLSX (garantindo que lê a planilha do início)
            df_raw = pd.read_excel(caminho_arquivo, header=None, engine='openpyxl')

        # 3. Pré-processamento e Criação do Cabeçalho
        # Linha 2 (índice 1): Datas. Preenche NaN com o último valor não vazio (ffill).
        dates_row = df_raw.iloc[1].mask(df_raw.iloc[1].isna()).ffill().fillna('')

        # Linha 3 (índice 2): Nomes das métricas.
        metrics_row = df_raw.iloc[2]

        # Cria um novo cabeçalho combinando Data e Métrica (ex: "2025-01-01 - Saldo")
        new_header = []
        for i in range(len(df_raw.columns)):
            metric = metrics_row.iloc[i]
            if i < 2:  # Colunas A e B: 'Centro de custo' e 'Categoria'
                new_header.append(metric)
            else:  # Colunas de dados (a partir da coluna C)
                date = dates_row.iloc[i]
                new_header.append(f"{date} - {metric}")

        # 4. Criar o DataFrame de dados (a partir da linha 4, índice 3) e aplicar o novo cabeçalho
        df = df_raw.iloc[3:].copy()
        df.columns = new_header

        # 5. Aplicar FFILL na Coluna A ('Centro de custo') conforme solicitado
        df['Centro de custo'].ffill(inplace=True)

        # 6. Melt (Unpivot) - Transformar wide para long
        id_vars = ['Centro de custo', 'Categoria']

        # Colunas a serem derretidas (melted): todas que contêm a data e a métrica no nome
        value_vars = [col for col in df.columns if ' - ' in str(col)]

        df_long = pd.melt(df,
                          id_vars=id_vars,
                          value_vars=value_vars,
                          var_name='Data_Metrica',
                          value_name='Valor')

        # 7. Separar 'Data' e 'Metrica'
        df_long[['Data', 'Metrica']] = df_long['Data_Metrica'].str.split(' - ', expand=True)
        df_long['Data'] = pd.to_datetime(df_long['Data'])
        df_long.drop(columns=['Data_Metrica'], inplace=True)

        # 8. Limpeza: Remover linhas onde o 'Valor' é nulo (células vazias)
        df_long.dropna(subset=['Valor'], inplace=True)

        # 9. Pivotar a Metrica de volta para colunas (Valor baixado (bruto) e Saldo)
        df_final = df_long.pivot_table(index=['Centro de custo', 'Categoria', 'Data'],
                                       columns='Metrica',
                                       values='Valor',
                                       aggfunc='first').reset_index()

        # 10. Reordenar as colunas no formato final: Centro de custo, Categoria, Data, V. Bruto, Saldo
        colunas_finais = ['Centro de custo', 'Categoria', 'Data']

        if 'Valor baixado (bruto)' in df_final.columns:
            colunas_finais.append('Valor baixado (bruto)')
        if 'Saldo' in df_final.columns:
            colunas_finais.append('Saldo')

        # Inclui quaisquer outras colunas de métrica que possam ter sido pivotadas (futuras adições)
        outras_colunas = [col for col in df_final.columns if col not in colunas_finais]
        colunas_finais.extend(outras_colunas)

        df_final = df_final[colunas_finais]

        # 11. Salvar o novo arquivo transformado como XLSX
        with pd.ExcelWriter(caminho_destino, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='Dados Longos', index=False, header=True)

        return f"Sucesso! O arquivo foi transformado para o formato longo e salvo em: {caminho_destino}"

    except Exception as e:
        return f"Ocorreu um erro durante o processamento: {e}"