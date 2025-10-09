import pandas as pd
import numpy as np
import openpyxl
import re

def trata_contratos(d8_do_main, conciliacao_do_main, caminho):
    d8_reduzido = d8_do_main[['Matrícula', 'CPF', 'Nome', 'Contrato', 'Cod na Instituição', 'Serviço', 'Valor original']].copy()
    conciliacao = conciliacao_do_main.copy()

    def extrair_contratos_com_referencia(df_sujo: pd.DataFrame, df_limpo: pd.DataFrame) -> pd.DataFrame:
        """
        Extrai e limpa números de contrato de um DataFrame usando outro como referência.

        Args:
            df_sujo (pd.DataFrame): O DataFrame correspondente à "Planilha A",
                                    com a coluna de contratos poluída (ex: 'CONTRATOS')
                                    e uma coluna de CPF (ex: 'CPF').
            df_limpo (pd.DataFrame): O DataFrame correspondente à "Planilha B",
                                     com colunas de contratos limpos e CPF.

        Returns:
            pd.DataFrame: O DataFrame original (df_sujo) com novas colunas para cada
                          contrato encontrado e limpo.
        """
        print("Iniciando o processo de extração de contratos...")

        # --- Passo 1: Criar o mapa de referência a partir da Planilha B ---
        # Garante que os contratos na planilha de referência sejam texto para busca
        df_limpo['CONTRATO'] = df_limpo['CONTRATO'].astype(str).str.strip()

        # Agrupa por CPF e cria uma lista de contratos limpos para cada um
        print("Criando mapa de referência CPF -> Contratos...")
        mapa_cpf_contratos = df_limpo.groupby('CPF')['CONTRATO'].apply(list).to_dict()

        # --- Passo 2: Definir a função que será aplicada em cada linha da Planilha A ---
        def encontrar_contratos_na_linha(row):
            cpf = row['CPF']
            texto_contratos_sujo = str(row['Cod na Instituição'])

            # Pega a lista de contratos válidos para este CPF. Se o CPF não existir no mapa, retorna lista vazia.
            contratos_validos_para_cpf = mapa_cpf_contratos.get(cpf, [])

            if not contratos_validos_para_cpf:
                return []  # Retorna vazio se não houver contratos de referência para este CPF

            encontrados = []
            # Para cada contrato limpo, verifica se ele está contido no texto sujo
            for contrato in contratos_validos_para_cpf:
                # A verificação "in" é simples e poderosa aqui
                if contrato in texto_contratos_sujo:
                    encontrados.append(contrato)

            return encontrados

        # --- Passo 3: Aplicar a função e criar as novas colunas ---
        print("Analisando a Planilha A e extraindo os contratos...")
        # Aplica a função em cada linha de df_sujo para obter uma lista de contratos encontrados
        lista_de_contratos_encontrados = df_sujo.apply(encontrar_contratos_na_linha, axis=1)

        # Converte a Série de listas em um novo DataFrame
        df_contratos_novos = pd.DataFrame(lista_de_contratos_encontrados.tolist(), index=df_sujo.index)

        # Renomeia as colunas do novo DataFrame para 'Contrato_Encontrado_1', 'Contrato_Encontrado_2', etc.
        df_contratos_novos.columns = [f'Contrato_Encontrado_{i + 1}' for i in df_contratos_novos.columns]

        # Junta o DataFrame original com as novas colunas de contratos
        df_resultado = pd.concat([df_sujo, df_contratos_novos], axis=1)

        print("Processo concluído com sucesso!")
        df_resultado.to_excel(fr'{caminho}\D8 Contratos tratados.xlsx', index=False)
        return df_resultado

    extrair_contratos_com_referencia(d8_reduzido, conciliacao)


