import pandas as pd
import openpyxl
import numpy as np
from itertools import combinations
import math

files_list = []


def metodo_soma(conciliacao, d8, folder):
    total_ades_usadas = []

    def soma_cpf_emprestimo(conciliacao_tratado, d8_geral, caminho):
        d8 = d8_geral[d8_geral['Serviço'] == 'Empréstimo Consignado'].copy()
        # print(f"DF D8: {d8}")
        conciliacao = conciliacao_tratado.copy()

        # total_ades_usadas = []

        conciliacao_encontrados = conciliacao.loc[conciliacao['PRODUTO'] == 'Empréstimo', ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO', 'AVERBAÇÃO - ATUALIZADA', 'PRODUTO', 'Lançou']]

        ades_usadas_cartao = soma_por_cpf(conciliacao_encontrados, d8, 'Empréstimo Consignado', caminho)

        total_ades_usadas.extend(ades_usadas_cartao)

        # Filtra o d8_geral_df UMA ÚLTIMA VEZ para garantir que temos o estado mais recente
        d8_final_restante = d8_geral[~d8_geral['Contrato'].isin(total_ades_usadas)]

        soma_cpf_beneficio(conciliacao, d8_final_restante, caminho)

    def soma_cpf_beneficio(conciliacao_tratado, d8_geral, caminho):
        d8 = d8_geral[d8_geral['Serviço'] == 'Cartão Benefício'].copy()
        conciliacao = conciliacao_tratado.copy()

        # total_ades_usadas = []

        conciliacao_encontrados = conciliacao.loc[
            conciliacao['PRODUTO'] == 'Cartão Benefício', ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO', 'AVERBAÇÃO - ATUALIZADA', 'PRODUTO']]

        ades_usadas_cartao = soma_por_cpf(conciliacao_encontrados, d8, 'Cartão Benefício', caminho)

        total_ades_usadas.extend(ades_usadas_cartao)

        # Filtra o d8_geral_df UMA ÚLTIMA VEZ para garantir que temos o estado mais recente
        d8_final_restante = d8_geral[~d8_geral['Contrato'].isin(total_ades_usadas)]
        soma_cpf_cartao_nao_lancado(conciliacao, d8_final_restante, caminho)


    def soma_cpf_cartao_nao_lancado(conciliacao_tratado, d8_geral, caminho):
        d8 = d8_geral[d8_geral['Serviço'] == 'Cartão de Crédito'].copy()
        conciliacao = conciliacao_tratado.copy()

        # total_ades_usadas = []

        conciliacao_encontrados = conciliacao.loc[
            conciliacao['PRODUTO'] == 'Cartão de Crédito', ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO',
                                                           'AVERBAÇÃO - ATUALIZADA', 'PRODUTO']]

        ades_usadas_cartao = soma_por_cpf(conciliacao_encontrados, d8, 'Cartão de Crédito NÃO LANÇADO', caminho)

        total_ades_usadas.extend(ades_usadas_cartao)

        # Filtra o d8_geral_df UMA ÚLTIMA VEZ para garantir que temos o estado mais recente
        d8_final_restante = d8_geral[~d8_geral['Contrato'].isin(total_ades_usadas)]
        soma_cpf_resto(conciliacao, d8_final_restante, caminho)

    def soma_cpf_resto(conciliacao_tratado, d8_geral, caminho):
        d8 = d8_geral.copy()
        conciliacao = conciliacao_tratado.copy()

        # total_ades_usadas = []

        conciliacao_encontrados = conciliacao.loc[
            conciliacao['PRODUTO'] == 'Cartão de Crédito', ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO',
                                                            'AVERBAÇÃO - ATUALIZADA', 'PRODUTO']]

        ades_usadas_cartao = soma_por_cpf(conciliacao_encontrados, d8, 'RESTO',caminho)

        total_ades_usadas.extend(ades_usadas_cartao)

    def soma_por_cpf(conciliacao_para_somar, d8, produto=None,caminho=None):
        # Conciliação e D8 já vai vir como DataFrame
        conciliacao_tratado = conciliacao_para_somar.copy()
        d8_geral = d8.copy()

        """
        Atribui ADEs da Planilha B para a Planilha A encontrando combinações de somas de parcelas por CPF,
        considerando uma tolerância de +20, +40 ou +60 reais.
        """
        print("Iniciando o processo com lógica de tolerância...")

        # 2. Preparar Dados
        conciliacao_tratado['CPF'] = conciliacao_tratado['CPF'].astype(str).str.strip()
        d8_geral['CPF'] = d8_geral['CPF'].astype(str).str.strip()
        conciliacao_tratado['PRESTAÇÃO'] = pd.to_numeric(conciliacao_tratado['PRESTAÇÃO'], errors='coerce')
        print(f'Parcelas da conciliação: {conciliacao_tratado['PRESTAÇÃO']}')

        print(f'Parcelas do D8: {d8_geral}')
        d8_geral['Valor original'] = d8_geral['Valor original'].astype(str).str.replace(".", "")
        d8_geral['Valor original'] = d8_geral['Valor original'].astype(str).str.replace(",", ".")
        d8_geral['Valor original'] = pd.to_numeric(d8_geral['Valor original'], errors='coerce')

        conciliacao_tratado.dropna(subset=['CPF', 'PRESTAÇÃO'], inplace=True)

        d8_geral.dropna(subset=['CPF', 'Valor original', 'Contrato'], inplace=True)

        # Adiciona as colunas de resultado na Planilha A
        conciliacao_tratado['ADE'] = None
        conciliacao_tratado['Soma_Calculada_A'] = None  # <-- ALTERADO: Nova coluna para auditoria
        conciliacao_tratado['Parcela_Encontrada_B'] = None  # <-- ALTERADO: Nova coluna para auditoria

        # 3. Estrutura de Busca Rápida para Planilha B
        b_lookup = {}
        for _, row in d8_geral.iterrows():
            cpf = row['CPF']
            parcela = round(row['Valor original'], 2)
            ade = row['Contrato']
            if cpf not in b_lookup:
                b_lookup[cpf] = {}
            if parcela not in b_lookup[cpf]:
                b_lookup[cpf][parcela] = [ade]
            else:
                b_lookup[cpf][parcela].append(ade)

        print("Estrutura de busca da Planilha B criada.")

        # 4. Processar por CPF
        cpfs_unicos = conciliacao_tratado['CPF'].unique()
        total_cpfs = len(cpfs_unicos)
        tolerancias = [0, 20, 40, 60]  # <-- ALTERADO: Lista de tolerâncias a serem testadas
        print(f"Encontrados {total_cpfs} CPFs únicos. Iniciando processamento com tolerâncias: {tolerancias}")

        ades_utilizadas = set()

        for i, cpf in enumerate(cpfs_unicos):
            '''if (i + 1) % 100 == 0:
                print(f"Processando CPF {i + 1}/{total_cpfs}...")'''

            if cpf not in b_lookup:
                continue

            itens_a_combinar = list(conciliacao_tratado[conciliacao_tratado['CPF'] == cpf][['PRESTAÇÃO']].itertuples())

            while itens_a_combinar:
                match_encontrado_nesta_iteracao = False

                for tamanho_comb in range(1, len(itens_a_combinar) + 1):
                    for comb in combinations(itens_a_combinar, tamanho_comb):
                        soma_parcelas = round(sum(item[1] for item in comb), 2)

                        for tol in tolerancias:
                            valor_alvo = round(soma_parcelas + tol, 2)

                            if valor_alvo in b_lookup.get(cpf, {}):
                                # 1. Pega a ADE disponível
                                ade_disponivel = b_lookup[cpf][valor_alvo].pop(0)

                                # 2. (CORREÇÃO) Adiciona IMEDIATAMENTE ao set de utilizadas
                                ades_utilizadas.add(ade_disponivel)

                                # Limpa a entrada do dicionário se não houver mais ADEs para essa parcela
                                if not b_lookup[cpf][valor_alvo]:
                                    del b_lookup[cpf][valor_alvo]

                                indices_para_atualizar = [item.Index for item in comb]

                                # 3. Atribui os valores ao DataFrame
                                conciliacao_tratado.loc[indices_para_atualizar, 'ADE'] = ade_disponivel
                                conciliacao_tratado.loc[indices_para_atualizar, 'Produto'] = produto
                                conciliacao_tratado.loc[indices_para_atualizar, 'Soma_Calculada_A'] = soma_parcelas
                                conciliacao_tratado.loc[indices_para_atualizar, 'Parcela_Encontrada_B'] = valor_alvo


                                # Remove os itens já combinados para não serem usados de novo neste CPF
                                itens_a_combinar = [item for item in itens_a_combinar if
                                                    item.Index not in indices_para_atualizar]

                                # A linha incorreta foi removida daqui

                                match_encontrado_nesta_iteracao = True
                                break  # Sai do loop de tolerâncias

                        if match_encontrado_nesta_iteracao:
                            break  # Sai do loop de combinações
                    if match_encontrado_nesta_iteracao:
                        break  # Sai do loop de tamanhos de combinação

                # Esta condição de 'break' pode fazer o loop do CPF parar prematuramente.
                # Se você quer que ele tente TODAS as combinações possíveis para um CPF,
                # mesmo depois de achar uma, a lógica de loop precisaria de ajuste.
                # Mas mantendo sua lógica original:
                if not match_encontrado_nesta_iteracao:
                    # Isso significa que nenhuma combinação foi encontrada para os itens restantes deste CPF
                    break

        # Faz uma copia sem ADEs vazias para juntar no módulo main.py
        conciliacao_tratado_sem_vazios = conciliacao_tratado[~(conciliacao_tratado['ADE'].isna()) | ~(conciliacao_tratado['ADE'] == '')]

        # 7. Salvar Resultado (fora do loop de CPF)
        caminho_saida = fr"{caminho}\Planilha_A_com_ADE_e_Tolerancia {produto}.xlsx"
        conciliacao_tratado.to_excel(caminho_saida, index=False)
        files_list.append(caminho_saida)
        print(f"\nProcesso concluído! O resultado foi salvo em '{caminho_saida}'.")

        return list(ades_utilizadas)

    # Inicia a primeira função
    soma_cpf_emprestimo(conciliacao, d8, folder)

    # Hora de retornar todos os arquivos
    df_averbacao_unificada = pd.concat([pd.read_excel(arquivo) for arquivo in files_list], ignore_index=True)
    nome_arquivo_metodo_soma = fr'{folder}\DADOS DE AVERBAÇÃO UNIFICADAS METODO SOMA.xlsx'
    df_averbacao_unificada.to_excel(nome_arquivo_metodo_soma, index=False)


    return df_averbacao_unificada, list(total_ades_usadas), list(files_list)