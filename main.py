import pandas as pd
import openpyxl
import numpy as np
import tkinter as tk
from tkinter import filedialog
from MetodoSoma import metodo_soma
from TrataContratos import trata_contratos


files_list = []

def selecionar_arquivo(titulo="Selecione um arquivo", multiplos=False):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    if multiplos:
        arquivos = filedialog.askopenfilenames(
            title=titulo,
            filetypes=[("Arquivos Excel", "*.xlsx *.xls *.csv"), ("Todos os arquivos", "*.*")]
        )
        root.destroy()
        return list(arquivos)
    else:
        arquivo = filedialog.askopenfilename(
            title=titulo,
            filetypes=[("Arquivos Excel", "*.xlsx *.xls *.csv"), ("Todos os arquivos", "*.*")]
        )
        root.destroy()
        return arquivo


def selecionar_pasta(titulo="Selecione uma pasta"):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    pasta = filedialog.askdirectory(title=titulo)
    root.destroy()
    if not pasta:
        return None
    return pasta

def selecionar_com_validacao(titulo, extensao_correta):
    while True:
        arquivo = selecionar_arquivo(titulo)
        if not arquivo:  # usuário cancelou ou fechou
            return None
        if arquivo.lower().endswith(extensao_correta):
            return arquivo
        else:
            print(f"Arquivo inválido! Selecione um arquivo com extensão {extensao_correta}")



# --- CONFIGURAÇÃO ---

# Credbase trabalhado
credbase_trabalhado = selecionar_com_validacao(r"Selecione o Credbase Trabalhado", 'xlsx')

# Conciliação são os dados averbados
conciliacao_bruto = selecionar_com_validacao(r"Selecione o arquivo de Conciliação bruto", "xlsx")

# D8 Geral do convenio
planilha_d8_geral = selecionar_com_validacao(r"Selecione a planilha geral de D8", "csv")

# D8 apenas para os casos de cartão
planilha_d8 = selecionar_com_validacao(r"Selecione o arquivo de retorno unificado do convênio", "csv")

folder = selecionar_pasta('Insira o caminho de saída')

# --------------------

def separacao_conciliacao(credbase, conciliacao):
    credbase_tratado = pd.read_excel(credbase)
    credbase_tratado['Codigo Credbase'] = credbase_tratado['Codigo Credbase'].astype(int)

    conciliacao_atratar = pd.read_excel(conciliacao)
    conciliacao_atratar['CONTRATO'] = conciliacao_atratar['CONTRATO'].astype(int)

    credbase_tratado['CONTSE'] = credbase_tratado.groupby('Codigo Credbase')['Codigo Credbase'].transform('count')

    # Puxa para a conciliação o contse que fizemos no cred
    conciliacao_atratar['Lançou'] = conciliacao_atratar['CONTRATO'].map(credbase_tratado.set_index('Codigo Credbase')['CONTSE'].to_dict())
    conciliacao_atratar['Lançou'] = conciliacao_atratar['Lançou'].fillna(0)

    # 1. Selecionar colunas com "d8" no nome e somar por linha (axis=1)
    # "D8 " precisa ficar com espaço para que a coluna "CONVENIO D8" não atrapalhe na hora da soma
    colunas_d8 = conciliacao_atratar.filter(regex=r'^(?!.*PRODUTO)D8').columns
    for col in colunas_d8:
        tipos = conciliacao_atratar[col].apply(type).value_counts()
        '''print(f"Coluna {col}:")
        print(tipos)
        print()'''
    conciliacao_atratar[colunas_d8] = conciliacao_atratar[colunas_d8].apply(pd.to_numeric, errors='coerce')

    soma_d8 = conciliacao_atratar.filter(regex=r'^(?!.*PRODUTO)D8').sum(axis=1)

    # 2. Calcular prestação * prazo
    prestacao_vezes_prazo = conciliacao_atratar['PRESTAÇÃO'] * conciliacao_atratar['PRAZO']

    # 3. Calcular o resultado final
    conciliacao_atratar['Pago'] = soma_d8 - prestacao_vezes_prazo
    conciliacao_atratar['Saldo'] = conciliacao_atratar['Pago'] + conciliacao_atratar['RECEBIDO GERAL']

    # Tira FUTURO, CANCELADOS, LIQUIDADOS, OBITO, E TÉRMINO DE CONTRATO
    conciliacao_tratado = conciliacao_atratar[
        ~conciliacao_atratar['BANCO'].str.contains('FUTURO')
        & ~conciliacao_atratar['ESTEIRA ATUAL'].str.contains('OBITO|CANCELADO|LIQUIDADO')
        & ~(conciliacao_atratar['Saldo'].fillna(float(-np.inf)) >= 0)
        ]


    return conciliacao_tratado


# --- Assumindo que essas funções já existem em outro lugar do seu código ---
# def separacao_conciliacao(credbase_trabalhado, conciliacao_bruto): ...
# def processar_alocacao_ade(dados, d8, produto): -> retorna (lista_ades, dataframe_modificado)
# def metodo_soma(dados, d8, pasta): ...
# def prepara_emprestimo(d8_restante, conciliacao_restante): ...

total_ades_usadas = []

def prepara_cartao(caminho_retorno_cartao: str,
                   caminho_d8_geral: str,
                   ):
    """
    Processa a alocação de ADEs para a modalidade 'CARTÃO', utilizando dois métodos
    e preparando os dados restantes para as próximas etapas.
    """
    print("--- INICIANDO ETAPA DE ALOCAÇÃO PARA CARTÃO ---")

    # ===================================================================
    # 1. SETUP E CARREGAMENTO DE DADOS (Sem alterações)
    # ===================================================================
    print("1. Carregando e preparando os dados iniciais...")

    d8_cartao_bruto_df = pd.read_csv(caminho_retorno_cartao, encoding="ISO-8859-1", sep=";", on_bad_lines="skip")
    conciliacao_base_df = separacao_conciliacao(credbase_trabalhado, conciliacao_bruto)


    colunas_relevantes = ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO', 'AVERBAÇÃO - ATUALIZADA', 'PRODUTO', 'Lançou']
    dados_cartao_para_alocar = (
        conciliacao_base_df.loc[conciliacao_base_df['Lançou'] == 1, colunas_relevantes]
        .sort_values(by=['CPF', 'AVERBAÇÃO - ATUALIZADA'], ascending=True)
        .copy()
    )

    # ===================================================================
    # 2. MÉTODO 1: ALOCAÇÃO DIRETA (Usando arquivo de retorno de cartão)
    # ===================================================================
    print("2. Executando Método 1: Alocação Direta para Cartão...")

    d8_cartao_df = d8_cartao_bruto_df.iloc[:, [2, 3, 7, 13]].copy()
    d8_cartao_df.columns = ["MATRÍCULA", "CPF", "PARCELA", "ADE"]

    # << --- ALTERAÇÃO AQUI --- >>
    # A chamada da função agora captura os dois valores retornados: a lista e o DataFrame.
    lista_ades_metodo_1, resultado_df_metodo_1 = processar_alocacao_ade(
        dados_cartao_para_alocar, d8_cartao_df, 'CARTÃO'
    )

    total_ades_usadas.extend(lista_ades_metodo_1)

    soma_exata()

def soma_exata():
    d8_ades_amenos = pd.read_csv(planilha_d8_geral, encoding="ISO-8859-1", sep=";", on_bad_lines="skip")
    d8_soma = d8_ades_amenos[~d8_ades_amenos['Contrato'].isin(total_ades_usadas)]

    colunas_relevantes = ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO', 'AVERBAÇÃO - ATUALIZADA', 'PRODUTO', 'Lançou']

    # prepara a conciliação separando apenas as colunas que vamos usar

    conciliacao_base_df = separacao_conciliacao(credbase_trabalhado, conciliacao_bruto)
    dados_cartao_para_alocar = (
        conciliacao_base_df.loc[conciliacao_base_df['Lançou'] == 0, colunas_relevantes]
        .sort_values(by=['CPF', 'AVERBAÇÃO - ATUALIZADA'], ascending=True)
        .copy()
    )

    # Recebe de volta o resultado da conciliação tratada, as ades utilizadas, e os arquivos gerados
    conciliacao_retorno_soma, ades_tulizadas_soma, files_list_soma = metodo_soma(dados_cartao_para_alocar, d8_soma, folder)
    # Coloca "vazio" nas linhas de ADE que não achou nada
    conciliacao_retorno_soma['ADE'] = conciliacao_retorno_soma['ADE'].fillna('')

    # Aumenta a lista de ADEs usadas para não correr o risco de utilizá-las novamente
    total_ades_usadas.extend(ades_tulizadas_soma)
    files_list.extend(files_list_soma)

    d8_para_contrato = d8_soma[~d8_soma['Contrato'].isin(total_ades_usadas)]
    conciliacao_para_contrato = conciliacao_retorno_soma[conciliacao_retorno_soma['ADE'] == '']

    # prepara_contratos(d8_ades_amenos, conciliacao_base_df)
    prepara_contratos(d8_para_contrato, conciliacao_para_contrato)


def prepara_contratos(d8_vindo_de_soma,
                      conciliacao_de_soma):
    """
    Usa os contratos encontrados para mapear ADEs em um novo conjunto de dados.

    Args:
        d8_vindo_de_soma: DataFrame D8 restante da etapa anterior.
        conciliacao_de_soma: DataFrame de conciliação restante da etapa anterior.

    Returns:
        (pd.DataFrame, pd.DataFrame, list): Retorna os DataFrames D8 e de conciliação
                                             prontos para a próxima etapa, e a lista
                                             total de ADEs utilizadas atualizada.
    """
    print("--- INICIANDO ETAPA DE MAPEAMENTO DE CONTRATOS PARA ADEs ---")

    # 1. ENCONTRAR CORRESPONDÊNCIAS
    # A função trata_contratos cria o mapeamento entre contratos "sujos" e "limpos"
    # O resultado é o D8 com as colunas 'Contrato_Encontrado_X'
    df_codigos_tratados = trata_contratos(d8_vindo_de_soma, conciliacao_de_soma, folder)
    # print(f'Comprimento do df_codigos_tratados: {len(df_codigos_tratados)}')
    # print(f'Comprimento do conciliacao_de_soma: {len(conciliacao_de_soma)}')

    # 2. CRIAR O MAPA: CONTRATO LIMPO -> ADE
    # Este mapa será usado para enriquecer outros dados
    print("Criando mapa de Contrato Limpo -> ADE...")
    mapa_contrato_para_ade = {}
    colunas_contratos = [col for col in df_codigos_tratados.columns if 'Contrato_Encontrado_' in col]

    # CORREÇÃO: Adicionado .iterrows()
    for _, row in df_codigos_tratados.iterrows():
        ade = row['Contrato']  # 'Contrato' no D8 é a ADE
        for col in colunas_contratos:
            contrato_encontrado = row.get(col)
            if pd.notna(contrato_encontrado):
                # Mapeia o contrato limpo (encontrado) para a sua ADE correspondente
                mapa_contrato_para_ade[str(contrato_encontrado).strip()] = ade

    # print(f'Comprimento do df_codigos_tratados: {len(df_codigos_tratados)}')

    # 3. APLICAR O MAPA EM OUTRO CONJUNTO DE DADOS
    # CORREÇÃO: Usa o DataFrame 'conciliacao_de_soma' que veio como argumento,
    # em vez de recarregar tudo.
    print("Aplicando mapa para encontrar ADEs nos contratos com 'Lançou == 0'...")
    print(f'Comprimento do conciliacao_de_soma == 1: {len(conciliacao_de_soma[conciliacao_de_soma['Lançou'] != 0])}')
    dados_para_alocar_ade = conciliacao_de_soma[conciliacao_de_soma['Lançou'] == 0].copy()
    # print(f'Comprimento do dados_para_alocar_ade: {len(dados_para_alocar_ade)}')

    # Usa o mapa para preencher a coluna 'ADE'
    dados_para_alocar_ade['ADE'] = dados_para_alocar_ade['CONTRATO'].astype(str).str.strip().map(mapa_contrato_para_ade)

    # 4. PREPARAR DADOS PARA A PRÓXIMA ETAPA
    print("Preparando dados para a próxima etapa (Empréstimo)...")

    # Atualiza a lista de ADEs utilizadas com as que acabamos de encontrar
    ades_encontradas_nesta_etapa = dados_para_alocar_ade['ADE'].dropna().unique().tolist()
    total_ades_usadas.extend(ades_encontradas_nesta_etapa)

    # Filtra o D8 para a próxima função, removendo as ADEs utilizadas
    d8_para_emprestimo = d8_vindo_de_soma[~d8_vindo_de_soma['Contrato'].isin(total_ades_usadas)]

    # CORREÇÃO: Filtra os contratos que AINDA não têm ADE usando .isna()
    conciliacao_para_emprestimo = dados_para_alocar_ade[dados_para_alocar_ade['ADE'].isna()]

    dados_para_alocar_ade.to_excel(fr'{folder}\Dados de Averbação - Contratos achados.xlsx', index=False)

    prepara_emprestimo(d8_para_emprestimo, conciliacao_para_emprestimo)

def prepara_emprestimo(geral_d8, conciliacao_soma_exata):
    d8_geral = geral_d8
    d8_geral_reduzido = d8_geral.loc[d8_geral['Serviço'] == 'Empréstimo Consignado',['Matrícula', 'CPF', 'Nome', 'Contrato', 'Serviço', 'Valor original']]

    d8_geral_colunas_novas = d8_geral_reduzido.rename(columns={'Matrícula': 'MATRÍCULA', 'Nome': 'NOME', 'Contrato': 'ADE', 'Valor original': 'PARCELA'})

    conciliacao = conciliacao_soma_exata.loc[conciliacao_soma_exata['PRODUTO'] == 'Empréstimo', ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO', 'AVERBAÇÃO - ATUALIZADA', 'PRODUTO', 'Lançou']].copy()

    # total_ades_usadas = []

    ades_usadas_cartao = processar_alocacao_ade(conciliacao, d8_geral_colunas_novas, 'EMPRÉSTIMO')

    total_ades_usadas.extend(ades_usadas_cartao)

    # Filtra o d8_geral_df UMA ÚLTIMA VEZ para garantir que temos o estado mais recente
    d8_final_restante = d8_geral[~d8_geral['Contrato'].isin(total_ades_usadas)]
    prepara_beneficio(d8_final_restante, conciliacao_soma_exata)

def prepara_beneficio(geral_d8, conciliacao_soma_exata):
    d8_geral = geral_d8
    d8_geral_reduzido = d8_geral.loc[d8_geral['Serviço'] == 'Cartão Benefício',['Matrícula', 'CPF', 'Nome', 'Contrato', 'Serviço', 'Valor original']]

    d8_geral_colunas_novas = d8_geral_reduzido.rename(columns={'Matrícula': 'MATRÍCULA', 'Nome': 'NOME', 'Contrato': 'ADE', 'Valor original': 'PARCELA'})

    conciliacao = conciliacao_soma_exata.loc[
            conciliacao_soma_exata['PRODUTO'] == 'Cartão Benefício', ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO', 'AVERBAÇÃO - ATUALIZADA', 'PRODUTO']].copy()

    # total_ades_usadas = []


    ades_usadas_cartao = processar_alocacao_ade(conciliacao, d8_geral_colunas_novas, 'CARTÃO BENEFÍCIO')

    total_ades_usadas.extend(ades_usadas_cartao)

    # Filtra o d8_geral_df UMA ÚLTIMA VEZ para garantir que temos o estado mais recente
    d8_final_restante = d8_geral[~d8_geral['Contrato'].isin(total_ades_usadas)]
    prepara_cartao_nao_lancado(d8_final_restante, conciliacao_soma_exata)


def prepara_cartao_nao_lancado(geral_d8, conciliacao_soma_exata):
    d8_geral = geral_d8
    d8_geral_reduzido = d8_geral.loc[d8_geral['Serviço'] == 'Cartão de Crédito',['Matrícula', 'CPF', 'Nome', 'Contrato', 'Serviço', 'Valor original']]

    d8_geral_colunas_novas = d8_geral_reduzido.rename(columns={'Matrícula': 'MATRÍCULA', 'Nome': 'NOME', 'Contrato': 'ADE', 'Valor original': 'PARCELA'})

    conciliacao = conciliacao_soma_exata.loc[
            conciliacao_soma_exata['PRODUTO'] == 'Cartão de Crédito', ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO',
                                                           'AVERBAÇÃO - ATUALIZADA', 'PRODUTO', 'Lançou']].copy()
    # total_ades_usadas = []

    mask_credito_nao_lancado = (conciliacao['PRODUTO'] == 'Cartão de Crédito') & (conciliacao['Lançou'] != 1)
    conciliacao_encontrados = conciliacao.loc[mask_credito_nao_lancado]


    ades_usadas_cartao = processar_alocacao_ade(conciliacao_encontrados, d8_geral_colunas_novas, 'CARTÃO NÃO LANÇADO')

    total_ades_usadas.extend(ades_usadas_cartao)

    # Filtra o d8_geral_df UMA ÚLTIMA VEZ para garantir que temos o estado mais recente
    d8_final_restante = d8_geral[~d8_geral['Contrato'].isin(total_ades_usadas)]

    prepara_resto(d8_final_restante, conciliacao_soma_exata)

def prepara_resto(geral_d8, conciliacao_soma_exata):
    d8_geral = geral_d8
    d8_geral_reduzido = d8_geral.loc[
        d8_geral['Serviço'] == 'Cartão de Crédito', ['Matrícula', 'CPF', 'Nome', 'Contrato', 'Serviço',
                                                     'Valor original']]

    d8_geral_colunas_novas = d8_geral_reduzido.rename(
        columns={'Matrícula': 'MATRÍCULA', 'Nome': 'NOME', 'Contrato': 'ADE', 'Valor original': 'PARCELA'})

    conciliacao = conciliacao_soma_exata
    # total_ades_usadas = []

    conciliacao_encontrados = conciliacao.loc[conciliacao['PRODUTO'] == '-',
    ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO', 'AVERBAÇÃO - ATUALIZADA', 'PRODUTO', 'Lançou']]

    ades_usadas_cartao = processar_alocacao_ade(conciliacao_encontrados, d8_geral_colunas_novas, 'RESTO')

    total_ades_usadas.extend(ades_usadas_cartao)

    # Junta todos os DataFrames
    concatena_resultados()

def processar_alocacao_ade(dados, d8, modalidade):
    """
    Aloca ADEs da planilha d8 para a planilha dados com base no CPF e valor da parcela,
    gerenciando o saldo de cada ADE.
    """
    try:
        df_dados = dados.copy()
        df_d8 = d8.copy()
    except FileNotFoundError:
        print(f"Erro: O arquivo '{dados}' ou '{d8}' não foi encontrado.")
        return
    except ValueError as e:
        print(
            f"Erro ao ler as planilhas. Verifique se os nomes '{dados}' e '{d8}' estão corretos. Detalhe: {e}")
        return

    # Garante que os tipos de dados estejam corretos para evitar erros
    df_dados['CPF'] = df_dados['CPF'].astype(str)
    df_d8['CPF'] = df_d8['CPF'].astype(str)

    # Garante que parcela em df_d8 é número
    df_d8['PARCELA'] = df_d8['PARCELA'].astype(str).str.replace(".", "")
    df_d8['PARCELA'] = df_d8['PARCELA'].astype(str).str.replace(",", ".")
    df_d8['PARCELA'] = df_d8['PARCELA'].astype(float)
    # df_d8['PARCELA'] = pd.to_numeric(df_d8['PARCELA'], errors='coerce')

    # Agrupa os dados da d8 por CPF para fácil acesso
    # Cria um dicionário: {cpf: [{ade: X, parcela: Y, saldo: Y}, ...]}
    d8_agrupado = {}
    for _, row in df_d8.iterrows():
        cpf = row['CPF']
        if cpf not in d8_agrupado:
            d8_agrupado[cpf] = []
        d8_agrupado[cpf].append({
            'ADE': row['ADE'],
            'parcela_original': row['PARCELA'],
            'saldo': row['PARCELA']  # Saldo inicial é o valor total da parcela
        })

    # print(d8_agrupado['004.436.613-24'])

    # Lista para armazenar os resultados
    dados_resultado = []

    # Dicionário para rastrear o índice da ADE atual para cada CPF
    ade_tracker = {cpf: 0 for cpf in d8_agrupado.keys()}

    # NOVO: Conjunto para guardar as ADEs que foram de fato utilizadas
    ades_utilizadas = set()

    # Itera sobre cada linha da planilha "dados"
    for _, row in df_dados.iterrows():
        cpf = row['CPF']
        parcela_a_cobrir = row['PRESTAÇÃO']

        if cpf not in d8_agrupado:
            # CPF não existe na d8, adiciona a linha com um aviso
            nova_linha = row.to_dict()
            nova_linha['ADE'] = 'CPF não encontrado em d8'
            # nova_linha['ADE'] = row['CONTRATO']
            dados_resultado.append(nova_linha)
            continue

        ades_disponiveis = d8_agrupado[cpf]

        while parcela_a_cobrir > 0.01:  # Tolerância para ponto flutuante
            idx_ade_atual = ade_tracker.get(cpf, 0)

            if idx_ade_atual >= len(ades_disponiveis):
                # Não há mais ADEs para cobrir o valor
                linha_artificial = row.to_dict()
                linha_artificial['PRESTAÇÃO'] = parcela_a_cobrir  # Valor que faltou
                # linha_artificial['ADE'] = 'ADEs insuficientes em d8'
                linha_artificial['ADE'] = row['CONTRATO']
                dados_resultado.append(linha_artificial)
                break

            ade_atual = ades_disponiveis[idx_ade_atual]

            # Valor que pode ser coberto por esta ADE
            valor_a_alocar = min(parcela_a_cobrir, ade_atual['saldo'])

            if valor_a_alocar > 0.01:
                linha_artificial = row.to_dict()
                # A parcela na nova linha é o valor que foi efetivamente coberto
                linha_artificial['PRESTAÇÃO'] = valor_a_alocar
                linha_artificial['ADE'] = ade_atual['ADE']
                dados_resultado.append(linha_artificial)

                # NOVO: Adiciona a ADE à nossa lista de utilizadas
                ades_utilizadas.add(ade_atual['ADE'])

                # Atualiza os saldos e valores
                ade_atual['saldo'] -= valor_a_alocar
                parcela_a_cobrir -= valor_a_alocar

            # Se o saldo da ADE zerou, move para a próxima
            if ade_atual['saldo'] < 0.01:
                ade_tracker[cpf] += 1

    # Cria o DataFrame final a partir da lista de resultados
    df_resultado = pd.DataFrame(dados_resultado)

    # Salva o resultado em uma nova planilha no mesmo arquivo Excel
    nome_arquivo = fr"{folder}\Dados de Averbacao Tratado {modalidade}.xlsx"
    df_resultado.to_excel(nome_arquivo, index=False)
    files_list.append(nome_arquivo)
    d8.to_excel(fr'{folder}\D8 ROBO.xlsx', index=False)

    print(f"Processo  de {modalidade} concluído com sucesso! Verifique a planilha 'Dados de Averbacao Tratado {modalidade}' no seu arquivo.")

    # NOVO: Retorna a lista de ADEs e o df_resultado
    return list(ades_utilizadas), df_resultado

# Funcao que concatena todos os arquivos
def concatena_resultados():
    df_averbacao_unificada = pd.concat([pd.read_excel(arquivo) for arquivo in files_list], ignore_index=True)

    # É necessário tirar as células vazias da coluna 'ADE'
    df_averbacao_unificada['ADE'] = df_averbacao_unificada['ADE'].fillna('')
    df_averbacao_unificada_sem_vazios = df_averbacao_unificada[df_averbacao_unificada['ADE'] != '']

    print(f"Tá vazio? {df_averbacao_unificada.loc[38364, 'ADE'] == ''}")

    df_averbacao_unificada_sem_vazios.to_excel(fr'{folder}\DADOS DE AVERBAÇÃO UNIFICADAS SEM VAZIOS.xlsx', index=False)
    df_averbacao_unificada.to_excel(fr'{folder}\DADOS DE AVERBAÇÃO UNIFICADAS.xlsx', index=False)


# Executa a função principal
if __name__ == "__main__":
    prepara_cartao(planilha_d8, planilha_d8_geral)