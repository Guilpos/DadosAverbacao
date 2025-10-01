import pandas as pd
import openpyxl
import numpy as np
import tkinter as tk
from tkinter import filedialog

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
# credbase_trabalhado = r"Z:\Python\DadosAverbacao\geral\CREDBASE TRABALHADO UNIFICADO GOV MA 09.2025.xlsx"
credbase_trabalhado = selecionar_com_validacao(r"Selecione o Credbase Trabalhado", 'xlsx')

# Conciliação são os dados averbados
# conciliacao_bruto = r"Z:\Python\DadosAverbacao\geral\Conciliação-Governo do Maranhão - 092025.xlsx"
conciliacao_bruto = selecionar_com_validacao(r"Selecione o arquivo de Conciliação bruto", "xlsx")

# D8 Geral do convenio
# planilha_d8_geral = r"Z:\Python\DadosAverbacao\geral\MA_D8 GERAL.CSV"
planilha_d8_geral = selecionar_com_validacao(r"Selecione a planilha geral de D8", "csv")

# D8 apenas para os casos de cartão
# planilha_d8 = r"Z:\Python\DadosAverbacao\geral\D8 GOV MA.csv"
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
    colunas_d8 = conciliacao_atratar.filter(like='D8 ').columns
    for col in colunas_d8:
        tipos = conciliacao_atratar[col].apply(type).value_counts()
        '''print(f"Coluna {col}:")
        print(tipos)
        print()'''
    conciliacao_atratar[colunas_d8] = conciliacao_atratar[colunas_d8].apply(pd.to_numeric, errors='coerce')

    soma_d8 = conciliacao_atratar.filter(like='D8 ').sum(axis=1)

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


def prepara_cartao(retorno, d8_tudo):
    print('Iniciando...')
    df_retorno = pd.read_csv(retorno, encoding="ISO-8859-1",sep=";", on_bad_lines="skip")

    d8_geral = pd.read_csv(d8_tudo, encoding="ISO-8859-1",sep=";", on_bad_lines="skip")

    conciliacao = separacao_conciliacao(credbase_trabalhado, conciliacao_bruto)

    conciliacao_contratos_encontrados = conciliacao.loc[conciliacao['Lançou'] == 1,
    ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO', 'AVERBAÇÃO - ATUALIZADA']]

    conciliacao_contratos_encontrados = conciliacao_contratos_encontrados.sort_values(by=['CPF', 'AVERBAÇÃO - ATUALIZADA'], ascending=True)

    total_ades_usadas = []

    # Seleciona as colunas desejadas pelos índices
    retorno_filtrado = df_retorno.iloc[:, [2, 3, 7, 13]]

    # Renomeia as colunas
    retorno_filtrado.columns = ["MATRÍCULA", "CPF", "PARCELA", "ADE"]

    conciliacao_contratos_encontrados.to_excel(fr'{folder}\CONCILIACAO_TESTE.xlsx', index=False)

    ades_usadas_cartao = processar_alocacao_ade(conciliacao_contratos_encontrados, retorno_filtrado, 'CARTÃO')
    total_ades_usadas.extend(ades_usadas_cartao)

    # Filtra o d8_geral_df UMA ÚLTIMA VEZ para garantir que temos o estado mais recente
    d8_final_restante = d8_geral[~d8_geral['Contrato'].isin(total_ades_usadas)]
    prepara_emprestimo(d8_final_restante)

def prepara_emprestimo(geral_d8):
    d8_geral = geral_d8
    d8_geral_reduzido = d8_geral.loc[d8_geral['Serviço'] == 'Empréstimo Consignado',['Matrícula', 'CPF', 'Nome', 'Contrato', 'Serviço', 'Valor original']]

    d8_geral_colunas_novas = d8_geral_reduzido.rename(columns={'Matrícula': 'MATRÍCULA', 'Nome': 'NOME', 'Contrato': 'ADE', 'Valor original': 'PARCELA'})

    conciliacao = separacao_conciliacao(credbase_trabalhado, conciliacao_bruto)

    total_ades_usadas = []

    conciliacao_encontrados = conciliacao.loc[conciliacao['PRODUTO'] == 'Empréstimo',
    ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO', 'AVERBAÇÃO - ATUALIZADA']]

    ades_usadas_cartao = processar_alocacao_ade(conciliacao_encontrados, d8_geral_colunas_novas, 'EMPRÉSTIMO')

    total_ades_usadas.extend(ades_usadas_cartao)

    # Filtra o d8_geral_df UMA ÚLTIMA VEZ para garantir que temos o estado mais recente
    d8_final_restante = d8_geral[~d8_geral['Contrato'].isin(total_ades_usadas)]
    prepara_beneficio(d8_final_restante)

def prepara_beneficio(geral_d8):
    d8_geral = geral_d8
    d8_geral_reduzido = d8_geral.loc[d8_geral['Serviço'] == 'Cartão Benefício',['Matrícula', 'CPF', 'Nome', 'Contrato', 'Serviço', 'Valor original']]

    d8_geral_colunas_novas = d8_geral_reduzido.rename(columns={'Matrícula': 'MATRÍCULA', 'Nome': 'NOME', 'Contrato': 'ADE', 'Valor original': 'PARCELA'})

    conciliacao = separacao_conciliacao(credbase_trabalhado, conciliacao_bruto)
    total_ades_usadas = []

    conciliacao_encontrados = conciliacao.loc[conciliacao['PRODUTO'] == 'Cartão Benefício',
    ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO', 'AVERBAÇÃO - ATUALIZADA']]

    ades_usadas_cartao = processar_alocacao_ade(conciliacao_encontrados, d8_geral_colunas_novas, 'CARTÃO BENEFÍCIO')

    total_ades_usadas.extend(ades_usadas_cartao)

    # Filtra o d8_geral_df UMA ÚLTIMA VEZ para garantir que temos o estado mais recente
    d8_final_restante = d8_geral[~d8_geral['Contrato'].isin(total_ades_usadas)]
    prepara_cartao_nao_lancado(d8_final_restante)


def prepara_cartao_nao_lancado(geral_d8):
    d8_geral = geral_d8
    d8_geral_reduzido = d8_geral.loc[d8_geral['Serviço'] == 'Cartão de Crédito',['Matrícula', 'CPF', 'Nome', 'Contrato', 'Serviço', 'Valor original']]

    d8_geral_colunas_novas = d8_geral_reduzido.rename(columns={'Matrícula': 'MATRÍCULA', 'Nome': 'NOME', 'Contrato': 'ADE', 'Valor original': 'PARCELA'})

    conciliacao = separacao_conciliacao(credbase_trabalhado, conciliacao_bruto)
    total_ades_usadas = []

    mask_credito_nao_lancado = (conciliacao['PRODUTO'] == 'Cartão de Crédito') & (conciliacao['Lançou'] != 1)
    conciliacao_encontrados = conciliacao.loc[mask_credito_nao_lancado, ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO', 'AVERBAÇÃO - ATUALIZADA']]


    ades_usadas_cartao = processar_alocacao_ade(conciliacao_encontrados, d8_geral_colunas_novas, 'CARTÃO NÃO LANÇADO')

    total_ades_usadas.extend(ades_usadas_cartao)

    # Filtra o d8_geral_df UMA ÚLTIMA VEZ para garantir que temos o estado mais recente
    d8_final_restante = d8_geral[~d8_geral['Contrato'].isin(total_ades_usadas)]

    prepara_resto(d8_final_restante)

def prepara_resto(geral_d8):
    d8_geral = geral_d8
    d8_geral_reduzido = d8_geral.loc[
        d8_geral['Serviço'] == 'Cartão de Crédito', ['Matrícula', 'CPF', 'Nome', 'Contrato', 'Serviço',
                                                     'Valor original']]

    d8_geral_colunas_novas = d8_geral_reduzido.rename(
        columns={'Matrícula': 'MATRÍCULA', 'Nome': 'NOME', 'Contrato': 'ADE', 'Valor original': 'PARCELA'})

    conciliacao = separacao_conciliacao(credbase_trabalhado, conciliacao_bruto)
    total_ades_usadas = []

    conciliacao_encontrados = conciliacao.loc[conciliacao['PRODUTO'] == '-',
    ['CONTRATO', 'CPF', 'NOME', 'PRESTAÇÃO', 'AVERBAÇÃO - ATUALIZADA']]

    ades_usadas_cartao = processar_alocacao_ade(conciliacao_encontrados, d8_geral_colunas_novas, 'RESTO')

    total_ades_usadas.extend(ades_usadas_cartao)

    # Filtra o d8_geral_df UMA ÚLTIMA VEZ para garantir que temos o estado mais recente
    d8_final_restante = d8_geral[~d8_geral['Contrato'].isin(total_ades_usadas)]

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
    df_resultado.to_excel(fr"{folder}\Dados de Averbacao Tratado {modalidade}.xlsx", index=False)
    d8.to_excel(fr'{folder}\D8 ROBO.xlsx', index=False)

    print(f"Processo  de {modalidade} concluído com sucesso! Verifique a planilha 'Dados de Averbacao Tratado {modalidade}' no seu arquivo.")

    # NOVO: Retorna a lista de ADEs
    return list(ades_utilizadas)

# Executa a função principal
if __name__ == "__main__":
    prepara_cartao(planilha_d8, planilha_d8_geral)
