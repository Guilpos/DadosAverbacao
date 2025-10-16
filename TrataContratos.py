import pandas as pd
import numpy as np
import openpyxl
import re
from thefuzz import fuzz # Importe a biblioteca no início do seu script

rejeitados = ['/']


def trata_contratos(d8_do_main, conciliacao_do_main, caminho):
    d8_reduzido = d8_do_main[
        ['Matrícula', 'CPF', 'Nome', 'Contrato', 'Cod na Instituição', 'Serviço', 'Valor original']].copy()
    conciliacao = conciliacao_do_main.copy()

    def extrair_contratos_com_referencia(df_sujo: pd.DataFrame, df_limpo: pd.DataFrame) -> pd.DataFrame:
        print("Iniciando o processo de extração de contratos...")

        # Função de limpeza (pode ser definida aqui ou fora)
        def limpar_contrato(texto: str) -> str:
            if not isinstance(texto, str):
                texto = str(texto)
            return re.sub(r'[^0-9a-zA-Z]', '', texto)  # Mantém letras e números

        # --- Passo 1: Criar o mapa de referência (sem alterações) ---
        df_limpo['CONTRATO'] = df_limpo['CONTRATO'].astype(str).str.strip()
        print("Criando mapa de referência CPF -> Contratos...")
        mapa_cpf_contratos = df_limpo.groupby('CPF')['CONTRATO'].apply(list).to_dict()

        # --- Passo 2: Definir a função que será aplicada em cada linha (LÓGICA ALTERADA) ---
        def encontrar_contratos_na_linha(row):
            cpf = row['CPF']
            texto_contratos_sujo = str(row['Cod na Instituição'])

            contratos_validos_para_cpf = mapa_cpf_contratos.get(cpf, [])
            if not contratos_validos_para_cpf:
                return []

            # << --- NOVA LÓGICA DE DIVISÃO --- >>
            # 1. DIVIDIR: Quebra o campo sujo em partes usando múltiplos separadores.
            #    Filtra strings vazias que podem surgir de separadores duplos (ex: '123//456')
            partes_sujas = [p for p in re.split(r'[/,;\s-]+', texto_contratos_sujo) if p]

            if not partes_sujas:
                return []

            encontrados_nesta_linha = []
            # Cria uma cópia da lista de contratos válidos para poder remover itens já encontrados
            contratos_disponiveis = list(contratos_validos_para_cpf)
            LIMIAR_DE_SIMILARIDADE = 70

            # 2. CONQUISTAR: Para cada parte suja, encontra o melhor match
            for parte in partes_sujas:
                parte_limpa = limpar_contrato(parte)
                if not parte_limpa:
                    continue

                melhor_match_para_parte = None
                maior_score = 0

                for contrato_valido in contratos_disponiveis:
                    contrato_valido_limpo = limpar_contrato(contrato_valido)

                    # Usa a comparação por similaridade
                    score = fuzz.ratio(parte_limpa, contrato_valido_limpo)

                    if score > LIMIAR_DE_SIMILARIDADE and score > maior_score:
                        maior_score = score
                        melhor_match_para_parte = contrato_valido

                # Se um bom match foi encontrado para esta parte...
                if melhor_match_para_parte:
                    # Adiciona à lista de resultados da linha
                    encontrados_nesta_linha.append(melhor_match_para_parte)
                    # Remove da lista de disponíveis para não ser encontrado novamente na mesma linha
                    contratos_disponiveis.remove(melhor_match_para_parte)

            # 3. JUNTAR: Retorna a lista com todos os matches encontrados
            return encontrados_nesta_linha

        # --- Passo 3: Aplicar a função e criar as novas colunas (sem alterações) ---
        print("Analisando a Planilha A e extraindo os contratos...")
        df_sujo['Cod na Instituição'] = df_sujo['Cod na Instituição'].astype(str).str.replace('nan', '')
        lista_de_contratos_encontrados = df_sujo.apply(encontrar_contratos_na_linha, axis=1)

        df_contratos_novos = pd.DataFrame(lista_de_contratos_encontrados.tolist(), index=df_sujo.index)
        df_contratos_novos.columns = [f'Contrato_Encontrado_{i + 1}' for i in df_contratos_novos.columns]

        df_resultado = pd.concat([df_sujo, df_contratos_novos], axis=1)

        print("Processo concluído com sucesso!")
        df_resultado.to_excel(fr'{caminho}\D8 Contratos tratados.xlsx', index=False)
        return df_resultado

    # Chama a função principal com os dataframes preparados
    df_codigos_tratados = extrair_contratos_com_referencia(d8_reduzido, conciliacao)
    return df_codigos_tratados
