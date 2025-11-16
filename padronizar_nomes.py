import pandas as pd
import json
import os


def carregar_dicionario_padronizacao(caminho_json):
    """Carrega um dicionário de padronização de um arquivo JSON."""
    try:
        with open(caminho_json, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Erro: Arquivo JSON não encontrado em '{caminho_json}'")
        return None
    except json.JSONDecodeError:
        print(f"Erro: Falha ao decodificar JSON em '{caminho_json}'. Verifique a sintaxe.")
        return None

def padronizar_coluna(df, coluna_nome, dicionario_padronizacao):
    """
    Aplica a padronização a uma coluna de um DataFrame.
    As chaves do dicionário devem estar prontas para serem comparadas com o valor original da coluna.
    O valor padronizado será o do dicionário.
    """
    if coluna_nome in df.columns and dicionario_padronizacao is not None:
        def mapear_valor(valor):
            if pd.isna(valor): 
                return valor
            str_valor = str(valor)
            if str_valor in dicionario_padronizacao:
                return dicionario_padronizacao[str_valor]
            elif str_valor.lower() in dicionario_padronizacao:
                return dicionario_padronizacao[str_valor.lower()]
            return valor

        df[coluna_nome] = df[coluna_nome].apply(mapear_valor)
    else:
        if coluna_nome not in df.columns:
            print(f"Aviso: Coluna '{coluna_nome}' não encontrada na planilha.")
    return df



def padronizar_e_sobrescrever_planilhas(pasta_alvo, json_portais, json_clientes):
    """
    Processa todos os arquivos Excel na pasta alvo, padroniza as colunas
    'Portal' e 'Cliente' e SOBRESCREVE os arquivos originais.

    Args:
        pasta_alvo (str): O caminho para a pasta contendo os arquivos Excel a serem modificados.
        json_portais (str): O caminho para o arquivo JSON de padronização de portais.
        json_clientes (str): O caminho para o arquivo JSON de padronização de clientes.
    """
    print(f"ATENÇÃO: Este script irá sobrescrever os arquivos Excel originais na pasta '{pasta_alvo}'.")
    print("Certifique-se de ter um backup se necessário.")

    dicionario_portais = carregar_dicionario_padronizacao(json_portais)
    dicionario_clientes = carregar_dicionario_padronizacao(json_clientes)

    if dicionario_portais is None or dicionario_clientes is None:
        print("Não foi possível carregar um ou ambos os dicionários de padronização. Abortando.")
        return

    for nome_arquivo in os.listdir(pasta_alvo):
        if nome_arquivo.endswith(('.xlsx', '.xls')):
            caminho_completo_arquivo = os.path.join(pasta_alvo, nome_arquivo)

            print(f"\nProcessando e sobrescrevendo: {nome_arquivo}")
            try:

                excel_file = pd.ExcelFile(caminho_completo_arquivo)
                todas_as_abas = {}

                for aba_nome in excel_file.sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=aba_nome)


                    df = padronizar_coluna(df, 'Portal', dicionario_portais)


                    df = padronizar_coluna(df, 'Cliente', dicionario_clientes)

                    todas_as_abas[aba_nome] = df


                with pd.ExcelWriter(caminho_completo_arquivo, engine='xlsxwriter') as writer:
                    for aba_nome, df_modificado in todas_as_abas.items():
                        df_modificado.to_excel(writer, sheet_name=aba_nome, index=False)
                print(f"Arquivo '{nome_arquivo}' padronizado e sobrescrito com sucesso.")

            except Exception as e:
                print(f"Erro ao processar o arquivo {nome_arquivo}: {e}")




PASTA_DAS_PLANILHAS = r'C:\Users\tec01\Desktop\pasta_de_trabalho\contagem_pdf\retornos'
JSON_PORTAIS = 'padrao_portais.json'
JSON_CLIENTES = 'padrao_clientes.json'


if __name__ == "__main__":
    padronizar_e_sobrescrever_planilhas(PASTA_DAS_PLANILHAS, JSON_PORTAIS, JSON_CLIENTES)