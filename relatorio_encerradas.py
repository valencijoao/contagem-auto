import os
import pandas as pd
from datetime import datetime

PASTA_RAIZ = r"C:\Users\tec01\Desktop\Sites"
dados_carinha = []
arquivos_desconhecidos = [] # Lista para o log

for endereco_atual, subpastas, arquivos in os.walk(PASTA_RAIZ):
    for carinha in arquivos:
        nome_low = carinha.lower()
        
        if nome_low.endswith('.pdf'):
            caminho_completo = os.path.join(endereco_atual, carinha)
            
            try:
                ano_arquivo = datetime.fromtimestamp(os.path.getmtime(caminho_completo)).year
            except:
                continue
            
            # Focamos apenas em 2025
            if ano_arquivo == 2025:
                rel = os.path.relpath(endereco_atual, PASTA_RAIZ)
                partes_c = rel.split(os.sep)
                nome_portal = partes_c[0] if len(partes_c) > 0 else "Raiz"
                nome_cliente = partes_c[1] if len(partes_c) > 1 else "Sem Cliente"
                
                identificador = None
                
                # --- REGRAS DE EXTRAÇÃO ---
                
                # 1. Caso Pregão
                if 'pregao' in nome_low or 'pregão' in nome_low:
                    id_pot = nome_low.replace('pregão', 'pregao').split('pregao')[1].replace('.pdf', '').strip()
                    if id_pot.replace('_', '').isdigit():
                        identificador = id_pot
                
                # 2. Caso LCT
                elif 'lct' in nome_low:
                    id_pot = nome_low.split('lct')[1].replace('.pdf', '').strip()
                    if id_pot.isdigit():
                        identificador = id_pot
                
                # 3. Caso Underscore (Exato 3 partes)
                elif '_' in nome_low:
                    nome_sem_ext = nome_low.replace('.pdf', '')
                    partes = nome_sem_ext.split('_')
                    if len(partes) == 3:
                        id_pot = partes[1].strip()
                        if id_pot.isdigit():
                            identificador = id_pot

                # --- LÓGICA DO LOG ---
                if identificador:
                    dados_carinha.append({
                        'Portal': nome_portal,
                        'Cliente': nome_cliente,
                        'Arquivo': carinha,
                        'ID': identificador
                    })
                else:
                    # Se for de 2025 e não encaixou nas regras, vai para o log
                    # Ignoramos apenas arquivos que já sabemos ser editais (_ac_, _edital, etc)
                    if '_ac_' not in nome_low and '_edital' not in nome_low:
                        arquivos_desconhecidos.append(f"Portal: {nome_portal} | Arquivo: {carinha}")

# 1. Gerar o Excel
if dados_carinha:
    df = pd.DataFrame(dados_carinha)
    df = df.drop_duplicates(subset=['Portal', 'Cliente', 'ID'])
    df.sort_values(by=['Portal', 'Cliente']).to_excel('relatorio_2025.xlsx', index=False)
    print(f"Relatório gerado: {len(dados_carinha)} arquivos.")

# 2. Gerar o Log de Conferência Manual
with open('log_conferencia_manual.txt', 'w', encoding='utf-8') as f:
    if arquivos_desconhecidos:
        f.write("--- ARQUIVOS DE 2025 COM PADRÃO DESCONHECIDO ---\n")
        f.write("Estes arquivos não entraram no Excel e podem precisar de nova regra:\n\n")
        for linha in arquivos_desconhecidos:
            f.write(linha + "\n")
        print(f"Atenção: {len(arquivos_desconhecidos)} arquivos no log para conferência.")
    else:
        f.write("Nenhum padrão desconhecido encontrado para 2025.")