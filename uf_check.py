import pandas as pd
import unicodedata

# 1. Função para normalizar strings (remove acentos, minúsculas e espaços extras)
def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    # Remove acentos
    nfkd_form = unicodedata.normalize('NFKD', texto)
    texto_sem_acento = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    return texto_sem_acento.lower().strip()

# 2. Lista completa dos municípios do ES
map_tremont = [
    "Afonso Cláudio", "Água Docedo Norte", "Águia Branca", "Alegre", "Alfredo Chaves", 
    "Alto Rio Novo", "Anchieta", "Apiacá", "Aracruz", "Atílio Vivácqua", 
    "Baixo Guandu", "Barra de São Francisco", "Boa Esperança", "Bom Jesus do Norte", "Brejetuba", 
    "Cachoeiro de Itapemirim", "Cariacica", "Castelo", "Colatina", "Conceição da Barra", 
    "Conceição do Castelo", "Divino de São Lourenço", "Domingos Martins", "Dores do Rio Preto", "Ecoporanga", 
    "Fundão", "Governador Lindenberg", "Guaçuí", "Guarapari", "Ibatiba", 
    "Ibiraçu", "Ibitirama", "Iconha", "Irupi", "Itaguaçu", 
    "Itapemirim", "Itarana", "Iúna", "Jaguaré", "Jerônimo Monteiro", 
    "João Neiva", "Laranja da Terra", "Linhares", "Mantenópolis", "Marataízes", 
    "Marechal Floriano", "Marilândia", "Mimoso do Sul", "Montanha", "Mucurici", 
    "Muniz Freire", "Muqui", "Nova Venécia", "Pancas", "Pedro Canário", 
    "Pinheiros", "Piúma", "Ponto Belo", "Presidente Kennedy", "Rio Bananal", 
    "Rio Novo do Sul", "Santa Leopoldina", "Santa Maria de Jetibá", "Santa Teresa", "São Domingos do Norte", 
    "São Gabriel da Palha", "São José do Calçado", "São Mateus", "São Roque do Canaã", "Serra", 
    "Sooretama", "Vargem Alta", "Venda Nova do Imigrante", "Viana", "Vila Pavão", 
    "Vila Valério", "Vila Velha", "Vitória"
]




municipios_normalizados = [normalizar_texto(m) for m in map_tremont]
arquivo = 'UFs_Tremont.xlsx' 
coluna_alvo = 'orgao_comprador'

df = pd.read_excel(arquivo)

# 4. Criar coluna temporária normalizada para busca
df['coluna_busca'] = df[coluna_alvo].apply(normalizar_texto)

# 5. Filtrar usando Regex (Busca Parcial)
# Isso identifica "Vitória" dentro de "Município de Vitória"
padrao_regex = '|'.join([rf"\b{m}\b" for m in municipios_normalizados])

df_filtrado = df[df['coluna_busca'].str.contains(padrao_regex, na=False, regex=True)].copy()

# 6. Limpeza e Resultado
df_filtrado = df_filtrado.drop(columns=['coluna_busca'])

print(f"Foram encontrados {len(df_filtrado)} registros do ES.")
print(df_filtrado.head())

# Opcional: Salvar o resultado
# df_filtrado.to_excel('resultado_es.xlsx', index=False)