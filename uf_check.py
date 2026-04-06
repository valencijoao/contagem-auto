import pandas as pd
import unicodedata

# 1. Função de normalização (essencial para ignorar acentos e maiúsculas)
def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    nfkd_form = unicodedata.normalize('NFKD', texto)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).lower().strip()

# 2. Dicionário de Mapeamentos (Adicione novos clientes aqui)
mapeamentos = {
    "tremont": [
        "Afonso Cláudio", "Água Doce do Norte", "Águia Branca", "Alegre", "Alfredo Chaves", 
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
    ],
    "auremar": ["Adamantina", "Adolfo", "Aguaí", "Águas da Prata", "Águas de Lindoia", "Águas de Santa Bárbara", 
        "Águas de São Pedro", "Agudos", "Alambari", "Alfredo Marcondes", "Altair", "Altinópolis", 
        "Alto Alegre", "Alumínio", "Álvares Florence", "Álvares Machado", "Álvaro de Carvalho", 
        "Alvinlândia", "Americana", "Américo Brasiliense", "Américo de Campos", "Amparo", "Analândia", 
        "Andradina", "Angatuba", "Anhembi", "Anhumas", "Aparecida", "Aparecida d'Oeste", "Apiaí", 
        "Araçariguama", "Araçatuba", "Araçoiaba da Serra", "Aramina", "Arandu", "Arapeí", "Araraquara", 
        "Araras", "Arco-Íris", "Arealva", "Areias", "Areiópolis", "Ariranha", "Artur Nogueira", "Arujá", 
        "Aspásia", "Assis", "Atibaia", "Auriflama", "Avaí", "Avanhandava", "Avaré", "Bady Bassitt", 
        "Balbinos", "Bálsamo", "Bananal", "Barão de Antonina", "Barbosa", "Bariri", "Barra Bonita", 
        "Barra do Chapéu", "Barra do Turvo", "Barretos", "Barrinha", "Barueri", "Bastos", "Batatais", 
        "Bauru", "Bebedouro", "Bento de Abreu", "Bernardino de Campos", "Bertioga", "Bilac", "Birigui", 
        "Biritiba Mirim", "Boa Esperança do Sul", "Bocaina", "Bofete", "Boituva", "Bom Jesus dos Perdões", 
        "Bom Sucesso de Itararé", "Borá", "Boraceia", "Borborema", "Borebi", "Botucatu", "Bragança Paulista", 
        "Braúna", "Brejo Alegre", "Brodowski", "Brotas", "Buri", "Buritama", "Buritizal", "Cabrália Paulista", 
        "Cabreúva", "Caçapava", "Cachoeira Paulista", "Caconde", "Caiabu", "Caieiras", "Caiuá", "Cajamar", 
        "Cajati", "Cajobi", "Cajuru", "Campina do Monte Alegre", "Campinas", "Campo Limpo Paulista", 
        "Campos do Jordão", "Campos Novos Paulista", "Cananeia", "Canas", "Cândido Mota", "Cândido Rodrigues", 
        "Canitar", "Capão Bonito", "Capela do Alto", "Capivari", "Caraguatatuba", "Carapicuíba", "Cardoso", 
        "Casa Branca", "Cássia dos Coqueiros", "Castilho", "Catanduva", "Catiguá", "Catobi", "Cedral", 
        "Cerqueira César", "Cerquilho", "Cesário Lange", "Charqueada", "Chavantes", "Clementina", "Colina", 
        "Colômbia", "Conchal", "Conchas", "Cordeirópolis", "Coroados", "Coronel Macedo", "Corumbataí", 
        "Cosmópolis", "Cosmorama", "Cotia", "Cristais Paulista", "Cruzália", "Cruzeiro", "Cubatão", 
        "Cunha", "Descalvado", "Diadema", "Dirce Reis", "Divinolândia", "Dobrada", "Dois Córregos", 
        "Dolcinópolis", "Dourado", "Dracena", "Duartina", "Dumont", "Echaporã", "Eldorado", "Elias Fausto", 
        "Elisiário", "Embu das Artes", "Embu-Guaçu", "Emilianópolis", "Engenheiro Coelho", "Espírito Santo do Pinhal", 
        "Espírito Santo do Turvo", "Estrela d'Oeste", "Estrela do Norte", "Euclides da Cunha Paulista", 
        "Fartura", "Fernando Prestes", "Fernandópolis", "Fernão", "Ferraz de Vasconcelos", "Flora Rica", 
        "Floreal", "Flórida Paulista", "Florínea", "Franca", "Francisco Morato", "Franco da Rocha", 
        "Gabriel Monteiro", "Gália", "Garça", "Gastão Vidigal", "Gavião Peixoto", "General Salgado", 
        "Getulina", "Glicério", "Guaiçara", "Guaimbê", "Guaíra", "Guapiaçu", "Guapiara", "Guará", "Guaraçaí", 
        "Guaraci", "Guarani d'Oeste", "Guarantã", "Guararapes", "Guararema", "Guaratinguetá", "Guareí", 
        "Guariba", "Guarujá", "Guarulhos", "Guatapará", "Guzolândia", "Herculândia", "Holambra", "Hortolândia", 
        "Iacanga", "Iacri", "Iaras", "Ibaté", "Ibirá", "Ibirarema", "Ibitinga", "Ibiúna", "Icém", "Iepê", 
        "Igaraçu do Tietê", "Igarapava", "Igaratá", "Iguape", "Ilhabela", "Ilha Comprida", "Ilha Solteira", 
        "Indaiatuba", "Indiana", "Indiaporã", "Inúbia Paulista", "Ipaussu", "Iperó", "Ipeúna", "Ipiguá", 
        "Iporanga", "Ipuã", "Iracemápolis", "Irapuã", "Irapuru", "Itaberá", "Itaí", "Itajobi", "Itaju", 
        "Itanhaém", "Itaoca", "Itapecerica da Serra", "Itapetininga", "Itapeva", "Itapevi", "Itapira", 
        "Itapirapuã Paulista", "Itápolis", "Itaporanga", "Itapuí", "Itapura", "Itaquaquecetuba", "Itararé", 
        "Itariri", "Itatiba", "Itatinga", "Itirapina", "Itirapuã", "Itobi", "Itu", "Itupeva", "Ituverava", 
        "Jaborandi", "Jaboticabal", "Jacareí", "Jaci", "Jacupiranga", "Jaguariúna", "Jales", "Jambeiro", 
        "Jandira", "Jardinópolis", "Jarinu", "Jaú", "Jeriquara", "Joanópolis", "João Ramalho", 
        "José Bonifácio", "Júlio Mesquita", "Jumirim", "Jundiaí", "Junqueirópolis", "Juquiá", "Juquitiba", 
        "Lagoinha", "Laranjal Paulista", "Lavínia", "Lavrinhas", "Leme", "Lençóis Paulista", "Limeira", 
        "Lindoia", "Lins", "Lorena", "Lourdes", "Louveira", "Lucélia", "Lucianópolis", "Luís Antônio", 
        "Luiziânia", "Lupércio", "Lutécia", "Macatuba", "Macaubal", "Macedônia", "Magda", "Mairinque", 
        "Mairiporã", "Manduri", "Marabá Paulista", "Maracaí", "Marapoama", "Mariápolis", "Marília", 
        "Marinópolis", "Martinópolis", "Matão", "Mauá", "Mendonça", "Meridiano", "Mesópolis", "Miguelópolis", 
        "Mineiros do Tietê", "Mira Estrela", "Miracatu", "Mirante do Paranapanema", "Mirassol", 
        "Mirassolândia", "Mococa", "Mogi das Cruzes", "Mogi Guaçu", "Mogi Mirim", "Mombuca", "Monções", 
        "Mongaguá", "Monte Alegre do Sul", "Monte Alto", "Monte Aprazível", "Monte Azul Paulista", 
        "Monte Castelo", "Monte Mor", "Monteiro Lobato", "Morungaba", "Morro Agudo", "Motuca", "Murutinga do Sul", 
        "Nantes", "Narandiba", "Natividade da Serra", "Nazaré Paulista", "Neves Paulista", "Nhandeara", 
        "Nipoã", "Nova Aliança", "Nova Campina", "Nova Canaã Paulista", "Nova Castilho", "Nova Europa", 
        "Nova Granada", "Nova Guataporanga", "Nova Independência", "Nova Luzitânia", "Nova Odessa", 
        "Novais", "Novo Horizonte", "Nuporanga", "Ocauçu", "Óleo", "Olímpia", "Onda Verde", "Oriente", 
        "Orindiúva", "Orlândia", "Osasco", "Oscar Bressane", "Osvaldo Cruz", "Ourinhos", "Ouro Verde", 
        "Ouroeste", "Pacaembu", "Palestina", "Palmares Paulista", "Palmeira d'Oeste", "Palmital", 
        "Panorama", "Paraguaçu Paulista", "Paraibuna", "Paraíso", "Paranapanema", "Paranapuã", "Parapuã", 
        "Pardinho", "Pariquera-Açu", "Parisi", "Patrocínio Paulista", "Pauliceia", "Paulínia", "Paulistânia", 
        "Paulo de Faria", "Pederneiras", "Pedra Bela", "Pedranópolis", "Pedregulho", "Pedreira", 
        "Pedrinhas Paulista", "Pedro de Toledo", "Penápolis", "Pereira Barreto", "Pereiras", "Peruíbe", 
        "Piacatu", "Piedade", "Pilar do Sul", "Pindamonhangaba", "Pindorama", "Pinhalzinho", "Piquerobi", 
        "Piquete", "Piracaia", "Piracicaba", "Piraju", "Pirajuí", "Pirangi", "Pirapora do Bom Jesus", 
        "Pirapozinho", "Pirassununga", "Piratininga", "Pitangueiras", "Planalto", "Platina", "Poá", 
        "Poloni", "Pompeia", "Pongaí", "Pontal", "Pontalinda", "Pontes Gestal", "Populina", "Porangaba", 
        "Porto Feliz", "Porto Ferreira", "Potim", "Potirendaba", "Pracinha", "Pradópolis", "Praia Grande", 
        "Pratânia", "Presidente Alves", "Presidente Bernardes", "Presidente Epitácio", "Presidente Prudente", 
        "Presidente Venceslau", "Promissão", "Quadra", "Quatá", "Queiroz", "Queluz", "Quintana", "Rafard", 
        "Rancharia", "Redenção da Serra", "Regente Feijó", "Reginópolis", "Registro", "Restinga", 
        "Ribeira", "Ribeirão Bonito", "Ribeirão Branco", "Ribeirão Corrente", "Ribeirão do Sul", 
        "Ribeirão dos Índios", "Ribeirão Grande", "Ribeirão Pires", "Ribeirão Preto", "Riversul", 
        "Rifaina", "Rincão", "Rinópolis", "Rio Claro", "Rio das Pedras", "Rio Grande da Serra", 
        "Riolândia", "Rosana", "Roseira", "Rubiácea", "Rubineia", "Sabino", "Sagres", "Sales", 
        "Sales Oliveira", "Salesópolis", "Salmourão", "Saltinho", "Salto", "Salto de Pirapora", 
        "Salto Grande", "Sandovalina", "Santa Adélia", "Santa Albertina", "Santa Bárbara d'Oeste", 
        "Santa Branca", "Santa Clara d'Oeste", "Santa Cruz da Conceição", "Santa Cruz da Esperança", 
        "Santa Cruz do Rio Pardo", "Santa Ernestina", "Santa Fé do Sul", "Santa Gertrudes", 
        "Santa Isabel", "Santa Lúcia", "Santa Maria da Serra", "Santa Mercedes", "Santa Rita d'Oeste", 
        "Santa Rita do Passa Quatro", "Santa Rosa de Viterbo", "Santa Salete", "Santana da Ponte Pensa", 
        "Santana de Parnaíba", "Santo Anastácio", "Santo André", "Santo Antônio da Alegria", 
        "Santo Antônio de Posse", "Santo Antônio do Aracanguá", "Santo Antônio do Jardim", 
        "Santo Antônio do Pinhal", "Santo Expedito", "Santópolis do Aguapeí", "Santos", "São Bento do Sapucaí", 
        "São Bernardo do Campo", "São Caetano do Sul", "São Carlos", "São Francisco", "São João da Boa Vista", 
        "São João das Duas Pontes", "São João de Iracema", "São João do Pau-d'Alho", "São Joaquim da Barra", 
        "São José da Bela Vista", "São José do Barreiro", "São José do Rio Pardo", "São José do Rio Preto", 
        "São José dos Campos", "São Lourenço da Serra", "São Luiz do Paraitinga", "São Manuel", "São Miguel Arcanjo", 
        "São Paulo", "São Pedro", "São Pedro do Turvo", "São Roque", "São Sebastião", "São Sebastião da Grama", 
        "São Vicente", "Sarapuí", "Sarutaiá", "Sebastianópolis do Sul", "Serra Azul", "Serra Negra", 
        "Serrana", "Sertãozinho", "Sete Barras", "Severínia", "Silveiras", "Socorro", "Sorocaba", 
        "Sud Mennucci", "Sumaré", "Suzanápolis", "Suzano", "Tabapuã", "Tabatinga", "Taboão da Serra", 
        "Taciba", "Taguaí", "Taiaçu", "Taiúva", "Tambaú", "Tanabi", "Tapiraí", "Tapiratiba", "Taquaral", 
        "Taquaritinga", "Taquarituba", "Taquarivaí", "Tarabai", "Tarumã", "Tatuí", "Taubaté", "Tejupá", 
        "Teodoro Sampaio", "Terra Roxa", "Tietê", "Timburi", "Torre de Pedra", "Torrinha", "Trabiju", 
        "Tremembé", "Três Fronteiras", "Tuiuti", "Tupã", "Tupi Paulista", "Turiúba", "Turmalina", "Ubarana", 
        "Ubatuba", "Ubirajara", "Uchoa", "União Paulista", "Urânia", "Uru", "Urupês", "Valentim Gentil", 
        "Valinhos", "Valparaíso", "Vargem", "Vargem Grande do Sul", "Vargem Grande Paulista", "Várzea Paulista"],
    "santos extintores":[],
    "rodobens":[] 
}

# 3. Interface de seleção
print("--- SISTEMA DE FILTRAGEM ---")
print("Clientes disponíveis:", ", ".join(mapeamentos.keys()))
cliente_escolhido = input("Digite o nome do cliente para a checagem: ").strip().lower()

if cliente_escolhido not in mapeamentos:
    print(f"Erro: Cliente '{cliente_escolhido}' não encontrado.")
else:
    # 4. Preparação da lista do cliente escolhido
    lista_referencia = [normalizar_texto(m) for m in mapeamentos[cliente_escolhido]]
    padrao_regex = '|'.join([rf"\b{m}\b" for m in lista_referencia])

    # 5. Leitura e Processamento
    try:
        arquivo = 'UFs_template.xlsx'
        coluna_alvo = 'Orgão'
        
        df = pd.read_excel(arquivo)
        
        # Cria coluna temporária para busca sem alterar a original
        df['coluna_busca'] = df[coluna_alvo].apply(normalizar_texto)
        
        # Filtra (Busca Parcial)
        df_filtrado = df[df['coluna_busca'].str.contains(padrao_regex, na=False, regex=True)].copy()
        
        # Remove a coluna de suporte
        df_filtrado = df_filtrado.drop(columns=['coluna_busca'])
        
        print(f"\nForam encontrados {len(df_filtrado)} registros para o cliente '{cliente_escolhido}'.")
        print(df_filtrado.head())
        

    except Exception as e:
        print(f"Ocorreu um erro ao processar a planilha: {e}")