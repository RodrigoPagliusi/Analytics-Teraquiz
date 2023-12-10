##################################################################################### INSTALA AS BIBLIOTECAS

import os
import subprocess

print('Verificando os pacotes necessários...')

# Verifica se os pacotes necessários estão instalados
# Instala os que não estiverem
try: subprocess.run(["pip3", "--version"])
except FileNotFoundError: os.system("python get-pip.py") ; os.system("pip install --upgrade pip")
try: import PyInstaller
except ModuleNotFoundError: os.system("python -m pip3 install pyinstaller")
try: import requests
except ModuleNotFoundError: os.system("pip3 install requests")
try: import dateutil
except ModuleNotFoundError: os.system("pip3 install python-dateutil")
try: import xlsxwriter
except ModuleNotFoundError: os.system("pip3 install xlsxwriter")
try: import numpy
except ModuleNotFoundError: os.system("pip3 install numpy")
try: import matplotlib
except ModuleNotFoundError: os.system("pip3 install -U matplotlib")

print('Tudo instalado!...\n')

##################################################################################### IMPORTA BIBLIOTECAS E DEFINE PATHS

import json
import requests
from dateutil import parser
from dateutil.relativedelta import relativedelta
import xlsxwriter
import statistics as st
from collections import Counter
from datetime import datetime
import numpy as np
import pprint as pp
import matplotlib.pyplot as plt

# Define aonde serão criados o arquivo excel e a pasta contendo todos os gráficos.
path_script = os.path.abspath(__file__) # Do jeito que está, eles serão criados na mesma pasta em que está este programa
path_excel = os.path.dirname(path_script)
path_figuras = os.path.dirname(path_script)

print('Lendo os dados...')

##################################################################################### LEITURA DOS DADOS

# Lê o arquivo com os dados em formato json e coloca eles como um dicionário em uma variável
with open('response.json','r', encoding='utf_8') as data_json: data_original = json.load(data_json)
print('Os dados foram lidos com sucesso...')

##################################################################################### DEFINE OS NOMES DE CADA VARIÁVEL DOS DADOS E OS PADRÕES DE FORMATAÇÃO

# Dados de datas possuem formatação especial
nomes_dados_datas = ["data_nascimento","data_app"]

# Chaves dos dados originais à esquerda e chaveS novas à direita. A ordem dos nomes aqui define a ordem que aparecerá no excel.
dados_reais_renom_reordenados = {
"Id":"Id",
"email":"email",
"phone":"telefone",
"firstName":"nome",
"lastName":"sobrenome",
"birthdate":nomes_dados_datas[0],
"non_original_idade":"idade",
"country":"pais",
"state":"estado",
"city":"cidade",
"cep":"cep",
"jobId":"especialidade",
"yearStartResidence":"ano_comeco_residencia",
"non_original_tempo_especialidade":"tempo_especialidade",
"institution":"instituicao",
"invite":"convite",
"confirmedAt":nomes_dados_datas[1]
    }

# Essas listas servem para dar tratamentos diferentes a diferentes tipos de dados, são usadas mais adiante no código
str_usuario_ativo = 'user_active'
job_id = "jobId"
criados_no_script = ["non_original_idade","non_original_tempo_especialidade"]
possivel_ni =  ["phone","firstName","lastName","country","state","city","cep","institution","invite"]
tratar_datas = ["birthdate","yearStartResidence","confirmedAt"]

# Formatação e cores gerais, independente da área
branco = "#FFFFFF"
cinza = "#333F48"
fonte = "Verdana"
tamanho_colunas_excel = 18

# Variáveis dos dados comuns a todas as áreas
dados_areas = {
    "time_all":"Tempo de Uso Total",
    "time_home":"Tempo Home",
    "time_quiz":"Tempo Quiz",
    "time_video":"Tempo Video",
    "time_community":"Tempo Comunidade",
    "click_link":"Click Link",
    "share_newsletter":"Compartilhou Newsletter",
    "click_newsletter":"Click Newsletter",
    "click_caso_clinico":"Click Caso Clinico",
    'usuario_ativo':"Usuário Ativo"
    }

# Nome de cada área com a sua respectiva cor
# Ao introzudir novas áreas, colocar elas acima
areas = {
    "onco":"#3FB1E5",
    "reuma":"#00EB80",
    "radio":"#FF7F32",
    "teraquiz":"#12233D"
    }

# Cores secundárias de cada área
cores_mais_claras = {
    "onco":"#7FF5F9",
    "reuma":"#AAFFE3",
    "radio":"#FFBF76",
    "teraquiz":"#45566F"
    }

# Como os nomes das áreas aparecerão nos gráficos
titulos_areas = ['Onco','Reumato','Radio','Teraquiz']

### Informações para mudanças futuras
# Incluir hermato e uro e Dermato 
# Radio - Entrada +4
# Onco - Entrada +3
# Reuma - Entrada+2
# Hermato -  Entrada +2
# Uro -  Entrada +3
# Dermato - Entrada +3

# Código e nome de todos jobs ou especialidades possíveis nos dados
legenda_jobs_id = {
    3:'Estudante de Medicina',
    4:'Médico não radioterapeuta',
    5:'Enfermeiro',
    6:'Tecnólogo',
    7:'Dosimetrista',
    8:'Físico Médico',
    10:'Radioterapeuta',
    11:'Oncologista',
    12:'Residente de Oncologia',
    13:'Mastologista',
    14:'Radioterapeuta',
    15:'Reumatologista',
    16:'Cirurgião',
    17:'Médico Outras Especialidades',
    18:'Hematologista',
    19:'Urologista',
    20:'Dermatologista'
    }

##################################################################################### VARIÁVEIS INICIAIS

# Números e strings necessários que se repetem com frequência no código
divisao_horas = 60*60 # Os tempos estão em segundos, com essa divisão se obtém em horas
tempo_ativo = 60 # 60 segundos = 1 minutos que é o tempo total de uso necessário para ser considerado ativo em cada área
str_user_info = "str_user_info"
chave_user = "user"
nao_informado = "N.I."
especialidades = ['Especialista', 'Residente', nao_informado]

# Listas dos dados originais e lista dos dados novos
dados_usuarios_originais = list(dados_reais_renom_reordenados.keys())
dados_usuarios_novos = list(dados_reais_renom_reordenados.values())
dados_originais_areas = list(dados_areas.keys())
dados_novos_areas = list(dados_areas.values())

# Serve para se obter o Job Id, independente da ordem do dicionário 'dados_reais_renom_reordenados'
indice_job_id = dados_usuarios_originais.index(job_id)

# nomes de cada área e da cor no formato RRGGBB de cada área
nomes_areas = list(areas.keys())
cores_areas = list(areas.values())

# Cria o arquivo xlsx e renomeia a planilha
arquivo_excel = xlsxwriter.Workbook(path_excel + "/Dados_Teraquiz.xlsx")
planilha = arquivo_excel.add_worksheet("Dados_Teraquiz")

##################################################################################### FORMATA O EXCEL

# Formatação das cédulas dos usuários no excel
formato_cedula_geral = arquivo_excel.add_format({"font":fonte, "font_color":cinza, "border":1, "bg_color":branco})
planilha.set_column("A:XFD", tamanho_colunas_excel, formato_cedula_geral)

# Formatação das cédulas de cada área
for pref_num in range(len(nomes_areas)):
    if pref_num == len(nomes_areas) - 1: globals()["formato_cedula_" + nomes_areas[pref_num]] = arquivo_excel.add_format({"font":fonte, "font_color":branco, "border":1, "bg_color":cores_areas[pref_num]})
    else: globals()["formato_cedula_" + nomes_areas[pref_num]] = arquivo_excel.add_format({"font":fonte, "font_color":cinza, "border":1, "bg_color":cores_areas[pref_num]})
    planilha.set_column(len(dados_usuarios_novos) + len(dados_originais_areas)*(pref_num), len(dados_usuarios_novos) + len(dados_originais_areas)*(pref_num+1) - 1, tamanho_colunas_excel, globals()["formato_cedula_" + nomes_areas[pref_num]])

# Formatação as colunas de datas
formato_cedula_data = arquivo_excel.add_format({"font":fonte, "bg_color":branco, "font_color":cinza, "border":1, "num_format":"dd/mm/yy hh:mm"})
for nome_data in nomes_dados_datas: planilha.set_column(dados_usuarios_novos.index(nome_data), dados_usuarios_novos.index(nome_data), tamanho_colunas_excel, formato_cedula_data)

# Escreve os títulos das colunas na planilha
## Títulos dos usuários
for numero_de_usuario in range(len(dados_usuarios_novos)):
    if dados_usuarios_novos[numero_de_usuario] in nomes_dados_datas: planilha.write(0, numero_de_usuario, dados_usuarios_novos[numero_de_usuario], formato_cedula_geral)
    else: planilha.write(0, numero_de_usuario, dados_usuarios_novos[numero_de_usuario])

## Títulos das áreas
auxiliar_contagem = 0
for area in nomes_areas:
    for dado in dados_originais_areas:
        col_areas = len(dados_usuarios_novos) + auxiliar_contagem
        planilha.write(0, col_areas, dado + "_" + area)
        auxiliar_contagem += 1

##################################################################################### COLOCA TODOS OS DADOS DE MANEIRA ORGANIZADA NUMA LISTA

# Nesta lista serão colocados os dicts de todos os usuários. Cada dict contém os dados de um usuário.
data_formatado = []

# Loop nos dados originais
for num_data_ori in range(len(data_original)):

    if num_data_ori < 2: continue

    registro_original = data_original[num_data_ori]

    # Cada dict_usuario conterá os dados de um usuário
    dict_usuario = {}
    dict_usuario[str_user_info] = {}
    for num_usuario in range(len(dados_usuarios_novos)):

        # Há dados que serão criados a partir dos existentes e não precisam ter nenhuma ação feita no momento
        if dados_usuarios_originais[num_usuario] in criados_no_script: continue

        # Verificar se determinados dados são nulos ou não representam nenhuma informação
        elif dados_usuarios_originais[num_usuario] in possivel_ni:
            if registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user][dados_usuarios_originais[num_usuario]] == ""\
            or registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user][dados_usuarios_originais[num_usuario]] == "Nao sei"\
            or registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user][dados_usuarios_originais[num_usuario]] == "-":
                dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]] = nao_informado
            else: dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]] = registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user][dados_usuarios_originais[num_usuario]]

        # Pega o job_id certo, que pode ser 'jobId' ou 'job_id' dos dados originais e usa o dict para obter a especilidade
        elif dados_usuarios_originais[num_usuario] == job_id:
            num_jobId = registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user]["jobId"]
            num_job_id = registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user]["job_id"]
            if num_jobId != 0: dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]] = legenda_jobs_id[num_jobId]
            else: dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]] = legenda_jobs_id[num_job_id]

        # Verificações necessárias nos dados de datas
        ## Data de nascimento
        elif dados_usuarios_originais[num_usuario] == tratar_datas[0]:
            dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]] = parser.isoparse(registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user][dados_usuarios_originais[num_usuario]]).replace(tzinfo=None)
            dict_usuario[str_user_info][dados_reais_renom_reordenados[criados_no_script[0]]] = relativedelta(datetime.now(),dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]]).years

        ## Ano de começo de residência
        elif dados_usuarios_originais[num_usuario] == tratar_datas[1]:
            # Define os tempos de residência necessários para ser considerado Especialista de acordo com a área
            if dict_usuario[str_user_info][dados_usuarios_novos[indice_job_id]] == legenda_jobs_id[10] or dict_usuario[str_user_info][dados_usuarios_novos[indice_job_id]] == legenda_jobs_id[14]: tempo_especialista = 4
            elif dict_usuario[str_user_info][dados_usuarios_novos[indice_job_id]] == legenda_jobs_id[11] or dict_usuario[str_user_info][dados_usuarios_novos[indice_job_id]] == legenda_jobs_id[19]\
            or dict_usuario[str_user_info][dados_usuarios_novos[indice_job_id]] == legenda_jobs_id[20]: tempo_especialista = 3
            elif dict_usuario[str_user_info][dados_usuarios_novos[indice_job_id]] == legenda_jobs_id[18] or dict_usuario[str_user_info][dados_usuarios_novos[indice_job_id]] == legenda_jobs_id[15]: tempo_especialista = 2
            else: tempo_especialista = 5

            # Verifica se o usuário possui tempo de residência necessário para ser especialista
            try:
                dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]] = int(registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user][dados_usuarios_originais[num_usuario]])
                if datetime.today().year - dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]] >= tempo_especialista: dict_usuario[str_user_info][dados_reais_renom_reordenados[criados_no_script[1]]] = especialidades[0]
                else: dict_usuario[str_user_info][dados_reais_renom_reordenados[criados_no_script[1]]] = especialidades[1]
            # Se não estiver informado o ano de começo de residência, esses dados serão 'nao_informado' = 'N.I.'
            except:
                dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]] = nao_informado
                dict_usuario[str_user_info][dados_reais_renom_reordenados[criados_no_script[1]]] = nao_informado

        ## Data em que o usuário entrou no app
        elif dados_usuarios_originais[num_usuario] == tratar_datas[2]:
            if registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user][dados_usuarios_originais[num_usuario]] == "0001-01-01T00:00:00Z": dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]] = nao_informado
            elif "." in registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user][dados_usuarios_originais[num_usuario]]: 
                dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]] = parser.isoparse(registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user][dados_usuarios_originais[num_usuario]].split(".")[0] + "Z").replace(tzinfo=None)
            else: dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]] = parser.isoparse(registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user][dados_usuarios_originais[num_usuario]]).replace(tzinfo=None)

        ## Se o dado não precisa de nenhuma verificação, ele é apenas inserido no dict
        else: dict_usuario[str_user_info][dados_usuarios_novos[num_usuario]] = str(registro_original[nomes_areas[-2] + nomes_areas[-1]][chave_user][dados_usuarios_originais[num_usuario]])

    # Insere dados de cada área no dict
    for area in nomes_areas[:-1]:
        dict_usuario[area + nomes_areas[-1]] = {}
        for dado in dados_originais_areas[:-1]:
            dict_usuario[area + nomes_areas[-1]][dado] = registro_original[area + nomes_areas[-1]][dado]
        if dict_usuario[area + nomes_areas[-1]][dados_originais_areas[0]] >= tempo_ativo: dict_usuario[area + nomes_areas[-1]][dados_originais_areas[-1]] = True
        else: dict_usuario[area + nomes_areas[-1]][dados_originais_areas[-1]] = False

    # Insere dados somados de todas as áreas no dict, na área teraquiz, que representa o aplicativo como um todo
    dict_usuario[nomes_areas[-1]] = {}
    for dado in dados_originais_areas[:-1]:
        dict_usuario[nomes_areas[-1]][dado] = 0
        for area in nomes_areas[:-1]: dict_usuario[nomes_areas[-1]][dado] += registro_original[area + nomes_areas[-1]][dado]
    if dict_usuario[nomes_areas[-1]][dados_originais_areas[0]] >= tempo_ativo: dict_usuario[nomes_areas[-1]][dados_originais_areas[-1]] = True
    else: dict_usuario[nomes_areas[-1]][dados_originais_areas[-1]] = False

    # Insere o dict de cada usuário na lista
    data_formatado.append(dict_usuario)

##################################################################################### INSERE OS DADOS NO EXCEL

# Insere todos os dados
for num_data_form in range(len(data_formatado)):

    ## Insere todos os dados dos usuários
    for numero_de_usuario in range(len(list(data_formatado[num_data_form][str_user_info].values()))):
        planilha.write(num_data_form + 1, numero_de_usuario, list(data_formatado[num_data_form][str_user_info].values())[numero_de_usuario])

    ## Insere dados de todas as áreas
    auxiliar_contagem = 0
    for area in nomes_areas:
        for dado in dados_originais_areas:
            col_areas = len(dados_usuarios_novos) + auxiliar_contagem
            if area == nomes_areas[-1]: planilha.write(num_data_form + 1, col_areas, data_formatado[num_data_form][area][dado])
            else: planilha.write(num_data_form + 1, col_areas, data_formatado[num_data_form][area + nomes_areas[-1]][dado])
            auxiliar_contagem += 1

arquivo_excel.close() # É necessário close() o arquivo_excel para salvar as mudanças

print('Os dados foram transformados com sucesso...')

##################################################################################### TRANSFORMAÇÕES E ESTATÍSTICAS

# Loop para produzir as listas personalizadas e agregações dos dados em cada área
for area in nomes_areas:
    nome_adicional = nomes_areas[-1]
    if area == nomes_areas[-1]: nome_adicional = ''

    # Criar as listas vazias que serão preenchidas
    globals()['media_tempo_mes_' + area + nome_adicional] = []
    globals()['media_tempo_semana_' + area + nome_adicional] = []
    globals()['regiao' + '_' + area + nome_adicional] = []
    globals()['tempo_especialidade' + '_' + area + nome_adicional] = []
    for dado in dados_originais_areas: globals()[dado + '_' + area + nome_adicional] = []


    for user in range(len(data_formatado)):
        # Preenche as listas com os dados de todas as áreas
        for dado in dados_originais_areas: globals()[dado + '_' + area + nome_adicional].append(data_formatado[user][area + nome_adicional][dado])

        # Calcula as médias de tempo por mês e por semana de cada usuário
        if data_formatado[user][str_user_info][dados_reais_renom_reordenados["confirmedAt"]] == 'N.I.': # Se não houver a dados_graficos de criação da conta do usuário, não tem fazer os cálculos abaixo
            globals()['media_tempo_mes_' + area + nome_adicional].append(data_formatado[user][str_user_info][dados_reais_renom_reordenados["confirmedAt"]])
            globals()['media_tempo_semana_' + area + nome_adicional].append(data_formatado[user][str_user_info][dados_reais_renom_reordenados["confirmedAt"]])
        else: 
            meses_criacao_conta = relativedelta(datetime.now(),data_formatado[user][str_user_info][dados_reais_renom_reordenados["confirmedAt"]]).months + 1
            globals()['media_tempo_mes_' + area + nome_adicional].append(round(data_formatado[user][area + nome_adicional][dados_originais_areas[0]]/meses_criacao_conta,0))
            semanas_criacao_conta = relativedelta(datetime.now(),data_formatado[user][str_user_info][dados_reais_renom_reordenados["confirmedAt"]]).weeks + 1
            globals()['media_tempo_semana_' + area + nome_adicional].append(round(data_formatado[user][area + nome_adicional][dados_originais_areas[0]]/semanas_criacao_conta,0))

        # São adquiridos os dados de usuários e tempo de uso por região
        # Também são adquiridas as listas de usuários que são residentes, especialistas ou não informados ('N.I.')
        if area != nomes_areas[-1]:
            if data_formatado[user][area + nome_adicional][dados_originais_areas[-1]] == True:
                if data_formatado[user][str_user_info][dados_reais_renom_reordenados["country"]] == 'Brasil':
                    globals()['regiao' + '_' + area + nome_adicional].append(data_formatado[user][str_user_info][dados_reais_renom_reordenados["state"]])

                else: globals()['regiao' + '_' + area + nome_adicional].append(data_formatado[user][str_user_info][dados_reais_renom_reordenados["country"]])

                globals()['tempo_especialidade' + '_' + area + nome_adicional].append(data_formatado[user][str_user_info][dados_reais_renom_reordenados["non_original_tempo_especialidade"]])

            else: globals()['regiao' + '_' + area + nome_adicional].append('Not_In_Area') ; globals()['tempo_especialidade' + '_' + area + nome_adicional].append('Not_In_Area')

        else:
            if data_formatado[user][str_user_info][dados_reais_renom_reordenados["country"]] == 'Brasil': globals()['regiao' + '_' + area + nome_adicional].append(data_formatado[user][str_user_info][dados_reais_renom_reordenados["state"]])
            else: globals()['regiao' + '_' + area + nome_adicional].append(data_formatado[user][str_user_info][dados_reais_renom_reordenados["country"]])
            globals()['tempo_especialidade' + '_' + area + nome_adicional].append(data_formatado[user][str_user_info][dados_reais_renom_reordenados["non_original_tempo_especialidade"]])

    # Agregações da média de tempo de uso mensal
    para_agregacoes_soma_media_tempo_mes = list(filter(lambda x: x != nao_informado, globals()['media_tempo_mes_' + area + nome_adicional]))
    globals()['soma_media_tempo_mes_' + area + nome_adicional] = round(sum(para_agregacoes_soma_media_tempo_mes) / divisao_horas, 1)
    globals()['media_media_tempo_mes_' + area + nome_adicional] = round(st.mean(para_agregacoes_soma_media_tempo_mes) / divisao_horas*60, 1)
    globals()['mediana_media_tempo_mes_' + area + nome_adicional] = round(st.median(para_agregacoes_soma_media_tempo_mes) / divisao_horas*60, 1)
    globals()['maximo_media_tempo_mes_' + area + nome_adicional] = round(max(para_agregacoes_soma_media_tempo_mes) / divisao_horas, 1)

    # Agregações de todos os dados de tempo de todas de cada área
    for dado in dados_originais_areas[:-1]:
        globals()['para_calculos_' + dado + '_' + area + nome_adicional] = list(filter(lambda x: x > tempo_ativo, globals()[dado + '_' + area + nome_adicional]))
        if len(globals()['para_calculos_' + dado + '_' + area + nome_adicional]) == 0: 
            globals()['soma_' + dado + '_' + area + nome_adicional] = 0
            globals()['media_' + dado + '_' + area + nome_adicional] = 0
            globals()['mediana_' + dado + '_' + area + nome_adicional] = 0
            globals()['maximo_' + dado + '_' + area + nome_adicional] = 0
        else:
            globals()['soma_' + dado + '_' + area + nome_adicional] = round(sum(globals()['para_calculos_' + dado + '_' + area + nome_adicional]) / divisao_horas, 1)
            globals()['media_' + dado + '_' + area + nome_adicional] = round(st.mean(globals()['para_calculos_' + dado + '_' + area + nome_adicional]) / divisao_horas*60, 1)
            globals()['mediana_' + dado + '_' + area + nome_adicional] = round(st.median(globals()['para_calculos_' + dado + '_' + area + nome_adicional]) / divisao_horas*60, 1)
            globals()['maximo_' + dado + '_' + area + nome_adicional] = round(max(globals()['para_calculos_' + dado + '_' + area + nome_adicional]) / divisao_horas, 1)

    # Número de usuários ativos e inativos. E suas porcentagens.
    globals()['numero_' + dados_originais_areas[-1] + '_' + area + nome_adicional] = dict(Counter(globals()[dados_originais_areas[-1] + '_' + area + nome_adicional]))
    globals()['porcent_ativos_' + dados_originais_areas[-1] + '_' + area + nome_adicional] = round(Counter(globals()[dados_originais_areas[-1] + '_' + area + nome_adicional])[True] / len(globals()[dados_originais_areas[-1] + '_' + area + nome_adicional]),2)
    globals()['porcent_inativos_' + dados_originais_areas[-1] + '_' + area + nome_adicional] = round(Counter(globals()[dados_originais_areas[-1] + '_' + area + nome_adicional])[False] / len(globals()[dados_originais_areas[-1] + '_' + area + nome_adicional]),2)

    # Número de usuários especialistas, residentes ou não informados. E suas porcentagens.
    globals()['numero_tempo_especialidade_' + area + nome_adicional] = dict(Counter(globals()['tempo_especialidade' + '_' + area + nome_adicional]))
    if area != nomes_areas[-1]: globals()['numero_tempo_especialidade_' + area + nome_adicional].pop('Not_In_Area')
    for especiali in especialidades:
        globals()['porcent_' + especiali + '_tempo_especialidade' + '_' + area + nome_adicional] = round(Counter(globals()['tempo_especialidade' + '_' + area + nome_adicional])[especiali] / sum(list(globals()['numero_tempo_especialidade_' + area + nome_adicional].values())),2)

    # Calcula o tempo de uso total por regiã
    globals()['tempo_total_regiao_' + area + nome_adicional] = {}
    for lugar in set(globals()['regiao' + '_' + area + nome_adicional]):
        for num in range(len(globals()['regiao' + '_' + area + nome_adicional])):
            if globals()['regiao' + '_' + area + nome_adicional][num] == lugar:
                if lugar in globals()['tempo_total_regiao_' + area + nome_adicional]: globals()['tempo_total_regiao_' + area + nome_adicional][lugar] += globals()[dados_originais_areas[0] + '_' + area + nome_adicional][num]
                else: globals()['tempo_total_regiao_' + area + nome_adicional][lugar] = globals()[dados_originais_areas[0] + '_' + area + nome_adicional][num]

    # Calcula o número de usuário por região
    globals()['numero_usuarios_regiao_' + area + nome_adicional] = dict(Counter(globals()['regiao' + '_' + area + nome_adicional]))

    # Remove os usuários que não estão nas respectivas áreas, seja reumateraquiz, radioteraquiz, etc
    if area != nomes_areas[-1]:
        globals()['numero_usuarios_regiao_' + area + nome_adicional].pop('Not_In_Area')
        globals()['tempo_total_regiao_' + area + nome_adicional].pop('Not_In_Area')

print('Os dados foram analisados com sucesso...')

##################################################################################### GRÁFICOS

# Formatação padrão dos gráficos
plt.rcParams['font.family'] = 'Verdana'
plt.rcParams['text.color'] = cinza

tempo_acima_de = 1
titulo_area = -1 # Variável para fazer um loop pelos títulos das áreas

# Função para produzir gráficos
def produzir_graficos(tipo_grafico, # Tipo de gráfico a ser criado
                      titulo_grafico, # Título que aparecerá em cima do gráfico
                      titulo_arquivo, # Título do arquivo.png
                      largura_figura, altura_figura, # Dimensões da figura do gráfico
                      dados_graficos, rotulos_dados, # Dados e seus rótulos
                      espessura_barra = 0.4, x_rotulos = [], y_rotulos = [], porcent = False, unidade_tempo = False, rotulos_inteiros = False):

    # Se aparece a unidade de tempo no título ou não
    if unidade_tempo != False: tempo_unidade = unidade_tempo + '\n'
    else: tempo_unidade = ''

    # Cria a figura e define as suas dimensões
    plt.figure(figsize=(largura_figura, altura_figura))

    if tipo_grafico == 'bar': # Gráfico de barras verticais

        plt.bar(rotulos_dados, dados_graficos, color = areas[area], width = espessura_barra)
    
        for i in range(len(rotulos_dados)):
            if porcent == True: rotulos = str(round(dados_graficos[i]*100,3)) + '%' # Verifica se é para aparecer o símbolo de % nos rótulos
            else: rotulos = str(dados_graficos[i])
            plt.text(rotulos_dados[i], dados_graficos[i] + max(dados_graficos)*0.01, rotulos, ha = 'center')

    if tipo_grafico == 'table': # Tabela

        table = plt.table(cellText =  dados_graficos, loc='center', cellLoc='center', colLabels=None)
        for cell in table.get_celld().values():
            if cell in list(table.get_celld().values())[0:5]: cell.set_text_props(weight = 'bold')
            if area == nomes_areas[-1]: cell.set_text_props(color = branco)
            cell.set_facecolor(areas[area])

    if tipo_grafico == 'h_bar': # Gráfico de barras horizontais

        h_bars = plt.barh(rotulos_dados, dados_graficos, color = areas[area], height = espessura_barra)
        if rotulos_inteiros == True:
            for bar in h_bars: plt.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height() / 2, f'{int(bar.get_width())}', ha='left', va='center') # Rótulos inteiros
        else:
            for bar in h_bars: plt.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height() / 2, f'{bar.get_width():.1f}', ha='left', va='center') # Rótulos com decimais

    if tipo_grafico == 'pie': # Gráfico de pizza

        plt.pie(dados_graficos, labels = rotulos_dados, autopct='%1.1f%%', startangle=90, colors=[areas[area], cores_mais_claras[area]])
        plt.axis('equal')

    # Configurações comuns a todos os gráficos
    plt.xticks(x_rotulos, x_rotulos, color = cinza)
    plt.yticks(y_rotulos)
    plt.title(titulo_grafico + '\n' + tempo_unidade + titulos_areas[titulo_area])
    plt.tick_params(axis='both', which='both', length=0, labelcolor = cinza)
    plt.grid(False)
    for position in ['right', 'top', 'bottom', 'left']: plt.gca().spines[position].set_visible(False)
    plt.savefig(path_figuras + r'/graficos/' + area + r'/' + titulo_arquivo + '_' + titulos_areas[titulo_area] + '.png', dpi=400, transparent=True) # Salva a figura com fundo transparente
    plt.close()

# Loop para produzir os gráficos de cada área
for area in nomes_areas:

    titulo_area += 1
    nome_adicional = nomes_areas[-1]
    if area == nomes_areas[-1]: nome_adicional = ''

    if not os.path.exists(path_figuras + r'/graficos/' + area): os.makedirs(path_figuras + r'/graficos/' + area) # Cria a pasta gráficos se ela não existir

    # Cada um desses blocos produz um gráfico diferente

    # Tempos Totais de Uso em Cada Área
    dados_plot = []
    for dado in dados_originais_areas[1:5]: dados_plot.append(globals()['soma_' + dado + '_' + area + nome_adicional])
    rotulos_plot = ['Tempo Home','Tempo Quiz','Tempo Video','Tempo Comunidade']
    produzir_graficos('bar','Tempos de Uso em Cada Área', 'tempos_por_area', 8, 6, dados_plot, rotulos_plot,\
                      x_rotulos = rotulos_plot, espessura_barra = 0.7, unidade_tempo = '(em horas)')


    # Porcentagem de Usuários Ativos e Inativos
    if area != nomes_areas[-1]: # Produz esse gráfico se não é da área geral, ou seja, não apenas teraquiz
        rotulos_plot = ['Ativos na Área','Outros']
        dados_plot = [globals()['porcent_ativos_' + dados_originais_areas[-1] + '_' + area + nome_adicional],
                    globals()['porcent_inativos_' + dados_originais_areas[-1] + '_' + area + nome_adicional]]
        produzir_graficos('bar','Porcentagem de Usuários Ativos na Área', 'porcent_numero_de_usuarios_ativos', 6, 4.5, dados_plot, rotulos_plot,
                          x_rotulos = rotulos_plot, porcent = True)
    else: # Produz esse gráfico se é da área geral, ou seja, teraquiz
        rotulos_plot = ['Ativos','inativos']
        dados_plot = [globals()['porcent_ativos_' + dados_originais_areas[-1] + '_' + area + nome_adicional],
                    globals()['porcent_inativos_' + dados_originais_areas[-1] + '_' + area + nome_adicional]]
        produzir_graficos('bar','Porcentagem de Usuários Ativos na Área', 'porcent_numero_de_usuarios_ativos', 6, 4.5, dados_plot, rotulos_plot,
                          x_rotulos = rotulos_plot, porcent = True)


    # Porcentagem de Residentes, Especialistas e N.I.
    rotulos_plot = especialidades
    dados_plot = []
    for especiali in especialidades: dados_plot.append(globals()['porcent_' + especiali + '_tempo_especialidade' + '_' + area + nome_adicional])
    produzir_graficos('bar','Porcentagem Residentes, Especialistas e Não Informados', 'porcent_residentes_especialistas_ni', 6, 4.5, dados_plot, rotulos_plot,
                      x_rotulos = rotulos_plot, porcent = True)


    # Porcentagem de Residentes, Especialistas e N.I.
    rotulos_plot = especialidades[:-1]
    dados_plot = []
    for especiali in especialidades[:-1]: dados_plot.append(globals()['porcent_' + especiali + '_tempo_especialidade' + '_' + area + nome_adicional])
    produzir_graficos('bar','Porcentagem Residentes e Especialistas', 'porcent_residentes_especialistas', 6, 4.5, dados_plot, rotulos_plot,
                      x_rotulos = rotulos_plot, porcent = True)


    # Número de Usuários Ativos com Média de Tempo Mensal Acima de 5 minutos
    minutos = 5
    filtro_ativos = [valor for valor, ativo in zip(globals()['media_tempo_mes_' + area + nome_adicional],
                                                   globals()[dados_originais_areas[-1] + '_' + area + nome_adicional])
                     if ativo == True]
    filtro_ni = list(filter(lambda x: x != nao_informado, filtro_ativos))
    ordenar = sorted(filtro_ni)
    em_minutos = np.round(np.array(ordenar) / 60,1)
    auxiliar = em_minutos.tolist()
    acima = list(filter(lambda x: x >= minutos, auxiliar))
    abaixo = list(filter(lambda x: x < minutos, auxiliar))
    dados_plot = [len(acima),len(abaixo)]
    rotulos_plot  = ['Acima de ' + str(minutos) + ' minutos', 'Abaixo de ' + str(minutos) + ' minutos']
    produzir_graficos('bar','Número de Usuários com Média Mensal \nAcima de ' + str(minutos) + ' minutos', 'numero_usuarios_media_mensal_5', 6, 6, dados_plot, rotulos_plot,
                      x_rotulos = rotulos_plot)

    # Número de Usuários Ativos com Média de Tempo Mensal Acima de 1 minuto1
    minutos = 1
    filtro_ativos = [valor for valor, ativo in zip(globals()['media_tempo_mes_' + area + nome_adicional],
                                                   globals()[dados_originais_areas[-1] + '_' + area + nome_adicional])
                     if ativo == True]
    filtro_ni = list(filter(lambda x: x != nao_informado, filtro_ativos))
    ordenar = sorted(filtro_ni)
    em_minutos = np.round(np.array(ordenar) / 60,1)
    auxiliar = em_minutos.tolist()
    acima = list(filter(lambda x: x >= minutos, auxiliar))
    abaixo = list(filter(lambda x: x < minutos, auxiliar))
    dados_plot = [len(acima),len(abaixo)]
    rotulos_plot  = ['Acima de ' + str(minutos) + ' minutos', 'Abaixo de ' + str(minutos) + ' minutos']
    produzir_graficos('bar','Número de Usuários com Média Mensal \nAcima de ' + str(minutos) + ' minutos', 'numero_usuarios_media_mensal_1', 6, 6, dados_plot, rotulos_plot,
                      x_rotulos = rotulos_plot)


    # Número de Usuários Ativos com Média de Tempo Semanal Acima de 1 Minuto
    minutos = 1
    filtro_ativos = [valor for valor, ativo in zip(globals()['media_tempo_semana' + '_' + area + nome_adicional],
                                                   globals()[dados_originais_areas[-1] + '_' + area + nome_adicional])
                     if ativo == True]
    filtro_ni = list(filter(lambda x: x != nao_informado, filtro_ativos))
    ordenar = sorted(filtro_ni)
    em_minutos = np.round(np.array(ordenar) / 60,1)
    auxiliar = em_minutos.tolist()
    acima = list(filter(lambda x: x >= minutos, auxiliar))
    abaixo = list(filter(lambda x: x < minutos, auxiliar))
    dados_plot = [len(acima),len(abaixo)]
    rotulos_plot  = ['Acima de ' + str(minutos) + ' minutos', 'Abaixo de ' + str(minutos) + ' minutos']
    produzir_graficos('bar','Número de Usuários com Média Semanal \nAcima de ' + str(minutos) + ' Minuto', 'numero_usuarios_media_semanal', 6, 6, dados_plot, rotulos_plot,
                      x_rotulos = rotulos_plot)


    # Estatísticas Tempo de Uso Total
    medidas = ['Usuários','Total','Média','Mediana','Máximo']
    unidades = ['(Total Ativos)','(Horas)','(minutos)','(minutos)','(Horas)']
    estatisticas = [list(globals()['numero_' + dados_originais_areas[-1] + '_' + area + nome_adicional].values())[0],
                    globals()['soma_' +  dados_originais_areas[0] + '_' + area + nome_adicional],
                    globals()['media_' +  dados_originais_areas[0] + '_' + area + nome_adicional],
                    globals()['mediana_' +  dados_originais_areas[0] + '_' + area + nome_adicional],
                    globals()['maximo_' +  dados_originais_areas[0] + '_' + area + nome_adicional]]
    dados_plot = [medidas, unidades, estatisticas]
    produzir_graficos('table','Estatísticas do Tempo de Uso Total\nUsuários Ativos', 'estatisticas_tempo_de_uso_total', 6, 6, dados_plot, rotulos_plot)


    # Número de Usuários por Estado
    ordenar_dict = dict(sorted(globals()['numero_usuarios_regiao_' + area + nome_adicional].items(), key=lambda item: item[1]))
    rotulos_plot = list(ordenar_dict.keys())
    dados_plot = list(ordenar_dict.values())
    produzir_graficos('h_bar','Número de Usuários por Estado e Não Informado', 'numero_usuarios_estado_ni', 16, len(rotulos_plot)*0.40, dados_plot, rotulos_plot,
                      espessura_barra = 0.8, y_rotulos=rotulos_plot, rotulos_inteiros = True)


    # Número de Usuários por Estado Sem N.I.
    ordenar_dict = dict(sorted(globals()['numero_usuarios_regiao_' + area + nome_adicional].items(), key=lambda item: item[1]))
    fora_ni = {key: value for key, value in ordenar_dict.items() if key != 'N.I.'}
    rotulos_plot = list(fora_ni.keys())
    dados_plot = list(fora_ni.values())
    produzir_graficos('h_bar','Número de Usuários por Estado', 'numero_usuarios_estado', 16, len(rotulos_plot)*0.40, dados_plot, rotulos_plot,
                      espessura_barra = 0.8, y_rotulos=rotulos_plot, rotulos_inteiros = True)


    # Número de Usuários fora Brasil
    if area == nomes_areas[-1]:
        ordenar_dict = dict(sorted(globals()['numero_usuarios_regiao_' + area + nome_adicional].items(), key=lambda item: item[1]))
        fora_ni = {key: value for key, value in ordenar_dict.items() if key != 'N.I.'}
        fora_brasil = {key: value for key, value in fora_ni.items() if len(key) > 2}
        rotulos_plot = list(fora_brasil.keys())
        dados_plot = list(fora_brasil.values())
        produzir_graficos('h_bar','Número de Usuários Fora do Brasil', 'numero_usuarios_estrangeiros', 16, len(rotulos_plot)*0.40, dados_plot, rotulos_plot,
                          espessura_barra = 0.8, y_rotulos=rotulos_plot, rotulos_inteiros = True)


    # Tempo Total por Estado (mais de 6 minutos == 0.1 horas)
    ordenar_dict = dict(sorted(globals()['tempo_total_regiao_' + area + nome_adicional].items(), key=lambda item: item[1]))
    ordenar_dict = {key: value for key, value in ordenar_dict.items() if value >= 360}
    rotulos_plot = list(ordenar_dict.keys())
    auxiliar = list(ordenar_dict.values())
    auxiliar = np.round(np.array(auxiliar) / divisao_horas,1)
    dados_plot = auxiliar.tolist()
    produzir_graficos('h_bar','Tempo Total por Estado', 'tempo_total_estado_ni', 6, len(rotulos_plot)*0.4, dados_plot, rotulos_plot,
                      espessura_barra = 0.8, y_rotulos=rotulos_plot, unidade_tempo = '(em horas)')


    # Tempo Total por Estado sem NI (mais de 6 minutos == 0.1 horas)
    ordenar_dict = dict(sorted(globals()['tempo_total_regiao_' + area + nome_adicional].items(), key=lambda item: item[1]))
    ordenar_dict = {key: value for key, value in ordenar_dict.items() if value >= 360}
    fora_ni = {key: value for key, value in ordenar_dict.items() if key != 'N.I.'}
    rotulos_plot = list(fora_ni.keys())
    auxiliar = list(fora_ni.values())
    auxiliar = np.round(np.array(auxiliar) / divisao_horas,1)
    dados_plot = auxiliar.tolist()
    produzir_graficos('h_bar','Tempo Total por Estado', 'tempo_total_estado', 6, len(rotulos_plot)*0.4, dados_plot, rotulos_plot,
                      espessura_barra = 0.8, y_rotulos=rotulos_plot, unidade_tempo = '(em horas)')


    # Média de Tempo Mensal dos Usuários Mais Ativos
    tempo_horas = 1
    filtro_ativos = [valor for valor, ativo in zip(globals()['media_tempo_mes_' + area + nome_adicional], globals()[dados_originais_areas[-1] + '_' + area + nome_adicional]) if ativo == True]
    filtro_ni = list(filter(lambda x: x != nao_informado, filtro_ativos))
    ordenar = sorted(filtro_ni)
    em_horas = np.round(np.array(ordenar) / divisao_horas,1)
    auxiliar = em_horas.tolist()
    dados_plot = list(filter(lambda x: x > tempo_horas, auxiliar))
    rotulos_plot  = [str(i) for i in range(1, len(dados_plot) + 1)]
    produzir_graficos('h_bar','Média de Tempo Mensal dos Usuários Mais Ativos', 'media_tempo_total_mes', 8, len(rotulos_plot)*0.35 + 3, dados_plot, rotulos_plot,
                      espessura_barra = 0.6, unidade_tempo = '(em horas)')


    # Categorias Média de Tempo Mensal dos Usuários Mais Ativos
    tempo_horas_mais_ativo = 1
    mais = 4 ; entre = 2
    filtro_ativos = [valor for valor, ativo in zip(globals()['media_tempo_mes_' + area + nome_adicional], globals()[dados_originais_areas[-1] + '_' + area + nome_adicional]) if ativo == True]
    filtro_ni = list(filter(lambda x: x != nao_informado, filtro_ativos))
    ordenar = sorted(filtro_ni)
    em_horas = np.round(np.array(ordenar) / divisao_horas,1)
    auxiliar = em_horas.tolist()
    dados_plot = [len(list(filter(lambda x: x < tempo_horas_mais_ativo, auxiliar))),
                  len(list(filter(lambda x: x >= tempo_horas_mais_ativo and x < entre, auxiliar))),
                  len(list(filter(lambda x: x >= entre and x < mais, auxiliar))),
                  len(list(filter(lambda x: x >= mais, auxiliar)))]
    rotulos_plot  = ['Entre 1 minuto e ' + str(tempo_horas_mais_ativo) + ' hora','Mais de ' + str(tempo_horas_mais_ativo) + ' hora','Entre ' + str(entre) + ' e ' + str(mais) + ' horas','Mais de ' + str(mais) + ' horas']
    produzir_graficos('h_bar','Número de Usuários por Tempo de Uso Médio Mensal\nUsuários Mais Ativos', 'categorias_tempo_medio_mensal_acima_1', 15, 6, dados_plot, rotulos_plot,
                      y_rotulos=rotulos_plot, espessura_barra = 0.5, rotulos_inteiros = True)

    # Categorias Média de Tempo Mensal dos Usuários Mais Ativos
    tempo_horas_mais_ativo = 1
    mais = 4 ; entre = 2
    filtro_ativos = [valor for valor, ativo in zip(globals()['media_tempo_mes_' + area + nome_adicional], globals()[dados_originais_areas[-1] + '_' + area + nome_adicional]) if ativo == True]
    filtro_ni = list(filter(lambda x: x != nao_informado, filtro_ativos))
    ordenar = sorted(filtro_ni)
    em_horas = np.round(np.array(ordenar) / divisao_horas,1)
    auxiliar = em_horas.tolist()
    dados_plot = [len(list(filter(lambda x: x >= tempo_horas_mais_ativo and x < entre, auxiliar))),
                  len(list(filter(lambda x: x >= entre and x < mais, auxiliar))),
                  len(list(filter(lambda x: x >= mais, auxiliar)))]
    rotulos_plot  = ['Mais de ' + str(tempo_horas_mais_ativo) + ' hora','Entre ' + str(entre) + ' e ' + str(mais) + ' horas','Mais de ' + str(mais) + ' horas']
    produzir_graficos('h_bar','Número de Usuários por Tempo de Uso Médio Mensal\nUsuários Mais Ativos', 'categorias_tempo_medio_mensal_abaixo_1', 11, 5, dados_plot, rotulos_plot,
                      y_rotulos=rotulos_plot, espessura_barra = 0.5, rotulos_inteiros = True)


    # Mais e menos 5 minutos Usuários Ativos
    rotulos_plot = ['Mais de ' + str(5) + ' minutos','Menos de ' + str(5) + ' minutos']
    filtro_ativos = [valor for valor, ativo in zip(globals()[dados_originais_areas[0] + '_' + area + nome_adicional], globals()[dados_originais_areas[-1] + '_' + area + nome_adicional]) if ativo == True]
    tamanho_lista = len(filtro_ativos)
    acima_ = list(filter(lambda x: x >= 300, filtro_ativos))
    abaixo_ = list(filter(lambda x: x < 300, filtro_ativos))
    dados_plot = [round(len(acima_)/tamanho_lista,3), round(len(abaixo_)/tamanho_lista,3)]
    produzir_graficos('pie','Porcentagens Usuários Ativos', 'porcent_usuarios_ativos', 9, 12, dados_plot, rotulos_plot)


print('Os gráficos foram criados com sucesso.')