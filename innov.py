from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd
import numpy as np
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import glob

def ultimo_arquivo():

    diretorio = r"C:\Users\TI\Downloads"
    arquivos_csv = [arquivo for arquivo in os.listdir(diretorio) if arquivo.endswith('.csv')]

    # ordenar os arquivos pela data de modificação, do mais recente para o mais antigo
    arquivos_csv = sorted(arquivos_csv, key=lambda arquivo: os.path.getmtime(os.path.join(diretorio, arquivo)), reverse=True)

    # pegar o arquivo mais recente
    ultimo_arquivo = arquivos_csv[0]

    return ultimo_arquivo, diretorio

scope = ['https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive"]
   
credentials = ServiceAccountCredentials.from_json_keyfile_name("service_account_cemag.json", scope)
client = gspread.authorize(credentials)
filename = 'service_account_cemag.json'
sa = gspread.service_account(filename)

# Conectando com google sheets e acessando Análise Previsão de Consumo (CMM / NTP ) DEE

sheet = 'Análise Previsão de Consumo (CMM / NTP ) DEE'
worksheet = 'Dados Simulação'
worksheet2 = 'Est. Produção'
worksheet3 = 'Dados Pedidos'

sh1 = sa.open(sheet)
wks1 = sh1.worksheet(worksheet)
wks2 = sh1.worksheet(worksheet2)
wks3 = sh1.worksheet(worksheet3)

df = wks1.get()
df2 = wks2.get()
df3 = wks3.get()

tabela = pd.DataFrame(df)
tabela2 = pd.DataFrame(df2)
tabela3 = pd.DataFrame(df3)

cabecalho = wks1.row_values(2)
cabecalho2 = wks2.row_values(2)
cabecalho3 = wks3.row_values(2)

#tratando planilha Análise Previsão de Consumo (CMM / NTP ) DEE
#df = df.set_axis(cabecalho, axis=1)

# Entrando no Iframe
def iframes(nav):

    iframe_list = nav.find_elements(By.CLASS_NAME,'tab-frame')

    for iframe in range(len(iframe_list)):
        time.sleep(1)
        try: 
            nav.switch_to.default_content()
            nav.switch_to.frame(iframe_list[iframe])
            print(iframe)
        except:
            pass
# Saindo Iframe
def saida_iframe(nav):
    nav.switch_to.default_content()
# Lista de classe para as abas do innovaro
def listar(nav, classe):
    
    lista_menu = nav.find_elements(By.CLASS_NAME, classe)
    
    elementos_menu = []

    for x in range (len(lista_menu)):
        a = lista_menu[x].text
        elementos_menu.append(a)

    test_lista = pd.DataFrame(elementos_menu)
    test_lista = test_lista.loc[test_lista[0] != ""].reset_index()

    return(lista_menu, test_lista)

# -------------------------------------------------------------------

# -----------------------------------------------------------------

# Plano Mestre e Simulação
nav = webdriver.Chrome()#chrome_options=options)
time.sleep(1)
nav.maximize_window()
time.sleep(1)
nav.get('https://cemag.innovaro.com.br/sistema') # Sistema de produção
# nav.get('http://devcemag.innovaro.com.br:81/sistema') # Base H - Testes


# Usuário e Senha Cemag
nav.find_element(By.ID, 'username').send_keys('ti.dev') #ti.dev ti.cemag
time.sleep(2)
nav.find_element(By.ID, 'password').send_keys('cem@#161010')
time.sleep(1)
nav.find_element(By.ID, 'submit-login').click()

# Innovaro -> Produção -> Plano Mestre -> Plano Mestre
WebDriverWait(nav,20).until(EC.presence_of_element_located((By.ID, 'bt_1892603865')))
time.sleep(1)
nav.find_element(By.ID, 'bt_1892603865').click()
time.sleep(3)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Produção'].reset_index(drop=True)['index'][0]
lista_menu[click_producao].click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Plano mestre e simulação (MPS)'].reset_index(drop=True)['index'][0]

lista_menu[click_producao].click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Plano mestre e simulação'].reset_index(drop=True)['index'][0]

lista_menu[click_producao].click()
time.sleep(6)
# Innovaro -> Estoque -> Consultas -> Saldo de Recursos
# nav.find_element(By.XPATH, '//*[@id="divTreeNavegation"]/div[4]/span[1]').click()
# Consultas
# nav.find_element(By.XPATH, '//*[@id="divTreeNavegation"]/div[7]/span[1]').click()
# Saldo de recursos
# nav.find_element(By.XPATH,'//*[@id="divTreeNavegation"]/div[18]/span[2]').click()

# Entrando no Iframe e acessando os icones e o input de localizar
iframes(nav)
WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div')))
time.sleep(1)
nav.find_element(By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div').click()
time.sleep(3)
input_localizar = nav.find_element(By.ID, 'grInputSearch_grSimulacoes')
time.sleep(1)
input_localizar.send_keys('Pendencia Diaria Carretas Compras')
time.sleep(1)
input_localizar.send_keys(Keys.ENTER)
time.sleep(1)

try: 
  while  nav.find_element(By.ID, 'progressMessageBox'):
    print('Carregando 2 ...')
except:
    print('Carregou 2') 
    time.sleep(1.5)

time.sleep(1.5)

# Excluindo os itens
nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td[1]/div').click()
time.sleep(1.5)
nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[3]/div').click()
time.sleep(1.5)
saida_iframe(nav)
time.sleep(1)
nav.find_element(By.XPATH, '//*[@id="answers_0"]').click()
time.sleep(1)

# Pendências de pedidos
nav.find_element(By.XPATH, '/html/body/div[2]/table/tbody/tr/td[9]/div/input').send_keys(Keys.CONTROL,Keys.SHIFT + 'e')
time.sleep(2)
iframes(nav)
time.sleep(1)

# Emissão Inicial
# emissao_inicial = nav.find_element(By.XPATH, '//*[@id="grFiltroDePedidos"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[17]/td[2]/table/tbody/tr/td[1]/input')
# time.sleep(1)
# emissao_inicial.send_keys(Keys.CONTROL + 'a')
# time.sleep(1)
# emissao_inicial.send_keys('30112022')
# time.sleep(1)
# emissao_inicial.send_keys(Keys.ENTER)
# time.sleep(1)

# Classe do Recurso 
classe_recursos = nav.find_element(By.XPATH, '//*[@id="grFiltroDePedidos"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input')
time.sleep(1)
classe_recursos.send_keys(Keys.CONTROL + 'a')
time.sleep(1)
classe_recursos.send_keys(Keys.BACKSPACE)
time.sleep(1)
classe_recursos.send_keys('Carretas Agrícolas')
time.sleep(2)
classe_recursos.send_keys(Keys.ENTER)
time.sleep(2)
input_carretas = nav.find_element(By.XPATH, '/html/body/div/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[3]/td[1]/input')
time.sleep(1)
input_carretas.click()
time.sleep(2)
input_carretas.send_keys(Keys.CONTROL, Keys.SHIFT + 'o')
time.sleep(3)
# Emissão Inicial 
emissao_inicial = nav.find_element(By.XPATH, '//*[@id="grFiltroDePedidos"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[17]/td[2]/table/tbody/tr/td[1]/input')
time.sleep(1)
emissao_inicial.send_keys(Keys.CONTROL + 'a')
time.sleep(1)
emissao_inicial.send_keys(Keys.BACKSPACE)
time.sleep(1)
emissao_inicial.send_keys('01/01/2021')
time.sleep(1)
emissao_inicial.send_keys(Keys.ENTER)
time.sleep(3)
# Emissão Final
emissao_final = nav.find_element(By.XPATH, '//*[@id="grFiltroDePedidos"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[17]/td[4]/table/tbody/tr/td[1]/input')
time.sleep(1)
emissao_final.send_keys(Keys.CONTROL + 'a')
time.sleep(1)
emissao_final.send_keys('h')
time.sleep(1)
emissao_final.send_keys(Keys.ENTER)
time.sleep(3)

# Executar busca de pedidos

emissao_final.send_keys(Keys.CONTROL,Keys.SHIFT + 'e') # Possivel mudança
time.sleep(2)
saida_iframe(nav)
time.sleep(2)

nav.find_element(By.ID, 'bt_1892603865').click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Plano mestre e simulação'].reset_index(drop=True)['index'][0]

lista_menu[click_producao].click() #clicando em producao
# nav.find_element(By.XPATH, '//*[@id="divTreeNavegation"]/div[15]/span[2]').click()
time.sleep(2)
WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[4]/div/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/span[2]')))

iframes(nav)
WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div')))
time.sleep(1)
nav.find_element(By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div').click()
time.sleep(3)
input_localizar = nav.find_element(By.ID, 'grInputSearch_grSimulacoes')
time.sleep(1)
input_localizar.send_keys('Pendencia Diaria Carretas Compras')
time.sleep(1)
input_localizar.send_keys(Keys.ENTER)
time.sleep(2)
input_localizar.send_keys(Keys.CONTROL, Keys.SHIFT + 'e')
time.sleep(2)
classe_recursos = nav.find_element(By.XPATH, '//*[@id="grFiltroDePedidos"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input')
time.sleep(1)
classe_recursos.send_keys(Keys.CONTROL + 'a')
time.sleep(1)
classe_recursos.send_keys(Keys.BACKSPACE)
time.sleep(1)
classe_recursos.send_keys('Mat Peças e Componentes')
time.sleep(1)
classe_recursos.send_keys(Keys.ENTER)
time.sleep(3)
classe_recursos.send_keys(Keys.CONTROL, Keys.SHIFT + 'e')

WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td[11]/div/div')))

# while len(nav.find_elements(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[10]/td[11]/div/div')) < 1:
#     time.sleep(1)
time.sleep(1.5)

nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[14]/td[11]/div/div').click()

time.sleep(2)

table = nav.find_element(By.XPATH,'//*[@id="grEspecificacao"]/tbody/tr[1]/td[1]/table')

table_html = table.get_attribute('outerHTML')

df = pd.read_html(str(table_html))

df1 = df.copy()

df1 = df1[0]

df1 = df1.rename(columns={0:'',1:'Recurso',3:'Quantidade',5:'Prev. Emissão Doc.',6:'Hora'})

df1.reset_index(drop=True)

df1 = df1[['Recurso','Quantidade','Prev. Emissão Doc.','Hora']]

df1 = df1.iloc[12:]

df1.replace(np.nan,'',inplace=True)

df1 = df1[df1['Recurso'] != ""]
df1 = df1[df1['Prev. Emissão Doc.'] == ""]

df1.reset_index(drop=True, inplace=True)
      
# hoje = datetime.now()

# data = hoje.strftime('%d/%m/%Y')

# hora = hoje.time()

# df1['Prev. Emissão Doc.'].replace('',data, inplace=True)

# hora_atual = hora.strftime('%H:%M')

# df1['Hora'].replace('', hora_atual ,inplace=True)
l = 4

# Percorrer as linhas do df1, ou seja, as que apresentam Prev. Emissão Doc. e Hora com input vazio 
for i in range(0,len(df1)):

    # if df1['Prev. Emissão Doc.'][i] == '':

    # Input campo de Data
    time.sleep(1)
    nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td[5]/div/div').click()  

    time.sleep(1)
    nav.find_element(By.XPATH, "//*[@id='0']/td[5]/div/input").send_keys('3112')
    time.sleep(1)
    nav.find_element(By.XPATH, "//*[@id='0']/td[5]/div/input").send_keys(Keys.ENTER)
    l = l+2
        # if df1['Hora'][i] == '': 

    time.sleep(1)
    # Input campo de Hora
    nav.find_element(By.XPATH, '//*[@id="0"]/td[6]/div/input').send_keys('h')
    time.sleep(1)
    nav.find_element(By.XPATH, '//*[@id="0"]/td[6]/div/input').send_keys(Keys.ENTER)
    time.sleep(1)
    nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[6]/td[5]/div/div').click()

# Explodir
time.sleep(2)
nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[6]/td[5]/div/input').send_keys(Keys.CONTROL,Keys.SHIFT + 'x')
time.sleep(2)
saida_iframe(nav)
time.sleep(1)
# Clicar em OK apos a tela de carregamento
while len(nav.find_elements(By.XPATH, '/html/body/div[10]/div[2]/table/tbody/tr[2]/td/div/button')) < 1:
    time.sleep(1)
time.sleep(1.5)

nav.find_element(By.XPATH,'//*[@id="confirm"]').click()
time.sleep(1.5)

saida_iframe(nav)
time.sleep(3)
nav.find_element(By.ID, 'bt_1892603865').click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Relatório de Logística de Compras da Simulação'].reset_index(drop=True)['index'][0]

lista_menu[click_producao].click()
time.sleep(6)
iframes(nav)
time.sleep(1.5)
simulacao = nav.find_element(By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input')
time.sleep(1.5)
simulacao.send_keys(Keys.BACKSPACE)
time.sleep(1.5)
simulacao.send_keys('Pendencia Diaria Carretas Compras') # Último Nível, Almox de Compras, Materiais e Produtos
time.sleep(2)
ult_niv = nav.find_element(By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
time.sleep(2)
ult_niv.send_keys(Keys.CONTROL + 'a')
time.sleep(2)
ult_niv.send_keys(Keys.BACKSPACE)
time.sleep(2)
ult_niv.send_keys('Último Nível')
time.sleep(1)
almo_comp = nav.find_element(By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[5]/td[2]/table/tbody/tr/td[1]/input')
time.sleep(2)
almo_comp.send_keys(Keys.CONTROL + 'a')
time.sleep(2)
almo_comp.send_keys(Keys.BACKSPACE)
time.sleep(1)
almo_comp.send_keys('Almox de Compras')
time.sleep(2)
mat_prod = nav.find_element(By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input')
time.sleep(2)
mat_prod.send_keys(Keys.CONTROL + 'a')
time.sleep(2)
mat_prod.send_keys(Keys.BACKSPACE)
time.sleep(1)
mat_prod.send_keys('Materiais e Produtos')
time.sleep(1)
nav.find_element(By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]/input').send_keys(Keys.CONTROL, Keys.SHIFT + 'e')

time.sleep(1)
while len(nav.find_elements(By.XPATH, '//*[@id="lid-0"]')) < 1:
    time.sleep(1)

time.sleep(3)
saida_iframe(nav)
time.sleep(2)
nav.find_element(By.XPATH, '/html/body/div[2]/table/tbody/tr/td[9]/div/input').send_keys(Keys.CONTROL, Keys.SHIFT + 'x')
time.sleep(2)
nav.find_element(By.ID, 'answers_3').click()
time.sleep(4)
nav.find_element(By.ID,'answers_0').click()
time.sleep(4)
nav.find_element(By.XPATH, '/html/body/div[2]/table/tbody/tr/td[9]/div/input').send_keys(Keys.CONTROL, Keys.SHIFT + 'e')
time.sleep(5)
iframes(nav)
time.sleep(1)
nav.find_element(By.XPATH, '//*[@id="_download_elt"]').click()
time.sleep(1)

# Organizando a planilha Recursos utilizados

ultimoArquivo, caminho = ultimo_arquivo()

tabela_recursos = pd.read_csv(r'C:/Users/TI/Downloads/' + ultimoArquivo, encoding='iso-8859-1', sep=';')

tabela_recursos = tabela_recursos.rename(columns={'="Recurso"':'Recurso','="Unid."':'Unid.','="Média"':'Média',
                                                  '="CMA"':'CMA', '="Simulado"':'Simulado','="Qtd.Est."':'Qtd.Est.','="Ped.Pend."':'Ped.Pend.',
                                                  '="Saldo"':'Saldo','="Cust.Unit."':'Cust.Unit.', '="TRP"':'TRP','="DEE"':'DEE'})

tabela_recursos_csv = tabela_recursos.copy()

tabela_recursos_csv = tabela_recursos_csv.replace('=','',regex=True)
tabela_recursos_csv.replace('"','',regex=True,inplace=True)

tabela_recursos_csv['Média'] = tabela_recursos_csv['Média'].replace(",",".",regex=True)
tabela_recursos_csv['Média'] = tabela_recursos_csv['Média'].replace("",0,regex=True)
tabela_recursos_csv['Média'] = tabela_recursos_csv['Média'].astype(float)

tabela_recursos_csv['CMA'] = tabela_recursos_csv['CMA'].replace(",",".",regex=True)
tabela_recursos_csv['CMA'] = tabela_recursos_csv['CMA'].replace("",0,regex=True)
tabela_recursos_csv['CMA'] = tabela_recursos_csv['CMA'].astype(float)

tabela_recursos_csv['Simulado'] = tabela_recursos_csv['Simulado'].replace(",",".",regex=True)
tabela_recursos_csv['Simulado'] = tabela_recursos_csv['Simulado'].replace("",0,regex=True)
tabela_recursos_csv['Simulado'] = tabela_recursos_csv['Simulado'].astype(float)

tabela_recursos_csv['Qtd.Est.'] = tabela_recursos_csv['Qtd.Est.'].replace(",",".",regex=True)
tabela_recursos_csv['Qtd.Est.'] = tabela_recursos_csv['Qtd.Est.'].replace("",0,regex=True)
tabela_recursos_csv['Qtd.Est.'] = tabela_recursos_csv['Qtd.Est.'].astype(float)

tabela_recursos_csv['Ped.Pend.'] = tabela_recursos_csv['Ped.Pend.'].replace(",",".",regex=True)
tabela_recursos_csv['Ped.Pend.'] = tabela_recursos_csv['Ped.Pend.'].replace("",0,regex=True)
tabela_recursos_csv['Ped.Pend.'] = tabela_recursos_csv['Ped.Pend.'].astype(float)

tabela_recursos_csv['Saldo'] = tabela_recursos_csv['Saldo'].replace(",",".",regex=True)
tabela_recursos_csv['Saldo'] = tabela_recursos_csv['Saldo'].replace("",0,regex=True)
tabela_recursos_csv['Saldo'] = tabela_recursos_csv['Saldo'].astype(float)

tabela_recursos_csv['Cust.Unit.'] = tabela_recursos_csv['Cust.Unit.'].replace(",",".",regex=True)
tabela_recursos_csv['Cust.Unit.'] = tabela_recursos_csv['Cust.Unit.'].replace("",0,regex=True)
tabela_recursos_csv['Cust.Unit.'] = tabela_recursos_csv['Cust.Unit.'].astype(float)

tabela_recursos_csv['DEE'] = tabela_recursos_csv['DEE'].replace(",",".",regex=True)
tabela_recursos_csv['DEE'] = tabela_recursos_csv['DEE'].replace("",0,regex=True)
tabela_recursos_csv['DEE'] = tabela_recursos_csv['DEE'].astype(float)

#Limpar valores nulos e transformar em lista
tabela_recursos_csv_lista = tabela_recursos_csv.fillna('').values.tolist()

#Apagar valores da planilha
sh1.values_clear("'Dados Simulação'!E2:O")

#Atualizar valores da planilha
wks1.update('E2:O', tabela_recursos_csv_lista)

# ------------------------------------
# Ajustando a planilha Análise Previsão de Consumo (CMM / NTP ) DEE

tabela2 = tabela.copy()
tabela2 = tabela2.iloc[:, 4:15]
tabela2 = tabela2.rename(columns={4:'Recurso',5:'Unid.',6:'Média',7:'CMA', 8:'Simulado',9:'Qtd.Est.',10:'Ped.Pend.', 11:'Saldo',12:'Cust.Unit.', 13:'TRP',14:'DEE'})
tabela2 = tabela2.reset_index()
tabela2 = tabela2.drop(0)
tabela2 = tabela2.drop('index',axis=1)

# ------------- Estoque --------------

saida_iframe(nav)
time.sleep(2)
nav.find_element(By.ID, 'bt_1892603865').click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Estoque'].reset_index(drop=True)['index'][0]

lista_menu[click_producao].click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Consultas'].reset_index(drop=True)['index'][0]

lista_menu[click_producao].click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Saldos de Recursos - CEMAG'].reset_index(drop=True)['index'][0]

lista_menu[click_producao].click()
time.sleep(2)

iframes(nav)
time.sleep(1.5)
data_base = nav.find_element(By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
time.sleep(1)

while len(data_base) < 1:
    time.sleep(1)
time.sleep(1.5)

data_base.send_keys('h')
time.sleep(1.5)
data_base.send_keys(Keys.ENTER)
time.sleep(2)

nav.find_element(By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[7]/td[2]/table/tbody/tr/td[1]/input').send_keys(Keys.CONTROL, Keys.SHIFT + 'x')
time.sleep(2)

saida_iframe(nav)
time.sleep(2)
try:    
    while nav.find_element(By.XPATH, '//*[@id="statusMessageBox"]'):
        print('Carregando...')
except:
    print('Carregou 1')

try: 
    while  nav.find_element(By.XPATH, '//*[@id="progressMessageBox"]'):
        print('Carregando 2 ...')
except:
    print('Carregou 2') 
    time.sleep(1)
time.sleep(1.5)

iframes(nav)
time.sleep(1.5)

while len(nav.find_element(By.ID, '_lbl_dadosGeradosComSucessoSelecionaAOpcaoExportarParaSelecionarOFormatoDeExportacao')):
    time.sleep(1)
time.sleep(1.5)

nav.find_element(By.XPATH, '/html/body').send_keys(Keys.CONTROL,Keys.SHIFT + 'x')

time.sleep(2)
saida_iframe(nav)
time.sleep(2)

nav.find_element(By.ID,'answers_0').click()
time.sleep(4)
nav.find_element(By.XPATH, '/html/body/div[2]/table/tbody/tr/td[9]/div/input').send_keys(Keys.CONTROL,Keys.SHIFT + 'e')
time.sleep(2)
try: 
    while  nav.find_element(By.XPATH, '//*[@id="progressMessageBox"]'):
        print('Carregando 2 ...')
except:
    print('Carregou 2') 
    time.sleep(1)
time.sleep(1.5)
iframes(nav)
time.sleep(1)
nav.find_element(By.XPATH, '//*[@id="_download_elt"]').click()
time.sleep(1)
# ------------------------------------------------------------------
# Organizando a planilha Saldo de Recursos - CEMAG

ultimoArquivo, caminho = ultimo_arquivo()

tabela_saldos = pd.read_csv(r'C:/Users/TI/Downloads/' + ultimoArquivo, encoding='iso-8859-1', sep=';')

tabela_saldos = tabela_saldos.rename(columns={'="1o. Agrupamento"':'1o. Agrupamento','=" "':'','="Depósito"':'Depósito',
                                                  '="Recurso#Classe"':'Recurso#Classe', '="Recurso#Unid. Medida"':'Recurso#Unid. Medida','="Saldo"':'Saldo','="Custo#Total"':'Custo#Total',
                                                  '="Custo#Médio"':'Custo#Médio'})

tabela_saldos_csv = tabela_saldos.copy()

tabela_saldos_csv = tabela_saldos_csv.replace('=','',regex=True)
tabela_saldos_csv.replace('"','',regex=True,inplace=True)

tabela_saldos_csv['Saldo'] = tabela_saldos_csv['Saldo'].replace(",",".",regex=True)
tabela_saldos_csv['Saldo'] = tabela_saldos_csv['Saldo'].replace("",0,regex=True)
tabela_saldos_csv['Saldo'] = tabela_saldos_csv['Saldo'].astype(float)

#Limpar valores nulos e transformar em lista
tabela_saldos_csv_lista = tabela_saldos_csv.fillna('').values.tolist()

sh1.values_clear("'Est. Produção'!N3:U")

#Atualizar valores da planilha
wks2.update('N3:U', tabela_saldos_csv_lista)

# ----------------Compra------------------
saida_iframe(nav)
time.sleep(2)
nav.find_element(By.ID,'bt_1892603865').click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Compra'].reset_index(drop=True)['index'][0]

lista_menu[click_producao].click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Consultas'].reset_index(drop=True)['index'][0]

lista_menu[click_producao].click()
time.sleep(2)

lista_menu, test_list = listar(nav, 'webguiTreeNodeLabel')
time.sleep(1)
click_producao = test_list.loc[test_list[0] == 'Análise de Pedidos Pendentes ou Baixados - CEMAG'].reset_index(drop=True)['index'][0]

lista_menu[click_producao].click()
time.sleep(15)
# mudar essa parte
iframes(nav)
time.sleep(1)

emissao = nav.find_element(By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[29]/td[4]/table/tbody/tr/td[1]/input')
time.sleep(1.5)
emissao.send_keys(Keys.CONTROL + 'a')
time.sleep(1.5)
emissao.send_keys('h')
time.sleep(1.5)
emissao.send_keys(Keys.ENTER)
time.sleep(1.5)
nav.find_element(By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input').send_keys(Keys.CONTROL, Keys.SHIFT + 'x')
WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="_lbl_dadosGeradosComSucessoSelecionaAOpcaoExportarParaSelecionarOFormatoDeExportacao"]')))

time.sleep(2)

nav.find_element(By.XPATH, '/html/body').send_keys(Keys.CONTROL,Keys.SHIFT + 'x')
time.sleep(1.5)
saida_iframe(nav)

time.sleep(2)
nav.find_element(By.ID,'answers_0').click()
time.sleep(2)
iframes(nav)
time.sleep(2)
nav.find_element(By.XPATH, '/html/body/div[2]/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input').send_keys(Keys.CONTROL,Keys.SHIFT + 'e')
time.sleep(1.5)

iframes(nav)

time.sleep(1.5)

WebDriverWait(nav,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="_download_elt"]')))

time.sleep(3)

nav.find_element(By.XPATH, '//*[@id="_download_elt"]').click()

time.sleep(1.5)

# ---------------------------------------------------------------------------
# Organizando a planilha de Análise de Pedidos Pendentes ou Baixados - CEMAG 

ultimoArquivo, caminho = ultimo_arquivo()

tabela_analise = pd.read_csv(r'C:/Users/TI/Downloads/' + ultimoArquivo, encoding='iso-8859-1', sep=';')

tabela_analise = tabela_analise.rename(columns={'="1o. Agrupamento"':'Região','="2o. Agrupamento"':'Estado','="3o. Agrupamento"':'Pessoa',
                                                  '="4o. Agrupamento"':'Classe Recurso', '="5o. Agrupamento"':'Data Entrega','="Chave ¹   Ch Criação ²"':'Chave ¹ Ch Criação ²','="Emissão"':'Emissão',
                                                  '="Dias Entrega"':'Dias Entrega', '="Classe"':'Classe','="Estabelecimento"':'Estabelecimento', '="Loc Escrit"':'Loc Escrit','="Recurso"':'Recurso',
                                                  '="Observação"':'Observação','="Núcleo"':'Núcleo','="Qde Ped"':'Qde Ped','="Unitário"':'Unitário','="Desc Venda"':'Desc Venda',
                                                  '="% Venda"':'% Venda','="Desc Item"':'Desc Item','="% Item"':'% Item','="Total"':'Total','="Qde Atend"':'Qde Atend',
                                                  '="Qde Canc"':'Qde Canc','="Qde Pend"':'Qde Pend'})

tabela_analise_csv = tabela_analise.copy()

coluna = tabela_analise.columns.tolist()
tabela_analise_csv = tabela_analise_csv.drop(['="Baixa"', '="Movimentação"', '="DP"', '="Tipo"', '="Número"', '="Qde Baixa"', '="Op. Vinculada"'],axis=1)
tabela_analise_csv = tabela_analise_csv.replace('=','',regex=True)
tabela_analise_csv.replace('"','',regex=True,inplace=True)
#tabela_analise_csv['Data Entrega'] = pd.to_datetime(tabela_analise_csv['Data Entrega'])

tabela_analise_csv['Qde Ped'] = tabela_analise_csv['Qde Ped'].replace(",",".",regex=True)
tabela_analise_csv['Qde Ped'] = tabela_analise_csv['Qde Ped'].replace("",0,regex=True)
tabela_analise_csv['Qde Ped'] = tabela_analise_csv['Qde Ped'].astype(float)

tabela_analise_csv['Qde Pend'] = tabela_analise_csv['Qde Pend'].replace(",",".",regex=True)
tabela_analise_csv['Qde Pend'] = tabela_analise_csv['Qde Pend'].replace("",0,regex=True)
tabela_analise_csv['Qde Pend'] = tabela_analise_csv['Qde Pend'].astype(float)

#Limpar valores nulos e transformar em lista
tabela_analise_csv_lista = tabela_analise_csv.fillna('').values.tolist()

sh1.values_clear("'Dados Pedidos'!B2:Y")

#Atualizar valores da planilha
wks3.update('B2:Y', tabela_analise_csv_lista)

time.sleep(1)

# data = nav.find_element(By.XPATH, '//*[@id="vars"]/tbody/tr[1]/td[1]/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/table/tbody/tr/td[1]/input')
# time.sleep(1)
# data.send_keys('h')
# time.sleep(1)
# data.send_keys(Keys.ENTER)
# time.sleep(1)
# saida_iframe(nav)
# nav.find_element(By.XPATH, '//*[@id="buttonsContainer_1"]/td[1]/span[2]').click()
# time.sleep(1.5)

# try:
#     while nav.find_element(By.ID, 'statusMessageBox'):
#         print('Carregando...')
# except:
#     print('Carregou 1')

# try: 
#     while  nav.find_element(By.ID, 'progressMessageBox'):
#         print('Carregando 2 ...')
# except:
#     print('Carregou 2') 

# time.sleep(1)
# nav.find_element(By.XPATH, '//*[@id="buttonsContainer_1"]/td[2]/span[2]').click()
# time.sleep(1)

# nav.find_element(By.ID, 'answers_0').click()
# time.sleep(1) 

# saida_iframe(nav)

# nav.find_element(By.XPATH, '/html/body/div[4]/div[2]/div[1]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/span[2]').click()
# time.sleep(1.5)

# saida_iframe(nav)

# try:
#     while  nav.find_element(By.ID, 'content_progressMessageBox'):
            
#             print('Carregando 3 ...')
# except:
#     print('Carregou 3') 

# iframes(nav)

# nav.find_element(By.XPATH, '//*[@id="_download_elt"]').click()