from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
import numpy as np
from datetime import datetime

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
# lista_menu, test_list = listar(nav, 'menuButtons')

# click_producao = test_list.loc[test_list[0] == 'Produção'].reset_index(drop=True)['index'][0]

# lista_menu[click_producao].click()



# options = webdriver.ChromeOptions()
# options.add_argument("--headless")
# -----------------------------------------------------------------

# Plano Mestre e Simulação
nav = webdriver.Chrome()#chrome_options=options)
time.sleep(1)
nav.maximize_window()
time.sleep(1)
# nav.get('https://cemag.innovaro.com.br/sistema') # Sistema de produção
nav.get('http://devcemag.innovaro.com.br:81/sistema') # Base H - Testes

# Usuário e Senha Cemag
nav.find_element(By.ID, 'username').send_keys('ti.cemag') 
time.sleep(2)
nav.find_element(By.ID, 'password').send_keys('cem@#161010')
time.sleep(1)
nav.find_element(By.ID, 'submit-login').click()

# Innovaro -> Produção -> Plano Mestre -> Plano Mestre
time.sleep(3)
nav.find_element(By.ID, 'bt_1892603865').click()
time.sleep(1)
nav.find_element(By.XPATH, '//*[@id="divTreeNavegation"]/div[10]/span[2]').click()
time.sleep(1)
nav.find_element(By.XPATH, '//*[@id="divTreeNavegation"]/div[13]/span[2]').click()
time.sleep(1)
nav.find_element(By.XPATH, '//*[@id="divTreeNavegation"]/div[16]/span[2]').click()
time.sleep(3)

# Innovaro -> Estoque -> Consultas -> Saldo de Recursos
# nav.find_element(By.XPATH, '//*[@id="divTreeNavegation"]/div[4]/span[1]').click()
# Consultas
# nav.find_element(By.XPATH, '//*[@id="divTreeNavegation"]/div[7]/span[1]').click()
# Saldo de recursos
# nav.find_element(By.XPATH,'//*[@id="divTreeNavegation"]/div[18]/span[2]').click()

# Entrando no Iframe e acessando os icones e o input de localizar
iframes(nav)
nav.find_element(By.XPATH, '/html/body/table/tbody/tr[1]/td/div/form/table/thead/tr[2]/td[1]/table/tbody/tr/td[1]/table/tbody/tr/td[1]/div').click()
time.sleep(3)
input_localizar = nav.find_element(By.ID, 'grInputSearch_grSimulacoes')
input_localizar.send_keys('Pendencia Diaria Carretas 29-03-2023 Compras')
input_localizar.send_keys(Keys.ENTER)
time.sleep(2)

# Excluindo os itens
nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[2]/td[1]/div').click()
time.sleep(1)
nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/thead/tr[1]/td[1]/table/tbody/tr/td[2]/table/tbody/tr/td[3]/div').click()
time.sleep(1)
saida_iframe(nav)
time.sleep(1)
nav.find_element(By.XPATH, '//*[@id="answers_0"]').click()
time.sleep(1)

# Pendências de pedidos
nav.find_element(By.XPATH, '//*[@id="buttonsContainer_1"]/td[1]/span[2]').click()
time.sleep(1)
iframes(nav)
time.sleep(1)
emissao_final = nav.find_element(By.XPATH, '//*[@id="grFiltroDePedidos"]/tbody/tr[1]/td[1]/table/tbody/tr/td/table/tbody/tr[17]/td[4]/table/tbody/tr/td[1]/input')
time.sleep(1)
emissao_final.send_keys(Keys.CONTROL + 'a')
time.sleep(1)
emissao_final.send_keys('h')
time.sleep(1)
emissao_final.send_keys(Keys.ENTER)
time.sleep(1)
saida_iframe(nav)
time.sleep(1)

# Executar busca de pedidos
nav.find_element(By.XPATH, '//*[@id="buttonsContainer_1"]/td/span[2]').click()

time.sleep(3)

iframes(nav)

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

# Percorrer as linhas do df1, ou seja, as que apresentam Prev. Emissão Doc. e Hora em Branco
for i in range(0,len(df1)):

    # if df1['Prev. Emissão Doc.'][i] == '':

    # Input campo de Data
    time.sleep(1)
    nav.find_element(By.XPATH, '/html/body/table/tbody/tr[2]/td/div/form/table/tbody/tr[1]/td[1]/table/tbody/tr[4]/td[5]/div/div').click()  

    time.sleep(1)
    nav.find_element(By.XPATH, "//*[@id='0']/td[5]/div/input").send_keys('h')
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
# time.sleep(2)
# saida_iframe(nav)
# nav.find_element(By.XPATH,'//*[@id="buttonsContainer_1"]/td[3]/span[2]').click()

# ------------- Estoque --------------
# time.sleep(1)

# input.send_keys(Keys.CONTROL + 'a')

# time.sleep(1)

# input.send_keys(Keys.DELETE)

# time.sleep(1)

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