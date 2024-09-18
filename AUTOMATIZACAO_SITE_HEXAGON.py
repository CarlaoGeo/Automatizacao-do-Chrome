import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import datetime, timedelta
from selenium.common.exceptions import ElementClickInterceptedException, NoSuchElementException,TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import os
from dotenv import load_dotenv
#ESSE PROGRAMA AUTOMATIZA O CHROME PARA QUE AS INFORMAÇÕES DO SITE HEXAGON SEJAM COLHIDAS AUTOMATICAMENTE E SEJAM LEVADAS PARA UM EXCEL 
#QUE SERVIRÁ DE BANCO DE DADOS E POSTERIORMENTE COM ESSES DADOS POSSA CRIAR UM POWERBI POR EXEMPLO COM OS DADOS
#NO MOMENTO ESSE PROGRAMA SÓ COLHE INFORMAÇÕES DOS EQUIPAMENTOS
#CADASTRADOS NO MÓDULO DE COLHEITA DA EMPRESA


#CARREGA O ARQUIVO ENV
load_dotenv()


#FUNÇÃO PARA VER SE O ELEMENTO ESTA "CLICAVEL" E VISÍVEL
def ensure_element_clickable(xpath):
    element = WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", element)
    WebDriverWait(driver, 20).until(EC.visibility_of(element))
    return element


# FUNÇÃO DE CLICK EM JAVASCRIPT
def click_using_js(xpath):
    element = driver.find_element(By.XPATH, xpath)
    driver.execute_script("arguments[0].click();", element)
    
#INSTALA E RODA O CHROMEDRIVER SEM A NECESSIDADE DE TER BAIXADO NO COMPUTADOR
service = Service(ChromeDriverManager().install())

driver = webdriver.Chrome(service=service)

#DEFINE COMO A JANELA VAI SE COMPORTAR
driver.maximize_window()

#ACESSO AO SITE
driver.get(os.getenv("SITE_HEXAGON"))

time.sleep(5)
#FAZER O LOGIN
#PEGA O LOCAL DE COLOCAR O USUARIO
username = driver.find_element(By.XPATH, "/html/body/app-root/app-login/app-access-container/div/div[2]/div[3]/form/div[1]/input")

#PEGA O LOCAL DE COLCOAR A SENHA
senha = driver.find_element(By.XPATH, "/html/body/app-root/app-login/app-access-container/div/div[2]/div[3]/form/div[2]/input")

#PEGA O USUARIO E A SENHA DO ARQUIVO ENV E COLOCA NOS RESPECTIVOS LOCAIS JA PEGOS
username.send_keys(os.getenv("USUARIO"))
senha.send_keys(os.getenv("SENHA"))

#FAZ O CLICK USANDO O PROPRIO SELENIUM MESMO
login = driver.find_element(By.XPATH, "/html/body/app-root/app-login/app-access-container/div/div[2]/div[3]/form/div[3]/p-button/button/span")
login.click()

#LIMPAR FILTRO DA PAGINA
limparfiltro = "/html/body/app-root/div/app-monitoring-grid-page/app-content-page/div/div/div/app-multiple-filter/div/div[1]/div/div[2]/app-button/button/div/app-text/div"

#VERIFICA SE O ELEMENTO ESTA ATIVO
ensure_element_clickable(limparfiltro)

#CLICA NO BOTÃO DE LIMPAR O FILTRO USANDO A FUNÇÃO DE CLICK DO JAVA SCRIPT
click_using_js(limparfiltro)

#ESPERA 5 SEGUNDOS E ROLA A PÁGINA PARA O TOPO
time.sleep(5)
driver.execute_script("window.scrollTo(0,0)")



#ESPECIFICA QUE O EXCEL (QUE USAMOS COMO BANCO DE DADOS AQUI) SEJA CRIADO NO MESMO DIRETÓRIO DO PROGRAMA
caminho = 'dados_equipamentos.xlsx'

#SE JA EXISTE ELE IRA LER OS DADOS
if os.path.exists(caminho):
    df = pd.read_excel(caminho)
    
#SE NÃO CRIARÁ UM DO ZERO
else:
    df = pd.DataFrame(columns=['Frota', 'Atividade', 'Grupo de atividade', 'Função', 'Em Atividade', 'Frente', 'Status', 'HORA/DATA','Latitude', 'Longitude', 'Cracha','UltimaCom'])
    df.to_excel(caminho, index=False)
    print(f"Arquivo criado em: {caminho}")
    
# GARANTE QUE A COLUNA FROTA É NUMÉRICA
df['Frota'] = pd.to_numeric(df['Frota'], errors='coerce')

# LOOP PARA ACESSAR S DADOS DO SITE
while True:
    try:
        #PEGA QUANTOS QUADRADOS DE EQUIPAMENTOS TEM NA PÁGINA
        num_colunas = len(driver.find_elements(By.XPATH, '/html/body/app-root/div/app-monitoring-grid-page/app-content-page/div/div/div/app-monitoring-grid/app-column/app-row/div/app-column'))
                                                          

        #FAZ O LOOP PARA IR DE UM EM UM PARA PASSAR EM TODOS OS QUADRADOS DE EQUIPAMENTO
        for i in range(1, num_colunas + 1):
                        
            try:
                #VERIFICA SE O ELEMENTO ESTA ATIVO E ASSIM PEGA SUAS MEDIDAS
                quadro = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/app-monitoring-grid-page/app-content-page/div/div/div/app-monitoring-grid/app-column/app-row/div/app-column[' + str(i) + ']'))
                )
                pos = quadro.location
                tam = quadro.size
                #COM AS MEDIDAS A PÁGINA VAI ROLANDO ATÉ O QUADRADO FICAR VISÍVEL
                driver.execute_script("window.scrollTo(0, " + str(pos['y'] - tam['height'] / 2) + ")")

                

                # CAPTURA DOS TEXTOS EM CADA QUADRO
                try:
                    #frota1 = WebDriverWait(driver, 10).until(
                        #EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/app-monitoring-grid-page/app-content-page/div/div/div/app-monitoring-grid/app-column/app-row/div/app-column[' + str(i) + ']/app-monitoring-card/app-content-page/div/app-row[1]/div/app-column/app-row/div/app-column[1]/app-row/div/app-column/app-label[1]/label'))
                    #)
                    frota1 = driver.find_element(By.XPATH,"/html/body/app-root/div/app-monitoring-grid-page/app-content-page/div/div/div/app-monitoring-grid/app-column/app-row/div/app-column[" + str(i) + "]/app-monitoring-card/app-content-page/div/app-row[1]/div/app-column/app-row/div/app-column[1]/app-row/div/app-column/app-label[1]/label")
                    frota = frota1.get_attribute('innerText').strip()
                    try:
                        frota = int(frota)
                    except ValueError:
                        frota = None
                    print(frota)
                    
                    #PEGA O TEMPO EM ATIVIDADE DO EQUIPAMENTO
                    tempo1 = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/app-monitoring-grid-page/app-content-page/div/div/div/app-monitoring-grid/app-column/app-row/div/app-column[' + str(i) + ']/app-monitoring-card/app-content-page/div/app-column[1]/app-column[1]/app-column[3]/app-row/div/app-text/div'))
                    )
                    tempo = tempo1.text
                    print(tempo)
                    
                    #PEGA A FRENTE DO EQUIPAMENTO
                    frente1 = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/app-monitoring-grid-page/app-content-page/div/div/div/app-monitoring-grid/app-column/app-row/div/app-column[' + str(i) + ']/app-monitoring-card/app-content-page/div/app-column[2]/app-row/div/app-text/div'))
                    )
                    frente = frente1.text
                    print(frente)
                    
                    #PEGA A ATIVIDADE DO EQUIPAMENTO
                    atividade = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, f"/html/body/app-root/div/app-monitoring-grid-page/app-content-page/div/div/div/app-monitoring-grid/app-column/app-row/div/app-column[" + str(i) + "]/app-monitoring-card/app-content-page/div/app-column[1]/app-column[1]/app-column[1]/app-row/div/app-text/div"))
                    )
                    atividade_text1 = atividade.text
                    #AQUI TEM UM FILTRO CASO TENHA ALGUM CARACTERE QUE NÃO SEJA LETRA OU ESPAÇO
                    atividade_text = ''.join(char for char in atividade_text1 if char.isalpha() or char.isspace())
                    
                    atividade_text = atividade_text.strip()
                    print(atividade_text)

                    #PEGA O GRUPO DE ATIVIDADE DO EQUIPAMENTO
                    grupativ = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/app-monitoring-grid-page/app-content-page/div/div/div/app-monitoring-grid/app-column/app-row/div/app-column[' + str(i) + ']/app-monitoring-card/app-content-page/div/app-row[2]/div/app-column[1]/app-tag/div/app-text/div'))
                    )
                    
                    #PEQUENA LÓGICA PARA MUDAR O VALOR DA VARIAVEL "GRUPATIV" POIS ESTAVA DANDO UM ERRO DO JEITO QUE ESTAVA VINDO
                    grupativ = grupativ.text.strip()
                    if grupativ == "1 - Produtiva":
                        grupativ = "Produtiva"
                    elif grupativ == "2 - Auxiliar":
                        grupativ = "Auxiliar"
                    print(grupativ)
                    
                    
                    #PEGA A FUNÇÃO DO EQUIPMAMENTO
                    funcao = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, '/html/body/app-root/div/app-monitoring-grid-page/app-content-page/div/div/div/app-monitoring-grid/app-column/app-row/div/app-column[' + str(i) + ']/app-monitoring-card/app-content-page/div/app-column[1]/app-column[1]/app-column[2]/app-row/div/app-text/div'))
                    )
                    func = funcao.text
                    print(func)
                    
                    #PEGA A ULTIMA COMUNICAÇÃO DO EQUIPAMENTO
                    recente = driver.find_element(By.XPATH, '/html/body/app-root/div/app-monitoring-grid-page/app-content-page/div/div/div/app-monitoring-grid/app-column/app-row/div/app-column[' + str(i) + ']/app-monitoring-card/app-content-page/div/app-column[1]/app-column[1]/app-column[4]/app-row/div/app-text/div')
                    recente_text = recente.text.strip()
                    ultima = recente_text
                    print(ultima)
                    
                    #TRANSFORMA A VARIAVEL EM VALOR DE TEMPO
                    try:
                        ultima_com = datetime.strptime(recente_text, "%d/%m/%Y %H:%M:%S")
                        
                        
                    except ValueError:
                        print("Erro")
                    agora = datetime.now()
                    diferenca = agora - ultima_com
                    limite = timedelta(minutes=10)
                    #LÓGICA PARA DEFINIR SE O EQUIPAMENTO ESTA ONLINE OU OFFLINE, COMPARANDO O TEMPO DE AGORA COM A VARIAVEL "ULTIMA_COM"
                    if diferenca > limite:
                        online3 = "Offline"
                    else:
                        online3 = "Online"
                        
                    
                    
                    print(online3)


                    #VARIAVEL QUE ARMAZENA O LOCAL DE UM ELEMENTO DO SITE
                    izinho = '/html/body/app-root/div/app-monitoring-grid-page/app-content-page/div/div/div/app-monitoring-grid/app-column/app-row/div/app-column['+str(i)+']/app-monitoring-card/app-content-page/div/app-row[3]/div/app-column/app-icon[2]/i'
                    
                    try:
                        #VERIFICA SE ESSE ELEMENTO ESTA ATIVO E PODE SER CLICADO E USA UMA FUNÇÃO COM JAVASCRIPT PARA FAZER O CLICK
                        ensure_element_clickable(izinho)
                        click_using_js(izinho)
                        time.sleep(1)
                    
                        #PEGA A LATITUDE LOCALIZANDO UMA CÉLULA QUE CONTEM A PALAVRA LATITUDE E PEGA O SEU "IRMÃO" QUE SERIA A CÉLULA A SUA FRENTE
                        lat = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//tr[contains(@class, "ui-selectable-row")]//td[contains(., "Latitude")]/following-sibling::td[1]'))
                            )
                        lati = lat.get_attribute('innerText').strip().replace(',','.')
                        #SE A LATUTUDE FOI PEGA
                        if lati:
                            #TENTA TRANSFORMA ESSE VALOR EM FLOAT
                            try:
                                latitude = float(lati)
                                print(latitude)
                            except ValueError:
                                print(f"Não convertido '{lati}' para float")
                        else:
                            print("Vazio")
                    
                        #PEGA A LONGITUDE LOCALIZANDO UMA CÉLULA QUE CONTEM A PALAVRA LONGITUDE E PEGA O SEU "IRMÃO" QUE SERIA A CÉLULA A SUA FRENTE
                        long = WebDriverWait(driver,10).until(
                            EC.presence_of_element_located((By.XPATH, '//tr[contains(@class, "ui-selectable-row")]//td[contains(., "Longitude")]/following-sibling::td[1]'))
                            )
                        longi = long.get_attribute('innerText').strip().replace(',','.')
                        #SE A LONGITUDE FOI PEGA
                        if longi:
                            #TENTA TRANSFORMAR ESSE VALOR EM FLOAT
                            try:
                                longitude = float(longi)
                                print(longitude)
                            except ValueError:
                                print(f"Não convertido '{longi}' para float")
                        else:
                            print("Vazio")
                            
                        #PEGA O NOME DO OPERADOR DO EQUIPAMENTO
                        cracha1 = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, "//*[@id='text-operator-name-id']//div"))
                            )
                        cracha = cracha1.text.strip()
                                     
                        print(cracha)
                        

                    except ElementClickInterceptedException:
                        print("ERRO")
                    
                    
           
                    #CLICA EM UM LUGAR DA TELA PARA FECHAR A TABELA QUE SE ABRIU
                    driver.execute_script("document.elementFromPoint(30, 0).click();")
                    
                 
                    

                #CASO DE ERRO EM ALGUM QUADRO MOSTRA EM QUAL DEU ERRO
                except Exception as e:
                    print(f"Erro ao capturar os dados do equipamento {i}: {e}")
                    

                #SEMPRE SALVA OS DADOS TEMPORARIAMENTE A CADA QUADRO
                finally:
                    if frota is not None:
                        if pd.to_numeric(df['Frota'], errors='coerce').isin([frota]).any():
                            #SE O VALOR DE FROTA JA EXISTE, SUBSTITUI OS VALORES, ATUALIZANDO
                            df.loc[df['Frota'] == frota, ['Atividade', 'Grupo de atividade', 'Função', 'Em Atividade', 'Frente', 'Status','Latitude','Longitude','Cracha','UltimaCom']] = [atividade_text, grupativ, func, tempo, frente, online3,latitude,longitude,cracha,ultima]
                        else:
                            #CASO NÃO EXISTA O VALOR DE FROTA ADICIONA EM UMA NOVA LINHA
                            atualizacao = pd.DataFrame([{'Frota': frota, 'Atividade': atividade_text, 'Grupo de atividade': grupativ, 'Função': func, 'Em Atividade': tempo, 'Frente': frente, 'Status': online3, 'Latitude': latitude,'Longitude': longitude, 'Cracha': cracha,'UltimaCom': ultima}])
                            
                            df = pd.concat([df, atualizacao], ignore_index=True)
                        print("Dados do equipamento salvos.")

            except Exception as e:
                print(f"Erro ao processar a coluna {i}: {e}")
                
                continue

    except Exception as e:
        print(f"Erro ao iterar sobre as colunas: {e}")

    # MÉTODO PARA SALVAR O DATAFRAME AO FINAL DO LOOP E AI SIM CONSEGUIMOS VER OS DADOS NO EXCEL
    try:
        
        df.drop_duplicates(subset=['Frota'], keep='last', inplace=True)  # REMOVE AS DUPLICATAS COM BASE NO VALOR DE "FROTA"
        print("Duplicatas removidas")
        #PEGA A HORA ATUAL E A ADICIONA NA PRIMEIRA LINHA, SOMENTE PARA TER UM CONTROLE DE TEMPO DE QUANDO FOI ATUALIZADO AS INFORMAÇÕES
        hora_atual = datetime.now()
        hora_formatada = hora_atual.strftime("%d/%m/%Y %H:%M:%S")
        print(hora_formatada)
        df.loc[0, 'HORA/DATA'] = hora_formatada
        df.to_excel(caminho, index=False)

        print(f"Arquivo Excel atualizado com sucesso em {os.path.abspath(caminho)}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")

    #O PROGRAMA ESPERA 10 SEGUNDOS
    time.sleep(10)
    # ATÉ QUE VOLTE PARA O TOPO DO SITE E COMEÇA O LOOP DENOVO
    driver.execute_script("window.scrollTo(0,0)")
    
