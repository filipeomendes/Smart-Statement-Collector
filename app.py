from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import StaleElementReferenceException
import glob
from time import sleep
import pandas as pd
import pyautogui
import os
import json

#Caminho onde os arquivos serão salvos
download_dir = os.path.join(os.getcwd(), "C:/Users/Arquivos Baixados")

# Configurações do Chrome
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument('--kiosk-printing')  # Ativa a impressão automática

# Configuração para o destino de impressão "Salvar como PDF"
app_state = {
    "recentDestinations": [{
        "id": "Save as PDF",
        "origin": "local",
        "account": ""
    }],
    "selectedDestinationId": "Save as PDF",
    "version": 2
}

# Configurando as preferências
prefs = {
    "download.default_directory": download_dir,  
    "download.prompt_for_download": False,       
    "download.directory_upgrade": True,          
    "safebrowsing.enabled": True,
    "printing.print_preview_sticky_settings.appState": json.dumps(app_state),  # Define "Salvar como PDF" como destino
    "savefile.default_directory": download_dir  # Garante que o PDF seja salvo no diretório especificado
}
chrome_options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=chrome_options)

# Aquivo Excel base
arquivo = pd.read_excel('Base.xlsx')

#Informamos o site que o programa irá entrar
driver.get('https://filipemendes.com.br')

#Na primeira tela do site, procura o botão "--------"
sleep(2)
WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.XPATH, '//*[@id="btn-----"]'))
).click()
sleep(10)
pyautogui.press('enter')
sleep(10)
pyautogui.press('enter')

#Selecionando o menu e escolhendo a opção "Solicitar ---------"
menu_inicial = WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/form/table[-]/select'))
)
dropdown = Select(menu_inicial)
sleep(2)
dropdown.select_by_index(11)

#Preenchendo o ---- e clicando em continuar
menu_base_da_conta = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, '/html/body/form/table[-]/select'))
)
dropdown2 = Select(menu_base_da_conta)
dropdown2.select_by_index(12)

conta_erro_10039 = 0
conta_erro_144 = 0

for index, row in arquivo.iterrows():
    #Colunas do Excel
    np = str(row['ID'])
    nome = str(row['NOME'])
    pis = str(row['PIS'])

    campo_pis = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.NAME, "-----"))
    )
    campo_pis.send_keys(pis)
    btn_continuar = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//img[contains(@src, '-------.gif')]"))
    ).click()
    sleep(0.5)

    

    try:
        nao_localizado = driver.find_element(By.XPATH, "//*[contains(text(), '(10039) - ------- não localizado. Verifique o número digitado ou informe outro atributo de pesquisa.	')]")
        if nao_localizado:
            print(f"O PIS:{pis} do colaborador {nome}-{np} não foi localizado! ERRO 10039.")
            driver.back()  # Volta para a página anterior
            sleep(1)
            campo_pis_alterado = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/form/table[-]/input'))
            )
            campo_pis_alterado.clear()
            conta_erro_10039 += 1
    except:
        try:
            erro_mensagem = driver.find_element(By.XPATH, "//*[contains(text(), '(144) - CONTA LOCALIZADA NAO ATENDE OS CRITERIOS ESTABELECIDOS PARA ACESSO VIA INTERNET')]")
            if erro_mensagem:
                print(f"O PIS:{pis} do colaborador {nome}-{np} não atende os requisitos! ERRO 144.")
                driver.back()  # Volta para a página anterior
                sleep(1)
                campo_pis_alterado = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/form/table[-]/input'))
                )
                campo_pis_alterado.clear()
                conta_erro_144 += 1
        except:
            try:
                visualizar_impressao = driver.find_element(By.XPATH, "//img[contains(@src, '-------.gif')]").click()
                sleep(1)

                # Identifica todas as abas abertas e troca para a aba da impressão
                abas = driver.window_handles
                driver.switch_to.window(abas[-1])

                valor_base = WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/table[-]"))
                ).text 
                print(f"Valor capturado para {nome}: {valor_base}")

                # Preenche na coluna correta
                for col in ['Conta 1 Valor Base', 'Conta 2 Valor Base', 'Conta 3 Valor Base', 'Conta 4 Valor Base', 'Conta 5 Valor Base']:
                    if pd.isna(arquivo.at[index, col]):  # Verifica se a célula está vazia
                        valor_base_limpo = valor_base.replace("R$ ", "")
                        arquivo.at[index, col] = valor_base_limpo  # Atualiza com o valor capturado
                        break

                valor_data = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/table[-]"))
                ).text  # Novo XPath
                print(f"Valor Data capturado para {nome}: {valor_data}")

                # Preenche na coluna de Data
                for col in ['Conta 1 Valor Base', 'Conta 2 Valor Base', 'Conta 3 Valor Base', 'Conta 4 Valor Base', 'Conta 5 Valor Base']:
                    if pd.isna(arquivo.at[index, col]):  # Verifica se a célula está vazia
                        arquivo.at[index, col] = valor_data  # Atualiza com o valor capturado
                        break

                # Ativa a impressão da página e salva como PDF automaticamente
                driver.execute_script('window.print();')
                sleep(2)
                pdf_path = max(glob.glob(os.path.join(download_dir, "*.pdf")), key=os.path.getctime)
                novo_nome_base = f"Extrato - {nome}.pdf"
                novo_caminho = os.path.join(download_dir, novo_nome_base)
                contador = 1
                while os.path.exists(novo_caminho):
                    novo_nome = f"Extrato - {nome}({contador}).pdf"
                    novo_caminho = os.path.join(download_dir, novo_nome)
                    contador += 1
                os.rename(pdf_path, novo_caminho)

                # Fechar a aba de impressão e volta para a aba original
                driver.close()
                driver.switch_to.window(abas[0])

                # Volta duas páginas
                driver.back()
                sleep(0.5)
                campo_pis_alterado = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/form/table[-]/input'))
                )
                campo_pis_alterado.clear()

            except:
                #Pagina onde verificar o número de contas
                btn_continuar = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//img[contains(@src, 'botao_continuar.gif')]"))
                ).click()
                sleep(0.5)

                #Segue para a página antes da impressão
                visualizar_impressao = driver.find_element(By.XPATH, "//img[contains(@src, '------.gif')]").click()
                sleep(1)

                # Identifica todas as abas abertas e troca para a aba da impressão
                abas = driver.window_handles
                driver.switch_to.window(abas[-1])

                valor_base = WebDriverWait(driver, 2).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/table[-]"))
                ).text
                print(f"Valor capturado para {nome}: {valor_base}")

                # Preenche na coluna correta
                for col in ['Conta 1 Valor Base', 'Conta 2 Valor Base', 'Conta 3 Valor Base', 'Conta 4 Valor Base', 'Conta 5 Valor Base']:
                    if pd.isna(arquivo.at[index, col]):  # Verifica se a célula está vazia
                        valor_base_limpo = valor_base.replace("R$ ", "")
                        arquivo.at[index, col] = valor_base_limpo  # Atualiza com o valor capturado
                        break

                valor_data = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/table[-]"))
                ).text  # Novo XPath
                print(f"Valor Data capturado para {nome}: {valor_data}")

                # Preenche na coluna de Data
                for col in ['Conta 1 Valor Base', 'Conta 2 Valor Base', 'Conta 3 Valor Base', 'Conta 4 Valor Base', 'Conta 5 Valor Base']:
                    if pd.isna(arquivo.at[index, col]):  # Verifica se a célula está vazia
                        arquivo.at[index, col] = valor_data  # Atualiza com o valor capturado
                        break

                # Ativa a impressão da página e salva como PDF automaticamente
                driver.execute_script('window.print();')
                sleep(2)
                pdf_path = max(glob.glob(os.path.join(download_dir, "*.pdf")), key=os.path.getctime)
                novo_nome_base = f"{np} - {nome}.pdf"
                novo_caminho = os.path.join(download_dir, novo_nome_base)
                contador = 1
                while os.path.exists(novo_caminho):
                    novo_nome = f"{np} - {nome}({contador}).pdf"
                    novo_caminho = os.path.join(download_dir, novo_nome)
                    contador += 1
                os.rename(pdf_path, novo_caminho)

                # Fechar a aba de impressão e volta para a aba original
                driver.close()
                driver.switch_to.window(abas[0])

                # Volta duas páginas e preenche o novo PIS
                driver.back()
                sleep(0.5)

                #Verifica se o funcionario tem mais contas
                # Agora buscamos todos os inputs de tipo 'radio' que são relacionados às contas
                contas = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//input[@type='radio'][contains(@name, '------')]"))
                )

                # Se houver mais de uma conta, percorre todas as contas e clica nas adicionais
                if len(contas) > 1:
                    if len(contas) > 1:  # Confirma que há mais de uma conta
                        for i in range(1, len(contas)):
                            try:
                                # Garantir que o índice ainda seja válido
                                if i < len(contas):
                                    contas[i].click()
                                    sleep(1)
                                    # Continue com o processo como antes
                                    btn_continuar = WebDriverWait(driver, 5).until(
                                        EC.element_to_be_clickable((By.XPATH, "//img[contains(@src, '-------.gif')]"))
                                    ).click()
                                    sleep(0.5)
                                    visualizar_impressao = driver.find_element(By.XPATH, "//img[contains(@src, '-------.gif')]").click()
                                    sleep(1)
                                    abas = driver.window_handles
                                    driver.switch_to.window(abas[-1])
                                    valor_base = WebDriverWait(driver, 2).until(
                                        EC.presence_of_element_located((By.XPATH, "/html/body/table[-]"))
                                    ).text 
                                    print(f"Valor capturado para {nome}: {valor_base}")

                                    # Preenche na coluna correta
                                    for col in ['Conta 1 Valor Base', 'Conta 2 Valor Base', 'Conta 3 Valor Base', 'Conta 4 Valor Base', 'Conta 5 Valor Base']:
                                        if pd.isna(arquivo.at[index, col]):  # Verifica se a célula está vazia
                                            valor_base_limpo = valor_base.replace("R$ ", "")
                                            arquivo.at[index, col] = valor_base_limpo  # Atualiza com o valor capturado
                                            break

                                    valor_data = WebDriverWait(driver, 5).until(
                                        EC.presence_of_element_located((By.XPATH, "/html/body/table[-]"))
                                    ).text  # Novo XPath
                                    print(f"Valor Data capturado para {nome}: {valor_data}")

                                    # Preenche na coluna de Data
                                    for col in ['Conta 1 Data', 'Conta 2 Data', 'Conta 3 Data', 'Conta 4 Data', 'Conta 5 Data']:
                                        if pd.isna(arquivo.at[index, col]):  # Verifica se a célula está vazia
                                            arquivo.at[index, col] = valor_data  # Atualiza com o valor capturado
                                            break

                                    driver.execute_script('window.print();')
                                    sleep(2)
                                    pdf_path = max(glob.glob(os.path.join(download_dir, "*.pdf")), key=os.path.getctime)
                                    novo_nome_base = f"{np} - {nome}.pdf"
                                    novo_caminho = os.path.join(download_dir, novo_nome_base)
                                    contador = 1
                                    while os.path.exists(novo_caminho):
                                        novo_nome = f"{np} - {nome}({contador}).pdf"
                                        novo_caminho = os.path.join(download_dir, novo_nome)
                                        contador += 1
                                    os.rename(pdf_path, novo_caminho)
                                    driver.close()
                                    driver.switch_to.window(abas[0])
                                    driver.back()
                                    sleep(0.5)
                            except StaleElementReferenceException:
                                # Caso ocorra um erro, tenta atualizar e continuar
                                print(f"Elemento de conta {i} não encontrado, atualizando lista de contas...")
                                contas = WebDriverWait(driver, 5).until(
                                    EC.presence_of_all_elements_located((By.XPATH, "//input[@type='radio'][contains(@name, '-----')]"))
                                )
                                if i < len(contas):
                                    contas[i].click()
                                    sleep(1)
                                    btn_continuar = WebDriverWait(driver, 5).until(
                                        EC.element_to_be_clickable((By.XPATH, "//img[contains(@src, '-------.gif')]"))
                                    ).click()
                                    sleep(0.5)
                                    visualizar_impressao = driver.find_element(By.XPATH, "//img[contains(@src, '------.gif')]").click()
                                    sleep(1)
                                    abas = driver.window_handles
                                    driver.switch_to.window(abas[-1])
                                    valor_base = WebDriverWait(driver, 2).until(
                                        EC.presence_of_element_located((By.XPATH, "/html/body/table[-]"))
                                    ).text
                                    print(f"Valor capturado para {nome}: {valor_base}")

                                    # Preenche na coluna correta
                                    for col in ['Conta 1 Valor Base', 'Conta 2 Valor Base', 'Conta 3 Valor Base', 'Conta 4 Valor Base', 'Conta 5 Valor Base']:
                                        if pd.isna(arquivo.at[index, col]):  # Verifica se a célula está vazia
                                            valor_base_limpo = valor_base.replace("R$ ", "")
                                            arquivo.at[index, col] = valor_base_limpo  # Atualiza com o valor capturado
                                            break

                                    valor_data = WebDriverWait(driver, 5).until(
                                        EC.presence_of_element_located((By.XPATH, "/html/body/table[-]"))
                                    ).text  # Novo XPath
                                    print(f"Valor Data capturado para {nome}: {valor_data}")

                                    # Preenche na coluna de Data
                                    for col in ['Conta 1 Data', 'Conta 2 Data', 'Conta 3 Data', 'Conta 4 Data', 'Conta 5 Data']:
                                        if pd.isna(arquivo.at[index, col]):  # Verifica se a célula está vazia
                                            arquivo.at[index, col] = valor_data  # Atualiza com o valor capturado
                                            break

                                    driver.execute_script('window.print();')
                                    pdf_path = max(glob.glob(os.path.join(download_dir, "*.pdf")), key=os.path.getctime)
                                    novo_nome_base = f"{np} - {nome}.pdf"
                                    novo_caminho = os.path.join(download_dir, novo_nome_base)
                                    contador = 1
                                    while os.path.exists(novo_caminho):
                                        novo_nome = f"{np} - {nome}({contador}).pdf"
                                        novo_caminho = os.path.join(download_dir, novo_nome)
                                        contador += 1
                                    os.rename(pdf_path, novo_caminho)
                                    sleep(2)
                                    driver.close()
                                    driver.switch_to.window(abas[0])
                                    driver.back()
                                    sleep(0.5)

                driver.back()
                campo_pis_alterado = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/form/table[-]/input'))
                )
                campo_pis_alterado.clear()
                
arquivo.to_excel('Base.xlsx', index=False)
print("Os arquivos disponíveis foram gerados com sucesso! :D")
print(f"Houveram {conta_erro_10039} erros de conta não localizada e {conta_erro_144} erros de contas que não atendem os critérios para acesso via internet.")
driver.quit()