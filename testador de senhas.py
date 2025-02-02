from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import openpyxl
import time

header_ascii = r"""           
  _____                                     _  _______        _            
 |  __ \                                   | ||__   __|      | |           
 | |__) |__ _______  __      _____  _ __ __| ||  | | ___  ___| |_ ________
 |  ___/ _ / __/ __\ \ \ /\ / / _ \| '__/ _| |   | |/ _ \/ __| __/ _ \ '__|
 | |  | (_|\__ \__ \  \ V  V / (_) | | | (_| |   | |  __/\__ \ ||  __/ |   
 |_|  \__,_|___/___/   \_/\_/ \___/|_|  \__,_|   |_|\___||___/\__\___|_|
     _             _____      _    _
    | |           |  __ \    | |  (_)                                      
    | |__  _   _  | |__) |_ _| | ___                                     
    | '_ \| | | | |  ___/ _  | |/ / |                                      
    | |_) | |_| | | |  | (_| |   <| |                                      
    |_.__/ \__, | |_|   \__,_|_|\_\_|                                      
            __/ |                                                          
           |___/   
                                                               
"""
print(header_ascii)

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

workbook = openpyxl.load_workbook('cnpj.xlsx')
sheet = workbook.active

dados = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    cnpj = row[3]  # Coluna D para CNPJ
    senha = row[4] # Coluna E para Senha
    dados.append({
        'cnpj': cnpj,
        'senha': senha
    })

for i, item in enumerate(dados):
    cnpj = item['cnpj']
    senha = item['senha']

    navegador.get('https://nfse.canoas.rs.gov.br/autenticacao/login')

    try:
        element = WebDriverWait(navegador, 10).until(
            EC.visibility_of_element_located((By.XPATH, '//*[@id="idUser"]'))
        )

        element.send_keys(cnpj)
    
    except Exception as e:
        print(f"Erro ao clicar no elemento: {e}")
        
    navegador.find_element('xpath', '//*[@id="idSenha"]').send_keys(senha)
    navegador.find_element('xpath', '//*[@id="btnLogin"]').click()

    try:
        WebDriverWait(navegador, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="MENU_MEU_CADASTRO/cadastros/meucadastro"]'))
        )
        print(f"{cnpj} OK")
        navegador.quit()
        navegador = webdriver.Chrome(service=servico)
        navegador.get('https://nfse.canoas.rs.gov.br/autenticacao/login')
        
    except Exception as e:
        print(f"{cnpj} Senha incorreta")

    time.sleep(2)

input("Aperte enter para finalizar")
