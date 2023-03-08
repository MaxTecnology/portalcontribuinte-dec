# Importações pra ler arquivo excel
import pandas as pd
import os
# import shutil
# Importações para automatizar a web
from selenium import webdriver  # Navegador
from selenium.webdriver.common.by import By  # Achar os elementos
from selenium.webdriver.common.keys import Keys  # Digitar teclado na web
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager  # Google Chrome
# from webdriver_manager.firefox import GeckoDriverManager # Firefox
from selenium.webdriver.chrome.service import Service  # Google Chrome Web Driver
# from selenium.webdriver.firefox.service import Service # Firefox Web Driver
# Verificar se o elemento existe
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.select import Select

import time
from PIL import Image

# 0 - Ler arquivo Excel
nome_do_arquivo = "empresas.xlsx"
df = pd.read_excel(nome_do_arquivo)

# Opções do navegador
chrome_options = Options()
arguments = ['--lang=pt-BR', '--start-maximized', '--disable-gpu']

for argument in arguments:
    chrome_options.add_argument(argument)

caminho_download = 'Portal Contribuinte Download - DEC - Real e Presumido'
chrome_options.add_experimental_option("prefs", {


    "plugins.always_open_pdf_externally": True,
    "download.open_pdf_in_system_reader": False,
    "profile.default_content_settings.popups": 0,
    # Visualizar PDF
    "download.default_directory": caminho_download,
    # Atualiza diretorio para diretorio a cima
    "download.directory_upgrade": True,
    # Seta se o navegador deve pedir ou não para fazer download
    "download.prompt_for_download": False,
    # Perguntar ao usuario onde vai ser salvo
    "profile.default_content_setting_values.notifications": 2,
    # desabilitar notificação
    "profile.default_content_setting_values.automatic_downloads": 1,
    # Multiplos downloads

})

servico = Service(ChromeDriverManager().install())
#servico = Service(GeckoDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=chrome_options)
#navegador = webdriver.Firefox(service=servico)
DATA_INI = input('Digite a data inicial (Somente números) : ')
DATA_FIM = input('Digite a data final (Somente números) : ')
site = "https://contribuinte.sefaz.al.gov.br/#/"

# Entrar no Site
navegador.get(site)
navegador.implicitly_wait(40)  # aguardar carregar site
# time.sleep(5)
# 1 - Loopar arquivo
for index, row in df.iterrows():
    print("Index: " + str(index) +
          " A empresa: " + str(row["EMPRESA"]) + " Login: " +
          str(row["LOGIN"]) + " Senha: " +
          str(row["SENHA"])


          )
    #   1.1 - Preencher dados lidos para cada linha no navegador

    navegador.get(site)
    navegador.implicitly_wait(40)  # aguardar carregar site
    time.sleep(5)

    # Clicar no Botão Login
    navegador.find_element(
        By.XPATH, '/html/body/div[3]/div[1]/div/div/div[3]/div/div[1]/a').click()
    # fazer login
    navegador.find_element(
        By.XPATH, '//*[@id="username"]').send_keys(row['LOGIN'])
    navegador.find_element(
        By.XPATH, '//*[@id="password"]').send_keys(row['SENHA'])
    navegador.find_element(
        By.XPATH, '/html/body/div[1]/div/div/div[2]/div/div[3]/form/button').click()
    # Emitir Relatório.
    # navegador.implicitly_wait(10)  # aguardar carregar site
    time.sleep(5)
    # Notas Fiscais de Entrada
    navegador.find_element(
        By.XPATH, '//*[@id="link-relatorio-notas-fiscais-entradas-dfe"]').click()
    # navegador.implicitly_wait(30)
    time.sleep(5)
    cnpj_dropdown = navegador.find_element(
        By.XPATH, '/html/body/jhi-main/div[2]/div/jhi-relatorio-contribuinte/div/div/div[1]/div/select').click()  # Campo CNPJ
    navegador.find_element(
        By.XPATH, '/html/body/jhi-main/div[2]/div/jhi-relatorio-contribuinte/div/div/div[1]/button').click()
    time.sleep(2)
    # navegador.find_element(
    #     By.XPATH, '/html/body/jhi-main/div[2]/div/jhi-relatorio-contribuinte/div/div/div[2]/div[1]/table/tbody/tr/td/div/button/span').click()
    # time.sleep(10)
    # navegador.find_element(
    #     By.XPATH, '/html/body/jhi-main/div[2]/div/jhi-relatorio-contribuinte/div/div/div[2]/div[2]/table/tbody/tr/td/div/button/span').click()
    # time.sleep(10)
    # navegador.find_element(
    #     By.XPATH, '/html/body/jhi-main/div[2]/div/jhi-relatorio-contribuinte/div/div/div[2]/div[3]/table/tbody/tr/td/div/button/span').click()
    # time.sleep(10)

    navegador.find_element(
        By.XPATH, '//*[@id="dataInicial"]').send_keys(DATA_INI)  # Data inicial
    time.sleep(1)
    navegador.find_element(
        By.XPATH, '//*[@id="dataFinal"]').send_keys(DATA_FIM)  # Vencimento final
    time.sleep(2)
    # formato_dropdown = navegador.find_element(
    #     By.XPATH, '//*[@id="tipoRelatorio"]')  # Campo Formato
    # opcoes = Select(formato_dropdown)
    # opcoes.select_by_index(1)
    # time.sleep(2)
    # Salvar janela atual
    janela_inicial = navegador.current_window_handle
    # Ação pra abrir outra aba
    navegador.find_element(
        By.XPATH, '/html/body/jhi-main/div[2]/div/jhi-relatorio-contribuinte/div/div/div[2]/div[4]/table/tbody/tr[1]/td/div/button/span').click()  # Botão Gerar Relatório
    time.sleep(5)
    
    # Remover Linha da planilha
    df = df.drop(index)
    df.to_excel('empresas.xlsx', index=False)

    # Verificar quais janelas estão abertas agora
    # janelas = navegador.window_handles
    # # Alterando foco para outra janela, se existir
    # for janela in janelas:
    #     print(janela)
    #     if janela not in janela_inicial:
    #         navegador.switch_to.window(janela)

    # navegador.close()
    # navegador.switch_to.window(janela_inicial)

    # Sair do portal
    navegador.find_element(
        By.XPATH, '//*[@id="account-menu"]').click()  # Botão Usuraio
    navegador.find_element(
        By.XPATH, '//*[@id="logout"]').click()  # Botão Sair

    # Remover Linha da planilha
    df = df.drop(index)
    df.to_excel('empresas.xlsx', index=False)
