# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------#
# Nome:     Consulta Certidão Negativa Campinas                         #
# Arquivo:  consulta_certidao_negativa_campinas.py                      #
# Versão:   1.0.0                                                       #
# Modulo:   Serviços Prefeitura                                         #
# Objetivo: Acessa e anota se possui ou não debitos                     #
# Autor:    Ryan Zimerman Leite                                         #
# Data:     17/11/2023                                                  #
# ----------------------------------------------------------------------#
from selenium import webdriver
from selenium.webdriver.common.by import By
import pyperclip, pyautogui as p
import os, time, re

from sys import path
path.append(r'..\..\_comum')
from captcha_comum import _solve_text_captcha
from chrome_comum import _initialize_chrome, _send_input, _find_by_id, _find_by_path
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, _where_to_start, _indice
from pyautogui_comum import _find_img, _click_img, _wait_img

def login(driver, cnpj, download_folder):
    # Faz o login no site
    print('>>> Acessando site')
    timer = 0
    timer_erro = 0
    try:
        # Tenta acessar o site
        driver.get('https://certidaoqualquerorigem.campinas.sp.gov.br/processos')
    except:
        return driver, 'erro'
    
    # ENQUANTO NÂO LOCALIZAR O BOTÂO DE DROPDOWN INICIAL ELE TENTA ABRIR O SITE
    while not _find_by_id('28778', driver):
        print('🕤')
        time.sleep(1)
        timer += 1
        timer_erro += 1
        
        if timer > 60:
            timer = 0
            try:
                # Tenta acessar o site lixo
                driver.get('https://certidaoqualquerorigem.campinas.sp.gov.br/processos')
            except:
                return driver, 'erro'
            
        if timer_erro > 60:
            return driver, 'erro'

    # clique no elemento da lista suspensa para abrir as opções
    driver.find_element(by=By.ID, value="s2id_28778").click()
    time.sleep(1)

    # FAZ UM CLIQUE NA OPCAO DE CNPJ
    while not _find_by_id('28780', driver):
        time.sleep(1)
        _click_img('cnpj.png', conf=0.9)

    time.sleep(1)

    # DIGITA NO CAMPO CNPJ
    driver.find_element(by=By.ID, value="28780").click()
    time.sleep(0.5)
    pyperclip.copy(cnpj)
    time.sleep(1)
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)

    # Faz a chamada para API resolver o captcha
    captcha = _solve_text_captcha(driver, '/html/body/div[2]/div[3]/div[1]/div/div[3]/div[9]/img', 'xpath')

    # Se não encontrar retorna erro e tenta dnv
    if not captcha:
        print('❌ Erro Login - não encontrou captcha')
        return driver, 'erro captcha'

    # Se resolver o captcha escreve o captcha e da um click em emitir certidão
    try:
        _send_input('28781', captcha, driver)
        driver.find_element(by=By.ID, value='28782').click()
        time.sleep(2)
        # SE APARECER A MENSAGEM DE CAPTCHA ERRADO ELE DA ERRO E TENTA DENOVO
        if _find_by_id('ui-id-1', driver):
            p.press('esc')
            return driver, 'erro'
        print('>>> Fazendo Login')
    except:
        print('❌ Erro ao validar a consulta, tentando novamente')
        return driver, 'erro'

    # Enquanto ele não encontrar o botão de voltar ele espera de 1 em 1 segundo
    while not _find_by_id('28739', driver):
        time.sleep(1)
        if _find_by_id('28727', driver):
            break
        elif _find_by_id('28793', driver):
            break

    # LOCALIZA EM REGEX SE POSSUI OU NAO ARQUIVO DE CERTIDÃO
    msg = re.compile('NÃO FOI POSSIVEL EMITIR A CERTIDÃO POR UM DOS MOTIVOS ABAIXO:').search(driver.page_source)
    msg2 = re.compile('CERTIDÃO NEGATIVA QUALQUER ORIGEM PJ.pdf').search(driver.page_source)
    msg3 = re.compile('CERTIDÃO POSITIVA COM EFEITOS DE NEGATIVA DE DÉBITOS DE QUALQUER ORIGEM PJ.pdf').search(driver.page_source)
    time.sleep(1)

    if (msg):
        # Se não encontrar o arquivo PDF então anota na planilha
        return driver, 'Constam débitos vencidos!'

    elif (msg2):
        # Faz o download do arquivo
        driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[3]/div[1]/div/div[3]/div[5]/table/tbody/tr/td/div[4]/table/tbody/tr[2]/td/div/table/tbody/tr/td[2]/a/img').click()
        time.sleep(2)
        # Renomeia o arquivo para o CNPJ da empresa
        pasta_arquivo = download_folder + '\CERTIDÃO NEGATIVA QUALQUER ORIGEM PJ.pdf'
        renomear_arquivo = download_folder + '\\' + cnpj + '.pdf'
        os.rename(pasta_arquivo, renomear_arquivo)
        return driver, 'Não possui débitos!'
    
    elif (msg3):
        driver.find_element(by=By.XPATH, value='/html/body/div[2]/div[3]/div[1]/div/div[3]/div[5]/table/tbody/tr/td/div[4]/table/tbody/tr[2]/td/div/table/tbody/tr/td[2]/a/img').click()
        time.sleep(2)
        pasta_arquivo = download_folder + '\CERTIDÃO POSITIVA COM EFEITOS DE NEGATIVA DE DÉBITOS DE QUALQUER ORIGEM PJ.pdf'
        renomear_arquivo = download_folder + '\\' + cnpj + '.pdf'
        os.rename(pasta_arquivo, renomear_arquivo)
        return driver, 'Certidão positiva com efeitos de negativa'

@_time_execution
def run():
    # DEFINE A PASTA AONDE VÃO OS ARQUIVOS PDF
    download_folder = "V:\\Setor Robô\\Scripts Python\\Serviços Prefeitura\\Consulta Certidão Negativa Campinas\\execução\\Certidões"
    # CRIA A PASTA DEFINIDA ACIMA
    os.makedirs(download_folder, exist_ok=True)
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    # options.add_argument('--window-size=1366,768')
    options.add_argument("--start-maximized")
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_experimental_option('prefs', {
        "download.default_directory": download_folder,  # muda o diretório padrão de download do navegador
        "download.prompt_for_download": False,  # faz o download automatico sem perguntar onde salvar
        "download.directory_upgrade": True,  # atualiza o diretório de download padrão do navegador
        "plugins.always_open_pdf_externally": True  # não irá abrir o PDF no navegador
    })

    # Abre a planilha de dados selecionada
    empresas = _open_lista_dados()
    if not empresas:
        return False

    # Da a opção de continuar a excecução anterior a partir do ultimo indice
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    # Percorre a planilha de dados atribuindo seus valores nas variaveis
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa, index)
        cnpj, nome = empresa

        while True:
            # iniciar o driver do chrome
            status, driver = _initialize_chrome(options)

            # coloca um timeout de 60 segundos para que o robô não fique esperando eternamente caso o site não carregue
            driver.set_page_load_timeout(60)

            # faz login no site
            driver, resultado = login(driver, cnpj, download_folder)  # se não der erro no login, sai do while e realiza a consulta
            if resultado != 'erro':
                break
            driver.close()

        print(f'❕ {resultado}')
        # Escreve os dados da consulta no excel
        _escreve_relatorio_csv(f'{cnpj};{nome};{resultado}',
                               nome='Resultado Certidão Negativa')
        driver.close()

# Cria o cabeça do excel
    _escreve_header_csv('CNPJ;NOME;STATUS',
                        nome='Resultado Certidão Negativa')

if __name__ == '__main__':
    run()