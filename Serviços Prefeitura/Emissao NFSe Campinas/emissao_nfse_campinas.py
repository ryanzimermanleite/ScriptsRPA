# -*- coding: utf-8 -*-
# -----------------------------------------------------------------------------#
# Nome:     Envia NFSe Campinas                                                #
# Arquivo:  emissao_nfse_campinas.py                                           #
# Versão:   1.0.0                                                              #
# Modulo:   Serviços Prefeitura                                                #
# Objetivo: Acessa e Envia NFSe de Campinas                                    #
# Autor:    Ryan Zimerman Leite                                                #
# Data:     10/10/2023                                                         #
# Script executa até certo momento dps para e da erro ai rodar novamente       #
# #
# -----------------------------------------------------------------------------#
import time, re, os
import pyautogui as p
from sys import path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from chrome_comum import _initialize_chrome, _send_input, _find_by_id, _find_by_path
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, \
    _where_to_start, _indice
from captcha_comum import _solve_text_captcha, _solve_hcaptcha

dados = "V:\Setor Robô\Scripts Python\Serviços Prefeitura\Emissao NFSe Campinas\dados.txt"
f = open(dados, 'r', encoding='utf-8')
user = f.read()
user = user.split('/')

def login(driver):

    print('>>> Acessando site')
    # aguarda o botão de habilitar a consultar aparecer
    timer = 0
    # Enquanto não achar o botão Acesso ao Sistema
    while not _find_by_path('/html/body/div[2]/div/div/div[1]/div/div[2]/ul/li[2]/a', driver):
        print('🕤')
        # abre o site da consulta e caso de erro é porque o site demorou pra responder,
        # nesse caso retorna um erro para tentar novamente
        try:
            driver.get('https://issdigital.campinas.sp.gov.br/NotaFiscal/index.php')
            # Encontra o elemento do frame pelo nome ou índice
            frame = driver.find_element(by=By.ID, value='principal')
            # frame = driver.find_element_by_index(0)

            # Alterna o driver para o contexto do frame
            driver.switch_to.frame(frame)
        except:
            return driver, 'erro'
        time.sleep(1)
        timer += 1
        if timer > 60:
            return driver, 'erro'

    # clica Acesso ao Sistema
    driver.find_element(by=By.XPATH, value='/html/body/div[2]/div/div/div[1]/div/div[2]/ul/li[2]/a').click()

    # aguarda o campo de CPF aparecer
    timer = 0
    while not _find_by_id('rLogin', driver):
        print('🕘')
        time.sleep(1)
        timer += 1
        if timer > 120:
            return driver, 'erro'

    # lista de campos para preencher
    login = user[0]
    senha = user[1]
    itens = [('rLogin', login),
             ('rSenha', senha)]

    # Faz um clique e digita no campo o login e senha
    for iten in itens:
        driver.find_element(by=By.ID, value=iten[0]).click()
        driver.find_element(by=By.ID, value=iten[0]).send_keys(iten[1])

    # Faz a chamada para API resolver o captcha
    captcha = _solve_text_captcha(driver, 'captcha_image')

    # Se não encontrar retorna erro e tenta dnv
    if not captcha:
        print('❌ Erro Login - não encontrou captcha')
        return driver, 'erro captcha'

    # Se resolver o captcha escreve o captcha e da um click
    try:
        _send_input('cap_text', captcha, driver)
        driver.find_element(by=By.ID, value='btnEntrar').click()
    except:
        print('❌ Erro ao validar a consulta, tentando novamente')
        return driver, 'erro'

    print('>>> Fazendo Login')

    # aguarda as informações do cadastro aparecerem, se demorar mais de 1 minuto ou a resposta do captcha estiver errada,
    # retorna um erro e tenta novamente
    timer = 0

    while not _find_by_path('/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/div[4]/ul/li/a',
                            driver):
        time.sleep(1)
        timer += 1
        if re.compile(r'Código de Verificação inválido.').search(driver.page_source):
            print('❌ A resposta não corresponde ao desafio gerado, tentando novamente')
            return driver, 'erro'
        if timer > 60:
            print('❌ O site demorou muito para responder, tentando novamente')
            return driver, 'erro'

    driver.find_element(by=By.XPATH, value='/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/div[4]/ul/li/a').click()

    time.sleep(3)
    while not _find_by_path('/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/div[8]/ul/li/a',driver):
        time.sleep(1)

    driver.find_element(by=By.XPATH, value='/html/body/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr[3]/td[2]/div[8]/ul/li/a').click()

    while not _find_by_id('rTomCpfCnpjSel', driver):
        print('🕘')
        time.sleep(1)
        timer += 1
        if timer > 120:
            return driver, 'erro'

    return driver, 'ok'
def consulta(driver, cod, valor, cnpj):
    cnpj = cnpj.replace('.', '').replace('/', '').replace('-', '')
    _send_input('rTomCpfCnpjSel', cnpj, driver)
    driver.find_element(by=By.ID, value='btnTomador').click()
    time.sleep(1)

    # BOTAO AVANÇAR
    driver.execute_script('document.getElementById("btnAvancar").click()')
    time.sleep(2)

    # Possuiveis mensagem de erro
    msg_erro_razao = re.compile('Razão Social não informada').search(driver.page_source)
    msg_erro_CEP = re.compile('CEP não informado').search(driver.page_source)
    msg_erro_razao_tipo_Logradouro = re.compile('Tipo de Logradouro não informado').search(driver.page_source)
    msg_erro_razao_Logradouro = re.compile('Logradouro não informado').search(driver.page_source)
    msg_erro_razao_numero = re.compile('Informe o Numero ou Complemento!').search(driver.page_source)
    msg_erro_razao_tipo_bairro = re.compile('Tipo de Bairro não informado').search(driver.page_source)
    msg_erro_razao_bairro = re.compile('Bairro não informado').search(driver.page_source)
    msg_erro_razao_email = re.compile('E-mail não informado').search(driver.page_source)
    msg_erro_site_parou = re.compile('Erro na consulta ao serviço de verificação da tributação especial.').search(driver.page_source)

    # Se encontrar alguma dessas mensagem ele printa e da um return
    if (msg_erro_razao or msg_erro_CEP or
            msg_erro_razao_tipo_Logradouro or
            msg_erro_razao_Logradouro or
            msg_erro_razao_numero or
            msg_erro_razao_tipo_bairro or
            msg_erro_razao_bairro or
            msg_erro_razao_email):
        print('>>> Empresa Nova, Nota Não Emitida!')
        return driver, str('Nota Não Emitida')

    # Se ele encontrar alguma mensagem de erro retorna False
    if msg_erro_site_parou:
        return False
        
     #  Message: no such element: Unable to locate element: {"method":"css selector","selector":"[id="rCodAtv"]"}
    # TRATAR ESSE ERRO
    time.sleep(2)
    timer = 0
    # Localize o elemento do dropdown DE ATIVIDADE
    while not _find_by_id('rCodAtv', driver):
        print('🕘')
        time.sleep(1)
        timer += 1
        if timer > 120:
            return driver, 'erro'

    # Seleciona o codigo da atividade
    atividade = Select(driver.find_element(by=By.ID, value='rCodAtv'))
    atividade.select_by_value('S620230002')
    time.sleep(1)
    
    timer = 0
    # Enquanto ele não achar o campo cidade ele espera
    while not _find_by_id('rCidadeCodigoPrestacao', driver):
        print('🕘')
        time.sleep(1)
        timer += 1
        if timer > 120:
            return driver, 'erro'

    # Localize o elemento do dropdown DE LOCAL DA PRESTAÇÂO
    while not driver.find_element(by=By.ID, value='rCidadeCodigoPrestacao'):
        time.sleep(1)
    time.sleep(5)
    while True:
        try:
            # Seleciona a cidade de Campinas
            local = Select(driver.find_element(by=By.ID, value='rCidadeCodigoPrestacao'))
            local.select_by_value('6291')
            break
        except:
            print('erro 6291')
            pass
    time.sleep(3)

    # Clcika no botão de avançar
    driver.execute_script('document.getElementById("btnAvancar").click()')
    time.sleep(2)

    # Escreve essa descrição no campo de descrição
    _send_input('rItemDescricao', 'Referente a utilização do emissor', driver)
    time.sleep(0.5)

    # Digita 1 para quantidade do item
    _send_input('rItemQtd', 1, driver)
    time.sleep(0.5)

    # Digita o valor do item se baseando na tabela do excel
    _send_input('rItemValUnit', valor, driver)
    time.sleep(0.5)

    # Clicka no botão para adicionar
    driver.find_element(by=By.ID, value='btnAdd').click()
    time.sleep(1)
    driver.find_element(by=By.ID, value='btnAdd').click()
    time.sleep(1)

    # Faz um clicke no botão de emitir nota fiscal
    driver.execute_script('document.getElementById("btnEmitir").click()')
    time.sleep(2)

    # Procura pelo botão que aparece no modal para confirma a emissão
    while not _find_by_id('btnConfirmaModal', driver):
        time.sleep(1)
    time.sleep(2)
    # Faz um clique no botão
    driver.execute_script('document.getElementById("btnConfirmaModal").click()')
    time.sleep(4)

    # Clicka para fechar o modal para ir a proxima consulta
    while not _find_by_id('btnFechaModalSucesso', driver):
        time.sleep(1)
    #driver.find_element(by=By.ID, value='btnFechaModalSucesso').click()

    # Faz um clique usando javascript para fechar o modal de sucesso
    try:
        driver.execute_script('document.getElementById("btnFechaModalSucesso").click()')
    except:
        pass
    return driver, 'Nota Emitida Com Sucesso'

# Função para verificar os dados da planilha excel
def verifica_dados(cod, valor, cnpj):
    infos = [(cod, 'Codigo'),
             (valor, 'Valor'),
             (cnpj, 'CNPJ')]

    for info in infos:
        if info[0] == '':
            print(f'❌ {info[1]} não informado')
            _escreve_relatorio_csv(f'{cod};{valor};{cnpj};{info[1]} não informado', nome='Emissao NFSe Campinas')
            return False
    return True
@_time_execution
def run():
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    #options.add_argument('--window-size=1366,768')
    options.add_argument("--start-maximized")

    empresas = _open_lista_dados()
    if not empresas:
        return False

    # Da a opção de continuar a excecução anterior a partir do ultimo indice
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return False

    # Percorre a planilha de dados atribuindo seus valores nas variaveis
    total_empresas = empresas[index:]

    while True:
        # iniciar o driver do chrome
        try:
            status, driver = _initialize_chrome(options)

            # coloca um timeout de 60 segundos para que o robô não fique esperando eternamente caso o site não carregue
            driver.set_page_load_timeout(60)

            # faz login no site
            driver, resultado = login(driver)
            # se não der erro no login, sai do while e realiza a consulta
            if resultado != 'erro':
                break
            driver.close()
        except:
            pass

    # Para cada empresa da planilha do excel ele faz a consulta
    for count, empresa in enumerate(empresas[index:], start=1):
        _indice(count, total_empresas, empresa, index)
        cod, valor, cnpj = empresa

        # verifica se não tem nenhum dado faltando
        if not verifica_dados(cod, valor, cnpj):
            continue
        while True:
            driver, resultado = consulta(driver, cod, valor, cnpj)
            if resultado:
                break

        # Escreve o resultado no relatorio em excel
        print(f'❕ {resultado}')
        _escreve_relatorio_csv(f'{cod};{valor};{cnpj};{resultado}',
                               nome='Emissao NFSe Campinas')

    _escreve_header_csv('COD;VALOR;CNPJ;STATUS',
                        nome='Emissao NFSe Campinas')
    driver.close()

if __name__ == '__main__':
    run()