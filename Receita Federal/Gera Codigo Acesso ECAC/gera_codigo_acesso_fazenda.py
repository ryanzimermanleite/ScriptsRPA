# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------#
# Nome:     Gera Codigo Acesso Fazenda Ecac                             #
# Arquivo:  gera_codigo_acesso_fazenda.py                               #
# Versão:   1.0.0                                                       #
# Modulo:   Receita Federal                                             #
# Objetivo: Acessa e gera codigo de acesso Ecac                         #
# Autor:    Ryan Zimerman Leite                                         #
# Data:     10/10/2023                                                  #
# ----------------------------------------------------------------------#
import time, re, os
import fitz
import pyautogui
from selenium import webdriver
from PIL import Image
from selenium.webdriver.common import alert
from selenium.webdriver.common.by import By
from sys import path

path.append(r'..\..\_comum')
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from pyautogui_comum import _find_img, _click_img, _wait_img
from chrome_comum import _initialize_chrome, _find_by_id, _find_by_path
from comum_comum import _time_execution, _escreve_relatorio_csv, _escreve_header_csv, _open_lista_dados, \
    _where_to_start, _indice
from captcha_comum import _solve_text_captcha, _solve_hcaptcha

# Variaveis aonde está localizado os arquivos PDFs
lista_2022 = "V:\Setor Robô\Scripts Python\Receita Federal\Gera Codigo Acesso ECAC\Lista\LISTA GERAL - 2022.pdf"
lista_2023 = "V:\Setor Robô\Scripts Python\Receita Federal\Gera Codigo Acesso ECAC\Lista\LISTA GERAL 2023.pdf"

# Função responsavel por verificar se o nome do cliente está nos PDF e extrair o numero do recibo.
def verifica_lista_pdf(cpf):
    if len(cpf) < 11:
        teste = cpf.zfill(11)
    cpf = '{}.{}.{}-{}'.format(cpf[:3], cpf[3:6], cpf[6:9], cpf[9:])
    # 123.456.789-00
    # Abre o PDF 2023
    with fitz.open(os.path.join(lista_2023)) as pdf:
        status_2023 = 'N'
        num_recibo_2023 = ''
        # Percorre o PDF procurando pela variavel nome e extraindo o numero de recibo
        for page in pdf:
            textinho = page.get_text('text', flags=1 + 2 + 8)
            recibo_2023 = re.compile(r'({})\n.+\n.+\n(.+)'.format(cpf)).search(textinho)

            # Se encontrar o nome e o recibo armaze em variaveis
            if recibo_2023:
                num_recibo_2023 = recibo_2023.group(2)
                status_2023 = 'S'
        # Procura na lista de 2022
        with fitz.open(os.path.join(lista_2022)) as pdf:
            status_2022 = 'N'
            num_recibo_2022 = ''
            for page in pdf:
                textinho = page.get_text('text', flags=1 + 2 + 8)
                recibo_2022 = re.compile(r'({})\n.+\n.+\n(.+)'.format(cpf)).search(textinho)

                if recibo_2022:
                    num_recibo_2022 = recibo_2022.group(2)
                    status_2022 = 'S'

    # Se o nome tiver em ambas as listas PDF ele retorna o status de script
    if status_2023 == 'S' and status_2022 == 'S':
        status = 'script'
        print(">>> Nome encontrado na lista de 2023 e 2022")
        return status, num_recibo_2023, num_recibo_2022

    # Se o nome tiver no PDF 2023 e não no PDF 2023 retorna o status de 'anota'
    if status_2023 == 'S' and status_2022 == 'N':
        status = 'anota'
        print(">>> Nome não encontrado na lista de 2022")
        return status, num_recibo_2023, num_recibo_2022

    # Se o nome não tiver na lista PDF 2023 ele retorna status como 'pula'
    if status_2023 == 'N':
        status = 'pula'
        print(">>> Nome não encontrado na lista de 2023")
        return status, num_recibo_2023, num_recibo_2022
    
def login(driver, cpf, data_nasc, senha, num_recibo_2023, num_recibo_2022):
    print('>>> Acessando site')
    timer = 0
    # Tenta carregar o site até encontrar o botão de avançar
    while not _find_by_id('btnAvancar', driver):
        print('🕤')
        try:
            # Tenta acessar o site
            driver.get('https://servicos.receita.fazenda.gov.br/servicos/codacesso/PFCodAcesso.aspx')
        except:
            return driver, 'erro'
        time.sleep(1)
        timer += 1
        if timer > 60:
            return driver, 'erro'

    timer = 0
    # Enquanto não achar o campo de CPF ele tenta carregar o site novamente
    while not _find_by_id('txtCPF', driver):
        print('🕘')
        time.sleep(1)
        timer += 1
        if timer > 120:
            return driver, 'erro'

    # Faz uma lista passando o CPF e Data de NAscimento
    itens = [('txtCPF', cpf),
             ('txtDtNascimento', data_nasc)
             ]
    # Para cada item da liste ele clicka no campo corresponde no site e digita seu valor
    for iten in itens:
        driver.find_element(by=By.ID, value=iten[0]).click()
        driver.find_element(by=By.ID, value=iten[0]).send_keys(iten[1])

    # Pega os dados necessario para resolver o hCaptcha
    data = {'sitekey': 'ed2e35b4-226f-4f8d-a605-42b8fc7d1238',
            'url': 'https://servicos.receita.fazenda.gov.br/servicos/codacesso/PFCodAcesso.aspx'}

    # Procura pelo captcha escondido atravez de regex no codigo da pagina
    id_captcha = re.compile(r'<textarea id=\"(h-captcha-response-.+)\" name=\"h-captcha-response\"').search(
        driver.page_source).group(1)

    # Envia para a API os dados do hCaptcha
    captcha = _solve_hcaptcha(data, visible=True)

    # Digita no campo oculto a resolução do hCaptcha
    driver.execute_script('document.getElementById("{}").innerHTML="{}"'.format(id_captcha, captcha))

    while not driver.find_element(by=By.ID, value='btnAvancar'):
        time.sleep(1)
    # Faz um clique no botão de Avançar
    driver.find_element(by=By.ID, value='btnAvancar').click()
    time.sleep(2)

    '''if re.compile(r'Data de Nascimento não confere.').search(driver.page_source):
        obs = 'Data nascimento errada'
        return driver, 'ok', 'Codigo não gerado', obs'''
    
    obj = driver.switch_to.alert
    obj.accept()

    # Formata as variaveis de recibo extraido do PDF tirando os . e recortando os 2 ultimos numeros
    recibo_2023 = num_recibo_2023.replace('.', '')[:-2]
    recibo_2022 = num_recibo_2022.replace('.', '')[:-2]

    # Cria uma lista passando o numero do recibo de 2022 e 2023, e a Senha extraida do arquivo excel
    itens = [('txtReciboAno1', recibo_2023),
             ('txtReciboAno2', recibo_2022),
             ('txtSenha', senha),
             ('txtConfirmaSenha', senha)
             ]
    while not driver.find_element(by=By.ID, value='btnGerarCodigo'):
        time.sleep(1)
    # Para cada item da lista acima ele da um clique no campo e digita o valor
    for iten in itens:
        driver.find_element(by=By.ID, value=iten[0]).click()
        driver.find_element(by=By.ID, value=iten[0]).send_keys(iten[1])

    # Procura pelo botão de gerar codigo e da um clique nele
    driver.find_element(by=By.ID, value='btnGerarCodigo').click()
    time.sleep(1)

    # Pega o codigo de acesso gerado no site atravez de regex
    codigo = re.compile(r'(class="spanCodigo">)(.+)(<\/span>)').search(driver.page_source)
    # Se achar o codigo ele captura, se não achar ele provavelmente deu que o Recibo não confere e retorna
    if codigo:
        num_codigo = codigo.group(2)
    else:
        recibo_nao_confere = re.compile(r'Recibo não confere.').search(driver.page_source)
        obs = 'Recibo não confere'
        return driver, 'ok', 'Codigo não gerado', obs
    # Caso gere o codigo ele retorna 'ok' e variavel do numero do codigo
    return driver, 'ok', num_codigo, 'Codigo gerado'

@_time_execution
def run():
    # opções para fazer com que o chrome trabalhe em segundo plano (opcional)
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    #options.add_argument('--window-size=1366,768')
    # options.add_argument("--start-maximized")

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
        nome, cpf, data_nasc, codigo, senha = empresa

        # Chama a função que verifica se o nome da planilha de dados está nos PDFS
        status, num_recibo_2023, num_recibo_2022 = verifica_lista_pdf(cpf)

        # Se a função verifica_lista_pdf retornar como status 'anota' quer dizer quem
        # o nome do cliente tem no pdf de 2023 mas nao no pdf de 2022
        if status == 'anota':
            obsv = 'Não encontrado na Lista 2022'
            num_codigo = 'Codigo não gerado'
            # Anota em um excel separado
            _escreve_relatorio_csv(f'{nome};{cpf};{data_nasc};{num_codigo};{senha};{obsv}',
                                   nome='Listagem Clientes')

        if status == 'pula':
            obsv = 'Não encontrado na Lista 2023'
            num_codigo = 'Codigo não gerado'
            # Anota em um excel separado
            _escreve_relatorio_csv(f'{nome};{cpf};{data_nasc};{num_codigo};{senha};{obsv}',
                                   nome='Listagem Clientes')

        # Se a função verifica_lista_pdf retornar como status 'script' quer dizer que
        # o nome do cliente tem em ambos os PDFs 2023 e 2022 assim faz a excução do script completo
        if status == 'script':
            while True:
                # iniciar o driver do chrome
                    status, driver = _initialize_chrome(options)

                    # coloca um timeout de 60 segundos para que o robô não fique esperando eternamente caso o site não carregue
                    driver.set_page_load_timeout(60)

                    # faz login no site
                    driver, resultado, num_codigo, obsv = login(driver, cpf, data_nasc, senha, num_recibo_2023, num_recibo_2022) # se não der erro no login, sai do while e realiza a consulta
                    if resultado != 'erro':
                        break
                    driver.close()

            print(f'❕ {resultado}')
            # Escreve os dados da consulta no excel
            _escreve_relatorio_csv(f'{nome};{cpf};{data_nasc};{num_codigo};{senha};{obsv}',
                                   nome='Listagem Clientes')
            driver.close()
    # Cria o cabeça do excel
    _escreve_header_csv('NOME;CPF;DATA DE NASCIMENTO;CODIGO ACESSO ECAC;SENHA ECAC;OBSERVACAO',
                        nome='Listagem Clientes')

if __name__ == '__main__':
    run()