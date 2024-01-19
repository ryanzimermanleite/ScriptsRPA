import time
# -*- coding: utf-8 -*-
import fitz, re
import os, pyperclip
import time, shutil, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status

def run():

    while True:
        abre_tela_inicial()
        status_gravar = exibe_gravar_declaracao()

        if status_gravar == 'ok':
            status = gravar_declaracao()
            if status == 'pendencia':
                imprimir_pendencia()
                excluir()
                renomeia_arquivo_pendencia()
            elif status == 'ok':
                status = transmitir()
                if status == 'ok':
                    salvar_recibo()
                    salvar_pdf()
                    renomeia_arquivo()
                    excluir()
                    move_arquivo_RFB()
                elif status != 'ok':
                    p.press('enter')
                    time.sleep(0.5)
                    p.press('esc')
                    time.sleep(0.5)
                    p.press('esc')
                    excluir()
                    move_arquivo_RFB_erro(status)
        elif status_gravar == 'Sem Declaração':
            print('Fim da execução!')
            break

def abre_tela_inicial():
    while not _find_img('tela_inicial.png', conf=0.9):
        time.sleep(1)
    _click_img('tela_inicial.png', conf=0.9)

def exibe_gravar_declaracao():
    timer = 0
    while not _find_img('gravar_declaracao.png', conf=0.9):
        p.hotkey('ctrl', 'g')
        time.sleep(1)
        timer += 1
        if timer > 5:
            return 'Sem Declaração'
    time.sleep(1)
    return 'ok'

def selecionar_primeiro():
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('up')
    time.sleep(1)

def gravar_declaracao():
    time.sleep(1)
    if _find_img('imagem1.png', conf=0.9):
        p.press('right')
    else:
        selecionar_primeiro()

    p.hotkey('alt', 'o')

    while not _find_img('verificando_pendencias.png', conf=0.9):
        time.sleep(1)
    time.sleep(1)
    p.hotkey('alt', 'g')
    while not _find_img('sucesso_gravar.png', conf=0.9):
        time.sleep(1)
        if _find_img('pendencia.png', conf=0.9):
            return 'pendencia'
        time.sleep(0.5)

    p.press('enter')
    time.sleep(0.5)
    p.press('esc')
    print('Gravado com sucesso!')
    return 'ok'

def imprimir_pendencia():
    while not _find_img('relacao_pendencia.png', conf=0.9):
        p.hotkey('alt', 'p')
        time.sleep(1)

    p.hotkey('alt', 'i')


    while not _find_img('salvar_imagem.png', conf=0.9):
        time.sleep(1)
    os.makedirs(r'{}\{}'.format(e_dir, 'Recibos'), exist_ok=True)
    p.write('Pendencia')

    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy('V:\Setor Robô\Scripts Python\DCTF Mensal\Envia DCTF e Salvar Recibo\{}\{}'.format(e_dir, 'Recibos'))
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(3)
    print('✔ Pendencia Salvo')
    p.hotkey('alt', 'f')
    time.sleep(1)
    p.hotkey('alt', 'c')
    time.sleep(1)
    p.press('esc')

def transmitir():
    while not _find_img('transmitir.png', conf=0.9):
        p.hotkey('ctrl', 'i')
        time.sleep(1)
    time.sleep(1)

    p.hotkey('alt', 'o')

    while not _find_img('transmitir_via_internet.png', conf=0.9):
        time.sleep(1)
    time.sleep(1)

    p.press('right')
    time.sleep(0.5)
    p.press('tab', presses=5, interval=0.2)
    time.sleep(0.5)
    p.press('space')
    time.sleep(0.5)

    p.hotkey('alt', 'o')
    assinar()

    while not _find_img('transmitido_sucesso.png', conf=0.9):
        time.sleep(1)
        if _find_img('erro_2_mes.png', conf=0.9):
            return 'DCTF a partir do 2º mês'
        elif _find_img('erro_cadastro_rfb.png', conf=0.9):
            return 'CPF/CNPJ diferente no cadastro da RFB'
        elif _find_img('erro_dctf_ativa.png', conf=0.9):
            return 'DCTF ativa referente a PA anterior do mesmo ano-calendário'
        elif _find_img('erro_nao_aplica.png', conf=0.9):
            return 'Opção Não se aplica no campo Critério de Reconhecimento...'
        elif _find_img('erro_procuracao.png', conf=0.9):
            return 'Não existe procuração eletrônica'
        elif _find_img('erro_real_estimativa.png', conf=0.9):
            return 'Tributação igual a Real/Estimativa'
        elif _find_img('erro_irpj.png', conf=0.9):
            return 'Tributação igual a Isenta do IRPJ'
        elif _find_img('erro_presumido.png', conf=0.9):
            return 'Tributação igual a Presumido'
        elif _find_img('erro_certificado.png', conf=0.9):
            return 'Certificado Digital Expirou'
        elif _find_img('erro_certificado_cancelado.png', conf=0.9):
            return 'Certificado Digital Foi Cancelado'
        elif _find_img('erro_ja_consta.png', conf=0.9):
            return 'Já Consta uma DCTF Normal'
        elif _find_img('erro_dispensadas.png', conf=0.9):
            return 'Naturezas Jurídicas Estão Dispensadas da Apresentação da DCTF'
        
        time.sleep(1)
    time.sleep(1)
    p.press('enter')
    time.sleep(0.5)
    
    p.press('esc', presses=2, interval=0.3)
    return 'ok'
def assinar():
    while not _find_img('assinar.png', conf=0.9):
        time.sleep(1)
    time.sleep(1)

    p.press('up', presses=3, interval=0.2)
    time.sleep(0.5)
    p.hotkey('alt', 'a')

def salvar_recibo():
    while not _find_img('imprimir_recibo.png', conf=0.9):
        p.hotkey('ctrl', 'b')
        time.sleep(1)
    time.sleep(1)

    p.hotkey('alt', 'o')

    while not _find_img('imprimir_recibo2.png', conf=0.9):
        time.sleep(1)
    time.sleep(1)

    p.press('right')
    time.sleep(0.5)
    p.hotkey('alt', 'o')

def salvar_pdf():
    while not _find_img('salvar_imagem.png', conf=0.9):
        time.sleep(1)
    time.sleep(1)
    os.makedirs(r'{}\{}'.format(e_dir, 'Recibos'), exist_ok=True)
    p.write('Arquivo')

    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy('V:\Setor Robô\Scripts Python\DCTF Mensal\Envia DCTF e Salvar Recibo\{}\{}'.format(e_dir,'Recibos'))
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(3)
    print('✔ Recibo Salvo')
    p.press('esc', presses=2, interval=0.3)

def renomeia_arquivo():
    
    arq_nome = 'Arquivo.pdf'
    pasta = 'V:\Setor Robô\Scripts Python\DCTF Mensal\Envia DCTF e Salvar Recibo\execução\Recibos'
    with fitz.open(pasta + '\\' + arq_nome) as pdf:
        # Percorre o PDF procurando pela variavel denominacao e extraindo os dados de creditos debitos e saldo
        for page in pdf:
            textinho = page.get_text('text', flags=1 + 2 + 8)
            cnpj = re.compile(r'(CNPJ: (.+))').search(textinho)

            if cnpj:
                cnpj = cnpj.group(2)
    status = 'Transmitido com sucesso!'
    cnpj_formatado = cnpj.replace('.', '').replace('/', '').replace('-', '').replace(' ', '')
    shutil.move(pasta + '\\' + arq_nome, pasta + '\\' + cnpj_formatado + '.pdf')
    _escreve_relatorio_csv(f'{cnpj_formatado};{status}', nome='Relatorio Envia DCTF')
def renomeia_arquivo_pendencia():
    arq_nome = 'Pendencia.pdf'
    pasta = 'V:\Setor Robô\Scripts Python\DCTF Mensal\Envia DCTF e Salvar Recibo\execução\Recibos'
    with fitz.open(pasta + '\\' + arq_nome) as pdf:
        # Percorre o PDF procurando pela variavel denominacao e extraindo os dados de creditos debitos e saldo
        for page in pdf:
            textinho = page.get_text('text', flags=1 + 2 + 8)
            cnpj = re.compile(r'(CNPJ: (.+))').search(textinho)

            if cnpj:
                cnpj = cnpj.group(2)
    status = 'Arquivo com pendencia!'
    cnpj_formatado = cnpj.replace('.', '').replace('/', '').replace('-', '').replace(' ', '')
    shutil.move(pasta + '\\' + arq_nome, pasta + '\\' + cnpj_formatado + '.pdf')
    _escreve_relatorio_csv(f'{cnpj_formatado};{status}', nome='Relatorio Envia DCTF')

def excluir():
    while not _find_img('excluir.png', conf=0.9):
        p.hotkey('ctrl', 'e')
        time.sleep(1)
    time.sleep(1)

    if _find_img('imagem1.png', conf=0.9):
        p.press('right')
    else:
        selecionar_primeiro()

    p.hotkey('alt', 'o')

    while not _find_img('confirma_exclusao.png', conf=0.9):
        time.sleep(1)
    time.sleep(1)

    p.hotkey('alt', 's')
    time.sleep(2)
    p.press('esc')

def move_arquivo_RFB():
    origem = 'C:\Arquivos de Programas RFB\DCTF Mensal 3.6\Declaracoes Gravadas\\'
    destino = 'V:\Setor Robô\Scripts Python\DCTF Mensal\Envia DCTF e Salvar Recibo\execução\Pastas RFB\\'

    allfiles = os.listdir(origem)

    for pasta in allfiles:
        shutil.move(origem + pasta, destino + pasta)

def move_arquivo_RFB_erro(status):
    origem = 'C:\Arquivos de Programas RFB\DCTF Mensal 3.6\Declaracoes Gravadas\\'
    destino = 'V:\Setor Robô\Scripts Python\DCTF Mensal\Envia DCTF e Salvar Recibo\execução\Pastas RFB\\'

    allfiles = os.listdir(origem)

    for pasta in allfiles:
        shutil.move(origem + pasta, destino + pasta)

    pasta = pasta.split("-")
    _escreve_relatorio_csv(f'{pasta[0]};{status}', nome='Relatorio Envia DCTF')

if __name__ == '__main__':
    run()