import time
# -*- coding: utf-8 -*-
import os, pyperclip
import time, shutil, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _escreve_relatorio_csv, e_dir, _open_lista_dados, _where_to_start, _barra_de_status


def abre_tela_inicial():
    while not _find_img('tela_inicial.png', conf=0.9):
        time.sleep(1)
    _click_img('tela_inicial.png', conf=0.9)

def abre_gravar_declaracao():
    while not _find_img('gravar_declaracao.png', conf=0.9):
        p.hotkey('ctrl', 'g')
        time.sleep(1)

def selecionar_primeiro():
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('up')
    time.sleep(0.5)

def gravar_declaracao():
    p.hotkey('alt', 'o')

    while not _find_img('unidade_c.png', conf=0.9):
        p.press('up')
        time.sleep(0.5)

    p.hotkey('alt', 'g')

    while not _find_img('sucesso_gravar.png', conf=0.9):
        time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.press('esc')
    print('Gravado com sucesso!')


def transmitir():
    while not _find_img('transmitir.png', conf=0.9):
        p.hotkey('ctrl', 'i')
        time.sleep(1)

    while not _find_img('unidade_c_2.png', conf=0.9):
        p.press('tab', presses=3, interval=0.3)
        time.sleep(0.5)
        p.press('up', presses=7, interval=0.2)
        p.press('tab')
        time.sleep(1)

    p.hotkey('alt', 'o')

    while not _find_img('transmitir_via_internet.png', conf=0.9):
        time.sleep(1)

    p.press('tab', presses=4, interval=0.3)
    time.sleep(0.5)
    p.press('space')
    time.sleep(1)
    p.press('down')
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('up')
    time.sleep(0.5)

    p.hotkey('alt', 'o')

    assinar()

    while not _find_img('transmitido_sucesso.png', conf=0.9):
        time.sleep(1)

    p.press('enter')
    time.sleep(0.5)
    p.press('esc', presses=2, interval=0.3)

def assinar():
    while not _find_img('assinar.png', conf=0.9):
        time.sleep(1)

    p.press('tab', presses=3, interval=0.3)
    time.sleep(0.5)
    p.hotkey('alt', 'a')

def salvar_recibo():
    while not _find_img('imprimir_recibo.png', conf=0.9):
        p.hotkey('ctrl', 'b')
        time.sleep(1)

    while not _find_img('unidade_c_3.png', conf=0.9):
        p.press('tab', presses=3, interval=0.3)
        time.sleep(0.5)
        p.press('up', presses=7, interval=0.2)
        p.press('tab')
        time.sleep(1)

    p.hotkey('alt', 'o')

    while not _find_img('imprimir_recibo2.png', conf=0.9):
        time.sleep(1)

def salvar_pdf(contador):
    p.hotkey('alt', 'o')
    while not _find_img('salvar.png', conf=0.9):
        time.sleep(1)
    os.makedirs(r'{}\{}'.format(e_dir, 'Recibos'), exist_ok=True)
    p.write(str(contador))

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
    time.sleep(2)
    print('✔ Recibo Salvo')

    p.press('esc', presses=2, interval=0.2)

def run():
    abre_tela_inicial()

    abre_gravar_declaracao()

    selecionar_primeiro()

    gravar_declaracao()

    transmitir()

    salvar_recibo()

    selecionar_primeiro()

    contador = 1

    salvar_pdf(contador)

    while True:
        abre_tela_inicial()

        abre_gravar_declaracao()

        selecionar_primeiro()

        p.press('down', presses=contador, interval=0.2)

        time.sleep(1)

        gravar_declaracao()

        transmitir()

        salvar_recibo()

        selecionar_primeiro()

        p.press('down', presses=contador, interval=0.2)

        time.sleep(1)

        contador += 1

        if _find_img('barra_no_fim2.png', conf=0.9):
            if _find_img('fim.png', conf=0.9):
                salvar_pdf(contador)
                print('Fim da execução!')
                break

        salvar_pdf(contador)

if __name__ == '__main__':
    run()