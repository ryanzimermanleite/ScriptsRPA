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

def abre_imprimir_pdf():
    while not _find_img('imprimir_declaracao.png', conf=0.9):
        p.hotkey('ctrl', 'l')
        time.sleep(1)

def selecionar_primeiro():
    time.sleep(0.5)
    p.press('down')
    time.sleep(0.5)
    p.press('up')
    time.sleep(0.5)

def run():

    abre_tela_inicial()

    abre_imprimir_pdf()

    selecionar_primeiro()

    contador = 1

    salvar_pdf(contador)

    while True:
        abre_tela_inicial()

        abre_imprimir_pdf()

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


def salvar_pdf(contador):
    p.hotkey('alt', 'o')
    while not _find_img('salvar.png', conf=0.9):
        time.sleep(1)
    os.makedirs(r'{}\{}'.format(e_dir, 'Arquivos'), exist_ok=True)
    p.write(str(contador))

    p.press('tab', presses=6)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    pyperclip.copy('V:\Setor Robô\Scripts Python\DCTF Mensal\Extrair Declaração DCTF Mensal\{}\{}'.format(e_dir,'Arquivos'))
    p.hotkey('ctrl', 'v')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.hotkey('alt', 'l')
    time.sleep(1)
    print('✔ PDF gerado')

if __name__ == '__main__':
    run()