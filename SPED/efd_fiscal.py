# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------#
# Nome:     SPED FISCAL                                                 #
# Arquivo:  efd_fiscal.py                                               #
# Versão:   1.0.0                                                       #
# Modulo:   EFD Fiscal                                                  #
# Objetivo: Importar, Assinar, Transmitir, Salvar Recibo                #
# Autor:    Ryan Zimerman Leite                                         #
# Data:     23/11/2023                                                  #
# Obs:      Mesma coisa que o EFD CONTRIBUIÇÕES muda so algumas coisas  #
# ----------------------------------------------------------------------#
from datetime import datetime
import fitz, re, time, os, pyautogui as p, pyperclip as clip
import pyperclip
from sys import path
from pathlib import Path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _time_execution, _open_lista_dados, \
    _where_to_start, _indice, _barra_de_status

pasta_sped = 'V:\!Setor SPED\EFD Fiscal\EFD Arquivos'
e_dir = Path('V:\!Setor SPED\EFD Fiscal\EFD Relatorios')
e_dir2 = 'V:\!Setor SPED\EFD Fiscal\EFD Relatorios\\Resultado Fiscal - '


# Recebe um texto 'texto' junta com 'end' e escreve num arquivo 'nome'
def escreve_relatorio_csv(texto, nome='resumo', local=e_dir, end='\n', encode='latin-1'):
    os.makedirs(local, exist_ok=True)

    try:
        f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.csv"), 'a', encoding=encode)

    f.write(texto + end)
    f.close()

# Recebe um cabeçalho 'texto' e escreve
# no comeco do arquivo 'nome'
def escreve_header_csv(texto, nome='resumo', local=e_dir, encode='latin-1'):
    os.makedirs(local, exist_ok=True)

    with open(os.path.join(local, f"{nome}.csv"), 'r', encoding=encode) as f:
        conteudo = f.read()

    with open(os.path.join(local, f"{nome}.csv"), 'w', encoding=encode) as f:
        f.write(texto + '\n' + conteudo)

def importar(window, andamentos, ano):
    p.click(833, 384)
    time.sleep(0.5)
    if os.path.exists(pasta_sped):
        if len(os.listdir(pasta_sped)) == 0:
            return '1'
        else:
            for arquivo in os.listdir(pasta_sped):
                print('>>> Importando...')
                window['-Mensagens-'].update(str(arquivo) + ' -> Importando...')
                try:
                    while not _find_img('importar_escrituracao.png', pasta='imgs', conf=0.9):
                        time.sleep(0.5)
                        p.hotkey('ctrl', 'i')
                    time.sleep(0.5)

                    pyperclip.copy(arquivo)
                    p.hotkey("ctrl", "v")

                    while not _find_img('EFD_arq.png', pasta='imgs', conf=0.9):
                        time.sleep(0.5)
                        _click_img('computador.png', conf=0.9)
                        time.sleep(0.5)
                        _click_img('DCA.png', conf=0.9, clicks=2)
                        time.sleep(0.5)
                        _click_img('Setor_SPED.png', conf=0.9, clicks=2)
                        time.sleep(0.5)
                        _click_img('EFD_Fiscal2.png', conf=0.9, clicks=2)
                        time.sleep(0.5)
                        _click_img('EFD_Pasta.png', conf=0.9, clicks=2)

                    p.press('enter')

                    while not _find_img('aviso1.png', pasta='imgs', conf=0.9):
                        time.sleep(0.5)
                    time.sleep(0.5)
                    p.press('enter')

                    while not _find_img('resultado_importar.png', pasta='imgs', conf=0.9):
                        time.sleep(0.5)
                        if  _find_img('mensagem1.png', pasta='imgs', conf=0.9):
                            time.sleep(0.5)
                            p.press('enter')
                        elif _find_img('aviso123.png', pasta='imgs', conf=0.9):
                            time.sleep(0.5)
                            p.press('enter')
                        elif _find_img('aviso6.png', pasta='imgs', conf=0.9):
                            time.sleep(0.5)
                            p.press('enter')
                        elif _find_img('pendencia.png', pasta='imgs', conf=0.9):
                            break
                    time.sleep(0.5)
                    p.hotkey('ctrl', 'f')

                    print(arquivo + ' -> Importado com sucesso!')
                    window['-Mensagens-'].update(str(arquivo) + ' -> Importado com sucesso!')

                    status = assinar(arquivo, window)
                    nome_relat = arquivo.split("l")
                    codigo = nome_relat[1].split(".")

                    if status == '4':
                        status = transmitir(arquivo, window)
                        if status == '5':
                            print(arquivo + ' -> Transmitido com sucesso!\n')
                            window['-Mensagens-'].update(str(arquivo) + ' -> Transmitido com sucesso!')

                            resultado = 'Transmitido com Sucesso!'


                            x = arquivo.split(".")

                            pasta_arquivo_REC = 'V:\!Setor SPED\EFD Fiscal\EFD Arquivos' + '\\' + x[0] + '.REC'
                            pasta_arquivo_TXT = 'V:\!Setor SPED\EFD Fiscal\EFD Arquivos' + '\\' + x[0] + '.txt'

                            pasta_transmitidos_REC = 'V:\!Setor SPED\EFD Fiscal\EFD Transmitidos' + '\\' + x[0] + '.REC'
                            pasta_transmitidos_TXT = 'V:\!Setor SPED\EFD Fiscal\EFD Transmitidos' + '\\' + x[0] + '.txt'

                            os.rename(pasta_arquivo_REC, pasta_transmitidos_REC)
                            os.rename(pasta_arquivo_TXT, pasta_transmitidos_TXT)
                            resultado = 'ok'
                        else:
                            print(">>> Falha ao transmitir!\n")
                            window['-Mensagens-'].update(str(arquivo) + ' -> Falha ao transmitir!')


                            resultado = 'Falha ao Transmitir!'
                            escreve_relatorio_csv(f'{codigo[0]};{resultado}', nome=andamentos)
                    else:
                        print('>>> Falha ao assinar!\n')
                        window['-Mensagens-'].update(str(arquivo) + ' -> Falha ao assinar!')

                        resultado = 'Falha ao assinar!'
                        escreve_relatorio_csv(f'{codigo[0]};{resultado}', nome=andamentos)
                except:
                    print("Arquivo não importado -> " + arquivo)
                    window['-Mensagens-'].update("Arquivo não importado -> " + str(arquivo))

                    resultado = 'Arquivo não importado!'
                    escreve_relatorio_csv(f'{codigo[0]};{resultado}', nome=andamentos)
            if resultado == 'ok':
                window['-Mensagens-'].update("Envios concluidos!")
                print('Execução Finalizada!')
                salva_relatorio(window, ano)
    else:
        print(">>> Diretorio não existe!")
        window['-Mensagens-'].update("-> Diretorio não existe!")

def assinar(arquivo, window):
    print('>>> Assinando...')
    window['-Mensagens-'].update(str(arquivo) + ' -> Assinando...')

    while not _find_img('assinar_escrituracao.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
        p.hotkey('ctrl', 's')

    time.sleep(1)

    while not _find_img('lista_certificado.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
        p.press('up')
        time.sleep(0.5)
        p.press('enter')

    time.sleep(0.5)
    p.press('up')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('tab')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)

    while not _find_img('mensagem3.png', pasta='imgs', conf=0.9):
        time.sleep(1)
    p.press('esc')
    time.sleep(0.5)
    p.press('esc')
    print(arquivo + ' -> Assinado com sucesso!')
    window['-Mensagens-'].update(str(arquivo) + ' -> Assinado com sucesso!')

    return '4'

def transmitir(arquivo, window):
    print('>>> Transmitindo...')
    window['-Mensagens-'].update(str(arquivo) + ' -> Transmitindo...')

    while not _find_img('EFD_Transmitir.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
        p.hotkey('ctrl', 't')
    time.sleep(1)

    p.press('up')
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)

    while not _find_img('aviso3.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)

    while not _find_img('recibo.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)

    while not _find_img('salvar.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
    x = arquivo.split(".")
    pyperclip.copy(x[0])
    p.hotkey("ctrl", "v")

    while not _find_img('efd_recibo2.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
        _click_img('computador.png', conf=0.9)
        time.sleep(0.5)
        _click_img('DCA.png', conf=0.9, clicks=2)
        time.sleep(0.5)
        _click_img('SETOR_EFD.png', conf=0.9, clicks=2)
        time.sleep(0.5)
        _click_img('EFD_Fiscal2.png', conf=0.9, clicks=2)
        time.sleep(0.5)
        _click_img('EFD_Pasta_Recibo.png', conf=0.9, clicks=2)

    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
    p.press('esc')
    return '5'

def salva_relatorio(window, ano):
    print('>>> Excluindo...')
    window['-Mensagens-'].update(' -> Excluindo...')

    local = e_dir2 + str(ano) + '.csv'

    time.sleep(3)

    # AGUARDA A TELA DE EXCLUSAO
    while not _find_img('excluir.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
        p.hotkey('ctrl', 'e')
    time.sleep(1)

    p.press('up')
    time.sleep(0.5)
    p.hotkey('ctrl', 'a')
    time.sleep(0.5)

    if os.path.exists(local):
        print('Ja Existe o Excel')
    else:
        with open(local, 'w') as csvfile:
            print('Arquivo Criado!')
            print(csvfile)

    while True:
        try:
            time.sleep(1)
            p.hotkey('ctrl', 'c')
            time.sleep(1)
            p.hotkey('ctrl', 'c')
            time.sleep(1)
            arquivos_excluir = pyperclip.paste()
            break
        except:
            pass

    while not _find_img('executar.png', conf=0.9):
        try:
            p.hotkey('winleft', 'r')
        except:
            pass
    time.sleep(1)
    p.press('del')
    time.sleep(1)
    pyperclip.copy(local)
    time.sleep(1)
    p.hotkey('ctrl', 'v')
    time.sleep(1)
    p.press('enter')

    while not _find_img('excel.png', conf=0.9):
        time.sleep(1)
    time.sleep(1)
    _wait_img('excel_clique.png', conf=0.9)
    time.sleep(1)
    _click_img('excel_clique.png', conf=0.9, clicks=2, timeout=1)
    time.sleep(1)
    p.press('up')
    time.sleep(1)
    p.hotkey('ctrl', 'down')
    time.sleep(1)
    p.hotkey('ctrl', 'down')
    time.sleep(1)
    p.hotkey('ctrl', 'up')
    time.sleep(1)
    p.press('down')
    time.sleep(1)
    pyperclip.copy(arquivos_excluir)
    time.sleep(1)
    p.hotkey('ctrl', 'v')
    time.sleep(1)
    p.hotkey('alt', 'f4')
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.write('Resultado Fiscal - ' + ano)
    time.sleep(1)
    p.press('tab')
    time.sleep(1)
    _click_img('tipo.png', conf=0.9)
    time.sleep(1)
    _click_img('csv.png', conf=0.9)
    time.sleep(1)
    _click_img('salvar2.png', conf=0.9)
    time.sleep(1)
    p.press('tab')
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    p.press('enter')
    while not _find_img('excluir.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
    time.sleep(1)
    p.press('enter')
    _wait_img('excluir2.png', conf=0.9)
    time.sleep(1)
    p.press('enter')
    time.sleep(1)
    _wait_img('resultado_ex.png.png', conf=0.9)
    time.sleep(1)
    p.press('esc')

@_barra_de_status
def run(window):
    ano = p.prompt(text='Qual mês-ano de competencia?', title='Script incrível', default='00-0000')
    time.sleep(2)
    andamentos = f'Resultado Fiscal {ano}'



    importar(window, andamentos, ano)

if __name__ == '__main__':
    run()