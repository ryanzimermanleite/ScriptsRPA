# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------#
# Nome:     Gera apuração                                               #
# Arquivo:  envia_apuracao.py                                           #
# Versão:   1.0.0                                                       #
# Modulo:   Dominio                                                     #
# Objetivo: Atualizar periodo de apuração das empresas                  #
# Autor:    Ryan Zimerman Leite                                         #
# Data:     30/10/2023                                                  #
# ----------------------------------------------------------------------#
import pyperclip, time, os, subprocess, pyautogui as p
from datetime import datetime
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _salvar_pdf
def verifica_empresa(cod):
    erro = 'sim'
    while erro == 'sim':
        try:
            p.click(1258, 82)

            while True:
                try:
                    time.sleep(1)
                    p.hotkey('ctrl', 'c')
                    time.sleep(1)
                    p.hotkey('ctrl', 'c')
                    time.sleep(1)
                    cnpj_codigo = pyperclip.paste()
                    break
                except:
                    pass

            time.sleep(0.5)
            codigo = cnpj_codigo.split('-')
            codigo = str(codigo[1])
            codigo = codigo.replace(' ', '')

            if codigo != cod:
                print(f'Código da empresa: {codigo}')
                print(f'Código encontrado no Domínio: {cod}')
                return False
            else:
                return True
        except:
            erro = 'sim'

def login(empresa, andamentos):
    cod, cnpj = empresa
    # espera a tela inicial do domínio
    while not _find_img('inicial.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)

    p.click(833, 384)

    # espera abrir a janela de seleção de empresa
    while not _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        p.press('f8')

    time.sleep(1)
    # clica para pesquisar empresa por código
    if _find_img('codigo.png', pasta='imgs_c', conf=0.9):
        p.click(p.locateCenterOnScreen(r'imgs_c\codigo.png', confidence=0.9))
    p.write(cod)
    time.sleep(3)

    # confirmar empresa
    p.hotkey('alt', 'a')
    # enquanto a janela estiver aberta verifica exceções
    while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
        time.sleep(1)
        if _find_img('sem_parametro.png', pasta='imgs_c', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Parametro não cadastrado para esta empresa']), nome=andamentos)
            print('❌ Parametro não cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            while not _find_img('parametros.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            p.press('esc')
            time.sleep(1)
            return False

        if _find_img('nao_existe_parametro.png', pasta='imgs_c', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Não existe parametro cadastrado para esta empresa']), nome=andamentos)
            print('❌ Não existe parametro cadastrado para esta empresa')
            p.press('enter')
            time.sleep(1)
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.press('esc')
            time.sleep(1)
            p.hotkey('alt', 'n')
            while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            return False

        if _find_img('empresa_nao_usa_sistema.png', pasta='imgs_c', conf=0.9) or _find_img('empresa_nao_usa_sistema_2.png', pasta='imgs_c', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Empresa não está marcada para usar este sistema']), nome=andamentos)
            print('❌ Empresa não está marcada para usar este sistema')
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5)
            while _find_img('trocar_empresa.png', pasta='imgs_c', conf=0.9):
                time.sleep(1)
            return False

        if _find_img('fase_dois_do_cadastro.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)
            p.hotkey('alt', 'n')

        if _find_img('conforme_modulo.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            time.sleep(1)

        if _find_img('aviso_regime.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'n')
            time.sleep(1)

        if _find_img('aviso.png', pasta='imgs_c', conf=0.9):
            p.hotkey('alt', 'o')
            time.sleep(1)

        if _find_img('erro_troca_empresa.png', pasta='imgs_c', conf=0.9):
            p.press('enter')
            time.sleep(1)
            p.press('esc', presses=5, interval=1)
            login(empresa, andamentos)

    if not verifica_empresa(cod):
        _escreve_relatorio_csv(';'.join([cod, cnpj, 'Empresa não encontrada']), nome=andamentos)
        print('❌ Empresa não encontrada')
        p.press('esc')
        return False

    p.press('esc', presses=5)
    time.sleep(1)

    return True
def apurar(mes_ano, empresa, andamentos):
    cod, cnpj = empresa

    # Se o mes escolhido for 12 ele define como 1 para a imagem
    mes = mes_ano.split('/')
    if mes[0] == '12':
        mes = '01'
    # Se não for mes 12 ele soma +1 ao mes escolhe para a imagem
    else:
        mes = str(int(mes[0]) + 1)

    # Se o mês da empresa estiver 1 mês na frente do mês escolhido ele pula a apuração
    while not _find_img('mes' + mes + '_.png', conf=0.9):
    
        if _find_img('empresas_2020.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Apuração 2020']), nome=andamentos)
            return 'ok'
        if _find_img('empresas_2021.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Apuração 2021']), nome=andamentos)
            return 'ok'
        if _find_img('empresas_2022.png', conf=0.9):
            _escreve_relatorio_csv(';'.join([cod, cnpj, 'Apuração 2022']), nome=andamentos)
            return 'ok'
        # espera o botão de movimentos do domínio aparecer na tela
        _wait_img('movimentos.png', conf=0.9, timeout=-1)

        print('>>> Apurando Empresa')
        # Enquanto não achar a janela de apuraçãp. Tenta abrir a janela de apuração
        while not _find_img('apuracao.png', conf=0.9):
            # Relatórios
            p.hotkey('alt', 'm')
            time.sleep(0.5)
            # gerador de relatórios
            p.press('r')
            time.sleep(0.5)

        # Escreve o mes_ano que foi escolhi e aperta o comando para gerar novo periodo
        time.sleep(0.5)
        p.write(mes_ano)
        time.sleep(0.5)
        p.press('tab')
        time.sleep(0.5)
        p.write(mes_ano)
        # gera novo periodo
        time.sleep(0.5)
        p.hotkey('alt', 'g')

        # Equanto não achar imagem do processo de apuração espera 1 segundo
        while not _find_img('progresso_apuracao.png', conf=0.9):
            time.sleep(1)

        # Enquanto não achar a imagem de fim de apuração ele procura pela imagem de avisos apuração
        while not _find_img('fim_apuracao.png', conf=0.9):
            time.sleep(1)
            # Se encontrar avisos ele fecha
            if _find_img('avisos_apuracao.png', conf=0.9):
                p.press('f')
            elif _find_img('atencao.png', conf=0.9):
                p.click(833, 384)
                time.sleep(1)
                p.press('y')

        # Fecha a janela de apuração
        time.sleep(1)
        p.press('n')
        time.sleep(0.5)
        p.press('esc')
        p.press('n')

        print('✔ Apuração Concluida')
        _escreve_relatorio_csv(';'.join([cod, cnpj, 'Apuração com Sucesso']), nome=andamentos)
        # fechar qualquer possível tela aberta
        p.press('esc', presses=5)
        time.sleep(1)
        return 'ok'
    _escreve_relatorio_csv(';'.join([cod, cnpj, 'Empresa já apurada!']), nome=andamentos)
    return 'ok'


@_barra_de_status
def run(window):
    # abre o Domínio Web e o módulo, no caso será o módulo Folha
    _login_web()
    _abrir_modulo('escrita_fiscal')
    
    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        _indice(count, total_empresas, empresa, index, window)

        while True:
            # abre a empresa no domínio
            if not login(empresa, andamentos):
                break
            # Chama a função de apurar
            resultado = apurar(str(ano), empresa, andamentos)

            if resultado == 'dominio fechou':
                _login_web()
                _abrir_modulo('escrita_fiscal')

            if resultado == 'modulo fechou':
                _abrir_modulo('escrita_fiscal')

            if resultado == 'ok':
                break
    _escreve_header_csv('CÓDIGO;CNPJ', nome=andamentos)


if __name__ == '__main__':
    # Captura o ano base que vai ser usado para apuração e envio de reinf
    ano = p.prompt(text='Qual periodo base?', title='Script incrível', default='00/0000')
    
    # Abre uma janela para escolher o arquivo excel que vai ser usado
    empresas = _open_lista_dados()
    
    # Define uma variavel para o nome do excel que vai ser gerado apos o final da execução
    andamentos = 'Resultado 1'
    
    # Abre uma janela para escolhe se quer continuar a ultima execução ou não
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
        