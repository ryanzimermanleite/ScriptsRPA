# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------#
# Nome:     Separa XML                                                  #
# Arquivo:  separa_xml.py                                               #
# Versão:   1.0.0                                                       #
# Modulo:   Analise de Arquivos                                         #
# Objetivo: Separar os arquivos XML que contenha CFOP ou 'CESTA BASICA' #
# Autor:    Ryan Zimerman Leite e Willian Rocha                         #
# Data:     14/09/2023                                                  #
# ----------------------------------------------------------------------#
from tkinter.filedialog import askopenfilename, askdirectory, Tk
from bs4 import BeautifulSoup as bs
from functools import wraps
import sys
import re, os

sys.path.append(r'..\..\_comum')
from comum_comum import _time_execution

pasta_cfop = "V:\Setor Robô\Scripts Python\Analise de Arquivos\Separa XML\ignore\cfop"
pasta_cesta = "V:\Setor Robô\Scripts Python\Analise de Arquivos\Separa XML\ignore\cesta"

# Função para o usuario selecionar a pasta
def ask_for_dir(title='Abrir pasta'):
    root = Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    folder = askdirectory(
        title=title, )
    return folder if folder else False
@_time_execution
def run():
    pasta_selecionada = ask_for_dir()
    # Permissão para sistema ler os arquivos da pasta
    os.chdir(str(pasta_selecionada))

    # Percorre a pasta abrindo, lendo e formatando os arquivos xml
    for arquivo in os.listdir():
        file = open(arquivo, encoding="utf-8")
        contents = file.read()
        soup = bs(contents, 'xml')

        # Procura atravez de regex as tags no XML,  <CFOP> de número 6910 ou 5910
        if re.compile(r'<CFOP>6910</CFOP>').search(str(soup)) or re.compile(r'<CFOP>5910</CFOP>').search(str(soup)):
            # Se achar fecha o arquivo
            file.close()
            # Move o arquivo para a pasta de destino
            os.rename(arquivo, pasta_cfop + "\\" + arquivo)
            print('Arquivo:', arquivo, '->', pasta_cfop)

        # Lista com as possiveis variaveis da palavra CESTA BASICA
        lista_cesta = [r'cesta basica',
                       r'cesta basicas',
                       r'cesta básica',
                       r'cesta básicas',
                       r'cestas basica',
                       r'cestas basicas',
                       r'cestas básica',
                       r'cestas básicas']

        # Percorre cada item da lista procurando pela palavra no XML
        for mensagem_regex in lista_cesta:
            # Procura nos XML as palavras que estão na lista_cesta
            if re.compile(mensagem_regex, re.IGNORECASE).search(str(soup)):
                # Fecha o arquivo
                file.close()
                # Move o arquivo para a pasta de destino
                os.rename(arquivo, pasta_cesta + "\\" + arquivo)
                print('Arquivo:', arquivo, '->', pasta_cesta)
                break

if __name__ == '__main__':
    run()