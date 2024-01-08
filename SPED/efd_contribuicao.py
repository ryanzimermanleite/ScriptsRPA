# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------#
# Nome:     SPED CONTRIBUIÇÃO                                           #
# Objetivo: Importar, Assinar, Transmitir, Salvar Recibo                #
# Autor:    Ryan Zimerman Leite                                         #
# Data:     22/11/2023                                                  #
# ----------------------------------------------------------------------#
import time, datetime, os, re, pyautogui as p
import pyperclip, shutil
from sys import path
from datetime import datetime
from pathlib import Path
path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _barra_de_status

diretorioArquivosTxt = 'V:\!Setor SPED\EFD Contribuições\EFD Arquivos'
dir_txt = Path('V:\!Setor SPED\EFD Contribuições\EFD Arquivos')
e_dir = Path('V:\!Setor SPED\EFD Contribuições\EFD Relatorios')

@_barra_de_status
def run(window):
    p.click(833, 384)
    if os.path.exists(diretorioArquivosTxt):
        if len(os.listdir(diretorioArquivosTxt)) == 0:
            window['-Mensagens-'].update(' -> Não Tem Arquivos na Pasta!')
        else:
            startImportacaoAssinaturaTransmissao(window)
    else:
        window['-Mensagens-'].update("-> Diretorio não existe!")
     
def startImportacaoAssinaturaTransmissao(window):
    for arquivo in os.listdir(diretorioArquivosTxt):
        window['-Mensagens-'].update(str(arquivo))
        try:
            abreJanelaImportacao(arquivo)
            verificaMsgImportacao()

            abreJanelaAssinatura()
            verificaMsgAssinatura()

            abreJanelaTransmitir()
            verificaMsgTransmitir(arquivo)
        
            geraRelatorioExcel(arquivo)
            moveArquivoTransmitidos(arquivo)
        except:
            window['-Mensagens-'].update("Arquivo não importado/assinado/transmitido -> " + str(arquivo))
    excluirEscrituracao()

def abreJanelaImportacao(arquivo):
    while not _find_img('importar_escrituracao.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
        p.hotkey('ctrl', 'i')
    time.sleep(0.5)
    pyperclip.copy(arquivo)
    p.hotkey("ctrl", "v")  # Digita o nome do arquivo no input
    localizaPastaArquivosTxt()

def localizaPastaArquivosTxt():
    while not _find_img('pasta_EFD_arquivo_selecionada.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
        _click_img('este_computador.png', conf=0.9)
        time.sleep(0.5)
        _click_img('unidade_de_rede_DCA.png', conf=0.9, clicks=2)
        time.sleep(0.5)
        _click_img('Setor_SPED.png', conf=0.9, clicks=2)
        time.sleep(0.5)
        _click_img('EFD_Contribuicoes.png', pasta='imgs_contri', conf=0.9, clicks=2)
        time.sleep(0.5)
        _click_img('EFD_Pasta_Arquivos.png', conf=0.9, clicks=2)
    p.press('enter')

def verificaMsgImportacao():
    while not _find_img('importacao_concluida.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
    p.press('enter')
    while not _find_img('escrituracao_pronta_assinatura_contri.png', pasta='imgs_contri', conf=0.9):
        time.sleep(0.5)
        if _find_img('atualizacao_tabela.png', pasta='imgs', conf=0.9):
            p.press('enter')
        if _find_img('arquivo_contem_avisos_contri.png', pasta='imgs_contri', conf=0.9):
            p.press('enter')
        if _find_img('arquivo_validado_sucesso_contri.png', pasta='imgs_contri', conf=0.9):
            p.press('enter')
    p.press('enter')
    while not _find_img('resultado_validacao.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
        if _find_img('pendencias_validacao.png', pasta='imgs', conf=0.9):
            break
    p.hotkey('ctrl', 'f')

def abreJanelaAssinatura():
    while not _find_img('assinar_escrituracao.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
        p.hotkey('ctrl', 's')
    while not _find_img('lista_certificado.png', pasta='imgs', conf=0.9):
        p.press('up')
        time.sleep(0.5)
        p.press('enter')
    p.press('up', presses=2)
    time.sleep(0.5)
    p.press('tab', presses=4, interval=0.2)
    time.sleep(0.5)
    p.press('enter')

def verificaMsgAssinatura():
    while not _find_img('mensagem2.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
    p.press('enter')

def abreJanelaTransmitir():
    while not _find_img('EFD_Transmitir.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
        p.hotkey('ctrl', 't')
    p.press('up', presses=2)
    time.sleep(0.5)
    p.press('enter')
    time.sleep(0.5)
 
def verificaMsgTransmitir(arquivo):
    while not _find_img('aviso3.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
    p.press('enter')
    while not _find_img('recibo_transmissao.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
    p.press('enter')
    while not _find_img('salvar.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
    x = arquivo.split(".")
    pyperclip.copy(x[0])
    p.hotkey("ctrl", "v")
    localizaPastaRecibos()

def localizaPastaRecibos():
    while not _find_img('EFD_Pasta_Recibo_Selecionada.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
        _click_img('este_computador.png', conf=0.9)
        time.sleep(0.5)
        _click_img('unidade_de_rede_DCA.png', conf=0.9, clicks=2)
        time.sleep(0.5)
        _click_img('Setor_SPED.png', conf=0.9, clicks=2)
        time.sleep(0.5)
        _click_img('EFD_Contribuicoes.png', pasta='imgs_contri', conf=0.9, clicks=2)
        time.sleep(0.5)
        _click_img('EFD_Pasta_Recibo.png', conf=0.9, clicks=2)
    p.press('enter')
    time.sleep(0.5)
    p.press('esc', presses=5)

def moveArquivoTransmitidos(arquivo):
    x = arquivo.split(".")
    # GUARDA O NOME DOS ARQUIVOS USADOS NA VARIAVAEL .txt E O .REC
    pasta_arquivo_REC = 'V:\!Setor SPED\EFD Contribuições\EFD Arquivos' + '\\' + x[0] + '.REC'
    pasta_arquivo_TXT = 'V:\!Setor SPED\EFD Contribuições\EFD Arquivos' + '\\' + x[0] + '.txt'

    with open(pasta_arquivo_TXT, 'r', errors="ignore") as f:
        dados = f.read()
    dados_arquivo = re.findall(r'\|([^|]*)', dados)

    dataCompetencia = dados_arquivo[6]
    mesCompetencia = dataCompetencia[2:4]
    anoCompetencia = dataCompetencia[4:]
    pastaAnoMes_REC = 'V:\!Setor SPED\EFD Contribuições\EFD Transmitidos\\' + anoCompetencia + '\\' + mesCompetencia + '\\' + x[0] + '.REC'
    pastaAnoMes_TXT = 'V:\!Setor SPED\EFD Contribuições\EFD Transmitidos\\' + anoCompetencia + '\\' + mesCompetencia + '\\' + x[0] + '.txt'

    # MOVE O ARQUIVO TXT E O REC QUE GEROU PARA PASTA DE TRANSMITIDOS
    os.rename(pasta_arquivo_REC, pastaAnoMes_REC)
    os.rename(pasta_arquivo_TXT, pastaAnoMes_TXT)

def geraRelatorioExcel(arquivo):
    caminhoNomeArquivo = diretorioArquivosTxt + '\\' + arquivo
    dia = datetime.now().day
    mes = datetime.now().month
    ano = datetime.now().year
    dataTransmissao = f'{dia}-{mes}-{ano}'

    with open(caminhoNomeArquivo, 'r', errors="ignore") as f:
        dados = f.read()
    dados_arquivo = re.findall(r'\|([^|]*)', dados)
    tipoArquivo = dados_arquivo[2]
    dataCompetencia = dados_arquivo[6]
    razaoSocial = dados_arquivo[7]
    cnpjCliente = dados_arquivo[8]

    if tipoArquivo == '0':
        tipoArquivo = 'Original'
    elif tipoArquivo == '1':
        tipoArquivo = 'Retificado'

    nomeRelatorioExcel = 'Relatorio Contribuições ' + dataCompetencia[2:4] + '-' + dataCompetencia[4:]
    escreve_relatorio_csv(';'.join([cnpjCliente, razaoSocial, tipoArquivo, dataTransmissao, 'Transmitida']), nome=nomeRelatorioExcel)

def escreve_relatorio_csv(texto, nome='resumo', local=e_dir, end='\n', encode='latin-1'):
    os.makedirs(local, exist_ok=True)
    try:
        f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
    except:
        f = open(os.path.join(local, f"{nome} - auxiliar.csv"), 'a', encoding=encode)
    f.write(texto + end)
    f.close()
    
def excluirEscrituracao():
    while not _find_img('excluir_escrituracao.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
        p.hotkey('ctrl', 'e')
    p.press('up')
    time.sleep(0.5)
    p.hotkey('ctrl', 'a')
    time.sleep(0.5)
    p.press('enter')
    while not _find_img('deseja_excluir.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
    p.press('enter')
    while not _find_img('resultado_da_exclusao.png', pasta='imgs', conf=0.9):
        time.sleep(0.5)
    p.press('esc', presses=5)

if __name__ == '__main__':
    run()
