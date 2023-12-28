import pyperclip, shutil, os, fitz, time, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _encerra_dominio, _login

def verifica_parametro(empresa, andamentos):
    cod, cnpj, nome = empresa
    
    _wait_img('controle.png', conf=0.9, timeout=-1)
    
    print('>>> Analisando o parametro')
    
    # tenta abrir a tela do gerador de relatórios até abrir
    while not _find_img('parametros.png', conf=0.9):
        # Relatórios
        p.hotkey('alt', 'c')
        time.sleep(0.5)
        p.press('p')
        time.sleep(0.5)
    time.sleep(1)
    
    while not _find_img('geral_decimo.png', conf=0.9):
        time.sleep(0.5)
        _click_img('decimo_terceiro.png', conf=0.9)
    time.sleep(0.5)


    _click_img('botao_13.png', conf=0.9)

    time.sleep(2)
    
    if _find_img('img3.png', conf=0.9):
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Nenhum Parametro Marcado']), nome=andamentos)
        p.press('esc', presses=5)
        return 'ok'
    elif _find_img('img1.png', conf=0.9):
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Parametro Marcado 1']), nome=andamentos)
        p.press('esc', presses=5)
        return 'ok'
    elif _find_img('img2.png', conf=0.9):
        _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Parametro Marcado 2']), nome=andamentos)
        p.press('esc', presses=5)
        return 'ok'
    
@_barra_de_status
def run(window):
    # Abre uma janela para escolher o arquivo excel que vai ser usado
    empresas = _open_lista_dados()

    # Define uma variavel para o nome do excel que vai ser gerado apos o final da execução
    andamentos = 'Resultado Verifica Parametro 13 - 1'

    # Abre uma janela para escolhe se quer continuar a ultima execução ou não
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is None:
        return andamentos

    # abre o Domínio Web e o módulo, no caso será o módulo Folha
    _login_web()
    _abrir_modulo('folha')

    total_empresas = empresas[index:]
    for count, empresa in enumerate(empresas[index:], start=1):
        # printa o indice da empresa que está sendo executada
        window['-Mensagens-'].update(f'{str(count)} / {str(len(total_empresas))}')
        _indice(count, total_empresas, empresa, index)

        while True:
            # abre a empresa no domínio
            if not _login(empresa, andamentos):
                break
            # Chama a função de apurar
            resultado = verifica_parametro(empresa, andamentos)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)
    _encerra_dominio()

if __name__ == '__main__':
    run()