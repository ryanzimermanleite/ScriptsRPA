import shutil, os, time, pyautogui as p
from sys import path

path.append(r'..\..\_comum')
from pyautogui_comum import _find_img, _click_img, _wait_img
from comum_comum import _indice, _time_execution, _open_lista_dados, _escreve_relatorio_csv, _where_to_start, \
    _barra_de_status, _escreve_header_csv
from dominio_comum import _login_web, _abrir_modulo, _login

def valor_folha_sn(empresa, andamentos):
    cod, cnpj, nome = empresa

    _wait_img('movimentos.png', conf=0.9, timeout=-1)

    print('>>> Importando e Gravando')

    # tenta abrir a tela do gerador de relatórios até abrir
    while not _find_img('valor_da_folha.png', conf=0.9):
        # Relatórios
        p.hotkey('alt', 'm')
        time.sleep(0.5)
        p.press('o')
        time.sleep(0.5)
        p.press('n')
        time.sleep(0.5)
        p.press('f')
        time.sleep(2)

    time.sleep(0.5)
    p.hotkey('alt', 'i')

    while not _find_img('botao_ativado.png', conf=0.9):
        time.sleep(1)

    p.hotkey('alt', 'g')

    while not _find_img('botao_desativado.png', conf=0.9):
        time.sleep(1)


    p.press('esc', presses=5)
    _escreve_relatorio_csv(';'.join([cod, cnpj, nome, 'Gravação com sucesso!']), nome=andamentos)
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
            if not _login(empresa, andamentos):
                break
            # Chama a função de apurar
            resultado = valor_folha_sn(empresa, andamentos)

            if resultado == 'ok':
                break
    # Escreve o cabeçalho do excel no final de todas as execuçoes
    _escreve_header_csv('COD;CNPJ;NOME;STATUS', nome=andamentos)


if __name__ == '__main__':
    # Abre uma janela para escolher o arquivo excel que vai ser usado
    empresas = _open_lista_dados()

    # Define uma variavel para o nome do excel que vai ser gerado apos o final da execução
    andamentos = 'Resultado Valor da Folha SN'

    # Abre uma janela para escolhe se quer continuar a ultima execução ou não
    index = _where_to_start(tuple(i[0] for i in empresas))
    if index is not None:
        run()
        