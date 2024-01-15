import PySimpleGUI as sg
from pyautogui import alert, confirm
from datetime import datetime
from pathlib import Path
import pandas as pd
import os
e_dir = Path('T:\# Projeto Robô\Carolina')

if __name__ == '__main__':
    def exibe_opcoes():
        names = ['Banco Itaú - 21', 'Banco do Brasil - 19', 'Banco Daycoval - 5042']
        return sg.Frame("", [[
            sg.Radio("ENTRADA", "entrada_saida", key="-entrada-", font=("Helvetica", 10), default=True),
            sg.Radio("SAÍDA", "entrada_saida", key="-saida-", font=("Helvetica", 10)),
            sg.Combo(names, font=("Helvetica", 10), expand_x=True, enable_events=True, readonly=False, key='-COMBO-')
        ]], pad=(5, 3), border_width=0)
    def seleciona_excel_cliente():
        return sg.Frame("", [[
            sg.FileBrowse('Pesquisar', button_color='grey80', key='-Abrir-', file_types=(('Planilhas Excel', '*.xlsx *.xls'),)),
            sg.InputText(key='-input_excel_cliente-', size=200, disabled=True)
        ]], pad=(5, 3), border_width=0)

    def botoes_acoes():
        return sg.Frame("", [[
            sg.Button('INICIAR', key='-INICIAR-', expand_x=True, expand_y=True, button_color='lightgreen'),
            sg.Button('AJUDA', key='-AJUDA-', expand_x=True, expand_y=True, button_color='light yellow'),
            sg.Button('SAIR', key='-SAIR-', expand_x=True, expand_y=True, button_color='Light Coral')
        ]], pad=(5, 3), expand_x=True, expand_y=True, border_width=0)

    def botao_resultado():
        return sg.Frame("", [[
            sg.Button('RELATORIO', key='-RELATORIO-', expand_x=True, expand_y=True, button_color='lightblue')
        ]], pad=(5, 3), expand_x=True, expand_y=True, border_width=0)

    sg.theme('GrayGrayGray')
    layout_arquivos_excel = [
        [sg.Frame("Cliente", [[seleciona_excel_cliente()]], pad=(5, 3), expand_x=True, expand_y=True)],
        [sg.Frame("Opções", [[exibe_opcoes()]], pad=(5, 3), expand_x=True, expand_y=True)],
    ]

    layout_botoes = [
        [sg.Frame("Ações", [[botoes_acoes()]], pad=(5, 3), expand_x=True, expand_y=True, title_location=sg.TITLE_LOCATION_TOP),
        sg.Frame("Resultado", [[botao_resultado()]], pad=(5, 3), expand_x=True, expand_y=True, title_location=sg.TITLE_LOCATION_TOP)],
    ]

    layout = [
         [sg.Frame("Arquivos Excel", layout_arquivos_excel, size=(500, 160), title_location=sg.TITLE_LOCATION_TOP),],
         [sg.Frame("Veiga & Postal", layout_botoes, size=(500, 100), title_location=sg.TITLE_LOCATION_TOP)]]

    window = sg.Window("Gera Planilha Entrada/Saída para Domínio", layout, margins=(2, 2), finalize=True)
    window['-COMBO-'].update('Banco Itaú - 21')
    dia = datetime.now().day
    mes = datetime.now().month
    ano = datetime.now().year
    hora = datetime.now().hour
    minuto = datetime.now().minute

    def run(window, input_excel_cliente, banco):

        dados_cliente, status, tipo = open_lista_dados(input_excel_cliente, banco)
        numero_banco_2 = banco.split(' -')

        arquivo = pd.read_excel(input_excel_cliente, index_col=1)
        coluna1 = (arquivo.iloc[0])
        coluna1_str = (str(coluna1))
        razao = coluna1_str.split()

        nomeRelatorio = f'{razao[0]} {tipo}-{numero_banco_2[0]} {mes}-{ano} ás ({hora}.{minuto})'
        if status == 'ok':
            for linha_cliente in dados_cliente.itertuples():
                index = linha_cliente[0]
                data_baixa = linha_cliente[1]
                data_vencimento = linha_cliente[2]
                cnpj = linha_cliente[3]
                numero_nota = linha_cliente[4]
                if tipo == 'Entrada':
                    numero_nota = int(numero_nota)
                else:
                    numero_nota = str(numero_nota)
                valor_baixa = linha_cliente[5]
                valor_total = linha_cliente[6]
                juros = linha_cliente[7]
                descontos = linha_cliente[8]

                multa = 'nada'
                numero_banco = banco.split('- ')
                cnpj_formatado = cnpj.replace('.', '').replace('/', '').replace('-', '')

                escreve_relatorio_csv(f'{numero_nota};{cnpj_formatado};{data_vencimento};{data_baixa};{str(valor_total).replace(".",",")};{str(juros).replace(".",",")};{str(multa).replace(".",",")};{str(descontos).replace(".",",")};{numero_banco[1]}', nome=nomeRelatorio)
            escreve_header_csv('Número da Nota;'
                               'CPF/CNPJ do Cliente;'
                               'Data de Vencimento da Parcela;'
                               'Data da Baixa da Parcela;'
                               'Valor Recebido;'
                               'Valor Juros;'
                               'Valor Multa;'
                               'Valor Desconto;'
                               'Código da Conta Banco/Caixa', nome=nomeRelatorio)

            alert(text=f'Relatorio concluido')

            window['-INICIAR-'].update(disabled=False)
        else:
            return

    def open_lista_dados(input_excel_cliente, banco):
        if banco == 'Banco Itaú - 21':
            banco = 'BANCO ITAÚ S/A'

        if banco == 'Banco do Brasil - 19':
            banco = 'BANCO DO BRASIL S/A'

        if banco == 'Banco Daycoval - 5042':
            banco = 'Banco Daycoval'

        if entrada is True:

            try:

                #ENTRADa
                arquivo_cliente_entrada = pd.read_excel(input_excel_cliente, skiprows=range(0, 4))
                arquivo_cliente_filtrado_entrada = arquivo_cliente_entrada['Nome Conta Banco Baixa'] == banco
                dados_cliente_entrada = arquivo_cliente_entrada[arquivo_cliente_filtrado_entrada].drop(
                    columns=['Nome Conta Banco Baixa', 'Emissão', 'Sacado', 'Complementos Histórico',
                             'Compl. Livre Histórico', 'Número da parcela'])

                return dados_cliente_entrada, 'ok', 'Entrada'
            except:
                alert(text=f'Você selecionou ENTRADA!\n\n Por favor selecione uma planilha de entrada.')
                window['-INICIAR-'].update(disabled=False)
                return arquivo_cliente_entrada, 'erro', 'Entrada'

        else:
            try:
                # SAIDA

                arquivo_cliente_saida = pd.read_excel(input_excel_cliente, skiprows=range(0, 4))
                arquivo_cliente_filtrado_saida = arquivo_cliente_saida['Nome Conta Banco Baixa'] == banco
                dados_cliente_saida = arquivo_cliente_saida[arquivo_cliente_filtrado_saida].drop(
                    columns=['Nome Conta Banco Baixa', 'Emissão', 'Vcto.Título', 'Cedente', 'Compl. Livre Histórico'])

                return dados_cliente_saida, 'ok', 'Saída'
            except:
                alert(text=f'Você selecionou SAÍDA!\n\n Por favor selecione uma planilha de saída.')
                window['-INICIAR-'].update(disabled=False)
                return arquivo_cliente_saida, 'erro', 'Saída'

    def escreve_relatorio_csv(texto, nome='resumo', local=e_dir, end='\n', encode='latin-1'):
        os.makedirs(local, exist_ok=True)

        try:
            f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
        except:
            f = open(os.path.join(local, f"{nome} - auxiliar.csv"), 'a', encoding=encode)

        f.write(texto + end)
        f.close()

    def escreve_header_csv(texto, nome='resumo', local=e_dir, encode='latin-1'):
        os.makedirs(local, exist_ok=True)

        with open(os.path.join(local, f"{nome}.csv"), 'r', encoding=encode) as f:
            conteudo = f.read()

        with open(os.path.join(local, f"{nome}.csv"), 'w', encoding=encode) as f:
            f.write(texto + '\n' + conteudo)

    while True:
        event, values = window.read()
        try:
            input_excel_cliente = values['-input_excel_cliente-']
            entrada = values['-entrada-']
            saida = values['-saida-']
            banco = values['-COMBO-']
        except:
            input_excel_cliente = 'Desktop'

        if event == sg.WIN_CLOSED or event == '-SAIR-':
            break

        if event == '-INICIAR-':
            if not input_excel_cliente:
                alert(text=f'Por favor selecione uma planilha do Cliente.')
            else:
                window['-INICIAR-'].update(disabled=True)
                run(window, input_excel_cliente, banco)
        if event == '-RELATORIO-':
            try:
                os.startfile('T:\# Projeto Robô\Carolina')
            except:
                alert(
                    text=f'Diretório do Relatório não encontrado...')
        if event == '-AJUDA-':
            alert(text=f'Fluxograma e Manual de Ajuda em desenvolvimento...\n\n  -> Desenvolvido por Ryan Zimerman Leite')
    window.close()