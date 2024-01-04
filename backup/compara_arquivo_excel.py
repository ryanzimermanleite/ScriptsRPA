import PySimpleGUI as sg
from pyautogui import alert, confirm
from datetime import datetime
from pathlib import Path
import pandas as pd
import os

e_dir = Path('T:\# Projeto Robô\Maria Eduarda')

if __name__ == '__main__':
    def excel_ae():
        return sg.Frame("", [[
            sg.FileBrowse('Pesquisar', button_color='grey80', key='-Abrir-', file_types=(('Planilhas Excel', '*.xlsx *.xls'),)),
            sg.InputText(key='-input_excel_ae-', size=200, disabled=True)
        ]], pad=(5, 3), border_width=0)

    def excel_sieg():
        return sg.Frame("", [[
            sg.FileBrowse('Pesquisar',button_color='grey80', key='-Abrir-', file_types=(('Planilhas Excel', '*.xlsx *.xls'),)),
            sg.InputText(key='-input_excel_sieg-', size=80, disabled=True)
        ]], pad=(5, 3), expand_x=True, expand_y=True, border_width=0)

    def acoes():
        return sg.Frame("", [[
            sg.Button('INICIAR',key='-INICIAR-', expand_x=True, expand_y=True, button_color='lightgreen'),
            sg.Button('AJUDA', expand_x=True, expand_y=True, button_color='light yellow'),
            sg.Button('SAIR',key='-SAIR-', expand_x=True, expand_y=True, button_color='Light Coral')
        ]], pad=(5, 3), expand_x=True, expand_y=True, border_width=0)


    def resultado():
        return sg.Frame("", [[
            sg.Button('RELATORIO', expand_x=True, expand_y=True, button_color='lightblue')
        ]], pad=(5, 3), expand_x=True, expand_y=True, border_width=0)

    sg.theme('GrayGrayGray')
    layout_frame1 = [
        [sg.Frame("A&C", [[excel_ae()]], pad=(5, 3), expand_x=True, expand_y=True)],
        [sg.Frame("SIEG", [[excel_sieg()]], pad=(5, 3), expand_x=True, expand_y=True)],
    ]

    layout_frame3 = [
        [sg.Frame("Ações", [[acoes()]], pad=(5, 3), expand_x=True, expand_y=True, title_location=sg.TITLE_LOCATION_TOP),
        sg.Frame("Resultado", [[resultado()]], pad=(5, 3), expand_x=True, expand_y=True, title_location=sg.TITLE_LOCATION_TOP)],
    ]

    layout = [
         [sg.Frame("Arquivos Excel", layout_frame1, size=(500, 150), title_location=sg.TITLE_LOCATION_TOP),],
         [sg.Frame("Veiga & Postal", layout_frame3, size=(500, 100), title_location=sg.TITLE_LOCATION_TOP)]]

    window = sg.Window("Comparação Certificado A&C no SIEG", layout, margins=(2, 2), finalize=True)

    def run(window, input_excel_ae, input_excel_sieg):
        dados_ae, dados_sieg = open_lista_dados(input_excel_ae, input_excel_sieg)
        dia = datetime.now().day
        mes = datetime.now().month
        ano = datetime.now().year

        dia_com_dois_digitos = f'{dia:02d}'
        mes_com_dois_digitos = f'{mes:02d}'
        data_atual = f'{dia_com_dois_digitos}/{mes_com_dois_digitos}/{ano}'
        data_atual_int = datetime.strptime(data_atual, "%d/%m/%Y")

        empresa_ae_valida = []
        empresa_sieg_valida = []
        for linha_ae in dados_ae.itertuples():
            index_ae = linha_ae[0]
            nome_ae = linha_ae[1]
            cnpj_ae = linha_ae[2]
            data_ae = linha_ae[3]
            data_ae_int = datetime.strptime(data_ae, "%d/%m/%Y")

            if data_ae_int >= data_atual_int:
                empresa_ae_valida.append(nome_ae, cnpj_ae)
            else:
                empresa_ae_invalida = [nome_ae, cnpj_ae]

        for linha_sieg in dados_sieg.itertuples():
            index_sieg = linha_sieg[0]
            nome_sieg = linha_sieg[1]
            cnpj_sieg = linha_sieg[2]
            data_sieg = linha_sieg[3]
            data_sieg_int = datetime.strptime(data_sieg, "%d/%m/%Y")

            if data_sieg_int >= data_atual_int:
                empresa_sieg_valida.append(cnpj_sieg)
            else:
                empresa_sieg_invalida = [nome_sieg, cnpj_sieg]


            print(empresa_ae_valida)








        #print(data_atual)

        # ARMAZENA NOME e CNPJ -> DADOS_AE [CD1.Vencim >= DATETIME.NOW()]
        # ARMAZENA NOME e CNPJ -> DADOS_SIEG [VENCIMENTO >= DATETIME.NOW()]

        # CRUZA OS DADOS CNPJ (AE, SIEG) OQUE FOR IGUAL = (IGUAL) DESCARTA
        # OQUE SOBRAR DOS DADOS CNPJ AE -> ARMAZENA*

        '''if dados_ae[CD1.Vencim] >= data_atual:
            cnpj_ae_ativo = dados_ae[CD1.Vencim]
        else:
            cnpj_ae_inativo = dados_ae[CD1.Vencim]


        if dados_sieg[Vencimento] >= data_atual:
            cnpj_sieg_ativo = dados_sieg[Vencimento]
        else:
            cnpj_sieg_inativo = dados_sieg[Vencimento]


        dados_new = cnpj_ae_ativo -  cnpj_sieg_ativo
        # escreve_relatorio_csv(';'.join(nome,  cnpj,  'Cadastrar no SIEG']), nome='tst')'''


        alert(text=f'Comparação concluida')

        window['-INICIAR-'].update(disabled=False)

    def open_lista_dados(input_excel_ae, input_excel_sieg):

        if input_excel_ae:
            lista_colunas_ae = ['Razao', 'CNPJ', 'CD1.Vencim']
            arquivo_ae = pd.read_excel(input_excel_ae, header=0, usecols=lista_colunas_ae)

        if input_excel_sieg:
            lista_colunas_sieg = ['Nome', 'CPF_CNPJ', 'Vencimento']
            arquivo_sieg = pd.read_excel(input_excel_sieg, header=0, usecols = lista_colunas_sieg)

        return arquivo_ae, arquivo_sieg

    def escreve_relatorio_csv(texto, nome='resumo', local=e_dir, end='\n', encode='latin-1'):
        os.makedirs(local, exist_ok=True)

        try:
            f = open(os.path.join(local, f"{nome}.csv"), 'a', encoding=encode)
        except:
            f = open(os.path.join(local, f"{nome} - auxiliar.csv"), 'a', encoding=encode)

        f.write(texto + end)
        f.close()

    while True:
        event, values = window.read()
        try:
            input_excel_ae = values['-input_excel_ae-']
            input_excel_sieg = values['-input_excel_sieg-']

        except:
            input_excel_ae = 'Desktop'
            input_excel_sieg = 'Desktop'

        if event == sg.WIN_CLOSED or event == '-SAIR-':
            break

        if event == '-INICIAR-':
            if not input_excel_ae or not input_excel_sieg:
                alert(text=f'Por favor selecione uma planilha do A&C e SIEG.')
            else:
                window['-INICIAR-'].update(disabled=True)
                run(window, input_excel_ae, input_excel_sieg)
    window.close()