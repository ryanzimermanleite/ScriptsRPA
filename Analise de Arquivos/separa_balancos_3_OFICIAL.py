# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------#
# Nome:     SEPARA BALANÇOS                                             #
# Arquivo:  separa_balancos_3_OFICIAL.py                                #
# Versão:   1.0.0                                                       #
# Modulo:   Analise de Arquivos                                         #
# Objetivo: Separa os arquivos de balanços enviados para a pasta        #
# Autor:    Ryan Zimerman Leite                                         #
# Data:     2/10/2023                                                  #
# ----------------------------------------------------------------------#
import os
import time

import PySimpleGUI as sg
from threading import Thread
from time import sleep
from email.message import EmailMessage
from smtplib import SMTP_SSL
from shutil import move
from pywhatkit import sendwhatmsg_instantly

# Executa as funçoes prinicipais de mover arquivos
def mover_arquivos():
    while True:

        if event == "Sair" or event == sg.WIN_CLOSED:
            return

        window['text2'].update('Movendo...')

        pasta01_to_pasta02(pasta_01)
        pasta02_to_pasta03(pasta_02)
        pasta04_to_pasta05(pasta_04)
        pasta05_to_pasta06(pasta_05)
        for cont in range(alterar_tempo, 0, -1):
            cont = cont % (24 * 3600)
            hour = cont // 3600
            cont %= 3600
            minutes = cont // 60
            cont %= 60
            formatado = "%d:%02d:%02d" % (hour, minutes, cont)
            window['text2'].update(formatado)
            sleep(1)
def led_scrip_thread():
    while True:
        if event == "Sair" or event == sg.WIN_CLOSED:
            return

        qtd_pasta1 = len(os.listdir(pasta_01))
        window['-CONTADOR1-'].update(qtd_pasta1)

        qtd_pasta2 = len(os.listdir(pasta_02))
        window['-CONTADOR2-'].update(qtd_pasta2 - 2)

        qtd_pasta3 = len(os.listdir(pasta_03))
        window['-CONTADOR3-'].update(qtd_pasta3)

        qtd_pasta4 = len(os.listdir(pasta_04))
        window['-CONTADOR4-'].update(qtd_pasta4 - 1)

        qtd_pasta5 = len(os.listdir(pasta_05))
        window['-CONTADOR5-'].update(qtd_pasta5 - 1)

        qtd_pasta6 = len(os.listdir(pasta_06))
        window['-CONTADOR6-'].update(qtd_pasta6)

        if len(os.listdir(pasta_01)) == 0:
            SetLED(window, '_cpu_1',
                   'red')
            sleep(1)
        else:
            SetLED(window, '_cpu_1',
                   'green1')
            sleep(1)
        if len(os.listdir(pasta_02)) > 2:
            SetLED(window, '_cpu_2',
                   'green1')
            sleep(1)
        else:
            SetLED(window, '_cpu_2',
                   'red')
            sleep(1)
        if len(os.listdir(pasta_03)) == 0:
            SetLED(window, '_cpu_3',
                   'red')
            sleep(1)
        else:
            SetLED(window, '_cpu_3',
                   'green1')
            sleep(1)
        if len(os.listdir(pasta_04)) > 1:
            SetLED(window, '_cpu_4',
                   'green1')
            sleep(1)
        else:
            SetLED(window, '_cpu_4',
                   'red')
            sleep(1)
        if len(os.listdir(pasta_05)) > 1:
            SetLED(window, '_cpu_5',
                   'green1')
            sleep(1)
        else:
            SetLED(window, '_cpu_5',
                   'red')
            sleep(1)
        if len(os.listdir(pasta_06)) == 0:
            SetLED(window, '_cpu_6',
                   'red')
            sleep(1)
        else:
            SetLED(window, '_cpu_6',
                   'green1')
            sleep(1)
def pasta01_to_pasta02(pasta_01):
    resultado = ''
    lista_zap = []
    if os.path.exists(pasta_01):
        if len(os.listdir(pasta_01)) == 0:
            xyz = 'nada'  # SO PRA TER ALGO NO IF##
        else:
            for arquivo in os.listdir(pasta_01):
                try:
                    move(pasta_01 + arquivo, pasta_02)
                    print("==================================")
                    print("Arquivo movido -> " + arquivo + ' 02 - Balanços para Veiga analisar')
                    lista_zap.append(arquivo)
                    resultado = 'OK'
                except:
                    print("==========ARQUIVO NÃO MOVIDO=============")
                    print(arquivo + ' -> já existe na pasta -> 02 - Balanços para Veiga analisar\n')
            if resultado == 'OK':
                print("==================================")
                print('---> WhatsApp Enviado para VEIGA!\n')
                envia_zap(zap_veiga, lista_zap)
                envia_email(email_mirella, 'Extratos Bancários', lista_zap)
                envia_email(email_ivonaide, 'Extratos Bancários', lista_zap)
    else:
        print("O diretório não existe! -> " + pasta_01)



def pasta02_to_pasta03(pasta_02):
    rejeitado = pasta_02 + 'rejeitado\\'
    autorizado = pasta_02 + 'autorizado\\'
    resultado = ''
    lista_rejeitado = []
    lista_autorizado = []

    ## REJEITADOS
    if os.path.exists(rejeitado):
        if len(os.listdir(rejeitado)) == 0:
            xyz = 'nada'  # SO PRA TER ALGO NO IF
        else:
            for arquivo in os.listdir(rejeitado):
                try:
                    move(rejeitado + arquivo, pasta_03)
                    print("==================================")
                    print("Arquivo movido -> " + arquivo + " 03 - Balanços rejeitados para ajuste")
                    lista_rejeitado.append(arquivo)
                    resultado = 'OK'
                except:
                    print("==========ARQUIVO NÃO MOVIDO=============")
                    print(arquivo + ' -> já existe na pasta -> 03 - Balanços rejeitados para ajuste\n')
            if resultado == 'OK':
                print("==================================")
                # envia_email(carolina, 'Balanços rejeitados para ajuste')
                envia_email(email_carolina, '03 - Balanços rejeitados para ajuste', lista_rejeitado)
    else:
        print("O diretório não existe! -> " + rejeitado)

    ## APROVADOS
    if os.path.exists(autorizado):
        if len(os.listdir(autorizado)) == 0:
            xyz = 'nada'  # SO PRA TER ALGO NO IF
        else:
            for arquivo in os.listdir(autorizado):
                try:
                    move(autorizado + arquivo, pasta_04)
                    print("==================================")
                    print("Arquivo movido -> " + arquivo + " 04 - Balanços aprovados para assinatura")
                    lista_autorizado.append(arquivo)
                    resultado = 'OK'
                except:
                    print("==========ARQUIVO NÃO MOVIDO=============")
                    print(arquivo + ' -> já existe na pasta -> 04 - Balanços aprovados para assinatura\n')
            if resultado == 'OK':
                print("==================================")
                envia_email(email_evandro, '04 - Balanços aprovados para assinatura', lista_autorizado)
    else:
        print("O diretório não existe! -> " + autorizado)

def pasta04_to_pasta05(pasta_04):
    assinado = pasta_04 + 'balanço assinado\\'
    resultado = ''
    lista_assinado = []
    if os.path.exists(assinado):
        if len(os.listdir(assinado)) == 0:
            xyz = 'nada'  # SO PRA TER ALGO NO IF
        else:
            for arquivo in os.listdir(assinado):
                try:
                    print("==================================")
                    print("Arquivo movido -> " + arquivo + " 05 - Balanços assinados para envio Onvio")
                    move(assinado + arquivo, pasta_05)
                    lista_assinado.append(arquivo)
                    resultado = 'OK'
                except:
                    print("==========ARQUIVO NÃO MOVIDO=============")
                    print(arquivo + ' -> já existe na pasta -> 05 - Balanços assinados para envio Onvio\n')
            if resultado == 'OK':
                print("==================================")
                envia_email(email_stephani, '05 - Balanços assinados para envio Onvio', lista_assinado)
    else:
        print("O diretório não existe! -> " + assinado)

def pasta05_to_pasta06(pasta_05):
    onvio = pasta_05 + 'enviado onvio\\'
    resultado = ''
    lista_onvio = []
    if os.path.exists(onvio):
        if len(os.listdir(onvio)) == 0:
            xyz = 'nada'  # SO PRA TER ALGO NO IF
        else:
            for arquivo in os.listdir(onvio):
                try:
                    move(onvio + arquivo, pasta_06)
                    print("==================================")
                    print("Arquivo movido -> " + arquivo + " 06 - Balanços enviados e finalizados")
                    lista_onvio.append(arquivo)
                    resultado = 'OK'
                except:
                    print("==========ARQUIVO NÃO MOVIDO=============")
                    print(arquivo + ' -> já existe na pasta -> 06 - Balanços enviados e finalizados\n')
            if resultado == 'OK':
                print("==================================")
                envia_email(email_carolina, '06 - Balanços enviados e finalizados', lista_onvio)
    else:
        print("O diretório não existe! -> " + onvio)

# Abre as pastas
def abrir_pasta(pasta):
    if pasta == 'pasta1':
        os.startfile(os.path.join(pasta_01))
    if pasta == 'pasta2':
        os.startfile(os.path.join(pasta_02))
    if pasta == 'pasta3':
        os.startfile(os.path.join(pasta_03))
    if pasta == 'pasta4':
        os.startfile(os.path.join(pasta_04))
    if pasta == 'pasta5':
        os.startfile(os.path.join(pasta_05))
    if pasta == 'pasta6':
        os.startfile(os.path.join(pasta_06))

# Função que faz o envio do email para os destinatios
def envia_email(destinatario, assunto, lista):
    email = "roboautorobot@gmail.com"

    # Abre o arquivo armazenado a senha do emial
    with open(r'senha_email.txt') as f:
        senha = f.readlines()
        f.close()
    # Guarda a senha do email em uma variavel
    senha_do_email = senha[0]

    # Define o assunto, destinatario
    msg = EmailMessage()
    msg['Subject'] = assunto
    msg['From'] = email
    msg['To'] = destinatario

    item_lista = ''

    for item in lista:
        item_lista += item + '\n'

    if assunto == 'Extratos Bancários':
        msg.set_content("A partir deste momento gostaria que fosse cobrado mensalmente os extratos ou conseguir os códigos de operadores das empresas abaixo.: \n\n" + item_lista + '\n')
    else:
        # Conteudo do corpo do email
        msg.set_content("BALANÇOS: \n\n" + item_lista + '\n')

    # Protocolo para o envio do email
    with SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email, senha_do_email)
        smtp.send_message(msg)
        print('---> E-mail enviado para: ' + destinatario + '\n')

# Função que faz o envio da mensagem para o zap do  veiga
def envia_zap(numero, lista_zap):
    item_lista_zap = ''

    for item in lista_zap:
        item_lista_zap += item + '\n'

    print('--->> Enviando Mensagem WhatsApp')
    if numero == 'numero':
        return 'erro'
    try:
        sendwhatmsg_instantly('+55' + numero, f"Olá Veiga!\n"
                                              f"Tem arquivos balanços para analisar!\n\n" + item_lista_zap,
                              7, True, 2)

        sleep(1)
    except:
        return 'erro'

# Define o cronometro de tempo
def hora():
    return sg.Frame("", [[

        sg.Text('0:01:00'
                '', key='text2', font=("Helvetica", 14, "bold"), justification='center',
                text_color=(sg.theme_text_color()))],

        [sg.Combo(names, font=('Helvetica', 13), expand_x=True, enable_events=True, readonly=False, key='-COMBO-')]

    ], pad=(5, 3), expand_x=True, expand_y=True, border_width=0, size=(85, 60))

# Define os texto da contagem de arquivos por pastas
def textos():
    return sg.Frame("", [[
        sg.Text("01 - Balanços para analise:", ), sg.Text('', key="-CONTADOR1-"), sg.Text(''),
        sg.Text("02 - Balanços para o Veiga:"), sg.Text('', key="-CONTADOR2-"), ],
        [sg.Text("03 - Balanços rejeitados:"), sg.Text('', key="-CONTADOR3-"), sg.Text('   '),
         sg.Text("04 - Balanços aprovados:"), sg.Text('', key="-CONTADOR4-"), ],
        [sg.Text("05 - Balanços assinados:"), sg.Text('', key="-CONTADOR5-"), sg.Text('  '),
         sg.Text("06 - Balanços finalizados:"), sg.Text('', key="-CONTADOR6-"),

         ]], pad=(5, 3), expand_x=True, expand_y=True, border_width=0)

# Define os botoes de iniciar, limpar  e sair
def botoes():
    return sg.Frame("", [[
        sg.Button("START", size=(10, 10), font=("Helvetica", 16, "bold"), button_color=('black', 'green1')),
        sg.Button('Limpar', size=(7, 10), font=("Helvetica", 16, "bold"), button_color=('black', 'white'), ),
        sg.Button('Sair', size=(6, 10), font=("Helvetica", 16, "bold"), button_color='red'),
    ]], pad=(5, 3), expand_x=True, expand_y=True, background_color='#404040', border_width=0)

# Define o campo de saida
def output():
    return sg.Frame("", [[
        sg.Output(size=(150, 40), font=20, key='-OUTPUT-', text_color='white', background_color='black')
    ]], pad=(5, 3), expand_x=True, expand_y=True, background_color='#404040', border_width=0)

# Define a imagem da logo
def imagem():
    return sg.Frame("", [[
        sg.Image('logo_veiga.png', expand_x=True, expand_y=True)
    ]], pad=(5, 3), expand_x=True, background_color='black', expand_y=True, border_width=0)

# Define os nomes e botões das pastas
def pastas(pasta):
    if pasta == 1:
        return sg.Frame("", [[
            sg.Text('01 -  Balanços para analise', text_color=(sg.theme_text_color())),
            LEDIndicator('_cpu_1')],
            [sg.Button("Abrir Pasta", key='-PASTA1-')]

        ], pad=(5, 3), expand_x=True, expand_y=True, border_width=0, size=(85, 60))

    if pasta == 2:
        return sg.Frame("", [[
            sg.Text('02 - Balanços para o Veiga', text_color=(sg.theme_text_color())),
            LEDIndicator('_cpu_2')],
            [sg.Button("Abrir Pasta", key='-PASTA2-')]

        ], pad=(5, 3), expand_x=True, expand_y=True, border_width=0, size=(85, 60))

    if pasta == 3:
        return sg.Frame("", [[
            sg.Text('03 - Balanços rejeitados', text_color=(sg.theme_text_color())),
            LEDIndicator('_cpu_3')],
            [sg.Button("Abrir Pasta", key='-PASTA3-')]
        ], pad=(5, 3), expand_x=True, expand_y=True, border_width=0, size=(85, 60))

    if pasta == 4:
        return sg.Frame("", [[
            sg.Text('04 - Balanços aprovados', text_color=(sg.theme_text_color())),
            LEDIndicator('_cpu_4')],
            [sg.Button("Abrir Pasta", key='-PASTA4-')]
        ], pad=(5, 3), expand_x=True, expand_y=True, border_width=0, size=(85, 60))

    if pasta == 5:
        return sg.Frame("", [[
            sg.Text('05 - Balanços assinados', text_color=(sg.theme_text_color())),
            LEDIndicator('_cpu_5')],
            [sg.Button("Abrir Pasta", key='-PASTA5-')]
        ], pad=(5, 3), expand_x=True, expand_y=True, border_width=0, size=(85, 60))

    if pasta == 6:
        return sg.Frame("", [[
            sg.Text('06 - Balanços finalizados', text_color=(sg.theme_text_color())),
            LEDIndicator('_cpu_6')],
            [sg.Button("Abrir Pasta", key='-PASTA6-')]
        ], pad=(5, 3), expand_x=True, expand_y=True, border_width=0, size=(85, 60))

def LEDIndicator(key=None, radius=30):
    return sg.Graph(canvas_size=(radius, radius),
                    graph_bottom_left=(-radius, -radius),
                    graph_top_right=(radius, radius),
                    pad=(0, 0), key=key)

def SetLED(window, key, color):
    graph = window[key]
    graph.erase()
    graph.draw_circle((0, 0), 12, fill_color=color, line_color=color)

if __name__ == '__main__':
    # Cria uma lista com tempos pre-determinados
    names = ['1 Minuto', '5 Minutos', '10 Minutos', '30 Minutos', '1 Hora']

    # Inicia a variavel tempo com 60 segundos
    alterar_tempo = 60

    # Define as pastas de balanços
    pasta_01 = "\\\\vpsrv04\Balanços\\01 - Colocar balanços aqui para analise\\"
    pasta_02 = "\\\\vpsrv04\Balanços\\02 - Balanços para o Veiga analisar\\"
    pasta_03 = "\\\\vpsrv04\Balanços\\03 - Balanços rejeitados para ajuste\\"
    pasta_04 = "\\\\vpsrv04\Balanços\\04 - Balanços aprovados para assinatura\\"
    pasta_05 = "\\\\vpsrv04\Balanços\\05 - Balanços assinados para envio Onvio\\"
    pasta_06 = "\\\\vpsrv04\Balanços\\06 - Balanços enviados e finalizados\\"

    # Abre o arquivo txt e separa em variaveis o resultado da linha lida
    with open(r'destinatarios.txt') as f:
        destinatario = f.readlines()
        f.close()

    zap_veiga = destinatario[0]
    email_carolina = destinatario[1]
    email_evandro = destinatario[2]
    email_stephani = destinatario[3]
    email_mirella = destinatario[4]
    email_ivonaide = destinatario[5]

    # Define o tema visual
    sg.theme('DarkGrey4')

    # Faz a criação dos frames
    layout_frame1 = [
        [pastas(1), pastas(2), pastas(3), pastas(4), pastas(5), pastas(6)],
        [sg.Frame("Output", [[output()]], pad=(5, 3), expand_x=True, expand_y=True, size=(200, 500), )],
        [sg.Frame("Ações", [[botoes()]], pad=(5, 3), expand_y=True, size=(360, 100),
                  title_location=sg.TITLE_LOCATION_TOP),
         sg.Frame("VEIGA & POSTAL", [[imagem()]], pad=(5, 3), expand_y=True, size=(280, 100),
                  title_location=sg.TITLE_LOCATION_TOP),

         sg.Frame("Quantidade Balanços por Pasta", [[textos()]], pad=(5, 3), expand_y=True, size=(430, 800),
                  title_location=sg.TITLE_LOCATION_TOP),
         sg.Frame("Próxima Execução", [[hora()]], size=(300, 400), element_justification='center',
                  title_location=sg.TITLE_LOCATION_TOP)],
    ]

    layout = [
        [sg.Frame("Balanços", layout_frame1, size=(1250, 700))],
    ]

    window = sg.Window("Separa Balanços 3", layout, margins=(2, 2), finalize=True, element_justification='center')

    # Atualiza o combo box para o texto de um minuto
    window['-COMBO-'].update('1 Minuto')
    while True:

        event, values = window.read()

        # Ao iniciar o robo ele inicia a função de mover e led em threads
        if event == 'START':
            script_thread = Thread(target=led_scrip_thread)
            script_thread.start()

            window["START"].update(disabled=True)

            script_thread = Thread(target=mover_arquivos)
            script_thread.start()
        elif event == "Limpar":
            window['-OUTPUT-'].update('')

        elif event == "Sair" or event == sg.WIN_CLOSED:
            break

        # Se mudar o tempo do  combobox ele atualiza a varivel assim alterando o tempo
        elif event == "-COMBO-":
            if values["-COMBO-"] == '1 Minuto':
                alterar_tempo = 60
            elif values["-COMBO-"] == '5 Minutos':
                alterar_tempo = 300
            elif values["-COMBO-"] == '10 Minutos':
                alterar_tempo = 600
            elif values["-COMBO-"] == '30 Minutos':
                alterar_tempo = 1800
            elif values["-COMBO-"] == '1 Hora':
                alterar_tempo = 3600

        # Botões para abrir a pasta dos balanços
        elif event == '-PASTA1-':
            abrir_pasta('pasta1')
        elif event == '-PASTA2-':
            abrir_pasta('pasta2')
        elif event == '-PASTA3-':
            abrir_pasta('pasta3')
        elif event == '-PASTA4-':
            abrir_pasta('pasta4')
        elif event == '-PASTA5-':
            abrir_pasta('pasta5')
        elif event == '-PASTA6-':
            abrir_pasta('pasta6')

    window.close()