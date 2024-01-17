# -*- coding: utf-8 -*-
import shutil

import fitz, re, os
from sys import path
path.append(r'..\..\_comum')
from comum_comum import ask_for_dir
from comum_comum import _escreve_relatorio_csv, _escreve_header_csv
pastaOrigem = 'V:\Setor Robô\Scripts Python\DCTF Mensal\Analisa Declaração DCTF Mensal\ignore'
pastaComMovimento = 'V:\Setor Robô\Scripts Python\DCTF Mensal\Analisa Declaração DCTF Mensal\execução\DCTF Com Movimento'
pastaSemMovimento = 'V:\Setor Robô\Scripts Python\DCTF Mensal\Analisa Declaração DCTF Mensal\execução\DCTF Sem Movimento'

def run():
    documentos = ask_for_dir()

    # Analiza cada pdf que estiver na pasta
    for arq_nome in os.listdir(documentos):
        if not arq_nome.endswith('.pdf'):
            continue

        # Abrir o pdf
        arq = os.path.join(documentos, arq_nome)

        with fitz.open(arq) as pdf:

            # Percorre o PDF procurando pela variavel denominacao e extraindo os dados de creditos debitos e saldo
            for page in pdf:
                textinho = page.get_text('text', flags=1 + 2 + 8)
                achou = 'não'
                denominacao = re.compile(r'(DENOMINAÇÃO: (.+))').search(textinho)
                cnpj = re.compile(r'(CNPJ: (.+) )').search(textinho)
                if cnpj:
                    cnpj = cnpj.group(2)

                if denominacao:
                    achou = 'sim'

                    debito_apurado = re.compile(r'(DÉBITO APURADO\n(.+))').search(textinho)
                    soma_creditos_vinculados = re.compile(r'(SOMA DOS CRÉDITOS VINCULADOS\n(.+))').search(textinho)
                    saldo_pagar_debito = re.compile(r'(SALDO A PAGAR DO DÉBITO\n(.+))').search(textinho)

                    denominacao = denominacao.group(2)
                    debito_apurado = debito_apurado.group(2)
                    soma_creditos_vinculados = soma_creditos_vinculados.group(2)
                    saldo_pagar_debito = saldo_pagar_debito.group(2)

                    _escreve_relatorio_csv(f'{cnpj};{denominacao};{debito_apurado};{soma_creditos_vinculados};{saldo_pagar_debito}', nome='DCTF Com Movimento')
        cnpj_formatado = cnpj.replace('.', '').replace('/', '').replace('-', '').replace(' ', '')
        if achou == 'não':
            status = 'Sem Movimento'
            _escreve_relatorio_csv(f'{cnpj};{status}', nome='DCTF Sem Movimento')
            shutil.move(pastaOrigem + '\\' + arq_nome, pastaSemMovimento + '\\' + cnpj_formatado + '.pdf')
        else:
            shutil.move(pastaOrigem + '\\' + arq_nome, pastaComMovimento + '\\' + cnpj_formatado + '.pdf')

if __name__ == '__main__':
    run()
    _escreve_header_csv('CNPJ;DENOMINAÇÃO;DÉBITO APURADO;SOMA DOS CRÉDITOS VINCULADOS;SALDO A PAGAR DO DÉBITO', nome='DCTF Com Movimento')