# AUTOMATIZA MENSAGENS PELO WHATSAPP PARA CLIENTES DETERMINANDO DIA E VALOR DA COBRANÇA

import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

# Ler planilha e guardar informações sobre nome, telefone, data de vencimento e valor.
workbook = openpyxl.load_workbook('clientes.xlsx')
pagina_clientes = workbook['Sheet1']

for linha in pagina_clientes.iter_rows(min_row=2):
    # nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    valor = round(float(linha[3].value), 2)

    # Variavel com a mensagem e data formatada.
    mensagem = f"Olá {nome} seu pagamento no valor de {valor}, vence no dia {vencimento.strftime('%d/%m/%y')}. Favor pagar no pix, 00000000 ou no link de pagamento https://www.link_do_pagamento.com"

    # links personalizados do whatsapp e enviar mensagens para cada cliente
    # com base nos dados da planilha
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(15)
        # Localiza e direciona o mouse para seta de enviar mensagem
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(1)
        # Clica com o botão esquerdo do mouse
        pyautogui.click(seta[0], seta[1])
        sleep(2)
        # Fecha a guia atual
        pyautogui.hotkey('ctrl', 'w')
        sleep(2)

    # Salva em uma tabela clientes que não foi possivel enviar a mensagem
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}{os.linesep}')
