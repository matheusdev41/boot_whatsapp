"""
PRECISO AUTOMATIZAR MINHAS MENSAGENS P/ MEUS CLIENTES, QUERO PODER 
MANDAR MENSAGENS DE COBRANÇA EM DETERMINADO DIA COM CLIENTES COM
VENCIMENTO DIFERENTE

"""

# 1 passo é descrever os passos manuais e depois transformar isso em codigo

import openpyxl
from urllib.parse import quote #permite formatar links para envio em links especiais 

import webbrowser #para conseguirmos abrir o nosso navegador
from time import sleep

import pyautogui 


webbrowser.open('https://web.whatsapp.com')
sleep(30)
# ler planilha e guardar informações sobre nome, telefone e data de vencimento

workbook = openpyxl.load_workbook('dados.xlsx')
pagina_clientes = workbook['Plan1']

for linha in pagina_clientes.iter_rows(min_row=2):
    #nome, telefone, vencimento
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value

    mensagem = f'Olá {nome} seu boleto vence no dia {vencimento.strftime('%d/%m/%Y')}. Favor pagar no link https://www.link_do_pagamento.com'

    
# Criar links personalizados do whatsapp e enviar mensagens para cada cliente com base nos dados da planilha


   
    try:
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f'Não foi possível enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')