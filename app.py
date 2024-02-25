import openpyxl
from time import sleep 
from urllib.parse import quote
import webbrowser
import pyautogui
import os

#deixando cliente logar no whatsapp web
webbrowser.open('https://web.whatsapp.com')
sleep(20)

#abrindo planilha
workbook = openpyxl.load_workbook('planilha.xlsx')#abrindo planilha

#abrindo pagina planilha
pagina_clientes = workbook['page1']

#encontrando os dados na planilha
for linha in pagina_clientes.iter_rows(min_row=2):
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    
#criando mensagem
    mensagem = (f'Olá {nome}, você possui um débito com vencimento para {vencimento.strftime('%d/%m/%Y')}, efetue o pagamento através do link abaixo: www.pagar.com.br')

    try:
#link da mensagem
        link_whatsapp = (f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}')
    
#acessando site
        webbrowser.open(link_whatsapp)
        sleep(12)

#clicando para enviar a mensagem
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0], seta[1])
        sleep(5)
#fechar guia
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possivel enviar mensagem para {nome}')
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome}, {telefone}, {os.linesep}')


