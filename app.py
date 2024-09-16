import openpyxl # type: ignore
from urllib.parse import quote 
import webbrowser
from time import sleep
import pyautogui

webbrowser.open('https://web.whatsapp.com/')
sleep(25)

workbook = openpyxl.load_workbook("lista_msg.xlsx")

pagina_msg = workbook["Planilha1"]

for linha in pagina_msg.iter_rows(min_row=2):
    
    nome = linha[0].value
    numero = linha[1].value
    
    mensagem = f'Olá {nome}. Favor me responder corretamente!'
    
    try:
        link_msg_zap = f'https://web.whatsapp.com/send?phone={numero}&text={quote(mensagem)}'
        # hhtps://web.wathsapp.com/send?phone=$text
        webbrowser.open(link_msg_zap)
        sleep(10)
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl','w')
        sleep(5)
    except:
        print(f"Não foi possivel enviar mensagem [nome]")
        with open('erros.csv','a',newline='',encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{numero}')
    
