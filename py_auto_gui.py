import pandas as pd
import pyautogui as gui
import pyperclip as clip
import time

gui.PAUSE = 0.5

def abrir_aba_com_link(link):
    gui.hotkey('ctrl', 't')  
    raw_string = link
    clip.copy(raw_string)
    gui.hotkey('ctrl', 'v')
    gui.press('enter')

abrir_aba_com_link('https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing')

time.sleep(5)
gui.click(x=603, y=409, clicks=2)     #ajustar cursor pointer se necessario

time.sleep(2)  # delay
gui.click(x=588, y=561) #arquivo      #ajustar cursor pointer se necessario
gui.click(x=1611, y=235) #menu        #ajustar cursor pointer se necessario
gui.click(x=1260, y=874) #download    #ajustar cursor pointer se necessario

time.sleep(3) #esperar o download do arquivo

tabela = pd.read_excel(r'~\Downloads\Vendas - Dez.xlsx')
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()
display(tabela)

abrir_aba_com_link('https://mail.google.com/')

time.sleep(3) #esperar o download do arquivo
gui.click(x=138, y=262) #compose email  #ajustar cursor pointer se necessario

time.sleep(1) #esperar o compose do gmail
gui.write('pyMail+diretoria@gmail.com')
gui.press('tab') #escolher o email

gui.press('tab') # assunto do email
assunto = "Relatório de Vendas"
clip.copy(assunto)
gui.hotkey('ctrl', 'v')

gui.press('tab') # corpo do email
mailBody = f''' 
Prezados, bom dia

O faturamento foi de R$ {faturamento:,.2f}
A quantidade de produtos foi de {quantidade:,}

Atensiosamente,
SENDER'''
gui.write(mailBody)

gui.hotkey('ctrl', 'enter') #btn SEND