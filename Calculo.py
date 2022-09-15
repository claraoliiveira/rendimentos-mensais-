import pandas as pd

tabela = pd.read_excel('report_mozi.xlsx')

display(tabela)

tabela[['Valor', 'Saldo', 'Saldo Sacável']].applymap(lambda x: x.replace('R$',''))

tab1 = tabela.loc[tabela['Descrição das Movimentações']=='Rendimentos',['Descrição das Movimentações','Valor']]  

display(tab1)

novo = tab1[['Valor']].applymap(lambda x: x.replace('R$','').replace(',','.'))

print(novo)

novo['Valor']= novo['Valor'].astype(float, errors = 'raise')

valor = novo['Valor'].sum()

print(valor)

#depois do cálculo enviar automaticamente com pyautogui

!pip install pyautogui

import pyautogui

import pyperclip

import time

pyautogui.hotkey('ctrl', 't')

pyperclip.copy('https://mail.google.com/mail/u/1/#inbox')

pyautogui.hotkey('ctrl','v')

pyautogui.press('enter')

time.sleep(5)

pyautogui.click(x=887, y=251)

pyautogui.click(x=1351, y=497)

time.sleep(0.5)

pyautogui.write('email pra quem voce quer enviar')

pyautogui.press('tab')

time.sleep(0.5)


pyperclip.copy('Relatório Mensal de Rendimentos')

pyautogui.hotkey('ctrl','v')

pyautogui.press('tab')

texto = f"""
escreva o texto aqui R${valor:.2f}
"""
print(texto)

pyperclip.copy(texto)

pyautogui.hotkey("ctrl", "v")

pyautogui.hotkey("ctrl", "enter")

#pyautogui.position() com esse comando você consegue saber onde seu mouse esta posicionado
#para utiliza-lo menlhor sempre use um time.sleep(6) 










