import pyautogui as pyauto
import pyperclip as pyclip
import time
import pandas as pd
pyauto.PAUSE = 1

pyauto.press("win")
pyauto.write("google chrome")
pyauto.press("enter")
pyclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyauto.hotkey("ctrl", "v")
pyauto.press("enter")

pyauto.sleep(5)

pyauto.click(x= 427, y= 277,clicks = 2)
time.sleep(2)

pyauto.click(x=438, y=416)
pyauto.click(x=1703, y=164)
pyauto.click(x=1513, y=560)

time.sleep(5)

tabela = pd.read_excel(r"D:/Downloads/Vendas - Dez.xlsx")
print(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

pyauto.hotkey("ctrl", "t")
pyclip.copy("https://outlook.live.com/mail/0/inbox")
pyauto.hotkey("ctrl", "v")
pyauto.press("enter")

time.sleep(5)

pyauto.click(x=159, y=146)

pyauto.write("pablo.bras@hotmail.com")
pyauto.press("tab")
pyclip.copy("Tabela Gerencial")
pyauto.hotkey("ctrl", "v")
pyauto.press("tab")

texto = f'''Prezados, boa tarde!

Segue planilha atualizada

Faturamento:R${faturamento:,.2f}
Quantidade: {quantidade:,}

Pablo Danillo
'''
pyclip.copy(texto)
pyauto.hotkey("ctrl", "v")
pyauto.hotkey("ctrl", "enter")
