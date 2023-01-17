import pyautogui
import pyperclip
import time
import pandas

link = 'https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga'
gmail = 'https://mail.google.com/mail/u/0/#inbox'
assunto = 'Relatório de vendas'
destinatario = 'Joao V'

pyautogui.PAUSE = 1

pyautogui.press('win')
pyautogui.write('chrome')
pyautogui.press('enter')

pyautogui.click(x=518, y=456)
pyautogui.hotkey('ctrl', 't')


pyperclip.copy(link)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)
pyautogui.click(x=354, y=308, clicks=2)
time.sleep(3)
pyautogui.click(x=366, y=349)
pyautogui.click(x=1167, y=186)
pyautogui.click(x=931, y=626)
time.sleep(5)

df = pandas.read_excel(r"C:\Users\joaov\Downloads\Vendas - Dez.xlsx")

faturamento = df['Valor Final'].sum()
qtde_produtos = df['Quantidade'].sum()

pyautogui.hotkey('ctrl', 't')
pyperclip.copy(gmail)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(6)
pyautogui.click(x=107, y=197)
pyperclip.copy(destinatario)
time.sleep(5)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")
pyautogui.press("tab")
pyperclip.copy(assunto)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab")

texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de:{qtde_produtos:,}

abs
João Pedro
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", 'v')
pyautogui.hotkey('ctrl', 'enter')

pyautogui.hotkey("ctrl", "w")
pyautogui.hotkey("ctrl", "w")
