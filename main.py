import pyautogui
import pandas as pd
import pyperclip
import time


pyautogui.PAUSE = 3

# Passo 1: Entrar no sistema (no nosso caso, entrar no link)
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(5)

# Passo 2: Navegar até o local do relatório (entrar na pasta Exportar)
pyautogui.click(x=417, y=296, clicks=2)
time.sleep(7)

# Passo 3: Fazer o download do relatório
pyautogui.click(x=475, y=426)
pyautogui.click(x=1219, y=190)
pyautogui.click(x=1043, y=617)
pyautogui.click(x=73, y=370)
pyautogui.click(x=670, y=682)


time.sleep(7)

# Passo 4: Calcular os indicadores

tabela = pd.read_excel(r"C:\Users\maaxg\Downloads/Vendas - Dez.xlsx")
print(tabela)
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()
time.sleep(10)

# Passo 5: Entrar no email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(10)

# Passo 6:Abrir corpo de mensagem
pyautogui.click(x=97, y=214)
time.sleep(5)
pyautogui.click(x=889, y=295)

# selecionar destinatario
pyperclip.copy("mgrodrigues920@gmail.com")
pyautogui.hotkey("ctrl", "v")
time.sleep(5)

#Digitar assunto do email
#pula pro campo de assunto
pyautogui.press("tab")
pyautogui.write("Relatório de vendas")
pyautogui.press("tab")

texto = f"""
Prezados, bom dia

O faturamento de ontem foi de: R${faturamento:,.2f}
A quantidade de produtos foi de: {quantidade:,}

Abs
Maxwell Gallucci Rodrigues"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
time.sleep(2)


# clicar no botão enviar

# apertar Ctrl Enter
pyautogui.hotkey("ctrl", "enter")
time.sleep(5)
pyautogui.position()
