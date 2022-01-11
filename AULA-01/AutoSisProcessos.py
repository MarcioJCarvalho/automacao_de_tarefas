# 1º desafio

# biblioteca usada: PANDAS, pyautogui, pyperclip
# instalar, pandas, numpy, openpyxl
import pyautogui
import pyperclip
import time
import pandas as pd

pyautogui.PAUSE = 1
# passo 1: Entrar no sistema da empresa https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)


# passo 2: navegar no sistema até a pasta exportar
pyautogui.click(x=360, y=278, clicks=2)
time.sleep(2)

# passo 3: fazer o downloade da base de vendas
pyautogui.click(x=338, y=342) #click no arquivo
pyautogui.click(x=1391, y=152) #click nos 3 pontos
pyautogui.click(x=1175, y=536) #click no fazer download

time.sleep(2)
pyautogui.click(x=1002, y=692)
time.sleep(5)

# passo 4: inportar a base de vendar para o python
# r na string é para ignorar a \
tabela = pd.read_excel(r"D:\Faculdade\Inensivão de Python\Vendas - Dez.xlsx")
print(tabela)

# passo 5: calcular o faturamento e quantidade de produtos vendidos (os indicadores)
faturamento = tabela["Valor Final"].sum()
qtde_produtos = tabela["Quantidade"].sum()
print(faturamento)
print(qtde_produtos)

# passo 6: Enviar e-mail para a diretoria 
# link gmail: https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox

# abrir email
pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/0/?tab=rm&ogbl#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# clicar no botão escrever
pyautogui.click(x=99, y=173)
time.sleep(2)

# preencer o destino
pyautogui.write("marciojosedecarvalho@gmail.com")
pyautogui.press("tab") # seleciona o e-mail
pyautogui.press("tab") # muda para o campo de assunto

# preencer o assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("tab") # ir para o campo corpo do email

# escrever o e-mail
texto = f'''
Prezados, bom dia

O faturamento de ontem foi  de RS: {faturamento:,.2f}
A quantodade de produtos foi de: {qtde_produtos:,}

Abs
Mário J Carvalho
'''

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")

# clicar em enviar
pyautogui.hotkey("ctrl", "enter")
