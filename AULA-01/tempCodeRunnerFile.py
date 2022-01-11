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