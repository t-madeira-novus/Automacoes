import pandas as pd
import pyautogui
from tqdm import tqdm
import time

df = pd.read_csv('responsaveis_empresas.csv', encoding='latin-1', sep=';')

time.sleep(2)
for i in tqdm(df.index):
    pyautogui.press('f8') # Abrir selecionador de empresa
    pyautogui.write(str(df.at[i, 'Códigos'])) # Digitar código da empresa
    pyautogui.press('enter') # Ativar empresa
    time.sleep(7) # Esperar carregar a empresa

    # Abrir menu 'Controle'
    pyautogui.keyDown('alt')  # Segurar Alt
    pyautogui.press('c')
    pyautogui.keyUp('alt')  # Soltar Alt

    # Abrir menu 'Empresas'
    pyautogui.press('down')
    pyautogui.press('enter')

    # Abrir menu 'Observações'
    pyautogui.click(725, 246)

    # Abrir menu 'Contábil', limpar o campo e escrever novo responsável
    pyautogui.press('tab')
    pyautogui.press('tab')
    pyautogui.press('right')
    pyautogui.press('right')

    pyautogui.click(732, 446)
    k = 0
    while k < 300:
        pyautogui.press('backspace')
        k += 1

    pyautogui.write(str(df.at[i, 'Contábil']))

    # Abrir menu 'Fiscal', limpar o campo e escrever novo responsável
    pyautogui.keyDown('shift')  # Segurar Shift
    pyautogui.press('tab')
    pyautogui.keyUp('shift')  # Soltar Shift
    pyautogui.press('right')

    pyautogui.click(732, 446)
    k = 0
    while k < 300:
        pyautogui.press('backspace')
        k += 1

    pyautogui.write(str(df.at[i, 'Fiscal']))

    # Abrir menu 'Pessoal', limpar o campo e escrever novo responsável
    pyautogui.keyDown('shift')  # Segurar Shift
    pyautogui.press('tab')
    pyautogui.keyUp('shift')  # Soltar Shift
    pyautogui.press('right')

    pyautogui.click(732, 446)
    k = 0
    while k < 300:
        pyautogui.press('backspace')
        k += 1

    pyautogui.write(str(df.at[i, 'Pessoal']))

    # Salvar
    pyautogui.click(605, 730)
    time.sleep(3)
    pyautogui.press('esc')