# Importando as bibliotecas necessárias
# Caso seja necessário baixar as bibliotecas '!pip install nome da biblioteca'
import pyautogui
import pyperclip
import pandas as pd
from time import sleep


# Buscando e baixando o arquivo
pyautogui.hotkey('ctrl', 't')
link = 'endereço do local de onde o arquivo será baixado'
pyperclip.copy(link)
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('enter')
sleep(5)
pyautogui.click(161, 285)
sleep(3)
pyautogui.click(94, 275)
sleep(3)
pyautogui.press('enter')

# Tratando os indicadores desejados
tabela = pd.read_excel('nome_do_arquivo.xlsx')
# display(tabela)
faturamento = tabela['Valor Final'].sum()
quantidade = tabela['Quantidade'].sum()
print(faturamento)
print(quantidade)

# Enviando o e-mail para diretoria
pyautogui.hotkey('ctrl', 't')
pyautogui.write('https://mail.google.com/mail')
pyautogui.press('enter')
sleep(6)

pyautogui.click(83, 184)
sleep(2)
pyautogui.write('seuemail+diretoria@gmail.com')
pyautogui.press('tab')
pyautogui.press('tab')
pyperclip.copy('Relatório de vendas')
pyautogui.hotkey('ctrl', 'v')
pyautogui.press('tab')

texto = f"""
Prezados, bom dia!
Envio do relatório de vendas de ontem.
Faturamento: R$: {faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade}
"""
pyperclip.copy(texto)
pyautogui.hotkey('ctrl', 'v')
sleep(1)
pyautogui.press('tab')
pyautogui.press('enter')

# Código para buscar a posição dos clicks do mouse
pyautogui.position()
sleep(5)
print(pyautogui.position())
