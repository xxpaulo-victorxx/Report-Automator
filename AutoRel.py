'''
Automação de Sistemas e Processos com Python
Desafio:
Todos os dias, o nosso sistema atualiza as vendas do dia anterior. O seu trabalho diário, como analista, é enviar um e-mail para a diretoria, assim que começar a trabalhar, com o faturamento e a quantidade de produtos vendidos no dia anterior

E-mail da diretoria: seugmail+diretoria@gmail.com
Local onde o sistema disponibiliza as vendas do dia anterior: https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing

Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado
'''

'''
Passo a Passo para resolver o desafio

Passo 1 - Entra no Sistema da Empresa (https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing)

Passo 2 - Navegar até o local onde está a base de dados (pasta Exportar)

Passo 3 - Baixar planilha de vendas

Passo 4 - Calcular o faturamento e quantidade ce produtos vendidos da base

Passo 5 - Enviar email para a diretoria com a quantidade e o faturamento de que calculamos
'''

# Início
#---------------------------------------
import pyautogui
import time
import pyperclip

pyautogui.PAUSE = 1

# Passo 1
#---------------------------------------

# Abrir uma nova aba
pyautogui.hotkey('ctrl','t')
'''
pyautogui.press('win')
pyautogui.press('Chrome')
pyautogui.press('enter')
'''
# Entrar no link do sistema
link = "https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga?usp=sharing"
pyperclip.copy(link)
pyautogui.hotkey('ctrl','v')
pyautogui.press('Enter')

# Passo 2
#---------------------------------------

time.sleep(5)
pyautogui.click(374,326, clicks=2)

# Passo 3
#---------------------------------------

pyautogui.click(374,398) # clicar no arquivo
pyautogui.click(1688,194) # clicar nos 3 pontos
pyautogui.click(1476,640) # clicar no botao de fazer download
time.sleep(10)

time.sleep(5)
pyautogui.position()

'''
Vamos agora ler o arquivo baixado para pegar os indicadores
	-> Faturamento1
	-> Quantidade de Produtos
'''

# Passo 4
#---------------------------------------
import pandas as pd

tabela = pd.read_excel(r"C:\Users\Usuario\Downloads\Vendas - Dez.xlsx")
faturamento = tabela["Valor Final"].sum() # soma da coluba Valor Final
quantidade = tabela["Quantidade"].sum() # tabela da coluba Quantidade
#print (tabela)
display(tabela)
#print(faturamento)
#print(quantidade)


'''
Vamos agora enviar um e-mail pelo gmail
'''

# Passo 5
#---------------------------------------

# Abrir uma aba
pyautogui.hotkey('ctrl','t')
# Entrar no Gmail
link = "https://mail.google.com/"
pyperclip.copy(link)
pyautogui.hotkey('ctrl','v')
pyautogui.press('Enter')
time.sleep(7)

# clicar no botao escrever
pyautogui.click(79, 198)

# digitar destinatario
email = "pv.innocencio@poli.ufrj.br"
pyperclip.copy(email)
pyautogui.hotkey('ctrl','v')
#pyautogui.write("pv.innocencio@gmail.com")
#pyautogui.press("tab") # escolher o email
pyautogui.press("tab") # passar para o campo de assunto

# digitar assunto
assunto = "Relatório de Vendas"
pyperclip.copy(assunto)
pyautogui.hotkey('ctrl','v')
#pyautogui.write(r"Relatório de Vendas")
pyautogui.press("tab") # passar para o campo de escrever email

# digitar o corpo do email
texto_email= f"""
Prezados, bom dia

O faturamento de ontem foi de: {faturamento}
A quantidade de produtos foi de: {quantidade}

Abs,
Paulo Victor Innocencio
"""
pyautogui.write(texto_email)

# clicar em enviar
pyautogui.hotkey('ctrl','Enter') # comando de enviar email (gmail )


'''
Use esse código para descobrir qual a posição de um item que queira clicar
	-> Lembre-se: a posição na sua tela é diferente da posição na minha tela

time.sleep(5)
pyautogui.position()
'''


