# compras de mercadoria p/ empresa 
# faturamento, quantidade e preço médio
import pyautogui # obs: tbm é chamada de "py" (apelido) -> automatiza mouse, teclado e tela do computador; não lê base de dados, p/ análises de dados e afins
import webbrowser
import time
import pandas

# pyautogui.click clique com o mouse
# pyautogui.write escrever um texto
# pyautogui.press apertar uma tecla
# pyautogui.hotkey apertar uma combinação de teclas (ex: Ctrl + D)
#pyautogui.PAUSE = 1 # espera em todos os comandos -> obs: time.sleep(x) : espera em determinado comando

# passo a passo

# passo 1: entrar no sistema da empresa

#pyautogui.press ("win")
#pyautogui.write("chrome")
#pyautogui.press("enter")
#  ou:
#webbrowser.open('x')
# colar o link de acesso escolhido (sistema da empresa vinculado ao drive -> p/ extrair a base de dados)
# ex: webbrowser.open('https://www.google.com/')
# pode ser que o navegador tenha que carregar
#time.sleep(5)

# passo 2: fazer login
# clicar no espaço de login
# escrever o login

#time.sleep(5) -> tempo estimado p/ calcular a posição
#print(pyautogui.position()) -> p/ calular a posição
# login:
#pyautogui.click(x=913, y=462)
#pyautogui.write("meu_login")
# senha:
#pyautogui.click(x=902, y=535)
#pyautogui.write("minha_senha")
# acessar:
#pyautogui.click(x=906, y=646)
#ime.sleep(3)

# passo 3: exportar a base de dados = 2-3 posições de click (forma padronizada = lugares fixos)
#pyautogui.click(x=470, y=416) # (download padronizado)
# pyautogui.click(x=611, y=302) # (download padronizado)
#time.sleep(3)
# ou: clicar c/botão direito + fazer download
#pyautogui.click(x=470, y=416,button= "right" )
#pyautogui.click(x=560, y=501) -> download
# obs: se for o botão esquerdo, usa-se: button= "left" 
# obs: se a tela tiver outra resolução, será preciso adaptar os pontos "x" e "y"


# passo 4: calcular os indicadores
import pandas as pd # -> p/ eitura e análise de dados
# obs: caso não esteja usando o jupyter, será preciso instalar: pandas, numpy e openpyxl

tabela = pandas.read_csv(r"C:\Users\X\Downloads\Compras (1).csv", sep=";") # copiar o caminho que se encontra o arquivo!
# obs: p/ copiar o caminho inteiro do arquivo, por estar em locais diferentes, deve-se ir até o arquivo e "copiar como caminho" = Ctrl + Shift + C
# o "r" diz que a leitura deve ser feita exatamente na forma que ele está escrito
print(tabela) # -> o "display" vem em formato de tabela, mais padronizada e de fácil leitura. também poderia ter usado o print, mas não sairia tão bonito
# no "csv" o normal é ter os dados separados por ",", mas tamém podem aparecer separados por ";". p/ resolver, basta usar o "sep"
# se oarquivo mudar o padrão, basta mudar o nome do arquivo, colocando uma variável dentro de um texto fixo. ex: compras -> compras(1) -> compras (2)...
# só se arredonda um número no final (aspecto monetário) -> p/ evitar mudanças visíveis no número trabalhado
total_gasto = tabela["ValorFinal"].sum()
quantidade = tabela["Quantidade"].sum()
preco_medio = total_gasto/quantidade
print(total_gasto)
print(quantidade)
print(preco_medio)

import pyperclip
# passo 5: enviar um e-mail para o chefe
# entrar no email: https://mail.google.com/mail/u/0/#inbox
webbrowser.open('https://mail.google.com/mail/u/0/#inbox')
time.sleep(5)
# ou: 
#pyautogui.press("win")
#pyautogui.write(outlook)
#pyautogui.press("enter")

# clicar no botão de escrever
pyautogui.click(x=38, y=225)

# preencher as informações de email
pyautogui.write("X@gmail.com") # email escolhido para envio
pyautogui.press("tab")  # escolher destinatário 
# p/ ir até o campo de "assunto". ou vc clica em "assunto", ou vc clica na tecla "tab " novamente
pyautogui.press("tab") # passar para o campo assunto
pyperclip.copy("Relatório de vendas") # -> p/ habilitar caracteres especiais
pyautogui.hotkey("ctrl", "v")

pyautogui.press("tab") # passando p/ o corpo do email
 
# p/ escrever em várias linhas, basta abrir 3 aspas simples/duplas
# p/ utilizar os dados da tebela no corpo do email, basta utlizar "{}" c/ os nomes das colunas/dados selecionados
# p/ utilizar a data do dia, basta utilizar a biblioteca "datetime" 
# p/ formatar os números é necessário utilizar um código de formatação -> ":,xf" -> x= quantidade de casas decimais. antes disso, e, antes mesmo de escrever o nome da tabela desejada, colocar o "R$"

texto = f"""
Prezados, 
Segue o relatório de compras

Total Gasto: R${total_gasto:,.2f} 
Quantidade de Produtos:  {quantidade:,}
Preço Médio: R${preco_medio:,.2f}

Qualquer dúvida, estou à disposição.
At.te,
G.B
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")


# enviar
pyautogui.hotkey("ctrl", "enter")

### obs: código base do projeto. alterações "ok" na formatalção "ipynb" (p/ rodar)