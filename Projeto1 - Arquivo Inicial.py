#!/usr/bin/env python
# coding: utf-8

# # Automação de Sistemas e Processos com Python
# 
# ### Desafio:
# 
# Para controle de custos, todos os dias, seu chefe pede um relatório com todas as compras de mercadorias da empresa.
# O seu trabalho, como analista, é enviar um e-mail para ele, assim que começar a trabalhar, com o total gasto, a quantidade de produtos compradas e o preço médio dos produtos.
# 
# E-mail do seu chefe: para o nosso exercício, coloque um e-mail seu como sendo o e-mail do seu chefe<br>
# Link de acesso ao sistema da empresa: https://pages.hashtagtreinamentos.com/aula1-intensivao-sistema
# 
# Para resolver isso, vamos usar o pyautogui, uma biblioteca de automação de comandos do mouse e do teclado

# In[10]:


import pyautogui
import time 

pyautogui.PAUSE = 1

pyautogui.hotkey("ctrl", "t")
pyautogui.write("https://pages.hashtagtreinamentos.com/aula1-intensivao-sistema")
pyautogui.press("enter")

time.sleep(5)

pyautogui.click(x=602, y=339)
pyautogui.write("meu_login")

pyautogui.click(x=634, y=414)
pyautogui.write("minha_senha")

pyautogui.click(x=630, y=495)
time.sleep(9)


# In[11]:


pyautogui.click(x=352, y=316) # clica no arquivo de compras
pyautogui.click(x=706, y=173) # clica nos 3 pontinhos
pyautogui.click(x=717, y=544) # faz o download
time.sleep(3)


# In[12]:


import pandas as pd 
tabela = pd.read_csv(r"C:\Users\Lenovo\Downloads\Compras.csv", sep = ";")
display(tabela)


# In[13]:


total_gasto = tabela["ValorFinal"].sum()
quantidade = tabela["Quantidade"].sum()
preco_medio = total_gasto / quantidade
print(total_gasto)
print(quantidade)
print(preco_medio)


# In[ ]:


import pyperclip

pyautogui.hotkey("ctrl", "t")
pyautogui.write("https://mail.google.com/mail/u/0/#inbox")
pyautogui.press("enter")

time.sleep(8)

pyautogui.click(x=51, y=159)

pyautogui.write("matvaroli14@gmail.com")
pyautogui.press("tab") # escolher o destinatario

pyautogui.press("tab") # passar para o campo assunto
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")

pyautogui.press("tab")

texto = f"""
Prezados,
Segue o relatório de compras

Total Gasto: R${total_gasto:,.2f}
Quantidade de Produtos: {quantidade:,}
Preço Médio: R${preco_medio:,.2f}

Qualquer dúvida, é só falar.
Att., Miguel
"""

pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")


# In[ ]:


#time.sleep(5)
#print(pyautogui.position())


# 
