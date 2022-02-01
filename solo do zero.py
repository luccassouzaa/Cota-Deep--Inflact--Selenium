#!/usr/bin/env python
# coding: utf-8

# In[1]:


#Importações
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd
import pyautogui
import pyperclip
import time


# In[2]:


#Abrir um novo navegador e pesquisar as cotações
web = webdriver.Chrome()

web.get("https://www.google.com.br/")
web.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação dolar")
web.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacaodolar = web.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
display(cotacaodolar)

web.get("https://www.google.com.br/")
web.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys("Cotação euro")
web.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
cotacaoeuro = web.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
display(cotacaoeuro)

web.get("https://www.melhorcambio.com/ouro-hoje")
cotacaoouro = web.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute("value")
cotacaoouro = cotacaoouro.replace(",", ".")
display(cotacaoouro)

web.quit()


# In[3]:


#Pegar essas cotações e atualizar a base de dados
tabela = pd.read_excel("Produtos.xlsx")


tabela.loc[tabela["Moeda"]=="Dólar" , "Cotação"] = float(cotacaodolar)
tabela.loc[tabela["Moeda"]=="Euro" , "Cotação"] = float(cotacaoeuro)
tabela.loc[tabela["Moeda"]=="Ouro" , "Cotação"] = float(cotacaoouro)


# In[4]:


#calcular a base
tabela["Preço de Compra"] = tabela["Preço Original"] * tabela["Cotação"]
tabela["Preço de Venda"] = tabela["Preço de Compra"] * tabela["Margem"]

display(tabela)


# In[5]:


#salvar tabela

tabela.to_excel("ProdutoSolo.xlsx")

#criar variantes para indicar o lucro
compra = tabela["Preço de Compra"].sum()
compraform = "R$ " f"{compra:,.0f}" " Valor da compra para revenda"
print (compraform)

venda = tabela["Preço de Venda"].sum()
vendaform = "R$ " f"{venda:,.0f}" " Valor no fim do mês"
print (vendaform)

lucro = venda - compra
lucroporcentagem = lucro / compra
porcentagem1 = round((lucroporcentagem), 2)
porcentagem = (f'Porcentagem de lucro : {porcentagem1} %')
print(porcentagem)

lucro2 = "RS " f"{lucro:,.0f}" " Esse foi o lucro " 
print (lucro2)


# In[6]:


#atualizar os numeros por extenso (usaremos o inflact)
import inflect

p = inflect.engine()

words = p.number_to_words(compraform)
words1 = p.number_to_words(vendaform)
words2 = p.number_to_words(lucro2)
words3 = p.number_to_words(porcentagem)

print (words)
print (words1)
print (words2)
print (words3)


# In[7]:


#Traduzir o texto extenso
from deep_translator import GoogleTranslator

tr1 = GoogleTranslator(source='auto', target='portuguese').translate(words)
tr2 = GoogleTranslator(source='auto', target='portuguese').translate(words1)
tr3 = GoogleTranslator(source='auto', target='portuguese').translate(words2)
tr4 = GoogleTranslator(source='auto', target='portuguese').translate(words3)

print (tr1)
print (tr2)
print (tr3)
print (tr4)


# In[11]:


#enviar email automatico

texto = f"""
Olá chefe, atualizei os dados das ultimas vendas, estao aqui

{compraform}
{tr1}

{vendaform}
{tr2}

{lucro2}
{tr3}

{porcentagem}
{tr4}

Atenciosamente, Luccas Souza
"""

web = webdriver.Chrome()

web.get("https://login.live.com/")
web.find_element(By.XPATH, '//*[@id="i0116"]').send_keys("luccasmelim@hotmail.com")
web.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()
web.find_element(By.XPATH, '//*[@id="i0118"]').send_keys("letycia10")
time.sleep(2)
pyautogui.click(x=647, y=538)
web.find_element(By.XPATH, '//*[@id="lightbox"]/div[3]/div/div[2]/div/div[3]/div[1]/div/label/span').click()
web.find_element(By.XPATH, '//*[@id="idSIButton9"]').click()

web.get('https://outlook.live.com/mail/0/')

time.sleep(10)

pyautogui.click(x=197, y=233)

time.sleep(1.3)

pyperclip.copy("luccasmelim@yahoo.com.br")
pyautogui.hotkey("Ctrl","v")
pyautogui.press("Tab")
pyperclip.copy("Olá chefe, está aqui os relatórios atualizados das vendas")
pyautogui.hotkey("ctrl","v")
pyautogui.press("Tab")
pyperclip.copy(f"{texto}")
pyautogui.hotkey("Ctrl","v")
pyautogui.hotkey("ctrl","enter")

web.quit()


# In[12]:


#provar que chegou o email

pyautogui.PAUSE = 1

pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
pyperclip.copy("https://mail.yahoo.com/")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

time.sleep(2)

pyautogui.click(x=538, y=313)


# In[ ]:




