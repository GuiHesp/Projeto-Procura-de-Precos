#!/usr/bin/env python
# coding: utf-8

# ## Inicialização do Navegador

# In[ ]:


#importar
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pandas as pd

#criar navegador
servico = Service(ChromeDriverManager().install())
nav = webdriver.Chrome(service=servico)

#importar base de dados e visualizar
tabela_pd = pd.read_excel(r'Buscas.xlsx')
display(tabela_pd)


# In[ ]:


tabela_pd.info()


# ## Definição das funções de busca do Google Shopping e Buscapé

# In[ ]:


import time

def verificar_tem_termos_banidos(lista_termos_banidos, nome):
    tem_termos_banidos = False
    for palavra in lista_termos_banidos:
        if palavra in nome:
            tem_termos_banidos = True
    return tem_termos_banidos

def verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome):
    tem_todos_termos_produtos = True
    for palavra in lista_termos_nome_produto:
        if palavra not in nome:
            tem_termos_banidos = False
    return tem_todos_termos_produtos


def busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    lista_ofertas = []
    
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)
    #termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_nome_produto = produto.split(" ")
    produto = produto.lower()
    
    #entrar no google
    nav.get("https://www.google.com/")

    #pesquisar o nome do produto no google
    nav.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys(produto)
    nav.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys(Keys.ENTER)
    
    #clicar na aba shopping
    lista_itens = nav.find_elements(By.CLASS_NAME, 'hdtb-mitem')
    for item in lista_itens:
        if "Shopping" in item.text:
            item.click()
            break
    
    #pegar o preco do produto no shopping
    lista_resultados = nav.find_elements(By.CLASS_NAME, 'i0X6df')

    for resultado in lista_resultados:
        nome = resultado.find_element(By.CLASS_NAME, 'tAxDx').text
        nome = nome.lower()

        #analisar se ele não tem nenhum termo banido
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)

        #analisar se ele tem TODOS os termos do nome do produto
        tem_todos_termos_produtos = verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome)

        #selecionar só os elementos que tem_termos_banidos = False e ao msm tempo tem_todos_termos_produtos = True
        if not tem_termos_banidos and tem_todos_termos_produtos:
            preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
            preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            preco = float(preco)
            
            #se o preco ta entre o preco min e max
            if preco_minimo <= preco <= preco_maximo:
                elemento_referencia = resultado.find_element('class name', 'bONr3b')
                elemento_pai = elemento_referencia.find_element('xpath' , '..')
                link = elemento_pai.get_attribute('href')

                lista_ofertas.append((nome, preco, link))
    return lista_ofertas

def busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo):
    #tratar valores
    lista_ofertas = []
    
    #termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_nome_produto = produto.split(" ")
    produto = produto.lower()
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)
    
    #buscar no buscape
    nav.get('https://www.buscape.com.br/')
    nav.find_element('xpath', '//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(produto, Keys.ENTER)
    
    #pegar os resultados
    #time.sleep(3)
    while len(nav.find_elements('class name', 'Select_Select__1S7HV')) < 1:
        time.sleep(1)
    
    lista_resultados = nav.find_elements('class name', 'SearchCard_ProductCard_Inner__7JhKb')
    
    for resultado in lista_resultados:
        preco = resultado.find_element('class name', 'Text_MobileHeadingS__Zxam2').text
        nome = resultado.find_element('class name', 'SearchCard_ProductCard_Name__ZaO5o').text
        nome = nome.lower()
        link = resultado.get_attribute('href')
        
        #analisar se ele não tem nenhum termo banido
        tem_termos_banidos = verificar_tem_termos_banidos(lista_termos_banidos, nome)

        #analisar se ele tem TODOS os termos do nome do produto
        tem_todos_termos_produtos = verificar_tem_todos_termos_produto(lista_termos_nome_produto, nome) 
    
        #analisar se o preco esta entre a faixa min-max estabelecido
        if not tem_termos_banidos and tem_todos_termos_produtos:
            preco = preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            preco = float(preco)
            if preco_minimo <= preco <= preco_maximo:
                lista_ofertas.append((nome, preco, link))
                
    #retornar a lista de ofertas do buscape
    return lista_ofertas


# ## Construindo a nossa lista de ofertas

# In[ ]:


tabela_ofertas = pd.DataFrame()

for linha in tabela_pd.index:

    produto = tabela_pd.loc[linha, 'Nome'] #separar o '64 gb' para pesquisar melhor.Como ñ existe iphone 64, n tem problema.
    termos_banidos = tabela_pd.loc[linha, 'Termos banidos']
    preco_minimo = tabela_pd.loc[linha, 'Preço mínimo']
    preco_maximo = tabela_pd.loc[linha, 'Preço máximo']

    lista_ofertas_google_shopping = busca_google_shopping(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping:
        tabela_google_shopping =pd.DataFrame(lista_ofertas_google_shopping, columns=['produto', 'preco', 'link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google_shopping])
    else:
        tabela_google_shopping = None
    
    lista_ofertas_buscape = busca_buscape(nav, produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_buscape =pd.DataFrame(lista_ofertas_buscape, columns=['produto', 'preco', 'link'])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_buscape])
    else:
        tabela_buscape = None
        
    
display(tabela_ofertas)    


# ## Exportando para Excel

# In[ ]:


#exportar para o excel
tabela_ofertas.to_excel("Ofertas.xlsx", index=False)


# ## Enviando E-mail

# In[ ]:


#enviar para o e-mail da tabela
import smtplib
import email.message

def enviar_email():  
    if len(tabela_ofertas) > 0:
        corpo_email = f"""
        <p>Produtos encontrados dentro da faixa de preço desejada:</p>
        {tabela_ofertas.to_html(index=False)}
        <p>Tamo Junto.</p>
        """

        msg = email.message.Message()
        msg['Subject'] = "Projeto 2 - Selenium"
        msg['From'] = 'guihesp@gmail.com'
        msg['To'] = 'guihesp@gmail.com'
        password = 'password' 
        msg.add_header('Content-Type', 'text/html')
        msg.set_payload(corpo_email )

        s = smtplib.SMTP('smtp.gmail.com: 587')
        s.starttls()
        # Login Credentials for sending the mail
        s.login(msg['From'], password)
        s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
        print('Email enviado')
        nav.quit()

enviar_email()        

