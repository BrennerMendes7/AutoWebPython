# webdriver -> Chrome -> chromedriver --- colocar o driver no local onde python está localizado
# pip install pandas, numpy, openpyxl, selenium
from selenium import webdriver
from selenium.webdriver.common.keys import Keys # importar as teclas
from selenium.webdriver.common.by import By # importar pesquisador de elementos da página
import pandas as pd

navegador = webdriver.Chrome()
# navegador = webdriver.Chrome("chromedriver.exe) --> quando o chromedriver tiver no mesmo local do projeto

# Passo 1: Pegar a cotação do Dólar
navegador.get("https://www.google.com.br/") # entrar no link
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div[2]/div[2]/input').send_keys("cotação dolar") # encontar campo de pesquisa do Google e escrever no mesmo, usar aspas simples no xpath
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div[2]/div[2]/input').send_keys(Keys.ENTER) # apertar a tecla enter no campo de pesquisa
cotacao_dolar = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value') # pegar valor do elemento
print(cotacao_dolar)

# Passo 2: Pegar a cotação do Euro
navegador.get("https://www.google.com.br/") # entrar no link
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div[2]/div[2]/input').send_keys("cotação euro") # encontar campo de pesquisa do Google e escrever no mesmo, usar aspas simples no xpath
navegador.find_element(By.XPATH, '/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div[2]/div[2]/input').send_keys(Keys.ENTER) # apertar a tecla enter no campo de pesquisa
cotacao_euro = navegador.find_element(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value') # pegar valor do elemento
print(cotacao_euro)

# Passo 3: Pegar a cotação do Ouro
navegador.get("https://www.melhorcambio.com/ouro-hoje") # entrar no link
cotacao_ouro = navegador.find_element(By.XPATH, '//*[@id="comercial"]').get_attribute('value') # pegar valor do elemento
cotacao_ouro = cotacao_ouro.replace(",", ".") # trocar vírgula por ponto
print(cotacao_ouro)
navegador.quit() # fechar navegador

# Passo 4: Importar a base e atualizar as cotações na minha base
tabela = pd.read_excel("Produtos.xlsx") # ler arquivo excel usando pandas
print(tabela)
tabela.loc[tabela['Moeda'] == 'Dólar', 'Cotação'] = float(cotacao_dolar) # localizar linha e coluna, verificar a moeda com == e atualizar com a nova cotação
tabela.loc[tabela['Moeda'] == 'Euro', 'Cotação'] = float(cotacao_euro) # localizar linha e coluna, verificar a moeda com == e atualizar com a nova cotação
tabela.loc[tabela['Moeda'] == 'Ouro', 'Cotação'] = float(cotacao_ouro) # localizar linha e coluna, verificar a moeda com == e atualizar com a nova cotação
print(tabela['Cotação']) # mostrar a coluna Cotação atualizada

# Passo 5: Calcular os novos preços e salvar/exportar a base de dados
# Preço de Compra = Preço Original * Cotação
tabela['Preço de Compra'] = tabela['Preço Original'] * tabela['Cotação'] # calcular e atualizar a coluna Preço de Compra linha por linha com o novo preço
print(tabela['Preço de Compra']) # mostrar a coluna Preço de Compra atualizada

# Preço de Venda = Preço de Compra * Margem
tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem'] # calcular e atualizar a coluna Preço de Venda linha por linha com o novo preço
print(tabela['Preço de Venda']) # mostrar a coluna Preço de Venda atualizada

tabela.to_excel("Produtos Novo.xlsx", index=False) # exportar a tabela atualizada para um novo arquivo excel sem os índices do Python

