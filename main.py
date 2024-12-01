from time import sleep
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys




#Entrando no site
opcoes = webdriver.ChromeOptions()
opcoes.add_argument(r'user-data-dir=chrome://version/profile path/usuario/nome do novo perfil')
opcoes.add_argument('--start-maximized')
opcoes.add_experimental_option("detach", True)
navegador = webdriver.Chrome(options=opcoes)
navegador.get('https://www.nuvemshop.com.br/login')
sleep(4)

#Pegando as informações na planilha
tabela = pd.read_excel('ControleNuvemshop.xlsx')
login = tabela.iloc[0, 0]
senha = tabela.iloc[0, 1]
print(tabela)

#Fazendo login
try:
    navegador.find_element(By.XPATH, '//*[@id="user-mail"]').send_keys(str(login))
    sleep(3)
    navegador.find_element(By.XPATH, '//*[@id="pass"]').send_keys(str(senha))
    sleep(3)
    navegador.find_element(By.XPATH, '//*[@id="register-page"]/div/div/div[2]/div[1]/div[3]/form/div[4]/p[2]/input').click()
    sleep(3)
except:
    pass
#Indo na aba de produtos
navegador.find_element(By.XPATH, '//*[@id="control-products"]/div[2]/p').click()
sleep(3)

#Pegando as informações na planilha
for i, nome in enumerate(tabela["Nome Produto"]):
    descricao = tabela.loc[i, 'Descricao']
    preco_venda = tabela.loc[i, 'Preco de Venda']
    cod_sku = tabela.loc[i, 'Codigo SKU']
    cod_barras = tabela.loc[i, 'Codigo de Barras']
    peso = tabela.loc[i, 'Peso (gramas)']
    comprimento = tabela.loc[i, 'Comprimento (centimetro)']
    altura = tabela.loc[i, 'Altura (centimetro)']
    largura = tabela.loc[i, 'Largura (centimetro)']
    categoria = tabela.loc[i, 'Categoria']
    subcategoria = tabela.loc[i, 'Subcategoria']

    # Iniciando cadastro de novo produto
    navegador.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/div/div/div[1]/div[3]/div/button[3]').click()
    sleep(2)
    #Nome e Descrição
    navegador.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/div/div/div[2]/div/form/div/div[1]/div[1]/div[2]/div/div/div/div[1]/div/input').send_keys(nome, Keys.TAB, str(descricao))
    sleep(1)
    #Preço
    navegador.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/div/div/div[2]/div/form/div/div[1]/div[4]/div[2]/div/div[1]/div/div[1]/div/div/div/input').send_keys(str(preco_venda))
    sleep(1)
    #Tipo de produto fisico
    navegador.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/div/div/div[2]/div/form/div/div[1]/div[6]/div[2]/div/div/div/div[2]/div/div/label[1]').click()
    #Estoque limitado
    navegador.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/div/div/div[2]/div/form/div/div[1]/div[6]/div[2]/div/div/div/div[2]/div/div/label[2]').click()
    sleep(1)
    #Código SKU
    navegador.find_element(By.XPATH, '//*[@id="input_sku"]').send_keys(str(cod_sku))
    #Código de barras e peso
    navegador.find_element(By.XPATH, '//*[@id="input_barcode"]').send_keys(str(cod_barras), Keys.TAB, str(peso))
    sleep(1)
    #Comprimento
    navegador.find_element(By.XPATH, '//*[@id="input_depth"]').send_keys(str(comprimento))
    sleep(1)
    #Altura
    navegador.find_element(By.XPATH, '//*[@id="input_height"]').send_keys(str(altura))
    sleep(1)
    #Largura
    navegador.find_element(By.XPATH, '//*[@id="input_width"]').send_keys(str(largura))
    sleep(1)
    #Categoria
    navegador.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/div/div/div[2]/div/form/div/div[1]/div[10]/div[2]/a').click()
    navegador.find_element(By.XPATH, '//*[@id="input_categorySearch"]').send_keys(str(categoria))
    sleep(1)
    navegador.find_element(By.XPATH, '//*[@id="input_categorySearch"]').send_keys(Keys.TAB, Keys.SPACE)
    sleep(1)
    navegador.find_element(By.XPATH, '//*[@id="input_categorySearch"]').send_keys(Keys.BACK_SPACE * len(categoria))
    sleep(1)
    #Subcategoria
    navegador.find_element(By.XPATH, '//*[@id="input_categorySearch"]').send_keys(str(subcategoria), Keys.TAB, Keys.SPACE, Keys.ESCAPE)
    sleep(1)
    #Salvando
    navegador.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/div/div/div[2]/div/form/div/div[2]/div[2]/button').click()
    sleep(3)

    #Voltando a página de produtos
    navegador.find_element(By.XPATH, '//*[@id="root"]/div/div[2]/div[2]/div/div/div[2]/div/form/div/div[2]/div[1]/button').click()
    sleep(3)

sleep(10)
navegador.quit()
