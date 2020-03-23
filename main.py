from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import datetime

# driver = webdriver.Chrome()
dt = datetime.datetime.today()
excel_att = pd.read_excel('Excel\Att-{}-{}-{}.xls'.format(dt.day, dt.month, dt.year))
excel_clean = excel_att.loc[excel_att['Último motivo'] == '-']
nomes = []
nomes = excel_clean['Descrição'].tolist()
lista_nomes = []

# for nome in nomes:
#     print(nome)
driver = webdriver.Chrome('ChromeDriver\chromedriver.exe')
driver.get("https://suporte.bridsolucoes.com.br/chamados?categoria=10&subcategoria=57&situacao=2&limit=200")

#Login
wait = WebDriverWait(driver, 1000)
user = driver.find_element_by_id('username')
user.send_keys("")
pasw = driver.find_element_by_id('password')
pasw.send_keys("")
wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="buscaid"]/div[4]/div[4]/input')))

#Chamados Existentes Em Andamento
table_id = driver.find_elements_by_class_name("tableOftickets")[0]
body = table_id.find_elements(By.TAG_NAME, "tbody")
rows = body[1].find_elements(By.TAG_NAME, "tr") # get all of the rows in the table
for row in rows:
    # Get the columns (all the column 2)        
    col = row.find_elements(By.TAG_NAME, "td")[3] #note: index start from 0, 1 is col 2
    lista_nomes.append(col.text)

#Chamados Existentes Pendentes
driver.get("https://suporte.bridsolucoes.com.br/chamados?categoria=10&subcategoria=57&situacao=3")
table_id = driver.find_elements_by_class_name("tableOftickets")[0]
body = table_id.find_elements(By.TAG_NAME, "tbody")
rows = body[1].find_elements(By.TAG_NAME, "tr") # get all of the rows in the table
for row in rows:
    # Get the columns (all the column 2)        
    col = row.find_elements(By.TAG_NAME, "td")[3] #note: index start from 0, 1 is col 2
    lista_nomes.append(col.text)


x = 0 #Count
for nome in lista_nomes:
    if '.' in nome:
        lista_nomes[x] = nome.split('. ')[1] #Removendo número do nome do Canal
    x += 1

#Print dos nomes de canais com chamado    
for nome_lista in lista_nomes:
    print(nome_lista)
print("// Fim Canais Com chamado aberto")

print(len(nomes))
for nome in lista_nomes:
    print(nome)
    if nome in nomes:
        nomes.remove(nome)  #Removendo Canais que já possuem chamados da lista de novos chamados
        print(nome+" --Canal Excluido")

    
print("//Fim Canais Excel")

for nome_final in nomes:
    print(nome_final)
print("//Fim Canais sem chamado aberto e desatualizados")
    


# Add Chamado
driver.get("https://suporte.bridsolucoes.com.br/chamados/addemmassa")
#wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="titulo"]')))
driver.find_element_by_xpath('//*[@id="categoria-id"]/option[2]').click()
driver.find_element_by_xpath('//*[@id="subcategoria-id"]/option[2]').click()
driver.find_element_by_xpath('//*[@id="grupo-responsavel"]/option[2]').click()
driver.find_element_by_xpath('//*[@id="responsavel"]/option[1]').click()
driver.find_element_by_xpath('//*[@id="situacao"]/option[2]').click()
driver.find_element_by_xpath('//*[@id="tipo"]/option[1]').click()
titulo = driver.find_element_by_id('titulo')
titulo.send_keys("Atualização - Agroview")
note = driver.find_element_by_xpath('//*[@id="camposChamado"]/div[3]/div/div[9]/div/div/div[3]/div[2]/p')
note.send_keys('''Boa tarde 



Favor, verificar a atualização do Canal e tomar as devidas providências 





Obrigado !!''')
canal = driver.find_element_by_xpath('//*[@id="camposChamado"]/div[3]/div/div[3]/div/span/span[1]/span/ul/li/input')
for nome in nomes:
    canal.send_keys(nome)
    canal.send_keys(Keys.ENTER)
btn = driver.find_element_by_xpath('//*[@id="enviarForm"]')
btn.click()


#Close
driver.quit() 




