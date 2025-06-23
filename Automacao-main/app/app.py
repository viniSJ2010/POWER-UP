from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
import openpyxl

driver = webdriver.Edge()
#1 - navegar ate o site
driver.get("https://contabilidade-devaprender.netlify.app")
sleep(5)
email = driver.find_element(By.XPATH,'//input[@id="email"]')
sleep(2)
#2 - digitar o email
email.send_keys("admin@contantabilidade.com")
#3 - digitar senha
senha = driver.find_element(By.XPATH,'//input[@id="senha"]')
sleep(2)
senha.send_keys("contabilidade1234")
#4 - clicar em entrar
botao_entrar = driver.find_element(By.XPATH,'//button[@id="Entrar"]')
sleep(2)
botao_entrar.click()
sleep(5)
#5 - extrair as informaçoes da planilia
empresas = openpyxl.load_workbook('./empresas.xlsx')
paginas_empresas = empresas['dados empresas']
#6 - clicar em cada campo e digitar as informaçoes
for linha in paginas_empresas.iter_rows(min_row=2,values_only=True):
    nome_empresa, email, telefone, endereço, cnpj, area_atuacao, quantidade_de_funcionarios, data_fundacao = linha
    driver.find_element(By.ID, 'nomeEmpresa').send_keys(nome_empresa)
    sleep(1)
    driver.find_element(By.ID, 'emailEmpresa').send_keys(email)
    sleep(1)
    driver.find_element(By.ID, 'telefoneEmpresa').send_keys(telefone)
    sleep(1)
    driver.find_element(By.ID, 'enderecoEmpresa').send_keys(endereço)
    sleep(1)
    driver.find_element(By.ID, 'cnpj').send_keys(cnpj)
    sleep(1)
    driver.find_element(By.ID, 'areaAtuacao').send_keys(area_atuacao)
    sleep(1)
    driver.find_element(By.ID, 'numeroFuncionarios').send_keys(quantidade_de_funcionarios)
    sleep(1)
    driver.find_element(By.ID, 'dataFundacao').send_keys(data_fundacao)
    sleep(1)

    driver.find_element(By.ID, 'Cadastrar').click()
    sleep(3)