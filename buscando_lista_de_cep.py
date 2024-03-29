from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service

from selenium.webdriver.common.by import By

#usar para esperar entre as ações definidas
import pyautogui as tempoPausa

#imports do excel
from openpyxl import load_workbook
import os

#instala versão atual do webdriver do navegador(no caso, chrome)
servico = Service(ChromeDriverManager().install())

navegador = webdriver.Chrome(service = servico)

#-------------------------- Excel -----------------------------------------------

#abre o arquivo .xlsx do caminho indicado e nome expecificado
nome_arquivo_cep = "E:\\dados_cep2.xlsx"
planilhaCriada = load_workbook(nome_arquivo_cep)

#variável que armazena a folha da planilha do excel onde os dados estão salvos
sheet_selecionada = planilhaCriada["CEP"]


#-------------- abre o navegador, acessa o site, pesquisa -----------------------
navegador.get("https://buscacepinter.correios.com.br/app/endereco/index.php")

tempoPausa.sleep(2)


#------------------ loop para buscar os cepse salvar -----------------------------

for linha in range(2, len(sheet_selecionada["A"]) + 1):

    tempoPausa.sleep(5)

    cep_pesquisa = sheet_selecionada['A%s' % linha].value
    navegador.find_element(By.NAME, "endereco").send_keys(cep_pesquisa)

    tempoPausa.sleep(5)

    navegador.find_element(By.NAME, "btn_pesquisar").click()

    tempoPausa.sleep(5)  

    rua = navegador.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[1]').text

    bairro = navegador.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[2]').text

    cidade = navegador.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[3]').text

    cep = navegador.find_element(By.XPATH, '//*[@id="resultado-DNEC"]/tbody/tr/td[4]').text

    sheet_dados = planilhaCriada["Dados"]
    linha_atual = len(sheet_dados['A']) + 1
      
    colunaA = "A" + str(linha_atual) #concatenando coluna e linha
    colunaB = "B" + str(linha_atual)
    colunaC = "C" + str(linha_atual)
    colunaD = "D" + str(linha_atual)

    sheet_dados[colunaA] = rua
    sheet_dados[colunaB] = bairro
    sheet_dados[colunaC] = cidade
    sheet_dados[colunaD] = cep

    tempoPausa.sleep(5)

    navegador.find_element(By.ID, "btn_nbusca").click()



planilhaCriada.save(filename=nome_arquivo_cep)

os.startfile(nome_arquivo_cep)

print('chegou no final')


