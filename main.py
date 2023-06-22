from openpyxl import load_workbook
import os
from selenium import webdriver as wd
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import pyautogui as tempoEspera


nomeCaminhoArquivo = 'C:\\Users\\bruno\\Documents\\Python_RPA\\Dados_Formulario\\preencheDados.xlsx'
planilha_aberta = load_workbook(filename=nomeCaminhoArquivo)

sheet_selecionada = planilha_aberta['Plan1'] 

for linha in range(2, len(sheet_selecionada['A']) +1):
    nome = sheet_selecionada['A%s' % linha].value
    email = sheet_selecionada['B%s' % linha].value
    telefone = sheet_selecionada['C%s' % linha].value
    sexo = sheet_selecionada['D%s' % linha].value
    sobre = sheet_selecionada['E%s' % linha].value
    tempoEspera.sleep(3)
    
    navegador_formulario = wd.Chrome()
    navegador_formulario.get("https://pt.surveymonkey.com/r/T658MBP")
    tempoEspera.sleep(3)


    navegador_formulario.find_element(By.XPATH,'//*[@id="138626355"]').send_keys(nome) 
    tempoEspera.sleep(1)
    
    navegador_formulario.find_element(By.XPATH,'//*[@id="138626515"]').send_keys(email) 
    tempoEspera.sleep(1)
    
    navegador_formulario.find_element(By.XPATH,'//*[@id="138626597"]').send_keys(telefone) 
    tempoEspera.sleep(1)
    
    if sexo == 'Masculino':
        navegador_formulario.find_element(By.XPATH,'//*[@id="138629629_1030175222_label"]/span[1]').click() 
        tempoEspera.sleep(1)
    else:
        navegador_formulario.find_element(By.XPATH,'//*[@id="138629629_1030175223_label"]/span[1]').click()

    navegador_formulario.find_element(By.XPATH,'//*[@id="138629673"]').send_keys(sobre) 
    tempoEspera.sleep(3)
    
    navegador_formulario.find_element(By.XPATH,'//*[@id="patas"]/main/article/section/form/div[2]/button').click()

    