import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class Robo:

    nome_arquivo = "Robo.xlsx"
    df = pd.read_excel(nome_arquivo)
    #df.to_excel("resultado.xlsx")
    driver = webdriver.Firefox()
    driver.get('https://hospitalar.postalsaudeservicos.com.br/hospitalarprod/Sistema/Abertura/Login.aspx')
    main_page = driver.current_window_handle
    login = driver.find_element_by_xpath('//*[@id="cphBody_txtLogin"]').send_keys("05657698630")
    senha = driver.find_element_by_xpath('//*[@id="cphBody_txtSenha"]').send_keys("total@23")
    entrar = webdriver.common.action_chains.ActionChains(driver)
    entrar = driver.find_element_by_xpath('//*[@id="cphBody_btnSalvar"]')
    entrar.click()
    time.sleep(4)
    for handle in driver.window_handles: 
        if handle != main_page: 
            login_page = handle 
    driver.switch_to.window(login_page)
    time.sleep(2)
    geral = driver.find_element_by_xpath('//*[@id="cphBody_dtPapeis_lblPapel_0"]').click()
    time.sleep(4)
    recepcao = driver.find_element_by_xpath('//*[@id="cphBody_dtlAplicativos_lblNomegrpApl_2"]').click()
    time.sleep(7)
    


    for index,row in df.iterrows(): 
        prestador = driver.find_element_by_xpath('//*[@id="cphBody_Conteudo_ddlUniAtendimento_TextBox"]').send_keys(row["PRESTADOR"])
        time.sleep(2)
            
        s_ocupacional = driver.find_element_by_xpath('//*[@id="cphBody_Conteudo_btnSaudeOcupacional"]').click()
        time.sleep(2)
        s_ocupacional_dois = driver.find_element_by_id('cphBody_Conteudo_btnSaudeOcupacional').click()
        time.sleep(4)
            
        #tido_de_exame = driver.find_element_by_xpath('//*[@id="cphBody_Conteudo_ddlTipoAtendimento_TextBox"]').clear()
        exame = driver.find_element_by_xpath('//*[@id="cphBody_Conteudo_ddlTipoAtendimento_TextBox"]').send_keys(row["EXAME"])
        time.sleep(2)
        tipo_matricula = driver.find_element_by_xpath('//*[@id="cphBody_Conteudo_ddlTipoPesquisa_Button"]').click().perfom()
        time.sleep(2)


        matricula = driver.find_element_by_xpath('/html/body/form/div[3]/div[1]/div[3]/div[3]/div[2]/div[1]/ul/li[2]/div/div/ul/li[1]').click()
        time.sleep(2)
        informar_matricula = driver.find_element_by_xpath('//*[@id="cphBody_Conteudo_txtPesquisa"]').send_keys(row["MATRICULA"])





