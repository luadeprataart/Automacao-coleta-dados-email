"""
Nome do Arquivo: app_email.py
Versão: V1.0
Python 3.11.4
geckodriver-v0.33.0-win64
Autor: Ana Julia Moraes
Data: 23 de Agosto de 2023

Descrição: Este arquivo apresenta um script de estudo para automação de tarefas, onde ele realiza interações automatizadas com o serviço
de e-mail Outlook por meio do navegador Firefox. O foco é verificar a chegada de novos e-mails (inscrições de candidatos, para mais detalhes
verifique D:\Projetos\integracao_whats-email\docs\emailTemplate.txt) de um remetente específico na caixa de entrada.

O script foi desenvolvido para cumprir as seguintes etapas:

 - Abertura do e-mail no navegador: O script abre a interface de e-mail no navegador Firefox e faz o login.
 - Verificação de novos e-mails: Ele verifica a caixa de entrada em busca de novos e-mails que atendam aos filtros estabelecidos.
 - Coleta do conteúdo do e-mail: Ao identificar um novo e-mail relevante, o script extrai o conteúdo do corpo do e-mail para análise.
 - Atualização do status do e-mail para "lido": Após coletar as informações necessárias, o script atualiza o status do e-mail para 
 "lido" na caixa de entrada.
 - Envio de mensagem via WhatsApp: Além disso, o script utiliza as informações obtidas no e-mail para obter um número de contato e envia 
 automaticamente uma mensagem para este número via WhatsappWeb no navegador padrão do computador.

Esse script demonstrativo foi criado como uma oportunidade de aprendizado para entender como a automação pode ser aplicada a tarefas 
rotineiras, como o monitoramento de e-mails e ações correspondentes. É importante observar que este script é apenas um exercício 
didático e pode ser expandido e ajustado para necessidades específicas.

"""

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from bs4 import BeautifulSoup as bs
import requests
from whats import _send_msg_whats
import config

driver = webdriver.Firefox() #Chamo o driver do browser que vou utilizar para acessar o email
driver.get("https://outlook.office.com/mail")
driver.implicitly_wait(5) 

#login do campo email no navegador
elem = driver.find_element(By.XPATH, '//*[@id="i0116"]')
elem.clear()
elem.send_keys(config.email)
elem.send_keys(Keys.ENTER)
driver.implicitly_wait(1) 

#login da senha no navegador
elem = driver.find_element(By.XPATH, '//*[@id="i0118"]')
elem.clear()
elem.send_keys(config.pass_email)
driver.implicitly_wait(5)
WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "lightbox-cover")))
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]'))).click()

WebDriverWait(driver, 30).until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "lightbox-cover")))
WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]'))).click()

#loop de busca de novos emails
while True:
    driver.implicitly_wait(30) 

    #busca o remetente 
    WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="topSearchInput"]'))).click()
    driver.implicitly_wait(30) # seconds
    elem = driver.find_element(By.XPATH, '//*[@id="topSearchInput"]')
    elem.clear()
    elem.send_keys(config.Search_email, Keys.ENTER )
    driver.implicitly_wait(10) # seconds

    #seleciona apenas novos emails
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div[1]/div/div/div[2]/div/div[3]/div/div[2]/div/div/button[2]/span/span'))).click()
    driver.implicitly_wait(20) # seconds

    #verifica se existem novos emails, caso contrario interrompe o while e finaliza a aplicação
    try:
        driver.implicitly_wait(5) # seconds
        WebDriverWait(driver, 40).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div/div[2]/div[2]/div/div/div/div[3]/div/div[2]/div[1]/div[1]/div/div/div/div/div/div/div/div[2]/div/div[2]/div/div/div[2]/div[1]/div'))).click()
        print("Encontrei um email")         
    except TimeoutException: 
        print("Sem novos candidatos encontrados, finalizando o programa!")
        break


    driver.implicitly_wait(20) # seconds

    #pega o corpo do email, caso não consiga ele reinicia a página
    try:
        html_content = driver.find_element(By.XPATH, '//*[@id="ReadingPaneContainerId"]/div/div/div/div[2]/div/div/div[1]/div/div/div/div/div[3]').get_attribute("innerHTML")
       
        print(html_content)
        driver.implicitly_wait(20) # seconds

        #marca o email como lido
        WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div/div[2]/div/div[2]/div[1]/div/div[2]/div[1]/div/div/div[1]/div/div/div/div/div/div/div/div/div[3]/div/div/div/div[2]/div'))).click() 
        
        conteudo = html_content.split('<br aria-hidden="true">') #converte html para poder trabalhar com os dados

        ##print(conteudo)

        #vamos utilizar apenas estes dados como modelo de teste
        nome = conteudo[2].split('Nome: ')
        ra = conteudo[4].split('RA: ')
        tel = conteudo[6].split('Celular: ')

        print(nome, '-', ra, '-',tel)

        if (len(nome) >= 2):
            nome = nome[1]
            ra = ra[1]
            tel = tel[1]
        else:
            nome = nome[0]
            ra = ra[0]
            tel = tel[0] 

        print(nome, '-', ra, '-',tel)
        
        #tratando o telefone para o padrão necessario para o whats
        telefone = '+55' + tel.replace('(','').replace(')','').replace('-','').replace(' ','')

        #criando o texto que será enviado para o whats do candidato 
        conteudo = f'''
        Olá {nome}, tudo bem?

        Obrigado por se inscrever no Processo Seletivo, abaixo segue o link do Manual do Candidato e convite para o grupo do WhatsApp

        manual: {config.link}
        link do grupo: {config.grupo}
        
        Em caso de dúvidas me mande uma mensagem wa.me/{config.emergency_contact}
        ''' 

        #chama a funçao de envio no whats
        _send_msg_whats(telefone, msg=conteudo)

    except NoSuchElementException:
        print('ERRO: não consegui ler o email, reiniciando página')

    
    #reinicio a página
    driver.get("https://outlook.office.com/mail")
    driver.implicitly_wait(70) 



#Fecha o browser após finalizar o serviço  
driver.close()
