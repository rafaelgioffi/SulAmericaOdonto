import configparser
import os.path
import pathlib
import subprocess
from datetime import datetime, date, timedelta

import pyautogui
import pymsgbox
from selenium import webdriver
from selenium.webdriver.common.by import By
from win32api import GetSystemMetrics
from selenium.webdriver.chrome.service import Service
#from winotify import Notification

cfg = configparser.ConfigParser()
cfg.read('CONFIG.ini')
empresa = cfg.getint('config', 'empresa')
codigo = cfg.getint('config', 'codigo')
chromedriver = cfg.get('config', 'driver')

data = ''
hoje = date.today()
# mes = hoje.month
dia_semana = hoje.isoweekday()
# data = datetime(data)
resp = ''

if dia_semana == 6:
    data = date.today() + timedelta(days=2)
    # print(data)
elif dia_semana == 7:
    data = date.today() + timedelta(days=1)
elif not dia_semana == 6 or not dia_semana == 7:
    data = date.today() + timedelta(days=1)

data = data.strftime(f'%d/%m/%Y')
# print(data)

inicial = datetime.now()
inicial = inicial.strftime('%H:%M:%S')
# print(inicial, type(inicial))
final = datetime.now() + timedelta(hours=+2)
final = final.strftime('%H:%M:%S')
# print(inicial)
# print(final)


def Acessar():
    navegador = webdriver.Chrome()#s)
    ##n = Notification(app_id='autogerador', title='Aguarde...', msg='Abriu o Chrome...')
    ##n.build().show()
    print('Abriu o Chrome...')

    navegador.get('https://portal.sulamericaseguros.com.br/pagamentos/#/odonto')
    ##n = Notification(app_id='autogerador', title='Aguarde...', msg='Abriu o site...')
    ##n.build().show()
    print('Abriu o site...')
    navegador.set_window_size(width=GetSystemMetrics(0) / 1.3, height=GetSystemMetrics(1) / 1.05)
    print('Alterou o tamanho da janela...')
    ##n = Notification(app_id='autogerador', title='Aguarde...', msg='Alterou o tamanho da janela...')
    ##n.build().show()
    # navegador.maximize_window()

    try:
        if navegador.find_element(By.XPATH, '//*[@id="sas-box-lgpd-info"]/div/div[2]/button').click():
            print('Clicou em Aceitar Cookies...')
            ##n = Notification(app_id='autogerador', title='Aguarde...', msg='Clicou em Aceitar Cookies...')
            ##n.build().show()
        else:
            print('Falha ao clicar em "Aceitar Cookies"')
            ##n = Notification(app_id='autogerador', title='Aguarde...', msg='Falha ao clicar em "Aceitar Cookies"')
            ##n.build().show()
    except:
        pass
        pyautogui.sleep(1)
    try:
        if navegador.find_element(By.XPATH, '/html/body/app-root/app-home/div/div/div[2]/app-odonto/div/div/div/button[2]').click():
            print('Clicou em Pessoa Jurídica...')
            ##n = Notification(app_id='autogerador', title='Aguarde...', msg='Clicou em Pessoa Jurídica...')
            ##n.build().show()
        else:
            print('Falha ao clicar em Pessoa Jurídica...')
            ##n = Notification(app_id='autogerador', title='Aguarde...', msg='Falha ao clicar em Pessoa Jurídica...')
            ##n.build().show()
    except:
        pass
    try:
        if navegador.find_element(By.XPATH, '//*[@id="cnpj"]').send_keys(empresa):
            print('Inseriu o CNPJ...')
            ##n = Notification(app_id='autogerador', title='Aguarde...', msg='Inseriu o CNPJ...')
            ##n.build().show()
        else:
            print('Falha ao inserir o CNPJ...')
            ##n = Notification(app_id='autogerador', title='Aguarde...', msg='Falha ao inserir o CNPJ...')
            ##n.build().show()
    except:
        pass
    try:
        if navegador.find_element(By.XPATH, '//*[@id="codEmpresa"]').send_keys(codigo):
            print('Inseriu o código...')
            #n = Notification(app_id='autogerador', title='Aguarde...', msg='Inseriu o código...')
            #n.build().show()
        else:
            print('Falha ao inserir o código...')
            #n = Notification(app_id='autogerador', title='Aguarde...', msg='Falha ao inserir o código...')
            #n.build().show()
    except:
        pass

    try:
        if navegador.find_element(By.XPATH, '/html/body/app-root/app-home/div/div/div[2]/app-odonto/div/div/form/div[3]/button').click():
            print('Clicou em Pesquisar...')
            #n = Notification(app_id='autogerador', title='Aguarde...', msg='Clicou em Pesquisar...')
            #n.build().show()
        else:
            print('Falha ao clicar em Pesquisar...')
            #n = Notification(app_id='autogerador', title='Aguarde...', msg='Clicou em Pesquisar...')
            #n.build().show()
    except:
        pass

    print(f'Aguarde 6 segundos...')
    #n = Notification(app_id='autogerador', title='Aguarde...', msg=f'Aguarde 6 segundos...')
    #n.build().show()
    pyautogui.sleep(6)

    try:
        navegador.find_element(By.XPATH, '//*[@id="rc-imageselect"]/div[2]')
        print('Captcha detectado.')
        #n = Notification(app_id='autogerador', title='Aguarde...', msg='Captcha detectado.')
        #n.build().show()
        pymsgbox.alert(f'Resolva o Captcha e clique em OK para continuar...','Aguardando o Captcha')
    except:
        print('Captcha não detectado. Prosseguindo...')
        #n = Notification(app_id='autogerador', title='Aguarde...', msg='Captcha não detectado. Prosseguindo...')
        #n.build().show()

    try:
        navegador.find_element(By.XPATH, '/html/body/app-root/app-home/div/div/div[2]/app-odonto/app-busca-odonto-pj/div/div/div/div[2]/div[2]/div[3]/div[5]/div/button[1]').click()
        print('Clicou em "Cód. de Barras"...')
        #n = Notification(app_id='autogerador', title='Aguarde...', msg='Clicou em "Cód. de Barras"...')
        #n.build().show()
    except:
        print('Falha ao clicar em "Cód de Barras"...')
        #n = Notification(app_id='autogerador', title='Aguarde...', msg='Falha ao clicar em "Cód de Barras"...')
        #n.build().show()

    print(f'Aguarde mais 3 segundos...')
    #n = Notification(app_id='autogerador', title='Aguarde...', msg=f'Aguarde mais 3 segundos...')
    #n.build().show()
    pyautogui.sleep(3)

    try:
        print('Boleto vencido')
        #n = Notification(app_id='autogerador', title='Aguarde...', msg=f'Boleto vencido')
        #n.build().show()

        # clicar no calendario...
        navegador.find_element(By.XPATH, '//*[@id="mat-dialog-0"]/app-modal-data/div/div/form/div[1]/mat-datepicker-toggle/button').click()
        pyautogui.sleep(1)
        if pyautogui.locateOnScreen('img/border.png'):
            border = pyautogui.locateOnScreen('img/border.png')
            pyautogui.click(border)
            # pyautogui.sleep(2)

        # navegador.find_element(By.XPATH, '//*[@id="date"]').send_keys(data)
        #n = Notification(app_id='autogerador', title='Aguarde...', msg=f'Data de vencimento inserida')
        #n.build().show()
        # pyautogui.sleep(2)
        navegador.find_element(By.XPATH, '//*[@id="mat-dialog-0"]/app-modal-data/div/div/form/div[2]/button[1]').click()
        #n = Notification(app_id='autogerador', title='Aguarde...', msg=f'Data definida')
        #n.build().show()
    except:
        print('Boleto não está vencido')
        #n = Notification(app_id='autogerador', title='Aguarde...', msg=f'Boleto não está vencido')
        #n.build().show()

    try:
        if navegador.find_element(By.XPATH, '//*[@id="mat-dialog-0"]/app-modal-boleto/div'):
            print('Boleto gerado')
            #n = Notification(app_id='autogerador', title='Aguarde...', msg='Boleto gerado')
            #n.build().show()
        else:
            print('Falha ao gerar o boleto...')
            #n = Notification(app_id='autogerador', title='Aguarde...', msg='Falha ao gerar o boleto...')
            #n.build().show()
    except:
        pass

    print(f'Aguarde mais 4 segundos...')
    #n = Notification(app_id='autogerador', title='Aguarde...', msg=f'Aguarde 4 segundos...')
    #n.build().show()
    pyautogui.sleep(4)

    try:
        if navegador.find_element(By.XPATH, '//*[@id="mat-dialog-1"]/app-erro-http/div'):
            print('Erro de comunicação, tente mais tarde')
            #n = Notification(app_id='autogerador', title='Aguarde...', msg=f'erro de comunicação, tente mais tarde...')
            #n.build().show()
            navegador.find_element(By.XPATH, '//*[@id="mat-dialog-1"]/app-erro-http/div/button').click()
            navegador.close()
            return False
    except:
        pass

    try:
        resp = pymsgbox.alert(f'Realize o pagamento a partir das {final}h e aperte OK para finalizar o script...','Realize o pagamento')
        if resp == 'OK':
            print('OK')
            # navegador.close()
            navegador.find_element(By.XPATH, '//*[@id="mat-dialog-0"]/app-modal-boleto/div/div/div[3]/button[1]').click()
            navegador.find_element(By.XPATH, '/html/body/app-root/app-home/div/div[2]/div/app-odonto/app-busca-odonto-pj/div/div/h2/span').click()

    except:
        pass

if os.path.exists(r'C:\\Program Files (x86)\\Google\\Chrome\\Application\\'):
    output = subprocess.check_output(
    r'wmic datafile where name="C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" get Version /value',
    shell=True
)
    chrome_version = output.decode('utf-8').split('.')
    chrome_version = chrome_version[0].replace('\n', '')
    chrome_version = chrome_version.replace('Version=', '')
    chrome_version = int(chrome_version)
    print('Chrome de 32 bits instalado: ', chrome_version)

    if chrome_version != '':
        s = r'' + chromedriver
        Acessar()
    else:
        print('Please, install the last version of Google Chrome in your system before use this software')
        print('Por favor, instale a última versão do Google Chrome antes de usar esse software')
elif os.path.exists(r'C:\\Program Files\\Google\\Chrome\\Application\\'):
    output = subprocess.check_output(
    r'wmic datafile where name="C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" get Version /value',
    shell=True
)
    chrome_version = output.decode('utf-8').split('.')
    chrome_version = chrome_version[0].replace('\n', '')
    chrome_version = chrome_version.replace('Version=', '')
    chrome_version = int(chrome_version)
    print('Chrome de 64 bits instalado: ', chrome_version)

    if chrome_version != '':
        s = r'' + chromedriver
        Acessar()
    else:
        print('Please, install the last version of Google Chrome in your system before use this software')
        print('Por favor, instale a última versão do Google Chrome antes de usar esse software')

else:
    print('Google chrome não instalado na pasta padrão...')