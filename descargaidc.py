from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import random 
import pandas as pd
import numpy as np
from selenium.webdriver.support.ui import Select
from selenium.webdriver import DesiredCapabilities
from selenium.webdriver import ActionChains
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import shutil

def ExtractIDC(NumeroSegSocial,CCC):
    chromeOptions = webdriver.ChromeOptions()
    '''prefs = {"download.default_directory" : r'C://Users//MiguelRodríguezBubbe//MooveCars//Ficheros Generales - Compartido Gestión Nóminas//DescargaIDC//'}
    chromeOptions.add_experimental_option("prefs",prefs)'''
    NumeroSegSocial = NumeroSegSocial.replace('/','').replace(' ','')
    CCC = CCC.replace('/','').replace(' ','')
    prefix = NumeroSegSocial[0:2]
    numero = NumeroSegSocial[2:]
    regimen = '0111'
    prefix_ccc = CCC[0:2]
    numero_ccc = CCC[2:]
    driver_path = 'chromedriver.exe'
    driver = webdriver.Chrome(driver_path, options=chromeOptions)
    driver.get('https://w2.seg-social.es/M/menuAFI-REMESAS.html')
    driver.get('https://w2.seg-social.es/Xhtml?JacadaApplicationName=SGIRED&TRANSACCION=ATR37&E=I&AP=AFIR')
    driver.find_element_by_id('SDFTESNAF').send_keys(prefix)
    time.sleep(2)
    driver.find_element_by_id('SDFNAF').send_keys(numero)
    time.sleep(2)
    driver.find_element_by_id('SDFREGCTA').send_keys(regimen)
    time.sleep(2)
    driver.find_element_by_id('SDFTESCTA').send_keys(prefix_ccc)
    time.sleep(2)
    driver.find_element_by_id('SDFCUENTA').send_keys(numero_ccc)

    ListaDesplegable = driver.find_element_by_id('ListaTipoImpresion')
    ElegirOpcion = Select(ListaDesplegable)
    ElegirOpcion.select_by_visible_text('OnLine')
    time.sleep(1)

   
    ContinueButton = driver.find_element_by_id('Sub2207601004')
    ContinueButton.click()
    time.sleep(3)

    if driver.find_element_by_id('Sub0900112078_3_4') == True:
        fecha_tabla = driver.find_element_by_id('Sub0900112078_3_4')
        action = ActionChains(driver)
        action.double_click(fecha_tabla).perform()
        time.sleep(5)
    elif driver.find_element_by_id('Sub0900112078_3_3') == True:
        fecha_tabla = driver.find_element_by_id('Sub0900112078_3_3')
        action = ActionChains(driver)
        action.double_click(fecha_tabla).perform()
        time.sleep(5)
    elif driver.find_element_by_id('Sub0900112078_3_2') == True:
        fecha_tabla = driver.find_element_by_id('Sub0900112078_3_2')
        action = ActionChains(driver)
        action.double_click(fecha_tabla).perform()
        time.sleep(5)
    elif driver.find_element_by_id('Sub0900112078_3_1') == True:
        fecha_tabla = driver.find_element_by_id('Sub0900112078_3_1')
        action = ActionChains(driver)
        action.double_click(fecha_tabla).perform()
        time.sleep(5)
    else:
        fecha_tabla = driver.find_element_by_id('Sub0900112078_3_0')
        action = ActionChains(driver)
        action.double_click(fecha_tabla).perform()
        time.sleep(5)

    



    parent = driver.window_handles[0]
    chld = driver.window_handles[1]
    driver.switch_to.window(chld)
    
    time.sleep(2)
    pyautogui.hotkey('ctrl','s')
    time.sleep(2)
    numero_entero = prefix+numero
    for i in numero_entero:
        pyautogui.press(i)
    pyautogui.press('enter')

    driver.close()
    driver.switch_to.window(parent)
    

    time.sleep(2)

    driver.find_element_by_id('Sub2205001006').click()


ObjetcToRead = pd.read_excel('idc.xls', converters={'CCC': lambda x: str(x), 'NAF': lambda x: str(x)})
ObjetcToRead.CCC = ObjetcToRead.CCC.astype(str)
ObjetcToRead.NAF = ObjetcToRead.NAF.astype(str)
lista_ccc = ObjetcToRead.CCC.values
lista_numerosseg = ObjetcToRead.NAF.values

for i in range(len(ObjetcToRead)):
    ExtractIDC(lista_numerosseg[i],lista_ccc[i])
    time.sleep(5)