from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import random 
import pandas as pd
import numpy as np
import os
from selenium.webdriver.support.ui import Select
from selenium.webdriver import DesiredCapabilities
from selenium.webdriver import ActionChains
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import pyautogui
import shutil
from threading import Thread

def ExtractIDC(NumeroSegSocial,CCC):
    driver_path = 'chromedriver.exe'
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(driver_path, options=options)


    def threaded_function(driver):
        #Calls the website
        driver = driver.get('https://w2.seg-social.es/Xhtml?JacadaApplicationName=SGIRED&TRANSACCION=ATR37&E=I&AP=AFIR')
        

    def threaded_function2():
        #Presses 10 times
        time.sleep(2)
        pyautogui.press('enter')

    thread2 = Thread(target = threaded_function2)
    thread2.start()
    thread = Thread(target = threaded_function(driver))
    thread.start()
    time.sleep(1)
    NumeroSegSocial = NumeroSegSocial.replace('/','').replace(' ','')
    CCC = CCC.replace('/','').replace(' ','')
    prefix = NumeroSegSocial[0:2]
    numero = NumeroSegSocial[2:]
    regimen = '0111'
    prefix_ccc = CCC[0:2]
    numero_ccc = CCC[2:]
    driver.find_element_by_id('SDFTESNAF').send_keys(prefix)
    time.sleep(1)
    driver.find_element_by_id('SDFNAF').send_keys(numero)
    time.sleep(1)
    driver.find_element_by_id('SDFREGCTA').send_keys(regimen)
    time.sleep(1)
    driver.find_element_by_id('SDFTESCTA').send_keys(prefix_ccc)
    time.sleep(1)
    driver.find_element_by_id('SDFCUENTA').send_keys(numero_ccc)

    ListaDesplegable = driver.find_element_by_id('ListaTipoImpresion')
    ElegirOpcion = Select(ListaDesplegable)
    ElegirOpcion.select_by_visible_text('OnLine')
    time.sleep(1)

   
    ContinueButton = driver.find_element_by_id('Sub2207601004')
    ContinueButton.click()
    time.sleep(3)
    try:
        fecha_tabla = driver.find_element_by_id('Sub0900112078_3_4')
        action = ActionChains(driver)
        action.double_click(fecha_tabla).perform()
        time.sleep(0.5)
    except: 
        pass
   
    try:
        fecha_tabla = driver.find_element_by_id('Sub0900112078_3_3')
        action = ActionChains(driver)
        action.double_click(fecha_tabla).perform()
        time.sleep(0.5)
    except:
        pass
    
    try:

        fecha_tabla = driver.find_element_by_id('Sub0900112078_3_3')
        action = ActionChains(driver)
        action.double_click(fecha_tabla).perform()
        time.sleep(0.5)
    except:
        pass


    try:
        fecha_tabla = driver.find_element_by_id('Sub0900112078_3_2')
        action = ActionChains(driver)
        action.double_click(fecha_tabla).perform()
        time.sleep(0.5)
    except:
        pass
    
    try:
        fecha_tabla = driver.find_element_by_id('Sub0900112078_3_1')
        action = ActionChains(driver)
        action.double_click(fecha_tabla).perform()
        time.sleep(0.5)
    except:
        pass
    try:    
        fecha_tabla = driver.find_element_by_id('Sub0900112078_3_0')
        action = ActionChains(driver)
        action.double_click(fecha_tabla).perform()
        time.sleep(0.5)
    except:
        pass

    time.sleep(2)
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
