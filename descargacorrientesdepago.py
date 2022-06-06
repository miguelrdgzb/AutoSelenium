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


def ExtractCorrientedePago(CuentadeCotizacion):  
    CCC = str(CuentadeCotizacion)
    regimen = '0111'
    chromeOptions = webdriver.ChromeOptions()
    driver_path = 'chromedriver.exe'
    driver = webdriver.Chrome(driver_path, options=chromeOptions)
    driver.get('https://w2.seg-social.es/Xhtml?JacadaApplicationName=SGIRED&TRANSACCION=RCR92&E=I&AP=DEUR')
    time.sleep(0.5)
    pyautogui.press('enter')
    driver.find_element_by_id('SDFWMIDENT').send_keys(CCC)
    driver.find_element_by_id('SDFWMRESU').send_keys(regimen)
    driver.find_element_by_id('detalle_deuda').click()
    ListaDesplegable = driver.find_element_by_id('ListaTipoImpresion')
    ElegirOpcion = Select(ListaDesplegable)
    ElegirOpcion.select_by_visible_text('OnLine')
    time.sleep(1)

    ContinueButton = driver.find_element_by_id('Sub2207101004_65')
    ContinueButton.click()
    time.sleep(1)
    NextButton = driver.find_element_by_id('Sub2204801005_7')
    NextButton.click()
    time.sleep(3)

    p = driver.current_window_handle
    parent = driver.window_handles[0]
    chld = driver.window_handles[1]
    driver.switch_to.window(chld)
    time.sleep(5)
    pyautogui.hotkey('ctrl','s')
    time.sleep(2)
    for i in CCC:
        pyautogui.press(i)
    pyautogui.press('enter')
    time.sleep(2)
    driver.close()

ObjetcToRead = pd.read_csv(r'CENTROS Y CCC.csv', sep=';')
ObjetcToRead.CCC = ObjetcToRead.CCC.apply(lambda x: x.replace(' ',''))
lista_ccc = ObjetcToRead.CCC.values

for i in range(len(ObjetcToRead)):
    ExtractCorrientedePago(lista_ccc[i])
    time.sleep(2)