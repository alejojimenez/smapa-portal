import os
import time
import shutil
import requests
import re

import pyautogui


from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys

class Scraper_Smapa():

    def __init__(self,url, email, password, driver_path):
        print(url, email, password, driver_path)
        self.url = url
        self.email = email
        self.password = password
        self.driver_path = driver_path

    def wait(self, seconds):
        return WebDriverWait(self.driver, seconds)

    def close(self):
        self.driver.close()
        self.driver = None

    def quit(self):
        self.driver.quit()
        self.driver = None    

    def login(self):
        
        driver_exe = '.domain\\chromedriver.exe'
        credencials = '.\\config\\credenciales.xlsx'

        print('Entrando en la funcion login...')
        print('----------------------------------------------------------------------')
        
        #Seteo variables
        email = self.email
        url = self.url
        driver_path = self.driver_path
        password  = self.password
        
        options = webdriver.ChromeOptions()
        options.add_experimental_option('prefs', {
        "download.default_directory": "C:\\roda\\smapa-portal\\input\\", #Change default directory for downloads
        "download.prompt_for_download": False, #To auto download the file
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True, #It will not show PDF directly in chrome
        })  
        
        self.driver = webdriver.Chrome(driver_path, options=options)
        self.driver.get(url)
        self.driver.maximize_window()
        self.driver.implicitly_wait(30)

        # Controlar evento alert() Notificacion
        try:
            alert = WebDriverWait(self.driver, 35).until(EC.alert_is_present())
            alert = Alert(self.driver)
            alert.accept()
                    
        except:
            print("No se encontró ninguna alerta.")  

        #Botones ID que utilizaremos para logear
        selector_ingreso_cuenta = 'Mail'
        selector_password_input = 'Password'
        selector_login_button = 'btn-primary'

        # Seleccionar campo cuenta
        intentos = 0
        reintentar = True
        while (reintentar):
            try:
                print('Try en la funcion login para el campo cuenta...', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                element_cuenta= self.driver.find_element(By.NAME, selector_ingreso_cuenta)
                element_cuenta.clear()
                element_cuenta.click()
                element_cuenta.send_keys(email)
                reintentar = False
            except:    
                print('Exception en la funcion click campo cuenta')
                print('----------------------------------------------------------------------')
                reintentar = intentos <= 3                
                
        # Seleccionar campo clave y setear clave
        intentos = 0
        reintentar = True
        while (reintentar):
            try:
                print('Try en la funcion click para el campo rut...', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                element_password = self.driver.find_element(By.NAME, selector_password_input)
                element_password.clear()
                element_password.click()
                element_password.send_keys(password)
                reintentar = False
            except:    
                print('Exception en la funcion click rut cliente')
                print('----------------------------------------------------------------------')
                reintentar = intentos <= 3

        # Seleccionar boton y hacer en boton ingresar
        intentos = 0
        reintentar = True
        while (reintentar):
            try:
                print('Try en la funcion click_element_xpath para el boton ingresar...', intentos)
                print('----------------------------------------------------------------------')
                intentos += 1
                element_button_ingresar = self.driver.find_element(By.CLASS_NAME, selector_login_button)
                element_button_ingresar.click()
                reintentar = False
            except:    
                print('Exception en la funcion click_element_xpath')
                print('----------------------------------------------------------------------')
                reintentar = intentos <= 3
    
    def scrapping_aguas(self,sociedad,limite):


        while sociedad < limite:
        
            #Buscamos la tabla que contiene las filas a buscar
            intento_sociedades = 0
            reintentar_sociedades = True
            while (reintentar_sociedades):
                try:
                    time.sleep(5)
                    elemento = self.driver.find_element(By.XPATH,f"/html/body/div[1]/div[2]/main/div/div/div/div/div/div[2]/div/div/div[2]/div/div[{sociedad}]/div[1]/a")
                    print(f'la sociedad es la numero {sociedad}')
                    numero_sociedad = elemento.text
                    elemento.click()
                    reintentar_sociedades = False
                except:    
                    print('menu sociedades no esta disponible')
                    print('----------------------------------------------------------------------')
                    self.driver.execute_script("window.scrollBy(0, 700);")
                    reintentar_sociedades = intento_sociedades <= 5 


            #Aqui, una vez que ya entramos a la sociedad, obtenemos la cantidad de boletas para comenzar descarga. Sera un bucle separado
            filas_tabla = WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located((By.TAG_NAME,"tr"))) 
            
            #Buscamos el largo de las descargas para utilizarlo como limite
            filas = self.driver.find_elements(By.TAG_NAME,"tr")
            
            cantidad_facturas =len(filas)
            print(f'cantidad de facturas es{cantidad_facturas}')
            print('----------------------------------------------------------')
            
            i=1
            while i < cantidad_facturas:

                #extraer el numero de factura
                intento = 0
                factura = True
                while (factura):
                    try:
                        time.sleep(5)
                        boton_factura = self.driver.find_element(By.XPATH,f'/html/body/div[1]/div[2]/main/div/div/div/div/div/div[3]/div/div[2]/div/div[2]/div/div/div/table/tbody/tr[{i}]/td[1]')
                        numero_fact = boton_factura.text
                        print('Factura_text: ', numero_fact)
                        lista = numero_fact.split()
                        factura_oficial = lista[2]
                        print(factura_oficial)

                        time.sleep(5)
                        boton_periodo = self.driver.find_element(By.XPATH,f'/html/body/div[1]/div[2]/main/div/div/div/div/div/div[3]/div/div[2]/div/div[2]/div/div/div/table/tbody/tr[{i}]/td[2]')
                        numero_period = boton_periodo.text
                        print('Period_text: ', numero_period)
                        lista_period = numero_period.split()
                        periodo_oficial = lista_period[1]
                        print(periodo_oficial)
                        
                        factura = False                        
                    except:    
                        print('Boton de descarga no se encuentra disponible....')
                        print('----------------------------------------------------------------------')
                        factura = intento <= 5    
                
                #Buscamos el boton de descarga en cada una de las filas
                intento = 0
                descarga = True
                while (descarga):
                    try:
                        time.sleep(5)
                        boton_descarga = self.driver.find_element(By.XPATH,f'/html/body/div[1]/div[2]/main/div/div/div/div/div/div[3]/div/div[2]/div/div[2]/div/div/div/table/tbody/tr[{i}]/td[6]/div/button')  
                        boton_descarga.click()
                        descarga = False
                    except:    
                        print('Boton de descarga no se encuentra disponible....')
                        print('----------------------------------------------------------------------')
                        descarga = intento <= 5      

                # Esperar a que se abra la ventana emergente
                intentos = 0
                reintentar_ventana = True
                while (reintentar_ventana):
                    try:
                        # Obtiene el identificador de la ventana actual
                        current_window = self.driver.current_window_handle
                        print('Ventana principal: ', current_window)
                        print('----------------------------------------------------------------------')
                        print('Try en la funcion manejo de ventanas abiertas...', intentos)
                        print('----------------------------------------------------------------------')
                        intentos += 1
                        window_handles_all = self.driver.window_handles
                        reintentar_ventana = False
                            
                    except:    
                        print('Exception en la funcion manejo de ventanas abiertas')
                        print('----------------------------------------------------------------------')
                        print('----------------------------------------------------------------------')
                        time.sleep(60) #espera para que cargue ventana emergente
                                    
                        reintentar_ventana = intentos <= 3
                                    
                # Obtiene los identificadores de las ventanas abiertas
                self.driver.implicitly_wait(35)
                print('Ventanas abiertas: ', window_handles_all, len(window_handles_all))
                print('----------------------------------------------------------------------')            
                                    
                # Cambiar al manejo de la ventana emergente
                for window_handle in window_handles_all:
                    if window_handle != current_window:
                        self.driver.switch_to.window(window_handle)
                        print('Ventana emergente: ', window_handle)
                        print('----------------------------------------------------------------------')
                        break
                
                
                #Revisar si este es el primer documento o no
                #print(os.listdir(folder_path))
                if i ==1 and sociedad==1:
                    hay_archivos = False
                elif i !=1 and sociedad ==1:
                    hay_archivos = True
                elif i ==1 and sociedad !=1:
                    hay_archivos = True
                elif i !=1 and sociedad !=1:
                    hay_archivos = True
                
                # Esperar hasta que el elemento esté presente en la página para descargar
                time.sleep(20) 
                pyautogui.click(500, 500)
                
                #AMBAS OPCIONES FUNCIONAN
                actions = ActionChains(self.driver)
                actions.send_keys(Keys.RETURN).perform()

                # Cerrar la ventana emergente
                time.sleep(5)
                self.driver.close()
                
                print(f'la pasada numero {i} entrega hay archivos igual a {hay_archivos}')
                
                #Ruta de descarga
                folder_path = './input/'
                
                if hay_archivos == False:
                
                    #Buscamos el archivo que se haya descargado que comience con el nombre boleta para poder moverlo a la carpeta input
                    for filename in os.listdir(folder_path):
                        if filename.endswith('.pdf'):
                            nombre_archivo = filename
                            os.rename(folder_path+nombre_archivo,folder_path+f'soc_{numero_sociedad}_fac_{factura_oficial}_{periodo_oficial}.pdf')
                
                elif hay_archivos == True:
                    
                    for filename in os.listdir(folder_path):
                        if not filename.startswith('soc_') and filename.endswith('.pdf'):
                            nombre_archivo = filename
                            os.rename(folder_path+nombre_archivo,folder_path+f'soc_{numero_sociedad}_fac_{factura_oficial}_{periodo_oficial}.pdf')   

                # Cambiar de nuevo al manejo de ventana principal
                self.driver.switch_to.window(current_window)
                print('Cual ventana es: ', current_window)
                print('----------------------------------------------------------------------')
                        
                time.sleep(10)
                
                print('pasamos al siguiente archivo')
                i+=1
            
            sociedad+=1  
            self.driver.execute_script("window.history.go(-1)")
            print('Pasamos a la siguiente sociedad')
