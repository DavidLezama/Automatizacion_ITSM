import os
import time 
import keyring
from cryptography.fernet import Fernet
from dotenv import load_dotenv
import tkinter as tk
from tkinter import simpledialog, messagebox
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
from requests.exceptions import ConnectionError

#INSTALADORES
#pip install msedge-selenium-tools
#pip install webdriver-manager


from subprocess import CREATE_NO_WINDOW

def solicitar_credenciales():
    
    root = tk.Tk()
    root.withdraw()  

    cuenta = simpledialog.askstring("Cuenta", "Ingrese su cuenta:")
    if not cuenta:
        messagebox.showerror("Error", "Debe ingresar una cuenta.")
        return None, None

    contrasena = simpledialog.askstring("Contraseña", "Ingrese su contraseña:", show="*")
    if not contrasena:
        messagebox.showerror("Error", "Debe ingresar una contraseña.")
        return None, None

    return cuenta, contrasena


def guardar_clave_en_credenciales(nombre, valor):
    
    try:
        keyring.set_password("my_application", nombre, valor) 
        print(f"Clave {nombre} guardada/actualizada en el Windows Credential Manager.")
    except Exception as e:
        print(f"Error al guardar la clave en el Credential Manager: {e}")

def obtener_clave_desde_credenciales(nombre):
    
    try:
        valor = keyring.get_password("my_application", nombre)
        if valor:
            return valor
        else:
            print(f"No se encontró la clave {nombre} en el Credential Manager.")
            return None
    except Exception as e:
        print(f"Error al obtener la clave desde el Credential Manager: {e}")
        return None


def generar_clave():
    nueva_clave = Fernet.generate_key().decode()  
    guardar_clave_en_credenciales("ENCRYPTION_KEY", nueva_clave)  
    return nueva_clave

def guardar_credenciales(cuenta, contrasena):
   
    if not cuenta or not contrasena:
        print("No se ingresaron credenciales válidas.")
        return

    clave = generar_clave()
    cipher = Fernet(clave.encode())

    cuenta_encriptada = cipher.encrypt(cuenta.encode()).decode()
    contrasena_encriptada = cipher.encrypt(contrasena.encode()).decode()

    
    with open("Input/.env", "w") as archivo:
        archivo.write(f"CUENTA={cuenta_encriptada}\n")
        archivo.write(f"CONTRASENA={contrasena_encriptada}\n")

    print("Credenciales encriptadas y guardadas correctamente en .env.")
    print(f"Clave actualizada en el Credential Manager: ENCRYPTION_KEY={clave}")

def cargar_credenciales():
    
    load_dotenv('Input/.env')

    
    clave = obtener_clave_desde_credenciales("ENCRYPTION_KEY")
    if not clave:
        print("Error: La clave de desencriptado no está configurada en el Windows Credential Manager.")
        return

    cipher = Fernet(clave.encode())

    cuenta_encriptada = os.getenv("CUENTA")
    contrasena_encriptada = os.getenv("CONTRASENA")

    if not cuenta_encriptada or not contrasena_encriptada:
        print("Error: No se encontraron credenciales en el archivo .env.")
        return

    
    cuenta = cipher.decrypt(cuenta_encriptada.encode()).decode()
    contrasena = cipher.decrypt(contrasena_encriptada.encode()).decode()

    #print("Credenciales descifradas:")
    #print(f"Cuenta: {cuenta}")
    #print(f"Contraseña: {contrasena}")
    #print(f"Clave: {clave}")

    return cuenta, contrasena
def  configuracion_navegador(cwd):
    options = Options()
    prefs = {
        "download.default_directory": cwd,  
        "safebrowsing.enabled": "false",    
        "profile.default_content_settings.popups": 0,
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--disable-breakpad") 
    options.add_argument("disable-infobars")    
    options.add_argument("--disable-sync")
    options.add_argument("--log-level=3")  
    options.add_argument("--inprivate")


    
    #options.add_argument("--headless") #Descomentar codigo al final 🔥
    return options
def instancia_webdriver_edge(options, url):
    try:
        service = Service(EdgeChromiumDriverManager().install())
        service.creationflags = getattr(os, "CREATE_NO_WINDOW", 0)
        driver = webdriver.Edge(service=service, options=options)
        #driver.minimize_window()  # Minimiza la ventana al iniciar
        driver.get(url)   
        
        return driver
    except ConnectionError:
        messagebox.showerror("Error", "No se pudo instalar el driver!\nRevisa tu conexión ⚠️")
        return None
        
    

def autenticacion_itsm(driver,cuenta,contrasena):
    print('Prueba...')
    try:
        if(driver.find_element(By.XPATH,'//*[@id="main-message"]/h1/span')):
            messagebox.showerror('No estas conectado a internet!\n Revise su conexion a internet⚠️')
            return False
    except NoSuchElementException:
        pass
    try:
        wait = WebDriverWait(driver ,20)
        print('Entre')
        btn_auten = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="microsoft-auth-button"]')))
        btn_auten.click()
        print('Entre2')
        input_cuenta= wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="i0116"]')))
        input_cuenta.send_keys(cuenta)
        btn_siguiente = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="idSIButton9"]')))
        btn_siguiente.click()
        input_contra = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="i0118"]')))
        input_contra.send_keys(contrasena)
        
        btn_iniciar = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="idSIButton9"]')))
        btn_iniciar.click()
        
        btn_no_iniciada = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="idBtn_Back"]')))
        btn_no_iniciada.click()
        print('Aqui')
        
        try:
            print('todo good')
            btn_continuar = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="login-submit"]')))
            btn_continuar.click()
            time.sleep(20)
            print('mini siesta terminada')
            return driver
        except NoSuchElementException:
            print('No encontre la monda esa😠')
            pass

        return driver
    except TimeoutException as e:
        print(f"Error: Timeout esperando un elemento. Detalles: {str(e)}")
        return False
def navegacion_itsm(driver):
    print('a')
    url_proyecto = 'https://servicios-it-corfi.atlassian.net/jira/servicedesk/projects/MS/queues/custom/50'
    wait = WebDriverWait(driver,10)
    driver.get(url_proyecto)
    time.sleep(30)
    url_filtros = 'https://servicios-it-corfi.atlassian.net/jira/filters'
    driver.get(url_filtros)
    input_filtro = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ak-main-content"]/div/div/div[1]/div[1]/div[2]/div/div[1]/div/div/input')))
    input_filtro.send_keys('Todos ABIERTOS NIVEL 2')
    time.sleep(20)
    click_filtro = wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ak-main-content"]/div/div/div[1]/div[2]/div[2]/table/tbody/tr/td[2]/div/div/a')))
    click_filtro.click()
    time.sleep(100)
    print('termine')
    return driver

        
def main ():   
    if not os.path.exists("Input/.env"):
        cuenta, contrasena = solicitar_credenciales()
        if cuenta and contrasena:
            guardar_credenciales(cuenta, contrasena)
    else:
        cuenta , contrasena = cargar_credenciales()
        print('Nice')
    
    cwd = str(Path().resolve())
    cwd=cwd+"\\Input"
    
    options = configuracion_navegador(cwd=cwd)
    url_jira = "https://id.atlassian.com/login"
    
    driver = instancia_webdriver_edge(options=options,url= url_jira)
    print(f'Driver: {driver}')
    driver = autenticacion_itsm(driver=driver,cuenta=cuenta,contrasena=contrasena)
    navegacion_itsm(driver=driver)


main()