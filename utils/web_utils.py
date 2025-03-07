import json
import os
import time
import traceback
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

def descargar_verificacion(data, carpeta_contenedora):
    """
    Descarga la verificación del SAT para una factura.
    
    Args:
        data (dict): Datos de la factura
        carpeta_contenedora (str): Carpeta donde guardar la verificación
        
    Returns:
        str: Ruta al archivo de verificación descargado o None en caso de error
    """
    # Configurar opciones de Chrome
    chrome_options = Options()
    
    # Nombre del archivo a generar
    nombre_archivo = f"verificacion_{data['Serie']}{data['Numero']}.pdf"
    
    # Asegurarse que el nombre del archivo termine en .pdf
    if not nombre_archivo.lower().endswith('.pdf'):
        nombre_archivo += '.pdf'
    
    # Configurar las preferencias de impresión para guardar como PDF
    settings = {
        "recentDestinations": [{
            "id": "Save as PDF",
            "origin": "local",
            "account": ""
        }],
        "selectedDestinationId": "Save as PDF",
        "version": 2,
        "isLandscapeEnabled": False,  # False para orientación vertical
        "isHeaderFooterEnabled": False,  # False para no incluir encabezado/pie de página
    }
    
    # Configurar las preferencias del navegador
    prefs = {
        'printing.print_preview_sticky_settings.appState': json.dumps(settings),
        'savefile.default_directory': carpeta_contenedora,
        'download.default_directory': carpeta_contenedora,
        'download.prompt_for_download': False,
        'download.directory_upgrade': True,
        'safebrowsing.enabled': True,
        'download.default_filename': nombre_archivo
    }
    
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument("--kiosk-printing")  # Impresión automática sin diálogo
    
    # Agregar modo headless para entornos sin interfaz gráfica
    #chrome_options.add_argument("--headless")  # Descomenta esta línea si quieres modo headless
    
    driver = None
    try:
        # Inicializar Chrome con WebDriver Manager
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )
        
        # Configurar tiempo de espera largo para cargar elementos
        wait = WebDriverWait(driver, 120)  # 2 minutos de espera máxima
        
        # Navegar a la página del SAT
        url = "https://verificacfdi.facturaelectronica.sat.gob.mx/"
        driver.get(url)
        
        # Llenar el formulario con los datos de la factura
        campos = {
            "ctl00_MainContent_TxtUUID": data['Folio_Fiscal'],
            "ctl00_MainContent_TxtRfcEmisor": data['Rfc_emisor'],
            "ctl00_MainContent_TxtRfcReceptor": data['Rfc_receptor'],
        }
        
        # Llenar cada campo del formulario
        for campo_id, valor in campos.items():
            try:
                elemento = wait.until(EC.presence_of_element_located((By.ID, campo_id)))
                elemento.clear()
                elemento.send_keys(Keys.HOME)
                elemento.send_keys(valor)
            except Exception as e:
                print(f"Error: No se encontró el elemento con ID {campo_id}")
                print(driver.page_source)  # Imprimir el HTML para depuración
                raise e
        
        # Posicionar en el campo del CAPTCHA para que el usuario lo resuelva
        elemento = wait.until(EC.presence_of_element_located((By.ID, "ctl00_MainContent_TxtCaptchaNumbers")))
        elemento.send_keys(Keys.HOME)
        
        # Esperar a que el botón de imprimir esté visible
        btn_imprimir = wait.until(EC.visibility_of_element_located((By.ID, "BtnImprimir")))
        
        # Hacer clic en el botón de imprimir
        btn_imprimir.click()
        time.sleep(2)  # Esperar a que se inicie la descarga
        
        # Obtener el nombre del archivo descargado (el más reciente en la carpeta)
        time.sleep(5)  # Dar tiempo para que se complete la descarga
        archivos_descargados = os.listdir(carpeta_contenedora)
        
        if not archivos_descargados:
            return None
            
        # Ordenar archivos por fecha de modificación (el más reciente primero)
        archivos_descargados.sort(
            key=lambda x: os.path.getmtime(os.path.join(carpeta_contenedora, x)), 
            reverse=True
        )
        
        archivo_descargado = archivos_descargados[0]
        return os.path.join(carpeta_contenedora, archivo_descargado)
        
    except Exception as e:
        print(f"Error durante la verificación: {e}")
        traceback.print_exc()  # Imprimir el stacktrace completo
        return None
        
    finally:
        # Cerrar el navegador
        if driver:
            driver.quit()
