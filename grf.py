import io
import os
import time
import traceback
from typing import Optional
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    StaleElementReferenceException, 
    TimeoutException,
    WebDriverException
)
from selenium.webdriver.chrome.options import Options

def crear_driver_browserless() -> webdriver.Remote:
    """
    Crea un driver remoto de Chrome usando Browserless.
    Configurado específicamente para Render.
    
    Returns:
        webdriver.Remote: Instancia del driver configurado
        
    Raises:
        ValueError: Si no se encuentra el token de Browserless
        WebDriverException: Si falla la conexión con Browserless
    """
    # Validar que existe el token
    browserless_token = os.getenv("BROWSERLESS_TOKEN")
    if not browserless_token:
        raise ValueError(
            "❌ BROWSERLESS_TOKEN no está configurado. "
            "Configúralo en Render Dashboard > Environment > Add Environment Variable"
        )
    
    print(f"✅ Token de Browserless detectado: {browserless_token[:8]}...")
    
    # Opciones de Chrome optimizadas para entorno remoto
    chrome_options = Options()
    
    # Opciones esenciales para headless
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    
    # Optimizaciones de rendimiento
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-infobars")
    
    # Configuración de ventana
    chrome_options.add_argument("--window-size=1920,1080")
    
    # Anti-detección
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Manejo de certificados SSL
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--allow-running-insecure-content")
    
    # 🔹 URL correcta moderna de Browserless
    command_executor_url = f"https://production-sfo.browserless.io/webdriver?token={browserless_token}"
    
    print(f"🌐 Conectando a Browserless...")
    
    try:
        driver = webdriver.Remote(
            command_executor=command_executor_url,
            options=chrome_options
        )
        
        # Configurar timeouts
        driver.set_page_load_timeout(60)
        driver.implicitly_wait(10)
        
        print("✅ Driver de Browserless creado exitosamente")
        return driver
        
    except WebDriverException as e:
        error_msg = str(e)
        if "legacy" in error_msg.lower():
            raise WebDriverException(
                "❌ Error: Estás usando un endpoint legacy. "
                "Verifica que la URL sea: https://production-sfo.browserless.io/webdriver"
            )
        elif "unauthorized" in error_msg.lower() or "401" in error_msg:
            raise WebDriverException(
                "❌ Error de autenticación. Verifica que BROWSERLESS_TOKEN sea válido."
            )
        else:
            raise WebDriverException(f"❌ Error conectando a Browserless: {error_msg}")


def consultar_en_grf(usuario: str, contra: str, archivo: str) -> Optional[io.BytesIO]:
    """
    Automatiza consultas en el sistema GRF usando Selenium y Browserless.
    
    Args:
        usuario: Usuario para login
        contra: Contraseña para login
        archivo: Ruta del archivo Excel con datos de entrada
        
    Returns:
        BytesIO con archivo Excel de resultados, o None si hay error
    """
    driver = None
    
    try:
        print("🚀 Inicializando Browserless ChromeDriver...")
        driver = crear_driver_browserless()
        wait = WebDriverWait(driver, 30)

        print("🌐 Abriendo página GRF...")
        driver.get("https://grf.claro.com.co:8202/GIT-web/")

        print("🔐 Ingresando credenciales...")
        driver.find_element(By.NAME, "j_idt41").send_keys(usuario)
        driver.find_element(By.NAME, "j_idt43").send_keys(contra)

        print("👆 Haciendo login...")
        btn_login = wait.until(EC.element_to_be_clickable((By.NAME, "j_idt47")))
        btn_login.click()

        print("📋 Navegando a Inventario...")
        menu_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "i.fa.fa-th-large")))
        menu_icon.click()

        inventarios_link = wait.until(EC.element_to_be_clickable((By.ID, "irVistaInventarioMenu")))
        inventarios_link.click()

        btn_consul = wait.until(EC.element_to_be_clickable((By.ID, "irVistaConsultaInventario")))
        btn_consul.click()

        print("✅ Sesión iniciada correctamente")
        
        # Configuración de tipo de consulta
        stic = 2  # 1: NACIONALES, 2: BOGOTÁ
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        
        # Cargar datos según tipo
        doc_envios = pd.read_excel(archivo)
        
        if stic == 1:
            lista_seriales = (
                doc_envios["NºSerieFab"]
                .dropna()
                .astype(str)
                .str.replace(r"[\s-]+", "", regex=True)
                .tolist()
            )
            lista_sap = doc_envios["Material"]
            lista_desc = doc_envios["Textobrevedematerial"]
            lista_ciudad = doc_envios["Destino"]
        else:  # stic == 2
            lista_seriales = (
                doc_envios["Serial"]
                .dropna()
                .astype(str)
                .str.replace(r"[\s-]+", "", regex=True)
                .tolist()
            )
            lista_sap = doc_envios["Codigo material"]
            lista_desc = doc_envios["Descripción SAP"]
            lista_ciudad = doc_envios["Departamento"]

        print(f"📊 Procesando {len(lista_seriales)} seriales...")
        resultados = []

        # Procesar cada serial
        for idx, serial_number in enumerate(lista_seriales):
            print(f"\n[{idx+1}/{len(lista_seriales)}] Procesando serial: {serial_number}")

            for intento in range(2):
                try:
                    # Ingresar serial
                    campo_serial = wait.until(EC.presence_of_element_located((By.ID, "idSerialBuscar")))
                    campo_serial.clear()
                    campo_serial.send_keys(serial_number)

                    # Buscar
                    buscar_link = wait.until(EC.element_to_be_clickable((By.ID, "btnInventarioBuscar")))
                    buscar_link.click()

                    # Esperar resultado
                    WebDriverWait(driver, 30).until(
                        EC.any_of(
                            EC.presence_of_element_located((By.XPATH, "//td[contains(text(),'No hay resultados en la Base de Datos')]")),
                            EC.presence_of_element_located((By.XPATH, "//span[contains(@id,'tablaInventarioTecnicoSerial')]"))
                        )
                    )

                    # Verificar si no hay resultados
                    try:
                        driver.find_element(By.XPATH, "//td[contains(text(),'No hay resultados en la Base de Datos')]")
                        print(f"⚠️  Serial {serial_number} no encontrado en BD")
                        resultados.append({
                            "SERIAL_BUSCADO": serial_number,
                            "SAP": lista_sap[idx],
                            "DESCRIPCIÓN": lista_desc[idx],
                            "CIUDAD": lista_ciudad[idx],
                            "ESTADO": "NO_ENCONTRADO"
                        })
                    except:
                        # Serial encontrado
                        try:
                            serial_span = wait.until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(@id,'tablaInventarioTecnicoSerial')]"))
                            )
                            serial_found = serial_span.get_attribute("textContent").strip()
                            
                            fecha_span = wait.until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(@id,'tablaFechaActualiza')]"))
                            )
                            fecha_actualizacion = fecha_span.get_attribute("textContent").strip()

                            print(f"✅ Serial {serial_found} encontrado - Fecha: {fecha_actualizacion}")

                            resultados.append({
                                "SERIAL_BUSCADO": serial_number,
                                "SAP": lista_sap[idx],
                                "DESCRIPCIÓN": lista_desc[idx],
                                "ESTADO": "OK",
                                "CIUDAD": lista_ciudad[idx],
                                "FECHA_ACTUALIZACION": fecha_actualizacion
                            })
                        except (TimeoutException, StaleElementReferenceException):
                            print(f"⚠️  Error obteniendo datos para {serial_number}")
                            resultados.append({
                                "SERIAL_BUSCADO": serial_number,
                                "SAP": lista_sap[idx],
                                "DESCRIPCIÓN": lista_desc[idx],
                                "ESTADO": "ERROR_DATOS",
                                "CIUDAD": lista_ciudad[idx]
                            })
                    break

                except (StaleElementReferenceException, TimeoutException) as e:
                    if intento == 0:
                        print(f"🔄 Reintentando serial {serial_number}...")
                        time.sleep(2)
                        continue
                    else:
                        print(f"❌ Fallo definitivo para {serial_number}: {type(e).__name__}")
                        resultados.append({
                            "SERIAL_BUSCADO": serial_number,
                            "SAP": lista_sap[idx],
                            "DESCRIPCIÓN": lista_desc[idx],
                            "ESTADO": "ERROR",
                            "CIUDAD": lista_ciudad[idx]
                        })
                        
        # Generar Excel de resultados
        print("\n📊 Generando archivo de resultados...")
        df_resultados = pd.DataFrame(resultados)
        output = io.BytesIO()
        df_resultados.to_excel(output, index=False, engine='xlsxwriter')
        output.seek(0)
        
        print(f"✅ Proceso completado: {len(resultados)} registros procesados")
        return output

    except Exception as e:
        print(f"❌ ERROR CRÍTICO en consultar_en_grf: {e}")
        traceback.print_exc()
        return None 
        
    finally:
        if driver:
            print("🔒 Cerrando driver...")
            try:
                driver.quit()
            except:
                pass


# Para testing local
if __name__ == "__main__":
    # Verificar configuración
    token = os.getenv("BROWSERLESS_TOKEN")
    if token:
        print(f"✅ BROWSERLESS_TOKEN configurado: {token[:8]}...")
    else:
        print("❌ BROWSERLESS_TOKEN NO configurado")
        print("💡 Configúralo con: export BROWSERLESS_TOKEN='tu_token'")
    
    # Opcional: test de conexión básico
    try:
        print("\n🧪 Probando conexión a Browserless...")
        driver = crear_driver_browserless()
        driver.get("https://www.google.com")
        print(f"✅ Título de página: {driver.title}")
        driver.quit()
    except Exception as e:
        print(f"❌ Error en test: {e}")