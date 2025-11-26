import io
import os
import time
import traceback
from typing import Optional
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    StaleElementReferenceException, 
    TimeoutException,
    WebDriverException
)
from selenium.webdriver.chrome.options import Options
import sys

# Rutas de Chrome y ChromeDriver en Render
BASE_CHROME_DIR = "/opt/render/project/src/.chrome"
CHROME_BIN = f"{BASE_CHROME_DIR}/chrome-linux64/chrome"
DRIVER_BIN = f"{BASE_CHROME_DIR}/chromedriver-linux64/chromedriver"


def crear_driver_local() -> webdriver.Chrome:
    """
    Crea un driver local de Chrome en Render.
    """
    print("🔍 Verificando instalación de Chrome...")
    print(f"Chrome Binary: {CHROME_BIN}")
    print(f"ChromeDriver: {DRIVER_BIN}")
    print(f"Chrome existe: {os.path.exists(CHROME_BIN)}")
    print(f"Driver existe: {os.path.exists(DRIVER_BIN)}")
    sys.stdout.flush()
    
    if not os.path.exists(CHROME_BIN):
        raise FileNotFoundError(f"❌ Chrome no encontrado en: {CHROME_BIN}")
    
    if not os.path.exists(DRIVER_BIN):
        raise FileNotFoundError(f"❌ ChromeDriver no encontrado en: {DRIVER_BIN}")
    
    chrome_options = Options()
    
    # Configuración headless y seguridad
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--remote-debugging-port=9222")
    
    # Anti-detección
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option("useAutomationExtension", False)
    
    # SSL
    chrome_options.add_argument("--ignore-certificate-errors")
    
    # Ruta del binario de Chrome
    chrome_options.binary_location = CHROME_BIN
    
    print("🚀 Inicializando ChromeDriver local...")
    sys.stdout.flush()
    
    try:
        service = Service(DRIVER_BIN)
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_page_load_timeout(60)
        driver.implicitly_wait(10)
        
        print("✅ Driver Chrome creado con éxito")
        sys.stdout.flush()
        return driver
        
    except Exception as e:
        print(f"❌ Error creando driver: {str(e)}")
        traceback.print_exc()
        sys.stdout.flush()
        raise


def imprimir_progreso(actual, total, serial, estado):
    """
    Imprime una barra de progreso visual en los logs.
    """
    porcentaje = (actual / total) * 100
    barra_longitud = 50
    barra_completa = int((actual / total) * barra_longitud)
    barra = "█" * barra_completa + "░" * (barra_longitud - barra_completa)
    
    # Calcular tiempo estimado restante
    # Asumiendo ~10 segundos por serial
    seriales_restantes = total - actual
    tiempo_restante_min = (seriales_restantes * 10) / 60
    
    print(f"\n{'='*70}")
    print(f"📊 PROGRESO: [{barra}] {porcentaje:.1f}% ({actual}/{total})")
    print(f"🔍 Serial actual: {serial}")
    print(f"📌 Estado: {estado}")
    print(f"⏱️  Tiempo restante estimado: ~{tiempo_restante_min:.1f} minutos")
    print(f"{'='*70}\n")
    sys.stdout.flush()


def consultar_en_grf(usuario: str, contra: str, archivo) -> Optional[io.BytesIO]:
    """
    Automatiza consultas en el sistema GRF usando Selenium.
    """
    driver = None
    
    try:
        print("🚀 Inicializando ChromeDriver...")
        sys.stdout.flush()
        driver = crear_driver_local()
        wait = WebDriverWait(driver, 30)

        print("🌐 Abriendo página GRF...")
        sys.stdout.flush()
        driver.get("https://grf.claro.com.co:8202/GIT-web/")

        print("🔐 Ingresando credenciales...")
        sys.stdout.flush()
        driver.find_element(By.NAME, "j_idt41").send_keys(usuario)
        driver.find_element(By.NAME, "j_idt43").send_keys(contra)

        print("👆 Haciendo login...")
        sys.stdout.flush()
        btn_login = wait.until(EC.element_to_be_clickable((By.NAME, "j_idt47")))
        btn_login.click()

        print("📋 Navegando a Inventario...")
        sys.stdout.flush()
        menu_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "i.fa.fa-th-large")))
        menu_icon.click()

        inventarios_link = wait.until(EC.element_to_be_clickable((By.ID, "irVistaInventarioMenu")))
        inventarios_link.click()

        btn_consul = wait.until(EC.element_to_be_clickable((By.ID, "irVistaConsultaInventario")))
        btn_consul.click()

        print("✅ Sesión iniciada correctamente")
        sys.stdout.flush()
        
        # Configuración de tipo de consulta
        stic = 2  # 1: NACIONALES, 2: BOGOTÁ
        
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

        total_seriales = len(lista_seriales)
        print(f"\n{'#'*70}")
        print(f"📊 INICIANDO PROCESAMIENTO DE {total_seriales} SERIALES")
        print(f"⏱️  Tiempo estimado total: ~{(total_seriales * 10) / 60:.1f} minutos")
        print(f"{'#'*70}\n")
        sys.stdout.flush()
        
        resultados = []

        # Procesar cada serial
        for idx, serial_number in enumerate(lista_seriales, start=1):
            # Mostrar progreso cada 10 seriales o en el primero/último
            if idx % 10 == 0 or idx == 1 or idx == total_seriales:
                imprimir_progreso(idx, total_seriales, serial_number, "Procesando...")

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
                        resultados.append({
                            "SERIAL_BUSCADO": serial_number,
                            "SAP": lista_sap[idx-1],
                            "DESCRIPCIÓN": lista_desc[idx-1],
                            "CIUDAD": lista_ciudad[idx-1],
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

                            resultados.append({
                                "SERIAL_BUSCADO": serial_number,
                                "SAP": lista_sap[idx-1],
                                "DESCRIPCIÓN": lista_desc[idx-1],
                                "ESTADO": "OK",
                                "CIUDAD": lista_ciudad[idx-1],
                                "FECHA_ACTUALIZACION": fecha_actualizacion
                            })
                        except (TimeoutException, StaleElementReferenceException):
                            resultados.append({
                                "SERIAL_BUSCADO": serial_number,
                                "SAP": lista_sap[idx-1],
                                "DESCRIPCIÓN": lista_desc[idx-1],
                                "ESTADO": "ERROR_DATOS",
                                "CIUDAD": lista_ciudad[idx-1]
                            })
                    break

                except (StaleElementReferenceException, TimeoutException) as e:
                    if intento == 0:
                        print(f"🔄 Reintentando serial {serial_number}...")
                        sys.stdout.flush()
                        time.sleep(2)
                        continue
                    else:
                        print(f"❌ Fallo definitivo en serial {serial_number}")
                        sys.stdout.flush()
                        resultados.append({
                            "SERIAL_BUSCADO": serial_number,
                            "SAP": lista_sap[idx-1],
                            "DESCRIPCIÓN": lista_desc[idx-1],
                            "ESTADO": "ERROR",
                            "CIUDAD": lista_ciudad[idx-1]
                        })
                        
        # Generar Excel de resultados
        print(f"\n{'#'*70}")
        print("📊 Generando archivo de resultados...")
        print(f"{'#'*70}\n")
        sys.stdout.flush()
        
        df_resultados = pd.DataFrame(resultados)
        output = io.BytesIO()
        df_resultados.to_excel(output, index=False, engine='xlsxwriter')
        output.seek(0)
        
        print(f"\n{'#'*70}")
        print(f"✅ PROCESO COMPLETADO: {len(resultados)} registros procesados")
        print(f"{'#'*70}\n")
        sys.stdout.flush()
        
        return output

    except Exception as e:
        print(f"❌ ERROR CRÍTICO: {e}")
        traceback.print_exc()
        sys.stdout.flush()
        return None 
        
    finally:
        if driver:
            print("🔒 Cerrando driver...")
            sys.stdout.flush()
            try:
                driver.quit()
            except:
                pass
