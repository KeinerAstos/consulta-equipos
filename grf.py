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
    
    # Calcular tiempo estimado restante (10 segundos por serial)
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
        wait = WebDriverWait(driver, 60)  # 60 segundos

        print("🌐 Abriendo página GRF...")
        sys.stdout.flush()
        driver.get("https://grf.claro.com.co:8202/GIT-web/")

        print(f"🔐 Ingresando credenciales para usuario: {usuario}")
        sys.stdout.flush()
        
        # Usar wait.until y clear() antes de send_keys
        usuario_field = wait.until(EC.presence_of_element_located((By.NAME, "j_idt41")))
        usuario_field.clear()
        usuario_field.send_keys(usuario)
        
        contra_field = wait.until(EC.presence_of_element_located((By.NAME, "j_idt43")))
        contra_field.clear()
        contra_field.send_keys(contra)

        print("👆 Haciendo login...")
        sys.stdout.flush()
        btn_login = wait.until(EC.element_to_be_clickable((By.NAME, "j_idt47")))
        btn_login.click()

        # Agregar espera después del login
        print("⏳ Esperando que cargue el dashboard (10 segundos)...")
        sys.stdout.flush()
        time.sleep(10)

        print("🔍 Buscando menú de inventario...")
        sys.stdout.flush()
        
        # Probar múltiples selectores
        menu_icon = None
        selectores = [
            ("CSS", "i.fa.fa-th-large"),
            ("XPATH", "//i[contains(@class, 'fa-th-large')]"),
            ("XPATH", "//i[@class='fa fa-th-large']"),
            ("CSS", ".fa.fa-th-large")
        ]
        
        for tipo, selector in selectores:
            try:
                if tipo == "CSS":
                    menu_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, selector)))
                else:
                    menu_icon = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                print(f"✅ Menú encontrado con selector {tipo}: {selector}")
                sys.stdout.flush()
                break
            except TimeoutException:
                print(f"❌ No se encontró con {tipo}: {selector}")
                sys.stdout.flush()
                continue
        
        if not menu_icon:
            print("❌ No se pudo encontrar el menú con ningún selector")
            print("📄 Primeros 1000 caracteres de la página:")
            print(driver.page_source[:1000])
            sys.stdout.flush()
            raise Exception("No se encontró el menú de inventario. Posible error de login.")

        print("📋 Haciendo click en el menú...")
        sys.stdout.flush()
        menu_icon.click()
        time.sleep(3)  # Espera después del click

        print("📂 Buscando link de Inventarios...")
        sys.stdout.flush()
        inventarios_link = wait.until(EC.element_to_be_clickable((By.ID, "irVistaInventarioMenu")))
        inventarios_link.click()
        time.sleep(2)  # Espera después del click

        print("🔎 Buscando botón Consultar Inventario...")
        sys.stdout.flush()
        btn_consul = wait.until(EC.element_to_be_clickable((By.ID, "irVistaConsultaInventario")))
        btn_consul.click()
        time.sleep(2)  # Espera después del click

        print("✅ Sesión iniciada correctamente")
        sys.stdout.flush()
        
        # Configuración de tipo de consulta (Se asume un estándar o se requiere una entrada)
        # Para mantener la lógica de doc.py:
        # stic = 2  # 1: NACIONALES, 2: BOGOTÁ
        
        # NOTA: En un entorno de servidor como Render, no se puede usar input(), 
        # por lo que el parámetro 'archivo' debe implicar el tipo de consulta
        # o se debe hardcodear uno de los tipos (por ejemplo, BOGOTÁ = 2).
        # Aquí se mantiene la flexibilidad basada en la estructura del archivo.
        
        # Detección simple para simular 'stic' (se recomienda pasar 'stic' como argumento)
        try:
             # Leer las primeras filas y determinar si es "Serial" o "NºSerieFab"
            df_temp = pd.read_excel(archivo, nrows=1)
            if "Serial" in df_temp.columns and "Codigo material" in df_temp.columns:
                stic = 2 # BOGOTÁ
            elif "NºSerieFab" in df_temp.columns and "Material" in df_temp.columns:
                stic = 1 # NACIONALES
            else:
                # Si no se detecta, asumir BOGOTÁ como en doc.py
                print("⚠️ No se pudo determinar el tipo de consulta, asumiendo BOGOTÁ (stic=2).")
                stic = 2
        except Exception as e:
            print(f"⚠️ Error al leer columnas para determinar stic: {e}. Asumiendo BOGOTÁ (stic=2).")
            stic = 2
        
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
            lista_desc = doc_envios["Texto breve de material"] # CAMBIO: Nombre de columna según doc.py
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
        print(f"📊 INICIANDO PROCESAMIENTO DE {total_seriales} SERIALES (TIPO: {'NACIONALES' if stic == 1 else 'BOGOTÁ'})")
        print(f"⏱️  Tiempo estimado total: ~{(total_seriales * 10) / 60:.1f} minutos")
        print(f"{'#'*70}\n")
        sys.stdout.flush()
        
        resultados = []

        # Procesar cada serial
        for idx, serial_number_raw in enumerate(lista_seriales, start=1):
            
            # Limpiar serial_number
            serial_number = str(serial_number_raw).strip()
            
            # Manejar serial vacío/invalido (Añadido)
            if not serial_number:
                imprimir_progreso(idx, total_seriales, serial_number_raw, "SERIAL VACÍO")
                resultados.append({
                    "SERIAL_BUSCADO": serial_number_raw,
                    "SAP": lista_sap[idx-1],
                    "DESCRIPCIÓN": lista_desc[idx-1],
                    "CIUDAD": lista_ciudad[idx-1],
                    "ESTADO": "VACÍO",
                    "BODEGA_GRF": "N/A" # CAMBIO: Añadir columna
                })
                continue
                
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
                            "ESTADO": "NO_ENCONTRADO",
                            "BODEGA_GRF": "N/A" # CAMBIO: Añadir columna
                        })
                    except:
                        # Serial encontrado
                        try:
                            # CAMBIO: Extracción de bodega, crucial para la funcionalidad de doc.py
                            serial_span = wait.until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(@id,'tablaInventarioTecnicoSerial')]"))
                            )
                            # serial_found = serial_span.get_attribute("textContent").strip() # No se usa en el resultado final, pero se puede mantener

                            fecha_span = wait.until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(@id,'tablaFechaActualiza')]"))
                            )
                            fecha_actualizacion = fecha_span.get_attribute("textContent").strip()
                            
                            bodega_span = wait.until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(@id,'tablaInventarioBodega')]"))
                            )
                            bodega_found = bodega_span.get_attribute("textContent").strip()

                            resultados.append({
                                "SERIAL_BUSCADO": serial_number,
                                "SAP": lista_sap[idx-1],
                                "DESCRIPCIÓN": lista_desc[idx-1],
                                "ESTADO": "OK",
                                "CIUDAD": lista_ciudad[idx-1],
                                "FECHA_ACTUALIZACION": fecha_actualizacion,
                                "BODEGA_GRF": bodega_found # CAMBIO: Añadir columna
                            })
                        except (TimeoutException, StaleElementReferenceException):
                            resultados.append({
                                "SERIAL_BUSCADO": serial_number,
                                "SAP": lista_sap[idx-1],
                                "DESCRIPCIÓN": lista_desc[idx-1],
                                "ESTADO": "ERROR_DATOS",
                                "CIUDAD": lista_ciudad[idx-1],
                                "BODEGA_GRF": "N/A" # CAMBIO: Añadir columna
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
                            "CIUDAD": lista_ciudad[idx-1],
                            "BODEGA_GRF": "N/A" # CAMBIO: Añadir columna
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