import io
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
import pandas as pd
import time
import os
from selenium.webdriver.chrome.options import Options

def crear_driver_browserless():
    chrome_options = Options()
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--window-size=1280,720")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--allow-running-insecure-content")

    browserless_token = os.getenv("BROWSERLESS_TOKEN")
    # 🔹 URL HTTPS correcta para Selenium
    command_executor_url = f"https://chrome.browserless.io/webdriver?token={browserless_token}"

    driver = webdriver.Remote(
        command_executor=command_executor_url,
        options=chrome_options
    )
    return driver

# -----------------------------
# Función principal
# -----------------------------
def consultar_en_grf(usuario, contra, archivo):

    print("🚀 Inicializando Browserless ChromeDriver...")
    driver = crear_driver_browserless()
    wait = WebDriverWait(driver, 30)

    try:
        print("🌐 Abriendo página...")
        driver.get("https://grf.claro.com.co:8202/GIT-web/")
#46250702
#Marzo026**
        print("Ingresando usuario...")
        driver.find_element(By.NAME, "j_idt41").send_keys(usuario)
        
        print("Ingresando contraseña...")
        driver.find_element(By.NAME, "j_idt43").send_keys(contra)

        print("Dando click en login...")
        btn_login = wait.until(EC.element_to_be_clickable((By.NAME, "j_idt47")))
        btn_login.click()

        print("Dando click en inventario...")
        menu_icon = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "i.fa.fa-th-large")))
        menu_icon.click()

        print("Dando click en inventario...")
        inventarios_link = wait.until(EC.element_to_be_clickable((By.ID, "irVistaInventarioMenu")))
        inventarios_link.click()

        print("Dando click en consultar inventario...")
        btn_consul = wait.until(EC.element_to_be_clickable((By.ID, "irVistaConsultaInventario")))
        btn_consul.click()

        print("BIENVENIDO KEINER PRO, ERES EL MEJOR DE TODOS BRO")
        print("1. NACIONALES")
        print("2. BOGOTÁ")
        stic = 2
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        
        if stic == 1:
            ruta = os.path.join(BASE_DIR, 'INGRESO', 'NACIONALES.xlsx')
            doc_envios = pd.read_excel(archivo)
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

        elif stic == 2:
            ruta = os.path.join(BASE_DIR, 'INGRESO', 'BOGOTA.xlsx')
            doc_envios = pd.read_excel(archivo)
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

        casca = 0 
        resultados = []

        for serial_number in lista_seriales:
            print("\n--- Procesando serial:", serial_number)

            for intento in range(2):
                try:
                    campo_serial = wait.until(EC.presence_of_element_located((By.ID, "idSerialBuscar")))
                    campo_serial.clear()
                    campo_serial.send_keys(serial_number)

                    buscar_link = wait.until(EC.element_to_be_clickable((By.ID, "btnInventarioBuscar")))
                    buscar_link.click()

                    WebDriverWait(driver, 30).until(
                        EC.any_of(
                            EC.presence_of_element_located((By.XPATH, "//td[contains(text(),'No hay resultados en la Base de Datos')]")),
                            EC.presence_of_element_located((By.XPATH, "//span[contains(@id,'tablaInventarioTecnicoSerial')]"))
                        )
                    )

                    try:
                        driver.find_element(By.XPATH, "//td[contains(text(),'No hay resultados en la Base de Datos')]")
                        print(f"❌ No hay datos para el serial {serial_number}.")
                        resultados.append({
                            "SERIAL_BUSCADO": serial_number,
                            "SAP": lista_sap[casca],
                            "DESCRIPCIÓN": lista_desc[casca],
                            "CIUDAD": lista_ciudad[casca],
                            "ESTADO": "NO_ENCONTRADO"
                        })
                        casca += 1
                    except:
                        try:
                            fecha_span = wait.until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(@id,'tablaFechaActualiza')]"))
                            )
                            fecha_actualizacion = fecha_span.get_attribute("textContent").strip()

                            serial_span = wait.until(
                                EC.presence_of_element_located((By.XPATH, "//span[contains(@id,'tablaInventarioTecnicoSerial')]"))
                            )
                            serial_found = serial_span.get_attribute("textContent").strip()

                            print(f"✅ La fecha de actualización para el serial {serial_found} es: {fecha_actualizacion}")

                            resultados.append({
                                "SERIAL_BUSCADO": serial_number,
                                "SAP": lista_sap[casca],
                                "DESCRIPCIÓN": lista_desc[casca],
                                "ESTADO": "OK",
                                "CIUDAD": lista_ciudad[casca]
                            })
                            casca += 1
                        except (TimeoutException, StaleElementReferenceException):
                            print(f"⚠️ No se encontraron datos de actualización para {serial_number}.")
                            resultados.append({
                                "SERIAL_BUSCADO": serial_number,
                                "SAP": lista_sap[casca],
                                "DESCRIPCIÓN": lista_desc[casca],
                                "ESTADO": "ERROR_DATOS",
                                "CIUDAD": lista_ciudad[casca]
                            })
                            casca += 1
                    break

                except (StaleElementReferenceException, TimeoutException) as e:
                    if intento == 0:
                        print(f"🔄 Reintentando serial {serial_number} por fallo: {type(e).__name__}")
                        time.sleep(2)
                        continue
                    else:
                        print(f"❌ Fallo al procesar {serial_number}: {type(e).__name__} - {str(e)}")
                        resultados.append({
                            "SERIAL_BUSCADO": serial_number,
                            "SAP": lista_sap[casca],
                            "DESCRIPCIÓN": lista_desc[casca],
                            "ESTADO": "ERROR",
                            "CIUDAD": lista_ciudad[casca]
                        })
                        casca += 1
                        
        df_resultados = pd.DataFrame(resultados)
        output = io.BytesIO()
        df_resultados.to_excel(output, index=False, engine='xlsxwriter')
        output.seek(0)
        
        return output

    except Exception as e:
        print(f"❌ ERROR CRÍTICO en consultar_en_grf: {e}")
        import traceback
        traceback.print_exc()
        return None 
        
    finally:
        if driver:
            print("Cerrando driver...")
            driver.quit()