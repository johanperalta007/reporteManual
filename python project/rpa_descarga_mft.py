"""
RPA - Descarga automática de archivos desde MFT Banco de Bogotá
Usa Selenium para automatizar el login y descarga de archivos.

Requisitos:
    pip install selenium webdriver-manager
"""

import os
import sys
import glob
import time
import shutil
import zipfile
import subprocess
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURACIÓN ---
URL_MFT = "https://mft.bancodebogota.com.co:4443/webclient/Login.xhtml"
USUARIO = "carlos.angulo"          # Reemplazar con el usuario real
PASSWORD = "m<t6fFui"      # Reemplazar con la contraseña real
CARPETA_DESCARGA = "/Users/johan.peralta/Downloads"
CARPETA_DESTINO = "/Users/johan.peralta/Documents/Banco de Bogota/Development/reporte manual /python project/desembolsos_diarios"
CARPETA_PROYECTO = "/Users/johan.peralta/Documents/Banco de Bogota/Development/reporte manual /python project"
SCRIPT_PRICING = "script_pricing_filtrado_ordenado_mantiene color AV.py"
TIMEOUT = 30  # segundos de espera máxima por elemento


def crear_driver():
    """Configura y retorna el driver de Chrome con carpeta de descarga personalizada."""

    chrome_options = Options()
    # chrome_options.add_argument("--headless")  # Descomentar para ejecutar sin ventana
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--allow-insecure-localhost")

    chrome_options.add_experimental_option("prefs", {
        "download.default_directory": CARPETA_DESCARGA,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    })

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.maximize_window()
    return driver


def esperar_elemento(driver, by, valor, timeout=TIMEOUT):
    """Espera a que un elemento esté presente y sea clickeable."""
    return WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, valor))
    )


def login(driver):
    """Paso 1: Navegar a la URL e iniciar sesión."""
    print("🌐 Navegando a MFT...")
    driver.get(URL_MFT)
    time.sleep(2)

    # Ingresar usuario
    print("👤 Ingresando usuario...")
    campo_usuario = esperar_elemento(driver, By.ID, "username")
    campo_usuario.clear()
    campo_usuario.send_keys(USUARIO)

    # Ingresar contraseña
    print("🔑 Ingresando contraseña...")
    campo_password = esperar_elemento(driver, By.ID, "value")
    campo_password.clear()
    campo_password.send_keys(PASSWORD)

    # Click en botón de login
    print("🔓 Iniciando sesión...")
    boton_login = esperar_elemento(driver, By.ID, "j_id_1o")
    boton_login.click()
    time.sleep(3)
    print("✅ Login exitoso.")


def navegar_a_carpeta(driver):
    """Paso 2: Hacer click en el elemento de navegación de carpeta."""
    print("📂 Navegando a la carpeta de archivos...")
    enlace_carpeta = esperar_elemento(driver, By.ID, "fileListForm:j_id_8l:0:j_id_8s")
    enlace_carpeta.click()
    time.sleep(3)
    print("✅ Carpeta abierta.")


def seleccionar_todos(driver):
    """Paso 3: Marcar el checkbox de 'seleccionar todos'."""
    print("☑️  Seleccionando todos los archivos...")
    try:
        # Intentar click en el checkbox usando el div contenedor
        checkbox = WebDriverWait(driver, TIMEOUT).until(
            EC.element_to_be_clickable((
                By.CSS_SELECTOR,
                "div.ui-chkbox-box.ui-widget.ui-corner-all.ui-state-default"
            ))
        )
        checkbox.click()
    except Exception:
        # Fallback: buscar por JavaScript si el click directo falla
        print("⚠️  Click directo falló, intentando con JavaScript...")
        checkbox = driver.find_element(
            By.CSS_SELECTOR,
            "div.ui-chkbox-box.ui-widget.ui-corner-all.ui-state-default"
        )
        driver.execute_script("arguments[0].click();", checkbox)

    time.sleep(2)
    print("✅ Todos los archivos seleccionados.")


def descargar_archivos(driver):
    """Paso 4: Click en el botón de descarga."""
    print("⬇️  Descargando archivos...")
    boton_descarga = esperar_elemento(driver, By.ID, "downloadFiles:downloadFiles")
    boton_descarga.click()
    time.sleep(5)
    print(f"✅ Descarga iniciada. Archivos se guardarán en: {CARPETA_DESCARGA}")


def esperar_descargas(timeout=120):
    """Espera a que las descargas terminen (no haya archivos .crdownload)."""
    print("⏳ Esperando a que finalicen las descargas...")
    inicio = time.time()
    while time.time() - inicio < timeout:
        archivos = os.listdir(CARPETA_DESCARGA)
        if archivos and not any(f.endswith(".crdownload") for f in archivos):
            print(f"✅ Descarga completada.")
            return True
        time.sleep(2)
    print("⚠️  Timeout esperando descargas.")
    return False


def buscar_zip_mas_reciente():
    """Busca el archivo .zip más reciente en la carpeta de descargas."""
    print("🔍 Buscando el ZIP más reciente en Downloads...")
    patron = os.path.join(CARPETA_DESCARGA, "*.zip")
    archivos_zip = glob.glob(patron)

    if not archivos_zip:
        raise FileNotFoundError(f"No se encontró ningún archivo .zip en {CARPETA_DESCARGA}")

    # Ordenar por fecha de modificación (más reciente primero)
    zip_reciente = max(archivos_zip, key=os.path.getmtime)
    print(f"✅ ZIP encontrado: {os.path.basename(zip_reciente)}")
    return zip_reciente


def descomprimir_y_copiar(ruta_zip):
    """Descomprime el ZIP en una carpeta temporal y copia todo a desembolsos_diarios."""
    carpeta_temp = os.path.join(CARPETA_DESCARGA, "_temp_rpa_extract")

    # Limpiar carpeta temporal si existe
    if os.path.exists(carpeta_temp):
        shutil.rmtree(carpeta_temp)

    # Descomprimir
    print(f"📦 Descomprimiendo {os.path.basename(ruta_zip)}...")
    with zipfile.ZipFile(ruta_zip, 'r') as zip_ref:
        zip_ref.extractall(carpeta_temp)

    # Determinar la carpeta raíz del contenido extraído
    contenido = os.listdir(carpeta_temp)
    if len(contenido) == 1 and os.path.isdir(os.path.join(carpeta_temp, contenido[0])):
        # Si el ZIP contiene una sola carpeta, usar su contenido
        carpeta_origen = os.path.join(carpeta_temp, contenido[0])
    else:
        carpeta_origen = carpeta_temp

    # Crear carpeta destino si no existe
    os.makedirs(CARPETA_DESTINO, exist_ok=True)

    # Copiar archivos sobrescribiendo los existentes
    archivos_copiados = 0
    archivos_nuevos = 0
    for item in os.listdir(carpeta_origen):
        origen = os.path.join(carpeta_origen, item)
        destino = os.path.join(CARPETA_DESTINO, item)

        ya_existia = os.path.exists(destino)

        if os.path.isfile(origen):
            shutil.copy2(origen, destino)  # copy2 preserva metadata
        elif os.path.isdir(origen):
            if os.path.exists(destino):
                shutil.rmtree(destino)
            shutil.copytree(origen, destino)

        if ya_existia:
            archivos_copiados += 1
        else:
            archivos_nuevos += 1

    print(f"📋 Archivos copiados a: {CARPETA_DESTINO}")
    print(f"   📄 Reemplazados: {archivos_copiados} | Nuevos: {archivos_nuevos}")

    # Limpiar carpeta temporal
    shutil.rmtree(carpeta_temp)
    print("🧹 Carpeta temporal eliminada.")


def ejecutar_script_pricing():
    """Ejecuta el script de pricing desde la carpeta del proyecto."""
    print("\n📊👷🏼‍♂️ Construyendo archivo de pricing con desembolsos...")
    ruta_script = os.path.join(CARPETA_PROYECTO, SCRIPT_PRICING)

    resultado = subprocess.run(
        [sys.executable, ruta_script],
        cwd=CARPETA_PROYECTO,
        capture_output=True,
        text=True
    )

    # Mostrar salida del script
    if resultado.stdout:
        print(resultado.stdout)
    if resultado.stderr:
        print(resultado.stderr)

    if resultado.returncode != 0:
        raise RuntimeError(f"El script de pricing falló con código {resultado.returncode}")

    print("=" * 60)
    print("✅ SE HA TERMINADO DE AGREGAR LOS DESEMBOLSOS DE MANERA EXITOSA")
    print("=" * 60)


def main():
    driver = None
    try:
        driver = crear_driver()
        login(driver)
        navegar_a_carpeta(driver)
        seleccionar_todos(driver)
        descargar_archivos(driver)
        esperar_descargas()

        # Post-procesamiento: descomprimir y copiar
        ruta_zip = buscar_zip_mas_reciente()
        descomprimir_y_copiar(ruta_zip)

        # Ejecutar script de pricing
        ejecutar_script_pricing()

        print("\n🎉 Proceso RPA completado exitosamente.")
    except Exception as e:
        print(f"\n❌ Error durante la ejecución: {e}")
        if driver:
            driver.save_screenshot("error_screenshot.png")
            print("📸 Screenshot de error guardado como 'error_screenshot.png'")
        raise
    finally:
        if driver:
            input("\nPresiona Enter para cerrar el navegador...")
            driver.quit()


if __name__ == "__main__":
    main()
