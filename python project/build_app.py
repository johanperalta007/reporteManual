"""
Script para generar el instalador/ejecutable de la app Reporte Pricing.

Uso:
    Mac:     python build_app.py
    Windows: python build_app.py

Genera:
    Mac:     dist/Reporte Pricing.app
    Windows: dist/Reporte Pricing/Reporte Pricing.exe

El ejecutable incluye todo: Python, dependencias y archivos del proyecto.
El usuario NO necesita instalar Python ni nada adicional.
"""

import os
import sys
import platform
import PyInstaller.__main__

# Rutas
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
APP_SCRIPT = os.path.join(SCRIPT_DIR, "app_pricing.py")

# Ícono según plataforma
if platform.system() == "Darwin":  # Mac
    ICON_PATH = os.path.join(SCRIPT_DIR, "images", "logoWhite.png")
else:  # Windows
    ICON_PATH = os.path.join(SCRIPT_DIR, "images", "logoWhite.ico")

IMAGES_DIR = os.path.join(SCRIPT_DIR, "images")

# Nombre de la app
APP_NAME = "Reporte Pricing"

# Archivos y carpetas a incluir en el ejecutable
datas = [
    (IMAGES_DIR, "images"),
]

# Construir argumentos de PyInstaller
args = [
    APP_SCRIPT,
    f"--name={APP_NAME}",
    "--onedir",
    "--windowed",
    "--noconfirm",
    "--clean",
    f"--distpath={os.path.join(SCRIPT_DIR, 'dist')}",
    f"--workpath={os.path.join(SCRIPT_DIR, 'build')}",
    f"--specpath={SCRIPT_DIR}",
]

# Agregar datos
for src, dest in datas:
    args.append(f"--add-data={src}{os.pathsep}{dest}")

# Agregar ícono si existe
if os.path.exists(ICON_PATH):
    if platform.system() == "Darwin":
        # En Mac, PyInstaller necesita .icns. Usamos el PNG como referencia
        args.append(f"--icon={ICON_PATH}")
    else:
        args.append(f"--icon={ICON_PATH}")

# Hidden imports necesarios
hidden_imports = [
    "customtkinter",
    "PIL",
    "PIL._tkinter_finder",
    "pandas",
    "openpyxl",
    "selenium",
    "selenium.webdriver",
    "selenium.webdriver.chrome",
    "selenium.webdriver.chrome.webdriver",
    "selenium.webdriver.chrome.service",
    "selenium.webdriver.chrome.options",
    "selenium.webdriver.common",
    "selenium.webdriver.common.by",
    "selenium.webdriver.support",
    "selenium.webdriver.support.ui",
    "selenium.webdriver.support.expected_conditions",
    "webdriver_manager",
    "webdriver_manager.chrome",
    "webdriver_manager.core",
    "urllib3",
    "requests",
]

for imp in hidden_imports:
    args.append(f"--hidden-import={imp}")

print(f"🔨 Generando ejecutable para {platform.system()}...")
print(f"   App: {APP_NAME}")
print(f"   Script: {APP_SCRIPT}")
print(f"   Salida: {os.path.join(SCRIPT_DIR, 'dist')}")
print()

PyInstaller.__main__.run(args)

print()
print("=" * 60)
if platform.system() == "Darwin":
    print(f"✅ App generada: dist/{APP_NAME}.app")
    print(f"   Doble click para abrir, o arrastra a Aplicaciones.")
else:
    print(f"✅ App generada: dist/{APP_NAME}/{APP_NAME}.exe")
    print(f"   Doble click en el .exe para ejecutar.")
print("=" * 60)
