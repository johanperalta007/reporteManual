# 📊 Reporte Pricing - Banco de Bogotá

> **Powered By ADL - JP Data Systems**

Herramienta para automatizar la generación del reporte de Pricing con desembolsos diarios.

---

## 🔐 Credenciales MFT

| Campo       | Valor                                                              |
|-------------|--------------------------------------------------------------------|
| Usuario     | `carlos.angulo`                                                    |
| Contraseña  | `m<t6fFui`                                                         |
| URL         | https://mft.bancodebogota.com.co:4443/webclient/Login.xhtml        |

---

## 🛠️ Instalación del entorno de desarrollo

### Requisitos previos

- Python 3.13 (instalar con `brew install python-tk@3.13` en Mac)
- Google Chrome instalado

### 1. Crear entorno virtual

```bash
cd "/Users/johan.peralta/Documents/Banco de Bogota/Development/reporte manual /python project"
python3.13 -m venv venv313
source venv313/bin/activate
```

### 2. Instalar dependencias

```bash
pip install pandas openpyxl customtkinter selenium webdriver-manager python-dotenv Pillow pyinstaller
```

---

### Script de pricing por terminal

```bash
source venv313/bin/activate
python "script_pricing_filtrado_ordenado_mantiene color AV.py"
```

### RPA por terminal (descarga + procesamiento)

```bash
source venv313/bin/activate
python rpa_descarga_mft.py
```

## 🚀 Ejecución

### Interfaz gráfica (recomendado)

```bash
python3.13 -m venv venv313
source venv313/bin/activate

python app_pricing.py
```
____________________________________________________________________________________

## 📦 Generar instalador

### 🍎 Mac (.dmg)

**Paso 1:** Ubicarse en la carpeta del proyecto y activar el entorno:

```bash
cd "/Users/johan.peralta/Documents/Banco de Bogota/Development/reporte manual /python project"
source venv313/bin/activate
```

**Paso 2:** Generar el `.app`:

```bash
python build_app.py
```

**Paso 3:** Generar el `.dmg` (instalador distribuible):

```bash
rm -f dist/*.dmg dist/Reporte_Pricing_Installer.dmg && create-dmg \
  --volname "Reporte Pricing" \
  --window-pos 200 120 \
  --window-size 600 400 \
  --icon-size 100 \
  --icon "Reporte Pricing.app" 150 190 \
  --app-drop-link 450 190 \
  --no-internet-enable \
  "dist/Reporte_Pricing_Installer.dmg" \
  "dist/Reporte Pricing.app"
```

**Resultado:** `dist/Reporte_Pricing_Installer.dmg`

> 💡 Si no tienes `create-dmg`, instálalo con: `brew install create-dmg`

---

### 🪟 Windows (.exe)

**Paso 1:** Abrir terminal git Bash y ubicarse en la carpeta del proyecto.

**Paso 2:** Crear entorno virtual *(solo la primera vez)*:

```bash
python -m venv venv313

pip install pandas openpyxl customtkinter selenium webdriver-manager Pillow pyinstaller
```

**Paso 3:** Activar entorno virtual *(si ya existe)*:

```bash
source venv313/Scripts/activate
pip install pandas openpyxl customtkinter selenium webdriver-manager pyinstaller
```

**Paso 4:** Generar el `.exe`:

```bash
python build_app.py
```

**Resultado:** `dist\Reporte Pricing\Reporte Pricing.exe`

> 💡 Para distribuir, comprimir la carpeta `dist\Reporte Pricing\` en un `.zip` y enviar a los compañeros.

---

## ⚠️ Notas importantes

- Cada vez que modifiques `app_pricing.py`, hay que regenerar el instalador.
- El instalador de **Mac solo funciona en Mac**, el de **Windows solo en Windows**. No hay cross-compile.
- Si agregas una nueva dependencia con `pip install`, hay que regenerar el instalador.
- El archivo `.env` solo contiene el nombre del archivo de Pricing de entrada. Es lo único que se edita manualmente cuando se corre por terminal.
