## CREDENCIALES BANCO

### Usuario: carlos.angulo
### Contraseña: m<t6fFui
### URL: (https://mft.bancodebogota.com.co:4443/webclient/Login.xhtml)


## PROCESO PARA GENERAR REPORTE CON DESEMBOLSOS

### 1. Limpiar variables de entorno e instalar los propias

```
python3 -m venv venv
source venv/bin/activate

pip install pandas openpyxl

pip install selenium webdriver-manager

pip install python-dotenv || pip3 install python-dotenv
```


### 2. Para ejecutar (RUN) el proyecto

```
python "script_pricing_filtrado_ordenado_DELETE_DUPLICADOS.py"
```

```
python "script_pricing_filtrado_ordenado_mantiene color AV.py"
```

### 3. Para ejecutar RPA del proyecto

```
python rpa_descarga_mft.py
```

