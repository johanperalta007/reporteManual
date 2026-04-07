
import pandas as pd
import os
from datetime import datetime

# --- CONFIGURACIÓN GENERAL ---
ARCHIVO_PRICING = "Reporte_Pricing_Actualizado 23-09-2025 V1.xlsx"
ARCHIVO_SUITE = "Reporte de Suite_Digital.xlsx"
CARPETA_TXT = "desembolsos_diarios"
SALIDA_FINAL = "Reporte_Pricing_Actualizado.xlsx"

def cargar_archivos_base():
    df_pricing = pd.read_excel(ARCHIVO_PRICING)
    df_suite = pd.read_excel(ARCHIVO_SUITE)
    return df_pricing, df_suite

def limpiar_cotizaciones(df):
    df['Cotizacion_Num'] = df['Número Cotización'].astype(str)
    df['Cotizacion_Num'] = df['Cotizacion_Num'].str.replace('SOL_PRC_', '', regex=False)
    df['Cotizacion_Num'] = df['Cotizacion_Num'].str.replace('PRI', '', regex=False)
    return df

def procesar_archivos_txt(carpeta):
    registros = []
    for archivo in os.listdir(carpeta):
        if archivo.endswith(".txt"):
            ruta = os.path.join(carpeta, archivo)
            with open(ruta, 'r', encoding='utf-8') as f:
                lineas = f.readlines()
                for linea in lineas[1:]:
                    partes = linea.strip().split('|')
                    if len(partes) >= 3:
                        cotizacion = partes[0].replace('SOL_PRC_', '')
                        numero_credito = partes[1]
                        fecha_raw = partes[2]
                        try:
                            fecha = datetime.strptime(fecha_raw, '%Y%m%d').date()
                        except:
                            fecha = None
                        registros.append({
                            'Cotizacion_Num': cotizacion,
                            'Número de Crédito 1': numero_credito,
                            'Fecha de Desembolso 1': fecha
                        })
    return pd.DataFrame(registros)

def unir_y_poblar(df_pricing, df_suite, df_txt):
    df_pricing = limpiar_cotizaciones(df_pricing)
    df_pricing['Cotizacion_Num'] = df_pricing['Cotizacion_Num'].astype(str)
    df_suite = df_suite.rename(columns={'x': 'Cotizacion_Num'})
    df_suite['Cotizacion_Num'] = df_suite['Cotizacion_Num'].astype(str)

    # Poblar desde Suite Digital
    df_suite_min = df_suite[['Cotizacion_Num', 'NUMERO_CREDITO', 'FECHA_ACTUALIZACION']].drop_duplicates()
    df_merge = pd.merge(df_pricing, df_suite_min, on='Cotizacion_Num', how='left')
    df_merge['Número de Crédito 1'] = df_merge['Número de Crédito 1'].fillna(df_merge['NUMERO_CREDITO'])
    df_merge['Fecha de Desembolso 1'] = df_merge['Fecha de Desembolso 1'].fillna(df_merge['FECHA_ACTUALIZACION'])

    # Poblar desde archivos TXT solo si sigue faltando
    df_txt['Cotizacion_Num'] = df_txt['Cotizacion_Num'].astype(str)
    df_txt_min = df_txt[['Cotizacion_Num', 'Número de Crédito 1', 'Fecha de Desembolso 1']].drop_duplicates()
    df_final = pd.merge(df_merge.drop(columns=['NUMERO_CREDITO', 'FECHA_ACTUALIZACION'], errors='ignore'),
                        df_txt_min,
                        on='Cotizacion_Num',
                        how='left',
                        suffixes=('', '_txt'))

    df_final['Número de Crédito 1'] = df_final['Número de Crédito 1'].fillna(df_final['Número de Crédito 1_txt'])
    df_final['Fecha de Desembolso 1'] = df_final['Fecha de Desembolso 1'].fillna(df_final['Fecha de Desembolso 1_txt'])

    # Filtrar registros no deseados
    df_final = df_final[
        ~df_final['Radicador'].str.upper().eq('MARIZA@BANCODEBOGOTA.COM.CO') &
        ~df_final['Estado'].str.upper().eq('BORRADOR')
    ]

    # Ordenar por 'Número Cotización'
    df_final = df_final.sort_values(by='Número Cotización', ascending=True)

    # Eliminar columnas intermedias
    df_final = df_final.drop(columns=['Cotizacion_Num', 'Número de Crédito 1_txt', 'Fecha de Desembolso 1_txt'], errors='ignore')

    return df_final

def guardar(df, salida=SALIDA_FINAL):
    df.to_excel(salida, index=False)
    print(f"✅ Reporte final guardado como: {salida}")

def main():
    df_pricing, df_suite = cargar_archivos_base()
    df_txt = procesar_archivos_txt(CARPETA_TXT)
    df_final = unir_y_poblar(df_pricing, df_suite, df_txt)
    guardar(df_final)

if __name__ == "__main__":
    main()
