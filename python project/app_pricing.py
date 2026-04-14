"""
App Reporte Pricing - Interfaz Gráfica
Permite seleccionar archivo de Pricing y carpeta de salida.
Muestra el progreso en tiempo real en una consola integrada.

Requisitos:
    pip3 install pandas openpyxl customtkinter
"""

import os
import sys
import threading
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


# --- Redirigir prints a la consola de la GUI ---
class ConsoleRedirector:
    def __init__(self, textbox):
        self.textbox = textbox

    def write(self, message):
        if message.strip() == "":
            return
        self.textbox.after(0, self._append, message)

    def _append(self, message):
        self.textbox.configure(state="normal")
        self.textbox.insert("end", message + "\n")
        self.textbox.see("end")
        self.textbox.configure(state="disabled")

    def flush(self):
        pass


# --- Lógica de procesamiento ---
def ejecutar_procesamiento(archivo_pricing, archivo_suite, carpeta_txt, salida_final):
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Color
    from copy import copy

    COL_SPREAD_REAL = "AT"
    COL_SPREAD_SUGERIDO = "AU"
    COL_DESVIACION = "AV"
    COLOR_VERDE = "FF00B050"
    COLOR_ROJO = "FFC00000"

    def normalizar_numero_credito(valor):
        if pd.isna(valor) or str(valor).strip() == '':
            return valor
        numero = str(valor).strip()
        if '.' in numero:
            numero = numero.split('.')[0]
        if len(numero) >= 14:
            return numero
        return numero.zfill(14)

    def _to_float(v):
        if v is None:
            return None
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip()
        if s == "":
            return None
        s = s.replace(",", ".")
        try:
            return float(s)
        except:
            return None

    try:
        print("📂 Cargando archivo de Pricing...")
        df_pricing = pd.read_excel(archivo_pricing)
        print(f"   ✅ {len(df_pricing)} registros cargados desde Pricing")

        print("📂 Cargando archivo de Suite Digital...")
        df_suite = pd.read_excel(archivo_suite)
        print(f"   ✅ {len(df_suite)} registros cargados desde Suite Digital")

        print(f"📄 Procesando archivos TXT de desembolsos...")
        registros = []
        archivos_txt = [f for f in os.listdir(carpeta_txt) if f.endswith(".txt")]
        print(f"   📁 {len(archivos_txt)} archivos TXT encontrados")

        for archivo in archivos_txt:
            ruta = os.path.join(carpeta_txt, archivo)
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
        df_txt = pd.DataFrame(registros)
        print(f"   ✅ {len(df_txt)} registros de desembolsos procesados")

        print("🔧 Limpiando cotizaciones...")
        df_pricing['Cotizacion_Num'] = df_pricing['Número Cotización'].astype(str)
        df_pricing['Cotizacion_Num'] = df_pricing['Cotizacion_Num'].str.replace('SOL_PRC_', '', regex=False)
        df_pricing['Cotizacion_Num'] = df_pricing['Cotizacion_Num'].str.replace('PRI', '', regex=False)

        print("🔗 Cruzando datos con Suite Digital...")
        df_suite = df_suite.rename(columns={'x': 'Cotizacion_Num'})
        df_suite['Cotizacion_Num'] = df_suite['Cotizacion_Num'].astype(str)
        df_pricing['Cotizacion_Num'] = df_pricing['Cotizacion_Num'].astype(str)

        df_suite_min = df_suite[['Cotizacion_Num', 'NUMERO_CREDITO', 'FECHA_ACTUALIZACION']].drop_duplicates()
        df_merge = pd.merge(df_pricing, df_suite_min, on='Cotizacion_Num', how='left')
        df_merge['Número de Crédito 1'] = df_merge['Número de Crédito 1'].fillna(df_merge['NUMERO_CREDITO'])
        df_merge['Fecha de Desembolso 1'] = df_merge['Fecha de Desembolso 1'].fillna(df_merge['FECHA_ACTUALIZACION'])

        print("🔗 Cruzando datos con archivos TXT...")
        df_txt['Cotizacion_Num'] = df_txt['Cotizacion_Num'].astype(str)
        df_txt_min = df_txt[['Cotizacion_Num', 'Número de Crédito 1', 'Fecha de Desembolso 1']].drop_duplicates()
        df_final = pd.merge(
            df_merge.drop(columns=['NUMERO_CREDITO', 'FECHA_ACTUALIZACION'], errors='ignore'),
            df_txt_min, on='Cotizacion_Num', how='left', suffixes=('', '_txt')
        )
        df_final['Número de Crédito 1'] = df_final['Número de Crédito 1'].fillna(df_final['Número de Crédito 1_txt'])
        df_final['Fecha de Desembolso 1'] = df_final['Fecha de Desembolso 1'].fillna(df_final['Fecha de Desembolso 1_txt'])

        print("🔥 Aplicando filtros...")
        df_final = df_final[~df_final['Radicador'].str.upper().eq('MARIZA@BANCODEBOGOTA.COM.CO')]
        df_final = df_final[~df_final['Estado'].str.upper().eq('BORRADOR')]

        print("🔥 Resolviendo duplicados (priorizando aprobados)...")
        prioridad_estado = {"APROBADO": 1, "APROBADA": 1, "APROBADA DIGITAL": 1, "APROBADA WEB": 1}
        df_final["prioridad"] = df_final["Estado"].str.upper().map(prioridad_estado).fillna(2)
        df_final = df_final.sort_values(by=["Número Cotización", "prioridad"], ascending=[True, True])
        df_final = df_final.drop_duplicates(subset=["Número Cotización"], keep="first")
        df_final['Número de Crédito 1'] = df_final['Número de Crédito 1'].apply(normalizar_numero_credito)
        df_final = df_final.sort_values(by="Número Cotización", ascending=True)
        df_final = df_final.drop(columns=[
            'Cotizacion_Num', 'Número de Crédito 1_txt', 'Fecha de Desembolso 1_txt', 'prioridad'
        ], errors='ignore')

        print(f"📊 Registros finales: {len(df_final)}")

        print(f"💾 Guardando reporte en: {os.path.basename(salida_final)}")
        df_final.to_excel(salida_final, index=False)
        print("✅ Archivo Excel guardado")

        print("🎨 Pintando columna Desviación (AV)...")
        wb = load_workbook(salida_final)
        ws = wb.active
        verde = Color(rgb=COLOR_VERDE)
        rojo = Color(rgb=COLOR_ROJO)
        pintadas = 0
        for r in range(2, ws.max_row + 1):
            spread = _to_float(ws[f"{COL_SPREAD_REAL}{r}"].value)
            sugerido = _to_float(ws[f"{COL_SPREAD_SUGERIDO}{r}"].value)
            if spread is None or sugerido is None:
                continue
            celda_av = ws[f"{COL_DESVIACION}{r}"]
            fuente_actual = celda_av.font if celda_av.font else Font()
            nueva = copy(fuente_actual)
            if spread > sugerido:
                nueva.color = verde
                celda_av.font = nueva
                pintadas += 1
            elif spread < sugerido:
                nueva.color = rojo
                celda_av.font = nueva
                pintadas += 1
        wb.save(salida_final)
        print(f"✅ {pintadas} celdas pintadas en columna AV")

        print("")
        print("=" * 50)
        print("✅ PROCESO COMPLETADO EXITOSAMENTE")
        print(f"📁 Archivo: {os.path.basename(salida_final)}")
        print("=" * 50)
        return True

    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        return False


# --- Interfaz Gráfica con CustomTkinter ---
class AppPricing(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Reporte Pricing - Banco de Bogotá")
        self.geometry("720x580")
        self.resizable(False, False)

        # Rutas fijas
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.archivo_suite = os.path.join(script_dir, "excelFiles", "Reporte de Suite_Digital.xlsx")
        self.carpeta_txt = os.path.join(script_dir, "desembolsos_diarios")
        self.default_salida = os.path.join(script_dir, "excelFiles")

        # Variables
        self.ruta_pricing = ""
        self.ruta_salida = self.default_salida

        self._crear_interfaz()

    def _crear_interfaz(self):
        # Título
        ctk.CTkLabel(self, text="Reporte Pricing", font=ctk.CTkFont(size=24, weight="bold")).pack(
            pady=(20, 0), padx=30, anchor="w")
        ctk.CTkLabel(self, text="Banco de Bogotá  —  Generador de reportes",
                     font=ctk.CTkFont(size=13), text_color="gray").pack(padx=30, anchor="w", pady=(2, 18))

        # --- Card 1: Archivo Pricing ---
        card1 = ctk.CTkFrame(self, corner_radius=10)
        card1.pack(fill="x", padx=30, pady=(0, 12))

        ctk.CTkLabel(card1, text="📄  Archivo de Pricing",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(padx=15, pady=(12, 0), anchor="w")
        ctk.CTkLabel(card1, text="Selecciona el archivo Excel (.xlsx) de Pricing que deseas procesar",
                     font=ctk.CTkFont(size=12), text_color="gray").pack(padx=15, anchor="w", pady=(2, 8))

        row1 = ctk.CTkFrame(card1, fg_color="transparent")
        row1.pack(fill="x", padx=15, pady=(0, 12))

        self.entry_pricing = ctk.CTkEntry(row1, placeholder_text="Ningún archivo seleccionado...",
                                          state="readonly", font=ctk.CTkFont(size=12), height=35)
        self.entry_pricing.pack(side="left", fill="x", expand=True)

        ctk.CTkButton(row1, text="Seleccionar archivo", width=160, height=35,
                      command=self._seleccionar_pricing).pack(side="right", padx=(10, 0))

        # --- Card 2: Carpeta de salida ---
        card2 = ctk.CTkFrame(self, corner_radius=10)
        card2.pack(fill="x", padx=30, pady=(0, 12))

        ctk.CTkLabel(card2, text="📁  Carpeta de salida",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(padx=15, pady=(12, 0), anchor="w")
        ctk.CTkLabel(card2, text="Carpeta donde se guardará el reporte generado automáticamente",
                     font=ctk.CTkFont(size=12), text_color="gray").pack(padx=15, anchor="w", pady=(2, 8))

        row2 = ctk.CTkFrame(card2, fg_color="transparent")
        row2.pack(fill="x", padx=15, pady=(0, 12))

        self.entry_salida = ctk.CTkEntry(row2, font=ctk.CTkFont(size=12), height=35, state="readonly")
        self.entry_salida.pack(side="left", fill="x", expand=True)
        # Mostrar ruta por defecto
        self.entry_salida.configure(state="normal")
        self.entry_salida.insert(0, self.ruta_salida)
        self.entry_salida.configure(state="readonly")

        ctk.CTkButton(row2, text="Seleccionar carpeta", width=160, height=35,
                      command=self._seleccionar_carpeta_salida).pack(side="right", padx=(10, 0))

        # --- Botón procesar ---
        self.btn_procesar = ctk.CTkButton(self, text="▶  Procesar Reporte", height=45,
                                          font=ctk.CTkFont(size=16, weight="bold"),
                                          command=self._iniciar_proceso)
        self.btn_procesar.pack(fill="x", padx=30, pady=(10, 12))

        # --- Consola ---
        ctk.CTkLabel(self, text="Consola de progreso:", font=ctk.CTkFont(size=12, weight="bold")).pack(
            anchor="w", padx=30)
        self.consola = ctk.CTkTextbox(self, font=ctk.CTkFont(family="Courier", size=12),
                                      fg_color="#1a1a1a", text_color="#00ff00",
                                      corner_radius=8, state="disabled")
        self.consola.pack(fill="both", expand=True, padx=30, pady=(4, 18))

    def _seleccionar_pricing(self):
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo de Pricing",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if ruta:
            self.ruta_pricing = ruta
            self.entry_pricing.configure(state="normal")
            self.entry_pricing.delete(0, "end")
            self.entry_pricing.insert(0, ruta)
            self.entry_pricing.configure(state="readonly")

    def _seleccionar_carpeta_salida(self):
        ruta = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if ruta:
            self.ruta_salida = ruta
            self.entry_salida.configure(state="normal")
            self.entry_salida.delete(0, "end")
            self.entry_salida.insert(0, ruta)
            self.entry_salida.configure(state="readonly")

    def _iniciar_proceso(self):
        if not self.ruta_pricing:
            messagebox.showwarning("Campo requerido", "Selecciona el archivo de Pricing.")
            return
        if not self.ruta_salida:
            messagebox.showwarning("Campo requerido", "Selecciona la carpeta de salida.")
            return

        # Limpiar consola
        self.consola.configure(state="normal")
        self.consola.delete("1.0", "end")
        self.consola.configure(state="disabled")

        self.btn_procesar.configure(state="disabled", text="⏳ Procesando...")

        sys.stdout = ConsoleRedirector(self.consola)
        sys.stderr = ConsoleRedirector(self.consola)

        hoy = datetime.now().strftime("%d-%m-%Y")
        salida = os.path.join(self.ruta_salida, f"Reporte_Pricing_Actualizado_{hoy}.xlsx")

        def proceso():
            resultado = ejecutar_procesamiento(
                archivo_pricing=self.ruta_pricing,
                archivo_suite=self.archivo_suite,
                carpeta_txt=self.carpeta_txt,
                salida_final=salida
            )
            self.after(0, self._proceso_terminado, resultado)

        threading.Thread(target=proceso, daemon=True).start()

    def _proceso_terminado(self, exito):
        sys.stdout = sys.__stdout__
        sys.stderr = sys.__stderr__
        self.btn_procesar.configure(state="normal", text="▶  Procesar Reporte")
        if exito:
            messagebox.showinfo("Listo", "El reporte se generó exitosamente.")


def main():
    app = AppPricing()
    app.mainloop()


if __name__ == "__main__":
    main()
