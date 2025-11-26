import pandas as pd
import glob
import os
import re
from send2trash import send2trash
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import threading
import sys

CARPETA_CANALIZADOR = "canalizador"
IATAS_VALIDOS = ["BUE", "IBUE", "GBAS", "GBAO", "GBAN", "LPG"]

def obtener_ruta_recurso(nombre_archivo):
    return os.path.abspath(os.path.join(CARPETA_CANALIZADOR, nombre_archivo))

def canalizadorLocalidad(df):
    ruta = obtener_ruta_recurso("canalizador referencia lucas.xlsx")
    try:
        canalizador = pd.read_excel(ruta)
        canalizador_reducido = canalizador[["CP Destino", "Distrito Destino"]]

        df = df.drop(columns=["Distrito Destino"], errors="ignore")
        merge = pd.merge(df, canalizador_reducido, on="CP Destino", how="left")

        if "Altura" in merge.columns:
            indice = merge.columns.get_loc("Altura") + 1
            columna_valores = merge.pop("Distrito Destino")
            merge.insert(indice, "Distrito Destino", columna_valores)

        return merge

    except Exception:
        return df

def canalizadorProvincia(df):
    ruta = obtener_ruta_recurso("canalizador referencia lucas.xlsx")
    try:
        canalizador = pd.read_excel(ruta)
        canalizador_reducido = canalizador[["CP Destino", "Provincia", "ZONIFICACION"]]

        df = df.drop(columns=["Provincia"], errors="ignore")
        merge = pd.merge(df, canalizador_reducido, on="CP Destino", how="left")

        if "Población" in merge.columns:
            indice = merge.columns.get_loc("Población") + 1
            columna_valores = merge.pop("Provincia")
            merge.insert(indice, "Provincia", columna_valores)

        return merge

    except Exception:
        return df

def borrarMHTML():
    for archivo in glob.glob("*MHTML"):
        os.remove(archivo)

def borrarXLSX():
    for archivo in glob.glob("*.xlsx"):
        send2trash(archivo)

def obtener_archivos():
    archivo_canalizador = obtener_ruta_recurso("canalizador referencia lucas.xlsx")
    return [archivo for archivo in glob.glob("*.xlsx") if os.path.abspath(archivo) != archivo_canalizador]

def cargar_datos():
    archivos = obtener_archivos()
    if not archivos:
        return None
    lista = [pd.read_excel(archivo) for archivo in archivos]
    return pd.concat(lista, ignore_index=True)

def direccion_necesita_corregir(direccion):
    if pd.isna(direccion):
        return False
    direccion = str(direccion)
    return bool(
        re.search(r"\bAV[\.\s]*$", direccion, re.IGNORECASE) or
        re.match(r"^(AV|AVDA|AVENIDA|AV\.)", direccion, re.IGNORECASE) or
        not re.search(r"\d{1,5}", direccion)
    )

def ordenar_y_corregir_direccion(direccion):
    if pd.isna(direccion):
        return ""
    direccion = str(direccion).strip()
    if re.search(r"\bAV[\.\s]*$", direccion, flags=re.IGNORECASE):
        direccion = re.sub(r"\bAV[\.\s]*$", "", direccion, flags=re.IGNORECASE).strip()
        direccion = "Avenida " + direccion
    direccion = re.sub(r"^(AV\.?|AVDA\.?|AVENIDA|AVEN\.?)\s+", "Avenida ", direccion, flags=re.IGNORECASE)
    direccion = re.sub(r"^(Dr\.?|DR\.?|Doc\.?)\s+", "Doctor ", direccion, flags=re.IGNORECASE)
    direccion = re.sub(r"[^\w\s]", "", direccion)
    direccion = re.sub(r"\s+", " ", direccion).strip()
    match = re.match(r"(.+?)\s+(\d{1,5})", direccion)
    if match:
        calle = match.group(1).strip()
        altura = match.group(2).strip()
        return f"{calle} {altura}"
    return direccion

def manipularDatos(df):
    df = df.copy()

    df["Equipo"] = df["Equipo"].astype(str).str.strip().str.upper()
    df["Nro. identificación pieza según cliente"] = df["Nro. identificación pieza según cliente"].astype(str).str.strip().str.upper()
    duplicados = df.duplicated(subset="Nro. identificación pieza según cliente", keep="first")
    df.loc[duplicados, "Nro. identificación pieza según cliente"] = df.loc[duplicados, "Equipo"]
    df = df.drop_duplicates(subset=["Equipo"], keep="first")

    df["Latitud"] = ""
    df["Longitud"] = ""
    df["Distrito Destino"] = ""
    df["Provincia"] = ""
    df["Atributo1"] = ""
    df["Atributo1"] = df["Tipo"]

    df.loc[df["Tipo"] == "Envio", "Tiempo espera"] = 10

    df.loc[df["Nombre Solicitante"] == "BIOMERIEUX ARGENTINA S.A.", "Hora Hasta"] = 1300
    df.loc[df["Nombre Solicitante"] == "GOBIERNO DE LA CIUDAD DE BUENOS AIR", "Hora Hasta"] = 1300
    
    condicion_oneto = df["Dirección destino"].str.contains(r"onett?o", case=False, na=False)
    filtro = condicion_oneto & (df["Atributo1"] == "Envio")
    df.loc[filtro, "Hora Hasta"] = 1100

    def normalizar(texto):
        if pd.isna(texto):
            return ""
        texto = str(texto).lower()
        texto = texto.replace(",", " ").replace(".", " ")
        texto = " ".join(texto.split())
        return texto

    patrones = [
        r"iriarte\s*3070",
        r"iriarte.*3070",
        r"iriart.*3070",
    ]
    regex_combinado = "(" + "|".join(patrones) + ")"

    direcciones_normalizadas = df["Dirección destino"].apply(normalizar)
    condicion_iriarte = direcciones_normalizadas.str.contains(regex_combinado, na=False)
    df = df.drop(df[condicion_iriarte].index)

    df.loc[df["Atributo1"] == "Retiro", "Tiempo espera"] = 20
    df.loc[df["Atributo1"] == "Retiro", "Volumen"] = 0.05
    df.loc[df["Altura"] == 0, "Altura"] = ""

    df = df.drop(df[df["Nombre Solicitante"] == "BOSTON SCIENTIFIC ARGENTINA S A"].index)
    df = df.drop(df[df["Nombre Solicitante"] == "REGISTRO NACIONAL DE LAS PERSONAS"].index)
    df = df.drop(df[df["Nombre Solicitante"] == "OCASA DISTRIBUCION POSTAL"].index)
    df = df.drop(df[df["Nombre Solicitante"] == "IBM Argentina S.R.L."].index)

    contengaCentra = df["Destinatario"].str.contains(r"CENTRA|centra|Centra", case=False, na=False)
    contengaVega = df["Dirección destino"].str.contains(r"vega|VEGA|Vega", case=False, na=False)
    filtro4 = contengaCentra & contengaVega
    df.loc[filtro4, "Ruta Virtual"] = 1

    contengaInaer = df["Destinatario"].str.contains(r"INAER|inaer|Inaer|ina|Ina", case=False, na=False)
    contengaArenales = df["Dirección destino"].str.contains(r"ARENALES|Arenales|arenales|aren|Aren", case=False, na=False)
    filtro5 = contengaInaer & contengaArenales
    df.loc[filtro5, "Ruta Virtual"] = 2

    contengaMaffei = df["Destinatario"].str.contains(r"MAFFEI|maffei|Maffei|maffe|Maffe", case=False, na=False)
    contengaCervi = df["Dirección destino"].str.contains(r"CERVIÑO|cerviño|Cerviño|Cervino|cervino|cervi|CERVI|Cervi", case=False, na=False)
    filtro6 = contengaMaffei & contengaCervi
    df.loc[filtro6, "Ruta Virtual"] = 3

    df["Altura"] = df["Altura"].astype(str)
    df.loc[df["Altura"] == "nan", "Altura"] = ""
    df["Altura"] = df["Altura"].str[:-2]

    df["Dirección destino"] = df["Dirección destino"] + " " + df["Altura"]

    return df

def procesar():
    df_completo = cargar_datos()
    if df_completo is None:
        messagebox.showwarning("Atención", "No hay archivos para procesar.")
        return
    
    #ver si ya existe el archivo antes de procesar uno nuevo
    if os.path.exists("subirUnigis-SALUD.xlsx"):
        messagebox.showerror("Error", "Eliminar archivo procesado Anteriormente.")
        os._exit(0) #mata la ejecucion del programa.

    df_completo = df_completo[df_completo["Destino"].isin(IATAS_VALIDOS)]
    df_total = pd.DataFrame()

    for iata in IATAS_VALIDOS:
        df = df_completo[df_completo["Destino"] == iata].copy()
        if df.empty:
            continue

        df = manipularDatos(df)
        df = canalizadorLocalidad(df)
        df = canalizadorProvincia(df)

        condicion_laPlata = df["Distrito Destino"] != "LA PLATA"
        df.loc[condicion_laPlata, "Dirección destino"] = (
            df.loc[condicion_laPlata, "Dirección destino"]
            .apply(ordenar_y_corregir_direccion)
        )

        df["Distrito Destino"] = df["Distrito Destino"].str.strip()
        condicionProvincia = df["Distrito Destino"] == "CAPITAL FEDERAL"
        condicionHora = df["Hora Desde"] >= 1400
        filtro2 = condicionProvincia & condicionHora
        df.loc[filtro2, "Ruta Virtual"] = 502

        condicionProvincia = df["Distrito Destino"] == "CAPITAL FEDERAL"
        condicionCliente = df["Nombre Solicitante"] == "GOBIERNO DE LA CIUDAD DE BUENOS AIR"
        filtro3 = condicionProvincia & condicionCliente
        df.loc[filtro3, "Ruta Virtual"] = 600

        df_total = pd.concat([df_total, df], ignore_index=True)

    if df_total.empty:
        print("No se generaron datos válidos.")
        return

    borrarMHTML()
    borrarXLSX()
    nombre_salida = "subirUnigis-SALUD.xlsx"
    df_total.to_excel(nombre_salida, index=False)
    os.startfile(nombre_salida)
    print("Proceso finalizado:", nombre_salida)

# ---------------------------
#   INTERFAZ + SPINNER
# ---------------------------

def ejecutar_proceso():
    try:
        procesar()
        ventana.after(0, ventana.destroy)
    except Exception as e:
        ventana.after(0, lambda: messagebox.showerror("Error", f"Ocurrió un error:"))

def ejecutar_en_thread():
    spinner.start(10)        # ← Inicia el spinner
    boton.config(state="disabled")
    hilo = threading.Thread(target=ejecutar_proceso)
    hilo.start()

ventana = tk.Tk()
ventana.title("Procesador de Canalizador")
ventana.geometry("400x250")
ventana.resizable(False, False)

frame = tk.Frame(ventana)
frame.pack(expand=True)

boton = tk.Button(
    frame,
    text="Procesar",
    font=("Arial", 18, "bold"),
    width=10,
    height=1,
    command=ejecutar_en_thread
)
boton.pack(pady=10)

# --- Spinner gráfico ---
spinner = ttk.Progressbar(
    frame,
    mode="determinate",
    length=180,
    

)
spinner.pack(pady=10)

# --- Footer ---
footer = tk.Label(
    ventana,
    text="SALUD -- Ruteo Centralizado.",
    font=("Arial", 9),
    fg="gray"
)
footer.pack(side="bottom", pady=5)

ventana.mainloop()
