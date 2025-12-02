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

def cargar_reglas():
    try:
        reglas = pd.read_excel("condicionales.xlsx")
        reglas["Regla"] = reglas["Regla"].str.strip().str.lower()
        reglas["Activa"] = reglas["Activa"].str.strip().str.upper()
        return dict(zip(reglas["Regla"], reglas["Activa"]))
    except Exception:
        return {}   # Por defecto, todas activas

REGLAS = cargar_reglas()

def regla_activa(nombre):
    return REGLAS.get(nombre, "SI") == "SI"   # SI por defecto

def obtener_ruta_recurso(nombre_archivo):
    return os.path.abspath(os.path.join(CARPETA_CANALIZADOR, nombre_archivo))

def canalizadorLocalidad(df):
    ruta = obtener_ruta_recurso("canalizador referencia lucas.xlsx")
    try:
        canalizador = pd.read_excel(ruta)
        #canalizador["CP Destino"] = canalizador["CP Destino"].astype(str).str[:-2]
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
        #canalizador["CP Destino"] = canalizador["CP Destino"].astype(str).str[:-2]
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
        if archivo == "condicionales.xlsx":
            continue
        else:
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

    df["Equipo"] = df["Equipo"].astype(str).str.strip()
    df["Nro. identificación pieza según cliente"] = df["Nro. identificación pieza según cliente"].astype(str).str.strip()

    duplicados = df.duplicated(subset="Nro. identificación pieza según cliente", keep="first")
    df.loc[duplicados, "Nro. identificación pieza según cliente"] = df.loc[duplicados, "Equipo"]
    df = df.drop_duplicates(subset=["Equipo"], keep="first")

    df["Latitud"] = ""
    df["Longitud"] = ""
    df["Distrito Destino"] = ""
    df["Provincia"] = ""
    df["Atributo1"] = df["Tipo"]
    #df["CP Destino"] = df["CP Destino"].astype(str)
    df.loc[df["Tipo"] == "Envio", "Tiempo espera"] = 10

    df.loc[df["Nombre Solicitante"] == "BIOMERIEUX ARGENTINA S.A.", "Hora Hasta"] = 1300
    df.loc[df["Nombre Solicitante"] == "GOBIERNO DE LA CIUDAD DE BUENOS AIR", "Hora Hasta"] = 1300

    # ONETO -- regla
    if regla_activa("aplicar_oneto_1100"):
        condicion_oneto = df["Dirección destino"].str.contains(r"onett?o", case=False, na=False)
        filtro = condicion_oneto & (df["Atributo1"] == "Envio")
        df.loc[filtro, "Hora Hasta"] = 1100

    # Excluir direcciones Iriarte 3070
    def normalizar(texto):
        if pd.isna(texto):
            return ""
        texto = str(texto).lower()
        texto = texto.replace(",", " ").replace(".", " ")
        texto = " ".join(texto.split())
        return texto

    if regla_activa("limpiar_iriarte_3070"):
        patrones = [r"iriarte\s*3070", r"iriarte.*3070", r"iriart.*3070",r"IRIART.*3070"]
        regex_combinado = "(" + "|".join(patrones) + ")"
        direcciones_normalizadas = df["Dirección destino"].apply(normalizar)
        condicion_iriarte = direcciones_normalizadas.str.contains(regex_combinado, na=False)
        #si contiene iriarte en la direccion y 3070 en altura, se borra
        condicion_iriarte2 = df["Dirección destino"].str.contains("iriart", case=False, na=False) & (df["Altura"] == 3070)
        df = df.drop(df[condicion_iriarte2].index)
        df = df.drop(df[condicion_iriarte].index)

    df.loc[df["Atributo1"] == "Retiro", "Tiempo espera"] = 20
    df.loc[df["Atributo1"] == "Retiro", "Volumen"] = 0.05
    df.loc[df["Altura"] == 0, "Altura"] = ""

    # Exclusiones por clientes
    if regla_activa("excluir_boston"):
        df = df.drop(df[df["Nombre Solicitante"] == "BOSTON SCIENTIFIC ARGENTINA S A"].index)
    if regla_activa("excluir_renaper"):
        df = df.drop(df[df["Nombre Solicitante"] == "REGISTRO NACIONAL DE LAS PERSONAS"].index)
    if regla_activa("excluir_ocasa"):
        df = df.drop(df[df["Nombre Solicitante"] == "OCASA DISTRIBUCION POSTAL"].index)
    if regla_activa("excluir_ibm"):
        df = df.drop(df[df["Nombre Solicitante"] == "IBM Argentina S.R.L."].index)

    # Rutas virtuales
    if regla_activa("aplicar_ruta_centra"):
        contengaCentra = df["Destinatario"].str.contains(r"CENTRA", case=False, na=False)
        contengaVega = df["Dirección destino"].str.contains(r"vega", case=False, na=False)
        #latitud = "-34.5863142398097"
        #longitud = "-58.4397811665643"
        df.loc[contengaCentra & contengaVega, "Ruta Virtual"] = 1
        #df.loc["Latitud"] = latitud
        #df.loc["Longitud"] = longitud

    if regla_activa("aplicar_ruta_inaer"):
        contengaInaer = df["Destinatario"].str.contains(r"INAER|ina", case=False, na=False)
        contengaArenales = df["Dirección destino"].str.contains(r"aren", case=False, na=False)
        df.loc[contengaInaer & contengaArenales, "Ruta Virtual"] = 2

    if regla_activa("aplicar_ruta_maffei"):
        contengaMaffei = df["Destinatario"].str.contains(r"MAFFEI", case=False, na=False)
        contengaCervi = df["Dirección destino"].str.contains(r"cervi", case=False, na=False)
        df.loc[contengaMaffei & contengaCervi, "Ruta Virtual"] = 3
        df.loc[contengaMaffei & contengaCervi, "CP Destino"] = 1426

    # Ajustes de altura
    df["Altura"] = df["Altura"].astype(str)
    df.loc[df["Altura"] == "nan", "Altura"] = ""
    df["Altura"] = df["Altura"].str[:-2]
    #concatenar altura con direccion
    df["Dirección destino"] = df["Dirección destino"] + " " + df["Altura"]

    return df

def procesar():
    df_completo = cargar_datos()
    if df_completo is None:
        messagebox.showwarning("Atención", "No hay archivos para procesar.")
        return

    # Si ya existe archivo previo
    if os.path.exists("subirUnigis-SALUD.xlsx"):
        messagebox.showerror("Error", "Eliminar archivo procesado Anteriormente.")
        os._exit(0)

    # Reglas externas
    if regla_activa("borrar_mhtml"):
        borrarMHTML()

    if regla_activa("borrar_xlsx_previos"):
        borrarXLSX()

    df_completo = df_completo[df_completo["Destino"].isin(IATAS_VALIDOS)]
    df_total = pd.DataFrame()

    for iata in IATAS_VALIDOS:
        df = df_completo[df_completo["Destino"] == iata].copy()
        if df.empty:
            continue

        df = manipularDatos(df)
        df = canalizadorLocalidad(df)
        df = canalizadorProvincia(df)

        # Corrección de direcciones
        if regla_activa("corregir_direcciones"):
            condicion_laPlata = df["Distrito Destino"] != "LA PLATA"
            df.loc[condicion_laPlata, "Dirección destino"] = (
                df.loc[condicion_laPlata, "Dirección destino"]
                .apply(ordenar_y_corregir_direccion)
            )

        df["Distrito Destino"] = df["Distrito Destino"].str.strip()

        if regla_activa("aplicar_ruta_502"):
            filtro2 = (df["Distrito Destino"] == "CAPITAL FEDERAL") & (df["Hora Desde"] >= 1400)
            df.loc[filtro2, "Ruta Virtual"] = 502

        if regla_activa("aplicar_ruta_600"):
            filtro3 = (df["Distrito Destino"] == "CAPITAL FEDERAL") & (df["Nombre Solicitante"] == "GOBIERNO DE LA CIUDAD DE BUENOS AIR") & (df["Atributo1"] == "Retiro")
            df.loc[filtro3, "Ruta Virtual"] = 600

        df_total = pd.concat([df_total, df], ignore_index=True)

    if df_total.empty:
        print("No se generaron datos válidos.")
        return

    nombre_salida = "subirUnigis-SALUD.xlsx"
    df_total.to_excel(nombre_salida, index=False)
    os.startfile(nombre_salida)
    print("Proceso finalizado:", nombre_salida)

#   INTERFAZ
def ejecutar_proceso():
    try:
        procesar()
        ventana.after(0, ventana.destroy)
    except:
        ventana.after(0, lambda: messagebox.showerror("Error", "Ocurrió un error."))

def ejecutar_en_thread():
    spinner.start(10)
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

spinner = ttk.Progressbar(frame, mode="determinate", length=180)
spinner.pack(pady=10)

footer = tk.Label(
    ventana,
    text="SALUD -- Ruteo Centralizado.",
    font=("Arial", 9),
    fg="gray"
)
footer.pack(side="bottom", pady=5)
ventana.mainloop()