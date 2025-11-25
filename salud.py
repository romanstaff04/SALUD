import pandas as pd
import glob
import os
import re
from send2trash import send2trash

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
    #eliminar duplicados
    duplicados = df.duplicated(subset="Nro. identificación pieza según cliente", keep="first")
    df.loc[duplicados, "Nro. identificación pieza según cliente"] = df.loc[duplicados, "Equipo"]
    df = df.drop_duplicates(subset=["Equipo"], keep="first")

    #vaciar columnas
    df["Latitud"] = ""
    df["Longitud"] = ""
    df["Distrito Destino"] = ""
    df["Provincia"] = ""
    df["Atributo1"] = ""
    #copiar columna AI en la columna Z
    df["Atributo1"] = df["Tipo"]

    #lo que es envio, en columna "m", 10
    df.loc[df["Tipo"] == "Envio", "Tiempo espera"] = 10

    df.loc[df["Nombre Solicitante"] == "BIOMERIEUX ARGENTINA S.A.", "Hora Hasta"] = 1300
    df.loc[df["Nombre Solicitante"] == "GOBIERNO DE LA CIUDAD DE BUENOS AIR", "Hora Hasta"] = 1300
    
    #si es envio y la direccion contiene onetto hora 1100
    condicion_oneto = df["Dirección destino"].str.contains(r"onett?o", case=False, na=False)
    filtro = condicion_oneto & (df["Atributo1"] == "Envio")
    df.loc[filtro, "Hora Hasta"] = 1100

    #si es org courrier y direccion contiene iriarte o lafayette, se borra esa fila.
    condicion_iriarte_lafayette = df["Dirección destino"].str.contains(r"iriarte|lafayette", case=False, na=False)
    filtro = condicion_iriarte_lafayette & (df["Nombre Solicitante"] == "ORG COURIER ARG")
    df = df.drop(df[filtro].index)

    return df

def procesar():
    df_completo = cargar_datos()
    if df_completo is None:
        print("No hay archivos para procesar.")
        return

    # Filtrar dataset solo a los IATAs permitidos
    df_completo = df_completo[df_completo["Destino"].isin(IATAS_VALIDOS)]

    df_total = pd.DataFrame()

    for iata in IATAS_VALIDOS:
        df = df_completo[df_completo["Destino"] == iata].copy()
        if df.empty:
            continue


        df = manipularDatos(df)
        df = canalizadorLocalidad(df)
        df = canalizadorProvincia(df)
        df["Dirección destino corregida"] = df["Dirección destino"].apply(ordenar_y_corregir_direccion)

        #capital federal y horaDesde mayor o igual a 1400, 502
        df["Distrito Destino"] = df["Distrito Destino"].str.strip()
        condicionProvincia = df["Distrito Destino"] == "CAPITAL FEDERAL"
        condicionHora = df["Hora Desde"] >= 1400
        filtro2 = condicionProvincia & condicionHora
        df.loc[filtro2, "Ruta Virtual"] = 502


        df_total = pd.concat([df_total, df], ignore_index=True)

    if df_total.empty:
        print("No se generaron datos válidos.")
        return

    #borrarMHTML()
    #borrarXLSX()
    nombre_salida = "subirUnigis-SALUD.xlsx"
    df_total.to_excel(nombre_salida, index=False)
    os.startfile(nombre_salida)
    print("Proceso finalizado:", nombre_salida)


if __name__ == "__main__":
    procesar()
