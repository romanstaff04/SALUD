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
import traceback

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

    df["Nombre Solicitante"] = df["Nombre Solicitante"].astype(str).str.strip()
    df.loc[df["Nombre Solicitante"] == "nan", "Nombre Solicitante"] = ""
    df["Dirección destino"] = df["Dirección destino"].astype(str).str.strip()

    df.loc[df["Nombre Solicitante"] == "BIOMERIEUX ARGENTINA S.A.", "Hora Hasta"] = 1300
    df.loc[df["Nombre Solicitante"] == "GOBIERNO DE LA CIUDAD DE BUENOS AIR", "Hora Hasta"] = 1300

    # ---------- Normalizar Altura pero conservar como int para reglas ----------
    # Convertir altura a numérico entero (NaN -> 0)
    df["Altura_num"] = pd.to_numeric(df.get("Altura", pd.Series()), errors="coerce").fillna(0).astype(int)

    # ---------- Concatenar dirección + altura ANTES de reglas ----------
    df["Dirección destino"] = df["Dirección destino"].astype(str).str.strip()

    df["Altura_txt"] = df["Altura_num"].replace(0, "").astype(str)

    df["Dirección destino"] = (
        df["Dirección destino"] +
        df["Altura_txt"].apply(lambda x: f" {x}" if x else "")
    )

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
            "RENÉ FAVALORO 4667": ("-34.5171957", "-58.7408939")
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
        destinatarioArgerich = df["Destinatario"].str.contains("ARGERICH", case=False, na=False)
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

        #CEMEDIC
        filtro = (df["Destinatario"].str.contains("CEMEDIC", case=False, na=False)) & (df["Dirección destino"]).str.contains("5206")
        latitud = "-34.63901564532383"
        longitud = "-58.50085418553161"
        df.loc[filtro, "Latitud"] = latitud
        df.loc[filtro, "Longitud"] = longitud

        #ITALIANO
        filtro = (df["Destinatario"].str.contains("ITALIANO", case=False, na=False)) & (df["CP Destino"] == "1900")
        latitud = "-34.936162457668175"
        longitud = "-57.97247957478854"
        df.loc[filtro, "Latitud"] = latitud
        df.loc[filtro, "Longitud"] = longitud

        #GARRAHAN
        filtro = df["Destinatario"].str.contains("GARRAHAN", case=False, na=False)
        latitud = "-34.63061257816345"
        longitud = "-58.39224010230858"
        df.loc[filtro, "Latitud"] = latitud
        df.loc[filtro, "Longitud"] = longitud

        #MARCOS SASTRE y TERRADA
        direcciones = {
            "MARCOS SASTRE 01088": ("-34.47823631330188", "-58.663775096023606"),
            "MARCOS SASTRE 1088": ("-34.47823631330188", "-58.663775096023606"),
            "M SASTRE 1088": ("-34.47823631330188", "-58.663775096023606"),
            "TERRADA 89": ("-34.62929514130666", "-58.46854209597703")
        }
        for direccion, (lat, lon) in direcciones.items():
            df.loc[df["Dirección destino"].str.contains(direccion, case=False, na=False), "Latitud"] = lat
            df.loc[df["Dirección destino"].str.contains(direccion, case=False, na=False), "Longitud"] = lon

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

        #FAMILIAR
        condicion = (df["Destinatario"].str.contains("FAMILIAR", case=False, na=False)) & (df["Dirección destino"].str.contains("3954"))
        latitud = "-34.60448842835383"
        longitud = "-58.50710449024052"
        df.loc[condicion, "Latitud"] = latitud
        df.loc[condicion, "Longitud"] = longitud

        #CESAC y FLEMING
        direcciones = {
            "CEMAR 1": ("-34.600796366581726", "-58.45917282831981"),
            "CEMAR 2": ("-34.65129815380956", "-58.39711870503457"),
            "CESAC 01": ("-34.64889917292924", "-58.38940564736295"),
            "CESAC 03": ("-34.68441182829896", "-58.46387817619654"),
            "CESAC 04": ("-34.655829572926805", "-58.50656996270609"),
            "CESAC 07": ("-34.678784695828455", "-58.492550062704574"),
            "CESAC 09": ("-34.63998099181188", "-58.36607646270705"),
            "CESAC 11": ("-34.5987054043804", "-58.41071647620206"),
            "CESAC 12": ("-34.56981333145599", "-58.47427493387591"),
            "CESAC 15": ("-34.62138537852804", "-58.36975700610132"),
            "CESAC 18": ("-34.67356501100675", "-58.46781952027218"),
            "CESAC 2": ("-34.576460202607336", "-58.5094446031882"),
            "CESAC 21": ("-34.582674772953055", "-58.37675013387495"),
            "CESAC 22": ("-34.600796366581726", "-58.45917282831981"),
            "CESAC 23": ("-34.63585344545349", "-58.505683805035616"),
            "CESAC 25": ("-34.581053692409824", "-58.384971543552275"),
            "CESAC 26": ("-34.58715277497324", "-58.42574827620286"),
            "CESAC 27": ("-34.54806833615256", "-58.4875114050413"),
            "CESAC 28": ("-34.67261657247125", "-58.48667208007605"),
            "CESAC 29": ("-34.68458200604573", "-58.486566920376"),
            "CESAC 33": ("-34.586326505700924", "-58.441846420382404"),
            "CESAC 34": ("-34.60657547294443", "-58.478236762709216"),
            "CESAC 36": ("-34.62180198601483", "-58.491866962708244"),
            "CESAC 37": ("-34.66564024585642", "-58.50413241852597"),
            "CESAC 38": ("-34.606502411739314", "-58.421465105037434"),
            "CESAC 41": ("-34.631346144975545", "-58.35782674736403"),
            "CESAC 42": ("-34.640572045954", "-58.423840905035206"),
            "CESAC 43": ("-34.67761410382098", "-58.46678720503287"),
            "CESAC 45": ("-34.625343015740896", "-58.402232920379795"),
            "CESAC 46": ("-34.64847496838738", "-58.36813040503474"),
            "CESAC 47": ("-34.584386336621534", "-58.37931694921843"),
            "CESAC 48": ("-34.64997829500097", "-58.431172533870665"),
            "CESAC 49": ("-34.66172417292471", "-58.39307354736199"),
            "CESAC 5": ("-34.670380768147936", "-58.49523437804877"),
            "CESAC 50": ("-34.6058112115926", "-58.524079305037546"),
            "CESAC Nº 13": ("-34.64333929961592", "-58.48178176270686"),
            "CESAC Nº 28": ("-34.67265186672981", "-58.48675791076326"),
            "CESAC Nº 40": ("-34.64490774641399", "-58.44577484736321"),
            "CESAC Nº 45": ("-34.62533418714105", "-58.4022007338721"),
            "CESAC ZN 17": ("-34.592624582622705", "-58.42004939154606"),
            "CESAC ZN 2": ("-34.57645136880963", "-58.5094553320241"),
            "CESAC ZN 21": ("-34.58268360608936", "-58.37672867620314"),
            "CESAC ZN 23": ("-34.63583579048974", "-58.50566234736382"),
            "CESAC ZN 25": ("-34.58102893971433", "-58.38495927620322"),
            "CESAC ZN 26": ("-34.587170440293", "-58.42575900503876"),
            "CESAC ZN 33": ("-34.58627350919217", "-58.44180350503881"),
            "CESAC ZN 36": ("-34.62181964396519", "-58.491845505036444"),
            "CESAC ZN 38": ("-34.606543419960474", "-58.4218710915452"),
            "CESAC ZN 47": ("-34.58430683998972", "-58.37926330503893"),
            "CESAC ZS 1": ("-34.64950259075587", "-58.389673462706504"),
            "CESAC ZS 10": ("-34.634542000555385", "-58.38327593387156"),
            "CESAC ZS 11": ("-34.59884751587126", "-58.410714376202115"),
            "CESAC ZS 13": ("-34.643288192866734", "-58.48166824736334"),
            "CESAC ZS 14": ("-34.66061889776996", "-58.473874662705775"),
            "CESAC ZS 15": ("-34.621179643897385", "-58.36975954736471"),
            "CESAC ZS 16": ("-34.65180717292822", "-58.37473374736266"),
            "CESAC ZS 18": ("-34.67451121456703", "-58.46300417501135"),
            "CESAC ZS 19": ("-34.64290224620122", "-58.442393962706966"),
            "CESAC ZS 20": ("-34.64945922231311", "-58.438196276904264"),
            "CESAC ZS 24": ("-34.65818389803021", "-58.455590233870076"),
            "CESAC ZS 29": ("-34.68476169649406", "-58.48671579705721"),
            "CESAC ZS 3": ("-34.68442065060832", "-58.46389963386834"),
            "CESAC ZS 30": ("-34.649565772928995", "-58.404417305034656"),
            "CESAC ZS 31": ("-34.64742429418616", "-58.43521540503478"),
            "CESAC ZS 32": ("-34.65390716848493", "-58.431162876198485"),
            "CESAC ZS 35": ("-34.65637529704205", "-58.399556562706046"),
            "CESAC ZS 37": ("-34.66559612429662", "-58.504175333869576"),
            "CESAC ZS 39": ("-34.63100948895055", "-58.40990117620002"),
            "CESAC ZS 4": ("-34.655551198311485", "-58.50722196270607"),
            "CESAC ZS 40": ("-34.6448936933789", "-58.44584076270674"),
            "CESAC ZS 41": ("-34.631174700914926", "-58.35773250503594"),
            "CESAC ZS 43": ("-34.67760253472395", "-58.46676546270465"),
            "CESAC ZS 44": ("-34.6674832488099", "-58.475723133869344"),
            "CESAC ZS 45": ("-34.62549540152134", "-58.40221103387208"),
            "CESAC ZS 5": ("-34.67119829663944", "-58.494453562705125"),
            "CESAC ZS 6": ("-34.666326248159216", "-58.44200862915649"),
            "CESAC ZS 7": ("-34.678730868046", "-58.492363930160735"),
            "CESAC ZS 8": ("-34.65595219690704", "-58.395053876198276"),
            "CESAC ZS 9": ("-34.639905110651256", "-58.36601493387122"),
            "FLEMING": ("-34.571903450473805", "-58.451773281445476"),
            "CIDEA": ("-34.59851522214249", "-58.39645392486862"),
            "CHARLES": ("-34.59876613725729", "-58.39436092825608"),
            "ENHUE": ("-34.584680148572055", "-58.43597793646526"),
            "ARSEMA": ("-34.578220932295366", "-58.45782284874288"),
            "DOM CENTRO": ("-34.59412972872068", "-58.3958281964178"),
            "QUICKFOOD": ("-34.42367767301062", "-58.97647018886677"),
            "LILLY": ("-34.54715057382337", "-58.49088492307369")
        }
        for cesac_direccion, (lat, lon) in direcciones.items():
            condicion = df["Destinatario"].str.contains(cesac_direccion, case=False, na=False)
            df.loc[condicion, "Latitud"] = lat
            df.loc[condicion, "Longitud"] = lon


    df = df.drop(columns=["Altura_num", "Altura_txt"], errors="ignore")

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