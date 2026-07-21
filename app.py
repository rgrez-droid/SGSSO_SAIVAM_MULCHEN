import base64
import glob
import io
import mimetypes
import os
import re
from datetime import datetime
from urllib.parse import quote

import pandas as pd
import requests
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

# Carpeta donde se encuentra este archivo. Permite cargar la planilla,
# logos y fondos aunque Streamlit se ejecute desde otro directorio.
APP_DIR = os.path.dirname(os.path.abspath(__file__))

def ruta_app(*partes):
    return os.path.join(APP_DIR, *partes)

st.set_page_config(
    page_title="Sistema de Gestión SSO SAIVAM Mulchén",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Reduce conflictos del DOM provocados por la traducción automática del navegador.
st.markdown(
    """
    <meta name="google" content="notranslate">
    <style>
        html, body, [data-testid="stAppViewContainer"] {
            translate: no;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

AUTOR = "Ricardo Grez"
EMPRESA = "SAIVAM"
CONTRATO = "CMPC Mulchén"
VERSION = "1.4.13"
REVISION_CODIGO = "21-07-2026-R38-PRG-ESTADOS-DESDE-SHEET"

print(
    f"[SSO] Ejecutando archivo corregido: {os.path.abspath(__file__)} "
    f"| revisión {REVISION_CODIGO}"
)

# Fecha base para días sin accidentes si no existe hoja Configuracion.
# Puedes modificarla directamente o dejarla en la hoja Configuracion.
FECHA_INICIO_SIN_ACCIDENTES_DEFAULT = "01/04/2023"

# =========================================================
# GOOGLE SHEETS
# =========================================================
# Documento principal que alimenta toda la aplicación.
# La planilla debe estar compartida como "Cualquier persona con el enlace:
# Lector" o publicada en la web.
GOOGLE_SHEET_ID = "1GrwPn86i7dYnYxkLrr-bVeueUSh5eXqL"

# Identificador exacto de la pestaña "Cumplimientos SSO".
# Se obtiene desde el parámetro gid del enlace compartido por el usuario.
# Leer esta pestaña por GID evita depender del nombre visible de la hoja.
CUMPLIMIENTOS_SSO_GID = "745436880"

# Si una pestaña no puede leerse desde Google Sheets, el sistema conserva
# como respaldo la lectura desde el archivo Excel local.
USAR_GOOGLE_SHEETS = True

ARCHIVOS_EXCEL_POSIBLES = [
    # Base alineada con el menú actual de la aplicación.
    "Base_Datos_SSO_SAIVAM_Mulchen.xlsx",
    # Compatibilidad con versiones anteriores.
    "Base_Datos_SGS_SAIVAM_Mulchen.xlsx",
    "Base_Datos_SGS_SAIVAM_Mulchen (1).xlsx",
    # También reconoce la versión prototipo creada inicialmente.
    "Base_Datos_SGS_SAIVAM_Mulchen_Prototipo.xlsx",
    "BASE SGS SAIVAM MULCHEN.xlsx",
    "SGS_SAIVAM_MULCHEN.xlsx",
    "Sistema_Gestion_SSO_SAIVAM_Mulchen.xlsx",
]

MESES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}

HOY = pd.Timestamp.today().normalize()


# =========================================================
# ESTRUCTURA DE HOJAS / BASE DE DATOS
# =========================================================

SHEETS = {
    "Incidentes": {
        # La pestaña visible en Google Sheets se llama "Reportabilidad",
        # pero internamente la aplicación conserva la clave "Incidentes"
        # para mantener compatibilidad con el panel general y sus indicadores.
        "nombres": [
            "Reportabilidad",
            "Incidentes",
            "Incidentes y reportes",
            "Incidentes_y_Reportes",
            "Reportes",
        ],
        "secret": "reportabilidad_url",
        "columnas": [
            "Fecha",
            "Área",
            "Tipo_Evento",
            "Descripcion",
            "Accion_Inmediata",
            "Responsable",
            "Estado",
            "Observacion",
            "Ruta_Link",
        ],
    },
    "Cumplimientos_SSO": {
        "nombres": [
            "Cumplimientos SSO",
            "Cumplimiento SSO",
            "Cumplimientos_SSO",
            "Cumplimiento_SSO",
        ],
        "secret": "cumplimientos_sso_url",
        "columnas": [
            "Observador", "Actividad",
            "ENE", "RE_ENE", "FEB", "RE_FEB", "MAR", "RE_MAR",
            "ABR", "RE_ABR", "MAY", "RE_MAY", "JUN", "RE_JUN",
            "JUL", "RE_JUL", "AGO", "RE_AGO", "SEP", "RE_SEP",
            "OCT", "RE_OCT", "NOV", "RE_NOV", "DIC", "RE_DIC",
        ],
    },
    "Capacitaciones": {
        "nombres": ["Capacitaciones", "Charlas", "Charlas_Capacitaciones"],
        "secret": "capacitaciones_url",
        # Estructura exacta de la pestaña "Capacitaciones" en Google Sheets.
        "columnas": [
            "Fecha", "Tema", "Tipo", "Área", "Responsable",
            "Vencimiento", "Estado", "Observacion", "Evidencia",
        ],
    },
    "Programa_Anual": {
        "nombres": [
            "PRG_SSO_2026",
            "PRG SSO 2026",
            "Programa_Anual",
            "Programa Anual de Seguridad",
        ],
        "secret": "programa_anual_url",
        # Estructura vigente de la pestaña PRG_SSO_2026.
        "columnas": [
            "Mes",
            "Eje_Trabajo",
            "Actividad",
            "Tipo_Actividad",
            "Fecha_Programada",
            "Fecha_Realizacion",
            "Responsable",
            "Estado",
            "Evidencia",
            "Observacion",
        ],
    },
    "Reconocimientos": {
        "nombres": ["Reconocimientos", "Reconocimiento", "Premios", "Destacados"],
        "secret": "reconocimientos_url",
        "columnas": [
            "Fecha", "Trabajador", "Cargo", "Motivo", "Periodo",
            "Estado", "Evidencia", "Observacion",
        ],
    },
    "Comite_Paritario": {
        "nombres": ["Comite_Paritario", "Comité Paritario", "Comite Paritario", "CPHS"],
        "secret": "comite_paritario_url",
        "columnas": [
            "Fecha", "Tipo_Reunion", "Área", "Tema", "Acuerdo",
            "Responsable", "Fecha_Compromiso", "Estado", "Evidencia",
            "Observacion",
        ],
    },
    "Protocolos_MINSAL": {
        "nombres": ["Protocolos_MINSAL", "Protocolos MINSAL", "MINSAL"],
        "secret": "protocolos_minsal_url",
        "columnas": [
            "Fecha", "Protocolo", "Etapa", "Área", "Actividad", "Expuestos",
            "Responsable", "Resultado", "Fecha_Compromiso", "Estado",
            "Evidencia", "Observacion",
        ],
    },
    "Certificaciones": {
        "nombres": ["Certificaciones", "Certificaciones_Maestro", "Certificados"],
        "secret": "certificaciones_url",
        "columnas": [
            "Fecha",
            "Categoria",
            "Subcategoria",
            "Nombre_Certificacion",
            "Entidad_Emisora",
            "Vencimiento",
            "Estado",
            "Dias_Para_Vencer",
            "Ruta_Link",
        ],
    },
    "Configuracion": {
        "nombres": ["Configuracion", "Configuración", "Config"],
        "secret": "configuracion_url",
        "columnas": ["Parametro", "Valor"],
    },
}


# =========================================================
# UTILIDADES GENERALES
# =========================================================

def normalizar_texto(texto):
    texto = str(texto).lower().strip()
    reemplazos = {
        "á": "a",
        "é": "e",
        "í": "i",
        "ó": "o",
        "ú": "u",
        "ñ": "n",
        "ü": "u",
    }
    for original, nuevo in reemplazos.items():
        texto = texto.replace(original, nuevo)
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    texto = re.sub(r"_+", "_", texto).strip("_")
    return texto


def limpiar_numero(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = re.sub(r"[^0-9,.-]", "", str(valor))
    if texto in ["", "-"]:
        return 0.0
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "." in texto and len(texto.split(".")[-1]) == 3:
        texto = texto.replace(".", "")
    try:
        return float(texto)
    except ValueError:
        return 0.0


def numero(valor):
    try:
        return f"{float(valor):,.0f}".replace(",", ".")
    except Exception:
        return "0"


def porcentaje(valor):
    try:
        return f"{float(valor):.0f}%".replace(".", ",")
    except Exception:
        return "0%"


def convertir_fecha(valor):
    if pd.isna(valor):
        return pd.NaT
    if isinstance(valor, pd.Timestamp):
        return valor
    if isinstance(valor, datetime):
        return pd.Timestamp(valor)
    if isinstance(valor, (int, float)):
        try:
            return pd.to_datetime(valor, unit="D", origin="1899-12-30", errors="coerce")
        except Exception:
            return pd.NaT
    return pd.to_datetime(valor, errors="coerce", dayfirst=True)


def fecha_texto(valor):
    if pd.isna(valor):
        return ""
    try:
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except Exception:
        return ""


def escape_html(texto):
    texto = str(texto)
    texto = texto.replace("&", "&amp;")
    texto = texto.replace("<", "&lt;")
    texto = texto.replace(">", "&gt;")
    texto = texto.replace('"', "&quot;")
    texto = texto.replace("'", "&#39;")
    return texto


def normalizar_columnas_dataframe(df):
    if df is None or df.empty:
        return df

    alias = {
        "fecha": "Fecha",
        "area": "Área",
        "área": "Área",
        "trabajador": "Trabajador",
        "supervisor": "Supervisor",
        "actividad": "Actividad",
        "mes": "Mes",
        "eje_trabajo": "Eje_Trabajo",
        "ejetrabajo": "Eje_Trabajo",
        "fecha_programada": "Fecha_Programada",
        "fechaprogramada": "Fecha_Programada",
        "fecha_realizacion": "Fecha_Realizacion",
        "fecharealizacion": "Fecha_Realizacion",
        "tipo_observacion": "Tipo_Observacion",
        "tipoobservacion": "Tipo_Observacion",
        "tipo_observación": "Tipo_Observacion",
        "conducta_segura": "Conducta_Segura",
        "conductasegura": "Conducta_Segura",
        "conducta_riesgo": "Conducta_Riesgo",
        "conductariesgo": "Conducta_Riesgo",
        "medida_correctiva": "Medida_Correctiva",
        "medidacorrectiva": "Medida_Correctiva",
        "responsable": "Responsable",
        "fecha_compromiso": "Fecha_Compromiso",
        "fechacompromiso": "Fecha_Compromiso",
        "estado": "Estado",
        "observacion": "Observacion",
        "observación": "Observacion",
        "tipo_evento": "Tipo_Evento",
        "tipoevento": "Tipo_Evento",
        "gravedad": "Gravedad",
        "descripcion": "Descripcion",
        "descripción": "Descripcion",
        "accion_inmediata": "Accion_Inmediata",
        "acción_inmediata": "Accion_Inmediata",
        "accioninmediata": "Accion_Inmediata",
        "tipo_inspeccion": "Tipo_Inspeccion",
        "tipoinspeccion": "Tipo_Inspeccion",
        "tipo_inspección": "Tipo_Inspeccion",
        "resultado": "Resultado",
        "hallazgos": "Hallazgos",
        "origen": "Origen",
        "hallazgo": "Hallazgo",
        "accion_correctiva": "Accion_Correctiva",
        "acción_correctiva": "Accion_Correctiva",
        "accioncorrectiva": "Accion_Correctiva",
        "evidencia": "Evidencia",
        "tema": "Tema",
        "tipo": "Tipo",
        "relator": "Relator",
        "asistentes": "Asistentes",
        "vencimiento": "Vencimiento",
        "vencimient": "Vencimiento",
        "dias_para_vencer": "Dias_Para_Vencer",
        "días_para_vencer": "Dias_Para_Vencer",
        "diasparavencer": "Dias_Para_Vencer",
        "díasparavencer": "Dias_Para_Vencer",
        "cargo": "Cargo",
        "epp": "EPP",
        "cantidad": "Cantidad",
        "proxima_reposicion": "Proxima_Reposicion",
        "próxima_reposición": "Proxima_Reposicion",
        "proximareposicion": "Proxima_Reposicion",
        "tipo_actividad": "Tipo_Actividad",
        "tipoactividad": "Tipo_Actividad",
        "cumplimiento": "Cumplimiento",
        "tipo_reconocimiento": "Tipo_Reconocimiento",
        "tiporeconocimiento": "Tipo_Reconocimiento",
        "motivo": "Motivo",
        "periodo": "Periodo",
        "tipo_reunion": "Tipo_Reunion",
        "tiporeunion": "Tipo_Reunion",
        "acuerdo": "Acuerdo",
        "tipo_trabajo": "Tipo_Trabajo",
        "tipotrabajo": "Tipo_Trabajo",
        "permiso": "Permiso",
        "tipo_documento": "Tipo_Documento",
        "tipodocumento": "Tipo_Documento",
        "nombre_documento": "Nombre_Documento",
        "nombredocumento": "Nombre_Documento",
        "version": "Version",
        "versión": "Version",
        "ruta_link": "Ruta_Link",
        "rutalink": "Ruta_Link",
        "link": "Ruta_Link",
        "parametro": "Parametro",
        "parámetro": "Parametro",
        "protocolo": "Protocolo",
        "etapa": "Etapa",
        "expuestos": "Expuestos",
        "categoria": "Categoria",
        "categoría": "Categoria",
        "subcategoria": "Subcategoria",
        "subcategoría": "Subcategoria",
        "titular_activo": "Titular_Activo",
        "titularactivo": "Titular_Activo",
        "nombre_certificacion": "Nombre_Certificacion",
        "nombre_certificación": "Nombre_Certificacion",
        "nombrecertificacion": "Nombre_Certificacion",
        "entidad_emisora": "Entidad_Emisora",
        "entidademisora": "Entidad_Emisora",
        "numero_certificado": "Numero_Certificado",
        "número_certificado": "Numero_Certificado",
        "numerocertificado": "Numero_Certificado",
        "valor": "Valor",
    }

    nuevas = {}
    for col in df.columns:
        original = str(col).replace("\n", " ").replace("\r", " ").strip()
        clave = normalizar_texto(original)
        clave_sin_guion = clave.replace("_", "")
        nuevas[col] = alias.get(clave, alias.get(clave_sin_guion, original))

    salida = df.rename(columns=nuevas)

    if salida.columns.duplicated().any():
        consolidado = pd.DataFrame(index=salida.index)
        for columna in dict.fromkeys(salida.columns):
            bloque = salida.loc[:, salida.columns == columna]
            if bloque.shape[1] == 1:
                consolidado[columna] = bloque.iloc[:, 0]
            else:
                consolidado[columna] = bloque.bfill(axis=1).iloc[:, 0]
        salida = consolidado

    return salida


def asegurar_columnas(df, columnas):
    salida = df.copy()
    for columna in columnas:
        if columna not in salida.columns:
            salida[columna] = ""
    return salida[columnas + [c for c in salida.columns if c not in columnas]]


def preparar_fechas(df):
    salida = df.copy()
    columnas_fecha = [
        "Fecha",
        "Fecha_Compromiso",
        "Vencimiento",
        "Proxima_Reposicion",
        "Fecha_Programada",
        "Fecha_Realizacion",
    ]
    for columna in columnas_fecha:
        if columna in salida.columns:
            salida[columna] = salida[columna].apply(convertir_fecha)
    return salida


def preparar_periodo(df):
    salida = df.copy()
    if "Fecha" not in salida.columns:
        salida["Fecha"] = pd.NaT
    salida["Fecha"] = salida["Fecha"].apply(convertir_fecha)
    salida["Año"] = salida["Fecha"].dt.year
    salida["Mes_Numero"] = salida["Fecha"].dt.month
    salida["Mes"] = salida["Mes_Numero"].map(MESES)
    salida["Periodo"] = salida["Mes"].fillna("Sin mes") + " " + salida["Año"].fillna(0).astype(int).astype(str)
    return salida


def estado_base(valor):
    texto = normalizar_texto(valor)
    if texto in ["", "nan", "none", "nat", "sin_estado"]:
        return "Sin estado"
    if "sin_vencimiento" in texto:
        return "Sin vencimiento"
    if "por_vencer" in texto or "proximo_a_vencer" in texto:
        return "Por vencer"
    if "vigente" in texto:
        return "Vigente"
    if "cerr" in texto or "realiz" in texto or "cumpl" in texto or "ok" in texto:
        return "Cerrada"
    if "proceso" in texto or "gestion" in texto or "pendiente_ejecucion" in texto:
        return "En proceso"
    if "venc" in texto or "atras" in texto:
        return "Vencida"
    if "pend" in texto or "abiert" in texto:
        return "Pendiente"
    if "no_cumple" in texto or "nocumple" in texto:
        return "No cumple"
    return str(valor).strip().capitalize()


def normalizar_estados(df):
    salida = df.copy()
    if "Estado" not in salida.columns:
        salida["Estado"] = "Sin estado"
    salida["Estado"] = salida["Estado"].apply(estado_base)
    return salida



def preparar_programa_anual(df):
    """
    Limpia y prepara la pestaña PRG_SSO_2026.

    El campo Estado se toma directamente desde Google Sheets. La aplicación no
    lo recalcula usando Fecha_Programada ni Fecha_Realizacion. De esta forma,
    cada actividad mantiene exactamente la condición de gestión registrada en
    la planilla: Cerrada, Pendiente o En proceso.
    """
    columnas = SHEETS["Programa_Anual"]["columnas"]

    if df is None:
        return pd.DataFrame(columns=columnas)

    salida = normalizar_columnas_dataframe(df.copy())
    salida = asegurar_columnas(salida, columnas)
    salida = preparar_fechas(salida)

    # Elimina filas sin una actividad real. Esto evita incorporar filas de
    # apoyo, listas auxiliares o formatos copiados hacia abajo en Google Sheets.
    mascara_registro = pd.Series(False, index=salida.index)
    for columna in [
        "Actividad",
        "Eje_Trabajo",
        "Tipo_Actividad",
        "Fecha_Programada",
        "Fecha_Realizacion",
        "Responsable",
    ]:
        valores = salida[columna]
        mascara_registro = mascara_registro | (
            valores.notna()
            & valores.astype(str).str.strip().ne("")
            & valores.astype(str).str.lower().ne("nan")
            & valores.astype(str).str.lower().ne("nat")
        )

    salida = salida.loc[mascara_registro].copy()

    if salida.empty:
        return pd.DataFrame(columns=columnas)

    # Limpieza de textos.
    for columna in [
        "Mes",
        "Eje_Trabajo",
        "Actividad",
        "Tipo_Actividad",
        "Responsable",
        "Estado",
        "Evidencia",
        "Observacion",
    ]:
        salida[columna] = (
            salida[columna]
            .fillna("")
            .astype(str)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
        )

    # Completa el mes desde la fecha programada cuando la celda Mes está vacía.
    mes_desde_fecha = salida["Fecha_Programada"].dt.month.map(MESES)
    mes_vacio = salida["Mes"].eq("") | salida["Mes"].str.lower().eq("nan")
    salida.loc[mes_vacio, "Mes"] = mes_desde_fecha.loc[mes_vacio].fillna("")

    def estado_desde_sheet(valor):
        """Normaliza únicamente los tres estados utilizados en PRG_SSO_2026."""
        estado = normalizar_texto(valor)

        if "cerr" in estado or "realiz" in estado or "cumpl" in estado:
            return "Cerrada"
        if "proceso" in estado or "gestion" in estado:
            return "En proceso"
        if "pend" in estado or "program" in estado:
            return "Pendiente"

        # Una celda vacía se mantiene como pendiente para que el registro no se
        # contabilice erróneamente como cerrado por la fecha de realización.
        return "Pendiente"

    # Fuente única de verdad: columna Estado de la planilla.
    salida["Estado"] = salida["Estado"].apply(estado_desde_sheet)

    # Periodo para filtros y ordenamiento del panel.
    salida["Año"] = salida["Fecha_Programada"].dt.year
    salida["Mes_Numero"] = salida["Fecha_Programada"].dt.month
    salida["Periodo"] = (
        salida["Mes"].replace("", "Sin mes")
        + " "
        + salida["Año"].fillna(0).astype(int).astype(str)
    )

    # Orden cronológico; las filas sin fecha quedan al final.
    salida = salida.sort_values(
        by=["Fecha_Programada", "Eje_Trabajo", "Actividad"],
        ascending=[True, True, True],
        na_position="last",
    )

    return salida.reset_index(drop=True)


def preparar_certificaciones(df):
    """Limpia la hoja y calcula vigencia y días restantes."""
    salida = df.copy()

    columnas_clave = [
        "Fecha",
        "Categoria",
        "Subcategoria",
        "Nombre_Certificacion",
        "Entidad_Emisora",
    ]

    # Elimina filas completamente vacías, incluso cuando Excel tiene
    # fórmulas copiadas hacia abajo.
    mascara_registro = pd.Series(False, index=salida.index)
    for columna in columnas_clave:
        if columna in salida.columns:
            valores = salida[columna]
            mascara_registro = mascara_registro | (
                valores.notna()
                & valores.astype(str).str.strip().ne("")
                & valores.astype(str).str.lower().ne("nan")
            )

    salida = salida.loc[mascara_registro].copy()

    if "Vencimiento" not in salida.columns:
        salida["Vencimiento"] = pd.NaT

    salida["Vencimiento"] = salida["Vencimiento"].apply(convertir_fecha)

    def calcular_vigencia(fecha):
        if pd.isna(fecha):
            return "Sin vencimiento", pd.NA

        dias = int((fecha.normalize() - HOY).days)

        if dias < 0:
            return "Vencida", dias
        if dias <= 30:
            return "Por vencer", dias
        return "Vigente", dias

    resultados = salida["Vencimiento"].apply(calcular_vigencia)
    salida["Estado"] = resultados.apply(lambda resultado: resultado[0])
    salida["Dias_Para_Vencer"] = pd.array(
        resultados.apply(lambda resultado: resultado[1]),
        dtype="Int64",
    )

    return salida.reset_index(drop=True)


def marcar_vencimientos(df, columna_fecha="Fecha_Compromiso"):
    salida = df.copy()
    if columna_fecha not in salida.columns:
        return salida
    salida[columna_fecha] = salida[columna_fecha].apply(convertir_fecha)
    if "Estado" not in salida.columns:
        salida["Estado"] = "Pendiente"
    cerrada = salida["Estado"].fillna("").astype(str).str.lower().str.contains("cerr|realiz|cumpl", regex=True)
    vencida = salida[columna_fecha].notna() & (salida[columna_fecha] < HOY) & (~cerrada)
    salida.loc[vencida, "Estado"] = "Vencida"
    return salida


def buscar_archivo_excel():
    """Busca la base Excel dentro de la misma carpeta de la aplicación."""
    # Primero revisa los nombres definidos, en el orden de prioridad indicado.
    for archivo in ARCHIVOS_EXCEL_POSIBLES:
        ruta = ruta_app(archivo)
        if os.path.isfile(ruta):
            return ruta

    # Como respaldo, detecta cualquier Excel compatible ubicado junto al código.
    candidatos = []
    for ruta in glob.glob(ruta_app("*.xlsx")):
        nombre_archivo = os.path.basename(ruta)
        if nombre_archivo.startswith("~$"):
            continue
        nombre = normalizar_texto(nombre_archivo)
        if any(clave in nombre for clave in ["base_datos_sgs", "sgs", "sso", "seguridad", "preventiva"]):
            candidatos.append(ruta)

    # Prefiere la planilla modificada más recientemente cuando existen varias.
    if candidatos:
        return max(candidatos, key=os.path.getmtime)
    return None


def construir_url_google_sheet(nombre_pestana):
    """Construye una URL CSV usando el nombre visible de una pestaña."""
    nombre_codificado = quote(str(nombre_pestana), safe="")
    return (
        f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/gviz/tq"
        f"?tqx=out:csv&sheet={nombre_codificado}"
    )


def construir_url_google_sheet_gid(gid):
    """Construye una URL CSV usando el identificador estable de la pestaña."""
    return (
        f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/gviz/tq"
        f"?tqx=out:csv&gid={quote(str(gid), safe='')}"
    )


def construir_url_exportacion_gid(gid):
    """URL oficial de exportación CSV para una pestaña identificada por GID."""
    return (
        f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export"
        f"?format=csv&gid={quote(str(gid), safe='')}"
    )


def descargar_csv_google(url, header=0):
    """
    Descarga un CSV de Google Sheets y detecta respuestas de inicio de sesión.

    El navegador puede abrir una planilla privada porque el usuario tiene una
    sesión de Google activa. Streamlit, pandas y requests no reciben esa sesión,
    por lo que la hoja debe estar compartida como lector mediante enlace o debe
    utilizarse autenticación con una cuenta de servicio.
    """
    respuesta = requests.get(
        url,
        timeout=30,
        allow_redirects=True,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 Chrome/124 Safari/537.36"
            )
        },
    )
    respuesta.raise_for_status()

    texto = respuesta.content.decode("utf-8-sig", errors="replace")
    inicio = texto.lstrip().lower()[:1500]
    url_final = respuesta.url.lower()

    respuesta_privada = (
        "accounts.google.com" in url_final
        or "serviceLogin".lower() in inicio
        or "sign in" in inicio
        or "iniciar sesión" in inicio
        or inicio.startswith("<!doctype html")
        or inicio.startswith("<html")
    )

    if respuesta_privada:
        raise PermissionError(
            "Google devolvió una página de inicio de sesión en lugar del CSV. "
            "La planilla no está disponible para lectura anónima."
        )

    return pd.read_csv(io.StringIO(texto), header=header)


@st.cache_data(ttl=60, show_spinner=False)
def descargar_libro_google_xlsx(sheet_id):
    """Descarga el Google Sheet completo como XLSX para usarlo como respaldo."""
    if not sheet_id:
        return b""

    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    respuesta = requests.get(
        url,
        timeout=35,
        allow_redirects=True,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 Chrome/124 Safari/537.36"
            )
        },
    )
    respuesta.raise_for_status()

    contenido = respuesta.content
    inicio = contenido[:1500].decode("utf-8", errors="ignore").lstrip().lower()
    tipo = respuesta.headers.get("content-type", "").lower()

    if (
        "accounts.google.com" in respuesta.url.lower()
        or "text/html" in tipo
        or inicio.startswith("<!doctype html")
        or inicio.startswith("<html")
        or "sign in" in inicio
        or "iniciar sesión" in inicio
    ):
        raise PermissionError(
            "Google devolvió una página de acceso en lugar del archivo XLSX. "
            "La planilla debe estar compartida como lector mediante enlace."
        )

    if not contenido.startswith(b"PK"):
        raise ValueError("Google no devolvió un archivo XLSX válido.")

    return contenido


def _dataframe_con_registros(df):
    """Valida que la respuesta contenga filas útiles y no solo encabezados vacíos."""
    if df is None or len(df.columns) == 0:
        return False

    if df.empty:
        return False

    prueba = df.copy()
    prueba = prueba.replace(r"^\s*$", pd.NA, regex=True)
    prueba = prueba.dropna(how="all")
    return not prueba.empty


def leer_hoja_desde_google(nombres_hoja, columnas_esperadas=None):
    """
    Lee una pestaña de Google Sheets mediante dos métodos:

    1. CSV GViz utilizando el nombre visible de la pestaña.
    2. Exportación XLSX del libro completo y lectura local de la pestaña.

    El segundo método evita que módulos como Reportabilidad queden vacíos
    cuando pandas no logra abrir directamente la URL CSV de Google.
    """
    if not USAR_GOOGLE_SHEETS or not GOOGLE_SHEET_ID:
        return None

    errores = []
    esperadas = {
        normalizar_texto(columna)
        for columna in (columnas_esperadas or [])
    }

    # Método 1: CSV por nombre de pestaña, descargado con requests.
    for nombre_pestana in nombres_hoja:
        url = construir_url_google_sheet(nombre_pestana)

        try:
            df = descargar_csv_google(url, header=0)
            df = df.dropna(how="all").reset_index(drop=True)

            if not _dataframe_con_registros(df):
                errores.append(f"{nombre_pestana}: respuesta CSV sin registros")
                continue

            if esperadas:
                columnas_recibidas = {
                    normalizar_texto(columna)
                    for columna in df.columns
                }
                coincidencias = len(esperadas.intersection(columnas_recibidas))
                if coincidencias < min(3, len(esperadas)):
                    errores.append(
                        f"{nombre_pestana}: encabezados no reconocidos "
                        f"({coincidencias} coincidencias)"
                    )
                    continue

            print(
                f"[SSO] Pestaña '{nombre_pestana}' leída mediante CSV GViz. "
                f"Registros: {len(df)}"
            )
            return df

        except Exception as error:
            errores.append(
                f"{nombre_pestana}: {type(error).__name__}: {error}"
            )

    # Método 2: descarga del libro completo como XLSX.
    try:
        contenido = descargar_libro_google_xlsx(GOOGLE_SHEET_ID)
        excel = pd.ExcelFile(io.BytesIO(contenido))
        hojas = {normalizar_texto(h): h for h in excel.sheet_names}

        for nombre_pestana in nombres_hoja:
            clave = normalizar_texto(nombre_pestana)
            if clave not in hojas:
                continue

            df = pd.read_excel(
                io.BytesIO(contenido),
                sheet_name=hojas[clave],
            )
            df = df.dropna(how="all").reset_index(drop=True)

            if _dataframe_con_registros(df):
                print(
                    f"[SSO] Pestaña '{hojas[clave]}' leída desde la "
                    f"exportación XLSX. Registros: {len(df)}"
                )
                return df

        errores.append("Exportación XLSX: no se encontró una pestaña compatible")

    except Exception as error:
        errores.append(f"Exportación XLSX: {type(error).__name__}: {error}")

    print(
        "[SSO] No fue posible leer la pestaña solicitada desde Google Sheets: "
        + " | ".join(errores[-5:])
    )
    return None




MESES_CORTOS = [
    "ENE", "FEB", "MAR", "ABR", "MAY", "JUN",
    "JUL", "AGO", "SEP", "OCT", "NOV", "DIC",
]

# =========================================================
# PERIODO OFICIAL PARA EL MÓDULO CUMPLIMIENTOS SSO
# =========================================================
# El usuario puede consultar el acumulado oficial de enero a julio de 2026
# o revisar el año completo, incluyendo las metas y resultados de ENE a DIC.
ANIO_CUMPLIMIENTOS = 2026
MES_CORTE_CUMPLIMIENTOS = 7
MESES_CUMPLIMIENTOS = MESES_CORTOS[:MES_CORTE_CUMPLIMIENTOS]
PERIODO_CUMPLIMIENTOS = "Enero a julio de 2026"
PERIODO_ANUAL_CUMPLIMIENTOS = "Año completo 2026"


def _es_texto_valido(valor):
    if pd.isna(valor):
        return False
    texto = str(valor).strip()
    return texto != "" and texto.lower() not in {"nan", "none", "nat"}


def _a_numero_cumplimiento(valor):
    if not _es_texto_valido(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace("%", "").replace(" ", "")
    texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except Exception:
        return 0.0


def normalizar_hoja_cumplimientos(df_raw):
    """
    Convierte la pestaña visual "Cumplimientos SSO" en una tabla utilizable.

    Admite dos formatos:
    1) La matriz visual del Excel original, con nombres combinados, encabezados
       ENE/RE repetidos y filas Total Act., Total Realizadas y % Cumplimiento.
    2) Una tabla normalizada con columnas Observador, Actividad, ENE, RE_ENE, etc.
    """
    columnas_salida = [
        "Observador", "Actividad",
        *[col for mes in MESES_CORTOS for col in (mes, f"RE_{mes}")],
    ]

    if df_raw is None or df_raw.empty:
        return pd.DataFrame(columns=columnas_salida)

    raw = df_raw.copy()
    raw = raw.dropna(how="all").reset_index(drop=True)

    # Caso 1: hoja ya normalizada.
    claves = {normalizar_texto(c): c for c in raw.columns}
    if "observador" in claves and "actividad" in claves:
        renombrar = {
            claves["observador"]: "Observador",
            claves["actividad"]: "Actividad",
        }
        for mes in MESES_CORTOS:
            posibles_meta = [normalizar_texto(mes), normalizar_texto(mes.lower())]
            posibles_real = [
                normalizar_texto(f"RE_{mes}"),
                normalizar_texto(f"RE {mes}"),
                normalizar_texto(f"REAL_{mes}"),
                normalizar_texto(f"REAL {mes}"),
            ]
            for clave in posibles_meta:
                if clave in claves:
                    renombrar[claves[clave]] = mes
                    break
            for clave in posibles_real:
                if clave in claves:
                    renombrar[claves[clave]] = f"RE_{mes}"
                    break

        salida = raw.rename(columns=renombrar)
        salida = asegurar_columnas(salida, columnas_salida)[columnas_salida]
        for col in columnas_salida[2:]:
            salida[col] = salida[col].apply(_a_numero_cumplimiento)
        salida["Observador"] = salida["Observador"].astype(str).str.strip()
        salida["Actividad"] = salida["Actividad"].astype(str).str.strip()
        salida = salida[
            salida["Observador"].ne("")
            & salida["Actividad"].ne("")
            & ~salida["Actividad"].apply(normalizar_texto).str.contains(
                r"^total|cumplimiento", regex=True, na=False
            )
        ]
        return salida.reset_index(drop=True)

    # Caso 2: matriz visual. Ubica la columna de actividad por contenido.
    puntajes = {}
    patrones_actividad = (
        "control operacional", "ops", "inspecciones", "check list",
        "observaciones", "bapp",
    )
    for col in raw.columns:
        serie = raw[col].astype(str).str.lower()
        puntajes[col] = int(sum(serie.str.contains(p, regex=False, na=False).sum() for p in patrones_actividad))

    if not puntajes or max(puntajes.values()) == 0:
        return pd.DataFrame(columns=columnas_salida)

    col_actividad = max(puntajes, key=puntajes.get)
    posicion_actividad = list(raw.columns).index(col_actividad)
    columnas_izquierda = list(raw.columns)[:posicion_actividad]
    columnas_datos = list(raw.columns)[posicion_actividad + 1: posicion_actividad + 25]

    registros = []
    observador_actual = ""

    for _, fila in raw.iterrows():
        # Recupera el nombre del observador desde las celdas combinadas.
        candidatos = []
        for col in columnas_izquierda:
            valor = fila.get(col, "")
            if not _es_texto_valido(valor):
                continue
            texto = str(valor).strip()
            clave = normalizar_texto(texto)
            if clave.isdigit() or clave in {"actividades", "actividad"}:
                continue
            if any(token in clave for token in ["seguimiento", "cumplimiento", "total"]):
                continue
            # Nombres de personas normalmente contienen al menos un espacio.
            if any(ch.isalpha() for ch in texto) and len(texto) >= 4:
                candidatos.append(texto)
        if candidatos:
            observador_actual = max(candidatos, key=len)

        actividad = fila.get(col_actividad, "")
        if not _es_texto_valido(actividad):
            continue

        actividad = str(actividad).strip()
        clave_actividad = normalizar_texto(actividad)

        if (
            not observador_actual
            or clave_actividad in {"actividad", "actividades", "ene", "re"}
            or clave_actividad.startswith("total")
            or "cumplimiento" in clave_actividad
        ):
            continue

        es_actividad = any(
            patron in clave_actividad
            for patron in [
                "control_operacional", "ops", "inspeccion", "check_list",
                "observacion", "bapp",
            ]
        )
        if not es_actividad:
            continue

        valores = [_a_numero_cumplimiento(fila.get(col, 0)) for col in columnas_datos]
        valores += [0.0] * (24 - len(valores))

        registro = {
            "Observador": observador_actual,
            "Actividad": actividad,
        }
        for i, mes in enumerate(MESES_CORTOS):
            registro[mes] = valores[i * 2]
            registro[f"RE_{mes}"] = valores[i * 2 + 1]
        registros.append(registro)

    salida = pd.DataFrame(registros, columns=columnas_salida)
    if salida.empty:
        return salida

    salida["Observador"] = salida["Observador"].astype(str).str.strip()
    salida["Actividad"] = salida["Actividad"].astype(str).str.strip()
    return salida.reset_index(drop=True)


def leer_cumplimientos_desde_google(nombres_hoja):
    """
    Lee la matriz de Cumplimientos SSO desde Google Sheets.

    Prioridad:
    1. Exportación CSV por GID.
    2. Consulta GViz por GID.
    3. Consulta GViz por nombres alternativos.
    """
    if not USAR_GOOGLE_SHEETS or not GOOGLE_SHEET_ID:
        return None

    intentos = []

    if CUMPLIMIENTOS_SSO_GID:
        intentos.extend([
            ("Exportación CSV por GID", construir_url_exportacion_gid(CUMPLIMIENTOS_SSO_GID)),
            ("GViz por GID", construir_url_google_sheet_gid(CUMPLIMIENTOS_SSO_GID)),
        ])

    intentos.extend(
        (f"GViz por nombre: {nombre}", construir_url_google_sheet(nombre))
        for nombre in nombres_hoja
    )

    errores = []

    for metodo, url in intentos:
        try:
            raw = descargar_csv_google(url, header=None)
            normalizado = normalizar_hoja_cumplimientos(raw)

            if normalizado is not None and not normalizado.empty:
                print(
                    f"[SSO] Cumplimientos SSO leído correctamente mediante {metodo}. "
                    f"Registros: {len(normalizado)}"
                )
                st.session_state["error_cumplimientos_google"] = ""
                return normalizado

            errores.append(f"{metodo}: el CSV fue leído, pero no se reconoció la matriz.")

        except Exception as error:
            errores.append(f"{metodo}: {type(error).__name__}: {error}")
            print(
                f"[SSO] No fue posible leer Cumplimientos SSO mediante {metodo} "
                f"desde {url}: {error}"
            )

    st.session_state["error_cumplimientos_google"] = " | ".join(errores[-3:])
    return None


def leer_cumplimientos_desde_excel(archivo_excel, nombres_hoja):
    if not archivo_excel:
        return None
    try:
        excel = pd.ExcelFile(archivo_excel)
        hojas = {normalizar_texto(h): h for h in excel.sheet_names}
        for nombre in nombres_hoja:
            clave = normalizar_texto(nombre)
            if clave in hojas:
                raw = pd.read_excel(archivo_excel, sheet_name=hojas[clave], header=None)
                normalizado = normalizar_hoja_cumplimientos(raw)
                if normalizado is not None and not normalizado.empty:
                    return normalizado
    except Exception:
        return None
    return None


def preparar_reportabilidad(df):
    """Normaliza, limpia y valida los registros de la pestaña Reportabilidad."""
    columnas = [
        "Fecha",
        "Área",
        "Tipo_Evento",
        "Descripcion",
        "Accion_Inmediata",
        "Responsable",
        "Estado",
        "Observacion",
        "Ruta_Link",
    ]

    if df is None:
        return pd.DataFrame(columns=columnas)

    salida = normalizar_columnas_dataframe(df.copy())
    salida = asegurar_columnas(salida, columnas)

    # Elimina filas vacías que suelen quedar formateadas hacia abajo en Sheets.
    claves_registro = [
        "Fecha",
        "Área",
        "Tipo_Evento",
        "Descripcion",
        "Responsable",
        "Estado",
    ]
    mascara = pd.Series(False, index=salida.index)
    for columna in claves_registro:
        valores = salida[columna]
        mascara = mascara | (
            valores.notna()
            & valores.astype(str).str.strip().ne("")
            & valores.astype(str).str.lower().ne("nan")
        )

    salida = salida.loc[mascara].copy()

    if salida.empty:
        return pd.DataFrame(columns=columnas)

    salida["Fecha"] = salida["Fecha"].apply(convertir_fecha)

    for columna in [
        "Área",
        "Tipo_Evento",
        "Descripcion",
        "Accion_Inmediata",
        "Responsable",
        "Estado",
        "Observacion",
        "Ruta_Link",
    ]:
        salida[columna] = salida[columna].fillna("").astype(str).str.strip()

    salida = normalizar_estados(salida)
    salida = preparar_periodo(salida)
    return salida.reset_index(drop=True)

def crear_datos_ejemplo(nombre_hoja):
    if nombre_hoja == "Cumplimientos_SSO":
        filas = []
        # Respaldo completo del módulo: se consideran los 6 colaboradores
        # definidos para el seguimiento SSO. Estos registros solo se utilizan
        # cuando no es posible leer la hoja desde Google Sheets ni desde Excel.
        ejemplo = {
            "Juan Fonseca Cuevas": [
                ("Control operacional CMPC", 2),
                ("Control operacional SAIVAM", 4),
                ("OPS de Seguridad CMPC", 8),
                ("Inspecciones/Check List de Seguridad SAIVAM", 4),
                ("Observaciones de Seguridad SAIVAM", 4),
            ],
            "José Acuña Ortiz": [
                ("Control operacional CMPC", 2),
                ("Control operacional SAIVAM", 4),
                ("OPS de Seguridad CMPC", 8),
                ("Inspecciones/Check List de Seguridad SAIVAM", 4),
                ("Observaciones de Seguridad SAIVAM", 4),
            ],
            "Daniel Carrasco G.": [
                ("Control operacional CMPC", 2),
                ("Control operacional SAIVAM", 4),
                ("OPS de Seguridad CMPC", 8),
                ("Inspecciones/Check List de Seguridad SAIVAM", 4),
                ("Observaciones de Seguridad SAIVAM", 4),
            ],
            "María Araya Parra": [
                ("Control operacional CMPC", 2),
                ("Control operacional SAIVAM", 4),
                ("OPS de Seguridad CMPC", 8),
                ("Inspecciones/Check List de Seguridad SAIVAM", 4),
                ("Observaciones de Seguridad SAIVAM", 4),
            ],
            "Ricardo Grez": [
                ("OPS de Seguridad CMPC", 8),
            ],
            "Esteban Cáceres": [
                ("OPS BAPP", 4),
            ],
        }
        for observador, actividades in ejemplo.items():
            for actividad, meta in actividades:
                fila = {"Observador": observador, "Actividad": actividad}
                for mes in MESES_CORTOS:
                    fila[mes] = meta
                    fila[f"RE_{mes}"] = meta if mes in MESES_CORTOS[:7] else 0
                filas.append(fila)
        return pd.DataFrame(filas)
    if nombre_hoja == "OPS":
        return pd.DataFrame([
            {
                "Fecha": "05/07/2026",
                "Área": "Aserradero",
                "Trabajador": "Ejemplo Trabajador",
                "Supervisor": "Supervisor Turno",
                "Actividad": "Limpieza operacional",
                "Tipo_Observacion": "Conducta segura",
                "Conducta_Segura": "Uso correcto de EPP y comunicación con el equipo.",
                "Conducta_Riesgo": "",
                "Medida_Correctiva": "Mantener estándar observado.",
                "Responsable": "Supervisor",
                "Fecha_Compromiso": "10/07/2026",
                "Estado": "Cerrada",
                "Observacion": "Registro de ejemplo.",
            },
            {
                "Fecha": "08/07/2026",
                "Área": "Drymill",
                "Trabajador": "Ejemplo Trabajador 2",
                "Supervisor": "Supervisor Turno",
                "Actividad": "Retiro de material",
                "Tipo_Observacion": "Conducta de riesgo",
                "Conducta_Segura": "",
                "Conducta_Riesgo": "Ingreso al área sin verificar segregación.",
                "Medida_Correctiva": "Reforzar segregación y control de ingreso.",
                "Responsable": "Líder de área",
                "Fecha_Compromiso": "12/07/2026",
                "Estado": "Pendiente",
                "Observacion": "Registro de ejemplo.",
            },
        ])

    if nombre_hoja == "Incidentes":
        return pd.DataFrame([
            {
                "Fecha": "03/07/2026",
                "Área": "Planta Térmica",
                "Tipo_Evento": "Hallazgo",
                "Gravedad": "Media",
                "Descripcion": "Condición subestándar detectada en punto de tránsito.",
                "Accion_Inmediata": "Se informa a supervisor y se controla el área.",
                "Responsable": "Supervisor",
                "Estado": "En proceso",
                "Observacion": "Registro de ejemplo.",
            },
            {
                "Fecha": "06/07/2026",
                "Área": "Aserradero",
                "Tipo_Evento": "Cuasi accidente",
                "Gravedad": "Alta",
                "Descripcion": "Interacción entre peatón y equipo móvil.",
                "Accion_Inmediata": "Detención de tarea y charla de refuerzo.",
                "Responsable": "Prevención",
                "Estado": "Pendiente",
                "Observacion": "Registro de ejemplo.",
            },
        ])

    if nombre_hoja == "Inspecciones":
        return pd.DataFrame([
            {
                "Fecha": "04/07/2026",
                "Área": "Mantención",
                "Tipo_Inspeccion": "Bloqueo de energías",
                "Resultado": "Cumple",
                "Hallazgos": "Procedimiento aplicado correctamente.",
                "Responsable": "Supervisor Mantención",
                "Fecha_Compromiso": "",
                "Estado": "Cerrada",
                "Observacion": "Registro de ejemplo.",
            },
            {
                "Fecha": "07/07/2026",
                "Área": "Descortezado",
                "Tipo_Inspeccion": "Orden y aseo",
                "Resultado": "No cumple",
                "Hallazgos": "Material acumulado en zona operacional.",
                "Responsable": "Líder área",
                "Fecha_Compromiso": "11/07/2026",
                "Estado": "Pendiente",
                "Observacion": "Registro de ejemplo.",
            },
        ])

    if nombre_hoja == "Plan_Accion":
        return pd.DataFrame([
            {
                "Fecha": "03/07/2026",
                "Origen": "Inspección",
                "Área": "Descortezado",
                "Hallazgo": "Material acumulado en zona de trabajo.",
                "Accion_Correctiva": "Realizar limpieza y reforzar estándar de orden.",
                "Responsable": "Supervisor Turno",
                "Fecha_Compromiso": "11/07/2026",
                "Estado": "Pendiente",
                "Evidencia": "",
                "Observacion": "Registro de ejemplo.",
            },
            {
                "Fecha": "01/07/2026",
                "Origen": "OPS",
                "Área": "Aserradero",
                "Hallazgo": "Falta de señalización temporal.",
                "Accion_Correctiva": "Instalar señalética y revisar segregación.",
                "Responsable": "Prevención",
                "Fecha_Compromiso": "06/07/2026",
                "Estado": "Cerrada",
                "Evidencia": "",
                "Observacion": "Registro de ejemplo.",
            },
        ])

    if nombre_hoja == "Capacitaciones":
        return pd.DataFrame([
            {
                "Fecha": "02/07/2026",
                "Tema": "Bloqueo de energías",
                "Tipo": "Capacitación",
                "Área": "Transversal",
                "Responsable": "María Araya",
                "Vencimiento": "31/07/2026",
                "Estado": "Pendiente",
                "Observacion": "Registro de ejemplo.",
                "Evidencia": "",
            },
            {
                "Fecha": "04/08/2026",
                "Tema": "Protocolo Psicosocial",
                "Tipo": "Capacitación",
                "Área": "Transversal",
                "Responsable": "María Araya",
                "Vencimiento": "31/08/2026",
                "Estado": "En proceso",
                "Observacion": "Registro de ejemplo.",
                "Evidencia": "",
            },
        ])

    if nombre_hoja == "EPP":
        return pd.DataFrame([
            {
                "Fecha": "01/07/2026",
                "Trabajador": "Ejemplo Trabajador",
                "Cargo": "Aseador Industrial",
                "EPP": "Guantes anticorte",
                "Cantidad": 1,
                "Proxima_Reposicion": "01/08/2026",
                "Estado": "Vigente",
                "Observacion": "Registro de ejemplo.",
            },
            {
                "Fecha": "03/07/2026",
                "Trabajador": "Ejemplo Trabajador 2",
                "Cargo": "Operador equipos",
                "EPP": "Lente de seguridad",
                "Cantidad": 1,
                "Proxima_Reposicion": "03/08/2026",
                "Estado": "Vigente",
                "Observacion": "Registro de ejemplo.",
            },
        ])

    if nombre_hoja == "Protocolos_MINSAL":
        return pd.DataFrame([
            {
                "Fecha": "15/03/2026", "Protocolo": "PREXOR", "Etapa": "Identificación",
                "Área": "Aserradero", "Actividad": "Evaluación de exposición a ruido",
                "Expuestos": 14, "Responsable": "María Araya",
                "Resultado": "Área incorporada al programa", "Fecha_Compromiso": "15/04/2026",
                "Estado": "Cerrada", "Evidencia": "Drive/MINSAL/PREXOR",
                "Observacion": "Registro de ejemplo.",
            },
            {
                "Fecha": "10/04/2026", "Protocolo": "TMERT", "Etapa": "Evaluación inicial",
                "Área": "Drymill", "Actividad": "Aplicación de lista de chequeo",
                "Expuestos": 10, "Responsable": "María Araya", "Resultado": "Riesgo medio",
                "Fecha_Compromiso": "15/08/2026", "Estado": "En proceso",
                "Evidencia": "Drive/MINSAL/TMERT", "Observacion": "Registro de ejemplo.",
            },
        ])

    if nombre_hoja == "Certificaciones":
        return pd.DataFrame([
            {
                "Fecha": "10/10/2025",
                "Categoria": "Equipos",
                "Subcategoria": "Grúa horquilla",
                "Nombre_Certificacion": "Certificación Sello Verde",
                "Entidad_Emisora": "C&S Certificación",
                "Vencimiento": "10/10/2026",
                "Estado": "Vigente",
                "Dias_Para_Vencer": 89,
                "Ruta_Link": "link",
            },
            {
                "Fecha": "10/10/2025",
                "Categoria": "Equipos",
                "Subcategoria": "Barredora hombre a bordo",
                "Nombre_Certificacion": "Certificación Sello Verde",
                "Entidad_Emisora": "C&S Certificación",
                "Vencimiento": "10/10/2026",
                "Estado": "Vigente",
                "Dias_Para_Vencer": 89,
                "Ruta_Link": "link",
            },
            {
                "Fecha": "10/10/2025",
                "Categoria": "Equipos",
                "Subcategoria": "Minicargador",
                "Nombre_Certificacion": "Certificación Sello Verde",
                "Entidad_Emisora": "C&S Certificación",
                "Vencimiento": "10/10/2026",
                "Estado": "Vigente",
                "Dias_Para_Vencer": 89,
                "Ruta_Link": "link",
            },
            {
                "Fecha": "10/10/2025",
                "Categoria": "Equipos",
                "Subcategoria": "Alzahombre",
                "Nombre_Certificacion": "Certificación Sello Verde",
                "Entidad_Emisora": "C&S Certificación",
                "Vencimiento": "10/10/2026",
                "Estado": "Vigente",
                "Dias_Para_Vencer": 89,
                "Ruta_Link": "link",
            },
            {
                "Fecha": "10/10/2025",
                "Categoria": "Equipos",
                "Subcategoria": "Camión",
                "Nombre_Certificacion": "Certificación Sello Verde",
                "Entidad_Emisora": "C&S Certificación",
                "Vencimiento": "10/10/2026",
                "Estado": "Vigente",
                "Dias_Para_Vencer": 89,
                "Ruta_Link": "link",
            },
            {
                "Fecha": "28/10/2025",
                "Categoria": "Personas",
                "Subcategoria": "Nibaldo Tobar",
                "Nombre_Certificacion": "Certificación Corma",
                "Entidad_Emisora": "ICCEN E.I.R.L",
                "Vencimiento": "28/10/2026",
                "Estado": "Vigente",
                "Dias_Para_Vencer": 107,
                "Ruta_Link": "link",
            },
            {
                "Fecha": "28/10/2025",
                "Categoria": "Personas",
                "Subcategoria": "Diego Cofré",
                "Nombre_Certificacion": "Certificación Corma",
                "Entidad_Emisora": "ICCEN E.I.R.L",
                "Vencimiento": "28/10/2026",
                "Estado": "Vigente",
                "Dias_Para_Vencer": 107,
                "Ruta_Link": "link",
            },
            {
                "Fecha": "28/10/2025",
                "Categoria": "Personas",
                "Subcategoria": "Camilo Aguayo",
                "Nombre_Certificacion": "Certificación Corma",
                "Entidad_Emisora": "ICCEN E.I.R.L",
                "Vencimiento": "28/10/2026",
                "Estado": "Vigente",
                "Dias_Para_Vencer": 107,
                "Ruta_Link": "link",
            },
            {
                "Fecha": "08/08/2025",
                "Categoria": "Empresa",
                "Subcategoria": "Comité Paritario",
                "Nombre_Certificacion": "Certificación categoría Oro",
                "Entidad_Emisora": "Mutual de Seguridad",
                "Vencimiento": "08/08/2026",
                "Estado": "Por vencer",
                "Dias_Para_Vencer": 26,
                "Ruta_Link": "link",
            },
            {
                "Fecha": "18/06/2026",
                "Categoria": "Empresa",
                "Subcategoria": "Sistema de Gestión SSO",
                "Nombre_Certificacion": "Mutual de Seguridad",
                "Entidad_Emisora": "Mutual de Seguridad",
                "Vencimiento": "18/06/2027",
                "Estado": "Vigente",
                "Dias_Para_Vencer": 340,
                "Ruta_Link": "link",
            },
        ])

    if nombre_hoja == "Programa_Anual":
        return pd.DataFrame([
            {
                "Mes": "Enero",
                "Eje_Trabajo": "Capacitaciones",
                "Actividad": "Capacitación Programa de Seguridad SAIVAM",
                "Tipo_Actividad": "Capacitación",
                "Fecha_Programada": "20/01/2026",
                "Fecha_Realizacion": "20/01/2026",
                "Responsable": "María Araya",
                "Estado": "Cerrada",
                "Evidencia": "Carpeta/PRG_SSO_2026/Enero",
                "Observacion": "Actividad ejecutada según programa.",
            },
            {
                "Mes": "Agosto",
                "Eje_Trabajo": "Procedimientos e instructivos",
                "Actividad": "Revisión de procedimientos críticos",
                "Tipo_Actividad": "Auditoría",
                "Fecha_Programada": "18/08/2026",
                "Fecha_Realizacion": "",
                "Responsable": "María Araya",
                "Estado": "Pendiente",
                "Evidencia": "",
                "Observacion": "Actividad programada.",
            },
        ])

    if nombre_hoja == "Reconocimientos":
        return pd.DataFrame([
            {
                "Fecha": "31/01/2026",
                "Trabajador": "María Araya P.",
                "Cargo": "Ingeniera en Prevención de Riesgos",
                "Motivo": "Gestión SSO 2025",
                "Periodo": "Enero 2026",
                "Estado": "Cerrada",
                "Evidencia": "",
                "Observacion": "Reconocimiento de seguridad CMPC, 2025.",
            },
            {
                "Fecha": "31/01/2026",
                "Trabajador": "Pedro Quezada L.",
                "Cargo": "Aseador Industrial Mantenedor",
                "Motivo": "Compromiso con la seguridad",
                "Periodo": "Enero 2026",
                "Estado": "Cerrada",
                "Evidencia": "",
                "Observacion": "Reconocimiento mensual correspondiente a enero de 2026.",
            },
            {
                "Fecha": "28/02/2026",
                "Trabajador": "Esteban Cáceres J.",
                "Cargo": "Aseador Industrial y Op. Riego",
                "Motivo": "Compromiso con la seguridad",
                "Periodo": "Febrero 2026",
                "Estado": "Cerrada",
                "Evidencia": "",
                "Observacion": "Reconocimiento mensual correspondiente a febrero de 2026.",
            },
            {
                "Fecha": "31/03/2026",
                "Trabajador": "Héctor Flores F.",
                "Cargo": "Aseador Industrial Mantenedor",
                "Motivo": "Compromiso con la seguridad",
                "Periodo": "Marzo 2026",
                "Estado": "Cerrada",
                "Evidencia": "",
                "Observacion": "Reconocimiento mensual correspondiente a marzo de 2026.",
            },
            {
                "Fecha": "30/04/2026",
                "Trabajador": "Omar Acevedo S.",
                "Cargo": "Aseador Industrial y Jardinero",
                "Motivo": "Compromiso con la seguridad",
                "Periodo": "Abril 2026",
                "Estado": "Cerrada",
                "Evidencia": "",
                "Observacion": "Reconocimiento mensual correspondiente a abril de 2026.",
            },
            {
                "Fecha": "31/05/2026",
                "Trabajador": "Pedro Morales C.",
                "Cargo": "Operador Riego",
                "Motivo": "Compromiso con la seguridad",
                "Periodo": "Mayo 2026",
                "Estado": "Cerrada",
                "Evidencia": "",
                "Observacion": "Reconocimiento mensual correspondiente a mayo de 2026.",
            },
            {
                "Fecha": "30/06/2026",
                "Trabajador": "Manuel Mardones B.",
                "Cargo": "Aseador Industrial Mantención",
                "Motivo": "Compromiso con la seguridad",
                "Periodo": "Junio 2026",
                "Estado": "Cerrada",
                "Evidencia": "",
                "Observacion": "Reconocimiento mensual correspondiente a junio de 2026.",
            },
            {
                "Fecha": "31/07/2026",
                "Trabajador": "Isaac Melgarejo R.",
                "Cargo": "Aseador Industrial",
                "Motivo": "Compromiso con la seguridad",
                "Periodo": "Julio 2026",
                "Estado": "Cerrada",
                "Evidencia": "",
                "Observacion": "Reconocimiento mensual correspondiente a julio de 2026.",
            },
        ])

    if nombre_hoja == "Comite_Paritario":
        return pd.DataFrame([
            {
                "Fecha": "12/06/2026",
                "Tipo_Reunion": "Reunión ordinaria",
                "Área": "Aserradero",
                "Tema": "Revisión de observaciones preventivas",
                "Acuerdo": "Reforzar rutas peatonales y segregación.",
                "Responsable": "Supervisor Aserradero",
                "Fecha_Compromiso": "25/06/2026",
                "Estado": "Cerrada",
                "Evidencia": "Carpeta/Comite_Paritario/Junio",
                "Observacion": "Acuerdo verificado en terreno.",
            },
            {
                "Fecha": "10/07/2026",
                "Tipo_Reunion": "Reunión ordinaria",
                "Área": "Mantención",
                "Tema": "Seguimiento de acciones correctivas",
                "Acuerdo": "Cerrar acciones vencidas y adjuntar evidencia.",
                "Responsable": "María Araya",
                "Fecha_Compromiso": "18/07/2026",
                "Estado": "En proceso",
                "Evidencia": "",
                "Observacion": "Seguimiento programado.",
            },
        ])

    if nombre_hoja == "Trabajos_Criticos":
        return pd.DataFrame([
            {
                "Fecha": "06/07/2026",
                "Área": "Mantención",
                "Tipo_Trabajo": "Bloqueo de energías",
                "Actividad": "Intervención de equipo detenido",
                "Responsable": "Supervisor Mantención",
                "Permiso": "Sí",
                "Estado": "Cerrada",
                "Observacion": "Registro de ejemplo.",
            },
            {
                "Fecha": "08/07/2026",
                "Área": "Planta Térmica",
                "Tipo_Trabajo": "Espacio confinado",
                "Actividad": "Limpieza interior",
                "Responsable": "Supervisor Turno",
                "Permiso": "Sí",
                "Estado": "En proceso",
                "Observacion": "Registro de ejemplo.",
            },
        ])

    if nombre_hoja == "Documentos":
        return pd.DataFrame([
            {
                "Tipo_Documento": "Procedimiento",
                "Nombre_Documento": "Procedimiento de bloqueo de energías",
                "Version": "1.0",
                "Fecha": "01/07/2026",
                "Vencimiento": "01/07/2027",
                "Estado": "Vigente",
                "Ruta_Link": "",
                "Observacion": "Registro de ejemplo.",
            },
            {
                "Tipo_Documento": "Matriz",
                "Nombre_Documento": "Matriz de riesgos operacionales",
                "Version": "1.0",
                "Fecha": "01/07/2026",
                "Vencimiento": "01/07/2027",
                "Estado": "Vigente",
                "Ruta_Link": "",
                "Observacion": "Registro de ejemplo.",
            },
        ])

    if nombre_hoja == "Configuracion":
        return pd.DataFrame([
            {"Parametro": "Inicio_Sin_Accidentes", "Valor": FECHA_INICIO_SIN_ACCIDENTES_DEFAULT},
            {"Parametro": "Meta_OPS_Mensual", "Valor": 30},
            {"Parametro": "Meta_Inspecciones_Mensual", "Valor": 12},
            {"Parametro": "Meta_Capacitaciones_Mensual", "Valor": 4},
        ])

    return pd.DataFrame(columns=SHEETS[nombre_hoja]["columnas"])


def leer_hoja_desde_excel(archivo_excel, nombres_hoja):
    if not archivo_excel:
        return None
    try:
        excel = pd.ExcelFile(archivo_excel)
        hojas_disponibles = {normalizar_texto(h): h for h in excel.sheet_names}
        for nombre in nombres_hoja:
            clave = normalizar_texto(nombre)
            if clave in hojas_disponibles:
                return pd.read_excel(archivo_excel, sheet_name=hojas_disponibles[clave])
    except Exception:
        return None
    return None


@st.cache_data(ttl=60)
def cargar_datos():
    """
    Carga todas las pestañas desde un único Google Sheet.

    Orden de prioridad:
    1. Google Sheets.
    2. Excel local como respaldo.
    3. Datos de ejemplo si no existe ninguna de las fuentes anteriores.
    """
    archivo_excel = buscar_archivo_excel()
    datos = {}
    fuentes = {}

    for nombre_hoja, config in SHEETS.items():
        if nombre_hoja == "Incidentes":
            df = leer_hoja_desde_google(
                config["nombres"],
                config["columnas"],
            )
            fuente = "Google Sheets"

            if df is None or df.empty:
                df = leer_hoja_desde_excel(
                    archivo_excel,
                    config["nombres"],
                )
                fuente = "Excel local"

            if df is None or df.empty:
                # Reportabilidad no utiliza registros ficticios. De esta forma
                # los indicadores nunca muestran información de ejemplo como si
                # correspondiera a la operación real.
                df = pd.DataFrame(columns=config["columnas"])
                fuente = "Sin datos"

            df = preparar_reportabilidad(df)
            datos[nombre_hoja] = df
            fuentes[nombre_hoja] = fuente
            continue

        if nombre_hoja == "Cumplimientos_SSO":
            df = leer_cumplimientos_desde_google(config["nombres"])
            fuente = "Google Sheets"

            if df is None or df.empty:
                df = leer_cumplimientos_desde_excel(
                    archivo_excel,
                    config["nombres"],
                )
                fuente = "Excel local"

            if df is None or df.empty:
                # No se usan datos ficticios en Cumplimientos SSO. Si Google
                # Sheets y el Excel local no responden, el módulo queda vacío
                # y muestra una advertencia, evitando porcentajes incorrectos.
                df = pd.DataFrame(columns=config["columnas"])
                fuente = "Sin datos"

            df = normalizar_hoja_cumplimientos(df)
            df = asegurar_columnas(df, config["columnas"])
            datos[nombre_hoja] = df
            fuentes[nombre_hoja] = fuente
            continue

        df = leer_hoja_desde_google(
            config["nombres"],
            config["columnas"],
        )
        fuente = "Google Sheets"

        if df is None or df.empty:
            df = leer_hoja_desde_excel(
                archivo_excel,
                config["nombres"],
            )
            fuente = "Excel local"

        if df is None or df.empty:
            df = crear_datos_ejemplo(nombre_hoja)
            fuente = "Datos de ejemplo"

        # La pestaña PRG_SSO_2026 se procesa por separado para conservar el
        # campo Mes y respetar el Estado informado directamente en Google Sheets.
        if nombre_hoja == "Programa_Anual":
            df = preparar_programa_anual(df)
            datos[nombre_hoja] = df
            fuentes[nombre_hoja] = fuente
            continue

        df = normalizar_columnas_dataframe(df)
        df = asegurar_columnas(df, config["columnas"])
        df = preparar_fechas(df)

        if nombre_hoja != "Configuracion":
            df = preparar_periodo(df)

        if "Estado" in df.columns:
            df = normalizar_estados(df)

        if nombre_hoja in [
            "OPS",
            "Inspecciones",
            "Plan_Accion",
            "Comite_Paritario",
            "Protocolos_MINSAL",
        ]:
            df = marcar_vencimientos(
                df,
                "Fecha_Compromiso",
            )

        # En Capacitaciones el estado se conserva tal como fue informado en
        # Google Sheets; no se recalcula automáticamente según Vencimiento.
        if nombre_hoja == "Certificaciones":
            df = preparar_certificaciones(df)

        datos[nombre_hoja] = df
        fuentes[nombre_hoja] = fuente

    return datos, archivo_excel, fuentes


def valor_config(configuracion, parametro, default=""):
    if configuracion is None or configuracion.empty:
        return default
    if "Parametro" not in configuracion.columns or "Valor" not in configuracion.columns:
        return default
    buscar = normalizar_texto(parametro)
    aux = configuracion.copy()
    aux["_param"] = aux["Parametro"].apply(normalizar_texto)
    fila = aux[aux["_param"] == buscar]
    if fila.empty:
        return default
    return fila.iloc[0]["Valor"]


def dias_sin_accidentes(configuracion, incidentes):
    fecha_inicio = convertir_fecha(valor_config(configuracion, "Inicio_Sin_Accidentes", FECHA_INICIO_SIN_ACCIDENTES_DEFAULT))

    if incidentes is not None and not incidentes.empty and "Tipo_Evento" in incidentes.columns:
        eventos = incidentes.copy()
        eventos["Fecha"] = eventos["Fecha"].apply(convertir_fecha)
        eventos["_tipo"] = eventos["Tipo_Evento"].apply(normalizar_texto)
        accidentes = eventos[eventos["_tipo"].str.contains("accidente", na=False)]
        accidentes = accidentes[~accidentes["_tipo"].str.contains("cuasi", na=False)]
        if not accidentes.empty:
            ultima_fecha = accidentes["Fecha"].max()
            if pd.notna(ultima_fecha):
                fecha_inicio = max(fecha_inicio, ultima_fecha)

    if pd.isna(fecha_inicio):
        fecha_inicio = convertir_fecha(FECHA_INICIO_SIN_ACCIDENTES_DEFAULT)

    return max(0, int((HOY - fecha_inicio.normalize()).days)), fecha_inicio


def aplicar_filtros(df, filtro_area, filtro_anio, filtro_mes):
    if df is None or df.empty:
        return df
    salida = df.copy()
    if filtro_area != "Todas las áreas" and "Área" in salida.columns:
        salida = salida[salida["Área"].astype(str) == filtro_area]
    if filtro_anio != "Todos" and "Año" in salida.columns:
        salida = salida[salida["Año"] == filtro_anio]
    if filtro_mes != "Todos" and "Mes" in salida.columns:
        salida = salida[salida["Mes"] == filtro_mes]
    return salida.copy()


def tabla_limpia(
    df,
    columnas=None,
    height=460,
    centrar_todo=False,
    modo_ultracompacto=False,
):
    """
    Muestra cualquier planilla del sistema como una tabla HTML compacta.

    El ancho de cada columna se calcula automáticamente según el tipo de
    información. De esta forma, todas las columnas quedan dentro del ancho
    disponible de la página principal, incluso en módulos con 10 a 12 campos.
    """
    if df is None or df.empty:
        st.info("Sin registros para mostrar.")
        return

    mostrar = df.copy()

    if columnas:
        columnas = [columna for columna in columnas if columna in mostrar.columns]
        mostrar = mostrar[columnas]

    # Formato uniforme de fechas.
    for columna in [
        "Fecha",
        "Fecha_Compromiso",
        "Vencimiento",
        "Proxima_Reposicion",
        "Fecha_Programada",
        "Fecha_Realizacion",
    ]:
        if columna in mostrar.columns:
            mostrar[columna] = mostrar[columna].apply(fecha_texto)

    # Encabezados visibles.
    mostrar = mostrar.rename(
        columns={
            "Tipo_Observacion": "Tipo observación",
            "Conducta_Segura": "Conducta segura",
            "Conducta_Riesgo": "Conducta de riesgo",
            "Medida_Correctiva": "Medida correctiva",
            "Fecha_Compromiso": "Fecha compromiso",
            "Tipo_Evento": "Tipo evento",
            "Accion_Inmediata": "Acción inmediata",
            "Tipo_Inspeccion": "Tipo inspección",
            "Accion_Correctiva": "Acción correctiva",
            "Proxima_Reposicion": "Próxima reposición",
            "Tipo_Actividad": "Tipo de actividad",
            "Eje_Trabajo": "Eje de trabajo",
            "Fecha_Programada": "Fecha programada",
            "Fecha_Realizacion": "Fecha realización",
            "Tipo_Reconocimiento": "Tipo de reconocimiento",
            "Tipo_Reunion": "Tipo de reunión",
            "Tipo_Trabajo": "Tipo trabajo",
            "Tipo_Documento": "Tipo documento",
            "Nombre_Documento": "Nombre documento",
            "Nombre_Certificacion": "Nombre certificación",
            "Entidad_Emisora": "Entidad emisora",
            "Numero_Certificado": "N.º certificado",
            "Dias_Para_Vencer": "Días para vencer",
            "Ruta_Link": "Ruta / enlace",
            "Observacion": "Observación",
            "Descripcion": "Descripción",
            "Detalle": "Descripción",
            "Categoria": "Categoría",
            "Subcategoria": "Subcategoría",
            "Protocolo": "Protocolo MINSAL",
            "Titular_Activo": "Titular / activo",
        }
    )

    cantidad_columnas = max(1, len(mostrar.columns))

    # El tamaño de letra disminuye gradualmente cuando hay más columnas.
    if cantidad_columnas <= 7:
        tamano_letra = "11px"
        padding_celda = "5px 6px"
    elif cantidad_columnas <= 9:
        tamano_letra = "10px"
        padding_celda = "4px 5px"
    elif cantidad_columnas <= 11:
        tamano_letra = "9px"
        padding_celda = "3px 4px"
    else:
        tamano_letra = "8.2px"
        padding_celda = "3px 3px"

    # Formato especial para tablas con muchas columnas, como el
    # cumplimiento anual por observador.
    if modo_ultracompacto:
        tamano_letra = "8.6px"
        padding_celda = "2px 2px"

    columnas_cortas = {
        "fecha",
        "vencimiento",
        "fecha_compromiso",
        "fecha_programada",
        "fecha_realizacion",
        "estado",
        "area",
        "mes",
        "meta",
        "resultado",
        "cumplimiento",
        "asistentes",
        "expuestos",
        "permiso",
        "version",
        "gravedad",
        "periodo",
        "origen",
        "tipo",
    }

    columnas_medias = {
        "trabajador",
        "supervisor",
        "responsable",
        "relator",
        "cargo",
        "tipo_evento",
        "tipo_observacion",
        "tipo_inspeccion",
        "tipo_actividad",
        "eje_de_trabajo",
        "tipo_reunion",
        "tipo_trabajo",
        "tipo_documento",
        "protocolo_minsal",
        "etapa",
        "registro",
        "categoria",
        "subcategoria",
        "ruta_enlace",
        "evidencia",
    }

    columnas_largas = {
        "actividad",
        "tema",
        "descripcion",
        "conducta_segura",
        "conducta_de_riesgo",
        "medida_correctiva",
        "accion_inmediata",
        "hallazgos",
        "hallazgo",
        "accion_correctiva",
        "observacion",
        "acuerdo",
        "motivo",
        "nombre_documento",
        "nombre_certificacion",
    }

    pesos = []

    meses_columnas = {
        "enero", "febrero", "marzo", "abril", "mayo", "junio",
        "julio", "agosto", "septiembre", "octubre",
        "noviembre", "diciembre",
    }

    for columna in mostrar.columns:
        clave = normalizar_texto(columna)

        if clave in meses_columnas:
            peso = 1.08
        elif clave in columnas_largas:
            peso = 2.05
        elif clave in columnas_medias:
            peso = 1.25
        elif clave in columnas_cortas:
            peso = 0.82
        else:
            peso = 1.0

        # Pequeño ajuste según la extensión real del contenido.
        valores = mostrar[columna].dropna().astype(str).head(80)
        longitud_media = valores.str.len().mean() if not valores.empty else 0

        if longitud_media >= 45:
            peso += 0.45
        elif longitud_media >= 25:
            peso += 0.25
        elif longitud_media <= 8:
            peso -= 0.08

        if modo_ultracompacto:
            # Reduce espacios innecesarios en la tabla anual.
            if clave == "observador":
                peso = 1.20
            elif clave == "tipo_observacion":
                peso = 1.25
            elif clave in {
                "ene", "feb", "mar", "abr", "may", "jun",
                "jul", "ago", "sep", "oct", "nov", "dic",
            }:
                peso = 0.68
            elif clave in {
                "real_ano", "teorica_ano", "avance",
                "real_a_la_fecha", "meta_a_la_fecha",
                "a_la_fecha",
            }:
                peso = 0.92

        pesos.append(max(0.54, peso))

    suma_pesos = sum(pesos)
    anchos = [(peso / suma_pesos) * 100 for peso in pesos]

    colgroup_html = "<colgroup>" + "".join(
        f'<col style="width:{ancho:.3f}%">'
        for ancho in anchos
    ) + "</colgroup>"

    encabezados_html = "".join(
        f"<th>{escape_html(columna)}</th>"
        for columna in mostrar.columns
    )

    columnas_centradas = {
        "fecha",
        "vencimiento",
        "fecha_compromiso",
        "fecha_programada",
        "fecha_realizacion",
        "estado",
        "meta",
        "resultado",
        "cumplimiento",
        "asistentes",
        "expuestos",
        "permiso",
        "version",
        "gravedad",
        "dias_para_vencer",
    }

    def valor_texto(valor):
        if pd.isna(valor):
            return ""

        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))

        return str(valor)

    def celda_html(valor, columna):
        texto = valor_texto(valor).strip()
        clave = normalizar_texto(columna)

        clase = (
            "tabla-celda-centrada"
            if centrar_todo or clave in columnas_centradas
            else ""
        )

        es_columna_enlace = clave in {
            "ruta_link",
            "ruta_enlace",
            "respaldo",
            "evidencia",
            "link",
            "enlace",
        }

        es_url = texto.lower().startswith(("http://", "https://"))

        if es_url:
            etiqueta = "📄 Abrir respaldo" if es_columna_enlace else "🔗 Abrir"

            contenido = (
                f'<a class="tabla-link-boton" '
                f'href="{escape_html(texto)}" '
                f'target="_blank" '
                f'rel="noopener noreferrer">'
                f'{escape_html(etiqueta)}'
                f'</a>'
            )
        elif es_columna_enlace and texto:
            # Si todavía existe texto como "link" o "enlace", se muestra como
            # texto normal porque no contiene una URL válida.
            contenido = escape_html(texto)
        else:
            contenido = escape_html(texto)

        return f'<td class="{clase}">{contenido}</td>'

    filas_html = []

    for _, fila in mostrar.iterrows():
        celdas = "".join(
            celda_html(fila[columna], columna)
            for columna in mostrar.columns
        )
        filas_html.append(f"<tr>{celdas}</tr>")

    st.markdown(
        f"""
<style>
.tabla-general-wrap {{
    width: 100%;
    max-width: 100%;
    overflow-x: hidden;
    border: 1px solid rgba(30, 180, 120, .42);
    border-radius: 13px;
    background: rgba(8, 13, 17, .94);
}}

.tabla-general-compacta {{
    width: 100%;
    max-width: 100%;
    border-collapse: collapse;
    table-layout: fixed;
    font-size: {tamano_letra};
    line-height: 1.08;
}}

.tabla-general-compacta th,
.tabla-general-compacta td {{
    box-sizing: border-box;
    padding: {padding_celda} !important;
    border-right: 1px solid rgba(110, 125, 140, .18);
    border-bottom: 1px solid rgba(110, 125, 140, .18);
    text-align: left;
    vertical-align: middle;
}}

.tabla-general-compacta th {{
    min-height: 25px;
    background: #1b2029;
    color: #b9bec8;
    font-weight: 700;
    white-space: nowrap !important;
    word-break: keep-all !important;
    overflow-wrap: normal !important;
    text-align: center;
}}

.tabla-general-compacta td {{
    min-height: 23px;
    color: #f3f7f5;
    white-space: normal !important;
    overflow-wrap: anywhere;
    word-break: normal;
}}

.tabla-general-compacta.tabla-ultracompacta th,
.tabla-general-compacta.tabla-ultracompacta td {{
    padding: 2px 2px !important;
    line-height: 1.02 !important;
}}

.tabla-general-compacta.tabla-ultracompacta td {{
    height: 22px;
    text-align: center !important;
}}

.tabla-general-compacta tr:last-child td {{
    border-bottom: none;
}}

.tabla-general-compacta th:last-child,
.tabla-general-compacta td:last-child {{
    border-right: none;
}}

.tabla-general-compacta tbody tr:hover td {{
    background: rgba(31, 197, 133, .055);
}}

.tabla-general-compacta .tabla-celda-centrada {{
    text-align: center;
}}

.tabla-general-compacta a {{
    color: #8EE7BE;
    font-weight: 700;
    text-decoration: none;
}}

.tabla-general-compacta .tabla-link-boton {{
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 4px;
    padding: 4px 8px;
    border: 1px solid rgba(52, 211, 153, .55);
    border-radius: 8px;
    background: rgba(16, 185, 129, .12);
    color: #A7F3D0 !important;
    font-size: 9px;
    font-weight: 800;
    white-space: nowrap;
    transition: all .16s ease;
}}

.tabla-general-compacta .tabla-link-boton:hover {{
    background: rgba(16, 185, 129, .24);
    border-color: rgba(110, 231, 183, .85);
    color: #ECFDF5 !important;
    transform: translateY(-1px);
}}
</style>
        """,
        unsafe_allow_html=True,
    )

    clase_tabla = (
        "tabla-general-compacta tabla-ultracompacta notranslate"
        if modo_ultracompacto
        else "tabla-general-compacta notranslate"
    )

    st.markdown(
        (
            '<div class="tabla-general-wrap notranslate" translate="no">'
            f'<table class="{clase_tabla}" translate="no">'
            f'{colgroup_html}'
            f'<thead><tr>{encabezados_html}</tr></thead>'
            f'<tbody>{"".join(filas_html)}</tbody>'
            '</table>'
            '</div>'
        ),
        unsafe_allow_html=True,
    )


# =========================================================
# ESTILO VISUAL — FORMATO TIPO PANEL EQUIPOS MÓVILES
# =========================================================

def obtener_data_uri_recurso(rutas):
    """Convierte un recurso local en data URI para usarlo en CSS/HTML."""
    for ruta in rutas:
        if ruta and os.path.exists(ruta) and os.path.isfile(ruta):
            try:
                mime = mimetypes.guess_type(ruta)[0] or "image/png"
                with open(ruta, "rb") as archivo:
                    contenido = base64.b64encode(archivo.read()).decode("utf-8")
                return f"data:{mime};base64,{contenido}"
            except Exception:
                continue
    return ""


def buscar_imagen_local(nombre_base):
    """Busca una imagen dentro de la carpeta de la aplicación y sus subcarpetas."""
    extensiones = [".png", ".jpg", ".jpeg", ".webp"]
    carpetas = ["", "fotos", "imagenes", "images", "img", "assets"]

    for carpeta in carpetas:
        for extension in extensiones:
            ruta = ruta_app(carpeta, f"{nombre_base}{extension}") if carpeta else ruta_app(f"{nombre_base}{extension}")
            if os.path.isfile(ruta):
                return ruta

    # Búsqueda adicional sin distinguir mayúsculas/minúsculas.
    objetivo = normalizar_texto(nombre_base)
    patrones = []
    for extension in extensiones:
        patrones.extend([
            ruta_app(f"*{extension}"),
            ruta_app("*", f"*{extension}"),
        ])

    for patron in patrones:
        for ruta in glob.glob(patron):
            base = os.path.splitext(os.path.basename(ruta))[0]
            if normalizar_texto(base) == objetivo and os.path.isfile(ruta):
                return ruta

    return ""


def obtener_imagen_html(nombre_base, clase="brand-logo-img", alt="SAIVAM"):
    ruta = buscar_imagen_local(nombre_base)
    imagen_uri = obtener_data_uri_recurso([ruta]) if ruta else ""
    if imagen_uri:
        return f'<img class="{clase}" src="{imagen_uri}" alt="{escape_html(alt)}">'
    return ""


def obtener_logo_respaldo_html(clase="brand-logo-img"):
    logo_uri = obtener_data_uri_recurso([
        ruta_app("logo_saivam.png"),
        ruta_app("saivam_logo.png"),
        ruta_app("SAIVAM.png"),
        ruta_app("saivam.png"),
        ruta_app("fotos", "logo_saivam.png"),
        ruta_app("fotos", "saivam.png"),
    ])
    if logo_uri:
        return f'<img class="{clase}" src="{logo_uri}" alt="SAIVAM">'
    return (
        '<div class="brand-fallback">'
        '<span class="brand-mark">S</span><span class="brand-name">SAIVAM</span>'
        '</div>'
    )


def obtener_logo_sidebar_html():
    # Imagen solicitada para el encabezado del menú lateral.
    return obtener_imagen_html("logoredondo", "brand-logo-img", "Logo redondo SAIVAM") or obtener_logo_respaldo_html()


def obtener_logo_principal_html():
    # Imagen solicitada para la esquina superior derecha del panel.
    return obtener_imagen_html("logo1", "brand-logo-img", "Logo SAIVAM") or obtener_logo_respaldo_html()



def obtener_carpetas_reconocimientos():
    # Localiza la carpeta aunque cambie mayúsculas o use plural.
    carpeta_static = ruta_app("static")
    if not os.path.isdir(carpeta_static):
        return []

    nombres_validos = {
        "reconocimiento",
        "reconocimientos",
        "recognition",
        "recognitions",
    }

    carpetas = []
    for raiz, subcarpetas, _ in os.walk(carpeta_static):
        for subcarpeta in subcarpetas:
            if normalizar_texto(subcarpeta) in nombres_validos:
                ruta = os.path.join(raiz, subcarpeta)
                if ruta not in carpetas:
                    carpetas.append(ruta)

    for ruta in [
        ruta_app("static", "reconocimiento"),
        ruta_app("static", "reconocimientos"),
    ]:
        if os.path.isdir(ruta) and ruta not in carpetas:
            carpetas.append(ruta)

    return carpetas


def listar_fotos_reconocimientos():
    # Lista todas las imágenes disponibles dentro de la carpeta del módulo.
    extensiones_validas = {
        ".png", ".jpg", ".jpeg", ".webp", ".jfif",
        ".bmp", ".gif", ".tif", ".tiff", ".avif",
    }

    archivos = []
    for carpeta in obtener_carpetas_reconocimientos():
        for raiz, _, nombres in os.walk(carpeta):
            for nombre in nombres:
                ruta = os.path.join(raiz, nombre)
                extension = os.path.splitext(nombre)[1].lower()
                if os.path.isfile(ruta) and extension in extensiones_validas:
                    archivos.append(ruta)

    return sorted(set(archivos))


def buscar_foto_reconocimiento(*nombres_base):
    # Busca una fotografía por uno o más nombres posibles.
    archivos = listar_fotos_reconocimientos()
    objetivos = [
        normalizar_texto(nombre)
        for nombre in nombres_base
        if str(nombre).strip()
    ]

    # Coincidencia exacta.
    for ruta in archivos:
        base = normalizar_texto(os.path.splitext(os.path.basename(ruta))[0])
        if base in objetivos:
            return ruta

    # Coincidencia flexible para sufijos añadidos por Windows.
    for ruta in archivos:
        base = normalizar_texto(os.path.splitext(os.path.basename(ruta))[0])
        for objetivo in objetivos:
            if base.startswith(objetivo) or objetivo.startswith(base):
                return ruta

    return ""




def obtener_fotos_reconocimientos():
    # Fotografías visibles en la galería de reconocimientos, sin duplicados.
    # Se aceptan variantes de escritura para facilitar la carga desde la carpeta.
    configuracion = [
        (
            ("claudioa", "clauidoa", "claudio", "claudia", "claudiaa"),
            "Claudioa",
        ),
        (
            ("mariaa", "maria", "maria_araya"),
            "Mariaa",
        ),
        (
            ("ricardog", "ricardo_g", "ricardo"),
            "Ricardog",
        ),
        (
            ("saivam500", "saivam_500"),
            "Saivam500",
        ),
        (
            ("saivam700", "saivam_700"),
            "Saivam700",
        ),
        (
            ("tresreconocidos", "tres_reconocidos"),
            "Tres reconocidos",
        ),
        (
            ("1200d", "1200_d", "1200dias", "1200_dias"),
            "1200 días sin accidentes",
        ),
        (
            ("pedroa", "pedro_a", "pedro"),
            "Pedro",
        ),
        (
            ("hectora", "hector_a", "hector"),
            "Héctor",
        ),
        (
            ("estebana", "esteban_a", "esteban"),
            "Esteban",
        ),
    ]

    fotos = []
    rutas_usadas = set()

    for nombres_posibles, titulo in configuracion:
        ruta = buscar_foto_reconocimiento(*nombres_posibles)

        if ruta and ruta not in rutas_usadas:
            fotos.append({
                "ruta": ruta,
                "titulo": titulo,
                "archivo": os.path.basename(ruta),
            })
            rutas_usadas.add(ruta)

    return fotos


def mostrar_fotos_reconocimientos():
    # Muestra una sola fotografía y los controles debajo de la imagen.
    fotos = obtener_fotos_reconocimientos()

    panel_titulo("Galería de Reconocimientos")

    if not fotos:
        st.info(
            "No se encontraron fotografías en `static/reconocimiento/`. "
            "Verifica que los archivos configurados estén dentro de esa carpeta."
        )
        return

    clave_indice = "indice_foto_reconocimiento"

    if clave_indice not in st.session_state:
        st.session_state[clave_indice] = 0

    if st.session_state[clave_indice] >= len(fotos):
        st.session_state[clave_indice] = 0

    indice_actual = st.session_state[clave_indice]
    foto = fotos[indice_actual]
    imagen_uri = obtener_data_uri_recurso([foto["ruta"]])

    # Fotografía sin nombre o texto inferior.
    if imagen_uri:
        html_foto = (
            '<div class="recognition-carousel-wrapper">'
            '<div class="recognition-photo-card">'
            f'<img class="recognition-photo-img" src="{imagen_uri}" '
            f'alt="Fotografía de reconocimiento">'
            '</div>'
            '</div>'
        )
        st.markdown(html_foto, unsafe_allow_html=True)
    else:
        st.image(
            foto["ruta"],
            width="stretch",
        )

    # Botones ubicados debajo de la fotografía.
    espacio_izq, boton_anterior, indicador, boton_siguiente, espacio_der = st.columns(
        [1.15, 1, 0.75, 1, 1.15]
    )

    with boton_anterior:
        if st.button(
            "⬅️ Anterior",
            key="foto_reconocimiento_anterior",
            use_container_width=True,
        ):
            st.session_state[clave_indice] = (
                st.session_state[clave_indice] - 1
            ) % len(fotos)
            st.rerun()

    with indicador:
        st.markdown(
            (
                '<div class="recognition-carousel-counter">'
                f'{indice_actual + 1} / {len(fotos)}'
                '</div>'
            ),
            unsafe_allow_html=True,
        )

    with boton_siguiente:
        if st.button(
            "Siguiente ➡️",
            key="foto_reconocimiento_siguiente",
            use_container_width=True,
        ):
            st.session_state[clave_indice] = (
                st.session_state[clave_indice] + 1
            ) % len(fotos)
            st.rerun()





def obtener_carpetas_certificaciones():
    # Localiza la carpeta de imágenes del módulo Certificaciones.
    carpetas = []

    rutas_directas = [
        ruta_app("static", "certificaciones"),
        ruta_app("static", "certificacion"),
        ruta_app("static-certificaciones"),
        ruta_app("static_certificaciones"),
    ]

    for ruta in rutas_directas:
        if os.path.isdir(ruta) and ruta not in carpetas:
            carpetas.append(ruta)

    carpeta_static = ruta_app("static")
    nombres_validos = {
        "certificacion",
        "certificaciones",
        "certification",
        "certifications",
    }

    if os.path.isdir(carpeta_static):
        for raiz, subcarpetas, _ in os.walk(carpeta_static):
            for subcarpeta in subcarpetas:
                if normalizar_texto(subcarpeta) in nombres_validos:
                    ruta = os.path.join(raiz, subcarpeta)
                    if ruta not in carpetas:
                        carpetas.append(ruta)

    return carpetas


def listar_fotos_certificaciones():
    extensiones_validas = {
        ".png", ".jpg", ".jpeg", ".webp", ".jfif",
        ".bmp", ".gif", ".tif", ".tiff", ".avif",
    }

    archivos = []

    for carpeta in obtener_carpetas_certificaciones():
        for raiz, _, nombres in os.walk(carpeta):
            for nombre in nombres:
                ruta = os.path.join(raiz, nombre)
                extension = os.path.splitext(nombre)[1].lower()

                if os.path.isfile(ruta) and extension in extensiones_validas:
                    archivos.append(ruta)

    return sorted(set(archivos))


def buscar_foto_certificacion(*nombres_base):
    archivos = listar_fotos_certificaciones()
    objetivos = {
        normalizar_texto(nombre)
        for nombre in nombres_base
        if str(nombre).strip()
    }

    # Coincidencia exacta.
    for ruta in archivos:
        nombre = os.path.splitext(os.path.basename(ruta))[0]
        if normalizar_texto(nombre) in objetivos:
            return ruta

    # Coincidencia flexible para nombres con sufijos.
    for ruta in archivos:
        nombre = normalizar_texto(
            os.path.splitext(os.path.basename(ruta))[0]
        )

        for objetivo in objetivos:
            if nombre.startswith(objetivo) or objetivo.startswith(nombre):
                return ruta

    return ""


def obtener_ficha_certificacion(df, aliases):
    if df is None or df.empty or "Subcategoria" not in df.columns:
        return None

    base = df.copy()

    if "Categoria" in base.columns:
        categoria = base["Categoria"].fillna("").apply(normalizar_texto)
        base = base[categoria == "equipos"].copy()

    if base.empty:
        return None

    subcategorias = base["Subcategoria"].fillna("").apply(normalizar_texto)
    aliases_normalizados = {
        normalizar_texto(alias)
        for alias in aliases
    }

    coincidencia = subcategorias.isin(aliases_normalizados)

    if not coincidencia.any():
        coincidencia = subcategorias.apply(
            lambda valor: any(
                alias in valor or valor in alias
                for alias in aliases_normalizados
                if alias and valor
            )
        )

    filas = base.loc[coincidencia]

    if filas.empty:
        return None

    return filas.iloc[0]


def clase_estado_certificacion(estado):
    estado_normalizado = normalizar_texto(estado)

    if estado_normalizado == "vigente":
        return "cert-status-vigente"

    if estado_normalizado == "por_vencer":
        return "cert-status-por-vencer"

    if estado_normalizado == "vencida":
        return "cert-status-vencida"

    return "cert-status-sin-vencimiento"


def texto_dias_certificacion(dias, estado):
    if pd.isna(dias) or str(dias).strip() in ["", "<NA>", "nan"]:
        return "Sin fecha de vencimiento"

    try:
        dias_numero = int(float(dias))
    except (TypeError, ValueError):
        return "Vigencia no disponible"

    estado_normalizado = normalizar_texto(estado)

    if estado_normalizado == "vencida":
        return f"Vencida hace {abs(dias_numero)} días"

    if dias_numero == 0:
        return "Vence hoy"

    if dias_numero == 1:
        return "Vence en 1 día"

    return f"Vence en {dias_numero} días"


def mostrar_equipos_certificados(df):
    configuracion = [
        {
            "archivo": ("alza_hombre", "alzahombre", "alza hombre"),
            "titulo": "Alzahombre",
            "aliases": (
                "Alzahombre",
                "Alza hombre",
                "Alza_hombre",
            ),
        },
        {
            "archivo": ("barredora", "barredora_hombre_a_bordo"),
            "titulo": "Barredora",
            "aliases": (
                "Barredora",
                "Barredora hombre a bordo",
                "Barredora Tennant",
            ),
        },
        {
            "archivo": ("camion_ford", "camion", "camión_ford"),
            "titulo": "Camión Ford",
            "aliases": (
                "Camion",
                "Camión",
                "Camion Ford",
                "Camión Ford",
            ),
        },
        {
            "archivo": (
                "grua_orquilla",
                "grua_horquilla",
                "grúa_horquilla",
            ),
            "titulo": "Grúa horquilla",
            "aliases": (
                "Grúa horquilla",
                "Grua horquilla",
                "Grúa orquilla",
                "Grua orquilla",
            ),
        },
        {
            "archivo": ("minicargador", "mini_cargador"),
            "titulo": "Minicargador",
            "aliases": (
                "Minicargador",
                "Mini cargador",
            ),
        },
    ]

    panel_titulo("Equipos certificados")

    columnas = st.columns(len(configuracion))

    for columna, equipo in zip(columnas, configuracion):
        ruta = buscar_foto_certificacion(*equipo["archivo"])
        imagen_uri = obtener_data_uri_recurso([ruta]) if ruta else ""
        fila = obtener_ficha_certificacion(df, equipo["aliases"])

        if fila is not None:
            certificacion = str(
                fila.get("Nombre_Certificacion", "")
            ).strip()
            entidad = str(
                fila.get("Entidad_Emisora", "")
            ).strip()
            estado = str(
                fila.get("Estado", "Sin vencimiento")
            ).strip()
            vencimiento = fecha_texto(
                fila.get("Vencimiento", pd.NaT)
            )
            dias = fila.get("Dias_Para_Vencer", pd.NA)
        else:
            certificacion = "Sin registro asociado"
            entidad = ""
            estado = "Sin vencimiento"
            vencimiento = ""
            dias = pd.NA

        clase_estado = clase_estado_certificacion(estado)
        texto_dias = texto_dias_certificacion(dias, estado)

        if imagen_uri:
            imagen_html = (
                f'<img class="cert-equipment-img" '
                f'src="{imagen_uri}" '
                f'alt="{escape_html(equipo["titulo"])}">'
            )
        else:
            imagen_html = (
                '<div class="cert-equipment-placeholder">'
                '📷'
                '</div>'
            )

        entidad_html = (
            f'<div class="cert-equipment-detail">'
            f'<span>Entidad:</span> {escape_html(entidad)}'
            f'</div>'
            if entidad
            else ""
        )

        vencimiento_html = (
            f'<div class="cert-equipment-detail">'
            f'<span>Vencimiento:</span> {escape_html(vencimiento)}'
            f'</div>'
            if vencimiento
            else ""
        )

        tarjeta = (
            '<div class="cert-equipment-card">'
            '<div class="cert-equipment-image-shell">'
            f'{imagen_html}'
            '</div>'
            f'<div class="cert-equipment-title">'
            f'{escape_html(equipo["titulo"])}'
            '</div>'
            f'<div class="cert-equipment-detail">'
            f'<span>Certificación:</span> '
            f'{escape_html(certificacion)}'
            '</div>'
            f'{entidad_html}'
            f'<div class="cert-equipment-status {clase_estado}">'
            f'{escape_html(estado)}'
            '</div>'
            f'{vencimiento_html}'
            f'<div class="cert-equipment-days">'
            f'{escape_html(texto_dias)}'
            '</div>'
            '</div>'
        )

        with columna:
            st.markdown(tarjeta, unsafe_allow_html=True)



def obtener_sello_certificaciones_html():
    # Sello de agua exclusivo del módulo Certificaciones.
    # Rutas reconocidas:
    # static/certificaciones/sello.png
    # static-certificaciones/sello.png
    ruta = buscar_foto_certificacion(
        "sello",
        "sello_certificaciones",
        "sello_certificacion",
    )

    imagen_uri = obtener_data_uri_recurso([ruta]) if ruta else ""

    if not imagen_uri:
        return ""

    return (
        '<div class="certification-watermark-layer" aria-hidden="true">'
        f'<img class="certification-watermark-img" '
        f'src="{imagen_uri}" alt="">'
        '</div>'
    )


def obtener_sello_reconocimientos_html():
    # Sello de agua exclusivo del módulo Reconocimientos.
    # Ruta principal esperada:
    # static/reconocimientos/saivam700.png
    ruta = buscar_foto_reconocimiento(
        "saivam700",
        "saivam_700",
    )

    imagen_uri = obtener_data_uri_recurso([ruta]) if ruta else ""

    if not imagen_uri:
        return ""

    return (
        '<div class="recognition-watermark-layer" aria-hidden="true">'
        f'<img class="recognition-watermark-img" src="{imagen_uri}" alt="">'
        '</div>'
    )


def obtener_sello_agua_html():
    """Carga la imagen 'agua' como sello de agua exclusivo del Panel General."""
    ruta = buscar_imagen_local("agua")
    imagen_uri = obtener_data_uri_recurso([ruta]) if ruta else ""
    if not imagen_uri:
        return ""
    return (
        '<div class="panel-watermark-layer" aria-hidden="true">'
        f'<img class="panel-watermark-img" src="{imagen_uri}" alt="">'
        '</div>'
    )


def obtener_sello_cumplimientos_html():
    """Carga la imagen 'saivam' como sello de agua de Cumplimientos SSO."""
    ruta = buscar_imagen_local("saivam")
    imagen_uri = obtener_data_uri_recurso([ruta]) if ruta else ""
    if not imagen_uri:
        return ""
    return (
        '<div class="cumplimientos-watermark-layer" aria-hidden="true">'
        f'<img class="cumplimientos-watermark-img" src="{imagen_uri}" alt="">'
        '</div>'
    )


def obtener_sello_saivam_global_html():
    """Carga la imagen 'saivam' completa como sello de agua para los módulos."""
    ruta = buscar_imagen_local("saivam")
    imagen_uri = obtener_data_uri_recurso([ruta]) if ruta else ""

    if not imagen_uri:
        return ""

    return (
        '<div class="saivam-page-watermark-layer" aria-hidden="true">'
        f'<img class="saivam-page-watermark-img" src="{imagen_uri}" alt="">'
        '</div>'
    )


def mostrar_sello_saivam_pagina():
    """Inserta el sello SAIVAM sobre el fondo y de forma tenue sobre los paneles."""
    sello = obtener_sello_saivam_global_html()
    if sello:
        st.markdown(sello, unsafe_allow_html=True)


def aplicar_estilo():
    fondo_uri = obtener_data_uri_recurso([
        ruta_app("fondo_seguridad.jpg"),
        ruta_app("fondo_seguridad.png"),
        ruta_app("fondo_sso.jpg"),
        ruta_app("fondo_sso.png"),
        ruta_app("fotos", "fondo_seguridad.jpg"),
        ruta_app("fotos", "fondo_sso.jpg"),
    ])

    if fondo_uri:
        fondo_css = (
            "linear-gradient(90deg, rgba(1,5,4,.92), rgba(3,15,11,.86)), "
            f"url('{fondo_uri}')"
        )
    else:
        fondo_css = (
            "radial-gradient(circle at 72% 8%, rgba(16,73,54,.34), transparent 34%), "
            "radial-gradient(circle at 18% 80%, rgba(26,137,96,.20), transparent 36%), "
            "linear-gradient(135deg, #010403 0%, #050b09 52%, #081712 100%)"
        )

    css = r"""
<style>
header[data-testid="stHeader"],
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
[data-testid="stDeployButton"],
[data-testid="collapsedControl"],
[data-testid="stSidebarCollapsedControl"],
[data-testid="stSidebarCollapseButton"],
[data-testid="stSidebarHeader"],
button[data-testid="stSidebarCollapseButton"],
button[aria-label="Collapse sidebar"],
button[aria-label="Expand sidebar"],
button[aria-label="Close sidebar"],
button[aria-label="Open sidebar"],
#MainMenu,
footer {
    display: none !important;
    visibility: hidden !important;
    height: 0 !important;
}

html, body, .stApp, [data-testid="stAppViewContainer"] {
    color: #F4FFF9 !important;
    background-color: #010403 !important;
}

:root {
    --sidebar-fixed-width: 250px;
}

.stApp {
    background-image: __FONDO__ !important;
    background-size: cover !important;
    background-position: center !important;
    background-attachment: fixed !important;
}

.stApp::before {
    content: "";
    position: fixed;
    inset: 0;
    pointer-events: none;
    background:
        linear-gradient(110deg, rgba(0,0,0,.34), rgba(0,0,0,.08) 52%),
        repeating-linear-gradient(90deg, rgba(110,231,183,.035) 0 1px, transparent 1px 92px);
    z-index: 0;
}

/*
El contenedor principal conserva siempre dos columnas reales: menú y contenido.
El menú permanece fijo visualmente, pero continúa ocupando su espacio dentro del
layout. De esta forma el contenido nunca queda debajo del menú al cambiar el
ancho de la ventana o el nivel de zoom.
*/
[data-testid="stAppViewContainer"] {
    display: flex !important;
    flex-direction: row !important;
    align-items: stretch !important;
    width: 100% !important;
    min-width: 0 !important;
    overflow-x: hidden !important;
}

section[data-testid="stSidebar"] {
    position: sticky !important;
    top: 0 !important;
    left: auto !important;
    bottom: auto !important;
    align-self: flex-start !important;
    flex: 0 0 var(--sidebar-fixed-width) !important;
    transform: none !important;
    visibility: visible !important;
    width: var(--sidebar-fixed-width) !important;
    min-width: var(--sidebar-fixed-width) !important;
    max-width: var(--sidebar-fixed-width) !important;
    height: 100vh !important;
    z-index: 9999 !important;
    background: linear-gradient(180deg, #043D31 0%, #075844 55%, #064735 100%) !important;
    border-right: 1px solid rgba(134,239,172,.28) !important;
    box-shadow: 10px 0 28px rgba(4,63,49,.24) !important;
    overflow: hidden !important;
    transition: none !important;
}

/* Compatibilidad con distintas versiones de Streamlit. */
[data-testid="stAppViewContainer"] > .main,
[data-testid="stAppViewContainer"] > section[data-testid="stMain"] {
    position: relative !important;
    z-index: 1 !important;
    flex: 1 1 auto !important;
    width: auto !important;
    max-width: none !important;
    min-width: 0 !important;
    margin-left: 0 !important;
    transform: none !important;
    overflow-x: hidden !important;
}

.main .block-container,
[data-testid="stMain"] .block-container,
[data-testid="stMainBlockContainer"] {
    box-sizing: border-box !important;
    padding-top: 1rem !important;
    padding-left: 1.6rem !important;
    padding-right: 1.45rem !important;
    padding-bottom: 1rem !important;
    width: 100% !important;
    max-width: 100% !important;
    min-width: 0 !important;
    margin: 0 !important;
}

section[data-testid="stSidebar"] > div {
    height: 100vh !important;
    box-sizing: border-box !important;
    overflow-y: auto !important;
    overflow-x: hidden !important;
    scrollbar-width: thin !important;
    scrollbar-color: rgba(167,243,208,.52) transparent !important;
    background: transparent !important;
    padding: 0 !important;
}


/* Compactación real del contenido interno del menú lateral. */
section[data-testid="stSidebar"] [data-testid="stSidebarContent"] {
    width: 100% !important;
    max-width: 100% !important;
    padding: 0 !important;
    margin: 0 !important;
    box-sizing: border-box !important;
}

section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
    width: 100% !important;
    max-width: 100% !important;
    padding: 8px 5px 10px 5px !important;
    margin: 0 !important;
    box-sizing: border-box !important;
}

section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] > div,
section[data-testid="stSidebar"] [data-testid="stSidebarUserContent"] [data-testid="stVerticalBlock"],
section[data-testid="stSidebar"] [data-testid="stElementContainer"] {
    width: 100% !important;
    max-width: 100% !important;
    box-sizing: border-box !important;
}

section[data-testid="stSidebar"] div[role="radiogroup"] {
    padding-right: 0 !important;
    margin-right: 0 !important;
}

section[data-testid="stSidebar"] > div::-webkit-scrollbar {
    width: 6px;
}

section[data-testid="stSidebar"] > div::-webkit-scrollbar-track {
    background: transparent;
}

section[data-testid="stSidebar"] > div::-webkit-scrollbar-thumb {
    background: rgba(167,243,208,.46);
    border-radius: 999px;
}

section[data-testid="stSidebar"] * {
    color: #ffffff !important;
}

section[data-testid="stSidebar"] hr {
    border-color: rgba(167,243,208,.25) !important;
    margin: 9px 0 !important;
}

.menu-brand {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 2px 2px 9px 2px;
    margin-bottom: 6px;
    border-bottom: 1px solid rgba(167,243,208,.24);
}

.menu-logo-shell {
    width: 44px;
    height: 44px;
    min-width: 44px;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(145deg, #ffffff 0%, #d7dee7 55%, #7b8796 100%);
    border: 1px solid rgba(255,255,255,.72);
    box-shadow: 0 8px 18px rgba(0,0,0,.34), inset 0 0 0 3px rgba(15,23,42,.16);
    overflow: hidden;
}

.menu-logo-shell .brand-logo-img {
    width: 100%;
    height: 100%;
    object-fit: cover;
    padding: 0;
    border-radius: 50%;
}

.brand-fallback {
    display: flex;
    align-items: center;
    gap: 5px;
}

.brand-mark {
    width: 42px;
    height: 27px;
    border-radius: 50%;
    background: #ffd500;
    color: #111827 !important;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-weight: 1000;
    font-style: italic;
    border: 5px solid #111827;
    transform: rotate(-12deg);
}

.brand-name {
    color: #111827 !important;
    font-size: 21px;
    font-weight: 1000;
    letter-spacing: 2px;
}

.menu-title {
    color: #ffffff !important;
    font-weight: 1000;
    font-size: 12.4px;
    line-height: 1.22;
    letter-spacing: .25px;
    text-transform: uppercase;
}

.menu-subtitle {
    color: #B7F7D4 !important;
    font-size: 9.9px;
    margin-top: 4px;
    font-weight: 900;
    letter-spacing: .3px;
}

section[data-testid="stSidebar"] div[role="radiogroup"] {
    display: flex !important;
    flex-direction: column !important;
    align-items: stretch !important;
    width: 100% !important;
    gap: 0 !important;
}

/* Todos los ítems tienen el mismo ancho y una altura compacta. */
section[data-testid="stSidebar"] div[role="radiogroup"] label {
    display: flex !important;
    align-items: center !important;
    width: 100% !important;
    min-width: 100% !important;
    max-width: 100% !important;
    min-height: 42px !important;
    box-sizing: border-box !important;
    border-radius: 13px !important;
    padding: 7px 7px !important;
    margin: 0 0 5px 0 !important;
    border: 1px solid rgba(133,213,175,.48) !important;
    background: rgba(4,61,49,.86) !important;
    box-shadow: inset 0 0 0 1px rgba(255,255,255,.015) !important;
    transition: all .18s ease !important;
}

section[data-testid="stSidebar"] div[role="radiogroup"] label:hover {
    transform: translateX(2px);
    background: rgba(10,104,78,.94) !important;
    border-color: rgba(110,231,183,.78) !important;
}

section[data-testid="stSidebar"] div[role="radiogroup"] label:has(input:checked) {
    background: linear-gradient(100deg, #19A96B 0%, #27C97F 100%) !important;
    border: 1px solid rgba(219,234,254,.96) !important;
    box-shadow: 0 9px 21px rgba(16,185,129,.28), inset 0 0 18px rgba(255,255,255,.10) !important;
}

section[data-testid="stSidebar"] div[role="radiogroup"] label > div:first-child {
    display: none !important;
}

section[data-testid="stSidebar"] div[role="radiogroup"] p {
    width: 100% !important;
    margin: 0 !important;
    color: #ffffff !important;
    font-weight: 900 !important;
    font-size: 12.2px !important;
    line-height: 1.2 !important;
    white-space: nowrap !important;
    overflow: visible !important;
    text-overflow: clip !important;
    text-shadow: 0 1px 1px rgba(0,0,0,.42);
}

section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span {
    font-weight: 800 !important;
}

section[data-testid="stSidebar"] [data-baseweb="select"] > div {
    min-height: 38px !important;
    background: #ffffff !important;
    border-radius: 10px !important;
    border: 1px solid rgba(255,255,255,.6) !important;
}

section[data-testid="stSidebar"] [data-baseweb="select"] * {
    color: #111827 !important;
    font-weight: 800 !important;
}

.sidebar-filter-title {
    color: #A7F3D0 !important;
    font-size: 10px;
    font-weight: 950;
    text-transform: uppercase;
    letter-spacing: .8px;
    margin-bottom: 8px;
}

.menu-footer-box {
    border: 1px solid rgba(133,213,175,.44);
    background: rgba(4,71,54,.90);
    border-radius: 15px;
    padding: 10px 11px;
    margin-top: 10px;
    box-shadow: 0 10px 20px rgba(0,0,0,.22);
}

.menu-info {
    color: #ffffff !important;
    font-size: 9.9px;
    line-height: 1.5;
    font-weight: 750;
}

.menu-info b {
    color: #A7F3D0 !important;
}

.app-topbar {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 18px;
    margin: 0 0 10px 0;
}

.title-main {
    font-size: clamp(30px, 2.35vw, 43px);
    font-weight: 1000;
    color: #D1FAE5;
    margin: 1px 0 2px 0;
    line-height: 1.04;
    letter-spacing: -1.5px;
    text-shadow: 0 2px 10px rgba(0,0,0,.78);
}

.subtitle-main {
    color: #B7E4D0;
    font-size: 14px;
    font-weight: 760;
    margin-top: 7px;
    max-width: 950px;
}

.main-logo-card {
    flex: 0 0 auto;
    min-width: 0;
    width: auto;
    height: auto;
    padding: 0;
    margin: 0;
    background: transparent;
    border: 0;
    border-radius: 0;
    box-shadow: none;
    display: flex;
    align-items: flex-start;
    justify-content: flex-end;
    overflow: visible;
}

.main-logo-card .brand-logo-img {
    display: block;
    width: 180px;
    max-width: 100%;
    height: auto;
    max-height: 58px;
    padding: 0;
    margin: 0;
    border: 0;
    border-radius: 0;
    background: transparent;
    box-shadow: none;
    object-fit: contain;
}

.main-logo-card .brand-fallback .brand-mark {
    width: 54px;
    height: 31px;
}

.main-logo-card .brand-name {
    font-size: 25px;
}

.kpi-card {
    position: relative;
    overflow: hidden;
    background: linear-gradient(145deg, rgba(3,10,8,.93), rgba(8,31,24,.84));
    border: 1px solid rgba(52,211,153,.36);
    border-radius: 17px;
    padding: 18px 20px 16px 20px;
    min-height: 190px;
    box-shadow: 0 12px 30px rgba(0,0,0,.34), inset 0 1px 0 rgba(110,231,183,.11);
    backdrop-filter: blur(11px);
    -webkit-backdrop-filter: blur(11px);
}

.kpi-card::after {
    content: "";
    position: absolute;
    right: -34px;
    bottom: -50px;
    width: 120px;
    height: 120px;
    border-radius: 50%;
    background: rgba(52,211,153,.07);
}

.kpi-icon {
    width: 58px;
    height: 58px;
    border-radius: 18px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 28px;
    margin-bottom: 13px;
    box-shadow: inset 0 0 0 1px rgba(255,255,255,.35);
}

.kpi-icon.azul { background: rgba(16,185,129,.18); }
.kpi-icon.verde { background: rgba(34,197,94,.18); }
.kpi-icon.morado { background: rgba(5,150,105,.17); }
.kpi-icon.ambar { background: rgba(132,204,22,.18); }
.kpi-icon.rojo { background: rgba(45,212,191,.18); }
.kpi-icon.celeste { background: rgba(20,184,166,.18); }

.kpi-title {
    font-size: 12.8px;
    color: #B7E4D0;
    font-weight: 930;
}

.kpi-value {
    font-size: 25px;
    font-weight: 1000;
    color: #FFFFFF;
    margin-top: 8px;
    line-height: 1.1;
}

.kpi-sub {
    font-size: 12px;
    color: #9FCBB9;
    margin-top: 9px;
    line-height: 1.32;
}

.panel-title {
    color: #D1FAE5;
    font-weight: 1000;
    font-size: 22px;
    margin-top: 16px;
    margin-bottom: 8px;
    letter-spacing: -.35px;
}

.subsection-label {
    color: #A7F3D0;
    font-size: 14px;
    font-weight: 900;
    text-transform: uppercase;
    letter-spacing: .7px;
    margin: 10px 0 8px 2px;
}

div[data-testid="stVerticalBlockBorderWrapper"] {
    background: linear-gradient(145deg, rgba(2,8,6,.92), rgba(7,27,21,.84)) !important;
    border: 1px solid rgba(52,211,153,.30) !important;
    border-radius: 21px !important;
    box-shadow: 0 12px 30px rgba(0,0,0,.32), inset 0 1px 0 rgba(110,231,183,.08) !important;
    backdrop-filter: blur(10px);
    -webkit-backdrop-filter: blur(10px);
}

div[data-testid="stVerticalBlockBorderWrapper"] > div {
    padding: 7px 10px 4px 10px !important;
}

[data-testid="stPlotlyChart"] {
    background: linear-gradient(145deg, rgba(2,8,6,.93), rgba(7,27,21,.86));
    border: 1px solid rgba(52,211,153,.28);
    border-radius: 20px;
    padding: 5px 8px 0 8px;
    box-shadow: 0 12px 30px rgba(0,0,0,.30), inset 0 1px 0 rgba(110,231,183,.08);
    backdrop-filter: blur(9px);
    -webkit-backdrop-filter: blur(9px);
    overflow: hidden;
}

div[data-testid="stDataFrame"] {
    background: rgba(1,7,5,.92);
    border: 1px solid rgba(52,211,153,.28);
    border-radius: 18px;
    padding: 3px;
    box-shadow: 0 10px 26px rgba(0,0,0,.30);
    overflow: hidden;
}

.alert-card,
.compromiso-card {
    background: rgba(3,12,9,.92);
    border: 1px solid rgba(52,211,153,.28);
    border-radius: 16px;
    padding: 13px 14px;
    margin-bottom: 10px;
    box-shadow: 0 7px 18px rgba(15,23,42,.055);
    backdrop-filter: blur(8px);
}

.alert-card {
    border-left: 6px solid #ef4444;
}

.alert-card.ok {
    border-left-color: #22c55e;
}

.alert-title,
.compromiso-title {
    font-weight: 950;
    color: #F4FFF9;
    font-size: 12.8px;
}

.alert-sub,
.compromiso-sub {
    color: #A9D5C3;
    font-size: 11.8px;
    margin-top: 5px;
    line-height: 1.35;
}

.compromiso-head {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    gap: 8px;
}

.compromiso-badge {
    white-space: nowrap;
    padding: 5px 9px;
    border-radius: 999px;
    font-size: 10.5px;
    font-weight: 950;
    color: #b54708;
    background: #fffaeb;
    border: 1px solid #fec84b;
}

.compromiso-badge.vencida {
    color: #b42318;
    background: #fef3f2;
    border-color: #fda29b;
}

.badge {
    display: inline-block;
    padding: 5px 10px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 950;
}

.badge-ok { background:#dcfce7; color:#166534; }
.badge-proceso { background:#dbeafe; color:#1d4ed8; }
.badge-pendiente { background:#fef9c3; color:#854d0e; }
.badge-vencida { background:#fee2e2; color:#991b1b; }
.badge-neutro { background:#e2e8f0; color:#334155; }

.stAlert {
    border-radius: 16px !important;
    background: rgba(3,12,9,.58) !important;
    border: 1px solid rgba(52,211,153,.30) !important;
}

/*
Transparencia uniforme para todos los módulos. Los paneles conservan contraste
para la lectura, pero permiten ver el sello de agua como en KPI SSO.
*/
.kpi-card {
    background: linear-gradient(
        145deg,
        rgba(3,10,8,.64),
        rgba(8,31,24,.50)
    ) !important;
    box-shadow: 0 10px 25px rgba(0,0,0,.24), inset 0 1px 0 rgba(110,231,183,.10) !important;
    backdrop-filter: blur(5px) !important;
    -webkit-backdrop-filter: blur(5px) !important;
}

div[data-testid="stVerticalBlockBorderWrapper"] {
    background: linear-gradient(
        145deg,
        rgba(2,8,6,.58),
        rgba(7,27,21,.46)
    ) !important;
    box-shadow: 0 10px 25px rgba(0,0,0,.22), inset 0 1px 0 rgba(110,231,183,.08) !important;
    backdrop-filter: blur(5px) !important;
    -webkit-backdrop-filter: blur(5px) !important;
}

[data-testid="stPlotlyChart"] {
    background: linear-gradient(
        145deg,
        rgba(2,8,6,.50),
        rgba(7,27,21,.40)
    ) !important;
    box-shadow: 0 10px 25px rgba(0,0,0,.20), inset 0 1px 0 rgba(110,231,183,.07) !important;
    backdrop-filter: blur(4px) !important;
    -webkit-backdrop-filter: blur(4px) !important;
}

div[data-testid="stDataFrame"],
.tabla-general-wrap,
.cert-table-wrap {
    background: rgba(1,7,5,.56) !important;
    box-shadow: 0 8px 22px rgba(0,0,0,.20) !important;
    backdrop-filter: blur(4px) !important;
    -webkit-backdrop-filter: blur(4px) !important;
}

.tabla-general-compacta th,
.cert-table th {
    background: rgba(27,32,41,.68) !important;
}

.alert-card,
.compromiso-card {
    background: rgba(3,12,9,.58) !important;
    box-shadow: 0 7px 18px rgba(0,0,0,.16) !important;
    backdrop-filter: blur(4px) !important;
    -webkit-backdrop-filter: blur(4px) !important;
}

.cert-equipment-card {
    background:
        radial-gradient(
            circle at 100% 100%,
            rgba(16,185,129,.08) 0,
            rgba(16,185,129,.08) 24%,
            transparent 25%
        ),
        rgba(4,24,18,.58) !important;
    box-shadow: 0 10px 24px rgba(0,0,0,.18) !important;
    backdrop-filter: blur(4px) !important;
    -webkit-backdrop-filter: blur(4px) !important;
}

.cumplimiento-hero {
    background: linear-gradient(
        135deg,
        rgba(15,23,42,.58),
        rgba(6,78,59,.26)
    ) !important;
    backdrop-filter: blur(4px) !important;
    -webkit-backdrop-filter: blur(4px) !important;
}

.persona-card {
    background: rgba(15,23,42,.48) !important;
    backdrop-filter: blur(4px) !important;
    -webkit-backdrop-filter: blur(4px) !important;
}

/*
Sello de agua exclusivo del Panel General. Ocupa toda el área principal,
desde el borde del menú lateral hasta el extremo derecho de la pantalla.
*/
.panel-watermark-layer {
    position: fixed;
    left: var(--sidebar-fixed-width);
    right: 0;
    top: 0;
    bottom: 0;
    width: calc(100vw - var(--sidebar-fixed-width));
    height: 100vh;
    display: flex;
    align-items: stretch;
    justify-content: stretch;
    overflow: hidden;
    pointer-events: none;
    z-index: 8;
}

.panel-watermark-img {
    display: block;
    width: 100%;
    height: 100%;
    max-width: none;
    max-height: none;
    object-fit: cover;
    object-position: center center;
    transform: scale(1.35);
    transform-origin: center center;
    opacity: .16;
    filter: saturate(.72) contrast(1.02) brightness(.84);
    user-select: none;
}

/*
Sello de agua SAIVAM para los módulos del menú lateral.
La imagen se muestra completa, sin recortes, ocupando el área disponible
entre el menú lateral y el borde derecho de la ventana.
*/
.saivam-page-watermark-layer {
    position: fixed;
    left: var(--sidebar-fixed-width);
    right: 0;
    top: 0;
    bottom: 0;
    width: calc(100vw - var(--sidebar-fixed-width));
    height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    overflow: hidden;
    pointer-events: none;
    z-index: 8;
    padding: 1.2rem 1.4rem;
    box-sizing: border-box;
}

.saivam-page-watermark-img {
    display: block;
    width: 100%;
    height: 100%;
    max-width: 100%;
    max-height: 100%;
    object-fit: contain;
    object-position: center center;
    opacity: .15;
    filter: saturate(.82) contrast(1.04) brightness(.88);
    user-select: none;
}

@media (max-width: 900px) {
    .saivam-page-watermark-layer {
        left: var(--sidebar-fixed-width);
        width: calc(100vw - var(--sidebar-fixed-width));
        height: 100vh;
        padding: .7rem;
    }

    .saivam-page-watermark-img {
        object-fit: contain;
        opacity: .12;
    }
}

/* Sello de agua exclusivo del módulo Cumplimientos SSO. */
.cumplimientos-watermark-layer {
    position: fixed;
    left: var(--sidebar-fixed-width);
    right: 0;
    top: 0;
    bottom: 0;
    width: calc(100vw - var(--sidebar-fixed-width));
    height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    overflow: hidden;
    pointer-events: none;
    z-index: 0;
}

.cumplimientos-watermark-img {
    display: block;
    width: 100%;
    height: 100%;
    max-width: none;
    max-height: none;
    object-fit: cover;
    object-position: center center;
    opacity: .24;
    filter: grayscale(.02) saturate(.95) contrast(1.12) brightness(.92);
    transform: scale(1.08);
    transform-origin: center center;
    user-select: none;
}

@media (max-width: 900px) {
    .cumplimientos-watermark-layer {
        left: var(--sidebar-fixed-width);
        width: calc(100vw - var(--sidebar-fixed-width));
        height: 100vh;
    }

    .cumplimientos-watermark-img {
        opacity: .19;
        transform: scale(1.12);
    }
}

.footer-app {
    width: 100%;
    margin: 28px auto 8px auto;
    padding: 0 18px;
    box-sizing: border-box;
    color: #9FCBB9;
    font-size: 10.4px;
    font-weight: 740;
    text-align: center;
    display: flex;
    justify-content: center;
    align-items: center;
    line-height: 1.45;
}


.alert-card {
    width: 100%;
    box-sizing: border-box;
    margin: 0 0 9px 0;
    padding: 11px 14px;
    border: 1px solid rgba(35, 189, 128, .42);
    border-left: 5px solid #F59E0B;
    border-radius: 12px;
    background: rgba(3, 18, 13, .88);
}

.alert-card-title {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
    color: #F3FFF8;
    font-size: 12px;
    font-weight: 850;
}

.alert-card-title span {
    flex: 0 0 auto;
    padding: 3px 8px;
    border: 1px solid rgba(251, 191, 36, .55);
    border-radius: 999px;
    background: rgba(245, 158, 11, .13);
    color: #FDE68A;
    font-size: 9px;
    font-weight: 800;
}

.alert-card-text {
    margin-top: 5px;
    color: #A9D2C1;
    font-size: 10px;
    line-height: 1.35;
}

.footer-app.footer-app-dos-lineas {
    display: block !important;
    text-align: center !important;
}
.footer-titulo {
    display: block;
    margin-bottom: 2px;
    font-size: 12px;
    font-weight: 900;
    color: #B8E8D3;
}
.footer-detalle {
    display: block;
    font-size: 10.4px;
    font-weight: 740;
    color: #9FCBB9;
    line-height: 1.35;
}







/* Sello de agua exclusivo del módulo Certificaciones. */
.certification-watermark-layer {
    position: fixed;
    left: var(--sidebar-fixed-width);
    top: 0;
    width: calc(100vw - var(--sidebar-fixed-width));
    height: 100vh;
    overflow: hidden;
    pointer-events: none;
    z-index: 0;
}

.certification-watermark-img {
    display: block;
    width: 100%;
    height: 100%;
    max-width: none;
    max-height: none;
    object-fit: cover;
    object-position: center center;
    opacity: .22;
    filter:
        grayscale(.08)
        saturate(.78)
        contrast(1.08)
        brightness(.72);
    transform: scale(1.06);
    transform-origin: center center;
    user-select: none;
}

section.main,
[data-testid="stAppViewContainer"] > .main {
    position: relative;
    z-index: 1;
}

@media (max-width: 900px) {
    .certification-watermark-layer {
        left: var(--sidebar-fixed-width);
        width: calc(100vw - var(--sidebar-fixed-width));
        height: 100vh;
    }

    .certification-watermark-img {
        width: 100%;
        height: 100%;
        object-fit: cover;
        opacity: .18;
        transform: scale(1.08);
    }
}

/* Tarjetas fotográficas del módulo Certificaciones. */
.cert-equipment-card {
    width: 100%;
    min-height: 285px;
    padding: 6px;
    margin-bottom: 8px;
    border: 1px solid rgba(52, 211, 153, .38);
    border-radius: 20px;
    background:
        radial-gradient(
            circle at 100% 100%,
            rgba(16, 185, 129, .12) 0,
            rgba(16, 185, 129, .12) 24%,
            transparent 25%
        ),
        rgba(4, 24, 18, .88);
    box-shadow: 0 13px 28px rgba(0, 0, 0, .24);
    box-sizing: border-box;
    overflow: hidden;
}

.cert-equipment-image-shell {
    width: 100%;
    height: 102px;
    margin-bottom: 10px;
    border-radius: 10px;
    overflow: hidden;
    background: rgba(232, 255, 245, .10);
    border: 1px solid rgba(184, 232, 212, .18);
}

.cert-equipment-img {
    display: block;
    width: 100%;
    height: 100%;
    object-fit: cover;
    object-position: center;
}

.cert-equipment-placeholder {
    width: 100%;
    height: 100%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 42px;
    background: rgba(255, 255, 255, .05);
}

.cert-equipment-title {
    min-height: 30px;
    color: #F4FFF9 !important;
    font-size: 13px;
    font-weight: 950;
    line-height: 1.22;
    margin: 2px 2px 4px 2px;
}

.cert-equipment-detail {
    min-height: 27px;
    color: #BBD7CA !important;
    font-size: 11.5px;
    line-height: 1.35;
    margin: 3px 2px;
}

.cert-equipment-detail span {
    color: #E4FFF2 !important;
    font-weight: 900;
}

.cert-equipment-status {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-height: 22px;
    margin: 4px 2px 4px 2px;
    padding: 3px 8px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 950;
    line-height: 1;
}

.cert-status-vigente {
    color: #4ADE80 !important;
    background: rgba(22, 101, 52, .30);
    border: 1px solid rgba(74, 222, 128, .42);
}

.cert-status-por-vencer {
    color: #FBBF24 !important;
    background: rgba(146, 64, 14, .30);
    border: 1px solid rgba(251, 191, 36, .42);
}

.cert-status-vencida {
    color: #FCA5A5 !important;
    background: rgba(153, 27, 27, .32);
    border: 1px solid rgba(252, 165, 165, .42);
}

.cert-status-sin-vencimiento {
    color: #CBD5E1 !important;
    background: rgba(71, 85, 105, .32);
    border: 1px solid rgba(203, 213, 225, .30);
}

.cert-equipment-days {
    color: #A7C8B9 !important;
    font-size: 12px;
    font-weight: 750;
    margin: 4px 2px 1px 2px;
}

@media (max-width: 1250px) {
    .cert-equipment-card {
        min-height: 270px;
        padding: 5px;
    }

    .cert-equipment-image-shell {
        height: 92px;
    }

    .cert-equipment-title {
        font-size: 12px;
        min-height: 28px;
    }

    .cert-equipment-detail {
        font-size: 9.8px;
        min-height: 24px;
    }
}

/* Sello de agua exclusivo del módulo Reconocimientos. */
.recognition-watermark-layer {
    position: fixed;
    left: var(--sidebar-fixed-width);
    top: 0;
    width: calc(100vw - var(--sidebar-fixed-width));
    height: 100vh;
    overflow: hidden;
    pointer-events: none;
    z-index: 0;
}

.recognition-watermark-img {
    display: block;
    width: 100%;
    height: 100%;
    max-width: none;
    max-height: none;
    object-fit: cover;
    object-position: center center;
    opacity: .13;
    filter: grayscale(.10) saturate(.78) contrast(1.08) brightness(.72);
    transform: scale(1.02);
    transform-origin: center center;
    user-select: none;
}

section.main,
[data-testid="stAppViewContainer"] > .main {
    position: relative;
    z-index: 1;
}

@media (max-width: 900px) {
    .recognition-watermark-layer {
        left: var(--sidebar-fixed-width);
        width: calc(100vw - var(--sidebar-fixed-width));
        height: 100vh;
    }

    .recognition-watermark-img {
        width: 100%;
        height: 100%;
        object-fit: cover;
        opacity: .11;
        transform: scale(1.05);
    }
}

/* Carrusel del módulo Reconocimientos. */
.recognition-carousel-wrapper {
    width: 100%;
    display: flex;
    justify-content: center;
    margin: 8px 0 8px 0;
}

.recognition-photo-card {
    width: min(100%, 980px);
    padding: 0;
    border: 1px solid rgba(52, 211, 153, .42);
    border-radius: 20px;
    background: transparent !important;
    box-shadow: none;
    box-sizing: border-box;
    overflow: hidden;
}

.recognition-photo-img {
    display: block;
    width: 100%;
    height: 500px;
    object-fit: contain;
    object-position: center;
    border-radius: 19px;
    background: transparent !important;
}
.recognition-carousel-counter {
    min-height: 40px;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #B8E8D4 !important;
    font-size: 13px;
    font-weight: 900;
    text-align: center;
}

@media (max-width: 1100px) {
    .recognition-photo-img {
        height: 390px;
    }
}

@media (max-width: 700px) {
    .recognition-photo-img {
        height: 280px;
    }
}

/* Refuerzo del tema oscuro para el área principal. El menú conserva su paleta verde. */
[data-testid="stAppViewContainer"] > .main,
[data-testid="stAppViewContainer"] > section[data-testid="stMain"],
[data-testid="stMainBlockContainer"] {
    background: transparent !important;
    color: #F4FFF9 !important;
}

[data-testid="stMain"] p,
[data-testid="stMain"] span,
[data-testid="stMain"] label,
[data-testid="stMain"] li {
    color: #D7F8E9;
}

[data-testid="stMain"] h1,
[data-testid="stMain"] h2,
[data-testid="stMain"] h3,
[data-testid="stMain"] h4 {
    color: #F4FFF9 !important;
}

[data-testid="stMain"] .stCaption,
[data-testid="stMain"] small {
    color: #9FCBB9 !important;
}

[data-testid="stMain"] [data-baseweb="select"] > div,
[data-testid="stMain"] input,
[data-testid="stMain"] textarea {
    background: rgba(2,10,7,.94) !important;
    color: #F4FFF9 !important;
    border-color: rgba(52,211,153,.34) !important;
}

[data-testid="stMain"] button {
    border-color: rgba(52,211,153,.38) !important;
}

@media (max-width: 1100px) {
    /* Se mantiene exactamente el mismo menú y separación del contenido. */
    section[data-testid="stSidebar"] div[role="radiogroup"] p {
        font-size: 12.2px !important;
    }

    .panel-watermark-layer {
        left: var(--sidebar-fixed-width);
        right: 0;
        top: 0;
        bottom: 0;
        width: calc(100vw - var(--sidebar-fixed-width));
        height: 100vh;
    }

    .panel-watermark-img {
        width: 100%;
        height: 100%;
        max-width: none;
        max-height: none;
        object-fit: cover;
        object-position: center center;
        transform: scale(1.35);
        transform-origin: center center;
    }
}
</style>
    """
    css = css.replace("__FONDO__", fondo_css)
    st.markdown(css, unsafe_allow_html=True)


def kpi_card(icono, titulo, valor, subtitulo=""):
    tonos = {
        "🛡️": "azul",
        "👷": "verde",
        "✅": "morado",
        "⚠️": "ambar",
        "🚨": "rojo",
        "📋": "azul",
        "📈": "verde",
        "🎓": "morado",
        "🦺": "ambar",
        "🔒": "celeste",
        "📁": "azul",
        "📦": "verde",
        "👥": "celeste",
        "🟢": "verde",
        "🟠": "ambar",
        "🔴": "rojo",
        "🔎": "azul",
        "📝": "morado",
        "📌": "ambar",
        "📜": "celeste",
        "🏆": "ambar",
        "🏢": "verde",
        "🗂️": "celeste",
        "❌": "rojo",
    }
    tono = tonos.get(icono, "azul")
    st.markdown(
        f"""
<div class="kpi-card notranslate" translate="no">
    <div class="kpi-icon {tono}" translate="no">{icono}</div>
    <div class="kpi-title notranslate" translate="no">{escape_html(titulo)}</div>
    <div class="kpi-value notranslate" translate="no">{escape_html(valor)}</div>
    <div class="kpi-sub notranslate" translate="no">{escape_html(subtitulo)}</div>
</div>
        """,
        unsafe_allow_html=True,
    )


def badge_estado(estado):
    estado_norm = normalizar_texto(estado)
    clase = "badge-neutro"
    if "cerr" in estado_norm or "cumpl" in estado_norm or "vigente" in estado_norm or "realiz" in estado_norm:
        clase = "badge-ok"
    elif "proceso" in estado_norm:
        clase = "badge-proceso"
    elif "pend" in estado_norm or "abiert" in estado_norm:
        clase = "badge-pendiente"
    elif "venc" in estado_norm or "no_cumple" in estado_norm:
        clase = "badge-vencida"
    return f"<span class='badge {clase}'>{escape_html(estado)}</span>"


def panel_titulo(texto):
    st.markdown(f"<div class='panel-title'>{texto}</div>", unsafe_allow_html=True)


def card_inicio():
    st.markdown("<div class='panel-card'>", unsafe_allow_html=True)


def card_fin():
    st.markdown("</div>", unsafe_allow_html=True)


# =========================================================
# GRÁFICOS
# =========================================================

PALETA_VERDE = ["#70D6A0", "#087B5B", "#35B779", "#A8E3C0", "#0B5D46", "#C9EED7"]


def aplicar_layout_fig(fig, height=360):
    fig.update_layout(
        title=dict(text=""),
        height=height,
        margin=dict(l=12, r=12, t=24, b=16),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        font=dict(color="#D1FAE5", size=12),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.24,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(1,8,6,.34)",
            font=dict(color="#D1FAE5"),
        ),
        hoverlabel=dict(bgcolor="#06100D", font_color="#F4FFF9", bordercolor="#34D399"),
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(110,231,183,.16)", zeroline=False, color="#C9F7E3")
    fig.update_yaxes(showgrid=True, gridcolor="rgba(110,231,183,.16)", zeroline=False, color="#C9F7E3")
    return fig


def grafico_barra(df, columna, titulo, orientacion="v", top=12):
    if df is None or df.empty or columna not in df.columns:
        st.info("Sin datos para graficar.")
        return
    base = df.copy()
    base[columna] = base[columna].fillna("Sin dato").astype(str).replace("", "Sin dato")
    conteo = base[columna].value_counts().reset_index()
    conteo.columns = [columna, "Cantidad"]
    conteo = conteo.head(top)
    if orientacion == "h":
        fig = px.bar(conteo.sort_values("Cantidad"), x="Cantidad", y=columna, orientation="h", title=titulo, text="Cantidad")
    else:
        fig = px.bar(conteo, x=columna, y="Cantidad", title=titulo, text="Cantidad")
    fig.update_traces(textposition="outside", marker_color="#62C990")
    st.plotly_chart(aplicar_layout_fig(fig), use_container_width=True)


def grafico_donut(df, columna, titulo):
    if df is None or df.empty or columna not in df.columns:
        st.info("Sin datos para graficar.")
        return
    base = df.copy()
    base[columna] = base[columna].fillna("Sin dato").astype(str).replace("", "Sin dato")
    conteo = base[columna].value_counts().reset_index()
    conteo.columns = [columna, "Cantidad"]
    fig = px.pie(conteo, names=columna, values="Cantidad", title=titulo, hole=0.48, color_discrete_sequence=PALETA_VERDE)
    fig.update_traces(textinfo="percent+label")
    st.plotly_chart(aplicar_layout_fig(fig), use_container_width=True)


def grafico_tendencia(df, titulo):
    if df is None or df.empty or "Fecha" not in df.columns:
        st.info("Sin datos para graficar.")
        return
    base = df.copy()
    base["Fecha"] = base["Fecha"].apply(convertir_fecha)
    base = base[base["Fecha"].notna()].copy()
    if base.empty:
        st.info("Sin fechas válidas para graficar.")
        return
    base["Periodo_Mes"] = base["Fecha"].dt.to_period("M").dt.to_timestamp()
    conteo = base.groupby("Periodo_Mes", as_index=False).size()
    conteo = conteo.rename(columns={"size": "Cantidad"})
    fig = px.line(conteo, x="Periodo_Mes", y="Cantidad", title=titulo, markers=True, color_discrete_sequence=PALETA_VERDE)
    fig.update_xaxes(title="Mes")
    fig.update_yaxes(title="Cantidad")
    st.plotly_chart(aplicar_layout_fig(fig), use_container_width=True)


# =========================================================
# PÁGINAS
# =========================================================

def resumen_mensual_cumplimientos(df):
    filas = []
    for mes in MESES_CORTOS:
        meta = pd.to_numeric(df.get(mes, 0), errors="coerce").fillna(0).sum()
        real = pd.to_numeric(df.get(f"RE_{mes}", 0), errors="coerce").fillna(0).sum()
        cumplimiento = (real / meta * 100) if meta > 0 else 0
        filas.append({"Mes": mes, "Meta": meta, "Realizadas": real, "Cumplimiento": cumplimiento})
    return pd.DataFrame(filas)


def pagina_panel_general(datos, filtros):
    sello_agua = obtener_sello_agua_html()
    if sello_agua:
        st.markdown(sello_agua, unsafe_allow_html=True)

    cumplimiento_df = datos.get("Cumplimientos_SSO", pd.DataFrame()).copy()
    incidentes = datos.get("Incidentes", pd.DataFrame()).copy()
    capacitaciones = datos.get("Capacitaciones", pd.DataFrame()).copy()
    config = datos.get("Configuracion", pd.DataFrame())

    dias, fecha_inicio = dias_sin_accidentes(config, incidentes)
    resumen = resumen_mensual_cumplimientos(cumplimiento_df) if not cumplimiento_df.empty else pd.DataFrame()
    resumen_corte = (
        resumen[resumen["Mes"].isin(MESES_CUMPLIMIENTOS)].copy()
        if not resumen.empty
        else pd.DataFrame()
    )

    # Todos los KPI de cumplimiento del panel general usan el mismo corte
    # oficial que la página Cumplimientos SSO: enero a julio de 2026.
    meta_corte = float(resumen_corte["Meta"].sum()) if not resumen_corte.empty else 0
    real_corte = float(resumen_corte["Realizadas"].sum()) if not resumen_corte.empty else 0
    porc_corte = (real_corte / meta_corte * 100) if meta_corte else 0
    observadores = cumplimiento_df["Observador"].nunique() if not cumplimiento_df.empty else 0

    accidentes = 0
    if not incidentes.empty and "Tipo_Evento" in incidentes.columns:
        tipos = incidentes["Tipo_Evento"].apply(normalizar_texto)
        accidentes = int((tipos.str.contains("accidente", na=False) & ~tipos.str.contains("cuasi", na=False)).sum())

    c1, c2, c3, c4, c5 = st.columns(5, gap="medium")
    with c1:
        kpi_card("🛡️", "Días sin accidentes", numero(dias), f"Desde {fecha_texto(fecha_inicio)}")
    with c2:
        kpi_card("👷", "Observadores", numero(observadores), "Personas con meta asignada")
    with c3:
        kpi_card("✅", "Actividades realizadas", numero(real_corte), f"Meta ene-jul: {numero(meta_corte)}")
    with c4:
        kpi_card("📈", "% cumplimiento", f"{porc_corte:.0f}%", PERIODO_CUMPLIMIENTOS)
    with c5:
        kpi_card("🚨", "Accidentes registrados", numero(accidentes), "Registros del año")

    col1, col2 = st.columns([1.15, 1], gap="large")
    with col1:
        panel_titulo("Meta versus resultado mensual")
        if resumen.empty:
            st.info("Sin datos en la hoja Cumplimientos SSO.")
        else:
            fig = go.Figure()
            fig.add_bar(x=resumen_corte["Mes"], y=resumen_corte["Meta"], name="Meta")
            fig.add_bar(x=resumen_corte["Mes"], y=resumen_corte["Realizadas"], name="Realizadas")
            fig.update_layout(barmode="group", xaxis_title="Mes", yaxis_title="Cantidad")
            st.plotly_chart(aplicar_layout_fig(fig, height=390), use_container_width=True)

    with col2:
        panel_titulo("Cumplimiento mensual")
        if resumen.empty:
            st.info("Sin datos para graficar.")
        else:
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=resumen_corte["Mes"], y=resumen_corte["Cumplimiento"],
                mode="lines+markers+text",
                text=[f"{v:.0f}%" for v in resumen_corte["Cumplimiento"]],
                textposition="top center", name="Cumplimiento",
            ))
            fig.add_hline(y=100, line_dash="dash", annotation_text="Meta 100%")
            fig.update_yaxes(title="Cumplimiento (%)", rangemode="tozero")
            st.plotly_chart(aplicar_layout_fig(fig, height=390), use_container_width=True)

    panel_titulo("Resumen por observador")
    if cumplimiento_df.empty:
        st.info("Sin registros para mostrar.")
    else:
        filas = []
        for observador, grupo in cumplimiento_df.groupby("Observador", dropna=False):
            meta = sum(
                pd.to_numeric(grupo[m], errors="coerce").fillna(0).sum()
                for m in MESES_CUMPLIMIENTOS
            )
            real = sum(
                pd.to_numeric(grupo[f"RE_{m}"], errors="coerce").fillna(0).sum()
                for m in MESES_CUMPLIMIENTOS
            )
            filas.append({
                "Observador": observador,
                "Meta ene-jul": meta,
                "Realizadas": real,
                "Cumplimiento": f"{(real / meta * 100) if meta else 0:.0f}%",
            })
        tabla_limpia(
            pd.DataFrame(filas),
            ["Observador", "Meta ene-jul", "Realizadas", "Cumplimiento"],
            centrar_todo=True,
        )

    # =============================================================
    # INFORMACIÓN AGREGADA: CERTIFICACIONES Y RECONOCIMIENTOS
    # =============================================================
    certificaciones_df = datos.get("Certificaciones", pd.DataFrame()).copy()
    reconocimientos_df = datos.get("Reconocimientos", pd.DataFrame()).copy()

    if certificaciones_df is None:
        certificaciones_df = pd.DataFrame()
    if reconocimientos_df is None:
        reconocimientos_df = pd.DataFrame()

    # Conserva solamente registros reales de certificaciones.
    if not certificaciones_df.empty:
        columnas_clave_cert = [
            columna
            for columna in [
                "Categoria",
                "Subcategoria",
                "Nombre_Certificacion",
                "Vencimiento",
            ]
            if columna in certificaciones_df.columns
        ]
        if columnas_clave_cert:
            mascara_cert = pd.Series(False, index=certificaciones_df.index)
            for columna in columnas_clave_cert:
                mascara_cert = mascara_cert | (
                    certificaciones_df[columna]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .ne("")
                )
            certificaciones_df = certificaciones_df.loc[mascara_cert].copy()

    total_certificaciones = len(certificaciones_df)
    vigentes_certificaciones = 0
    por_vencer_certificaciones = 0
    vencidas_certificaciones = 0

    if not certificaciones_df.empty:
        if "Estado" in certificaciones_df.columns:
            estados_cert = certificaciones_df["Estado"].fillna("").apply(estado_base)
        else:
            estados_cert = pd.Series("Sin estado", index=certificaciones_df.index)

        vigentes_certificaciones = int((estados_cert == "Vigente").sum())
        por_vencer_certificaciones = int((estados_cert == "Por vencer").sum())
        vencidas_certificaciones = int((estados_cert == "Vencida").sum())
        certificaciones_df["Estado"] = estados_cert

    cobertura_certificaciones = (
        (vigentes_certificaciones + por_vencer_certificaciones)
        / total_certificaciones
        * 100
        if total_certificaciones
        else 0
    )

    proximo_dias = None
    proximo_elemento = "Sin vencimientos registrados"
    proximos_vencimientos = pd.DataFrame()

    if not certificaciones_df.empty and "Vencimiento" in certificaciones_df.columns:
        certificaciones_df["Vencimiento"] = certificaciones_df["Vencimiento"].apply(convertir_fecha)
        certificaciones_df["Días restantes"] = certificaciones_df["Vencimiento"].apply(
            lambda fecha: int((fecha.normalize() - HOY).days) if pd.notna(fecha) else pd.NA
        )
        proximos_vencimientos = certificaciones_df[
            certificaciones_df["Días restantes"].notna()
            & (certificaciones_df["Días restantes"] >= 0)
        ].copy()
        proximos_vencimientos = proximos_vencimientos.sort_values(
            "Días restantes",
            ascending=True,
        )

        if not proximos_vencimientos.empty:
            primera_cert = proximos_vencimientos.iloc[0]
            proximo_dias = int(primera_cert["Días restantes"])
            proximo_elemento = str(
                primera_cert.get("Subcategoria", "")
                or primera_cert.get("Nombre_Certificacion", "")
                or "Certificación"
            ).strip()

    # Reconocimientos del año 2026, alineados con el periodo del panel KPI.
    if not reconocimientos_df.empty and "Fecha" in reconocimientos_df.columns:
        reconocimientos_df["Fecha"] = reconocimientos_df["Fecha"].apply(convertir_fecha)
        reconocimientos_2026 = reconocimientos_df[
            reconocimientos_df["Fecha"].dt.year.eq(2026)
        ].copy()
    else:
        reconocimientos_2026 = reconocimientos_df.copy()

    # Elimina filas visualmente vacías de la hoja.
    if not reconocimientos_2026.empty:
        columnas_clave_rec = [
            columna
            for columna in ["Trabajador", "Motivo", "Periodo", "Estado"]
            if columna in reconocimientos_2026.columns
        ]
        if columnas_clave_rec:
            mascara_rec = pd.Series(False, index=reconocimientos_2026.index)
            for columna in columnas_clave_rec:
                mascara_rec = mascara_rec | (
                    reconocimientos_2026[columna]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .ne("")
                )
            reconocimientos_2026 = reconocimientos_2026.loc[mascara_rec].copy()

    total_reconocimientos = len(reconocimientos_2026)
    reconocimientos_entregados = 0
    reconocimientos_corporativos = 0
    personas_reconocidas = 0

    if not reconocimientos_2026.empty:
        if "Estado" in reconocimientos_2026.columns:
            reconocimientos_entregados = int(
                reconocimientos_2026["Estado"]
                .fillna("")
                .astype(str)
                .str.contains("cerrada|entregada|realizada", case=False, regex=True, na=False)
                .sum()
            )

        es_corporativo = pd.Series(False, index=reconocimientos_2026.index)
        for columna in ["Trabajador", "Cargo", "Motivo", "Observacion"]:
            if columna in reconocimientos_2026.columns:
                texto_columna = reconocimientos_2026[columna].fillna("").apply(normalizar_texto)
                es_corporativo = es_corporativo | texto_columna.str.contains(
                    "empresa_saivam|reconocimiento_institucional|equipo_saivam",
                    regex=True,
                    na=False,
                )

        reconocimientos_corporativos = int(es_corporativo.sum())

        if "Trabajador" in reconocimientos_2026.columns:
            personas_reconocidas = int(
                reconocimientos_2026.loc[~es_corporativo, "Trabajador"]
                .replace("", pd.NA)
                .dropna()
                .nunique()
            )

    panel_titulo("Certificaciones y reconocimientos")

    st.markdown(
        "<div class='subsection-label notranslate' translate='no'>Estado de certificaciones</div>",
        unsafe_allow_html=True,
    )
    cc1, cc2, cc3, cc4 = st.columns(4, gap="medium")
    with cc1:
        kpi_card(
            "📜",
            "Certificaciones controladas",
            numero(total_certificaciones),
            f"{cobertura_certificaciones:.0f}% con vigencia",
        )
    with cc2:
        kpi_card(
            "✅",
            "Certificaciones vigentes",
            numero(vigentes_certificaciones),
            "Vigencia superior a 30 días",
        )
    with cc3:
        kpi_card(
            "⚠️",
            "Por vencer",
            numero(por_vencer_certificaciones),
            "Vencen dentro de 30 días",
        )
    with cc4:
        valor_proximo = f"{proximo_dias} días" if proximo_dias is not None else "—"
        kpi_card(
            "📌",
            "Próximo vencimiento",
            valor_proximo,
            proximo_elemento,
        )

    st.markdown(
        "<div class='subsection-label notranslate' translate='no'>Reconocimientos 2026</div>",
        unsafe_allow_html=True,
    )
    cr1, cr2, cr3, cr4 = st.columns(4, gap="medium")
    with cr1:
        kpi_card(
            "🏆",
            "Reconocimientos 2026",
            numero(total_reconocimientos),
            "Registros del año",
        )
    with cr2:
        kpi_card(
            "👷",
            "Personas reconocidas",
            numero(personas_reconocidas),
            "Trabajadores destacados",
        )
    with cr3:
        kpi_card(
            "🏢",
            "Reconocimientos corporativos",
            numero(reconocimientos_corporativos),
            "Reconocimientos a SAIVAM",
        )
    with cr4:
        kpi_card(
            "✅",
            "Reconocimientos entregados",
            numero(reconocimientos_entregados),
            "Registros cerrados",
        )

    col_cert, col_rec = st.columns(2, gap="large")

    with col_cert:
        panel_titulo("Certificaciones por categoría")
        if certificaciones_df.empty or "Categoria" not in certificaciones_df.columns:
            st.info("Sin datos de certificaciones para graficar.")
        else:
            categorias_cert = (
                certificaciones_df["Categoria"]
                .fillna("Sin categoría")
                .astype(str)
                .replace("", "Sin categoría")
                .value_counts()
                .rename_axis("Categoría")
                .reset_index(name="Cantidad")
            )
            fig_cert = px.bar(
                categorias_cert.sort_values("Cantidad"),
                x="Cantidad",
                y="Categoría",
                orientation="h",
                text="Cantidad",
            )
            fig_cert.update_traces(textposition="outside", marker_color="#62C990")
            st.plotly_chart(
                aplicar_layout_fig(fig_cert, height=330),
                use_container_width=True,
            )

    with col_rec:
        panel_titulo("Reconocimientos por mes")
        if reconocimientos_2026.empty or "Fecha" not in reconocimientos_2026.columns:
            st.info("Sin datos de reconocimientos para graficar.")
        else:
            reconocimientos_2026["Mes_Numero_KPI"] = reconocimientos_2026["Fecha"].dt.month
            resumen_rec_mensual = (
                reconocimientos_2026
                .groupby("Mes_Numero_KPI", as_index=False)
                .size()
                .rename(columns={"size": "Cantidad"})
            )
            resumen_rec_mensual["Mes"] = resumen_rec_mensual["Mes_Numero_KPI"].map(MESES)
            resumen_rec_mensual = resumen_rec_mensual.sort_values("Mes_Numero_KPI")
            fig_rec = px.bar(
                resumen_rec_mensual,
                x="Mes",
                y="Cantidad",
                text="Cantidad",
            )
            fig_rec.update_traces(textposition="outside", marker_color="#62C990")
            st.plotly_chart(
                aplicar_layout_fig(fig_rec, height=330),
                use_container_width=True,
            )

    col_vencimientos, col_ultimos = st.columns(2, gap="large")

    with col_vencimientos:
        panel_titulo("Próximos vencimientos")
        if proximos_vencimientos.empty:
            st.info("Sin vencimientos próximos registrados.")
        else:
            tabla_vencimientos = proximos_vencimientos.head(5).copy()
            tabla_limpia(
                tabla_vencimientos,
                [
                    "Categoria",
                    "Subcategoria",
                    "Vencimiento",
                    "Días restantes",
                    "Estado",
                ],
                centrar_todo=True,
            )

    with col_ultimos:
        panel_titulo("Últimos reconocimientos")
        if reconocimientos_2026.empty:
            st.info("Sin reconocimientos registrados durante 2026.")
        else:
            ultimos_reconocimientos = reconocimientos_2026.sort_values(
                "Fecha",
                ascending=False,
            ).head(5)
            tabla_limpia(
                ultimos_reconocimientos,
                ["Fecha", "Trabajador", "Motivo", "Periodo", "Estado"],
                centrar_todo=True,
            )


def pagina_reportabilidad(datos, filtros):
    """
    Panel de Reportabilidad basado en las columnas:

    Fecha, Área, Tipo_Evento, Descripcion, Accion_Inmediata,
    Responsable, Estado, Observacion y Ruta_Link.
    """
    mostrar_sello_saivam_pagina()

    # Usa .get() para que el módulo no se caiga si la pestaña todavía
    # no está disponible o existe un problema temporal de conexión.
    df = aplicar_filtros(
        datos.get("Incidentes", pd.DataFrame()),
        *filtros,
    )

    if df is None:
        df = pd.DataFrame()

    if df.empty:
        st.warning(
            "No se encontraron registros en la pestaña Reportabilidad. "
            "Verifique que el Google Sheet esté compartido como lector "
            "mediante enlace y que la pestaña contenga registros válidos."
        )

    # Asegura compatibilidad aunque una columna todavía no exista
    # en la pestaña de Google Sheets.
    columnas_requeridas = [
        "Fecha",
        "Área",
        "Tipo_Evento",
        "Descripcion",
        "Accion_Inmediata",
        "Responsable",
        "Estado",
        "Observacion",
        "Ruta_Link",
    ]

    for columna in columnas_requeridas:
        if columna not in df.columns:
            df[columna] = ""

    estados = (
        df["Estado"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    pendientes = int(
        estados.str.contains(
            "Pendiente|Vencida",
            case=False,
            regex=True,
            na=False,
        ).sum()
    )

    en_proceso = int(
        estados.str.contains(
            "En proceso|En gestión|En gestion",
            case=False,
            regex=True,
            na=False,
        ).sum()
    )

    cerrados = int(
        estados.str.contains(
            "Cerrada|Cerrado|Finalizada|Finalizado",
            case=False,
            regex=True,
            na=False,
        ).sum()
    )

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi_card(
            "📝",
            "Eventos registrados",
            numero(len(df)),
            "Total de reportes",
        )

    with c2:
        kpi_card(
            "⚠️",
            "Pendientes",
            numero(pendientes),
            "Requieren gestión",
        )

    with c3:
        kpi_card(
            "🔄",
            "En proceso",
            numero(en_proceso),
            "Actualmente en seguimiento",
        )

    with c4:
        kpi_card(
            "✅",
            "Cerrados",
            numero(cerrados),
            "Reportes finalizados",
        )

    # --------------------------------------------------------------
    # GRÁFICOS PRINCIPALES
    # --------------------------------------------------------------
    col_a, col_b = st.columns(2)

    with col_a:
        card_inicio()
        grafico_barra(
            df,
            "Tipo_Evento",
            "Eventos por tipo",
            orientacion="h",
        )
        card_fin()

    with col_b:
        card_inicio()
        grafico_barra(
            df,
            "Área",
            "Eventos por área",
            orientacion="h",
        )
        card_fin()

    col_c, col_d = st.columns(2)

    with col_c:
        card_inicio()
        grafico_donut(
            df,
            "Estado",
            "Estado de reportabilidad",
        )
        card_fin()

    with col_d:
        card_inicio()
        grafico_tendencia(
            df,
            "Tendencia mensual de eventos",
        )
        card_fin()

    # --------------------------------------------------------------
    # DETALLE
    # --------------------------------------------------------------
    panel_titulo("Detalle de reportabilidad")

    tabla_limpia(
        df,
        [
            "Fecha",
            "Área",
            "Tipo_Evento",
            "Descripcion",
            "Accion_Inmediata",
            "Responsable",
            "Estado",
            "Observacion",
            "Ruta_Link",
        ],
        height=430,
    )

    # --------------------------------------------------------------
    # ALERTAS DE SEGUIMIENTO
    # --------------------------------------------------------------
    abiertos = df[
        df["Estado"]
        .fillna("")
        .astype(str)
        .str.contains(
            "Pendiente|En proceso|Vencida",
            case=False,
            regex=True,
            na=False,
        )
    ].copy()

    if not abiertos.empty:
        panel_titulo("Eventos que requieren seguimiento")

        for _, fila in abiertos.iterrows():
            tipo_evento = str(
                fila.get("Tipo_Evento", "Evento")
            ).strip() or "Evento"

            area = str(
                fila.get("Área", "Sin área")
            ).strip() or "Sin área"

            responsable = str(
                fila.get("Responsable", "Sin responsable")
            ).strip() or "Sin responsable"

            estado = str(
                fila.get("Estado", "Pendiente")
            ).strip() or "Pendiente"

            observacion = str(
                fila.get("Observacion", "")
            ).strip()

            fecha = fecha_texto(
                fila.get("Fecha", pd.NaT)
            )

            detalle_secundario = (
                f"{area} · Responsable: {responsable}"
            )

            if fecha:
                detalle_secundario += f" · {fecha}"

            if observacion:
                detalle_secundario += f" · {observacion}"

            st.markdown(
                f"""
<div class="alert-card">
    <div class="alert-card-title">
        ⚠️ {escape_html(tipo_evento)}
        <span>{escape_html(estado)}</span>
    </div>
    <div class="alert-card-text">
        {escape_html(detalle_secundario)}
    </div>
</div>
                """,
                unsafe_allow_html=True,
            )



def pagina_cumplimientos_sso(datos, filtros):
    """Dashboard ejecutivo de cumplimiento de actividades SSO."""
    mostrar_sello_saivam_pagina()

    df = datos.get("Cumplimientos_SSO", pd.DataFrame()).copy()

    if df is None or df.empty:
        st.warning(
            "No fue posible descargar la pestaña 'Cumplimientos SSO'. "
            "El GID configurado es correcto, pero Google no está entregando el contenido CSV."
        )
        st.markdown(
            """
**Configuración necesaria en Google Sheets**  
1. Presiona **Compartir**.  
2. En **Acceso general**, selecciona **Cualquier persona con el enlace**.  
3. Selecciona el permiso **Lector** y guarda.  
4. Reinicia Streamlit o presiona **Actualizar Base de Datos**.

El enlace puede abrir normalmente en Chrome porque tu cuenta de Google está
iniciada, pero el proceso de Python no utiliza esa sesión.
            """
        )

        detalle_error = st.session_state.get("error_cumplimientos_google", "")
        if detalle_error:
            with st.expander("Ver detalle técnico de la conexión"):
                st.code(detalle_error, language="text")

        st.info(
            "Como respaldo, también puedes guardar el archivo "
            "'Base_Datos_SGS_SAIVAM_Mulchen.xlsx' junto a app.py."
        )
        return

    # ------------------------------------------------------------------
    # Preparación y homologación de actividades
    # ------------------------------------------------------------------
    def homologar_actividad(nombre):
        clave = normalizar_texto(nombre)
        if "control_operacional" in clave and "cmpc" in clave:
            return "Control operacional CMPC"
        if "control_operacional" in clave and "saivam" in clave:
            return "Control operacional SAIVAM"
        if "ops" in clave and "bapp" in clave:
            return "OPS BAPP"
        if "ops" in clave and "cmpc" in clave:
            return "OPS de Seguridad CMPC"
        if "inspeccion" in clave or "check_list" in clave or "checklist" in clave:
            return "Inspecciones / Check List SAIVAM"
        if "observacion" in clave:
            return "Observaciones de Seguridad SAIVAM"
        return str(nombre).strip() or "Otra actividad"

    orden_actividades = [
        "Control operacional CMPC",
        "Control operacional SAIVAM",
        "OPS de Seguridad CMPC",
        "OPS BAPP",
        "Inspecciones / Check List SAIVAM",
        "Observaciones de Seguridad SAIVAM",
    ]

    df["Actividad estándar"] = df["Actividad"].apply(homologar_actividad)
    for columna in MESES_CORTOS + [f"RE_{mes}" for mes in MESES_CORTOS]:
        if columna not in df.columns:
            df[columna] = 0
        df[columna] = df[columna].apply(_a_numero_cumplimiento)

    df["Observador"] = df["Observador"].fillna("").astype(str).str.strip()
    df = df[df["Observador"].ne("")].copy()

    st.markdown(
        """
        <style>
        .cumplimiento-hero {
            padding: 22px 24px;
            border: 1px solid rgba(52, 211, 153, .25);
            border-radius: 18px;
            background: linear-gradient(135deg, rgba(15,23,42,.98), rgba(6,78,59,.34));
            margin-bottom: 14px;
        }
        .cumplimiento-hero h2 { margin: 0; color: #f8fafc; font-size: 28px; }
        .cumplimiento-hero p { margin: 7px 0 0 0; color: #a7b0bf; font-size: 13px; }
        .persona-card {
            border: 1px solid rgba(148,163,184,.20);
            border-radius: 15px;
            padding: 14px 16px;
            background: rgba(15,23,42,.68);
            margin-bottom: 10px;
        }
        .persona-title { color:#f8fafc; font-weight:800; font-size:15px; margin-bottom:8px; }
        .persona-meta { color:#94a3b8; font-size:12px; margin-top:7px; }
        .progress-track { width:100%; height:10px; border-radius:99px; background:rgba(100,116,139,.28); overflow:hidden; }
        .progress-fill { height:100%; border-radius:99px; }
        .activity-pill {
            display:inline-block; padding:5px 9px; border-radius:999px;
            background:rgba(16,185,129,.12); border:1px solid rgba(52,211,153,.28);
            color:#a7f3d0; font-size:11px; font-weight:700; margin:2px 3px 2px 0;
        }
        </style>
        <div class="cumplimiento-hero">
            <h2>🎯 Cumplimiento SSO 2026</h2>
            <p>Seguimiento ejecutivo de actividades preventivas por colaborador, actividad y mes.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ------------------------------------------------------------------
    # Filtros
    # ------------------------------------------------------------------
    observadores = sorted(df["Observador"].unique().tolist())
    actividades_disponibles = [
        actividad for actividad in orden_actividades
        if actividad in df["Actividad estándar"].unique()
    ]
    actividades_extra = sorted(
        set(df["Actividad estándar"].unique()) - set(actividades_disponibles)
    )
    actividades_disponibles += actividades_extra

    f1, f2, f3 = st.columns([1.1, 1.55, .75], gap="medium")
    with f1:
        seleccion_observador = st.selectbox(
            "Colaborador",
            ["Todos"] + observadores,
            key="cumplimientos_observador_v3",
        )
    with f2:
        seleccion_actividad = st.multiselect(
            "Actividades",
            actividades_disponibles,
            default=actividades_disponibles,
            key="cumplimientos_actividad_v3",
        )
    with f3:
        seleccion_periodo = st.selectbox(
            "Periodo considerado",
            [PERIODO_CUMPLIMIENTOS, PERIODO_ANUAL_CUMPLIMIENTOS],
            index=0,
            key="cumplimientos_periodo_v5",
        )

    if seleccion_periodo == PERIODO_ANUAL_CUMPLIMIENTOS:
        meses_periodo = list(MESES_CORTOS)
        detalle_periodo = (
            "Cálculo anual: total de actividades realizadas entre ENE y DIC "
            "dividido por el total programado para todo 2026."
        )
    else:
        meses_periodo = list(MESES_CUMPLIMIENTOS)
        detalle_periodo = (
            "Cálculo oficial: total de actividades realizadas entre ENE y JUL "
            "dividido por el total programado del mismo periodo. "
            "Los registros de AGO a DIC no afectan estos indicadores."
        )

    st.caption(detalle_periodo)

    filtrado = df.copy()
    if seleccion_observador != "Todos":
        filtrado = filtrado[filtrado["Observador"] == seleccion_observador]
    if seleccion_actividad:
        filtrado = filtrado[filtrado["Actividad estándar"].isin(seleccion_actividad)]
    else:
        filtrado = filtrado.iloc[0:0]

    if filtrado.empty:
        st.info("No existen registros para los filtros seleccionados.")
        return

    # Periodo dinámico según la selección: enero-julio o año completo.
    reales_periodo = [f"RE_{mes}" for mes in meses_periodo]

    meta_total = float(filtrado[meses_periodo].sum().sum())
    real_total = float(filtrado[reales_periodo].sum().sum())
    pendientes = max(0.0, meta_total - real_total)
    cumplimiento_total = (real_total / meta_total * 100) if meta_total else 0.0

    def color_cumplimiento(valor):
        if valor >= 100:
            return "#22c55e"
        if valor >= 90:
            return "#10b981"
        if valor >= 80:
            return "#eab308"
        if valor >= 70:
            return "#f97316"
        return "#ef4444"

    # ------------------------------------------------------------------
    # KPI ejecutivos
    # ------------------------------------------------------------------
    k1, k2, k3, k4 = st.columns(4, gap="medium")
    with k1:
        kpi_card("📈", "Cumplimiento", f"{cumplimiento_total:.0f}%", seleccion_periodo)
    with k2:
        kpi_card("✅", "Realizadas", numero(real_total), "Actividades ejecutadas")
    with k3:
        kpi_card("🎯", "Programadas", numero(meta_total), "Meta del periodo")
    with k4:
        kpi_card("⏳", "Pendientes", numero(pendientes), "Brecha respecto de la meta")

    # ------------------------------------------------------------------
    # Resúmenes consolidados
    # ------------------------------------------------------------------
    filas_persona = []
    for observador, grupo in filtrado.groupby("Observador", dropna=False):
        meta = float(grupo[meses_periodo].sum().sum())
        real = float(grupo[reales_periodo].sum().sum())
        filas_persona.append({
            "Colaborador": observador,
            "Meta": meta,
            "Realizadas": real,
            "Pendientes": max(0.0, meta - real),
            "Cumplimiento": (real / meta * 100) if meta else 0.0,
        })
    por_persona = pd.DataFrame(filas_persona).sort_values(
        ["Cumplimiento", "Realizadas"], ascending=[False, False]
    )

    filas_actividad = []
    for actividad, grupo in filtrado.groupby("Actividad estándar", dropna=False):
        meta = float(grupo[meses_periodo].sum().sum())
        real = float(grupo[reales_periodo].sum().sum())
        filas_actividad.append({
            "Actividad": actividad,
            "Meta": meta,
            "Realizadas": real,
            "Pendientes": max(0.0, meta - real),
            "Cumplimiento": (real / meta * 100) if meta else 0.0,
        })
    por_actividad = pd.DataFrame(filas_actividad)
    por_actividad["_orden"] = por_actividad["Actividad"].apply(
        lambda x: orden_actividades.index(x) if x in orden_actividades else 99
    )
    por_actividad = por_actividad.sort_values(["_orden", "Actividad"]).drop(columns="_orden")

    # ------------------------------------------------------------------
    # Vista principal: ranking y actividad
    # ------------------------------------------------------------------
    izquierda, derecha = st.columns([1.05, 1.25], gap="large")

    with izquierda:
        panel_titulo("Ranking de cumplimiento por colaborador")
        ranking = por_persona.sort_values("Cumplimiento", ascending=True)
        colores = [color_cumplimiento(v) for v in ranking["Cumplimiento"]]
        fig = go.Figure(go.Bar(
            x=ranking["Cumplimiento"],
            y=ranking["Colaborador"],
            orientation="h",
            text=[f"{v:.0f}%" for v in ranking["Cumplimiento"]],
            textposition="outside",
            marker_color=colores,
            customdata=ranking[["Meta", "Realizadas", "Pendientes"]],
            hovertemplate=(
                "%{y}<br>Cumplimiento: %{x:.0f}%"
                "<br>Meta: %{customdata[0]:.0f}"
                "<br>Realizadas: %{customdata[1]:.0f}"
                "<br>Pendientes: %{customdata[2]:.0f}<extra></extra>"
            ),
        ))
        fig.add_vline(x=100, line_dash="dash", line_color="#94a3b8")
        fig.update_layout(title=dict(text=""), showlegend=False)
        fig.update_xaxes(title="Cumplimiento (%)", range=[0, max(110, ranking["Cumplimiento"].max() * 1.15)])
        fig.update_yaxes(title="", automargin=True)
        st.plotly_chart(
            aplicar_layout_fig(fig, height=max(360, 120 + len(ranking) * 48)),
            use_container_width=True,
            config={"displaylogo": False},
        )

    with derecha:
        panel_titulo("Cumplimiento por actividad")
        actividad_graf = por_actividad.sort_values("Cumplimiento", ascending=True)
        colores = [color_cumplimiento(v) for v in actividad_graf["Cumplimiento"]]
        fig = go.Figure(go.Bar(
            x=actividad_graf["Cumplimiento"],
            y=actividad_graf["Actividad"],
            orientation="h",
            text=[f"{v:.0f}%" for v in actividad_graf["Cumplimiento"]],
            textposition="outside",
            marker_color=colores,
            customdata=actividad_graf[["Meta", "Realizadas", "Pendientes"]],
            hovertemplate=(
                "%{y}<br>Cumplimiento: %{x:.0f}%"
                "<br>Meta: %{customdata[0]:.0f}"
                "<br>Realizadas: %{customdata[1]:.0f}"
                "<br>Pendientes: %{customdata[2]:.0f}<extra></extra>"
            ),
        ))
        fig.add_vline(x=100, line_dash="dash", line_color="#94a3b8")
        fig.update_layout(title=dict(text=""), showlegend=False)
        fig.update_xaxes(title="Cumplimiento (%)", range=[0, max(110, actividad_graf["Cumplimiento"].max() * 1.15)])
        fig.update_yaxes(title="", automargin=True)
        st.plotly_chart(
            aplicar_layout_fig(fig, height=max(360, 120 + len(actividad_graf) * 48)),
            use_container_width=True,
            config={"displaylogo": False},
        )

    # ------------------------------------------------------------------
    # Tendencia mensual y distribución
    # ------------------------------------------------------------------
    resumen_mes = []
    for mes in MESES_CORTOS:
        meta = float(filtrado[mes].sum())
        real = float(filtrado[f"RE_{mes}"].sum())
        resumen_mes.append({
            "Mes": mes,
            "Meta": meta,
            "Realizadas": real,
            "Cumplimiento": (real / meta * 100) if meta else 0.0,
        })
    resumen_mes = pd.DataFrame(resumen_mes)
    resumen_visible = resumen_mes[resumen_mes["Mes"].isin(meses_periodo)].copy()

    c1, c2 = st.columns([1.35, .85], gap="large")
    with c1:
        panel_titulo("Evolución mensual")
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=resumen_visible["Mes"], y=resumen_visible["Meta"],
            name="Programadas", opacity=.45,
            hovertemplate="%{x}<br>Programadas: %{y:.0f}<extra></extra>",
        ))
        fig.add_trace(go.Bar(
            x=resumen_visible["Mes"], y=resumen_visible["Realizadas"],
            name="Realizadas",
            hovertemplate="%{x}<br>Realizadas: %{y:.0f}<extra></extra>",
        ))
        fig.add_trace(go.Scatter(
            x=resumen_visible["Mes"], y=resumen_visible["Cumplimiento"],
            name="Cumplimiento %", mode="lines+markers+text", yaxis="y2",
            text=[f"{v:.0f}%" for v in resumen_visible["Cumplimiento"]],
            textposition="top center",
            hovertemplate="%{x}<br>Cumplimiento: %{y:.0f}%<extra></extra>",
        ))
        fig.update_layout(
            barmode="group",
            yaxis=dict(title="Cantidad"),
            yaxis2=dict(title="Cumplimiento (%)", overlaying="y", side="right", rangemode="tozero"),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
        )
        st.plotly_chart(
            aplicar_layout_fig(fig, height=410),
            use_container_width=True,
            config={"displaylogo": False},
        )

    with c2:
        panel_titulo("Distribución de realizadas")
        fig = go.Figure(go.Pie(
            labels=por_actividad["Actividad"],
            values=por_actividad["Realizadas"],
            hole=.62,
            textinfo="percent",
            texttemplate="%{percent:.0%}",
            hovertemplate="%{label}<br>Realizadas: %{value:.0f}<br>%{percent:.0%}<extra></extra>",
        ))
        fig.update_layout(showlegend=True, legend=dict(orientation="h", y=-.15))
        st.plotly_chart(
            aplicar_layout_fig(fig, height=410),
            use_container_width=True,
            config={"displaylogo": False},
        )

    # ------------------------------------------------------------------
    # Tarjetas individuales
    # ------------------------------------------------------------------
    panel_titulo("Detalle por colaborador")
    columnas_tarjetas = st.columns(2, gap="medium")
    for indice, fila in por_persona.reset_index(drop=True).iterrows():
        valor = float(fila["Cumplimiento"])
        ancho = min(100.0, max(0.0, valor))
        color = color_cumplimiento(valor)
        with columnas_tarjetas[indice % 2]:
            st.markdown(
                f"""
                <div class="persona-card">
                    <div class="persona-title">👷 {escape_html(fila['Colaborador'])}</div>
                    <div class="progress-track">
                        <div class="progress-fill" style="width:{ancho:.1f}%; background:{color};"></div>
                    </div>
                    <div class="persona-meta">
                        <b style="color:{color};">{valor:.0f}%</b> · Meta {fila['Meta']:.0f} ·
                        Realizadas {fila['Realizadas']:.0f} · Pendientes {fila['Pendientes']:.0f}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # ------------------------------------------------------------------
    # Vistas de control
    # ------------------------------------------------------------------
    # Se utiliza un selector horizontal en lugar de st.tabs(). Esto evita el
    # error del navegador "removeChild" que puede aparecer cuando Chrome
    # Translate o alguna extensión modifica el DOM que Streamlit administra.
    vista_control = st.radio(
        "Vista de información",
        [
            "📋 Resumen por actividad",
            "🗓️ Matriz mensual",
            "🔥 Mapa de calor",
        ],
        horizontal=True,
        label_visibility="collapsed",
        key="vista_control_cumplimientos_sso",
    )

    if vista_control == "📋 Resumen por actividad":
        resumen_tabla = por_actividad.copy()
        resumen_tabla["Cumplimiento"] = resumen_tabla["Cumplimiento"].map(
            lambda valor: f"{float(valor):.0f}%"
        )
        tabla_limpia(
            resumen_tabla,
            ["Actividad", "Meta", "Realizadas", "Pendientes", "Cumplimiento"],
            centrar_todo=False,
        )

    elif vista_control == "🗓️ Matriz mensual":
        detalle = filtrado[["Observador", "Actividad estándar"]].copy()
        detalle = detalle.rename(columns={"Actividad estándar": "Actividad"})

        for mes in meses_periodo:
            metas_mes = pd.to_numeric(filtrado[mes], errors="coerce").fillna(0)
            realizadas_mes = pd.to_numeric(
                filtrado[f"RE_{mes}"], errors="coerce"
            ).fillna(0)

            detalle[mes] = [
                f"{real:.0f}/{meta:.0f} ({(real / meta * 100 if meta else 0):.0f}%)"
                for meta, real in zip(metas_mes, realizadas_mes)
            ]

        tabla_limpia(
            detalle,
            modo_ultracompacto=True,
            centrar_todo=True,
        )

    else:
        heat_rows = []
        heat_labels = []

        for _, fila in filtrado.iterrows():
            valores = []

            for mes in meses_periodo:
                meta = _a_numero_cumplimiento(fila.get(mes, 0))
                real = _a_numero_cumplimiento(fila.get(f"RE_{mes}", 0))
                cumplimiento_mes = (real / meta * 100) if meta > 0 else 0.0
                valores.append(cumplimiento_mes)

            heat_rows.append(valores)
            heat_labels.append(
                f"{fila['Observador']} · {fila['Actividad estándar']}"
            )

        if heat_rows:
            maximo_heatmap = max(max(fila) for fila in heat_rows)
            fig_heatmap = go.Figure(
                go.Heatmap(
                    z=heat_rows,
                    x=meses_periodo,
                    y=heat_labels,
                    zmin=0,
                    zmax=max(120, maximo_heatmap),
                    colorscale=[
                        [0.00, "#7f1d1d"],
                        [0.58, "#ef4444"],
                        [0.70, "#f97316"],
                        [0.80, "#eab308"],
                        [0.90, "#10b981"],
                        [1.00, "#22c55e"],
                    ],
                    colorbar=dict(title="Cumpl. %"),
                    hovertemplate=(
                        "%{y}<br>%{x}: %{z:.0f}%<extra></extra>"
                    ),
                )
            )
            fig_heatmap.update_layout(title=dict(text=""))
            fig_heatmap.update_xaxes(title="Mes")
            fig_heatmap.update_yaxes(title="", automargin=True)

            alto_heatmap = max(
                420,
                min(880, 160 + len(heat_labels) * 31),
            )

            st.plotly_chart(
                aplicar_layout_fig(fig_heatmap, height=alto_heatmap),
                use_container_width=True,
                config={"displaylogo": False},
                key="heatmap_cumplimientos_sso",
            )
        else:
            st.info("No existen datos para generar el mapa de calor.")

def pagina_ops(datos, filtros):
    """
    Módulo compatible con dos estructuras de la hoja Observaciones_SSO_BAPP:

    1. Formato histórico por registro individual.
    2. Formato anual por observador, con columnas Enero-Diciembre,
       Real Año, Teórica Año y Avance.

    La detección es automática, por lo que no es necesario cambiar el nombre
    de la hoja ni crear una segunda planilla.
    """
    df_original = datos["OPS"].copy()

    columnas_por_clave = {
        normalizar_texto(columna): columna
        for columna in df_original.columns
    }

    es_matriz_anual = (
        "observador" in columnas_por_clave
        and "tipo_observacion" in columnas_por_clave
        and "enero" in columnas_por_clave
        and "real_ano" in columnas_por_clave
    )

    if es_matriz_anual:
        meses = [
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre",
            "Diciembre",
        ]

        nombres_estandar = {
            "observador": "Observador",
            "tipo_observacion": "Tipo Observación",
            "enero": "Enero",
            "febrero": "Febrero",
            "marzo": "Marzo",
            "abril": "Abril",
            "mayo": "Mayo",
            "junio": "Junio",
            "julio": "Julio",
            "agosto": "Agosto",
            "septiembre": "Septiembre",
            "octubre": "Octubre",
            "noviembre": "Noviembre",
            "diciembre": "Diciembre",
            "real_ano": "Real Año",
            "teorica_ano": "Teórica Año",
            "avance": "Avance",
        }

        renombrar = {
            columna_original: nombres_estandar[clave]
            for clave, columna_original in columnas_por_clave.items()
            if clave in nombres_estandar
        }

        matriz = df_original.rename(columns=renombrar).copy()

        columnas_matriz = [
            "Observador",
            "Tipo Observación",
            *meses,
            "Real Año",
            "Teórica Año",
            "Avance",
        ]

        for columna in columnas_matriz:
            if columna not in matriz.columns:
                matriz[columna] = ""

        matriz = matriz[columnas_matriz].copy()

        for columna in meses + ["Real Año", "Teórica Año"]:
            matriz[columna] = matriz[columna].apply(limpiar_numero)

        # Excel puede entregar Avance como número decimal (0,56), número entero
        # (56) o texto con porcentaje ("56%"). Se fuerza tipo texto/objeto para
        # permitir mostrar el símbolo % sin provocar un TypeError de pandas.
        matriz["Avance"] = (
            matriz["Avance"]
            .fillna("")
            .astype("object")
        )

        matriz["Observador"] = (
            matriz["Observador"]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        matriz["Tipo Observación"] = (
            matriz["Tipo Observación"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        tipo_normalizado = matriz["Tipo Observación"].apply(normalizar_texto)
        observador_normalizado = matriz["Observador"].apply(normalizar_texto)

        es_resumen = (
            tipo_normalizado.str.contains("teorica|cumplimiento", na=False)
            | observador_normalizado.str.contains("teorica|cumplimiento", na=False)
        )

        registros = matriz.loc[~es_resumen].copy()
        registros = registros[
            registros["Observador"].ne("")
            & registros["Tipo Observación"].ne("")
        ].copy()

        # Completa automáticamente Real Año cuando la celda viene vacía o en cero.
        suma_meses = registros[meses].sum(axis=1)
        mascara_real_vacio = registros["Real Año"].fillna(0).eq(0)
        registros.loc[mascara_real_vacio, "Real Año"] = suma_meses.loc[
            mascara_real_vacio
        ]

        # Completa el avance anual desde Real/Teórica cuando corresponde.
        registros["Avance calculado"] = registros.apply(
            lambda fila: (
                fila["Real Año"] / fila["Teórica Año"] * 100
                if fila["Teórica Año"] > 0
                else 0
            ),
            axis=1,
        )

        def avance_visible(fila):
            avance_original = str(fila.get("Avance", "")).strip()

            if avance_original and avance_original.lower() not in {
                "nan", "none", "nat"
            }:
                if "%" in avance_original:
                    return avance_original

                valor = limpiar_numero(avance_original)

                if valor <= 1:
                    valor *= 100

                return porcentaje(valor)

            return porcentaje(fila["Avance calculado"])

        registros["Avance"] = registros.apply(avance_visible, axis=1)

        # Actualiza en la matriz solo el acumulado numérico. El avance formateado
        # se conserva en `registros`, evitando mezclar porcentajes de texto con
        # columnas numéricas provenientes de Excel.
        for indice in registros.index:
            matriz.loc[indice, "Real Año"] = registros.loc[indice, "Real Año"]

        total_real = registros["Real Año"].sum()
        total_teorico = registros["Teórica Año"].sum()
        avance_total = (
            total_real / total_teorico * 100
            if total_teorico > 0
            else 0
        )

        # -------------------------------------------------------------
        # CUMPLIMIENTO REAL A LA FECHA
        # -------------------------------------------------------------
        # Solo considera los meses transcurridos hasta el mes actual.
        # No proyecta ni incorpora metas de los meses siguientes.
        mes_actual_numero = int(HOY.month)
        meses_a_la_fecha = meses[:mes_actual_numero]

        real_a_la_fecha = registros[meses_a_la_fecha].sum(axis=1).sum()

        # Las filas TEORICAS de la planilla contienen las metas mensuales
        # consolidadas de Seguridad y BAPP.
        filas_teoricas = matriz.loc[
            tipo_normalizado.str.contains("teorica", na=False)
            | observador_normalizado.str.contains("teorica", na=False)
        ].copy()

        teorico_a_la_fecha = 0.0

        if not filas_teoricas.empty:
            teorico_a_la_fecha = filas_teoricas[meses_a_la_fecha].sum(axis=1).sum()

        # Respaldo: si la hoja no contiene filas TEORICAS, distribuye la
        # meta anual proporcionalmente hasta el mes actual.
        if teorico_a_la_fecha <= 0 and total_teorico > 0:
            teorico_a_la_fecha = total_teorico * (mes_actual_numero / 12)

        porcentaje_a_la_fecha = (
            real_a_la_fecha / teorico_a_la_fecha * 100
            if teorico_a_la_fecha > 0
            else 0
        )

        seguridad = registros[
            registros["Tipo Observación"]
            .apply(normalizar_texto)
            .eq("seguridad")
        ]
        bapp = registros[
            registros["Tipo Observación"]
            .apply(normalizar_texto)
            .eq("bapp")
        ]

        total_seguridad = seguridad["Real Año"].sum()
        total_bapp = bapp["Real Año"].sum()

        c1, c2, c3, c4, c5 = st.columns(5)

        with c1:
            kpi_card(
                "👷",
                "Observadores",
                numero(registros["Observador"].nunique()),
                "Personas con meta asignada",
            )

        with c2:
            kpi_card(
                "🟢",
                "Observaciones SSO",
                numero(total_seguridad),
                "Acumulado real del año",
            )

        with c3:
            kpi_card(
                "👀",
                "Observaciones BAPP",
                numero(total_bapp),
                "Acumulado real del año",
            )

        with c4:
            kpi_card(
                "📈",
                "Avance anual",
                porcentaje(avance_total),
                f"{numero(total_real)} de {numero(total_teorico)}",
            )

        with c5:
            kpi_card(
                "📅",
                "% a la fecha",
                porcentaje(porcentaje_a_la_fecha),
                (
                    f"{numero(real_a_la_fecha)} de "
                    f"{numero(teorico_a_la_fecha)} · "
                    f"hasta {MESES.get(mes_actual_numero, '')}"
                ),
            )

        # Gráfico 1: avance por observador.
        grafico_observador = registros[
            ["Observador", "Tipo Observación", "Real Año", "Teórica Año"]
        ].copy()

        grafico_observador = grafico_observador.melt(
            id_vars=["Observador", "Tipo Observación"],
            value_vars=["Real Año", "Teórica Año"],
            var_name="Indicador",
            value_name="Cantidad",
        )

        # Gráfico 2: evolución mensual por tipo de observación.
        # Solo se incluyen en el gráfico los meses con información cargada.
        meses_grafico = [
            mes
            for mes in meses
            if registros[mes].sum() > 0
        ]

        mensual = registros.melt(
            id_vars=["Observador", "Tipo Observación"],
            value_vars=meses_grafico,
            var_name="Mes",
            value_name="Cantidad",
        )

        mensual["Mes"] = pd.Categorical(
            mensual["Mes"],
            categories=meses_grafico,
            ordered=True,
        )

        mensual = (
            mensual.groupby(
                ["Mes", "Tipo Observación"],
                observed=False,
                as_index=False,
            )["Cantidad"]
            .sum()
        )

        col_a, col_b = st.columns(2)

        with col_a:
            card_inicio()
            fig_observador = px.bar(
                grafico_observador,
                x="Observador",
                y="Cantidad",
                color="Indicador",
                barmode="group",
                title="Resultado real versus meta anual",
                text_auto=".0f",
            )
            fig_observador = aplicar_layout_fig(
                fig_observador,
                height=410,
            )
            fig_observador.update_layout(
                legend=dict(
                    orientation="h",
                    yanchor="top",
                    y=-0.34,
                    xanchor="center",
                    x=0.5,
                    title_text="",
                ),
                xaxis_title=None,
                margin=dict(l=18, r=18, t=58, b=110),
            )
            st.plotly_chart(
                fig_observador,
                use_container_width=True,
            )
            card_fin()

        with col_b:
            card_inicio()
            if mensual.empty:
                st.info("No existen datos mensuales cargados para graficar.")
            else:
                fig_mensual = px.line(
                    mensual,
                    x="Mes",
                    y="Cantidad",
                    color="Tipo Observación",
                    markers=True,
                    title="Evolución mensual de observaciones",
                )
                fig_mensual = aplicar_layout_fig(
                    fig_mensual,
                    height=410,
                )
                fig_mensual.update_layout(
                    legend=dict(
                        orientation="h",
                        yanchor="top",
                        y=-0.34,
                        xanchor="center",
                        x=0.5,
                        title_text="",
                    ),
                    margin=dict(l=18, r=18, t=58, b=110),
                )
                st.plotly_chart(
                    fig_mensual,
                    use_container_width=True,
                )
            card_fin()

        panel_titulo("Cumplimiento anual por observador")

        tabla_registros = registros[
            [
                "Observador",
                "Tipo Observación",
                *meses,
                "Real Año",
                "Teórica Año",
                "Avance",
            ]
        ].copy()

        abreviaturas_meses = {
            "Enero": "ENE.",
            "Febrero": "FEB.",
            "Marzo": "MAR.",
            "Abril": "ABR.",
            "Mayo": "MAY.",
            "Junio": "JUN.",
            "Julio": "JUL.",
            "Agosto": "AGO.",
            "Septiembre": "SEP.",
            "Octubre": "OCT.",
            "Noviembre": "NOV.",
            "Diciembre": "DIC.",
        }

        # Se muestran únicamente los meses que tienen datos reales.
        meses_con_datos = [
            mes
            for mes in meses
            if registros[mes].sum() > 0
        ]

        tabla_registros = tabla_registros.rename(
            columns=abreviaturas_meses
        )

        # Porcentaje individual a la fecha: solo meses transcurridos.
        meses_abreviados_a_la_fecha = [
            abreviaturas_meses[mes]
            for mes in meses_a_la_fecha
            if mes in meses_con_datos
        ]

        if meses_abreviados_a_la_fecha:
            tabla_registros["Real a la fecha"] = tabla_registros[
                meses_abreviados_a_la_fecha
            ].sum(axis=1)
        else:
            tabla_registros["Real a la fecha"] = 0

        tabla_registros["Meta a la fecha"] = tabla_registros.apply(
            lambda fila: (
                fila["Teórica Año"] * (mes_actual_numero / 12)
                if fila["Teórica Año"] > 0
                else 0
            ),
            axis=1,
        )

        tabla_registros["% a la fecha"] = tabla_registros.apply(
            lambda fila: porcentaje(
                fila["Real a la fecha"] / fila["Meta a la fecha"] * 100
                if fila["Meta a la fecha"] > 0
                else 0
            ),
            axis=1,
        )

        tabla_limpia(
            tabla_registros,
            [
                "Observador",
                "Tipo Observación",
                *[
                    abreviaturas_meses[mes]
                    for mes in meses_con_datos
                ],
                "Real Año",
                "Teórica Año",
                "Avance",
                "Real a la fecha",
                "Meta a la fecha",
                "% a la fecha",
            ],
            height=430,
            centrar_todo=True,
            modo_ultracompacto=True,
        )

        return

    # -----------------------------------------------------------------
    # FORMATO HISTÓRICO ORIGINAL: una fila por observación.
    # -----------------------------------------------------------------
    df = aplicar_filtros(df_original, *filtros)

    seguras = int(
        df["Tipo_Observacion"]
        .astype(str)
        .str.contains(
            "segura|positivo",
            case=False,
            regex=True,
            na=False,
        )
        .sum()
    ) if not df.empty else 0

    riesgos = int(
        df["Tipo_Observacion"]
        .astype(str)
        .str.contains(
            "riesgo|subestandar|subestándar",
            case=False,
            regex=True,
            na=False,
        )
        .sum()
    ) if not df.empty else 0

    pendientes = int(
        df["Estado"]
        .astype(str)
        .str.contains(
            "Pendiente|En proceso|Vencida",
            case=False,
            regex=True,
            na=False,
        )
        .sum()
    ) if not df.empty else 0

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi_card(
            "👷",
            "OPS totales",
            numero(len(df)),
            "Observaciones preventivas",
        )

    with c2:
        kpi_card(
            "🟢",
            "Conductas seguras",
            numero(seguras),
            "Refuerzo positivo",
        )

    with c3:
        kpi_card(
            "🟠",
            "Conductas de riesgo",
            numero(riesgos),
            "Requieren control",
        )

    with c4:
        kpi_card(
            "⚠️",
            "OPS pendientes",
            numero(pendientes),
            "Con acción abierta",
        )

    col_a, col_b = st.columns(2)

    with col_a:
        card_inicio()
        grafico_barra(
            df,
            "Área",
            "OPS por área",
            orientacion="h",
        )
        card_fin()

    with col_b:
        card_inicio()
        grafico_donut(
            df,
            "Tipo_Observacion",
            "Tipo de observación",
        )
        card_fin()

    panel_titulo("Detalle OPS preventivas")

    tabla_limpia(
        df,
        [
            "Fecha",
            "Área",
            "Trabajador",
            "Supervisor",
            "Actividad",
            "Tipo_Observacion",
            "Conducta_Segura",
            "Conducta_Riesgo",
            "Medida_Correctiva",
            "Responsable",
            "Fecha_Compromiso",
            "Estado",
        ],
    )



def pagina_inspecciones(datos, filtros):
    df = aplicar_filtros(datos["Inspecciones"], *filtros)
    total = len(df)
    cumple = int(df["Resultado"].astype(str).str.contains("Cumple", case=False, na=False).sum()) if not df.empty else 0
    no_cumple = int(df["Resultado"].astype(str).str.contains("No cumple", case=False, na=False).sum()) if not df.empty else 0
    cumplimiento = (cumple / total * 100) if total else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("📋", "Inspecciones", numero(total), "Registros del periodo")
    with c2:
        kpi_card("✅", "Cumplimiento", porcentaje(cumplimiento), "Resultado cumple")
    with c3:
        kpi_card("❌", "No cumple", numero(no_cumple), "Hallazgos detectados")
    with c4:
        vencidas = int(df["Estado"].astype(str).str.contains("Vencida", case=False, na=False).sum()) if not df.empty else 0
        kpi_card("⚠️", "Vencidas", numero(vencidas), "Fuera de plazo")

    col_a, col_b = st.columns(2)
    with col_a:
        card_inicio()
        grafico_donut(df, "Resultado", "Resultado de inspecciones")
        card_fin()
    with col_b:
        card_inicio()
        grafico_barra(df, "Tipo_Inspeccion", "Inspecciones por tipo", orientacion="h")
        card_fin()

    panel_titulo("Detalle de inspecciones")
    tabla_limpia(df, ["Fecha", "Área", "Tipo_Inspeccion", "Resultado", "Hallazgos", "Responsable", "Fecha_Compromiso", "Estado", "Observacion"])


def pagina_plan_accion(datos, filtros):
    df = aplicar_filtros(datos["Plan_Accion"], *filtros)
    total = len(df)
    cerradas = int(df["Estado"].astype(str).str.contains("Cerrada", case=False, na=False).sum()) if not df.empty else 0
    vencidas = int(df["Estado"].astype(str).str.contains("Vencida", case=False, na=False).sum()) if not df.empty else 0
    pendientes = int(df["Estado"].astype(str).str.contains("Pendiente|En proceso|Vencida", case=False, regex=True, na=False).sum()) if not df.empty else 0
    cumplimiento = (cerradas / total * 100) if total else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("✅", "Acciones totales", numero(total), "Plan de acción SSO")
    with c2:
        kpi_card("🟢", "Cerradas", numero(cerradas), f"{porcentaje(cumplimiento)} de cumplimiento")
    with c3:
        kpi_card("🟠", "Pendientes", numero(pendientes), "Seguimiento requerido")
    with c4:
        kpi_card("🔴", "Vencidas", numero(vencidas), "Prioridad alta")

    col_a, col_b = st.columns(2)
    with col_a:
        card_inicio()
        grafico_donut(df, "Estado", "Estado del plan de acción")
        card_fin()
    with col_b:
        card_inicio()
        grafico_barra(df, "Origen", "Acciones por origen", orientacion="h")
        card_fin()

    panel_titulo("Detalle del plan de acción")
    tabla_limpia(df, ["Fecha", "Origen", "Área", "Hallazgo", "Accion_Correctiva", "Responsable", "Fecha_Compromiso", "Estado", "Evidencia", "Observacion"])


def pagina_capacitaciones(datos, filtros):
    mostrar_sello_saivam_pagina()

    # La vista utiliza la misma estructura y los mismos estados registrados en
    # la pestaña "Capacitaciones" de Google Sheets.
    df = aplicar_filtros(datos["Capacitaciones"], *filtros)

    if df is not None and not df.empty:
        df = df.sort_values("Fecha", ascending=True, na_position="last").copy()
        estados = df["Estado"].apply(estado_base)
        cerradas = int((estados == "Cerrada").sum())
        pendientes = int((estados == "Pendiente").sum())
        en_proceso = int((estados == "En proceso").sum())
    else:
        cerradas = 0
        pendientes = 0
        en_proceso = 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("🎓", "Capacitaciones", numero(len(df)), "Total programado")
    with c2:
        kpi_card("✅", "Cerradas", numero(cerradas), "Estado informado en Sheet")
    with c3:
        kpi_card("🟠", "Pendientes", numero(pendientes), "Estado informado en Sheet")
    with c4:
        kpi_card("🔵", "En proceso", numero(en_proceso), "Estado informado en Sheet")

    col_a, col_b = st.columns(2)
    with col_a:
        card_inicio()
        grafico_barra(df, "Tema", "Capacitaciones por tema", orientacion="h")
        card_fin()
    with col_b:
        card_inicio()
        grafico_donut(df, "Estado", "Estado de las capacitaciones")
        card_fin()

    panel_titulo("Detalle de capacitaciones")
    tabla_limpia(
        df,
        [
            "Fecha",
            "Tema",
            "Tipo",
            "Área",
            "Responsable",
            "Vencimiento",
            "Estado",
            "Observacion",
            "Evidencia",
        ],
    )


def pagina_protocolos_minsal(datos, filtros):
    mostrar_sello_saivam_pagina()

    df = aplicar_filtros(datos["Protocolos_MINSAL"], *filtros)
    total = len(df)
    expuestos = int(df["Expuestos"].apply(limpiar_numero).sum()) if not df.empty and "Expuestos" in df.columns else 0
    cerrados = int(df["Estado"].astype(str).str.contains("Cerrada", case=False, na=False).sum()) if not df.empty else 0
    abiertos = int(df["Estado"].astype(str).str.contains("Pendiente|En proceso|Vencida", case=False, regex=True, na=False).sum()) if not df.empty else 0
    protocolos = df["Protocolo"].nunique() if not df.empty and "Protocolo" in df.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("🦺", "Registros MINSAL", numero(total), "Evaluaciones y seguimientos")
    with c2:
        kpi_card("📑", "Protocolos controlados", numero(protocolos), "Tipos de protocolo")
    with c3:
        kpi_card("👥", "Trabajadores expuestos", numero(expuestos), "Expuestos registrados")
    with c4:
        kpi_card("⚠️", "Seguimientos abiertos", numero(abiertos), f"{cerrados} registros cerrados")

    col_a, col_b = st.columns(2)
    with col_a:
        card_inicio()
        grafico_barra(df, "Protocolo", "Registros por protocolo MINSAL", orientacion="h")
        card_fin()
    with col_b:
        card_inicio()
        grafico_donut(df, "Estado", "Estado de protocolos MINSAL")
        card_fin()

    panel_titulo("Detalle de Protocolos MINSAL")
    tabla_limpia(
        df,
        ["Fecha", "Protocolo", "Etapa", "Área", "Actividad", "Expuestos", "Responsable", "Resultado", "Fecha_Compromiso", "Estado", "Evidencia", "Observacion"],
    )


def pagina_programa_anual(datos, filtros):
    mostrar_sello_saivam_pagina()

    df = datos.get("Programa_Anual", pd.DataFrame()).copy()
    df = aplicar_filtros(df, *filtros)

    if df is None or df.empty:
        st.warning(
            "No se encontraron registros en la pestaña PRG_SSO_2026. "
            "Verifica que la hoja esté compartida como lector y que conserve "
            "los encabezados definidos."
        )
        return

    # Filtros propios del programa anual.
    meses_presentes = {
        str(valor).strip()
        for valor in df.get("Mes", pd.Series(dtype=str)).dropna()
        if str(valor).strip()
    }
    meses_ordenados = [mes for mes in MESES.values() if mes in meses_presentes]
    meses_extra = sorted(meses_presentes.difference(meses_ordenados))

    ejes = sorted(
        valor
        for valor in df.get("Eje_Trabajo", pd.Series(dtype=str)).dropna().astype(str).str.strip().unique()
        if valor
    )
    estados = sorted(
        valor
        for valor in df.get("Estado", pd.Series(dtype=str)).dropna().astype(str).str.strip().unique()
        if valor
    )

    f1, f2, f3 = st.columns(3)
    with f1:
        filtro_mes = st.selectbox(
            "Mes",
            ["Todos", *meses_ordenados, *meses_extra],
            key="prg_sso_filtro_mes",
        )
    with f2:
        filtro_eje = st.selectbox(
            "Eje de trabajo",
            ["Todos", *ejes],
            key="prg_sso_filtro_eje",
        )
    with f3:
        filtro_estado = st.selectbox(
            "Estado",
            ["Todos", *estados],
            key="prg_sso_filtro_estado",
        )

    if filtro_mes != "Todos":
        df = df[df["Mes"].astype(str).eq(filtro_mes)]
    if filtro_eje != "Todos":
        df = df[df["Eje_Trabajo"].astype(str).eq(filtro_eje)]

    # El avance a la fecha se calcula antes de aplicar el filtro Estado.
    # Así el indicador conserva una base completa y no se transforma en 100 %
    # cuando el usuario selecciona únicamente las actividades cerradas.
    df_avance = df.copy()

    if filtro_estado != "Todos":
        df = df[df["Estado"].astype(str).eq(filtro_estado)]

    total = len(df)
    estados_norm = df["Estado"].fillna("").apply(normalizar_texto)
    cerradas = int(estados_norm.eq("cerrada").sum())
    pendientes = int(estados_norm.eq("pendiente").sum())
    en_proceso = int(estados_norm.eq("en_proceso").sum())
    cumplimiento = (cerradas / total * 100) if total else 0.0
    tipos = df["Tipo_Actividad"].nunique() if "Tipo_Actividad" in df.columns else 0

    # Fecha dinámica de Chile continental. Se recalcula en cada ejecución de
    # la página, por lo que el corte avanza automáticamente con el calendario.
    try:
        hoy_programa = (
            pd.Timestamp.now(tz="America/Santiago")
            .tz_localize(None)
            .normalize()
        )
    except Exception:
        hoy_programa = pd.Timestamp.today().normalize()

    fechas_programadas = pd.to_datetime(
        df_avance.get("Fecha_Programada", pd.Series(index=df_avance.index, dtype="datetime64[ns]")),
        errors="coerce",
    )
    actividades_hasta_hoy = fechas_programadas.notna() & (
        fechas_programadas.dt.normalize() <= hoy_programa
    )
    programadas_a_fecha = int(actividades_hasta_hoy.sum())

    estados_a_fecha = (
        df_avance.loc[actividades_hasta_hoy, "Estado"]
        .fillna("")
        .apply(normalizar_texto)
    )
    cerradas_a_fecha = int(estados_a_fecha.eq("cerrada").sum())
    avance_a_fecha = (
        cerradas_a_fecha / programadas_a_fecha * 100
        if programadas_a_fecha
        else 0.0
    )
    fecha_corte = fecha_texto(hoy_programa)

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        kpi_card(
            "🗓️",
            "Actividades programadas",
            numero(total),
            f"{tipos} tipos de actividad",
        )
    with c2:
        kpi_card(
            "✅",
            "Actividades cerradas",
            numero(cerradas),
            f"{porcentaje(cumplimiento)} del total filtrado",
        )
    with c3:
        kpi_card(
            "🟠",
            "Actividades pendientes",
            numero(pendientes),
            "Estado informado en Sheet",
        )
    with c4:
        kpi_card(
            "🔵",
            "Actividades en proceso",
            numero(en_proceso),
            "Estado informado en Sheet",
        )
    with c5:
        kpi_card(
            "📈",
            "Avance a la fecha",
            porcentaje(avance_a_fecha),
            (
                f"{cerradas_a_fecha} de {programadas_a_fecha} cerradas "
                f"al {fecha_corte}"
            ),
        )

    col_a, col_b = st.columns(2)
    with col_a:
        card_inicio()
        grafico_donut(df, "Estado", "Estado del programa anual")
        card_fin()
    with col_b:
        card_inicio()
        grafico_barra(
            df,
            "Eje_Trabajo",
            "Actividades por eje de trabajo",
            orientacion="h",
        )
        card_fin()

    panel_titulo("Detalle del Programa Anual de Seguridad")
    tabla_limpia(
        df,
        [
            "Mes",
            "Eje_Trabajo",
            "Actividad",
            "Tipo_Actividad",
            "Fecha_Programada",
            "Fecha_Realizacion",
            "Responsable",
            "Estado",
            "Evidencia",
            "Observacion",
        ],
        height=520,
    )


def pagina_reconocimientos(datos, filtros):
    mostrar_sello_saivam_pagina()

    df = aplicar_filtros(datos["Reconocimientos"], *filtros)

    if df is None:
        df = pd.DataFrame(columns=SHEETS["Reconocimientos"]["columnas"])

    total = len(df)

    # Identificar reconocimientos institucionales realizados a SAIVAM.
    # Estos registros no se contabilizan como personas reconocidas.
    if not df.empty and "Trabajador" in df.columns:
        trabajador_normalizado = df["Trabajador"].fillna("").apply(normalizar_texto)

        es_empresa = trabajador_normalizado.str.contains(
            "saivam|empresa",
            case=False,
            regex=True,
            na=False,
        )

        # Respaldo para registros donde SAIVAM fue escrito en el motivo
        # o en la observación y no directamente en Trabajador.
        if "Motivo" in df.columns:
            motivo_normalizado = df["Motivo"].fillna("").apply(normalizar_texto)
            es_empresa = es_empresa | motivo_normalizado.str.contains(
                "empresa_saivam|reconocimiento_institucional",
                case=False,
                regex=True,
                na=False,
            )

        if "Observacion" in df.columns:
            observacion_normalizada = df["Observacion"].fillna("").apply(normalizar_texto)
            es_empresa = es_empresa | observacion_normalizada.str.contains(
                "empresa_saivam|reconocimiento_institucional",
                case=False,
                regex=True,
                na=False,
            )

        personas_df = df.loc[~es_empresa].copy()
        trabajadores = (
            personas_df["Trabajador"]
            .replace("", pd.NA)
            .dropna()
            .nunique()
        )

        # Contar cada registro corporativo, aunque corresponda a la misma empresa.
        # En la base actual existen dos reconocimientos corporativos a SAIVAM.
        reconocimientos_corporativos = int(es_empresa.sum())
    else:
        trabajadores = 0
        reconocimientos_corporativos = 0

    entregados = (
        int(
            df["Estado"]
            .astype(str)
            .str.contains(
                "Cerrada|Entregada",
                case=False,
                regex=True,
                na=False,
            )
            .sum()
        )
        if not df.empty and "Estado" in df.columns
        else 0
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("🏆", "Reconocimientos", numero(total), "Registros del período")
    with c2:
        kpi_card("👷", "Personas reconocidas", numero(trabajadores), "Trabajadores destacados")
    with c3:
        kpi_card("✅", "Reconocimientos entregados", numero(entregados), "Registros cerrados")
    with c4:
        kpi_card("🏢", "Reconocimientos corporativos", numero(reconocimientos_corporativos), "Reconocimientos a SAIVAM")

    mostrar_fotos_reconocimientos()

    col_a, col_b = st.columns(2)
    with col_a:
        card_inicio()
        grafico_barra(df, "Periodo", "Reconocimientos por período", orientacion="h")
        card_fin()
    with col_b:
        card_inicio()
        grafico_barra(df, "Cargo", "Reconocimientos por cargo", orientacion="h")
        card_fin()

    panel_titulo("Detalle de Reconocimientos")
    tabla_limpia(
        df,
        ["Fecha", "Trabajador", "Cargo", "Motivo", "Periodo", "Estado", "Evidencia", "Observacion"],
    )


def pagina_comite_paritario(datos, filtros):
    mostrar_sello_saivam_pagina()

    df = aplicar_filtros(datos["Comite_Paritario"], *filtros)
    total = len(df)
    cerradas = int(df["Estado"].astype(str).str.contains("Cerrada", case=False, na=False).sum()) if not df.empty else 0
    pendientes = int(df["Estado"].astype(str).str.contains("Pendiente|En proceso|Vencida", case=False, regex=True, na=False).sum()) if not df.empty else 0
    vencidas = int(df["Estado"].astype(str).str.contains("Vencida", case=False, na=False).sum()) if not df.empty else 0
    cumplimiento = (cerradas / total * 100) if total else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("👥", "Reuniones y acuerdos", numero(total), "Registros del comité")
    with c2:
        kpi_card("✅", "Acuerdos cerrados", numero(cerradas), f"{porcentaje(cumplimiento)} de cumplimiento")
    with c3:
        kpi_card("⚠️", "Acuerdos pendientes", numero(pendientes), "Seguimiento requerido")
    with c4:
        kpi_card("🚨", "Acuerdos vencidos", numero(vencidas), "Prioridad de cierre")

    col_a, col_b = st.columns(2)
    with col_a:
        card_inicio()
        grafico_donut(df, "Estado", "Estado de acuerdos del comité")
        card_fin()
    with col_b:
        card_inicio()
        grafico_barra(df, "Tipo_Reunion", "Reuniones por tipo", orientacion="h")
        card_fin()

    panel_titulo("Detalle del Comité Paritario")
    tabla_limpia(
        df,
        ["Fecha", "Tipo_Reunion", "Área", "Tema", "Acuerdo", "Responsable", "Fecha_Compromiso", "Estado", "Evidencia", "Observacion"],
    )

def pagina_trabajos_criticos(datos, filtros):
    df = aplicar_filtros(datos["Trabajos_Criticos"], *filtros)
    con_permiso = int(df["Permiso"].astype(str).str.contains("Si|Sí|Con permiso", case=False, regex=True, na=False).sum()) if not df.empty else 0
    abiertos = int(df["Estado"].astype(str).str.contains("Pendiente|En proceso|Vencida", case=False, regex=True, na=False).sum()) if not df.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("🔒", "Trabajos críticos", numero(len(df)), "Registros del periodo")
    with c2:
        kpi_card("📝", "Con permiso", numero(con_permiso), "Permiso informado")
    with c3:
        kpi_card("⚠️", "Abiertos", numero(abiertos), "En seguimiento")
    with c4:
        tipos = df["Tipo_Trabajo"].nunique() if not df.empty and "Tipo_Trabajo" in df.columns else 0
        kpi_card("📌", "Tipos de trabajo", numero(tipos), "Categorías críticas")

    col_a, col_b = st.columns(2)
    with col_a:
        card_inicio()
        grafico_barra(df, "Tipo_Trabajo", "Trabajos críticos por tipo", orientacion="h")
        card_fin()
    with col_b:
        card_inicio()
        grafico_donut(df, "Estado", "Estado de trabajos críticos")
        card_fin()

    panel_titulo("Detalle de trabajos críticos")
    tabla_limpia(df, ["Fecha", "Área", "Tipo_Trabajo", "Actividad", "Responsable", "Permiso", "Estado", "Observacion"])


def pagina_documentos(datos, filtros):
    df = datos["Documentos"].copy()
    if filtros[1] != "Todos" and "Año" in df.columns:
        df = df[df["Año"] == filtros[1]]
    if filtros[2] != "Todos" and "Mes" in df.columns:
        df = df[df["Mes"] == filtros[2]]

    vencidos = int(df["Estado"].astype(str).str.contains("Vencida", case=False, na=False).sum()) if not df.empty else 0
    vigentes = int(df["Estado"].astype(str).str.contains("Cerrada|Vigente", case=False, regex=True, na=False).sum()) if not df.empty else 0
    tipos = df["Tipo_Documento"].nunique() if not df.empty and "Tipo_Documento" in df.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("📁", "Documentos", numero(len(df)), "Registros SGS")
    with c2:
        kpi_card("✅", "Vigentes", numero(vigentes), "Documentos al día")
    with c3:
        kpi_card("⚠️", "Vencidos", numero(vencidos), "Requieren actualización")
    with c4:
        kpi_card("🗂️", "Tipos", numero(tipos), "Procedimientos, matrices, registros")

    col_a, col_b = st.columns(2)
    with col_a:
        card_inicio()
        grafico_barra(df, "Tipo_Documento", "Documentos por tipo", orientacion="h")
        card_fin()
    with col_b:
        card_inicio()
        grafico_donut(df, "Estado", "Estado documental")
        card_fin()

    panel_titulo("Detalle documentos SGS")
    tabla_limpia(df, ["Tipo_Documento", "Nombre_Documento", "Version", "Fecha", "Vencimiento", "Estado", "Ruta_Link", "Observacion"])




def pagina_certificaciones(datos, filtros):
    mostrar_sello_saivam_pagina()

    df = aplicar_filtros(datos["Certificaciones"], *filtros)

    total = len(df)

    if not df.empty and "Estado" in df.columns:
        estados = df["Estado"].fillna("").apply(estado_base)
        vigentes = int((estados == "Vigente").sum())
        por_vencer = int((estados == "Por vencer").sum())
        vencidas = int((estados == "Vencida").sum())
    else:
        vigentes = 0
        por_vencer = 0
        vencidas = 0

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        kpi_card(
            "📜",
            "Certificaciones",
            numero(total),
            "Registros totales",
        )

    with c2:
        kpi_card(
            "✅",
            "Vigentes",
            numero(vigentes),
            "Más de 30 días de vigencia",
        )

    with c3:
        kpi_card(
            "⚠️",
            "Por vencer",
            numero(por_vencer),
            "Vencen dentro de 30 días",
        )

    with c4:
        kpi_card(
            "🚨",
            "Vencidas",
            numero(vencidas),
            "Requieren renovación",
        )

    mostrar_equipos_certificados(df)

    col_a, col_b = st.columns(2)

    with col_a:
        card_inicio()
        grafico_barra(
            df,
            "Categoria",
            "Certificaciones por categoría",
            orientacion="h",
        )
        card_fin()

    with col_b:
        card_inicio()
        grafico_donut(
            df,
            "Estado",
            "Estado de vigencia",
        )
        card_fin()

    panel_titulo("Detalle de Certificaciones")

    columnas = [
        "Fecha",
        "Categoria",
        "Subcategoria",
        "Nombre_Certificacion",
        "Entidad_Emisora",
        "Vencimiento",
        "Estado",
        "Dias_Para_Vencer",
        "Ruta_Link",
    ]

    if df is None or df.empty:
        st.info("Sin certificaciones para mostrar.")
        return

    mostrar = df[columnas].copy()
    mostrar["Fecha"] = mostrar["Fecha"].apply(fecha_texto)
    mostrar["Vencimiento"] = mostrar["Vencimiento"].apply(fecha_texto)

    mostrar = mostrar.sort_values(
        by="Dias_Para_Vencer",
        ascending=True,
        na_position="last",
    )

    mostrar = mostrar.rename(
        columns={
            "Categoria": "Categoría",
            "Subcategoria": "Subcategoría",
            "Nombre_Certificacion": "Nombre certificación",
            "Entidad_Emisora": "Entidad emisora",
            "Dias_Para_Vencer": "Días para vencer",
            "Ruta_Link": "Ruta / link",
        }
    ).fillna("")

    # Tabla HTML compacta para que las nueve columnas entren en la página.
    st.markdown(
        """
<style>
.cert-table-wrap {
    width: 100%;
    overflow-x: hidden;
    border: 1px solid rgba(30, 180, 120, .42);
    border-radius: 13px;
    background: rgba(8, 13, 17, .94);
}
.cert-table {
    width: 100%;
    border-collapse: collapse;
    table-layout: fixed;
    font-size: clamp(8.5px, .58vw, 10.5px);
    line-height: 1.05;
}
.cert-table th,
.cert-table td {
    box-sizing: border-box;
    padding: 3px 4px !important;
    border-right: 1px solid rgba(110, 125, 140, .18);
    border-bottom: 1px solid rgba(110, 125, 140, .18);
    text-align: left;
    vertical-align: middle;
    white-space: normal;
    overflow-wrap: anywhere;
}
.cert-table th {
    height: 25px;
    background: #1b2029;
    color: #b9bec8;
    font-weight: 650;
}
.cert-table td {
    height: 23px;
    color: #f3f7f5;
}
.cert-table th:nth-child(1),
.cert-table td:nth-child(1) { width: 7%; }
.cert-table th:nth-child(2),
.cert-table td:nth-child(2) { width: 8%; }
.cert-table th:nth-child(3),
.cert-table td:nth-child(3) { width: 13%; }
.cert-table th:nth-child(4),
.cert-table td:nth-child(4) { width: 22%; }
.cert-table th:nth-child(5),
.cert-table td:nth-child(5) { width: 17%; }
.cert-table th:nth-child(6),
.cert-table td:nth-child(6) { width: 8%; }
.cert-table th:nth-child(7),
.cert-table td:nth-child(7) { width: 8%; }
.cert-table th:nth-child(8),
.cert-table td:nth-child(8) {
    width: 9%;
    text-align: center;
}
.cert-table th:nth-child(9),
.cert-table td:nth-child(9) {
    width: 8%;
    text-align: center;
}

.cert-link-button {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    padding: 4px 8px;
    border: 1px solid rgba(52, 211, 153, .60);
    border-radius: 8px;
    background: rgba(16, 185, 129, .14);
    color: #A7F3D0 !important;
    font-size: 9px;
    font-weight: 800;
    text-decoration: none !important;
    white-space: nowrap;
}

.cert-link-button:hover {
    background: rgba(16, 185, 129, .28);
    border-color: rgba(110, 231, 183, .90);
    color: #ECFDF5 !important;
}
</style>
        """,
        unsafe_allow_html=True,
    )

    encabezados_html = "".join(
        f"<th>{escape_html(columna)}</th>"
        for columna in mostrar.columns
    )

    filas_html = []

    for _, fila in mostrar.iterrows():
        celdas = []

        for columna in mostrar.columns:
            valor = fila[columna]
            texto = "" if pd.isna(valor) else str(valor).strip()

            if columna == "Ruta / link":
                if texto.lower().startswith(("http://", "https://")):
                    contenido = (
                        f'<a class="cert-link-button" '
                        f'href="{escape_html(texto)}" '
                        f'target="_blank" '
                        f'rel="noopener noreferrer">'
                        f'📄 Abrir'
                        f'</a>'
                    )
                elif texto:
                    contenido = escape_html(texto)
                else:
                    contenido = ""
            else:
                contenido = escape_html(texto)

            celdas.append(f"<td>{contenido}</td>")

        filas_html.append(
            f"<tr>{''.join(celdas)}</tr>"
        )

    st.markdown(
        (
            '<div class="cert-table-wrap">'
            '<table class="cert-table">'
            f'<thead><tr>{encabezados_html}</tr></thead>'
            f'<tbody>{"".join(filas_html)}</tbody>'
            '</table>'
            '</div>'
        ),
        unsafe_allow_html=True,
    )

    if por_vencer > 0 or vencidas > 0:
        panel_titulo("Alertas de vencimiento")

        alertas = df[
            df["Estado"].isin(["Por vencer", "Vencida"])
        ].copy()

        alertas = alertas.sort_values(
            by="Dias_Para_Vencer",
            ascending=True,
            na_position="last",
        )

        for _, fila in alertas.iterrows():
            estado = str(fila.get("Estado", ""))
            dias = fila.get("Dias_Para_Vencer", pd.NA)
            nombre = fila.get("Nombre_Certificacion", "Certificación")
            subcategoria = fila.get("Subcategoria", "")
            vencimiento = fecha_texto(fila.get("Vencimiento"))

            if estado == "Vencida":
                detalle = (
                    f"{subcategoria} · vencida el {vencimiento} "
                    f"({abs(int(dias))} días de atraso)"
                )
            else:
                detalle = (
                    f"{subcategoria} · vence el {vencimiento} "
                    f"({int(dias)} días restantes)"
                )

            st.markdown(
                f"""
<div class="alert-card">
    <div class="alert-title">⚠️ {escape_html(nombre)}</div>
    <div class="alert-sub">{escape_html(detalle)}</div>
</div>
                """,
                unsafe_allow_html=True,
            )


# =========================================================
# APP PRINCIPAL
# =========================================================

aplicar_estilo()
datos, archivo_excel, fuentes_datos = cargar_datos()

logo_sidebar = obtener_logo_sidebar_html()
st.sidebar.markdown(
    f"""
<div class="menu-brand">
    <div class="menu-logo-shell">{logo_sidebar}</div>
    <div>
        <div class="menu-title">Sistema de Gestión<br>SSO</div>
        <div class="menu-subtitle">SAIVAM · MULCHÉN</div>
    </div>
</div>
    """,
    unsafe_allow_html=True,
)

menu = st.sidebar.radio(
    "Menú",
    [
        "🛡️ KPI SSO",
        "🗓️ PRG SSO 2026",
        "⚠️ Reportabilidad",
        "🎯 Cumplimientos SSO",
        "🎓 Capacitaciones",
        "🏆 Reconocimientos",
        "👥 Comité Paritario",
        "🦺 Protocolos MINSAL",
        "📊 Certificaciones",
    ],
    label_visibility="collapsed",
)

# Filtros ocultos.
# Todas las páginas muestran la información completa disponible.
filtros = ("Todas las áreas", "Todos", "Todos")

st.sidebar.markdown(
    f"""
<div class="menu-footer-box">
    <div class="menu-info">
        <b>Contrato:</b> {escape_html(CONTRATO)}<br>
        <b>Empresa:</b> {escape_html(EMPRESA)}<br>
        <b>Versión:</b> {escape_html(VERSION)}
    </div>
</div>
    """,
    unsafe_allow_html=True,
)


logo_principal = obtener_logo_principal_html()
st.markdown(
    f"""
<div class="app-topbar">
    <div>
        <div class="title-main">Seguimiento y Control de Seguridad y Salud Ocupacional</div>
        <div class="subtitle-main">Sistema de Gestión SSO SAIVAM Mulchén · Programa anual, reportabilidad, cumplimientos preventivos y seguimiento de la gestión.</div>
    </div>
    <div class="main-logo-card">{logo_principal}</div>
</div>
    """,
    unsafe_allow_html=True,
)

if menu == "🛡️ KPI SSO":
    pagina_panel_general(datos, filtros)
elif menu == "🗓️ PRG SSO 2026":
    pagina_programa_anual(datos, filtros)
elif menu == "⚠️ Reportabilidad":
    pagina_reportabilidad(datos, filtros)
elif menu == "🎯 Cumplimientos SSO":
    pagina_cumplimientos_sso(datos, filtros)
elif menu == "🎓 Capacitaciones":
    pagina_capacitaciones(datos, filtros)
elif menu == "🏆 Reconocimientos":
    pagina_reconocimientos(datos, filtros)
elif menu == "👥 Comité Paritario":
    pagina_comite_paritario(datos, filtros)
elif menu == "🦺 Protocolos MINSAL":
    pagina_protocolos_minsal(datos, filtros)
elif menu == "📊 Certificaciones":
    pagina_certificaciones(datos, filtros)

st.markdown(
    f"""
<div class="footer-app footer-app-dos-lineas">
    <div class="footer-titulo">Panel desarrollado por</div>
    <div class="footer-detalle">
        {escape_html(AUTOR)} – Administrador de Contrato |
        María Araya – SSO |
        {escape_html(EMPRESA)} – {escape_html(CONTRATO)} –
        Versión {escape_html(VERSION)}
    </div>
</div>
    """,
    unsafe_allow_html=True,
)
