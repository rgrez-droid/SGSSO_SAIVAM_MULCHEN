import base64
import glob
import mimetypes
import os
import re
from datetime import datetime
from urllib.parse import quote

import pandas as pd
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

AUTOR = "Ricardo Grez"
EMPRESA = "SAIVAM"
CONTRATO = "CMPC Mulchén"
VERSION = "1.2"
REVISION_CODIGO = "14-07-2026-R20-VERSION-1.2"

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
GOOGLE_SHEET_ID = "1vMT3Xd68RR4KVIMRU5eaGhBtQxc4OcfI"

# Si una pestaña no puede leerse desde Google Sheets, el sistema conserva
# como respaldo la lectura desde el archivo Excel local.
USAR_GOOGLE_SHEETS = True

ARCHIVOS_EXCEL_POSIBLES = [
    # Base alineada con el menú actual de la aplicación.
    "Base_Datos_SSO_SAIVAM_Mulchen.xlsx",
    # Compatibilidad con versiones anteriores.
    "Base_Datos_SGS_SAIVAM_Mulchen.xlsx",
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
    "OPS": {
        "nombres": ["Observaciones_SSO_BAPP", "Observaciones SSO y BAPP", "OPS", "Observaciones", "Observaciones_Preventivas"],
        "secret": "ops_url",
        "columnas": [
            "Fecha", "Área", "Trabajador", "Supervisor", "Actividad",
            "Tipo_Observacion", "Conducta_Segura", "Conducta_Riesgo",
            "Medida_Correctiva", "Responsable", "Fecha_Compromiso",
            "Estado", "Observacion",
        ],
    },
    "Incidentes": {
        "nombres": ["Reportabilidad", "Incidentes", "Eventos", "Seguridad"],
        "secret": "incidentes_url",
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
    "Inspecciones": {
        "nombres": ["Inspecciones_Seguridad", "Inspecciones de Seguridad", "Inspecciones", "Checklist", "Checklists"],
        "secret": "inspecciones_url",
        "columnas": [
            "Fecha", "Área", "Tipo_Inspeccion", "Resultado", "Hallazgos",
            "Responsable", "Fecha_Compromiso", "Estado", "Observacion",
        ],
    },
    "Plan_Accion": {
        "nombres": ["Control_Operacional", "Control Operacional", "Plan_Accion", "Plan de Acción", "Acciones_Correctivas"],
        "secret": "plan_accion_url",
        "columnas": [
            "Fecha", "Origen", "Área", "Hallazgo", "Accion_Correctiva",
            "Responsable", "Fecha_Compromiso", "Estado", "Evidencia",
            "Observacion",
        ],
    },
    "Capacitaciones": {
        "nombres": ["Capacitaciones", "Charlas", "Charlas_Capacitaciones"],
        "secret": "capacitaciones_url",
        "columnas": [
            "Fecha", "Tema", "Tipo", "Área", "Relator", "Asistentes",
            "Vencimiento", "Estado", "Observacion",
        ],
    },
    "Programa_Anual": {
        "nombres": ["PRG_SSO_2026", "PRG SSO 2026", "Programa_Anual", "Programa Anual de Seguridad"],
        "secret": "programa_anual_url",
        "columnas": [
            "Fecha", "Mes", "Actividad", "Tipo_Actividad", "Área",
            "Responsable", "Meta", "Resultado", "Cumplimiento", "Estado",
            "Evidencia", "Observacion",
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
    columnas_fecha = ["Fecha", "Fecha_Compromiso", "Vencimiento", "Proxima_Reposicion"]
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
    """
    Construye una URL CSV para una pestaña específica del Google Sheet.

    Se utiliza el endpoint GViz porque permite leer cada pestaña por su nombre
    usando un único ID de documento.
    """
    nombre_codificado = quote(str(nombre_pestana), safe="")
    return (
        f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/gviz/tq"
        f"?tqx=out:csv&sheet={nombre_codificado}"
    )


def leer_hoja_desde_google(nombres_hoja):
    """
    Prueba los nombres alternativos configurados para cada módulo y devuelve
    la primera pestaña válida encontrada.
    """
    if not USAR_GOOGLE_SHEETS or not GOOGLE_SHEET_ID:
        return None

    for nombre_pestana in nombres_hoja:
        url = construir_url_google_sheet(nombre_pestana)

        try:
            df = pd.read_csv(url)

            # Google puede devolver una tabla vacía cuando el nombre no existe.
            if df is not None and len(df.columns) > 0:
                return df
        except Exception:
            continue

    return None



def crear_datos_ejemplo(nombre_hoja):
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
                "Tipo": "Charla operacional",
                "Área": "Mantención",
                "Relator": "Prevención",
                "Asistentes": 12,
                "Vencimiento": "02/07/2027",
                "Estado": "Cerrada",
                "Observacion": "Registro de ejemplo.",
            },
            {
                "Fecha": "05/07/2026",
                "Tema": "Uso correcto de EPP",
                "Tipo": "Charla 5 minutos",
                "Área": "Aserradero",
                "Relator": "Supervisor",
                "Asistentes": 8,
                "Vencimiento": "05/07/2027",
                "Estado": "Cerrada",
                "Observacion": "Registro de ejemplo.",
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
                "Fecha": "10/01/2026",
                "Mes": "Enero",
                "Actividad": "Difusión del programa anual de seguridad",
                "Tipo_Actividad": "Gestión preventiva",
                "Área": "Todas las áreas",
                "Responsable": "María Araya",
                "Meta": 1,
                "Resultado": 1,
                "Cumplimiento": 100,
                "Estado": "Cerrada",
                "Evidencia": "Carpeta/Programa_Anual/Enero",
                "Observacion": "Programa comunicado a supervisores y trabajadores.",
            },
            {
                "Fecha": "18/07/2026",
                "Mes": "Julio",
                "Actividad": "Revisión de procedimientos críticos",
                "Tipo_Actividad": "Auditoría",
                "Área": "Planta Térmica",
                "Responsable": "María Araya",
                "Meta": 1,
                "Resultado": 0,
                "Cumplimiento": 0,
                "Estado": "Pendiente",
                "Evidencia": "",
                "Observacion": "Actividad programada para julio.",
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
        df = leer_hoja_desde_google(config["nombres"])
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

        if nombre_hoja == "Capacitaciones":
            df = marcar_vencimientos(
                df,
                "Vencimiento",
            )

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
    # Solo se muestran estas cuatro fotografías, sin duplicados.
    # Se incluyen variantes de escritura para asegurar la carga de claudioa.
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
            "Verifica los archivos `claudioa`, `mariaa`, `ricardog` "
            ", `saivam500` y `saivam700`."
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
    background: rgba(3,12,9,.94) !important;
    border: 1px solid rgba(52,211,153,.30) !important;
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
        "🗂️": "celeste",
        "❌": "rojo",
    }
    tono = tonos.get(icono, "azul")
    st.markdown(
        f"""
<div class="kpi-card">
    <div class="kpi-icon {tono}">{icono}</div>
    <div class="kpi-title">{escape_html(titulo)}</div>
    <div class="kpi-value">{escape_html(valor)}</div>
    <div class="kpi-sub">{escape_html(subtitulo)}</div>
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
        height=height,
        margin=dict(l=12, r=12, t=48, b=16),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        font=dict(color="#D1FAE5", size=12),
        title_font=dict(size=20, color="#F4FFF9", family="Arial Black"),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.24,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(1,8,6,.68)",
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

def pagina_panel_general(datos, filtros):
    sello_agua = obtener_sello_agua_html()
    if sello_agua:
        st.markdown(sello_agua, unsafe_allow_html=True)

    ops = aplicar_filtros(datos["OPS"], *filtros)
    incidentes = aplicar_filtros(datos["Incidentes"], *filtros)
    inspecciones = aplicar_filtros(datos["Inspecciones"], *filtros)
    plan = aplicar_filtros(datos["Plan_Accion"], *filtros)
    capacitaciones = aplicar_filtros(datos["Capacitaciones"], *filtros)
    config = datos["Configuracion"]

    dias, fecha_inicio = dias_sin_accidentes(config, datos["Incidentes"])

    total_plan = len(plan)
    cerradas = int(plan["Estado"].astype(str).str.contains("Cerrada", case=False, na=False).sum()) if not plan.empty else 0
    vencidas = int(plan["Estado"].astype(str).str.contains("Vencida", case=False, na=False).sum()) if not plan.empty else 0
    pendientes = int(plan["Estado"].astype(str).str.contains("Pendiente|En proceso|Vencida", case=False, regex=True, na=False).sum()) if not plan.empty else 0
    cumplimiento = (cerradas / total_plan * 100) if total_plan else 0

    accidentes_periodo = 0
    if not incidentes.empty and "Tipo_Evento" in incidentes.columns:
        aux = incidentes.copy()
        aux["_tipo"] = aux["Tipo_Evento"].apply(normalizar_texto)
        accidentes_periodo = len(
            aux[
                aux["_tipo"].str.contains("accidente", na=False)
                & ~aux["_tipo"].str.contains("cuasi", na=False)
            ]
        )

    # Primera fila equivalente a las cinco tarjetas del panel de equipos.
    c1, c2, c3, c4, c5 = st.columns(5, gap="medium")
    with c1:
        kpi_card("🛡️", "Días sin accidentes", numero(dias), f"Desde {fecha_texto(fecha_inicio)}")
    with c2:
        kpi_card("👷", "Observaciones registradas", numero(len(ops)), "Observaciones del período")
    with c3:
        kpi_card("✅", "Cumplimiento del plan", porcentaje(cumplimiento), f"{cerradas} cerradas de {total_plan}")
    with c4:
        kpi_card("⚠️", "Acciones pendientes", numero(pendientes), f"{vencidas} acciones vencidas")
    with c5:
        kpi_card("🚨", "Accidentes registrados", numero(accidentes_periodo), "Período seleccionado")

    # Histórico preventivo + gráfico de distribución, siguiendo la estructura de la referencia.
    col_hist, col_dist = st.columns([1.58, 0.92], gap="large")

    with col_hist:
        panel_titulo("Histórico de Gestión Preventiva")
        historico = []

        if not ops.empty:
            ops_ultimas = ops.sort_values("Fecha", ascending=False).head(6)
            for _, fila in ops_ultimas.iterrows():
                detalle = (
                    fila.get("Conducta_Riesgo", "")
                    or fila.get("Conducta_Segura", "")
                    or fila.get("Actividad", "")
                )
                historico.append({
                    "Fecha": fila.get("Fecha", pd.NaT),
                    "Registro": "Observación",
                    "Área": fila.get("Área", ""),
                    "Detalle": detalle,
                    "Responsable": fila.get("Responsable", fila.get("Supervisor", "")),
                    "Estado": fila.get("Estado", ""),
                })

        if not plan.empty:
            plan_ultimas = plan.sort_values("Fecha", ascending=False).head(6)
            for _, fila in plan_ultimas.iterrows():
                historico.append({
                    "Fecha": fila.get("Fecha", pd.NaT),
                    "Registro": "Plan de acción",
                    "Área": fila.get("Área", ""),
                    "Detalle": fila.get("Accion_Correctiva", fila.get("Hallazgo", "")),
                    "Responsable": fila.get("Responsable", ""),
                    "Estado": fila.get("Estado", ""),
                })

        historico_df = pd.DataFrame(historico)
        if not historico_df.empty:
            historico_df = historico_df.sort_values("Fecha", ascending=False).head(9)
            historico_df["Fecha"] = historico_df["Fecha"].apply(fecha_texto)
            tabla_limpia(
                historico_df,
                ["Fecha", "Registro", "Área", "Detalle", "Responsable", "Estado"],
                height=335,
            )
        else:
            st.info("Sin registros preventivos para mostrar.")

    with col_dist:
        grafico_donut(plan, "Estado", "Distribución del Plan de Acción")

    # Bloques inferiores: alertas, compromisos y evolución.
    col_alertas, col_compromisos, col_evolucion = st.columns([0.95, 1.05, 1.25], gap="large")

    with col_alertas:
        panel_titulo("Alertas Preventivas")
        alertas = []
        if vencidas > 0:
            alertas.append(("Acciones vencidas", f"{vencidas} acciones correctivas están fuera de plazo."))
        if accidentes_periodo > 0:
            alertas.append(("Accidentes registrados", f"Se registran {accidentes_periodo} accidentes en el período."))
        if not inspecciones.empty and "Resultado" in inspecciones.columns:
            no_cumple = int(inspecciones["Resultado"].astype(str).str.contains("No cumple", case=False, na=False).sum())
            if no_cumple > 0:
                alertas.append(("Inspecciones con hallazgos", f"{no_cumple} inspecciones presentan resultado No cumple."))
        if not capacitaciones.empty and "Estado" in capacitaciones.columns:
            cap_vencidas = int(capacitaciones["Estado"].astype(str).str.contains("Vencida", case=False, na=False).sum())
            if cap_vencidas > 0:
                alertas.append(("Capacitaciones vencidas", f"{cap_vencidas} registros requieren actualización."))

        if not alertas:
            st.markdown(
                """
<div class="alert-card ok">
    <div class="alert-title">✅ Sin alertas críticas</div>
    <div class="alert-sub">Los registros filtrados no presentan alertas de prioridad alta.</div>
</div>
                """,
                unsafe_allow_html=True,
            )
        else:
            for titulo, detalle in alertas[:4]:
                st.markdown(
                    f"""
<div class="alert-card">
    <div class="alert-title">⚠️ {escape_html(titulo)}</div>
    <div class="alert-sub">{escape_html(detalle)}</div>
</div>
                    """,
                    unsafe_allow_html=True,
                )

    with col_compromisos:
        panel_titulo("Próximos Compromisos")
        proximos = plan.copy()
        if not proximos.empty:
            proximos["Fecha_Compromiso"] = proximos["Fecha_Compromiso"].apply(convertir_fecha)
            cerrada_mask = proximos["Estado"].astype(str).str.contains("Cerrada|Cumplida|Realizada", case=False, regex=True, na=False)
            proximos = proximos[~cerrada_mask & proximos["Fecha_Compromiso"].notna()]
            proximos = proximos.sort_values("Fecha_Compromiso", ascending=True).head(4)

        if proximos.empty:
            st.markdown(
                """
<div class="compromiso-card">
    <div class="compromiso-title">Sin compromisos abiertos</div>
    <div class="compromiso-sub">No existen acciones con fecha de compromiso pendiente.</div>
</div>
                """,
                unsafe_allow_html=True,
            )
        else:
            for _, fila in proximos.iterrows():
                fecha_comp = convertir_fecha(fila.get("Fecha_Compromiso"))
                dias_restantes = int((fecha_comp.normalize() - HOY).days) if pd.notna(fecha_comp) else 0
                if dias_restantes < 0:
                    texto_badge = f"Vencida {abs(dias_restantes)} días"
                    clase_badge = "vencida"
                elif dias_restantes == 0:
                    texto_badge = "Vence hoy"
                    clase_badge = ""
                else:
                    texto_badge = f"Faltan {dias_restantes} días"
                    clase_badge = ""

                titulo = fila.get("Accion_Correctiva", fila.get("Hallazgo", "Compromiso preventivo"))
                subtitulo = f"{fila.get('Área', 'Sin área')} · Responsable: {fila.get('Responsable', 'Sin asignar')} · {fecha_texto(fecha_comp)}"
                st.markdown(
                    f"""
<div class="compromiso-card">
    <div class="compromiso-head">
        <div class="compromiso-title">{escape_html(titulo)}</div>
        <span class="compromiso-badge {clase_badge}">{escape_html(texto_badge)}</span>
    </div>
    <div class="compromiso-sub">{escape_html(subtitulo)}</div>
</div>
                    """,
                    unsafe_allow_html=True,
                )

    with col_evolucion:
        # Serie consolidada para emular la sección "Evolución" del panel de referencia.
        series = []
        for nombre, dataframe in [
            ("Observaciones", ops),
            ("Eventos", incidentes),
            ("Inspecciones", inspecciones),
        ]:
            if dataframe is None or dataframe.empty or "Fecha" not in dataframe.columns:
                continue
            base = dataframe.copy()
            base["Fecha"] = base["Fecha"].apply(convertir_fecha)
            base = base[base["Fecha"].notna()]
            if base.empty:
                continue
            base["Periodo_Mes"] = base["Fecha"].dt.to_period("M").dt.to_timestamp()
            agrupado = base.groupby("Periodo_Mes", as_index=False).size().rename(columns={"size": "Cantidad"})
            agrupado["Indicador"] = nombre
            series.append(agrupado)

        if series:
            evolucion = pd.concat(series, ignore_index=True)
            fig = px.line(
                evolucion,
                x="Periodo_Mes",
                y="Cantidad",
                color="Indicador",
                markers=True,
                title="Evolución de Indicadores",
                color_discrete_sequence=PALETA_VERDE,
            )
            fig.update_xaxes(title="Mes")
            fig.update_yaxes(title="Registros", rangemode="tozero")
            st.plotly_chart(aplicar_layout_fig(fig, height=365), use_container_width=True)
        else:
            panel_titulo("Evolución de Indicadores")
            st.info("Sin fechas válidas para construir la evolución mensual.")


def pagina_reportabilidad(datos, filtros):
    """
    Panel de Reportabilidad basado en las columnas:

    Fecha, Área, Tipo_Evento, Descripcion, Accion_Inmediata,
    Responsable, Estado, Observacion y Ruta_Link.
    """
    df = aplicar_filtros(
        datos["Incidentes"],
        *filtros,
    )

    if df is None:
        df = pd.DataFrame()

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
    df = aplicar_filtros(datos["Capacitaciones"], *filtros)
    total_asistentes = int(df["Asistentes"].apply(limpiar_numero).sum()) if not df.empty and "Asistentes" in df.columns else 0
    vencidas = int(df["Estado"].astype(str).str.contains("Vencida", case=False, na=False).sum()) if not df.empty else 0
    realizadas = int(df["Estado"].astype(str).str.contains("Cerrada|Realizada", case=False, regex=True, na=False).sum()) if not df.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("🎓", "Actividades", numero(len(df)), "Charlas/capacitaciones")
    with c2:
        kpi_card("👥", "Asistencias", numero(total_asistentes), "Total asistentes")
    with c3:
        kpi_card("✅", "Realizadas", numero(realizadas), "Registros cerrados")
    with c4:
        kpi_card("⚠️", "Vencidas", numero(vencidas), "Revisar competencias")

    col_a, col_b = st.columns(2)
    with col_a:
        card_inicio()
        grafico_barra(df, "Tema", "Capacitaciones por tema", orientacion="h")
        card_fin()
    with col_b:
        card_inicio()
        grafico_donut(df, "Tipo", "Tipo de actividad")
        card_fin()

    panel_titulo("Detalle de charlas y capacitaciones")
    tabla_limpia(df, ["Fecha", "Tema", "Tipo", "Área", "Relator", "Asistentes", "Vencimiento", "Estado", "Observacion"])


def pagina_protocolos_minsal(datos, filtros):
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
    df = aplicar_filtros(datos["Programa_Anual"], *filtros)
    total = len(df)
    cerradas = int(df["Estado"].astype(str).str.contains("Cerrada", case=False, na=False).sum()) if not df.empty else 0
    pendientes = int(df["Estado"].astype(str).str.contains("Pendiente|En proceso|Vencida", case=False, regex=True, na=False).sum()) if not df.empty else 0
    cumplimiento = float(df["Cumplimiento"].apply(limpiar_numero).mean()) if not df.empty and "Cumplimiento" in df.columns else 0
    tipos = df["Tipo_Actividad"].nunique() if not df.empty and "Tipo_Actividad" in df.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        kpi_card("🗓️", "Actividades programadas", numero(total), "Programa anual de seguridad")
    with c2:
        kpi_card("✅", "Actividades ejecutadas", numero(cerradas), "Registros cerrados")
    with c3:
        kpi_card("⚠️", "Actividades pendientes", numero(pendientes), "Requieren seguimiento")
    with c4:
        kpi_card("📊", "Cumplimiento promedio", porcentaje(cumplimiento), f"{tipos} tipos de actividad")

    col_a, col_b = st.columns(2)
    with col_a:
        card_inicio()
        grafico_donut(df, "Estado", "Estado del programa anual")
        card_fin()
    with col_b:
        card_inicio()
        grafico_barra(df, "Tipo_Actividad", "Actividades por tipo", orientacion="h")
        card_fin()

    panel_titulo("Detalle del Programa Anual de Seguridad")
    tabla_limpia(
        df,
        ["Fecha", "Mes", "Actividad", "Tipo_Actividad", "Área", "Responsable", "Meta", "Resultado", "Cumplimiento", "Estado", "Evidencia", "Observacion"],
    )


def pagina_reconocimientos(datos, filtros):
    sello_reconocimientos = obtener_sello_reconocimientos_html()
    if sello_reconocimientos:
        st.markdown(sello_reconocimientos, unsafe_allow_html=True)

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
    sello_certificaciones = obtener_sello_certificaciones_html()

    if sello_certificaciones:
        st.markdown(
            sello_certificaciones,
            unsafe_allow_html=True,
        )

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
        "👀 Observaciones SSO y BAPP",
        "📋 Inspecciones de Seguridad",
        "✅ Control Operacional",
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
        <div class="subtitle-main">Sistema de Gestión SSO SAIVAM Mulchén · Programa anual, reportabilidad, observaciones preventivas y seguimiento de la gestión.</div>
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
elif menu == "👀 Observaciones SSO y BAPP":
    pagina_ops(datos, filtros)
elif menu == "📋 Inspecciones de Seguridad":
    pagina_inspecciones(datos, filtros)
elif menu == "✅ Control Operacional":
    pagina_plan_accion(datos, filtros)
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
