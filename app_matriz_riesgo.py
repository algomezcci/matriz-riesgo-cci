# ============================================================
# CCI Puesto de Bolsa, S.A. — Plataforma Matriz de Riesgo AML/CFT
# Ley 155-17 / Reglamento Art. 44
# ============================================================
# Cómo correr:
#   1. pip install streamlit pandas openpyxl reportlab
#      pip install google-cloud-bigquery google-auth db-dtypes pyarrow
#   2. gcloud auth application-default login   (solo primera vez)
#   3. streamlit run app_matriz_riesgo.py
#   4. Abre http://localhost:8501 en el navegador
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import json
import os
import hashlib
from datetime import date, datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

import warnings
warnings.filterwarnings("ignore")

st.set_page_config(
    page_title="CCI — Matriz de Riesgo AML/CFT",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ════════════════════════════════════════════════════════════
# COLORES CCI
# ════════════════════════════════════════════════════════════
CCI_VERDE       = "#9EC229"
CCI_VERDE_DARK  = "#7A9A1E"
CCI_GRIS        = "#414042"
CCI_GRIS_MID    = "#6B6869"
CCI_AMARILLO    = "#FBDE64"
CCI_GRIS_CLARO  = "#D1D3D4"

# ════════════════════════════════════════════════════════════
# SEGURIDAD — Contraseña para Configuración
# ════════════════════════════════════════════════════════════
# Hash SHA-256 de "cci2026" — cambiar este hash para cambiar la contraseña
CONFIG_PASSWORD_HASH = "4addc8f7004cda7833a41e1bef09018670e3338ab2b06fecc28a55491ce939fe"

def hash_password(p: str) -> str:
    return hashlib.sha256(p.encode()).hexdigest()

# ════════════════════════════════════════════════════════════
# ARCHIVO DE CONFIGURACIÓN PERSISTENTE
# ════════════════════════════════════════════════════════════
CONFIG_FILE = "config_matriz.json"

# ════════════════════════════════════════════════════════════
# VALORES POR DEFECTO — Matriz Mayo 2026
# ════════════════════════════════════════════════════════════
DEFAULT_SECTOR_SCORES = {
    "COMERCIO DE VEHICULOS": 8, "CONSTRUCCION": 7, "JUEGOS DEL AZAR": 10,
    "SERVICIOS DE TRANSPORTE Y ALMACENAMIENTO": 7,
    "SERVICIOS LEGALES Y JURIDICOS": 7, "PROMOTOR INMOBILIARIO": 7,
    "RELIGIOSOS": 9, "OTROS SECTORES INDUSTRIALES": 8,
    "OTROS SERVICIOS COMERCIALES": 8, "OTROS SERVICIOS PROFESIONALES": 8,
    "INTERMEDIACION FINANCIERA NO REGULADA": 8,
    "ACTIVIDADES INMOBILIARIAS Y DE ALQUILER": 8,
    "ADMINISTRACION PUBLICA Y DEFENSA": 8, "ZONA FRANCA": 8,
    "EXPLOTACION DE MINAS": 6, "GANADERIA SILVICULTURA Y PESCA": 6,
    "HOTELES BARES Y RESTAURANTES": 6, "SALUD": 5, "AGRICULTURA": 5,
    "INDUSTRIA DE ALIMENTOS": 5, "INTERMEDIACION FINANCIERA REGULADA": 5,
    "DISTRIBUCION DE ENERGIA Y AGUA": 5, "EDUCACION": 5,
    "BEBIDAS ALCOHOLICAS Y TABACO": 5, "COMUNICACIONES": 4,
    "COMERCIO DE ALIMENTOS Y BIENES ESCENCIALES": 3,
    "COMERCIO DE COMBUSTIBLES": 3,
    "COMERCIO DE ROPA CALZADO ACCESORIOS": 3,
    "CASAS DE CAMBIO AGENTE DE REMESAS FINANCIERAS NO REGULADAS Y PRESTAMISTA": 10,
    "00-NO APLICA": 3, "01-AGRICOLA": 5, "02-COMERCIAL": 3,
    "03-COMUNICACIONES": 4, "04-CONSTRUCCION": 7, "05-EDUCATIVO": 5,
    "06-RELIGIOSO": 9, "07-FINANCIERO": 5, "09-HOTELERO": 6,
    "10-INDUSTRIAL": 8, "11-AGROPECUARIO": 5, "13-TECNOLOGIA": 5,
    "14-TURISMO": 6, "15-PUBLICO": 8, "16-SALUD": 5, "17-LEGAL": 7,
    "18-SERVICIOS": 5, "22-TURISMO": 6, "28-OTROS": 8, "29-INVERSIONES": 6,
}

DEFAULT_COND_LABORAL_SCORES = {
    "EMPLEADO": 3, "EJERCICIO INDEPENDIENTE": 5,
    "EJERCICIO INDEPENDIENTE DE LA PROFESION": 5, "INDEPENDIENTE": 5,
    "DUENO DE EMPRESA": 7, "DUEÑO DE EMPRESA": 7, "PROPIETARIO": 7,
    "PENSIONADO": 8, "RETIRADO": 8, "PENSIONADO/RETIRADO": 8,
    "DESEMPLEADO": 10, "ESTUDIANTE": 6,
}

DEFAULT_PROFESION_SCORES = {
    "OTRO": 3, "MERCADEO": 1, "INGENIERA INDUSTRIAL": 3,
    "INGENIERIA INDUSTRIAL": 3, "CONTADURIA": 5, "MEDICO": 1,
    "INGENIERIA CIVIL": 5, "INGENIERIA EN COMPUTACION": 1,
    "ECONOMIA": 1, "INFORMATICA": 1, "DERECHO": 5,
    "RELACIONES INTERNACIONALES": 1,
    "DISENO Y COMUNICACION VISUAL": 1, "DISEÑO Y COMUNICACION VISUAL": 1,
    "DENTISTA": 1, "INGENIERIA ELECTRICA Y ELECTRONICA": 1,
    "COMUNICACION SOCIAL - PERIODISMO": 1, "RECURSOS HUMANOS": 1,
    "ADMINISTRACION / NEGOCIOS": 1, "PSICOLOGIA": 1, "ARQUITECTURA": 5,
    "ENFERMERIA": 1, "FARMACEUTICO": 1, "EDUCACION / PEDAGOGIA": 1,
    "TRABAJO SOCIAL": 1, "MATEMATICAS": 1, "QUIMICO": 1,
    "MILITAR": 5, "POLICIA": 5, "NOTARIO": 5,
    "INVERSIONISTA EN CRIPTOMONEDAS": 9, "INVERSIONISTA": 5,
    "AGENTE O INTERMEDIARIO FINANCIERO": 5, "DIPLOMATICO": 5,
    "CONTRATISTA DEL ESTADO": 5, "CORREDOR DE BIENES RAICES": 5,
}

DEFAULT_PATRIMONIO_SCORES = {
    "MENOS DE 500,000": 1, "DE 500,001 A 1,000,000": 2,
    "DE 1,000,001 A 10,000,000": 3,
    "MÁS DE 10,000,000": 6, "MAS DE 10,000,000": 6,
}

DEFAULT_INGRESOS_SCORES = {
    "NO TENGO": 2, "MENOS DE 500,000": 2,
    "DE 500,001 A 1,500,000": 4, "DE 501,000.00 A 1,500,000": 4,
    "DE 1,500,001 A 3,500,000": 6, "DE DOP 1,500,001 A DOP 3,500,000": 6,
    "DE 3,500,001 A 10,000,000": 8, "DE DOP 3,500,001 A DOP 10,000,000": 8,
    "MÁS DE 10,000,000": 10, "MAS DE 10,000,000": 10,
    "MAS DE DOP 10,000,000": 10,
}

DEFAULT_MONTO_MOVILIZAR_SCORES = {
    "MENOS DE DOP 500,000": 2,
    "DE DOP 500,001 A 1,500,000": 4,
    "DE DOP 1,500,001 A 3,500,000": 6,
    "DE DOP 3,500,001 A 10,000,000": 8,
    "MAS DE DOP 10,000,000": 10,
    "MÁS DE DOP 10,000,000": 10,
}

DEFAULT_PERFIL_SCORES = {
    "CONSERVADOR": 3, "C": 3, "MODERADO": 6, "M": 6, "AGRESIVO": 8, "A": 8,
}

DEFAULT_PROVINCIA_SCORES = {
    "05": 5, "31": 5, "29": 3, "22": 6, "14": 10, "15": 4,
    "25": 6, "09": 1, "13": 7, "18": 7, "04": 9, "01": 5, "33": 5,
}

DEFAULT_COUNTRY_SCORES = {
    "REPUBLICA DOMINICANA": 5,
    "ESTADOS UNIDOS": 1, "ESPANA": 1, "HUNGRIA": 3,
    "RUSIA": 10, "VENEZUELA": 10, "CUBA": 10, "HAITI": 10,
    "IRAN": 10, "IRAQ": 10, "COREA DEL NORTE": 10,
    "PAKISTAN": 10, "SIRIA": 10, "SUDAN": 10,
    "NICARAGUA": 8, "PANAMA": 8, "COLOMBIA": 6,
    "MEXICO": 5, "BRASIL": 3, "ARGENTINA": 3,
    "CHILE": 3, "CHINA": 3, "ALEMANIA": 1,
    "FRANCIA": 1, "CANADA": 1, "AUSTRALIA": 1,
}

DEFAULT_TIEMPO_RESIDENCIA_SCORES = {
    "MENOS DE 1 AÑO": 8, "DE 1 A 3 AÑOS": 5, "MAS DE 3 AÑOS": 3, "MÁS DE 3 AÑOS": 3,
}

DEFAULT_PRODUCTOS_SCORES = {
    "DERIVADOS (SPOT NO LIQUIDADOS, FORWARDS)": 10,
    "DERIVADOS LIQUIDADOS POR DIFERENCIAS": 9,
    "DERIVADOS": 10,
    "FINANCIAMIENTOS (PRESTAMOS DE MARGEN, BUY SELL BACK, REPORTO)": 8,
    "FINANCIAMIENTOS": 8,
    "TRC (ACCIONES, FONDOS DE INVERSION, FIDEICOMISOS DE OFERTA PUBLICA, FIDEICOMISOS NO OBJETO DE OFERTA PUBLICA)": 7,
    "TRC": 7, "ACCIONES": 7, "FONDOS DE INVERSION": 7,
    "FIDEICOMISOS DE OFERTA PUBLICA": 7,
    "OTROS ACTIVOS DE DEUDA (FACTURAS, ACUERDOS DE RECONOCIMIENTO DE DEUDA, VALORES TITULARIZADOS)": 6,
    "OTROS ACTIVOS DE DEUDA": 6,
    "CUOTAS DE PARTICIPACION FONDOS CERRADOS": 6,
    "VALORES DE PARTICIPACION DE FIDEICOMISOS DE OFERTA PUBLICA": 6,
    "MUTUO O PRESTAMO DE VALORES SIMPLE": 5,
    "ESTRUCTURADOS (SELL BUY BACK, SELL BUY BACK CON PRESTAMO DE VALORES, REPORTO)": 4,
    "ESTRUCTURADOS": 4,
    "VALORES DE DEUDA DE FIDEICOMISOS DE OFERTA PUBLICA": 4,
    "OPERACIONES ESTRUCTURADAS DE VENTA CON PACTO DE RECOMPRA": 4,
    "OPERACIONES ESTRUCTURADAS DE VENTA CON PACTO DE RECOMPRA + MUTUO": 4,
    "TVD (BONOS, PAPELES COMERCIALES, FIDEICOMISOS DE DEUDA, PRESTAMOS DE VALORES)": 3,
    "TVD": 3, "BONOS O INSTRUMENTOS DE DEUDA": 3, "BONOS": 3,
    "PAPELES COMERCIALES": 3,
    "RENTA FIJA": 2, "RENTA VARIABLE": 3,
    "FONDOS": 2, "NO TENGO": 2,
}

# Origen de fondos — valores por fuente (se calcula el promedio de las fuentes marcadas)
DEFAULT_ORIGEN_FONDOS_SCORES = {
    "SALARIO": 1,        # Fuente formal, trazable
    "INVERSIONES": 2,    # Trazable vía estados financieros
    "INMUEBLES": 3,      # Documentable con contratos
    "DONACIONES": 6,     # Difícil de verificar
    "OTROS": 8,          # Sin trazabilidad
}

# Destino / Objetivo de la inversión
DEFAULT_DESTINO_FONDOS_SCORES = {
    "CONSERVACION DE CAPITAL": 1,
    "APRECIACION DEL CAPITAL Y FLUJO PERIODICO": 3,
    "APRECIACION DEL CAPITAL": 3,
    "FLUJO PERIODICO": 4,
    "ESPECULACION": 8,
    "OTROS": 5,
}

DEFAULT_TIPO_CLIENTE_SCORES = {"PROFESIONAL": 0, "NO PROFESIONAL": 0}
DEFAULT_METAMAP_SCORES = {"approved": 0, "revisado": 2, "declined": 5}

# Pesos crudos Matriz Mayo 2026 (luego se normalizan a 100%)
# Al agregar Origen y Destino, todo se normaliza automáticamente
DEFAULT_WEIGHTS_RAW = {
    "pep":               0.0,
    "tipo_cliente":      0.03,
    "apertura":          0.03,
    "cond_laboral":      0.10,
    "sector":            0.10,
    "profesion":         0.10,
    "tiempo_residencia": 0.09,
    "scoring_credito":   0.03,
    "productos":         0.10,
    "fecha_nac":         0.03,
    "patrimonio":        0.045,
    "activos_liquidos":  0.045,
    "ingresos":          0.10,
    "monto_movilizar":   0.10,
    "perfil_inv":        0.05,
    "duracion":          0.05,
    "metamap":           0.05,
    "origen_fondos":     0.08,
    "destino_fondos":    0.05,
}

DEFAULT_VALUES = {
    "cond_laboral": 4, "sector": 4, "profesion": 2,
    "pais_origen": 3, "pais_residencia": 3, "provincia": 5,
    "tiempo_residencia": 3, "scoring_credito": 2,
    "productos": 3, "patrimonio": 2, "activos_liquidos": 2,
    "ingresos": 3, "monto_movilizar": 4, "perfil_inv": 3,
    "fecha_nac": 3, "duracion": 5, "tipo_cliente": 0,
    "origen_fondos": 5, "destino_fondos": 5,
}

DEFAULT_UMBRALES = {"bajo": 5.0, "alto": 6.0}

STATUS_MAP = {
    "VERIFIED": "approved", "APPROVED": "approved",
    "REVIEWNEEDED": "revisado", "REVIEWNEED": "revisado",
    "REJECTED": "declined", "DECLINED": "declined",
}

# ════════════════════════════════════════════════════════════
# CARGA DE CONFIGURACIÓN PERSISTENTE
# ════════════════════════════════════════════════════════════
def cargar_config():
    """Carga config desde archivo JSON. Si no existe, usa defaults."""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {
        "weights_raw": dict(DEFAULT_WEIGHTS_RAW),
        "umbrales": dict(DEFAULT_UMBRALES),
        "sector_scores": dict(DEFAULT_SECTOR_SCORES),
        "cond_laboral_scores": dict(DEFAULT_COND_LABORAL_SCORES),
        "profesion_scores": dict(DEFAULT_PROFESION_SCORES),
        "patrimonio_scores": dict(DEFAULT_PATRIMONIO_SCORES),
        "ingresos_scores": dict(DEFAULT_INGRESOS_SCORES),
        "monto_movilizar_scores": dict(DEFAULT_MONTO_MOVILIZAR_SCORES),
        "perfil_scores": dict(DEFAULT_PERFIL_SCORES),
        "productos_scores": dict(DEFAULT_PRODUCTOS_SCORES),
        "tiempo_residencia_scores": dict(DEFAULT_TIEMPO_RESIDENCIA_SCORES),
        "country_scores": dict(DEFAULT_COUNTRY_SCORES),
        "provincia_scores": dict(DEFAULT_PROVINCIA_SCORES),
        "origen_fondos_scores": dict(DEFAULT_ORIGEN_FONDOS_SCORES),
        "destino_fondos_scores": dict(DEFAULT_DESTINO_FONDOS_SCORES),
    }

def guardar_config(config):
    """Guarda config a archivo JSON."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        st.error(f"Error guardando configuración: {e}")
        return False

def normalizar_pesos(weights_raw):
    """Convierte pesos crudos a porcentajes que suman 1.0"""
    total = sum(v for v in weights_raw.values() if v > 0)
    if total == 0:
        return {k: 0 for k in weights_raw}
    return {k: v/total for k, v in weights_raw.items()}

# ════════════════════════════════════════════════════════════
# CSS / ESTILO CCI
# ════════════════════════════════════════════════════════════
st.markdown(f"""
<style>
    html, body, [class*="css"] {{ font-family: 'Calibri', 'Segoe UI', sans-serif; }}
    .cci-header {{
        background-color: {CCI_GRIS}; padding: 1rem 1.5rem;
        border-left: 6px solid {CCI_VERDE};
        margin-bottom: 1.5rem; border-radius: 4px;
    }}
    .cci-header h1 {{ color: white; margin: 0; font-size: 1.6rem; }}
    .cci-header p {{ color: {CCI_GRIS_CLARO}; margin: 0.3rem 0 0 0; font-size: 0.9rem; }}
    .stat-box {{
        background-color: white; padding: 1.2rem; border-radius: 6px;
        border-left: 4px solid {CCI_VERDE};
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    }}
    .stat-box h2 {{ margin: 0; font-size: 2.2rem; color: {CCI_GRIS}; }}
    .stat-box p {{
        margin: 0.2rem 0 0 0; color: {CCI_GRIS_MID};
        font-size: 0.85rem; text-transform: uppercase; letter-spacing: 1px;
    }}
    .stButton button {{
        background-color: {CCI_VERDE}; color: white;
        border: none; font-weight: bold;
    }}
    .stButton button:hover {{ background-color: {CCI_VERDE_DARK}; color: white; }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 12px; }}
    .stTabs [data-baseweb="tab"] {{
        background-color: #f5f5f5; color: {CCI_GRIS};
        font-weight: 600; padding: 0.5rem 1.5rem;
        border-radius: 4px 4px 0 0;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {CCI_VERDE} !important; color: white !important;
    }}
    section[data-testid="stSidebar"] {{ background-color: {CCI_GRIS}; }}
    section[data-testid="stSidebar"] * {{ color: white !important; }}
</style>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# UTILIDADES
# ════════════════════════════════════════════════════════════
def normalize(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).upper().strip()
    for a, b in [("Á","A"),("É","E"),("Í","I"),("Ó","O"),("Ú","U"),("Ü","U"),("Ñ","N")]:
        s = s.replace(a, b)
    return s

def is_empty(val):
    if val is None: return True
    if isinstance(val, float) and pd.isna(val): return True
    return str(val).strip().upper() in ("#N/A","N/A","NA","","NAN","NONE","NULL")

def fuzzy_get(val, table, default):
    v = normalize(val)
    if not v: return default
    if v in table: return table[v]
    for k, s in table.items():
        if normalize(k) in v: return s
    return default

def normalize_status(val):
    try:
        if val is None or isinstance(val, float): return "revisado"
        s = str(val).strip()
        if s.upper() in ("#N/A","N/A","","NAN","NONE","NULL"): return "revisado"
        return STATUS_MAP.get(s.upper().replace(" ","").replace("_",""), "revisado")
    except:
        return "revisado"

def calc_age(dob):
    try:
        if is_empty(dob): return None
        d = dob.date() if hasattr(dob, "date") else dob
        t = date.today()
        return t.year - d.year - ((t.month, t.day) < (d.month, d.day))
    except:
        return None

def score_edad(age):
    if age is None: return 3
    if age < 30: return 5
    elif age < 40: return 3
    elif age < 55: return 1
    elif age < 65: return 3
    else: return 2

def score_duracion(years):
    if years is None: return 5
    if years == 0: return 7
    elif years <= 5: return 6
    else: return 3

def score_pais(pais_res, provincia, country_scores, provincia_scores):
    pais = normalize(pais_res)
    if is_empty(pais_res): return DEFAULT_VALUES["pais_residencia"]
    if "DOMINICANA" in pais:
        prov = normalize(provincia)
        if not prov: return 5
        code = prov.split()[0]
        return provincia_scores.get(code, 5)
    return country_scores.get(pais, 3)

def score_credito(val):
    if is_empty(val): return DEFAULT_VALUES["scoring_credito"]
    try:
        n = float(str(val))
        if n >= 750: return 1
        elif n >= 500: return 2
        else: return 3
    except: pass
    return fuzzy_get(val, {
        "EXCELENTE": 1, "MUY BUENO": 1, "BUENO": 2,
        "REGULAR": 2, "BAJO": 3, "MUY BAJO": 3,
    }, DEFAULT_VALUES["scoring_credito"])

def score_origen_fondos(texto, origen_scores):
    """Toma promedio de fuentes marcadas. Texto viene tipo 'Salario, Inmuebles, Donaciones'."""
    if is_empty(texto): return DEFAULT_VALUES["origen_fondos"]
    texto_up = normalize(texto)
    valores = []
    for fuente, score in origen_scores.items():
        if normalize(fuente) in texto_up:
            valores.append(score)
    if not valores: return DEFAULT_VALUES["origen_fondos"]
    return sum(valores) / len(valores)

def score_destino_fondos(val, destino_scores):
    return fuzzy_get(val, destino_scores, DEFAULT_VALUES["destino_fondos"])

# ════════════════════════════════════════════════════════════
# CÁLCULO DE PUNTAJE — recibe la configuración completa
# ════════════════════════════════════════════════════════════
def calcular_scores(df, config):
    """Calcula scores usando los pesos y tablas de la configuración."""
    weights = normalizar_pesos(config["weights_raw"])
    umbral_bajo = config["umbrales"]["bajo"]
    umbral_alto = config["umbrales"]["alto"]

    if "fecha_nacimiento" in df.columns:
        df["fecha_nacimiento"] = pd.to_datetime(df["fecha_nacimiento"], errors="coerce", dayfirst=True)
        df["edad"] = df["fecha_nacimiento"].apply(calc_age)
    else:
        df["edad"] = None

    if "duracion_relacion" in df.columns:
        df["anos_relac"] = pd.to_numeric(df["duracion_relacion"], errors="coerce").apply(
            lambda x: int(x) if pd.notna(x) else None)
    else:
        df["anos_relac"] = None

    df["razon_inadmisible"] = df.apply(
        lambda r: f"Status: {r.get('metamap','')}"
        if normalize_status(r.get("metamap","")) == "declined" else "", axis=1)

    df["v_pep"] = df.get("pep", pd.Series([""]*len(df))).apply(
        lambda x: 10 if normalize(x) in ("SI","S","YES") else 0)
    df["v_apertura"] = 5
    df["v_cond"] = df.get("cond_laboral", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, config["cond_laboral_scores"], DEFAULT_VALUES["cond_laboral"]))
    df["v_sector"] = df.get("sector", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, config["sector_scores"], DEFAULT_VALUES["sector"]))
    df["v_profesion"] = df.apply(lambda r:
        1 if "PENSION" in normalize(r.get("cond_laboral","")) or "RETIR" in normalize(r.get("cond_laboral",""))
        else fuzzy_get(r.get("profesion",""), config["profesion_scores"], DEFAULT_VALUES["profesion"]), axis=1)
    df["v_pais"] = df.apply(lambda r: score_pais(
        r.get("pais_residencia",""), r.get("provincia",""),
        config["country_scores"], config["provincia_scores"]), axis=1)
    df["v_tiempo_res"] = df.get("tiempo_residencia", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, config["tiempo_residencia_scores"], DEFAULT_VALUES["tiempo_residencia"]))
    df["v_credito"] = df.get("scoring_credito", pd.Series([""]*len(df))).apply(score_credito)
    df["v_productos"] = df.get("productos", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, config["productos_scores"], DEFAULT_VALUES["productos"]))
    df["v_edad"] = df["edad"].apply(score_edad)
    df["v_patrimonio"] = df.get("patrimonio", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, config["patrimonio_scores"], DEFAULT_VALUES["patrimonio"]))
    df["v_activos"] = df.get("activos_liquidos", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, config["patrimonio_scores"], DEFAULT_VALUES["activos_liquidos"]))
    df["v_ingresos"] = df.get("ingresos", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, config["ingresos_scores"], DEFAULT_VALUES["ingresos"]))
    df["v_monto"] = df.get("monto_movilizar", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, config["monto_movilizar_scores"], DEFAULT_VALUES["monto_movilizar"]))
    df["v_perfil"] = df.get("perfil_inv", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, config["perfil_scores"], DEFAULT_VALUES["perfil_inv"]))
    df["v_duracion"] = df["anos_relac"].apply(score_duracion)
    df["v_metamap"] = df.get("metamap", pd.Series([""]*len(df))).apply(
        lambda x: DEFAULT_METAMAP_SCORES.get(normalize_status(x), 2))
    df["v_tipo"] = 0
    df["v_origen"] = df.get("origen_fondos", pd.Series([""]*len(df))).apply(
        lambda x: score_origen_fondos(x, config["origen_fondos_scores"]))
    df["v_destino"] = df.get("destino_fondos", pd.Series([""]*len(df))).apply(
        lambda x: score_destino_fondos(x, config["destino_fondos_scores"]))

    df["puntaje_final"] = (
        df["v_tipo"]       * weights.get("tipo_cliente", 0) +
        df["v_apertura"]   * weights.get("apertura", 0) +
        df["v_cond"]       * weights.get("cond_laboral", 0) +
        df["v_sector"]     * weights.get("sector", 0) +
        df["v_profesion"]  * weights.get("profesion", 0) +
        df["v_tiempo_res"] * weights.get("tiempo_residencia", 0) +
        df["v_credito"]    * weights.get("scoring_credito", 0) +
        df["v_productos"]  * weights.get("productos", 0) +
        df["v_edad"]       * weights.get("fecha_nac", 0) +
        df["v_patrimonio"] * weights.get("patrimonio", 0) +
        df["v_activos"]    * weights.get("activos_liquidos", 0) +
        df["v_ingresos"]   * weights.get("ingresos", 0) +
        df["v_monto"]      * weights.get("monto_movilizar", 0) +
        df["v_perfil"]     * weights.get("perfil_inv", 0) +
        df["v_duracion"]   * weights.get("duracion", 0) +
        df["v_metamap"]    * weights.get("metamap", 0) +
        df["v_origen"]     * weights.get("origen_fondos", 0) +
        df["v_destino"]    * weights.get("destino_fondos", 0)
    ).round(2)

    def classify(row):
        if row["razon_inadmisible"]: return "Inadmisible"
        if row["v_pep"] == 10: return "Alto"
        s = row["puntaje_final"]
        if s <= umbral_bajo: return "Bajo"
        elif s <= umbral_alto: return "Medio"
        else: return "Alto"

    df["nivel_riesgo"] = df.apply(classify, axis=1)

    def flag_incoherence(row):
        flags = []
        patr = normalize(row.get("patrimonio",""))
        ingr = normalize(row.get("ingresos",""))
        if any(x in patr for x in ["MAS DE 10","MÁS DE 10","1,000,001"]) and \
           any(x in ingr for x in ["MENOS DE 500","NO TENGO","501,000"]):
            flags.append("Patrimonio alto / Ingresos bajos")
        # Incongruencia origen vs destino
        if "DONACIONES" in normalize(row.get("origen_fondos","")) and \
           "ESPECULACION" in normalize(row.get("destino_fondos","")):
            flags.append("Origen donaciones + destino especulativo")
        return "; ".join(flags) if flags else ""
    df["alerta_incoherencia"] = df.apply(flag_incoherence, axis=1)

    CAMPOS = {
        "sector": "Sector", "profesion": "Profesión",
        "pais_residencia": "País Residencia", "provincia": "Provincia",
        "scoring_credito": "Score Crédito", "patrimonio": "Patrimonio",
        "activos_liquidos": "Activos Líq.", "ingresos": "Ingresos",
        "perfil_inv": "Perfil", "fecha_nacimiento": "Fecha Nac.",
        "monto_movilizar": "Monto Mov.", "origen_fondos": "Origen fondos",
        "destino_fondos": "Destino fondos",
    }
    df["campos_faltantes"] = df.apply(
        lambda r: " | ".join([n for c,n in CAMPOS.items() if is_empty(r.get(c))]) , axis=1)

    return df

# ════════════════════════════════════════════════════════════
# COLUMN MAP — BigQuery → app
# ════════════════════════════════════════════════════════════
COLUMN_MAP = {
    "Email": "email", "Codigo": "codigo", "Nombre_Completo": "nombre",
    "PEP": "pep", "tipo_contrato": "tipo_contrato", "Sucursal": "sucursal",
    "edad": "edad", "tipo_cliente": "tipo_cliente", "tipo_apertura": "apertura",
    "condicion_laboral": "cond_laboral", "sector_economico": "sector",
    "profesion": "profesion", "pais_origen": "pais_origen",
    "pais_residencia": "pais_residencia", "Provincia": "provincia",
    "tiempo_residencia": "tiempo_residencia", "score": "scoring_credito",
    "fecha_nacimiento": "fecha_nacimiento", "patrimonio_dop": "patrimonio",
    "activos_liquidos": "activos_liquidos", "ingreso_mensual_dop": "ingresos",
    "monto_movilizar": "monto_movilizar", "perfil_inversionista": "perfil_inv",
    "duracion_relacion": "duracion_relacion",
    "perfil_transaccional": "perfil_legitimacion", "productos": "productos",
    "origen_fondos": "origen_fondos", "destino_fondos": "destino_fondos",
    "status": "metamap",
    # Compatibilidad Excel viejo
    "ocupacion": "cond_laboral", "Actividad_Economica": "sector",
    "patrimonio_total": "patrimonio", "patrimonio_liquido": "activos_liquidos",
    "ingresos_anuales": "ingresos", "Nacionalidad": "nacionalidad",
}

# ════════════════════════════════════════════════════════════
# CONSULTA SQL BIGQUERY
# ════════════════════════════════════════════════════════════
SQL_BIGQUERY = """
SELECT 
  s.email AS Email,
  s.codigo_cliente AS Codigo,
  TRIM(CONCAT(COALESCE(s.nombre, ''), ' ', COALESCE(s.apellido, ''))) AS Nombre_Completo,
  s.pep AS PEP,
  s.tipo_contrato AS tipo_contrato,
  s.sucursal AS Sucursal,
  s.edad,
  CASE WHEN is2.is_unique_selection = FALSE THEN 'Profesional' ELSE 'No profesional' END AS tipo_cliente,
  CASE WHEN kyc.tipo_apertura = "ANALOGA" THEN "FISICA" ELSE kyc.tipo_apertura END AS tipo_apertura,
  kyc.condicion_laboral,
  kyc.sector_economico,
  kyc.profesion AS profesion,
  kyc.pais_nacimiento AS pais_origen,
  kyc.pais_residencia AS pais_residencia,
  kyc.provincia AS Provincia,
  kyc.tiempo_residencia AS tiempo_residencia,
  sc.score AS score,
  s.fecha_nacimiento AS fecha_nacimiento,
  kyc.patrimonio_dop,
  kyc.activos_liquidos,
  fon.ingreso_mensual_dop,
  kyc.monto_transaccion AS monto_movilizar,
  s.perfil_inversionista AS perfil_inversionista,
  DATE_DIFF(CURRENT_DATE('America/Santo_Domingo'), s.fecha_registro, YEAR) - 
    IF(FORMAT_DATE('%m%d', CURRENT_DATE('America/Santo_Domingo')) < FORMAT_DATE('%m%d', s.fecha_registro), 1, 0) 
    AS duracion_relacion,
  pe.perfil_transaccional,
  CASE
    WHEN fon.producto_derivado       = TRUE THEN 'DERIVADOS'
    WHEN fon.producto_financiamiento = TRUE THEN 'FINANCIAMIENTOS'
    WHEN fon.producto_trc            = TRUE THEN 'TRC'
    WHEN fon.producto_otros_activos  = TRUE THEN 'OTROS ACTIVOS DE DEUDA'
    WHEN fon.producto_estructurado   = TRUE THEN 'ESTRUCTURADOS'
    WHEN fon.producto_tvd            = TRUE THEN 'TVD'
    ELSE 'NO TENGO'
  END AS productos,
  TRIM(REGEXP_REPLACE(CONCAT(
    IF(fon.fuente_salario    = TRUE, 'Salario, ', ''),
    IF(fon.fuente_inmuebles  = TRUE, 'Inmuebles, ', ''),
    IF(fon.fuente_inversion  = TRUE, 'Inversiones, ', ''),
    IF(fon.fuente_donaciones = TRUE, 'Donaciones, ', ''),
    IF(fon.fuente_otros      = TRUE, COALESCE(CONCAT('Otros (', fon.fuente_detalle, ')'), 'Otros'), '')
  ), ', $', '')) AS origen_fondos,
  fon.objetivo_inversion AS destino_fondos
FROM `picasso-364722.salvadordali.clientes-looker-studio` AS s
LEFT JOIN `picasso-364722.salvadordali.clientes_legitimacion` AS pe
  ON s.codigo_cliente = pe.codigo_cliente
LEFT JOIN `picasso-364722.kyc.informacion_inversionista` AS kyc
  ON s.email = kyc.email
LEFT JOIN `picasso-364722.raffaello.users` AS u 
  ON s.email = u.email
LEFT JOIN `picasso-364722.kyc.informacion_fondos` AS fon
  ON u.uuid = fon.usuario_uuid
LEFT JOIN `picasso-364722.raffaello.user_investor_statements` AS uis
  ON u.id = uis.user_id
LEFT JOIN `picasso-364722.raffaello.investor_statements` AS is2 
  ON uis.investor_statement_id = is2.id
LEFT JOIN `picasso-364722.api_transunion_hcicvscvestimator.score` AS sc
  ON REPLACE(s.documento, '-', '') = sc.identificacion
WHERE (s.codigo_cliente NOT LIKE "CLI01%" 
       AND s.codigo_cliente NOT LIKE "CLI00%" 
       AND s.apellido NOT LIKE "%NO USAR%")
  AND (s.cuenta_custodio != "0" AND s.cuenta_custodio != "1")
ORDER BY RAND()
LIMIT 200
"""

# ════════════════════════════════════════════════════════════
# GENERACIÓN DE PDF
# ════════════════════════════════════════════════════════════
def generar_pdf_resultados(cliente, observacion, analista, posicion, weights):
    """Genera la hoja de Resultados imprimible en PDF con identidad CCI."""
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
        leftMargin=0.7*inch, rightMargin=0.7*inch,
        topMargin=0.7*inch, bottomMargin=0.7*inch)

    styles = getSampleStyleSheet()
    verde = colors.HexColor(CCI_VERDE)
    gris = colors.HexColor(CCI_GRIS)
    grisM = colors.HexColor(CCI_GRIS_MID)
    grisC = colors.HexColor(CCI_GRIS_CLARO)
    blanco = colors.white

    title_style = ParagraphStyle("title", parent=styles["Title"],
        fontName="Helvetica-Bold", fontSize=18, textColor=gris,
        alignment=TA_LEFT, spaceAfter=6)
    sub_style = ParagraphStyle("sub", parent=styles["Normal"],
        fontName="Helvetica", fontSize=10, textColor=grisM,
        alignment=TA_LEFT, spaceAfter=12)
    sect_style = ParagraphStyle("sect", parent=styles["Heading2"],
        fontName="Helvetica-Bold", fontSize=12, textColor=verde,
        spaceBefore=14, spaceAfter=8)
    body_style = ParagraphStyle("body", parent=styles["Normal"],
        fontName="Helvetica", fontSize=10, textColor=gris,
        alignment=TA_LEFT, spaceAfter=6)

    story = []

    header_data = [
        [Paragraph("<b>CCI Puesto de Bolsa, S.A.</b><br/>Área de Cumplimiento", body_style),
         Paragraph(f"<para align='right'>Fecha: {date.today().strftime('%d/%m/%Y')}<br/>"
                   f"Ley 155-17 — Reglamento Art. 44</para>", body_style)]
    ]
    header_tbl = Table(header_data, colWidths=[3.8*inch, 3.4*inch])
    header_tbl.setStyle(TableStyle([
        ("LINEBELOW", (0,0), (-1,-1), 1.5, verde),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
    ]))
    story.append(header_tbl)
    story.append(Spacer(1, 12))

    story.append(Paragraph("MATRIZ DE RIESGO DEL CLIENTE — PERSONA FÍSICA", title_style))
    story.append(Paragraph(
        "En cumplimiento del Art. 44 del Reglamento que regula la Prevención de Lavado de "
        "Activos, Financiamiento del Terrorismo y de la Proliferación de Armas de Destrucción "
        "Masiva. Ley 155-17.",
        sub_style))

    story.append(Paragraph("DATOS DEL CLIENTE", sect_style))
    cli_data = [
        ["Código de cliente:", str(cliente.get("codigo", "—"))],
        ["Nombre o razón social:", str(cliente.get("nombre", "—"))],
        ["Fecha de nacimiento:",
         cliente.get("fecha_nacimiento").strftime("%d/%m/%Y") if pd.notna(cliente.get("fecha_nacimiento")) else "—"],
        ["Edad:", str(int(cliente.get("edad", 0))) if pd.notna(cliente.get("edad")) else "—"],
    ]
    t = Table(cli_data, colWidths=[2.2*inch, 5.0*inch])
    t.setStyle(TableStyle([
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTNAME", (1,0), (1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 10),
        ("TEXTCOLOR", (0,0), (-1,-1), gris),
        ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#F5F5F5")),
        ("LINEBELOW", (0,0), (-1,-1), 0.3, grisC),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("RIGHTPADDING",(0,0), (-1,-1), 8),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING",(0,0), (-1,-1), 6),
    ]))
    story.append(t)

    story.append(Paragraph("EVALUACIÓN DE RIESGO", sect_style))

    criterios = [
        ("PEP's", "SÍ" if cliente.get("v_pep", 0) == 10 else "NO PEP", cliente.get("v_pep", 0), 0),
        ("Tipo de cliente", cliente.get("tipo_cliente", "No Profesional"), cliente.get("v_tipo", 0), weights.get("tipo_cliente", 0)),
        ("Método de Apertura", cliente.get("apertura", "Digital"), cliente.get("v_apertura", 5), weights.get("apertura", 0)),
        ("Condición laboral", cliente.get("cond_laboral", "—"), cliente.get("v_cond", 0), weights.get("cond_laboral", 0)),
        ("Sector Económico", cliente.get("sector", "—"), cliente.get("v_sector", 0), weights.get("sector", 0)),
        ("Profesión u Oficio", cliente.get("profesion", "—"), cliente.get("v_profesion", 0), weights.get("profesion", 0)),
        ("País de Origen", cliente.get("pais_origen", "—"), cliente.get("v_pais", 0), 0),
        ("País de Residencia", cliente.get("pais_residencia", "—"), cliente.get("v_pais", 0), 0),
        ("Provincia", cliente.get("provincia", "—"), cliente.get("v_pais", 0), 0),
        ("Tiempo residiendo en país", cliente.get("tiempo_residencia", "—"), cliente.get("v_tiempo_res", 0), weights.get("tiempo_residencia", 0)),
        ("Calificación de crédito", cliente.get("scoring_credito", "—"), cliente.get("v_credito", 0), weights.get("scoring_credito", 0)),
        ("Productos que utiliza", cliente.get("productos", "—"), cliente.get("v_productos", 0), weights.get("productos", 0)),
        ("Fecha de Nacimiento",
         cliente.get("fecha_nacimiento").strftime("%d/%m/%Y") if pd.notna(cliente.get("fecha_nacimiento")) else "—",
         cliente.get("v_edad", 0), weights.get("fecha_nac", 0)),
        ("Patrimonio total", cliente.get("patrimonio", "—"), cliente.get("v_patrimonio", 0), weights.get("patrimonio", 0)),
        ("Activos líquidos", cliente.get("activos_liquidos", "—"), cliente.get("v_activos", 0), weights.get("activos_liquidos", 0)),
        ("Ingresos anuales", cliente.get("ingresos", "—"), cliente.get("v_ingresos", 0), weights.get("ingresos", 0)),
        ("Monto promedio a movilizar", cliente.get("monto_movilizar", "—"), cliente.get("v_monto", 0), weights.get("monto_movilizar", 0)),
        ("Perfil de Inversionista", cliente.get("perfil_inv", "—"), cliente.get("v_perfil", 0), weights.get("perfil_inv", 0)),
        ("Duración de la Relación",
         "Cliente Nuevo" if cliente.get("anos_relac") == 0 else f"{int(cliente.get('anos_relac', 0))} años",
         cliente.get("v_duracion", 0), weights.get("duracion", 0)),
        ("Metamap", normalize_status(cliente.get("metamap", "")), cliente.get("v_metamap", 0), weights.get("metamap", 0)),
        ("Origen de Fondos", cliente.get("origen_fondos", "—"), round(cliente.get("v_origen", 0), 2), weights.get("origen_fondos", 0)),
        ("Destino / Objetivo", cliente.get("destino_fondos", "—"), cliente.get("v_destino", 0), weights.get("destino_fondos", 0)),
    ]

    eval_data = [["Criterio", "Respuesta", "Valor", "Ponderación"]]
    for crit, resp, val, peso in criterios:
        eval_data.append([str(crit), str(resp)[:50], str(val),
                          f"{peso*100:.2f}%" if peso > 0 else "—"])

    eval_tbl = Table(eval_data, colWidths=[2.0*inch, 3.0*inch, 0.7*inch, 1.0*inch], repeatRows=1)
    eval_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), gris),
        ("TEXTCOLOR", (0,0), (-1,0), blanco),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
        ("FONTNAME", (0,1), (-1,-1), "Helvetica"),
        ("TEXTCOLOR", (0,1), (-1,-1), gris),
        ("ALIGN", (2,1), (3,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.3, grisC),
        ("LEFTPADDING", (0,0), (-1,-1), 5),
        ("RIGHTPADDING", (0,0), (-1,-1), 5),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [blanco, colors.HexColor("#F8F8F8")]),
    ]))
    story.append(eval_tbl)

    story.append(Spacer(1, 6))
    nivel = cliente.get("nivel_riesgo", "Medio")
    nivel_color = {"Bajo": colors.HexColor("#5C8A1E"),
                   "Medio": colors.HexColor("#B8860B"),
                   "Alto": colors.HexColor("#C0392B"),
                   "Inadmisible": gris}.get(nivel, gris)

    result_data = [
        ["Promedio ponderado", f"{cliente.get('puntaje_final', 0):.2f}"],
        ["Promedio ponderado redondeado", f"{round(cliente.get('puntaje_final', 0)):.0f}"],
    ]
    rt = Table(result_data, colWidths=[5.7*inch, 1.5*inch])
    rt.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 10),
        ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#F5F5F5")),
        ("BACKGROUND", (1,0), (1,-1), colors.HexColor("#EAF4CE")),
        ("TEXTCOLOR", (0,0), (-1,-1), gris),
        ("ALIGN", (1,0), (1,-1), "CENTER"),
        ("GRID", (0,0), (-1,-1), 0.3, grisC),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    story.append(rt)

    story.append(Paragraph("CATEGORÍAS DE RIESGO", sect_style))
    cat_data = [
        ["Bajo", "Clientes que representan menos riesgo. Se realizará una debida diligencia simplificada."],
        ["Medio", "Clientes con riesgo moderado. Se realizará una debida diligencia regular."],
        ["Alto", "Clientes con mayor riesgo. Se realizará una debida diligencia ampliada (EDD)."],
    ]
    ct = Table(cat_data, colWidths=[0.9*inch, 6.3*inch])
    ct.setStyle(TableStyle([
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTNAME", (1,0), (1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("TEXTCOLOR", (0,0), (-1,-1), gris),
        ("BACKGROUND", (0,0), (0,0), colors.HexColor("#EAF4CE")),
        ("BACKGROUND", (0,1), (0,1), colors.HexColor("#FFF8DC")),
        ("BACKGROUND", (0,2), (0,2), colors.HexColor("#FDECEA")),
        ("ALIGN", (0,0), (0,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.3, grisC),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    story.append(ct)

    story.append(Paragraph("RESULTADOS", sect_style))
    tipo_dd = {"Bajo": "Debida Diligencia Simplificada",
               "Medio": "Debida Diligencia Regular",
               "Alto": "Debida Diligencia Ampliada (EDD)",
               "Inadmisible": "RECHAZAR — No establecer relación"}.get(nivel, "—")

    result_final = [
        ["Nivel de Riesgo del cliente", f"Riesgo {nivel}"],
        ["Tipo de debida diligencia", tipo_dd],
    ]
    rf = Table(result_final, colWidths=[2.5*inch, 4.7*inch])
    rf.setStyle(TableStyle([
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTNAME", (1,0), (1,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 11),
        ("BACKGROUND", (0,0), (0,-1), gris),
        ("TEXTCOLOR", (0,0), (0,-1), blanco),
        ("BACKGROUND", (1,0), (1,0), nivel_color),
        ("TEXTCOLOR", (1,0), (1,0), blanco),
        ("BACKGROUND", (1,1), (1,1), colors.HexColor("#F5F5F5")),
        ("TEXTCOLOR", (1,1), (1,1), gris),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.3, grisC),
        ("TOPPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
    ]))
    story.append(rf)

    # Info complementaria
    def safe_val(v):
        if v is None or (isinstance(v, float) and pd.isna(v)) or not str(v).strip():
            return "Sin información"
        return str(v).strip()

    story.append(Paragraph("INFORMACIÓN COMPLEMENTARIA", sect_style))
    info_data = [
        ["Perfil de Legitimación:", safe_val(cliente.get("perfil_legitimacion", ""))],
        ["Origen de Fondos:",       safe_val(cliente.get("origen_fondos", ""))],
        ["Destino / Objetivo:",     safe_val(cliente.get("destino_fondos", ""))],
    ]
    info_tbl = Table(info_data, colWidths=[2.0*inch, 5.2*inch])
    info_tbl.setStyle(TableStyle([
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTNAME", (1,0), (1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("TEXTCOLOR", (0,0), (-1,-1), gris),
        ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#F5F5F5")),
        ("LINEBELOW", (0,0), (-1,-1), 0.3, grisC),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
        ("TOPPADDING", (0,0), (-1,-1), 6),
        ("BOTTOMPADDING", (0,0), (-1,-1), 6),
    ]))
    story.append(info_tbl)

    if observacion and observacion.strip():
        story.append(Paragraph("OBSERVACIÓN DEL ANALISTA", sect_style))
        obs_style = ParagraphStyle("obs", parent=styles["Normal"],
            fontName="Helvetica", fontSize=9, textColor=gris,
            leading=12, alignment=TA_LEFT)
        obs_text = observacion.replace("\n", "<br/>")
        story.append(Paragraph(obs_text, obs_style))

    story.append(Spacer(1, 30))
    firma_data = [
        ["Preparado por:", f"{analista}", "Fecha:", date.today().strftime("%d/%m/%Y %H:%M")],
        ["Posición:", f"{posicion}", "Firma:", "_____________________"],
    ]
    fm = Table(firma_data, colWidths=[1.2*inch, 2.8*inch, 1.0*inch, 2.2*inch])
    fm.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTNAME", (2,0), (2,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("TEXTCOLOR", (0,0), (-1,-1), gris),
        ("LINEBELOW", (1,0), (1,-1), 0.3, grisC),
        ("LINEBELOW", (3,0), (3,-1), 0.3, grisC),
        ("TOPPADDING", (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ("LINEABOVE", (0,0), (-1,0), 1, verde),
    ]))
    story.append(fm)

    doc.build(story)
    buf.seek(0)
    return buf.getvalue()

def generar_excel(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        cols_export = [c for c in [
            "codigo","nombre","edad","anos_relac","pep","cond_laboral","sector",
            "profesion","pais_residencia","provincia","scoring_credito",
            "patrimonio","activos_liquidos","ingresos","monto_movilizar",
            "perfil_inv","metamap","origen_fondos","destino_fondos",
            "perfil_legitimacion","puntaje_final","nivel_riesgo",
            "razon_inadmisible","alerta_incoherencia","campos_faltantes"
        ] if c in df.columns]
        df[cols_export].to_excel(writer, sheet_name="Clientes", index=False)
        dist = df["nivel_riesgo"].value_counts().reindex(
            ["Bajo","Medio","Alto","Inadmisible"], fill_value=0).reset_index()
        dist.columns = ["Nivel", "Clientes"]
        dist["%"] = (dist["Clientes"] / len(df) * 100).round(1).astype(str) + "%"
        dist.to_excel(writer, sheet_name="Distribución", index=False)
    return buf.getvalue()

# ════════════════════════════════════════════════════════════
# SESSION STATE — Configuración cargada
# ════════════════════════════════════════════════════════════
if "config" not in st.session_state:
    st.session_state.config = cargar_config()
if "config_authenticated" not in st.session_state:
    st.session_state.config_authenticated = False

# ════════════════════════════════════════════════════════════
# SIDEBAR
# ════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"""
    <div style="text-align: center; padding: 1.5rem 0 1rem 0;
                border-bottom: 2px solid {CCI_VERDE}; margin-bottom: 1rem;">
        <div style="font-family: 'Georgia', serif; color: white;
                    font-size: 2.2rem; font-weight: 300; letter-spacing: 4px;
                    line-height: 1;">CCI</div>
        <div style="color: {CCI_VERDE}; font-size: 0.7rem;
                    letter-spacing: 3px; margin-top: 0.4rem;
                    text-transform: uppercase; font-weight: 600;">Puesto de Bolsa</div>
        <div style="color: {CCI_GRIS_CLARO}; font-size: 0.65rem;
                    letter-spacing: 2px; margin-top: 0.2rem; opacity: 0.7;">
            Cumplimiento AML/CFT</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"### <span style='color:{CCI_VERDE}'>●</span> Carga de datos", unsafe_allow_html=True)

    fuente = st.radio(
        "Fuente de clientes:",
        ["Data Warehouse (BigQuery)", "Subir archivo Excel"],
    )

    df_clientes = None

    if fuente == "Subir archivo Excel":
        file = st.file_uploader("Selecciona archivo:", type=["xlsx","csv"])
        if file:
            if file.name.endswith(".csv"):
                df_clientes = pd.read_csv(file)
            else:
                df_clientes = pd.read_excel(file)
            df_clientes = df_clientes.rename(columns=COLUMN_MAP)
            st.success(f"{len(df_clientes)} clientes cargados correctamente")
    else:
        try:
            from google.cloud import bigquery
            client_bq = bigquery.Client(project="picasso-364722")

            @st.cache_data(ttl=600, show_spinner="Consultando data warehouse...")
            def cargar_clientes_bigquery():
                return client_bq.query(SQL_BIGQUERY).to_dataframe()

            df_clientes = cargar_clientes_bigquery()
            df_clientes = df_clientes.rename(columns=COLUMN_MAP)
            st.success(f"{len(df_clientes)} clientes cargados desde el data warehouse")
        except Exception as e:
            st.error("No se pudo conectar al data warehouse")
            st.caption(f"Error: {str(e)[:200]}")
            st.info("Por ahora usa la opción 'Subir archivo Excel'")

    st.markdown("---")
    st.markdown(f"### <span style='color:{CCI_VERDE}'>●</span> Umbrales de clasificación", unsafe_allow_html=True)
    st.session_state.config["umbrales"]["bajo"] = st.slider(
        "Bajo ≤", 1.0, 6.0, float(st.session_state.config["umbrales"]["bajo"]), 0.5)
    st.session_state.config["umbrales"]["alto"] = st.slider(
        "Alto >", st.session_state.config["umbrales"]["bajo"], 9.0,
        float(st.session_state.config["umbrales"]["alto"]), 0.5)

    st.markdown("---")
    st.markdown(f"""
    <div style="font-size: 0.75rem; color: {CCI_GRIS_CLARO}; padding: 0.5rem 0;">
        <p>Ley 155-17 — Art. 44<br/>
        Reglamento PLAFT<br/>
        Versión Matriz: Mayo 2026</p>
    </div>
    """, unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# HEADER PRINCIPAL
# ════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="cci-header">
    <h1>Matriz de Riesgo AML/CFT</h1>
    <p>Plataforma de evaluación de riesgo para Prevención de Lavado de Activos</p>
</div>
""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
# CONTENIDO PRINCIPAL
# ════════════════════════════════════════════════════════════
if df_clientes is None or len(df_clientes) == 0:
    st.warning("Carga datos desde la barra lateral para comenzar.")
    st.markdown("### Acerca de la plataforma")
    col1, col2, col3 = st.columns(3)
    for col, (n, t) in zip([col1, col2, col3], [
        ("1", "Conexión al data warehouse del KYC"),
        ("2", "Cálculo automático con matriz Mayo 2026"),
        ("3", "Hoja de resultados PDF imprimible")]):
        with col:
            st.markdown(f"""
            <div class="stat-box">
                <h2 style="color:{CCI_VERDE}">{n}</h2>
                <p>{t}</p>
            </div>
            """, unsafe_allow_html=True)
    st.stop()

df = calcular_scores(df_clientes.copy(), st.session_state.config)
weights = normalizar_pesos(st.session_state.config["weights_raw"])

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "Dashboard",
    "Clientes Procesados",
    "Alertas",
    "Datos Faltantes",
    "Hoja de Resultados",
    "Exportar",
    "Configuración"
])

# ── TAB 1: DASHBOARD ─────────────────────────────────────
with tab1:
    st.markdown("### Resumen ejecutivo")
    dist = df["nivel_riesgo"].value_counts().reindex(
        ["Bajo","Medio","Alto","Inadmisible"], fill_value=0)
    total = len(df)

    c1, c2, c3, c4, c5 = st.columns(5)
    metrics = [
        ("Total Clientes", total, CCI_VERDE),
        ("Bajo", dist["Bajo"], "#5C8A1E"),
        ("Medio", dist["Medio"], "#B8860B"),
        ("Alto", dist["Alto"], "#C0392B"),
        ("Inadmisible", dist["Inadmisible"], CCI_GRIS),
    ]
    for col, (label, val, color) in zip([c1,c2,c3,c4,c5], metrics):
        with col:
            pct = val/total*100 if total > 0 else 0
            st.markdown(f"""
            <div class="stat-box" style="border-left-color:{color}">
                <h2 style="color:{color}">{val}</h2>
                <p>{label} <span style="color:{CCI_GRIS_MID}">({pct:.1f}%)</span></p>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("---")
    cA, cB = st.columns([2, 1])
    with cA:
        st.markdown("#### Distribución por nivel de riesgo")
        chart_data = pd.DataFrame({
            "Nivel": ["Bajo","Medio","Alto","Inadmisible"],
            "Clientes": [dist["Bajo"], dist["Medio"], dist["Alto"], dist["Inadmisible"]],
        })
        st.bar_chart(chart_data.set_index("Nivel"), color=CCI_VERDE, height=300)
    with cB:
        st.markdown("#### Estadísticas")
        st.metric("Puntaje promedio", f"{df['puntaje_final'].mean():.2f}")
        st.metric("Puntaje mediana", f"{df['puntaje_final'].median():.2f}")
        st.metric("Puntaje máximo", f"{df['puntaje_final'].max():.2f}")
        st.metric("Clientes PEP", f"{(df['v_pep']==10).sum()}")

# ── TAB 2: CLIENTES PROCESADOS ───────────────────────────
with tab2:
    st.markdown("### Clientes procesados")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        filtro_nivel = st.multiselect("Nivel de riesgo",
            ["Bajo","Medio","Alto","Inadmisible"],
            default=["Bajo","Medio","Alto","Inadmisible"])
    with c2:
        filtro_pep = st.selectbox("PEP", ["Todos", "Solo PEP=SI", "Solo PEP=NO"])
    with c3:
        if "sucursal" in df.columns:
            sucursales = ["Todas"] + sorted(df["sucursal"].dropna().unique().tolist())
            filtro_suc = st.selectbox("Sucursal", sucursales)
        else:
            filtro_suc = "Todas"
    with c4:
        busqueda = st.text_input("Buscar nombre o código")

    df_filtro = df[df["nivel_riesgo"].isin(filtro_nivel)]
    if filtro_pep == "Solo PEP=SI":
        df_filtro = df_filtro[df_filtro["v_pep"] == 10]
    elif filtro_pep == "Solo PEP=NO":
        df_filtro = df_filtro[df_filtro["v_pep"] != 10]
    if filtro_suc != "Todas":
        df_filtro = df_filtro[df_filtro["sucursal"] == filtro_suc]
    if busqueda:
        mask = (df_filtro.get("nombre", "").astype(str).str.contains(busqueda, case=False, na=False) |
                df_filtro.get("codigo", "").astype(str).str.contains(busqueda, case=False, na=False))
        df_filtro = df_filtro[mask]

    st.caption(f"Mostrando {len(df_filtro)} de {len(df)} clientes")
    cols_show = [c for c in ["codigo","nombre","edad","cond_laboral","sector",
                              "profesion","ingresos","perfil_inv",
                              "puntaje_final","nivel_riesgo"] if c in df_filtro.columns]
    st.dataframe(df_filtro[cols_show], use_container_width=True, height=500,
        column_config={
            "puntaje_final": st.column_config.NumberColumn("Puntaje", format="%.2f"),
            "nivel_riesgo": st.column_config.TextColumn("Nivel"),
        })

# ── TAB 3: ALERTAS ───────────────────────────────────────
with tab3:
    st.markdown("### Alertas de incoherencia patrimonial")
    alertas = df[df["alerta_incoherencia"] != ""]
    if len(alertas) == 0:
        st.success("No hay alertas de incoherencia en la muestra actual.")
    else:
        st.warning(f"{len(alertas)} clientes con incoherencia detectada")
        cols = [c for c in ["codigo","nombre","patrimonio","activos_liquidos",
                            "ingresos","nivel_riesgo","alerta_incoherencia"]
                if c in alertas.columns]
        st.dataframe(alertas[cols], use_container_width=True, height=400)

# ── TAB 4: DATOS FALTANTES ───────────────────────────────
with tab4:
    st.markdown("### Clientes con KYC incompleto")
    faltantes = df[df["campos_faltantes"] != ""]
    if len(faltantes) == 0:
        st.success("Todos los clientes tienen datos completos.")
    else:
        st.info(f"{len(faltantes)} clientes con datos faltantes — actualizar KYC")
        cols = [c for c in ["codigo","nombre","nivel_riesgo","campos_faltantes"]
                if c in faltantes.columns]
        st.dataframe(faltantes[cols], use_container_width=True, height=400)

# ── TAB 5: HOJA DE RESULTADOS ────────────────────────────
with tab5:
    st.markdown("### Generar hoja de Resultados para expediente")
    st.caption("Selecciona un cliente, agrega la observación del caso y genera el PDF imprimible.")

    if "nombre" in df.columns:
        nombres = df["nombre"].dropna().astype(str).tolist()
        codigos = df["codigo"].astype(str).tolist() if "codigo" in df.columns else [""]*len(df)
        opciones = [f"{c} — {n}" for c, n in zip(codigos, nombres)]
        sel = st.selectbox("Cliente:", opciones)
        idx = opciones.index(sel)
        cliente = df.iloc[idx]

        col1, col2 = st.columns([2, 1])
        with col1:
            st.markdown("**Datos del cliente:**")
            datos_show = {
                "Código": cliente.get("codigo", "—"),
                "Nombre": cliente.get("nombre", "—"),
                "Edad": int(cliente.get("edad", 0)) if pd.notna(cliente.get("edad")) else "—",
                "Sector": cliente.get("sector", "—"),
                "Profesión": cliente.get("profesion", "—"),
                "Ingresos": cliente.get("ingresos", "—"),
                "Puntaje final": f"{cliente.get('puntaje_final', 0):.2f}",
                "Nivel de Riesgo": cliente.get("nivel_riesgo", "—"),
            }
            for k, v in datos_show.items():
                st.text(f"{k:<20} {v}")

        with col2:
            nivel = cliente.get("nivel_riesgo", "—")
            color = {"Bajo":"#5C8A1E","Medio":"#B8860B","Alto":"#C0392B",
                     "Inadmisible":CCI_GRIS}.get(nivel, CCI_GRIS)
            st.markdown(f"""
            <div class="stat-box" style="border-left-color:{color}; text-align:center">
                <h2 style="color:{color}">{nivel}</h2>
                <p>Nivel asignado</p>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("---")

        # Info complementaria visible
        def clean_value(v):
            if v is None or (isinstance(v, float) and pd.isna(v)) or not str(v).strip():
                return "Sin información"
            return str(v).strip()

        info_cards = [
            ("Perfil de Legitimación", clean_value(cliente.get("perfil_legitimacion", ""))),
            ("Origen de Fondos",       clean_value(cliente.get("origen_fondos", ""))),
            ("Destino / Objetivo de la Inversión", clean_value(cliente.get("destino_fondos", ""))),
        ]
        for label, value in info_cards:
            st.markdown(f"""
            <div style="background-color: #F5F5F5; border-left: 4px solid {CCI_VERDE};
                        padding: 0.7rem 1rem; border-radius: 4px; margin-bottom: 0.6rem;">
                <div style="font-size: 0.72rem; color: {CCI_GRIS_MID}; 
                            text-transform: uppercase; letter-spacing: 1px;
                            font-weight: 600; margin-bottom: 0.25rem;">{label}</div>
                <div style="font-size: 0.95rem; color: {CCI_GRIS}; font-weight: 500;">{value}</div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("**Observación del analista**")
        observacion = st.text_area(
            "Escribe el análisis del caso, origen de fondos, sustentos documentales, etc.",
            height=180,
            placeholder="Ej: El cliente labora desde 2010 en empresa XYZ..."
        )

        col_a, col_b = st.columns(2)
        with col_a:
            analista = st.text_input("Preparado por:", value="Shelsy Alix")
        with col_b:
            posicion = st.text_input("Posición:", value="Analista de Cumplimiento")

        if st.button("Generar PDF imprimible", use_container_width=True):
            try:
                pdf_bytes = generar_pdf_resultados(cliente, observacion, analista, posicion, weights)
                st.success("PDF generado correctamente")
                st.download_button(
                    "Descargar PDF", pdf_bytes,
                    file_name=f"Matriz_Riesgo_{cliente.get('codigo','cliente')}_{date.today().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf", use_container_width=True,
                )
            except Exception as e:
                st.error(f"Error al generar PDF: {e}")

# ── TAB 6: EXPORTAR ──────────────────────────────────────
with tab6:
    st.markdown("### Exportar reportes")
    st.markdown("Descarga el dataset completo procesado.")
    excel_bytes = generar_excel(df)
    st.download_button(
        "Descargar Excel completo", excel_bytes,
        file_name=f"Matriz_Riesgo_AML_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ── TAB 7: CONFIGURACIÓN ─────────────────────────────────
with tab7:
    st.markdown("### Configuración de la Matriz de Riesgo")
    st.caption("Modifica pesos, umbrales y valores. Los cambios se guardan en archivo y se aplican al recalcular.")

    # Autenticación
    if not st.session_state.config_authenticated:
        st.warning("Esta sección requiere autenticación")
        col_pwd, _ = st.columns([1, 2])
        with col_pwd:
            pwd = st.text_input("Contraseña:", type="password")
            if st.button("Acceder"):
                if hash_password(pwd) == CONFIG_PASSWORD_HASH:
                    st.session_state.config_authenticated = True
                    st.rerun()
                else:
                    st.error("Contraseña incorrecta")
        st.stop()

    # Botón de cerrar sesión
    col_a, col_b, col_c = st.columns([3, 1, 1])
    with col_c:
        if st.button("Cerrar sesión"):
            st.session_state.config_authenticated = False
            st.rerun()

    # Sub-pestañas dentro de Configuración
    sub1, sub2, sub3 = st.tabs([
        "Pesos por criterio",
        "Tablas de valores",
        "Importar / Exportar"
    ])

    # ── SUB-TAB: PESOS ──
    with sub1:
        st.markdown("#### Pesos por criterio")
        st.caption("Ajusta cuánto pesa cada criterio. La aplicación normaliza automáticamente para que sumen 100%.")

        cfg = st.session_state.config["weights_raw"]
        nombres = {
            "tipo_cliente": "Tipo de cliente",
            "apertura": "Método de apertura",
            "cond_laboral": "Condición laboral",
            "sector": "Sector económico",
            "profesion": "Profesión",
            "tiempo_residencia": "Tiempo residiendo",
            "scoring_credito": "Score de crédito",
            "productos": "Productos que utiliza",
            "fecha_nac": "Edad (fecha nacimiento)",
            "patrimonio": "Patrimonio total",
            "activos_liquidos": "Activos líquidos",
            "ingresos": "Ingresos anuales",
            "monto_movilizar": "Monto a movilizar",
            "perfil_inv": "Perfil de inversionista",
            "duracion": "Duración de relación",
            "metamap": "Status Metamap",
            "origen_fondos": "Origen de fondos",
            "destino_fondos": "Destino de fondos",
        }

        # Mostrar pesos en dos columnas
        col_l, col_r = st.columns(2)
        keys_lista = list(nombres.keys())
        mitad = len(keys_lista) // 2 + 1

        nuevos_pesos = {}
        for i, k in enumerate(keys_lista):
            col = col_l if i < mitad else col_r
            with col:
                nuevos_pesos[k] = st.number_input(
                    f"{nombres[k]}",
                    min_value=0.0, max_value=1.0,
                    value=float(cfg.get(k, 0.0)),
                    step=0.01, format="%.4f",
                    key=f"peso_{k}"
                )
        nuevos_pesos["pep"] = 0.0  # PEP siempre 0 (es trigger directo a Alto)

        # Indicador de suma
        total_pesos = sum(nuevos_pesos.values())
        normalizado = normalizar_pesos(nuevos_pesos)

        st.markdown("---")
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            color_total = "#5C8A1E" if abs(total_pesos - 1.0) < 0.01 else "#B8860B"
            st.markdown(f"""
            <div class="stat-box" style="border-left-color:{color_total}">
                <h2 style="color:{color_total}">{total_pesos*100:.2f}%</h2>
                <p>Suma total de pesos crudos</p>
            </div>
            """, unsafe_allow_html=True)
        with col_t2:
            st.markdown(f"""
            <div class="stat-box" style="border-left-color:{CCI_VERDE}">
                <h2 style="color:{CCI_VERDE}">100.00%</h2>
                <p>Después de normalizar (siempre)</p>
            </div>
            """, unsafe_allow_html=True)

        # Tabla con pesos normalizados (lo que se usa realmente)
        with st.expander("Ver pesos normalizados que se aplicarán"):
            df_pesos = pd.DataFrame([
                {"Criterio": nombres.get(k, k),
                 "Peso crudo": f"{v:.4f}",
                 "Peso normalizado (%)": f"{normalizado.get(k, 0)*100:.2f}%"}
                for k, v in nuevos_pesos.items() if k != "pep"
            ])
            st.dataframe(df_pesos, use_container_width=True, hide_index=True)

        # Botones de acción
        st.markdown("---")
        col_b1, col_b2, col_b3 = st.columns(3)
        with col_b1:
            if st.button("Aplicar y guardar", use_container_width=True):
                st.session_state.config["weights_raw"] = nuevos_pesos
                if guardar_config(st.session_state.config):
                    st.success("Configuración guardada. Recargando…")
                    st.rerun()
        with col_b2:
            if st.button("Restaurar por defecto", use_container_width=True):
                st.session_state.config["weights_raw"] = dict(DEFAULT_WEIGHTS_RAW)
                if guardar_config(st.session_state.config):
                    st.success("Pesos restaurados a valores por defecto.")
                    st.rerun()
        with col_b3:
            if st.button("Solo aplicar (sin guardar)", use_container_width=True):
                st.session_state.config["weights_raw"] = nuevos_pesos
                st.info("Cambios aplicados solo en esta sesión.")
                st.rerun()

    # ── SUB-TAB: TABLAS DE VALORES ──
    with sub2:
        st.markdown("#### Tablas de valores por criterio")
        st.caption("Modifica los valores que se asignan a cada respuesta posible.")

        tabla_sel = st.selectbox("Selecciona la tabla a editar:", [
            "Sectores económicos",
            "Condición laboral",
            "Profesiones",
            "Patrimonio",
            "Ingresos",
            "Monto a movilizar",
            "Perfil de inversionista",
            "Productos",
            "Tiempo residiendo",
            "Origen de fondos",
            "Destino de fondos",
        ])

        tabla_map = {
            "Sectores económicos": "sector_scores",
            "Condición laboral": "cond_laboral_scores",
            "Profesiones": "profesion_scores",
            "Patrimonio": "patrimonio_scores",
            "Ingresos": "ingresos_scores",
            "Monto a movilizar": "monto_movilizar_scores",
            "Perfil de inversionista": "perfil_scores",
            "Productos": "productos_scores",
            "Tiempo residiendo": "tiempo_residencia_scores",
            "Origen de fondos": "origen_fondos_scores",
            "Destino de fondos": "destino_fondos_scores",
        }

        key_tabla = tabla_map[tabla_sel]
        tabla_actual = st.session_state.config[key_tabla]

        # Editor con st.data_editor
        df_editor = pd.DataFrame([
            {"Respuesta": k, "Valor de riesgo": v}
            for k, v in tabla_actual.items()
        ])

        st.markdown(f"**Editando: {tabla_sel}** ({len(df_editor)} valores)")
        st.caption("Edita los valores en la columna 'Valor de riesgo'. Valores de 0 a 10.")

        df_edited = st.data_editor(
            df_editor,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "Respuesta": st.column_config.TextColumn("Respuesta", width="large"),
                "Valor de riesgo": st.column_config.NumberColumn(
                    "Valor de riesgo", min_value=0, max_value=10, step=1, default=5),
            },
            hide_index=True,
            key=f"editor_{key_tabla}",
        )

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            if st.button("Guardar cambios", key=f"save_{key_tabla}", use_container_width=True):
                nueva_tabla = {}
                for _, row in df_edited.iterrows():
                    if pd.notna(row["Respuesta"]) and str(row["Respuesta"]).strip():
                        nueva_tabla[str(row["Respuesta"]).strip()] = int(row["Valor de riesgo"])
                st.session_state.config[key_tabla] = nueva_tabla
                if guardar_config(st.session_state.config):
                    st.success(f"Tabla '{tabla_sel}' guardada.")
                    st.rerun()
        with col_b2:
            if st.button("Restaurar esta tabla", key=f"reset_{key_tabla}", use_container_width=True):
                defaults_map = {
                    "sector_scores": DEFAULT_SECTOR_SCORES,
                    "cond_laboral_scores": DEFAULT_COND_LABORAL_SCORES,
                    "profesion_scores": DEFAULT_PROFESION_SCORES,
                    "patrimonio_scores": DEFAULT_PATRIMONIO_SCORES,
                    "ingresos_scores": DEFAULT_INGRESOS_SCORES,
                    "monto_movilizar_scores": DEFAULT_MONTO_MOVILIZAR_SCORES,
                    "perfil_scores": DEFAULT_PERFIL_SCORES,
                    "productos_scores": DEFAULT_PRODUCTOS_SCORES,
                    "tiempo_residencia_scores": DEFAULT_TIEMPO_RESIDENCIA_SCORES,
                    "origen_fondos_scores": DEFAULT_ORIGEN_FONDOS_SCORES,
                    "destino_fondos_scores": DEFAULT_DESTINO_FONDOS_SCORES,
                }
                st.session_state.config[key_tabla] = dict(defaults_map[key_tabla])
                if guardar_config(st.session_state.config):
                    st.success(f"Tabla '{tabla_sel}' restaurada por defecto.")
                    st.rerun()

    # ── SUB-TAB: IMPORTAR/EXPORTAR ──
    with sub3:
        st.markdown("#### Importar / Exportar configuración")
        st.caption("Respalda o transfiere la configuración como archivo JSON.")

        col_i, col_e = st.columns(2)
        with col_i:
            st.markdown("##### Exportar configuración actual")
            config_json = json.dumps(st.session_state.config, indent=2, ensure_ascii=False)
            st.download_button(
                "Descargar config.json",
                config_json,
                file_name=f"config_matriz_{date.today().strftime('%Y%m%d')}.json",
                mime="application/json",
                use_container_width=True,
            )

        with col_e:
            st.markdown("##### Importar configuración")
            archivo_cfg = st.file_uploader("Sube un archivo config.json", type=["json"])
            if archivo_cfg:
                try:
                    nueva_cfg = json.loads(archivo_cfg.read().decode("utf-8"))
                    if st.button("Aplicar configuración importada", use_container_width=True):
                        st.session_state.config = nueva_cfg
                        if guardar_config(nueva_cfg):
                            st.success("Configuración importada y guardada.")
                            st.rerun()
                except Exception as e:
                    st.error(f"Archivo inválido: {e}")

        st.markdown("---")
        st.markdown("##### Restaurar TODO a valores por defecto")
        st.warning("Esta acción revierte todos los pesos y tablas a los valores originales de la Matriz Mayo 2026.")
        if st.button("Restaurar TODO a por defecto", type="secondary"):
            st.session_state.config = {
                "weights_raw": dict(DEFAULT_WEIGHTS_RAW),
                "umbrales": dict(DEFAULT_UMBRALES),
                "sector_scores": dict(DEFAULT_SECTOR_SCORES),
                "cond_laboral_scores": dict(DEFAULT_COND_LABORAL_SCORES),
                "profesion_scores": dict(DEFAULT_PROFESION_SCORES),
                "patrimonio_scores": dict(DEFAULT_PATRIMONIO_SCORES),
                "ingresos_scores": dict(DEFAULT_INGRESOS_SCORES),
                "monto_movilizar_scores": dict(DEFAULT_MONTO_MOVILIZAR_SCORES),
                "perfil_scores": dict(DEFAULT_PERFIL_SCORES),
                "productos_scores": dict(DEFAULT_PRODUCTOS_SCORES),
                "tiempo_residencia_scores": dict(DEFAULT_TIEMPO_RESIDENCIA_SCORES),
                "country_scores": dict(DEFAULT_COUNTRY_SCORES),
                "provincia_scores": dict(DEFAULT_PROVINCIA_SCORES),
                "origen_fondos_scores": dict(DEFAULT_ORIGEN_FONDOS_SCORES),
                "destino_fondos_scores": dict(DEFAULT_DESTINO_FONDOS_SCORES),
            }
            if guardar_config(st.session_state.config):
                st.success("Toda la configuración restaurada.")
                st.rerun()

# ════════════════════════════════════════════════════════════
# FOOTER
# ════════════════════════════════════════════════════════════
st.markdown("---")
st.markdown(f"""
<div style="text-align: center; padding: 1rem 0; color: {CCI_GRIS_MID}; font-size: 0.8rem;">
    <strong>CCI Puesto de Bolsa, S.A.</strong> — Área de Cumplimiento — {date.today().strftime('%B %Y')}<br/>
    Plataforma Matriz de Riesgo AML/CFT — Ley 155-17, Reglamento Art. 44
</div>
""", unsafe_allow_html=True)
