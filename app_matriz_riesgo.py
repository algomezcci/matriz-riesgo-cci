# ============================================================
# CCI Puesto de Bolsa, S.A. — Plataforma Matriz de Riesgo AML/CFT
# Ley 155-17 / Reglamento Art. 44
# ============================================================
# Cómo correr:
#   1. pip install streamlit pandas openpyxl reportlab
#   2. streamlit run app_matriz_riesgo.py
#   3. Abre http://localhost:8501 en el navegador
# ============================================================
 
import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, datetime
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
 
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
 
CCI_VERDE       = "#9EC229"
CCI_VERDE_DARK  = "#7A9A1E"
CCI_GRIS        = "#414042"
CCI_GRIS_MID    = "#6B6869"
CCI_AMARILLO    = "#FBDE64"
CCI_GRIS_CLARO  = "#D1D3D4"
 
st.markdown(f"""
<style>
    html, body, [class*="css"] {{
        font-family: 'Calibri', 'Segoe UI', sans-serif;
    }}
    .cci-header {{
        background-color: {CCI_GRIS};
        padding: 1rem 1.5rem;
        border-left: 6px solid {CCI_VERDE};
        margin-bottom: 1.5rem;
        border-radius: 4px;
    }}
    .cci-header h1 {{
        color: white;
        margin: 0;
        font-size: 1.6rem;
    }}
    .cci-header p {{
        color: {CCI_GRIS_CLARO};
        margin: 0.3rem 0 0 0;
        font-size: 0.9rem;
    }}
    .stat-box {{
        background-color: white;
        padding: 1.2rem;
        border-radius: 6px;
        border-left: 4px solid {CCI_VERDE};
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
    }}
    .stat-box h2 {{
        margin: 0;
        font-size: 2.2rem;
        color: {CCI_GRIS};
    }}
    .stat-box p {{
        margin: 0.2rem 0 0 0;
        color: {CCI_GRIS_MID};
        font-size: 0.85rem;
        text-transform: uppercase;
        letter-spacing: 1px;
    }}
    .stButton button {{
        background-color: {CCI_VERDE};
        color: white;
        border: none;
        font-weight: bold;
    }}
    .stButton button:hover {{
        background-color: {CCI_VERDE_DARK};
        color: white;
    }}
   .stTabs [data-baseweb="tab-list"] {{
        gap: 12px;
    }}
    .stTabs [data-baseweb="tab"] {{
        background-color: #f5f5f5;
        color: {CCI_GRIS};
        font-weight: 600;
        padding: 0.5rem 1.5rem;
        border-radius: 4px 4px 0 0;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: {CCI_VERDE} !important;
        color: white !important;
    }}
    section[data-testid="stSidebar"] {{
        background-color: {CCI_GRIS};
    }}
    section[data-testid="stSidebar"] * {{
        color: white !important;
    }}
    section[data-testid="stSidebar"] .stRadio label {{
        color: white !important;
    }}
</style>
""", unsafe_allow_html=True)
 
# ════════════════════════════════════════════════════════════
# TABLAS DE SCORING — Matriz Mayo 2026
# ════════════════════════════════════════════════════════════
SECTOR_SCORES = {
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
 
COND_LABORAL_SCORES = {
    "EMPLEADO": 3, "EJERCICIO INDEPENDIENTE": 5,
    "EJERCICIO INDEPENDIENTE DE LA PROFESION": 5, "INDEPENDIENTE": 5,
    "DUENO DE EMPRESA": 7, "DUEÑO DE EMPRESA": 7, "PROPIETARIO": 7,
    "PENSIONADO": 8, "RETIRADO": 8, "PENSIONADO/RETIRADO": 8,
    "DESEMPLEADO": 10, "ESTUDIANTE": 6,
}
 
PROFESION_SCORES = {
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
 
PATRIMONIO_SCORES = {
    "MENOS DE 500,000": 1, "DE 500,001 A 1,000,000": 2,
    "DE 1,000,001 A 10,000,000": 3,
    "MÁS DE 10,000,000": 6, "MAS DE 10,000,000": 6,
}
ACTIVOS_SCORES = dict(PATRIMONIO_SCORES)
 
INGRESOS_SCORES = {
    "NO TENGO": 2, "MENOS DE 500,000": 2,
    "DE 500,001 A 1,500,000": 4, "DE 501,000.00 A 1,500,000": 4,
    "DE 1,500,001 A 3,500,000": 6, "DE DOP 1,500,001 A DOP 3,500,000": 6,
    "DE 3,500,001 A 10,000,000": 8, "DE DOP 3,500,001 A DOP 10,000,000": 8,
    "MÁS DE 10,000,000": 10, "MAS DE 10,000,000": 10,
    "MAS DE DOP 10,000,000": 10,
}
 
MONTO_MOVILIZAR_SCORES = {
    "MENOS DE DOP 500,000": 2,
    "DE DOP 500,001 A 1,500,000": 4,
    "DE DOP 1,500,001 A 3,500,000": 6,
    "DE DOP 3,500,001 A 10,000,000": 8,
    "MAS DE DOP 10,000,000": 10,
    "MÁS DE DOP 10,000,000": 10,
}
 
PERFIL_SCORES = {
    "CONSERVADOR": 3, "C": 3, "MODERADO": 6, "M": 6, "AGRESIVO": 8, "A": 8,
}
 
PROVINCIA_SCORES = {
    "05": 5, "31": 5, "29": 3, "22": 6, "14": 10, "15": 4,
    "25": 6, "09": 1, "13": 7, "18": 7, "04": 9, "01": 5, "33": 5,
}
 
COUNTRY_SCORES = {
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
 
TIEMPO_RESIDENCIA_SCORES = {
    "MENOS DE 1 AÑO": 8, "DE 1 A 3 AÑOS": 5, "MAS DE 3 AÑOS": 3, "MÁS DE 3 AÑOS": 3,
}
 
PRODUCTOS_SCORES = {
    "DERIVADOS": 10, "DERIVADOS (SPOT NO LIQUIDADOS, FORWARDS)": 10,
    "ESTRUCTURADOS": 4, "OTROS ACTIVOS DE DEUDA": 6,
    "RENTA FIJA": 2, "RENTA VARIABLE": 3,
    "FONDOS DE INVERSION": 2, "FONDOS": 2,
}
 
TIPO_CLIENTE_SCORES = {"PROFESIONAL": 0, "NO PROFESIONAL": 0}
METAMAP_SCORES = {"approved": 0, "revisado": 2, "declined": 5}
 
WEIGHTS = {
    "pep": 0.0, "tipo_cliente": 0.028571, "apertura": 0.028571,
    "cond_laboral": 0.095238, "sector": 0.095238, "profesion": 0.095238,
    "pais_origen": 0.0, "pais_residencia": 0.0, "provincia": 0.0,
    "tiempo_residencia": 0.085714, "scoring_credito": 0.028571,
    "productos": 0.095238, "fecha_nac": 0.028571,
    "patrimonio": 0.042857, "activos_liquidos": 0.042857,
    "ingresos": 0.095238, "monto_movilizar": 0.095238,
    "perfil_inv": 0.047619, "duracion": 0.047619, "metamap": 0.047619,
}
 
DEFAULT_VALUES = {
    "cond_laboral": 4, "sector": 4, "profesion": 2,
    "pais_origen": 3, "pais_residencia": 3, "provincia": 5,
    "tiempo_residencia": 3, "scoring_credito": 2,
    "productos": 3, "patrimonio": 2, "activos_liquidos": 2,
    "ingresos": 3, "monto_movilizar": 4, "perfil_inv": 3,
    "fecha_nac": 3, "duracion": 5, "tipo_cliente": 0,
}
 
STATUS_MAP = {
    "VERIFIED": "approved", "APPROVED": "approved",
    "REVIEWNEEDED": "revisado", "REVIEWNEED": "revisado",
    "REJECTED": "declined", "DECLINED": "declined",
}
 
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
 
def score_pais(pais_res, provincia):
    pais = normalize(pais_res)
    if is_empty(pais_res): return DEFAULT_VALUES["pais_residencia"]
    if "DOMINICANA" in pais:
        prov = normalize(provincia)
        if not prov: return 5
        code = prov.split()[0]
        return PROVINCIA_SCORES.get(code, 5)
    return COUNTRY_SCORES.get(pais, 3)
 
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
 
def calcular_scores(df, umbral_bajo=5.0, umbral_alto=6.0):
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
        lambda x: fuzzy_get(x, COND_LABORAL_SCORES, DEFAULT_VALUES["cond_laboral"]))
    df["v_sector"] = df.get("sector", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, SECTOR_SCORES, DEFAULT_VALUES["sector"]))
    df["v_profesion"] = df.apply(lambda r:
        1 if "PENSION" in normalize(r.get("cond_laboral","")) or "RETIR" in normalize(r.get("cond_laboral",""))
        else fuzzy_get(r.get("profesion",""), PROFESION_SCORES, DEFAULT_VALUES["profesion"]), axis=1)
    df["v_pais"] = df.apply(lambda r: score_pais(r.get("pais_residencia",""), r.get("provincia","")), axis=1)
    df["v_tiempo_res"] = df.get("tiempo_residencia", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, TIEMPO_RESIDENCIA_SCORES, DEFAULT_VALUES["tiempo_residencia"]))
    df["v_credito"] = df.get("scoring_credito", pd.Series([""]*len(df))).apply(score_credito)
    df["v_productos"] = df.get("productos", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, PRODUCTOS_SCORES, DEFAULT_VALUES["productos"]))
    df["v_edad"] = df["edad"].apply(score_edad)
    df["v_patrimonio"] = df.get("patrimonio", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, PATRIMONIO_SCORES, DEFAULT_VALUES["patrimonio"]))
    df["v_activos"] = df.get("activos_liquidos", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, ACTIVOS_SCORES, DEFAULT_VALUES["activos_liquidos"]))
    df["v_ingresos"] = df.get("ingresos", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, INGRESOS_SCORES, DEFAULT_VALUES["ingresos"]))
    df["v_monto"] = df.get("monto_movilizar", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, MONTO_MOVILIZAR_SCORES, DEFAULT_VALUES["monto_movilizar"]))
    df["v_perfil"] = df.get("perfil_inv", pd.Series([""]*len(df))).apply(
        lambda x: fuzzy_get(x, PERFIL_SCORES, DEFAULT_VALUES["perfil_inv"]))
    df["v_duracion"] = df["anos_relac"].apply(score_duracion)
    df["v_metamap"] = df.get("metamap", pd.Series([""]*len(df))).apply(
        lambda x: METAMAP_SCORES.get(normalize_status(x), 2))
    df["v_tipo"] = 0
 
    df["puntaje_final"] = (
        df["v_tipo"] * WEIGHTS["tipo_cliente"] +
        df["v_apertura"] * WEIGHTS["apertura"] +
        df["v_cond"] * WEIGHTS["cond_laboral"] +
        df["v_sector"] * WEIGHTS["sector"] +
        df["v_profesion"] * WEIGHTS["profesion"] +
        df["v_tiempo_res"] * WEIGHTS["tiempo_residencia"] +
        df["v_credito"] * WEIGHTS["scoring_credito"] +
        df["v_productos"] * WEIGHTS["productos"] +
        df["v_edad"] * WEIGHTS["fecha_nac"] +
        df["v_patrimonio"] * WEIGHTS["patrimonio"] +
        df["v_activos"] * WEIGHTS["activos_liquidos"] +
        df["v_ingresos"] * WEIGHTS["ingresos"] +
        df["v_monto"] * WEIGHTS["monto_movilizar"] +
        df["v_perfil"] * WEIGHTS["perfil_inv"] +
        df["v_duracion"] * WEIGHTS["duracion"] +
        df["v_metamap"] * WEIGHTS["metamap"]
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
        return "; ".join(flags) if flags else ""
    df["alerta_incoherencia"] = df.apply(flag_incoherence, axis=1)
 
    CAMPOS = {
        "sector": "Sector", "profesion": "Profesión",
        "pais_residencia": "País Residencia", "provincia": "Provincia",
        "scoring_credito": "Score Crédito", "patrimonio": "Patrimonio",
        "activos_liquidos": "Activos Líq.", "ingresos": "Ingresos",
        "perfil_inv": "Perfil", "fecha_nacimiento": "Fecha Nac.",
        "metamap": "Status", "monto_movilizar": "Monto Mov.",
    }
    df["campos_faltantes"] = df.apply(
        lambda r: " | ".join([n for c,n in CAMPOS.items() if is_empty(r.get(c))]) , axis=1)
 
    return df
 
COLUMN_MAP = {
    "Email": "email", "Codigo": "codigo", "Nombre_Completo": "nombre",
    "PEP": "pep", "tipo_contrato": "tipo_contrato",
    "ocupacion": "cond_laboral", "Actividad_Economica": "sector",
    "profesion": "profesion", "Nacionalidad": "nacionalidad",
    "pais_residencia": "pais_residencia", "pais_origen": "pais_origen",
    "Sucursal": "sucursal", "Provincia": "provincia",
    "fecha_nacimiento": "fecha_nacimiento",
    "patrimonio_total": "patrimonio",
    "patrimonio_liquido": "activos_liquidos",
    "ingresos_anuales": "ingresos",
    "perfil_inversionista": "perfil_inv",
    "duracion_relacion": "duracion_relacion",
    "score": "scoring_credito", "status": "metamap",
    "tiempo_residencia": "tiempo_residencia",
    "monto_movilizar": "monto_movilizar",
    "productos": "productos", "tipo_cliente": "tipo_cliente",
}
 
def generar_pdf_resultados(cliente, observacion, analista, posicion):
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
        ("TEXTCOLOR", (0,0), (0,-1), gris),
        ("TEXTCOLOR", (1,0), (1,-1), gris),
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
        ("PEP's", "SÍ" if cliente.get("v_pep", 0) == 10 else "NO PEP", cliente.get("v_pep", 0), WEIGHTS["pep"]),
        ("Tipo de cliente", cliente.get("tipo_cliente", "No Profesional"), cliente.get("v_tipo", 0), WEIGHTS["tipo_cliente"]),
        ("Método de Apertura", "Digital", cliente.get("v_apertura", 5), WEIGHTS["apertura"]),
        ("Condición laboral", cliente.get("cond_laboral", "—"), cliente.get("v_cond", 0), WEIGHTS["cond_laboral"]),
        ("Sector Económico", cliente.get("sector", "—"), cliente.get("v_sector", 0), WEIGHTS["sector"]),
        ("Profesión u Oficio", cliente.get("profesion", "—"), cliente.get("v_profesion", 0), WEIGHTS["profesion"]),
        ("País de Origen", cliente.get("pais_origen", "—"), cliente.get("v_pais", 0), WEIGHTS["pais_origen"]),
        ("País de Residencia", cliente.get("pais_residencia", "—"), cliente.get("v_pais", 0), WEIGHTS["pais_residencia"]),
        ("Provincia", cliente.get("provincia", "—"), cliente.get("v_pais", 0), WEIGHTS["provincia"]),
        ("Tiempo residiendo en país", cliente.get("tiempo_residencia", "Más de 3 años"), cliente.get("v_tiempo_res", 0), WEIGHTS["tiempo_residencia"]),
        ("Calificación de crédito", cliente.get("scoring_credito", "—"), cliente.get("v_credito", 0), WEIGHTS["scoring_credito"]),
        ("Productos que utiliza", cliente.get("productos", "—"), cliente.get("v_productos", 0), WEIGHTS["productos"]),
        ("Fecha de Nacimiento",
         cliente.get("fecha_nacimiento").strftime("%d/%m/%Y") if pd.notna(cliente.get("fecha_nacimiento")) else "—",
         cliente.get("v_edad", 0), WEIGHTS["fecha_nac"]),
        ("Patrimonio total", cliente.get("patrimonio", "—"), cliente.get("v_patrimonio", 0), WEIGHTS["patrimonio"]),
        ("Activos líquidos", cliente.get("activos_liquidos", "—"), cliente.get("v_activos", 0), WEIGHTS["activos_liquidos"]),
        ("Ingresos anuales", cliente.get("ingresos", "—"), cliente.get("v_ingresos", 0), WEIGHTS["ingresos"]),
        ("Monto promedio a movilizar", cliente.get("monto_movilizar", "—"), cliente.get("v_monto", 0), WEIGHTS["monto_movilizar"]),
        ("Perfil de Inversionista", cliente.get("perfil_inv", "—"), cliente.get("v_perfil", 0), WEIGHTS["perfil_inv"]),
        ("Duración de la Relación",
         "Cliente Nuevo" if cliente.get("anos_relac") == 0 else f"{int(cliente.get('anos_relac', 0))} años",
         cliente.get("v_duracion", 0), WEIGHTS["duracion"]),
        ("Metamap", normalize_status(cliente.get("metamap", "")), cliente.get("v_metamap", 0), WEIGHTS["metamap"]),
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
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("FONTNAME", (0,1), (-1,-1), "Helvetica"),
        ("TEXTCOLOR", (0,1), (-1,-1), gris),
        ("ALIGN", (2,1), (3,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.3, grisC),
        ("LEFTPADDING", (0,0), (-1,-1), 5),
        ("RIGHTPADDING", (0,0), (-1,-1), 5),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
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
            "perfil_inv","metamap","puntaje_final","nivel_riesgo",
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
# SIDEBAR
# ════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"""
    <div style="text-align: center; padding: 1.5rem 0 1rem 0;
                border-bottom: 2px solid {CCI_VERDE}; margin-bottom: 1rem;">
        <div style="font-family: 'Georgia', serif; color: white;
                    font-size: 2.2rem; font-weight: 300; letter-spacing: 4px;
                    line-height: 1;">
            CCI
        </div>
        <div style="color: {CCI_VERDE}; font-size: 0.7rem;
                    letter-spacing: 3px; margin-top: 0.4rem;
                    text-transform: uppercase; font-weight: 600;">
            Puesto de Bolsa
        </div>
        <div style="color: {CCI_GRIS_CLARO}; font-size: 0.65rem;
                    letter-spacing: 2px; margin-top: 0.2rem; opacity: 0.7;">
            Cumplimiento AML/CFT
        </div>
    </div>
    """, unsafe_allow_html=True)
 
    st.markdown(f"### <span style='color:{CCI_VERDE}'>●</span> Carga de datos", unsafe_allow_html=True)
 
    fuente = st.radio(
        "Fuente de clientes:",
        ["Data Warehouse (KYC)", "Subir archivo Excel"],
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
        st.info("Conexión al DW pendiente (Fase 2). Por ahora usa la opción de subir Excel.")
 
    st.markdown("---")
    st.markdown(f"### <span style='color:{CCI_VERDE}'>●</span> Umbrales de clasificación", unsafe_allow_html=True)
    umbral_bajo = st.slider("Bajo ≤", 1.0, 6.0, 5.0, 0.5)
    umbral_alto = st.slider("Alto >", umbral_bajo, 9.0, 6.0, 0.5)
 
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
    with col1:
        st.markdown(f"""
        <div class="stat-box">
            <h2 style="color:{CCI_VERDE}">1</h2>
            <p>Conexión al data warehouse del KYC</p>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="stat-box">
            <h2 style="color:{CCI_VERDE}">2</h2>
            <p>Cálculo automático con matriz Mayo 2026</p>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="stat-box">
            <h2 style="color:{CCI_VERDE}">3</h2>
            <p>Hoja de resultados PDF imprimible</p>
        </div>
        """, unsafe_allow_html=True)
    st.stop()
 
df = calcular_scores(df_clientes.copy(), umbral_bajo=umbral_bajo, umbral_alto=umbral_alto)
 
tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "I · Dashboard",
    "II · Clientes Procesados",
    "III · Alertas",
    "IV · Datos Faltantes",
    "V · Hoja de Resultados",
    "VI · Exportar"
])
 
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
 
    st.dataframe(
        df_filtro[cols_show],
        use_container_width=True,
        height=500,
        column_config={
            "puntaje_final": st.column_config.NumberColumn("Puntaje", format="%.2f"),
            "nivel_riesgo": st.column_config.TextColumn("Nivel"),
        }
    )
 
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
        st.markdown("**Observación del analista**")
        observacion = st.text_area(
            "Escribe el análisis del caso, origen de fondos, sustentos documentales, etc.",
            height=180,
            placeholder="Ej: El cliente labora desde 2010 en empresa XYZ, los fondos provienen de prestaciones laborales conforme documentación presentada..."
        )
 
        col_a, col_b = st.columns(2)
        with col_a:
            analista = st.text_input("Preparado por:", value="Shelsy Alix")
        with col_b:
            posicion = st.text_input("Posición:", value="Analista de Cumplimiento")
 
        if st.button("Generar PDF imprimible", use_container_width=True):
            try:
                pdf_bytes = generar_pdf_resultados(cliente, observacion, analista, posicion)
                st.success("PDF generado correctamente")
                st.download_button(
                    "Descargar PDF",
                    pdf_bytes,
                    file_name=f"Matriz_Riesgo_{cliente.get('codigo','cliente')}_{date.today().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Error al generar PDF: {e}")
 
with tab6:
    st.markdown("### Exportar reportes")
    st.markdown("Descarga el dataset completo procesado con todos los criterios, scores y niveles.")
 
    excel_bytes = generar_excel(df)
    st.download_button(
        "Descargar Excel completo",
        excel_bytes,
        file_name=f"Matriz_Riesgo_AML_{date.today().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Incluye: hoja de clientes con todos los scores y hoja de distribución resumida.")
 
st.markdown("---")
st.markdown(f"""
<div style="text-align: center; padding: 1rem 0; color: {CCI_GRIS_MID}; font-size: 0.8rem;">
    <strong>CCI Puesto de Bolsa, S.A.</strong> — Área de Cumplimiento — {date.today().strftime('%B %Y')}<br/>
    Plataforma Matriz de Riesgo AML/CFT — Ley 155-17, Reglamento Art. 44
</div>
""", unsafe_allow_html=True)
