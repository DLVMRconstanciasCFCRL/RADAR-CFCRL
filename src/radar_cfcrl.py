"""Radar CFCRL: monitoreo por registro de solicitudes de constancia."""

from playwright.sync_api import sync_playwright
from rapidfuzz import fuzz
from datetime import datetime
from zoneinfo import ZoneInfo
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
import pandas as pd
import unicodedata
import hashlib
import os
import json
import re
import smtplib
from email.message import EmailMessage


def limpiar_texto_correo(valor):
    if valor is None:
        return ""
    return str(valor).replace("\xa0", " ").replace("\u200b", "").strip()


def limpiar_password_correo(valor):
    if valor is None:
        return ""
    return str(valor).replace("\xa0", "").replace("\u200b", "").replace(" ", "").strip()


try:
    import streamlit as st
    CORREO_REMITENTE = limpiar_texto_correo(st.secrets["CORREO_REMITENTE"])
    PASSWORD_CORREO = limpiar_password_correo(st.secrets["PASSWORD_CORREO"])
    CORREO_DESTINATARIO = limpiar_texto_correo(st.secrets["CORREO_DESTINATARIO"])
except Exception:
    CORREO_REMITENTE = limpiar_texto_correo(os.getenv("CORREO_REMITENTE", ""))
    PASSWORD_CORREO = limpiar_password_correo(os.getenv("PASSWORD_CORREO", ""))
    CORREO_DESTINATARIO = limpiar_texto_correo(os.getenv("CORREO_DESTINATARIO", ""))

URL = "https://servicios.centrolaboral.gob.mx/constancia-de-representatividad/solicitudes"
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CLIENTES_PATH = os.path.join(BASE_DIR, "clientes", "clientes.csv")
EVIDENCIAS_DIR = os.path.join(BASE_DIR, "evidencias")
REPORTES_DIR = os.path.join(BASE_DIR, "reportes")
os.makedirs(EVIDENCIAS_DIR, exist_ok=True)
os.makedirs(REPORTES_DIR, exist_ok=True)


def normalizar(texto):
    if texto is None:
        return ""
    texto = unicodedata.normalize("NFKD", str(texto).lower())
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.replace("walmart", "wal mart").replace("wal-mart", "wal mart")
    return " ".join(re.sub(r"[^a-z0-9]+", " ", texto).split())


def cargar_clientes():
    df = pd.read_csv(CLIENTES_PATH, dtype=str, keep_default_na=False)
    if not {"cliente", "keywords"}.issubset(df.columns):
        raise ValueError("El CSV debe tener las columnas cliente y keywords")
    resultado = []
    for _, fila in df.iterrows():
        nombre = fila["cliente"].strip()
        keywords = [s.strip() for s in fila["keywords"].split(";") if s.strip()]
        if nombre and keywords:
            resultado.append({"cliente": nombre, "keywords": keywords})
    if not resultado:
        raise ValueError("No se encontraron clientes configurados")
    return resultado


def detectar_coincidencias(empresa, clientes):
    """Busca exclusivamente en el nombre de la EMPRESA de la fila, no en sindicato ni domicilio."""
    texto = normalizar(empresa)
    coincidencias = []
    if not texto:
        return coincidencias
    for cliente in clientes:
        mejor = None
        for keyword in cliente["keywords"]:
            patron = normalizar(keyword)
            if not patron:
                continue
            # Evita falsos positivos como 'office' dentro de palabras mayores.
            if re.search(r"(?<!\w)" + re.escape(patron) + r"(?!\w)", texto):
                candidato = {"cliente": cliente["cliente"], "keyword": keyword,
                             "tipo": "coincidencia exacta", "score": 100}
            elif len(patron) >= 18:
                score = round(fuzz.partial_ratio(patron, texto), 1)
                candidato = ({"cliente": cliente["cliente"], "keyword": keyword,
                              "tipo": "coincidencia aproximada", "score": score}
