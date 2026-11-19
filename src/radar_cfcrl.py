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
                             if score >= 96 else None)
            else:
                candidato = None
            if candidato and (mejor is None or candidato["score"] > mejor["score"]
                              or (candidato["score"] == mejor["score"]
                                  and len(normalizar(candidato["keyword"])) > len(normalizar(mejor["keyword"])))):
                mejor = candidato
        if mejor:
            coincidencias.append(mejor)
    return coincidencias


# El sitio usa encabezados Fecha, Empresa, Sindicato, Domicilio..., Folio.
# Se buscan tablas por sus encabezados, no por su posicion fija en el DOM.
EXTRAER_FILAS_JS = r"""
() => {
  const norm = s => (s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
     .toLowerCase().replace(/\s+/g, ' ').trim();
  const txt = e => (e?.innerText || e?.textContent || '').replace(/\s+/g, ' ').trim();
  const resultado = [];
  const tablas = Array.from(document.querySelectorAll('table'));
  for (const tabla of tablas) {
    const tr = Array.from(tabla.querySelectorAll('tr'));
    if (!tr.length) continue;
    let indiceEncabezado = -1;
    let indices = null;
    for (let i=0; i<Math.min(tr.length, 6); i++) {
      const c = Array.from(tr[i].querySelectorAll('th,td')).map(e=>norm(txt(e)));
      const buscar = p => c.findIndex(p);
      const fecha = buscar(s=>s==='fecha'||s.startsWith('fecha '));
      const empresa = buscar(s=>s.includes('empresa'));
      const sindicato = buscar(s=>s.includes('sindicato'));
      const domicilio = buscar(s=>s.includes('domicilio'));
      const folio = buscar(s=>s.includes('folio'));
      if ([fecha,empresa,sindicato,domicilio,folio].every(x=>x>=0)) {
        indiceEncabezado = i;
        indices = {fecha, empresa, sindicato, domicilio, folio};
        break;
      }
    }
    if (!indices) continue;
    for (const fila of tr.slice(indiceEncabezado+1)) {
      const celdas = Array.from(fila.children).filter(e=>['TD','TH'].includes(e.tagName));
      if (celdas.length <= Math.max(...Object.values(indices))) continue;
      const leer = campo => txt(celdas[indices[campo]]);
      const dato = {
        fecha: leer('fecha'), empresa: leer('empresa'), sindicato: leer('sindicato'),
        domicilio: leer('domicilio'), folio: leer('folio')
      };
      if (!dato.empresa || (!dato.folio && !dato.sindicato)) continue;
      const enlace = Array.from(fila.querySelectorAll('a[href]'))
         .map(a=>a.href).find(h=>h.startsWith('http')) || '';
      dato.enlace = enlace;
      resultado.push(dato);
    }
  }
  return resultado;
}
"""


def extraer_filas(page):
    return page.evaluate(EXTRAER_FILAS_JS)


def huella_filas(filas):
    return hashlib.sha256(json.dumps(filas, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()


def avanzar_pagina(page):
    """Conserva la estrategia de paginacion del radar anterior."""
    return page.evaluate("""
        () => {
            const elementos = Array.from(document.querySelectorAll('a, button, span'));
            const candidatos = elementos.filter(el => {
                const texto = (el.textContent || '').trim();
                const aria = (el.getAttribute('aria-label') || '').trim().toLowerCase();
                const title = (el.getAttribute('title') || '').trim().toLowerCase();
                return texto === '»' || texto === '›' || texto.toLowerCase() === 'siguiente'
                    || aria.includes('siguiente') || aria.includes('next')
                    || title.includes('siguiente') || title.includes('next');
            });
            if (!candidatos.length) return false;
            const boton = candidatos[candidatos.length - 1];
            const clase = String(boton.className || '').toLowerCase();
            const padre = String(boton.parentElement?.className || '').toLowerCase();
            if (boton.disabled || clase.includes('disabled') || padre.includes('disabled')
                || boton.getAttribute('aria-disabled') === 'true') return false;
            boton.click();
            return true;
        }
    """)


def ajustar_fuente(celda, puntos=7):
    celda.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for p in celda.paragraphs:
        p.paragraph_format.space_after = Pt(2)
        for r in p.runs:
            r.font.size = Pt(puntos)


def crear_reporte_word(reporte):
    fecha_archivo = datetime.now(ZoneInfo("America/Mexico_City")).strftime("%Y%m%d_%H%M%S")
    ruta = os.path.join(REPORTES_DIR, f"reporte_cfcrl_{fecha_archivo}.docx")
    documento = Document()
    sec = documento.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Inches(11.7), Inches(8.3)
    sec.left_margin = Inches(.55)
    sec.right_margin = Inches(.55)
    sec.top_margin = Inches(.65)
    sec.bottom_margin = Inches(.65)

    h = documento.add_heading("Reporte de Monitoreo CFCRL", 0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = documento.add_paragraph("Solicitudes de Constancia de Representatividad")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    documento.add_paragraph(f"Fecha de consulta: {reporte['fecha_consulta']}")
    documento.add_paragraph(f"Fuente revisada: {reporte['url']}")
    documento.add_paragraph(f"Solicitudes con coincidencias: {reporte['total_hallazgos']}")
    documento.add_paragraph(f"Páginas revisadas: {reporte['paginas_revisadas']}")
    documento.add_heading("Solicitudes detectadas", 1)
    if not reporte["hallazgos"]:
        documento.add_paragraph("No se identificaron coincidencias con la lista de clientes configurada.")
    else:
        columnas = ["Fecha", "Empresa", "Sindicato", "Domicilio donde se desarrolla la actividad", "Folio"]
        tabla = documento.add_table(rows=1, cols=5)
        tabla.style = "Table Grid"
        tabla.alignment = WD_TABLE_ALIGNMENT.CENTER
        tabla.autofit = False
        anchos = [0.9, 2.3, 2.55, 3.35, 1.45]
        for i, titulo in enumerate(columnas):
            tabla.rows[0].cells[i].text = titulo
            tabla.rows[0].cells[i].width = Inches(anchos[i])
            ajustar_fuente(tabla.rows[0].cells[i], 8)
            for parrafo in tabla.rows[0].cells[i].paragraphs:
                for run in parrafo.runs:
                    run.bold = True
        for h in reporte["hallazgos"]:
            celdas = tabla.add_row().cells
            for i, key in enumerate(["fecha", "empresa", "sindicato", "domicilio", "folio"]):
                celdas[i].text = h.get(key) or "No disponible"
                celdas[i].width = Inches(anchos[i])
                ajustar_fuente(celdas[i], 7)
        documento.add_paragraph("")
        documento.add_heading("Detalle de coincidencias", 1)
        detalle = documento.add_table(rows=1, cols=5)
        detalle.style = "Table Grid"
        titulos = ["Folio", "Cliente monitoreado", "Palabra detectada", "Tipo / score", "Página CFCRL"]
        for i, t in enumerate(titulos):
            detalle.rows[0].cells[i].text = t
            ajustar_fuente(detalle.rows[0].cells[i], 8)
        for h in reporte["hallazgos"]:
            for c in h["coincidencias"]:
                cel = detalle.add_row().cells
                valores = [h.get("folio", ""), c["cliente"], c["keyword"],
                           f"{c['tipo']} ({c['score']})", str(h["pagina_cfcrl"])]
                for i, v in enumerate(valores):
                    cel[i].text = v or "No disponible"
                    ajustar_fuente(cel[i], 8)

    documento.add_heading("Alcance y revisión", 1)
    documento.add_paragraph(
        "El monitoreo identifica coincidencias de palabras clave con el nombre de la empresa "
        "publicado en cada solicitud. Los resultados requieren corroboración en el portal oficial. "
        "Este reporte es una alerta interna y no constituye notificación procesal."
    )
    documento.add_paragraph("Evidencias: disponibles como archivos adjuntos de la ejecución de GitHub Actions (Artifacts).")
    documento.save(ruta)
    return ruta


def enviar_correo_reporte(reporte, word_path):
    asunto = ("ALERTA CFCRL: posibles coincidencias detectadas" if reporte["total_hallazgos"]
              else "Reporte CFCRL sin hallazgos")
    resumen = []
    for h in reporte["hallazgos"][:15]:
        resumen.append(f"Empresa: {h['empresa']}\nSindicato: {h['sindicato']}\n"
                       f"Fecha: {h['fecha']}\nFolio: {h['folio']}\n"
                       f"Domicilio: {h['domicilio']}\n"
                       f"Cliente monitoreado: {', '.join(sorted({c['cliente'] for c in h['coincidencias']}))}")
    contenido_resumen = "\n\n".join(resumen) if resumen else "No se identificaron coincidencias."
    if len(reporte["hallazgos"]) > 15:
        contenido_resumen += "\n\nEl resto de las solicitudes aparece en el reporte Word adjunto."
    cuerpo = f"""Estimado Oscar:

Se ejecutó el monitoreo automático de avisos de solicitud de constancia de representatividad del Centro Federal de Conciliación y Registro Laboral.

Fecha de consulta: {reporte['fecha_consulta']}
Solicitudes con coincidencias: {reporte['total_hallazgos']}
Páginas revisadas: {reporte['paginas_revisadas']}

{contenido_resumen}

Se adjunta el reporte Word con la fecha, empresa, sindicato, domicilio, folio y palabras detectadas.

Fuente: {reporte['url']}

Este correo constituye una alerta interna de monitoreo y no sustituye la confirmación de la publicación ni constituye notificación procesal.

Saludos cordiales.
Radar CFCRL
"""
    if not all([CORREO_REMITENTE, PASSWORD_CORREO, CORREO_DESTINATARIO]):
        raise ValueError("Faltan datos de correo en los secrets de GitHub o Streamlit")
    mensaje = EmailMessage()
    mensaje["From"] = CORREO_REMITENTE
    mensaje["To"] = CORREO_DESTINATARIO
    mensaje["Subject"] = asunto
    mensaje.set_content(cuerpo)
    with open(word_path, "rb") as f:
        mensaje.add_attachment(
            f.read(), maintype="application",
            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=os.path.basename(word_path)
        )
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(CORREO_REMITENTE, PASSWORD_CORREO)
        smtp.send_message(mensaje)


def revisar_cfcrl():
    clientes = cargar_clientes()
    fecha = datetime.now(ZoneInfo("America/Mexico_City"))
    fecha_archivo = fecha.strftime("%Y%m%d_%H%M%S")
    screenshot_path = os.path.join(EVIDENCIAS_DIR, f"cfcrl_{fecha_archivo}.png")
    html_path = os.path.join(EVIDENCIAS_DIR, f"cfcrl_{fecha_archivo}.html")
    reporte_path = os.path.join(REPORTES_DIR, f"reporte_{fecha_archivo}.json")
    hallazgos = []
    vistos = set()
    huellas = set()
    paginas = 0
    registros_analizados = 0

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1440, "height": 900})
        try:
            page.goto(URL, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(8000)
            for numero_pagina in range(1, 101):
                print(f"Revisando página {numero_pagina}", flush=True)
                page.wait_for_timeout(1500)
                filas = extraer_filas(page)
                if not filas:
                    # No declarar sin coincidencias si el sitio cambió o no cargó la tabla.
                    raise RuntimeError(
                        f"No se pudieron extraer solicitudes de la página {numero_pagina}. "
                        "Revisar estructura del portal antes de confiar en el reporte."
                    )
                huella = huella_filas(filas)
                if huella in huellas:
                    raise RuntimeError(
                        f"La página {numero_pagina} repitió registros anteriores. "
                        "La paginación no avanzó correctamente; revisión incompleta."
                    )
                huellas.add(huella)
                paginas += 1
                registros_analizados += len(filas)
                for fila in filas:
                    coincidencias = detectar_coincidencias(fila["empresa"], clientes)
                    if not coincidencias:
                        continue
                    # Misma solicitud publicada en páginas distintas: contar una vez.
                    clave = (normalizar(fila.get("folio")) if fila.get("folio") else
                             (normalizar(fila["empresa"]), normalizar(fila["sindicato"]),
                              normalizar(fila["domicilio"]), fila["fecha"]))
                    if clave in vistos:
                        continue
                    vistos.add(clave)
                    hallazgos.append({**fila, "pagina_cfcrl": numero_pagina,
                                      "coincidencias": coincidencias})
                    print("COINCIDENCIA:", fila["empresa"], "|", fila["folio"], flush=True)
                if numero_pagina == 100:
                    raise RuntimeError("Se alcanzó el límite de 100 páginas; revisión incompleta")
                pudo_avanzar = avanzar_pagina(page)
                if not pudo_avanzar:
                    print("Fin de la paginación.", flush=True)
                    break
                # Se deja tiempo para que el sitio actualice la tabla.
                page.wait_for_timeout(3500)
            page.screenshot(path=screenshot_path, full_page=True)
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(page.content())
        except Exception:
            page.screenshot(path=screenshot_path, full_page=True)
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(page.content())
            raise
        finally:
            browser.close()

    reporte = {
        "fecha_consulta": fecha.strftime("%d/%m/%Y %H:%M:%S (CDMX)"),
        "url": URL, "total_hallazgos": len(hallazgos),
        "paginas_revisadas": paginas, "registros_analizados": registros_analizados,
        "hallazgos": hallazgos, "captura": screenshot_path, "html": html_path
    }
    with open(reporte_path, "w", encoding="utf-8") as f:
        json.dump(reporte, f, ensure_ascii=False, indent=2)
    return reporte


if __name__ == "__main__":
    reporte = revisar_cfcrl()
    word_path = crear_reporte_word(reporte)
    print("REPORTE CFCRL")
    print("Fecha:", reporte["fecha_consulta"])
    print("Páginas revisadas:", reporte["paginas_revisadas"])
    print("Registros analizados:", reporte["registros_analizados"])
    print("Solicitudes con coincidencias:", reporte["total_hallazgos"])
    print("Word:", word_path)
    try:
        enviar_correo_reporte(reporte, word_path)
        print("Correo enviado a:", CORREO_DESTINATARIO)
    except Exception as e:
        print("No se pudo enviar el correo:", repr(e))
        raise

