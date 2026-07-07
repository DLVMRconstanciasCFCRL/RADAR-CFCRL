if st.button("Ejecutar Radar CFCRL"):
    with st.spinner("Preparando navegador de revisión..."):
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            capture_output=True,
            text=True,
            cwd=BASE_DIR
        )

    with st.spinner("Revisando portal del CFCRL. Esto puede tardar algunos minutos..."):
        resultado = subprocess.run(
            [sys.executable, RADAR_PATH],
            capture_output=True,
            text=True,
            cwd=BASE_DIR
        )
import streamlit as st
import os
import glob
import subprocess
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORTES_DIR = os.path.join(BASE_DIR, "reportes")
RADAR_PATH = os.path.join(BASE_DIR, "src", "radar_cfcrl.py")

st.set_page_config(
    page_title="Radar CFCRL",
    page_icon="⚖️",
    layout="wide"
)

st.title("Radar CFCRL")
st.subheader("Monitoreo de solicitudes de constancia de representatividad")

st.write(
    "Esta herramienta revisa el portal del Centro Federal de Conciliación y Registro Laboral, "
    "genera un reporte Word y envía una alerta por correo cuando detecta coincidencias."
)

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.metric("Sistema", "Activo")

with col2:
    st.metric("Última revisión", datetime.now().strftime("%d/%m/%Y %H:%M"))

st.divider()

if st.button("Ejecutar Radar CFCRL"):
    with st.spinner("Revisando portal del CFCRL. Esto puede tardar algunos minutos..."):
        resultado = subprocess.run(
            ["python3", RADAR_PATH],
            capture_output=True,
            text=True,
            cwd=BASE_DIR
        )

    if resultado.returncode == 0:
        st.success("Radar ejecutado correctamente.")
        st.text_area("Resultado", resultado.stdout, height=300)
    else:
        st.error("Ocurrió un error al ejecutar el radar.")
        st.text_area("Error", resultado.stderr, height=300)

st.divider()

st.header("Reportes generados")

reportes_word = sorted(
    glob.glob(os.path.join(REPORTES_DIR, "*.docx")),
    reverse=True
)

if not reportes_word:
    st.info("Todavía no hay reportes Word generados.")
else:
    for reporte in reportes_word[:10]:
        nombre = os.path.basename(reporte)

        with open(reporte, "rb") as archivo:
            st.download_button(
                label=f"Descargar {nombre}",
                data=archivo,
                file_name=nombre,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

st.divider()

st.header("Archivos de sistema")

st.write("Carpeta de reportes:")
st.code(REPORTES_DIR)

st.write("Archivo principal del radar:")
st.code(RADAR_PATH)
