import streamlit as st
import pandas as pd
import tempfile, shutil, os, re
from funciones_asistencia import generar_asistencia, cargar_postulantes
from funciones_op1 import generar_op1, cargar_excel_con_encabezado_correcto
from funciones_op2 import generar_op2

# =========================
# CONFIGURACIÓN VISUAL
# =========================
st.set_page_config(page_title="PE3 - Generador de Reportes", layout="wide")

st.markdown("""
    <style>
        body {
            background-color: #f5f5f5;
            color: #333;
        }
        .block-container {
            padding-top: 1rem;
            padding-bottom: 1rem;
        }
        .stTabs [role="tablist"] {
            border-bottom: 1px solid #ccc;
        }
        .stTabs [role="tab"] {
            color: #444;
            background-color: #fff;
            border: 1px solid #ddd;
            border-bottom: none;
            padding: 0.5rem 1rem;
            border-radius: 5px 5px 0 0;
        }
        .stTabs [aria-selected="true"] {
            background-color: #eaeaea;
            border-color: #bbb;
            font-weight: 600;
        }
        .stButton>button {
            background-color: #4a90e2;
            color: white;
            border-radius: 8px;
            border: none;
            padding: 0.5rem 1.2rem;
        }
        .stButton>button:hover {
            background-color: #357ABD;
        }
    </style>
""", unsafe_allow_html=True)

# =========================
# CABECERA
# =========================
st.markdown("""
<div style="text-align:center; margin-bottom:10px;">
    <h2 style="color:#333;">Sistema de Generación de Reportes PE</h2>
    <p style="color:#666;">Cargue los archivos correspondientes y genere los reportes de Asistencia, OP1 y OP2</p>
</div>
""", unsafe_allow_html=True)

# =========================
# PLANTILLA BASE
# =========================
PLANTILLA_PATH = os.path.join("plantillas", "PE3 - Reporte.xlsx")
if not os.path.exists(PLANTILLA_PATH):
    st.error("No se encontró la plantilla base en /plantillas/.")
    st.stop()

def get_temp_copy():
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    shutil.copyfile(PLANTILLA_PATH, tmp.name)
    return tmp.name

# =========================
# VALIDACIÓN DE ARCHIVOS
# =========================
def validar_archivo(file, tipo, label=None):
    """
    Valida el archivo según su tipo (postulantes, instrumentos o FA)
    y su nombre (ASC, NOM, ACC). Retorna (True, columnas) si es válido,
    o (False, mensaje de error).
    """
    from openpyxl import load_workbook

    try:
        nombre = file.name.upper()

        # -------------------------------
        # 1️⃣ Validar tipo y nombre esperado
        # -------------------------------
        if label:
            # Normalizamos: eliminamos espacios, guiones y caracteres especiales
            label_norm = re.sub(r'[^A-Z0-9]', '', label.upper())
            nombre_norm = re.sub(r'[^A-Z0-9]', '', nombre)
            if label_norm not in nombre_norm:
                raise ValueError(f"El archivo cargado no corresponde a {label}. Subiste: {file.name}")

        # -------------------------------
        # 2️⃣ Validación de estructura por tipo
        # -------------------------------
        if tipo == "postulantes":
            df = cargar_postulantes(file)

            columnas_requeridas = [
                "Sede",
                "Postulantes",
                "Asistencia al Local",
                "Asistencia en Aula",
                "Casos de inconsistencia"
            ]
            faltantes = [c for c in columnas_requeridas if c not in df.columns]
            if faltantes:
                raise ValueError(f"Faltan columnas requeridas: {', '.join(faltantes)}")

            return True, list(df.columns)

        elif tipo in ["instrumentos", "fa"]:
            df = cargar_excel_con_encabezado_correcto(file)
            columnas_requeridas = ["Sede Operativa", "Local", "Tipo", "Inventario en campo"]
            faltantes = [c for c in columnas_requeridas if c not in df.columns]
            if faltantes:
                raise ValueError(f"Faltan columnas requeridas: {', '.join(faltantes)}")
            return True, list(df.columns)

        else:
            raise ValueError("Tipo de archivo no reconocido.")

    except Exception as e:
        return False, str(e)

# =========================
# INTERFAZ EN PESTAÑAS
# =========================
tabs = st.tabs(["Postulantes", "Instrumentos", "Formatos Auxiliares", "Generar Reportes"])

# --- TAB 1: POSTULANTES ---
with tabs[0]:
    st.subheader("Archivos de Postulantes")
    col1, col2, col3 = st.columns(3)
    with col1: asc = st.file_uploader("ASC - Postulantes", type=["xlsx"])
    with col2: nom = st.file_uploader("NOM - Postulantes", type=["xlsx"])
    with col3: acc = st.file_uploader("ACC - Postulantes", type=["xlsx"])

    for label, file in zip(["ASC", "NOM", "ACC"], [asc, nom, acc]):
        if file:
            ok, info = validar_archivo(file, "postulantes", label)
            if ok:
                st.success(f"{label} cargado correctamente.")
                st.caption(f"Columnas detectadas: {', '.join(info[:5])}...")
            else:
                st.error(f"{label} no válido: {info}")

# --- TAB 2: INSTRUMENTOS ---
with tabs[1]:
    st.subheader("Archivos de Instrumentos")
    col1, col2, col3 = st.columns(3)
    with col1: asc_inst = st.file_uploader("ASC - Instrumentos", type=["xlsx"])
    with col2: nom_inst = st.file_uploader("NOM - Instrumentos", type=["xlsx"])
    with col3: acc_inst = st.file_uploader("ACC - Instrumentos", type=["xlsx"])

    for label, file in zip(["ASC", "NOM", "ACC"], [asc_inst, nom_inst, acc_inst]):
        if file:
            ok, info = validar_archivo(file, "instrumentos", label)
            if ok:
                st.success(f"{label} válido.")
            else:
                st.error(f"{label} no válido: {info}")

# --- TAB 3: FORMATOS AUXILIARES ---
with tabs[2]:
    st.subheader("Archivos de Formatos Auxiliares (FA)")
    col1, col2 = st.columns(2)
    with col1: asc_fa = st.file_uploader("ASC - FA", type=["xlsx"])
    with col2: acc_fa = st.file_uploader("ACC - FA", type=["xlsx"])

    for label, file in zip(["ASC FA", "ACC FA"], [asc_fa, acc_fa]):
        if file:
            ok, info = validar_archivo(file, "fa", label)
            if ok:
                st.success(f"{label} válido.")
            else:
                st.error(f"{label} no válido: {info}")

# --- TAB 4: GENERACIÓN ---
with tabs[3]:
    st.subheader("Generación de Reportes")
    st.info("Los botones se activan cuando los archivos requeridos están correctamente cargados.")

    colA, colB, colC = st.columns(3)
    with colA:
        if st.button("Generar ASISTENCIA", disabled=not all([asc, nom, acc])):
            base_temp = get_temp_copy()
            generar_asistencia(base_temp, asc, nom, acc)
    with colB:
        if st.button("Generar OP1", disabled=not all([asc_inst, nom_inst, asc_fa])):
            base_temp = get_temp_copy()
            generar_op1(base_temp, asc_fa, asc_inst, nom_inst)
    with colC:
        if st.button("Generar OP2", disabled=not all([acc_inst, acc_fa])):
            base_temp = get_temp_copy()
            generar_op2(base_temp, acc_fa, acc_inst)

    st.divider()

    if "asistencia_generada" in st.session_state:
        st.download_button("Descargar Reporte ASISTENCIA", st.session_state["asistencia_generada"], file_name="PE3 - Reporte_ASISTENCIA.xlsx")
    if "op1_generada" in st.session_state:
        st.download_button("Descargar Reporte OP1", st.session_state["op1_generada"], file_name="PE3 - Reporte_OP1.xlsx")
    if "op2_generada" in st.session_state:
        st.download_button("Descargar Reporte OP2", st.session_state["op2_generada"], file_name="PE3 - Reporte_OP2.xlsx")