import os
import re
import math
from io import BytesIO
from datetime import datetime

import pandas as pd
import streamlit as st
from supabase import create_client


# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
st.set_page_config(
    page_title="Sistema interno Mobil | Carga de datos",
    page_icon="🛢️",
    layout="wide",
)



# =========================================================
# PROTECCIÓN CON CONTRASEÑA
# =========================================================
PASSWORD_APP = "1256"

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.markdown(
        """
        <div style="
            background: linear-gradient(135deg, #0B2E63 0%, #003A8F 70%, #D71920 100%);
            padding: 35px;
            border-radius: 24px;
            color: white;
            text-align: center;
            margin-top: 80px;
            box-shadow: 0 14px 34px rgba(11,46,99,0.22);
        ">
            <h1>🔒 Acceso restringido</h1>
            <p>Ingresa la contraseña para utilizar el sistema.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    password_input = st.text_input(
        "Contraseña",
        type="password",
        placeholder="Ingresa la contraseña"
    )

    if st.button("Ingresar"):
        if password_input == PASSWORD_APP:
            st.session_state.autenticado = True
            st.rerun()
        else:
            st.error("❌ Contraseña incorrecta")

    st.stop()

# =========================================================
# ESTILO VISUAL SIMPLE / MOBIL
# =========================================================
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1180px;
    }

    .hero {
        background: linear-gradient(135deg, #0B2E63 0%, #003A8F 70%, #D71920 100%);
        border-radius: 24px;
        padding: 30px 34px;
        color: white;
        box-shadow: 0 14px 34px rgba(11,46,99,0.22);
        margin-bottom: 24px;
    }

    .hero-badge {
        display: inline-block;
        background: white;
        color: #0B2E63;
        padding: 7px 14px;
        border-radius: 999px;
        font-weight: 800;
        font-size: 0.86rem;
        margin-bottom: 14px;
    }

    .hero h1 {
        margin: 0;
        font-size: 2.15rem;
        font-weight: 850;
        letter-spacing: -0.5px;
    }

    .hero p {
        margin-top: 10px;
        font-size: 1.03rem;
        opacity: 0.95;
        max-width: 850px;
    }

    .section-card {
        border: 1px solid #E5E7EB;
        border-radius: 20px;
        padding: 22px 24px;
        background: #FFFFFF;
        box-shadow: 0 8px 22px rgba(16,24,40,0.06);
        margin-bottom: 18px;
    }

    .simple-step {
        background: #F8FAFC;
        border: 1px solid #E5E7EB;
        border-radius: 18px;
        padding: 16px 18px;
        height: 100%;
    }

    .step-number {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        width: 30px;
        height: 30px;
        background: #0B2E63;
        color: white;
        border-radius: 999px;
        font-weight: 800;
        margin-bottom: 8px;
    }

    .step-title {
        font-weight: 800;
        color: #0B2E63;
        margin-bottom: 4px;
    }

    .step-text {
        color: #667085;
        font-size: 0.93rem;
    }

    .metric-card {
        border-radius: 18px;
        background: #F8FAFC;
        border: 1px solid #E5E7EB;
        padding: 18px;
        min-height: 112px;
    }

    .metric-label {
        font-size: 0.88rem;
        color: #667085;
        margin-bottom: 8px;
    }

    .metric-value {
        font-size: 2.05rem;
        color: #0B2E63;
        font-weight: 850;
        line-height: 1.05;
    }

    .ok-box {
        background: #EAF7EF;
        border-left: 5px solid #2E7D32;
        padding: 14px 16px;
        border-radius: 14px;
        color: #174A28;
        margin: 12px 0;
        font-weight: 650;
    }

    .warn-box {
        background: #FFF7E6;
        border-left: 5px solid #F59E0B;
        padding: 14px 16px;
        border-radius: 14px;
        color: #7A4B00;
        margin: 12px 0;
        font-weight: 650;
    }

    .error-box {
        background: #FDECEC;
        border-left: 5px solid #D71920;
        padding: 14px 16px;
        border-radius: 14px;
        color: #7F1D1D;
        margin: 12px 0;
        font-weight: 650;
    }

    div.stButton > button:first-child {
        border-radius: 14px;
        border: 1px solid #0B2E63;
        background: #0B2E63;
        color: white;
        font-weight: 800;
        padding: 0.65rem 1.15rem;
    }

    div.stDownloadButton > button:first-child {
        border-radius: 14px;
        border: 1px solid #003A8F;
        color: #003A8F;
        font-weight: 800;
        padding: 0.65rem 1.15rem;
    }

    [data-testid="stFileUploader"] {
        border: 1px dashed #B8C5D6;
        border-radius: 20px;
        padding: 18px;
        background: #F8FAFC;
    }

    .footer-note {
        color: #667085;
        font-size: 0.86rem;
        margin-top: 18px;
        text-align: center;
    }

    .small-muted {
        color: #667085;
        font-size: 0.9rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <div class="hero-badge">Mobil | Sistema interno</div>
        <h1>🛢️ Carga de análisis de lubricantes</h1>
        <p>
            Herramienta para preparar y cargar archivos de análisis al chatbot de lubricación.
            Solo debes subir el Excel, revisar el resumen y confirmar la carga.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# CONEXIÓN INTERNA
# =========================================================
def obtener_secret(nombre: str, default: str = "") -> str:
    try:
        return str(st.secrets.get(nombre, default)).strip()
    except Exception:
        return os.getenv(nombre, default).strip()


SUPABASE_URL = obtener_secret("SUPABASE_URL")
SUPABASE_KEY = obtener_secret("SUPABASE_KEY")
SUPABASE_TABLE = obtener_secret("SUPABASE_TABLE", "ECOPETROL_COLOMBIA_DIC_2025")

conexion = None
if SUPABASE_URL and SUPABASE_KEY:
    try:
        conexion = create_client(SUPABASE_URL, SUPABASE_KEY)
    except Exception:
        conexion = None


# =========================================================
# UTILIDADES
# =========================================================
def col_index_to_letter(idx: int) -> str:
    s = ""
    i = int(idx)
    while i >= 0:
        s = chr(i % 26 + 65) + s
        i = i // 26 - 1
    return s


def df_to_xlsx_bytes(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Sheet1")
    buf.seek(0)
    return buf


def df_to_csv_bytes(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    df.to_csv(buf, index=False, encoding="utf-8-sig")
    buf.seek(0)
    return buf


def normalizar(col: str) -> str:
    return (
        str(col)
        .strip()
        .replace("–", "-")
        .replace("μ", "Μ")
        .replace("  ", " ")
        .upper()
    )


def limpiar_nombre_archivo(texto: str) -> str:
    texto = str(texto).strip()
    texto = re.sub(r"[\\/*?:\"<>|]", "", texto)
    texto = re.sub(r"\s+", "_", texto)
    texto = re.sub(r"_+", "_", texto)
    return texto[:80] if texto else "CUENTA_SIN_NOMBRE"


def obtener_nombre_cuenta(df: pd.DataFrame) -> str:
    for col in ["NOMBRE_CLIENTE", "NOMBRE_OPERACION"]:
        if col in df.columns:
            serie = df[col].dropna().astype(str).str.strip()
            serie = serie[serie != ""]
            if not serie.empty:
                return limpiar_nombre_archivo(serie.iloc[0])
    return "CUENTA_SIN_NOMBRE"


def limpiar_valor(v):
    if v is None:
        return None

    try:
        if pd.isna(v):
            return None
    except Exception:
        pass

    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return None

    if isinstance(v, pd.Timestamp):
        return v.isoformat()

    if isinstance(v, str):
        t = v.strip()
        if t.lower() in ["nan", "nat", "none", "null", "<na>"]:
            return None
        return t

    return v


def preparar_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.copy().astype(object)
    df2 = df2.where(pd.notna(df2), None)
    for col in df2.columns:
        df2[col] = df2[col].map(limpiar_valor)
    return df2


def dataframe_a_registros(df: pd.DataFrame) -> list[dict]:
    df2 = preparar_dataframe(df)
    return [
        {k: limpiar_valor(v) for k, v in row.items()}
        for row in df2.to_dict(orient="records")
    ]


def normalizar_n_muestra(v) -> str:
    if v is None:
        return ""
    t = str(v).strip()
    if t.lower() in ["nan", "none", "null", "nat", "<na>"]:
        return ""
    return t


def faltan_datos_conexion() -> bool:
    return not (SUPABASE_URL and SUPABASE_KEY and SUPABASE_TABLE and conexion)


def consultar_registros_existentes(n_muestras: list[str]) -> set[str]:
    existentes = set()
    valores = [normalizar_n_muestra(x) for x in n_muestras if normalizar_n_muestra(x)]
    valores = list(dict.fromkeys(valores))

    for i in range(0, len(valores), 500):
        lote = valores[i:i + 500]
        respuesta = (
            conexion
            .table(SUPABASE_TABLE)
            .select("N_MUESTRA")
            .in_("N_MUESTRA", lote)
            .execute()
        )
        for row in respuesta.data or []:
            existentes.add(normalizar_n_muestra(row.get("N_MUESTRA")))

    return existentes



def cargar_registros(df: pd.DataFrame, batch_size: int = 300) -> int:
    registros = dataframe_a_registros(df)
    total = len(registros)

    if total == 0:
        return 0

    barra = st.progress(0)
    cargados = 0

    for i in range(0, total, batch_size):
        lote = registros[i:i + batch_size]
        conexion.table(SUPABASE_TABLE).insert(lote).execute()
        cargados += len(lote)
        barra.progress(min(cargados / total, 1.0))

    return cargados


# =========================================================
# ENCABEZADOS BASE
# =========================================================
REQUERIDOS = [
    "NOMBRE_CLIENTE","NOMBRE_OPERACION","N_MUESTRA","CORRELATIVO","FECHA_MUESTREO","FECHA_INGRESO",
    "FECHA_RECEPCION","FECHA_INFORME","EDAD_COMPONENTE","UNIDAD_EDAD_COMPONENTE","EDAD_PRODUCTO",
    "UNIDAD_EDAD_PRODUCTO","CANTIDAD_ADICIONADA","UNIDAD_CANTIDAD_ADICIONADA","PRODUCTO","TIPO_PRODUCTO",
    "EQUIPO","TIPO_EQUIPO","MARCA_EQUIPO","MODELO_EQUIPO","COMPONENTE","MARCA_COMPONENTE","MODELO_COMPONENTE",
    "DESCRIPTOR_COMPONENTE","ESTADO_REPORTE","NIVEL_DE_SERVICIO",
    "ÍNDICE PQ (PQI) - 3","PLATA (AG) - 19","ALUMINIO (AL) - 20","CROMO (CR) - 24",
    "COBRE (CU) - 25","HIERRO (FE) - 26","TITANIO (TI) - 38","PLOMO (PB) - 35",
    "NÍQUEL (NI) - 32","MOLIBDENO (MO) - 30","SILICIO (SI) - 36","SODIO (NA) - 31",
    "POTASIO (K) - 27","VANADIO (V) - 39","BORO (B) - 18","BARIO (BA) - 21",
    "CALCIO (CA) - 22","CADMIO (CD) - 23","MAGNESIO (MG) - 28","MANGANESO (MN) - 29",
    "FÓSFORO (P) - 34","ZINC (ZN) - 40","CÓDIGO ISO (4/6/14) - 47",
    "CONTEO PARTÍCULAS >= 4 ΜM - 49","CONTEO PARTÍCULAS >= 6 ΜM - 50",
    "CONTEO PARTÍCULAS >= 14 ΜM - 48","OXIDACIÓN - 80","NITRACIÓN - 82",
    "NÚMERO ÁCIDO (AN) - 43","NÚMERO BÁSICO (BN) - 12","NÚMERO BÁSICO (BN) - 17",
    "HOLLÍN - 79","DILUCIÓN POR COMBUSTIBLE - 46","AGUA (IR) - 81",
    "CONTENIDO AGUA (KARL FISCHER) - 41","CONTENIDO GLICOL - 105",
    "VISCOSIDAD A 100 °C - 13","VISCOSIDAD A 40 °C - 14",
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51",
    "AGUA CUALITATIVA (PLANCHA) - 360",
    "AGUA LIBRE - 416","ANÁLISIS ANTIOXIDANTES (AMINA) - 44",
    "ANÁLISIS ANTIOXIDANTES (FENOL) - 45","COBRE (CU) - 119",
    "ESPUMA SEC 1 - ESTABILIDAD - 60","ESPUMA SEC 1 - TENDENCIA - 59",
    "ESTAÑO (SN) - 37","ÍNDICE VISCOSIDAD - 359","RPVOT - 10",
    "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6",
    "SEPARABILIDAD AGUA A 54 °C (AGUA) - 7",
    "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8",
    "SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83",
    "**ULTRACENTRÍFUGA (UC) - 1",
    "ESTADO_PRODUCTO","ESTADO_DESGASTE","ESTADO_CONTAMINACION",
    "N_SOLICITUD","CAMBIO_DE_PRODUCTO","CAMBIO_DE_FILTRO",
    "TEMPERATURA_RESERVORIO","UNIDAD_TEMPERATURA_RESERVORIO",
    "COMENTARIO_CLIENTE","TIPO_DE_COMBUSTIBLE","TIPO_DE_REFRIGERANTE",
    "USUARIO","COMENTARIO_REPORTE","id_muestra"
]

NUEVAS_ESTADO = [
    "ESTADO_MUESTRA",
    "AGUA (IR) - 74",
    "AGUA (IR) - 74 - Estado",
    "AGUA (IR) - 81 - Estado",
    "AGUA LIBRE - 416 - Estado",
    "AGUA CUALITATIVA (PLANCHA) - 360 - Estado",
    "ALUMINIO (AL) - 20 - Estado",
    "BARIO (BA) - 21 - Estado",
    "BORO (B) - 18 - Estado",
    "CALCIO (CA) - 22 - Estado",
    "CADMIO (CD) - 23 - Estado",
    "COBRE (CU) - 25 - Estado",
    "COBRE (CU) - 119 - Estado",
    "CROMO (CR) - 24 - Estado",
    "HIERRO (FE) - 26 - Estado",
    "MAGNESIO (MG) - 28 - Estado",
    "MANGANESO (MN) - 29 - Estado",
    "MOLIBDENO (MO) - 30 - Estado",
    "NÍQUEL (NI) - 32 - Estado",
    "PLATA (AG) - 19 - Estado",
    "PLOMO (PB) - 35 - Estado",
    "POTASIO (K) - 27 - Estado",
    "SILICIO (SI) - 36 - Estado",
    "SODIO (NA) - 31 - Estado",
    "TITANIO (TI) - 38 - Estado",
    "VANADIO (V) - 39 - Estado",
    "ZINC (ZN) - 40 - Estado",
    "ESTAÑO (SN) - 37 - Estado",
    "FÓSFORO (P) - 34 - Estado",
    "CÓDIGO ISO (4/6/14) - 47 - Estado",
    "CONTEO PARTÍCULAS >= 4 ΜM - 49 - Estado",
    "CONTEO PARTÍCULAS >= 6 ΜM - 50 - Estado",
    "CONTEO PARTÍCULAS >= 14 ΜM - 48 - Estado",
    "OXIDACIÓN - 80 - Estado",
    "NITRACIÓN - 82 - Estado",
    "ÍNDICE PQ (PQI) - 3 - Estado",
    "NÚMERO ÁCIDO (AN) - 43 - Estado",
    "NÚMERO BÁSICO (BN) - 12 - Estado",
    "NÚMERO BÁSICO (BN) - 17 - Estado",
    "CONTENIDO AGUA (KARL FISCHER) - 41 - Estado",
    "ANÁLISIS ANTIOXIDANTES (AMINA) - 44 - Estado",
    "ANÁLISIS ANTIOXIDANTES (FENOL) - 45 - Estado",
    "HOLLÍN - 73",
    "HOLLÍN - 73 - Estado",
    "HOLLÍN - 79 - Estado",
    "DILUCIÓN POR COMBUSTIBLE - 46 - Estado",
    "VISCOSIDAD A 40 °C - 14 - Estado",
    "VISCOSIDAD A 100 °C - 13 - Estado",
    "ÍNDICE VISCOSIDAD - 359 - Estado",
    "ESPUMA SEC 1 - ESTABILIDAD - 60 - Estado",
    "ESPUMA SEC 1 - TENDENCIA - 59 - Estado",
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51 - Estado",
    "RESIDUO CARBÓN (MCR) - 361",
    "RESIDUO CARBÓN (MCR) - 361 - Estado",
    "PUNTO DE INFLAMACIÓN (PMA) - 61",
    "PUNTO DE INFLAMACIÓN (PMA) - 61 - Estado",
    "RPVOT - 10 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (ACEITE) - 6 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (AGUA) - 7 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (EMULSIÓN) - 8 - Estado",
    "SEPARABILIDAD AGUA A 54 °C (TIEMPO) - 83 - Estado",
    "**ULTRACENTRÍFUGA (UC) - 1 - Estado"
]

ALIASES_ENTRADA = {
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51": [
        "** COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51"
    ],
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51 - Estado": [
        "** COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51 - Estado"
    ],
}


def posibles_entradas(nombre_salida: str) -> list[str]:
    return [nombre_salida] + ALIASES_ENTRADA.get(nombre_salida, [])


def encontrar_columna_origen(cols_norm_map: dict, nombre_salida: str) -> str | None:
    for candidato in posibles_entradas(nombre_salida):
        key = normalizar(candidato)
        if key in cols_norm_map:
            return cols_norm_map[key]
    return None


COLUMNAS_USADAS = REQUERIDOS + NUEVAS_ESTADO


# =========================================================
# FLUJO SIMPLE
# =========================================================
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("Proceso de carga")

p1, p2, p3 = st.columns(3)
with p1:
    st.markdown(
        """
        <div class="simple-step">
            <div class="step-number">1</div>
            <div class="step-title">Subir archivo</div>
            <div class="step-text">Selecciona el Excel exportado desde SmartAssistence.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with p2:
    st.markdown(
        """
        <div class="simple-step">
            <div class="step-number">2</div>
            <div class="step-title">Revisar resumen</div>
            <div class="step-text">Revisa cuántos registros están listos para cargar.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
with p3:
    st.markdown(
        """
        <div class="simple-step">
            <div class="step-number">3</div>
            <div class="step-title">Cargar datos</div>
            <div class="step-text">Confirma la carga para alimentar el chatbot.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
st.markdown("</div>", unsafe_allow_html=True)

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.subheader("1. Subir archivo Excel")
files = st.file_uploader(
    "Seleccione uno o varios archivos Excel",
    type="xlsx",
    accept_multiple_files=True,
)
st.markdown("</div>", unsafe_allow_html=True)

if files:
    dfs_out = []
    extras_global = []

    for f in files:
        df = pd.read_excel(f, dtype=str, engine="openpyxl")
        cols = df.columns.tolist()
        cols_norm = {normalizar(c): c for c in cols}

        faltantes = []
        for col_salida in COLUMNAS_USADAS:
            col_origen = encontrar_columna_origen(cols_norm, col_salida)
            if col_origen is None:
                faltantes.append(col_salida)

        if faltantes:
            st.markdown(
                '<div class="error-box">❌ El archivo no tiene la estructura esperada. Revise que sea el Excel exportado desde SmartAssistence.</div>',
                unsafe_allow_html=True,
            )
            with st.expander("Ver detalle para soporte"):
                st.write(f"Archivo: **{f.name}**")
                st.dataframe(
                    pd.DataFrame({"Encabezado faltante": faltantes}),
                    use_container_width=True,
                )
            st.stop()

        usadas_norm = set()
        for c in COLUMNAS_USADAS:
            for cand in posibles_entradas(c):
                usadas_norm.add(normalizar(cand))

        for idx, c in enumerate(cols):
            if normalizar(c) in usadas_norm:
                continue

            serie = df[c].astype(str).str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
            n = int(serie.notna().sum())

            if n > 0:
                extras_global.append({
                    "Archivo": f.name,
                    "Columna no usada": c,
                    "Registros con datos": n,
                    "Posición": col_index_to_letter(idx),
                })

        df_out = pd.DataFrame()

        for col_salida in REQUERIDOS:
            col_origen = encontrar_columna_origen(cols_norm, col_salida)
            df_out[col_salida] = df[col_origen]

        df_out.rename(columns={"ESTADO_REPORTE": "ESTADO"}, inplace=True)
        df_out["Archivo_Origen"] = f.name

        for col_salida in NUEVAS_ESTADO:
            col_origen = encontrar_columna_origen(cols_norm, col_salida)
            df_out[col_salida] = df[col_origen]

        dfs_out.append(df_out)

    df_final = pd.concat(dfs_out, ignore_index=True)

    df_final["N_MUESTRA"] = df_final["N_MUESTRA"].map(normalizar_n_muestra)

    registros_sin_numero = df_final[df_final["N_MUESTRA"] == ""]
    df_final = df_final[df_final["N_MUESTRA"] != ""].copy()

    df_final_limpio = preparar_dataframe(df_final)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("2. Resumen del archivo")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">Archivos procesados</div><div class="metric-value">{len(files)}</div></div>',
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">Registros listos</div><div class="metric-value">{len(df_final_limpio)}</div></div>',
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">Registros incompletos removidos</div><div class="metric-value">{len(registros_sin_numero)}</div></div>',
            unsafe_allow_html=True,
        )

    st.markdown(
        '<div class="ok-box">✅ Archivo validado correctamente y listo para cargar.</div>',
        unsafe_allow_html=True,
    )

    with st.expander("Ver vista previa de los registros"):
        st.dataframe(df_final_limpio.head(30), use_container_width=True)

    if extras_global:
        with st.expander("Avisos para soporte: columnas con datos no usadas"):
            st.dataframe(pd.DataFrame(extras_global), use_container_width=True)

    st.markdown("</div>", unsafe_allow_html=True)

    nombre_cuenta = obtener_nombre_cuenta(df_final_limpio)
    fecha_hora = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_excel = f"{nombre_cuenta}_{fecha_hora}.xlsx"
    nombre_csv = f"{nombre_cuenta}_{fecha_hora}.csv"

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Descarga de respaldo")
    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            "📥 Descargar Excel preparado",
            df_to_xlsx_bytes(df_final_limpio),
            file_name=nombre_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    with d2:
        st.download_button(
            "📥 Descargar CSV preparado",
            df_to_csv_bytes(df_final_limpio),
            file_name=nombre_csv,
            mime="text/csv",
        )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("3. Cargar datos al chatbot")

    if faltan_datos_conexion():
        st.markdown(
            '<div class="error-box">❌ La aplicación no está lista para cargar datos. Contacte al administrador.</div>',
            unsafe_allow_html=True,
        )
        st.stop()

    if st.button("🔎 Verificar registros existentes"):
        with st.spinner("Verificando registros..."):
            existentes = consultar_registros_existentes(df_final_limpio["N_MUESTRA"].tolist())
        st.session_state["registros_existentes"] = list(existentes)

    existentes = set(st.session_state.get("registros_existentes", []))

    nuevos_df = df_final_limpio[~df_final_limpio["N_MUESTRA"].isin(existentes)].copy()
    existentes_df = df_final_limpio[df_final_limpio["N_MUESTRA"].isin(existentes)].copy()

    r1, r2 = st.columns(2)
    with r1:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">Registros nuevos</div><div class="metric-value">{len(nuevos_df)}</div></div>',
            unsafe_allow_html=True,
        )
    with r2:
        st.markdown(
            f'<div class="metric-card"><div class="metric-label">Registros ya cargados</div><div class="metric-value">{len(existentes_df)}</div></div>',
            unsafe_allow_html=True,
        )

    if len(existentes_df) > 0:
        with st.expander("Ver registros ya cargados"):
            st.dataframe(
                existentes_df[["N_MUESTRA", "COMPONENTE", "FECHA_INFORME", "Archivo_Origen"]],
                use_container_width=True,
            )

    df_cargar = nuevos_df.copy()

    st.info(f"Registros que se cargarán: {len(df_cargar)}")

    confirmar = st.checkbox("Confirmo que revisé el resumen y autorizo la carga")

    if confirmar and st.button("🚀 Cargar datos al chatbot"):
        if len(df_cargar) == 0:
            st.warning("No hay registros nuevos para cargar.")
        else:
            try:
                with st.spinner("Cargando datos. Por favor espere..."):
                    total_cargados = cargar_registros(df_cargar, batch_size=300)

                st.markdown(
                    f'<div class="ok-box">✅ Carga completada correctamente. Registros cargados: {total_cargados}</div>',
                    unsafe_allow_html=True,
                )

            except Exception:
                st.markdown(
                    '<div class="error-box">❌ No fue posible completar la carga. Contacte al administrador.</div>',
                    unsafe_allow_html=True,
                )

    st.markdown("</div>", unsafe_allow_html=True)

st.markdown(
    '<div class="footer-note">Mobil LubeSoporte · Herramienta interna para alimentar el chatbot. La base realizará su control automático posterior.</div>',
    unsafe_allow_html=True,
)
