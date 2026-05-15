import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
import re
from supabase import create_client


# ===================== CONFIGURACIÓN =====================
st.set_page_config(page_title="SmartAssistence a Power BI / Supabase", layout="wide")
st.title("📄 Convertidor SmartAssistence a Power BI + Supabase")

st.info(
    "Instrucciones importantes:\n\n"
    "1. Sube el archivo Excel original exportado desde SmartAssistence.\n"
    "2. Descarga el archivo generado por esta herramienta.\n"
    "3. El archivo descargado NO debe ser modificado.\n"
    "4. No cambies el nombre del archivo generado.\n"
    "5. No edites columnas, encabezados, hojas, datos ni formato interno.\n"
    "6. Verifica todo antes de presionar el botón de carga a Supabase.\n"
    "7. La columna N_MUESTRA se controla como identificador único para evitar duplicados."
)


# ===================== CONEXIÓN SUPABASE =====================
def obtener_secret(nombre: str, requerido: bool = True, defecto: str = "") -> str:
    """Lee secretos desde Streamlit Cloud. No pongas credenciales en GitHub."""
    try:
        valor = st.secrets.get(nombre, defecto)
    except Exception:
        valor = defecto
    valor = str(valor or "").strip()
    if requerido and not valor:
        st.error(f"❌ Falta configurar el secreto: {nombre}")
        st.stop()
    return valor

SUPABASE_URL = obtener_secret("SUPABASE_URL")
SUPABASE_KEY = obtener_secret("SUPABASE_KEY")
SUPABASE_TABLE = obtener_secret("SUPABASE_TABLE", requerido=False, defecto="ECOPETROL_COLOMBIA_DIC_2025")

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

# ===================== UTILIDADES =====================
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


def normalizar_n_muestra(valor) -> str:
    if valor is None:
        return ""
    texto = str(valor).strip()
    if texto.lower() in ["nan", "none", "null", "na", "n/a"]:
        return ""
    return texto

def preparar_dataframe_para_supabase(df: pd.DataFrame) -> pd.DataFrame:
    """Convierte NaN a None y limpia textos para insertar en Supabase."""
    df2 = df.copy()
    df2 = df2.astype(object)
    df2 = df2.where(pd.notna(df2), None)
    for col in df2.columns:
        df2[col] = df2[col].apply(lambda x: None if x is None or str(x).strip().lower() in ["", "nan", "none", "null"] else str(x).strip())
    return df2

def dividir_en_lotes(lista: list, tamano: int = 500):
    for i in range(0, len(lista), tamano):
        yield lista[i:i + tamano]

def obtener_muestras_existentes(n_muestras: list[str]) -> set[str]:
    """Consulta en Supabase cuáles N_MUESTRA ya existen."""
    existentes = set()
    muestras_limpias = [m for m in n_muestras if m]
    for lote in dividir_en_lotes(muestras_limpias, 500):
        try:
            resp = (
                supabase.table(SUPABASE_TABLE)
                .select("N_MUESTRA")
                .in_("N_MUESTRA", lote)
                .execute()
            )
            for row in resp.data or []:
                valor = normalizar_n_muestra(row.get("N_MUESTRA"))
                if valor:
                    existentes.add(valor)
        except Exception as e:
            st.error(f"❌ Error consultando muestras existentes en Supabase: {e}")
            st.stop()
    return existentes

def insertar_en_supabase(df: pd.DataFrame, reemplazar_existentes: bool = False) -> dict:
    """Inserta registros. Si reemplazar_existentes=True, elimina primero esos N_MUESTRA en Supabase."""
    df_carga = preparar_dataframe_para_supabase(df)
    registros = df_carga.to_dict(orient="records")
    n_muestras = [normalizar_n_muestra(r.get("N_MUESTRA")) for r in registros]

    if reemplazar_existentes:
        for lote in dividir_en_lotes([m for m in n_muestras if m], 500):
            supabase.table(SUPABASE_TABLE).delete().in_("N_MUESTRA", lote).execute()

    total = len(registros)
    barra = st.progress(0)
    insertados = 0

    for i, lote in enumerate(dividir_en_lotes(registros, 500), start=1):
        supabase.table(SUPABASE_TABLE).insert(lote).execute()
        insertados += len(lote)
        barra.progress(min(insertados / total, 1.0))

    return {"insertados": insertados}

# ===================== ENCABEZADOS BASE =====================
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

# ===================== ENCABEZADOS ESTADO =====================
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

# ===================== ALIAS DE ENTRADA =====================
ALIASES_ENTRADA = {
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51": [
        "** COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51"
    ],
    "COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51 - Estado": [
        "** COLORIMETRÍA MEMBRANA DE PARCHE (MPC) - 51 - Estado"
    ],
}

# ===================== FUNCIONES PARA MAPEAR ENCABEZADOS =====================
def posibles_entradas(nombre_salida: str) -> list[str]:
    return [nombre_salida] + ALIASES_ENTRADA.get(nombre_salida, [])

def encontrar_columna_origen(cols_norm_map: dict, nombre_salida: str) -> str | None:
    for candidato in posibles_entradas(nombre_salida):
        key = normalizar(candidato)
        if key in cols_norm_map:
            return cols_norm_map[key]
    return None

# ===================== COLUMNAS USADAS =====================
COLUMNAS_USADAS = REQUERIDOS + NUEVAS_ESTADO

# ===================== CARGA DE ARCHIVOS =====================
files = st.file_uploader(
    "📤 Sube uno o varios Excel exportados desde SmartAssistence (.xlsx)",
    type="xlsx",
    accept_multiple_files=True
)

if files:
    dfs_out = []

    for f in files:
        df = pd.read_excel(f, dtype=str, engine="openpyxl")
        cols = df.columns.tolist()
        cols_norm = {normalizar(c): c for c in cols}

        # -------- VALIDACIÓN DE ENCABEZADOS --------
        faltantes = []
        for col_salida in COLUMNAS_USADAS:
            col_origen = encontrar_columna_origen(cols_norm, col_salida)
            if col_origen is None:
                faltantes.append(col_salida)

        if faltantes:
            st.error(f"❌ {f.name} – Faltan encabezados requeridos")
            st.dataframe(
                pd.DataFrame({"Encabezado faltante esperado en salida": faltantes}),
                use_container_width=True
            )
            st.stop()

        # -------- DETECCIÓN DE COLUMNAS CON DATOS NO USADAS --------
        usadas_norm = set()
        for c in COLUMNAS_USADAS:
            for cand in posibles_entradas(c):
                usadas_norm.add(normalizar(cand))

        extras = []
        for idx, c in enumerate(cols):
            if normalizar(c) in usadas_norm:
                continue

            serie = df[c].astype(str).str.strip().replace({"": pd.NA, "nan": pd.NA})
            n = int(serie.notna().sum())

            if n > 0:
                extras.append({
                    "Archivo": f.name,
                    "Encabezado NO usado": c,
                    "Registros con datos": n,
                    "Posición": col_index_to_letter(idx)
                })

        if extras:
            st.warning(f"⚠️ {f.name}: columnas con datos NO usadas en la salida")
            st.dataframe(pd.DataFrame(extras), use_container_width=True)

        # -------- CONSTRUCCIÓN DEL ARCHIVO FINAL --------
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

    # -------- CONTROL OBLIGATORIO DE N_MUESTRA ÚNICO --------
    if "N_MUESTRA" not in df_final.columns:
        st.error("❌ No existe la columna N_MUESTRA en el archivo transformado.")
        st.stop()

    df_final["N_MUESTRA"] = df_final["N_MUESTRA"].apply(normalizar_n_muestra)
    sin_muestra = df_final[df_final["N_MUESTRA"] == ""]

    if not sin_muestra.empty:
        st.error("❌ Hay filas sin N_MUESTRA. Corrige el archivo original antes de cargar a Supabase.")
        st.dataframe(sin_muestra.head(50), use_container_width=True)
        st.stop()

    duplicados_archivo = df_final[df_final.duplicated(subset=["N_MUESTRA"], keep=False)].copy()

    if not duplicados_archivo.empty:
        st.warning("⚠️ Se encontraron N_MUESTRA duplicados dentro del archivo cargado. Se conservará una sola fila por muestra.")
        columnas_mostrar = [c for c in ["N_MUESTRA", "NOMBRE_CLIENTE", "NOMBRE_OPERACION", "COMPONENTE", "FECHA_INFORME", "Archivo_Origen"] if c in duplicados_archivo.columns]
        st.dataframe(duplicados_archivo[columnas_mostrar].head(100), use_container_width=True)

        # Si viene duplicado, se conserva la última fila según el orden final del archivo procesado.
        df_final = df_final.drop_duplicates(subset=["N_MUESTRA"], keep="last").reset_index(drop=True)
        st.success("✅ Duplicados internos controlados: quedó una sola fila por N_MUESTRA.")

    st.success("✅ Conversión de SmartAssistence a Power BI completada correctamente")

    st.warning(
        "Antes de subir a la base de datos: revisa la vista previa y valida los avisos. "
        "La carga a Supabase se hará únicamente cuando presiones el botón final."
    )

    st.subheader("📌 Resumen de verificación")
    c1, c2, c3 = st.columns(3)
    c1.metric("Filas finales", len(df_final))
    c2.metric("N_MUESTRA únicos", df_final["N_MUESTRA"].nunique())
    c3.metric("Archivos procesados", len(files))

    st.dataframe(df_final.head(20), use_container_width=True)

    nombre_cuenta = obtener_nombre_cuenta(df_final)
    fecha_hora = datetime.now().strftime("%Y%m%d_%H%M%S")

    nombre_excel = f"{nombre_cuenta}_{fecha_hora}.xlsx"
    nombre_csv = f"{nombre_cuenta}_{fecha_hora}.csv"

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "📥 Descargar archivo final en Excel",
            df_to_xlsx_bytes(df_final),
            file_name=nombre_excel,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col2:
        st.download_button(
            "📥 Descargar archivo final en CSV",
            df_to_csv_bytes(df_final),
            file_name=nombre_csv,
            mime="text/csv"
        )

    # ===================== CARGA A SUPABASE =====================
    st.divider()
    st.subheader("🚀 Carga controlada a Supabase")
    st.caption(f"Tabla destino: {SUPABASE_TABLE}")

    n_muestras_finales = df_final["N_MUESTRA"].dropna().astype(str).str.strip().tolist()

    if st.button("🔎 Verificar contra Supabase antes de cargar"):
        existentes = obtener_muestras_existentes(n_muestras_finales)
        st.session_state["muestras_existentes"] = sorted(existentes)
        st.session_state["verificacion_supabase_ok"] = True

    if st.session_state.get("verificacion_supabase_ok"):
        existentes = set(st.session_state.get("muestras_existentes", []))
        nuevas = [m for m in n_muestras_finales if m not in existentes]

        col_a, col_b = st.columns(2)
        col_a.metric("Muestras nuevas", len(nuevas))
        col_b.metric("Muestras ya existentes en Supabase", len(existentes))

        if existentes:
            with st.expander("Ver N_MUESTRA que ya existen en Supabase"):
                st.dataframe(pd.DataFrame({"N_MUESTRA existente": sorted(existentes)}), use_container_width=True)

        modo_carga = st.radio(
            "¿Qué hacer si N_MUESTRA ya existe en Supabase?",
            [
                "Insertar solo muestras nuevas",
                "Reemplazar muestras existentes y luego insertar",
            ],
            index=0,
        )

        confirmar = st.checkbox("Confirmo que revisé la transformación y autorizo la carga a Supabase")

        if confirmar:
            if modo_carga == "Insertar solo muestras nuevas":
                df_cargar = df_final[~df_final["N_MUESTRA"].isin(existentes)].copy()
                reemplazar = False
            else:
                df_cargar = df_final.copy()
                reemplazar = True

            st.info(f"Filas que se cargarán: {len(df_cargar)}")

            if len(df_cargar) == 0:
                st.warning("⚠️ No hay muestras nuevas para cargar.")
            elif st.button("🚀 Subir definitivamente a Supabase"):
                try:
                    resultado = insertar_en_supabase(df_cargar, reemplazar_existentes=reemplazar)
                    st.success(f"✅ Carga completada. Registros insertados: {resultado['insertados']}")
                    st.session_state["verificacion_supabase_ok"] = False
                    st.session_state["muestras_existentes"] = []
                except Exception as e:
                    st.error(f"❌ Error cargando a Supabase: {e}")

