# limites_condenatorios.py
# Autor: Javier Parada
# Entrada: Excel TAL CUAL viene de SmartAssintence (sin modificar el archivo)
# Cambios solicitados:
# 1) Filtro por nombre del lubricante (PRODUCTO)
# 2) Variables: solo variables "lógicas" (pareja: VARIABLE y "VARIABLE - Estado")
#    -> el usuario puede escoger cuáles variables usar
# 3) Opción para excluir valores atípicos según el ESTADO:
#    -> si "VARIABLE - Estado" está en ALERTA (o las que el usuario elija), ese registro NO se usa para el cálculo.

import streamlit as st
import pandas as pd
import numpy as np

# =========================
# Configuración
# =========================
st.set_page_config(page_title="Límites Condenatorios - por Equipo", layout="wide")
st.title("📏 Límites Condenatorios por Equipo (Histórico)")

st.markdown("""
### Cómo funciona (paso a paso)
1) **Cargas el Excel** tal cual viene de SmartAssintence (la app no lo modifica).  
2) **Filtras** si quieres por **Lubricante (PRODUCTO)**, Operación, Tipo de equipo y luego eliges el/los **equipos** a evaluar.  
3) En **Variables**, la app solo te muestra variables “lógicas”: aquellas que vienen en pareja **(Variable + “Variable - Estado”)**.  
   - Ejemplo: **Hierro** y **Hierro - Estado** ⇒ se usa **Hierro** para el cálculo, y **Hierro - Estado** para decidir si se incluye o no el dato.  
4) Si lo decides, puedes **excluir del cálculo** los registros donde el estado de la variable esté en **ALERTA** (o cualquier estado que marques).  
5) Con el histórico filtrado, por cada **Equipo + Variable** calcula límites:
   - **Si n ≥ umbral**: usa **Percentiles** (Precaución=P90, Condenatorio=P95 por defecto).  
   - **Si n < umbral**: usa **Media + k·Desviación** (Prec=μ+2σ, Cond=μ+3σ por defecto).  
6) Te entrega la tabla final y la puedes **exportar a CSV** para Power BI / reportes.
""")

# =========================
# Carga (tal cual)
# =========================
@st.cache_data(show_spinner=False)
def load_excel(file) -> pd.DataFrame:
    return pd.read_excel(file)

archivo = st.file_uploader("📁 Sube tu Excel (.xlsx) de SmartAssintence (tal cual)", type=["xlsx"])
if not archivo:
    st.stop()

df = load_excel(archivo).copy()

# =========================
# Validaciones mínimas
# =========================
required_min = ['EQUIPO', 'FECHA_INFORME']
missing = [c for c in required_min if c not in df.columns]
if missing:
    st.error(f"❌ Falta(n) columna(s) mínima(s) requerida(s): {missing}")
    st.stop()

df['FECHA_INFORME'] = pd.to_datetime(df['FECHA_INFORME'], errors='coerce')
if df['FECHA_INFORME'].isna().all():
    st.error("❌ FECHA_INFORME no tiene fechas válidas (NaT). Revisa el Excel.")
    st.stop()

# Columnas opcionales (si existen, las usamos)
col_lub = 'PRODUCTO' if 'PRODUCTO' in df.columns else None
col_op = 'NOMBRE_OPERACION' if 'NOMBRE_OPERACION' in df.columns else None
col_tipo = 'TIPO_EQUIPO' if 'TIPO_EQUIPO' in df.columns else None

# =========================
# Detectar variables "lógicas" (parejas: base + 'base - Estado')
# =========================
def normalize_state_col_name(colname: str) -> str:
    # normaliza espacios: " - Estado ", " - Estado", etc.
    s = str(colname).strip()
    s = s.replace(" - Estado ", " - Estado")
    return s

df.columns = [normalize_state_col_name(c) for c in df.columns]

# Encuentra columnas de estado y arma pares
state_cols = [c for c in df.columns if str(c).endswith(" - Estado")]
pairs = []
for sc in state_cols:
    base = str(sc).replace(" - Estado", "").strip()
    if base in df.columns:
        pairs.append((base, sc))

# Si no hay pares, no se puede continuar con tu lógica solicitada
if not pairs:
    st.error("❌ No encontré variables en pareja tipo 'Variable' y 'Variable - Estado'. Revisa el formato del Excel.")
    st.stop()

# Lista de variables base disponibles (solo las “lógicas”)
bases = sorted({b for b, _ in pairs})

# =========================
# Filtros
# =========================
st.markdown("### 1) Filtros")

df_f = df.copy()

c1, c2, c3 = st.columns(3)

with c1:
    if col_lub:
        lubs = st.multiselect("Filtrar por Lubricante (PRODUCTO)", df_f[col_lub].dropna().unique())
        if lubs:
            df_f = df_f[df_f[col_lub].isin(lubs)].copy()
    else:
        st.info("No existe columna PRODUCTO. (Filtro por lubricante no disponible)")

with c2:
    if col_op:
        ops = st.multiselect("Filtrar por Operación (NOMBRE_OPERACION)", df_f[col_op].dropna().unique())
        if ops:
            df_f = df_f[df_f[col_op].isin(ops)].copy()

with c3:
    if col_tipo:
        tipos = st.multiselect("Filtrar por Tipo de equipo (TIPO_EQUIPO)", df_f[col_tipo].dropna().unique())
        if tipos:
            df_f = df_f[df_f[col_tipo].isin(tipos)].copy()

if df_f.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()

# =========================
# Selección de equipos
# =========================
st.markdown("### 2) Selección de equipo(s)")

equipos_disponibles = sorted(df_f['EQUIPO'].dropna().astype(str).unique())
equipos_sel = st.multiselect("Equipo(s) a calcular", equipos_disponibles, default=equipos_disponibles[:1] if equipos_disponibles else [])
if not equipos_sel:
    st.warning("Selecciona al menos un equipo.")
    st.stop()

df_f = df_f[df_f['EQUIPO'].astype(str).isin(set(map(str, equipos_sel)))].copy()
if df_f.empty:
    st.warning("No hay registros para los equipos seleccionados con los filtros actuales.")
    st.stop()

# =========================
# Selección de variables lógicas
# =========================
st.markdown("### 3) Selección de variables (solo variables en pareja con '- Estado')")

vars_sel = st.multiselect(
    "Selecciona variables a calcular",
    options=bases,
    default=bases[:1] if bases else []
)
if not vars_sel:
    st.warning("Selecciona al menos una variable.")
    st.stop()

# Mapa base -> estado_col
base_to_state = {b: sc for b, sc in pairs if b in vars_sel}

# =========================
# Opción: excluir ALERTA (u otros estados) del cálculo
# =========================
st.markdown("### 4) Tratamiento de valores atípicos por estado (opcional)")

# Estados presentes (en las columnas de estado seleccionadas)
def unique_states_for_selected(df_in: pd.DataFrame, state_columns: list[str]) -> list[str]:
    vals = set()
    for c in state_columns:
        if c in df_in.columns:
            s = df_in[c].dropna().astype(str).str.strip().str.upper().unique()
            vals.update(s)
    vals.discard("")
    return sorted(vals)

all_states = unique_states_for_selected(df_f, list(base_to_state.values()))
if not all_states:
    all_states = ["NORMAL", "PRECAUCION", "ALERTA"]  # fallback

exclude_by_state = st.checkbox(
    "Excluir del cálculo registros según el estado de la variable (recomendado si quieres ignorar ALERTAS)",
    value=True
)

exclude_states = []
if exclude_by_state:
    # por defecto, excluir ALERTA
    default_excl = ["ALERTA"] if "ALERTA" in all_states else []
    exclude_states = st.multiselect(
        "Selecciona los estados que NO se tendrán en cuenta para el cálculo",
        options=all_states,
        default=default_excl
    )

# =========================
# Parámetros de cálculo (dos métodos por n)
# =========================
st.markdown("### 5) Parámetros para cálculo de límites")

min_n = st.number_input("Mínimo de datos válidos por equipo (por variable) para calcular", min_value=2, value=3, step=1)

colA, colB = st.columns(2)
with colA:
    n_switch = st.number_input("Umbral n para usar percentiles (si n ≥ umbral)", min_value=3, value=10, step=1)
    p_prec = st.slider("Percentil para Precaución (n ≥ umbral)", min_value=50, max_value=99, value=90, step=1)
    p_cond = st.slider("Percentil para Condenatorio (n ≥ umbral)", min_value=50, max_value=99, value=95, step=1)

with colB:
    k_prec = st.number_input("k Precaución si n < umbral: μ + k·σ", min_value=0.0, value=2.0, step=0.5)
    k_cond = st.number_input("k Condenatorio si n < umbral: μ + k·σ", min_value=0.0, value=3.0, step_


