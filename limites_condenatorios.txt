# limites_condenatorios.py
# Autor: Javier Parada
# Entrada: Excel TAL CUAL viene de SmartAssintence (sin modificar el archivo)
# Objetivo: Calcular límites de Precaución y Condenatorio por equipo, usando histórico por variable.
#
# Notas:
# - La app NO cambia tu Excel; solo lo lee y calcula.
# - Te deja seleccionar 1 o varios equipos para calcular límites específicamente para ellos.
# - Métodos automáticos según n (cantidad de análisis válidos por equipo):
#     * n ≥ umbral -> Percentiles (Prec=P90, Cond=P95 por defecto)
#     * n < umbral -> Media + k·Desv (Prec=μ+2σ, Cond=μ+3σ por defecto)
# - Opción de limpieza de outliers (IQR) antes del cálculo (recomendado para PQ).

import streamlit as st
import pandas as pd
import numpy as np

# =========================
# Configuración de la app
# =========================
st.set_page_config(page_title="Límites Condenatorios - por Equipo", layout="wide")
st.title("📏 Límites Condenatorios por Equipo (Histórico)")

st.markdown("""
### ¿Qué hace esta aplicación y cómo calcula los límites?

Esta herramienta calcula **límites de Precaución y Condenatorio para un equipo específico** usando el histórico del archivo **tal cual** viene de SmartAssintence. El flujo es:

1) **Cargas el Excel** (sin modificarlo). La app valida que existan columnas mínimas como *EQUIPO* y *FECHA_INFORME*.  
2) **Seleccionas la variable** a evaluar (por ejemplo: Fe, Cu, PQ, Visc@40, Agua, etc.). La app convierte esa columna a numérica para analizarla bien.  
3) La app te muestra un **inventario de equipos** con cuántos datos válidos tiene cada uno y el rango de fechas.  
4) **Seleccionas el/los equipos** a los que quieres calcular límites.  
5) Para cada equipo, según el tamaño del histórico (**n** datos válidos), aplica automáticamente:
   - **Si n ≥ umbral (por defecto 10): Percentiles** → Precaución = **P90** y Condenatorio = **P95** (configurable).
   - **Si n < umbral: Media + k·Desviación** → Precaución = **μ + k₁·σ** y Condenatorio = **μ + k₂·σ** (configurable).  
6) (Opcional) Puedes activar **limpieza de outliers (IQR)** antes del cálculo para evitar que valores extremos inflen los límites.  
7) Obtienes una tabla final con límites por equipo y puedes **exportarla a CSV** para Power BI o reportes.

> Si un equipo no cumple el mínimo de datos válidos configurado, se marca como **Insuficiente** y no se calculan límites para evitar conclusiones débiles.
""")

# --------------------
# Carga (tal cual)
# --------------------
@st.cache_data(show_spinner=False)
def load_excel(file) -> pd.DataFrame:
    return pd.read_excel(file)

archivo = st.file_uploader("📁 Sube tu Excel (.xlsx) de SmartAssintence (tal cual)", type=["xlsx"])
if not archivo:
    st.stop()

df = load_excel(archivo).copy()

# --------------------
# Validaciones mínimas (sin tocar el archivo)
# --------------------
required_min = ['EQUIPO', 'FECHA_INFORME']
missing = [c for c in required_min if c not in df.columns]
if missing:
    st.error(f"❌ Falta(n) columna(s) mínima(s) requerida(s): {missing}")
    st.stop()

df['FECHA_INFORME'] = pd.to_datetime(df['FECHA_INFORME'], errors='coerce')
if df['FECHA_INFORME'].isna().all():
    st.error("❌ FECHA_INFORME no tiene fechas válidas (NaT). Revisa el Excel.")
    st.stop()

# Columnas opcionales (si existen, las usamos para enriquecer inventario/filtros)
col_op = 'NOMBRE_OPERACION' if 'NOMBRE_OPERACION' in df.columns else None
col_tipo = 'TIPO_EQUIPO' if 'TIPO_EQUIPO' in df.columns else None

# --------------------
# Selección de variable
# --------------------
st.markdown("### 1) Selecciona la variable a analizar")

exclude = set(['EQUIPO', 'FECHA_INFORME'])
if col_op: exclude.add(col_op)
if col_tipo: exclude.add(col_tipo)
# también excluir columnas de estado si existen (no son la variable numérica)
# (no afecta si no existen)
exclude |= {c for c in df.columns if ' - Estado' in str(c)}

candidates = []
for c in df.columns:
    if c in exclude:
        continue

    if pd.api.types.is_numeric_dtype(df[c]):
        candidates.append(c)
    else:
        sample = df[c].dropna().astype(str).head(60)
        if not sample.empty and (sample.str.contains(r"\d", regex=True).mean() >= 0.6):
            candidates.append(c)

# quitar duplicados conservando orden
seen = set()
candidates = [x for x in candidates if not (x in seen or seen.add(x))]

if not candidates:
    st.warning("No se encontraron columnas numéricas/convertibles para calcular límites.")
    st.stop()

var = st.selectbox("Variable", candidates)

# Convertir la variable a numérica (sin cambiar el excel, solo el dataframe en memoria)
df[var] = pd.to_numeric(
    df[var].astype(str).str.replace(r"[^0-9\.\-]", "", regex=True),
    errors='coerce'
)

# --------------------
# Filtros opcionales (si existen columnas)
# --------------------
st.markdown("### 2) Filtros (opcional)")

df_f = df.copy()

if col_op:
    ops = st.multiselect(f"Filtrar por operación ({col_op})", df_f[col_op].dropna().unique())
    if ops:
        df_f = df_f[df_f[col_op].isin(ops)].copy()

if col_tipo:
    tipos = st.multiselect(f"Filtrar por tipo de equipo ({col_tipo})", df_f[col_tipo].dropna().unique())
    if tipos:
        df_f = df_f[df_f[col_tipo].isin(tipos)].copy()

if df_f.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()

# --------------------
# Inventario de equipos
# --------------------
st.markdown("### 3) Inventario de equipos (histórico disponible)")

group_cols = ['EQUIPO']
if col_op: group_cols.append(col_op)
if col_tipo: group_cols.append(col_tipo)

inv = (
    df_f.groupby(group_cols)
        .agg(
            muestras_totales=('EQUIPO', 'size'),
            n_validas=(var, lambda s: s.notna().sum()),
            fecha_min=('FECHA_INFORME', 'min'),
            fecha_max=('FECHA_INFORME', 'max'),
            ultima_muestra=('FECHA_INFORME', 'max')
        )
        .reset_index()
        .sort_values(['n_validas', 'muestras_totales'], ascending=False)
)

st.dataframe(inv, use_container_width=True)

# --------------------
# Selección de equipo(s)
# --------------------
st.markdown("### 4) Selecciona el/los equipo(s) a calcular")

equipos_disponibles = inv['EQUIPO'].dropna().astype(str).unique()
equipos_sel = st.multiselect(
    "Equipo(s)",
    options=equipos_disponibles,
    default=list(equipos_disponibles[:1]) if len(equipos_disponibles) > 0 else []
)

if not equipos_sel:
    st.warning("Selecciona al menos un equipo para calcular límites.")
    st.stop()

df_calc = df_f[df_f['EQUIPO'].astype(str).isin(set(map(str, equipos_sel)))].copy()
if df_calc.empty:
    st.warning("No hay registros para los equipos seleccionados (con los filtros actuales).")
    st.stop()

st.markdown("---")

# --------------------
# Parámetros de cálculo
# --------------------
st.markdown("### 5) Parámetros para cálculo de límites")

min_n = st.number_input("Mínimo de datos válidos por equipo para calcular límites", min_value=2, value=3, step=1)

colA, colB = st.columns(2)
with colA:
    n_switch = st.number_input("Umbral n para usar percentiles (si n ≥ umbral)", min_value=3, value=10, step=1)
    p_prec = st.slider("Percentil para Precaución (n ≥ umbral)", min_value=50, max_value=99, value=90, step=1)
    p_cond = st.slider("Percentil para Condenatorio (n ≥ umbral)", min_value=50, max_value=99, value=95, step=1)

with colB:
    k_prec = st.number_input("k para Precaución (n < umbral): μ + k·σ", min_value=0.0, value=2.0, step=0.5)
    k_cond = st.number_input("k para Condenatorio (n < umbral): μ + k·σ", min_value=0.0, value=3.0, step=0.5)

limpieza = st.checkbox("Aplicar limpieza de outliers antes de calcular (IQR) — recomendado para PQ", value=False)

def clean_outliers_iqr(x: pd.Series) -> pd.Series:
    x = x.dropna()
    if len(x) < 4:
        return x
    q1, q3 = x.quantile(0.25), x.quantile(0.75)
    iqr = q3 - q1
    lo = q1 - 1.5 * iqr
    hi = q3 + 1.5 * iqr
    return x[(x >= lo) & (x <= hi)]

# --------------------
# Cálculo por equipo
# --------------------
st.markdown("### 6) Resultados: límites por equipo")

def compute_limits(series: pd.Series) -> dict:
    s = series.dropna()
    n = len(s)

    if n < min_n:
        return {
            'n': n, 'metodo': 'Insuficiente',
            'precaucion': np.nan, 'condenatorio': np.nan,
            'mean': np.nan, 'std': np.nan, 'median': np.nan,
            'min': np.nan, 'max': np.nan
        }

    if limpieza:
        s = clean_outliers_iqr(s)
        n = len(s)
        if n < min_n:
            return {
                'n': n, 'metodo': 'Insuficiente (post-limpieza)',
                'precaucion': np.nan, 'condenatorio': np.nan,
                'mean': np.nan, 'std': np.nan, 'median': np.nan,
                'min': np.nan, 'max': np.nan
            }

    mean = float(s.mean())
    std = float(s.std(ddof=1)) if n > 1 else 0.0
    median = float(s.median())
    vmin = float(s.min())
    vmax = float(s.max())

    if n >= n_switch:
        prec = float(s.quantile(p_prec / 100))
        cond = float(s.quantile(p_cond / 100))
        metodo = f"Percentiles P{p_prec}/P{p_cond}"
    else:
        prec = mean + (k_prec * std)
        cond = mean + (k_cond * std)
        metodo = f"Media+Desv (k={k_prec}/{k_cond})"

    return {
        'n': n, 'metodo': metodo,
        'precaucion': prec, 'condenatorio': cond,
        'mean': mean, 'std': std, 'median': median,
        'min': vmin, 'max': vmax
    }

rows = []
for equipo, g in df_calc.groupby(df_calc['EQUIPO'].astype(str)):
    out = compute_limits(g[var])
    row = {
        'EQUIPO': equipo,
        'VARIABLE': var,
        'fecha_min': g['FECHA_INFORME'].min(),
        'fecha_max': g['FECHA_INFORME'].max(),
        **out
    }
    # agrega columnas descriptivas si existen
    if col_op:
        row[col_op] = g[col_op].dropna().iloc[0] if g[col_op].notna().any() else None
    if col_tipo:
        row[col_tipo] = g[col_tipo].dropna().iloc[0] if g[col_tipo].notna().any() else None

    rows.append(row)

res = pd.DataFrame(rows)

# Visualización
dec = st.number_input("Decimales para mostrar (solo visual)", min_value=0, value=0, step=1)
res_show = res.copy()
for c in ['precaucion', 'condenatorio', 'mean', 'std', 'median', 'min', 'max']:
    if c in res_show.columns:
        res_show[c] = res_show[c].round(dec)

# Orden sugerido de columnas
cols_order = []
for c in ['EQUIPO', col_op, col_tipo, 'VARIABLE', 'n', 'metodo', 'precaucion', 'condenatorio', 'mean', 'std', 'median', 'min', 'max', 'fecha_min', 'fecha_max']:
    if c and c in res_show.columns:
        cols_order.append(c)

st.dataframe(res_show[cols_order].sort_values(['n', 'EQUIPO'], ascending=[False, True]), use_container_width=True)

# --------------------
# Exportación
# --------------------
st.markdown("### 7) Exportar resultados")

csv = res.to_csv(index=False).encode("utf-8")
st.download_button(
    "⬇️ Descargar CSV (límites por equipo seleccionado)",
    data=csv,
    file_name=f"limites_{var.replace(' ', '_')}_equipos.csv",
    mime="text/csv"
)
