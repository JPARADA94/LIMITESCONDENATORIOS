# limites_condenatorios.py
# Autor: Javier Parada
# Entrada: Excel TAL CUAL viene de SmartAssintence
# Objetivo: Calcular límites de Precaución y Condenatorio por equipo usando SOLO variables “lógicas”
#           (variables que tienen su pareja "<Variable>" y "<Variable> - Estado").
#
# Cambios solicitados:
# 1) Filtro por nombre del lubricante (PRODUCTO si existe).
# 2) Variables: SOLO las que están en pareja (ej: "Fe (Iron)" y "Fe (Iron) - Estado").
#    El usuario escoge cuáles variables usar.
# 3) Filtro de atípicos: opción para excluir del cálculo los registros donde la variable esté en ALERTA
#    según su columna "<Variable> - Estado".

import streamlit as st
import pandas as pd
import numpy as np

# =========================
# Configuración de la app
# =========================
st.set_page_config(page_title="Límites Condenatorios - por Equipo", layout="wide")
st.title("📏 Límites Condenatorios por Equipo (Histórico)")

st.markdown("""
### Cómo funciona (paso a paso)

1) **Cargas el Excel** (tal cual viene de SmartAssintence).  
2) La app detecta las **variables “lógicas”**, es decir, aquellas que vienen en pareja:  
   **`<Variable>`** y **`<Variable> - Estado`**.  
   - Ejemplo: **HIERRO** y **HIERRO - Estado** → la variable usable es **HIERRO**.  
   - La columna **`- Estado`** se usa únicamente para saber si ese resultado está en **NORMAL / PRECAUCIÓN / ALERTA**.  
3) Aplicamos filtros (Operación / Tipo de equipo / **Lubricante** / Fechas) y luego eliges el/los **equipos**.  
4) Para cada equipo y cada variable seleccionada, calculamos límites con dos métodos (según el histórico disponible):  
   - **Si n ≥ umbral**: **Percentiles** (Precaución=P90, Condenatorio=P95 por defecto).  
   - **Si n < umbral**: **Media + k·Desviación** (Precaución=μ+2σ, Condenatorio=μ+3σ por defecto).  
5) **Opcional (atípicos):** si activas “Excluir ALERTAS”, los registros donde **`<Variable> - Estado` = ALERTA**
   **NO se tienen en cuenta** para el cálculo de esa variable (por equipo).  
6) Descargas el CSV para Power BI / reportes.
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

# Columnas opcionales que usaremos si existen
col_op = 'NOMBRE_OPERACION' if 'NOMBRE_OPERACION' in df.columns else None
col_tipo = 'TIPO_EQUIPO' if 'TIPO_EQUIPO' in df.columns else None
col_lub = 'PRODUCTO' if 'PRODUCTO' in df.columns else ('Tested Lubricant' if 'Tested Lubricant' in df.columns else None)

# --------------------
# Detectar variables “lógicas” (pares: Var y Var - Estado)
# --------------------
def normalize_estado_colname(s: str) -> str:
    # Maneja casos con espacios extra: " - Estado " vs " - Estado"
    return str(s).replace(" - Estado ", " - Estado").strip()

cols = [str(c) for c in df.columns]
estado_cols = [c for c in cols if " - Estado" in c]  # detecta todas las columnas estado

logic_vars = []
var_to_estado = {}

for c_estado in estado_cols:
    c_estado_norm = normalize_estado_colname(c_estado)
    base = c_estado_norm.replace(" - Estado", "").strip()
    # Solo si la base existe como columna
    if base in df.columns:
        logic_vars.append(base)
        var_to_estado[base] = c_estado  # guardamos el nombre real como viene (para indexar)

# Quitar duplicados conservando orden
seen = set()
logic_vars = [v for v in logic_vars if not (v in seen or seen.add(v))]

if not logic_vars:
    st.error("❌ No se detectaron variables lógicas (pares '<Variable>' y '<Variable> - Estado').")
    st.stop()

# --------------------
# Filtros (incluye lubricante)
# --------------------
st.markdown("### 1) Filtros")

df_f = df.copy()

if col_op:
    ops = st.multiselect(f"Filtrar por operación ({col_op})", sorted(df_f[col_op].dropna().unique()))
    if ops:
        df_f = df_f[df_f[col_op].isin(ops)].copy()

if col_tipo:
    tipos = st.multiselect(f"Filtrar por tipo de equipo ({col_tipo})", sorted(df_f[col_tipo].dropna().unique()))
    if tipos:
        df_f = df_f[df_f[col_tipo].isin(tipos)].copy()

# ✅ Filtro por lubricante
if col_lub:
    lubs = st.multiselect(f"Filtrar por lubricante ({col_lub})", sorted(df_f[col_lub].dropna().unique()))
    if lubs:
        df_f = df_f[df_f[col_lub].isin(lubs)].copy()
else:
    st.info("No encontré columna de lubricante (PRODUCTO / Tested Lubricant). Se omite ese filtro.")

# Filtro por fechas (opcional)
use_dates = st.checkbox("Filtrar por rango de fechas", value=False)
if use_dates:
    min_d = df_f['FECHA_INFORME'].min()
    max_d = df_f['FECHA_INFORME'].max()
    if pd.isna(min_d) or pd.isna(max_d):
        st.warning("No se puede filtrar por fechas porque FECHA_INFORME tiene NaT.")
    else:
        d1, d2 = st.date_input(
            "Rango de fechas",
            value=[min_d.date(), max_d.date()],
            min_value=min_d.date(),
            max_value=max_d.date()
        )
        df_f = df_f[(df_f['FECHA_INFORME'] >= pd.to_datetime(d1)) &
                    (df_f['FECHA_INFORME'] <= pd.to_datetime(d2))].copy()

if df_f.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()

# --------------------
# Selección de variables lógicas
# --------------------
st.markdown("### 2) Variables (solo variables lógicas con columna `- Estado`)")

vars_sel = st.multiselect(
    "Selecciona las variables para calcular límites",
    options=logic_vars,
    default=logic_vars[:1] if len(logic_vars) > 0 else []
)

if not vars_sel:
    st.warning("Selecciona al menos una variable.")
    st.stop()

# Convertir variables seleccionadas a numéricas (solo en memoria)
for v in vars_sel:
    df_f[v] = pd.to_numeric(
        df_f[v].astype(str).str.replace(r"[^0-9\.\-]", "", regex=True),
        errors="coerce"
    )

# --------------------
# Inventario de equipos (con n válidas)
# --------------------
st.markdown("### 3) Inventario de equipos (n de datos válidos por variable)")

group_cols = ['EQUIPO']
if col_op: group_cols.append(col_op)
if col_tipo: group_cols.append(col_tipo)
if col_lub: group_cols.append(col_lub)

# inventario básico
inv = (
    df_f.groupby(group_cols)
        .agg(
            muestras_totales=('EQUIPO', 'size'),
            fecha_min=('FECHA_INFORME', 'min'),
            fecha_max=('FECHA_INFORME', 'max'),
            ultima_muestra=('FECHA_INFORME', 'max')
        )
        .reset_index()
)

# agregamos columnas n_validas_<var> por cada variable seleccionada (para que el usuario vea el histórico)
for v in vars_sel:
    inv[f"n_validas_{v}"] = (
        df_f.groupby(group_cols)[v]
            .apply(lambda s: int(s.notna().sum()))
            .values
    )

# Orden: mayor histórico primero (sumatoria de n_validas)
inv["_score"] = 0
for v in vars_sel:
    inv["_score"] += inv[f"n_validas_{v}"]
inv = inv.sort_values(["_score", "muestras_totales"], ascending=False).drop(columns=["_score"])

st.dataframe(inv, use_container_width=True)

# --------------------
# Selección de equipo(s)
# --------------------
st.markdown("### 4) Selecciona el/los equipo(s)")

equipos_disponibles = sorted(df_f['EQUIPO'].dropna().astype(str).unique())
equipos_sel = st.multiselect(
    "Equipo(s)",
    options=equipos_disponibles,
    default=equipos_disponibles[:1] if len(equipos_disponibles) > 0 else []
)

if not equipos_sel:
    st.warning("Selecciona al menos un equipo.")
    st.stop()

df_calc = df_f[df_f['EQUIPO'].astype(str).isin(set(map(str, equipos_sel)))].copy()
if df_calc.empty:
    st.warning("No hay registros para los equipos seleccionados con los filtros actuales.")
    st.stop()

st.markdown("---")

# --------------------
# Parámetros de cálculo + atípicos
# --------------------
st.markdown("### 5) Parámetros de cálculo")

min_n = st.number_input("Mínimo de datos válidos por equipo (por variable) para calcular límites", min_value=2, value=3, step=1)

c1, c2 = st.columns(2)
with c1:
    n_switch = st.number_input("Umbral n para usar percentiles (si n ≥ umbral)", min_value=3, value=10, step=1)
    p_prec = st.slider("Percentil Precaución (n ≥ umbral)", 50, 99, 90, 1)
    p_cond = st.slider("Percentil Condenatorio (n ≥ umbral)", 50, 99, 95, 1)
with c2:
    k_prec = st.number_input("k Precaución (n < umbral): μ + k·σ", min_value=0.0, value=2.0, step=0.5)
    k_cond = st.number_input("k Condenatorio (n < umbral): μ + k·σ", min_value=0.0, value=3.0, step=0.5)

# ✅ Atípicos por estado
st.markdown("#### Filtro de atípicos por estado (según `<Variable> - Estado`)")

excluir_alertas = st.checkbox(
    "Excluir del cálculo los registros en ALERTA (por variable)",
    value=True
)
# Si más adelante quieres también excluir PRECAUCION, aquí está listo:
excluir_precaucion = st.checkbox(
    "Excluir también los registros en PRECAUCIÓN (opcional)",
    value=False
)

limpieza_iqr = st.checkbox("Aplicar limpieza adicional de outliers (IQR) después del filtro por estado", value=False)

def clean_outliers_iqr(x: pd.Series) -> pd.Series:
    x = x.dropna()
    if len(x) < 4:
        return x
    q1, q3 = x.quantile(0.25), x.quantile(0.75)
    iqr = q3 - q1
    lo = q1 - 1.5 * iqr
    hi = q3 + 1.5 * iqr
    return x[(x >= lo) & (x <= hi)]

def apply_estado_filter(g: pd.DataFrame, v: str) -> pd.Series:
    """
    Devuelve serie numérica de v, filtrada opcionalmente por estado:
    - excluye ALERTA si el usuario lo activa
    - excluye PRECAUCION si el usuario lo activa
    """
    s = g[v]
    estado_col = var_to_estado.get(v)
    if not estado_col:
        return s  # si no existe columna estado, no filtramos por estado

    est = g[estado_col].astype(str).str.strip().str.upper()

    mask = pd.Series(True, index=g.index)
    if excluir_alertas:
        mask &= (est != "ALERTA")
    if excluir_precaucion:
        mask &= (est != "PRECAUCION")

    return s[mask]

def compute_limits(series: pd.Series) -> dict:
    s = series.dropna()
    n = int(len(s))

    if n < min_n:
        return {
            'n': n, 'metodo': 'Insuficiente',
            'precaucion': np.nan, 'condenatorio': np.nan,
            'mean': np.nan, 'std': np.nan, 'median': np.nan,
            'min': np.nan, 'max': np.nan
        }

    if limpieza_iqr:
        s2 = clean_outliers_iqr(s)
        n2 = int(len(s2))
        if n2 < min_n:
            return {
                'n': n2, 'metodo': 'Insuficiente (post-IQR)',
                'precaucion': np.nan, 'condenatorio': np.nan,
                'mean': np.nan, 'std': np.nan, 'median': np.nan,
                'min': np.nan, 'max': np.nan
            }
        s = s2
        n = n2

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

# --------------------
# Cálculo: por equipo x variable
# --------------------
st.markdown("### 6) Resultados (límites por equipo y variable)")

rows = []
for equipo, g_eq in df_calc.groupby(df_calc['EQUIPO'].astype(str)):
    for v in vars_sel:
        # aplicar filtro por estado (ALERTA / PRECAUCION) si corresponde
        serie_filtrada = apply_estado_filter(g_eq, v)

        out = compute_limits(serie_filtrada)

        row = {
            "EQUIPO": equipo,
            "VARIABLE": v,
            "n": out["n"],
            "metodo": out["metodo"],
            "precaucion": out["precaucion"],
            "condenatorio": out["condenatorio"],
            "mean": out["mean"],
            "std": out["std"],
            "median": out["median"],
            "min": out["min"],
            "max": out["max"],
            "fecha_min": g_eq['FECHA_INFORME'].min(),
            "fecha_max": g_eq['FECHA_INFORME'].max(),
        }

        # columnas descriptivas (si existen)
        if col_op:
            row[col_op] = g_eq[col_op].dropna().iloc[0] if g_eq[col_op].notna().any() else None
        if col_tipo:
            row[col_tipo] = g_eq[col_tipo].dropna().iloc[0] if g_eq[col_tipo].notna().any() else None
        if col_lub:
            row[col_lub] = g_eq[col_lub].dropna().iloc[0] if g_eq[col_lub].notna().any() else None

        # info del filtro por estado aplicado
        estado_col = var_to_estado.get(v)
        row["estado_col_usada"] = estado_col if estado_col else ""
        row["excluye_ALERTA"] = "SI" if excluir_alertas else "NO"
        row["excluye_PRECAUCION"] = "SI" if excluir_precaucion else "NO"
        row["iqr"] = "SI" if limpieza_iqr else "NO"

        rows.append(row)

res = pd.DataFrame(rows)

# Redondeo visual (no altera cálculos internos)
dec = st.number_input("Decimales para mostrar (solo visual)", min_value=0, value=0, step=1)
res_show = res.copy()
for c in ["precaucion", "condenatorio", "mean", "std", "median", "min", "max"]:
    res_show[c] = res_show[c].round(dec)

# Orden sugerido
cols_order = []
for c in ["EQUIPO", col_op, col_tipo, col_lub, "VARIABLE", "n", "metodo",
          "precaucion", "condenatorio", "mean", "std", "median", "min", "max",
          "fecha_min", "fecha_max",
          "excluye_ALERTA", "excluye_PRECAUCION", "iqr"]:
    if c and c in res_show.columns:
        cols_order.append(c)

st.dataframe(
    res_show[cols_order].sort_values(["EQUIPO", "VARIABLE"]),
    use_container_width=True
)

# --------------------
# Exportación
# --------------------
st.markdown("### 7) Exportar resultados")

csv = res.to_csv(index=False).encode("utf-8")
st.download_button(
    "⬇️ Descargar CSV (límites por equipo y variable)",
    data=csv,
    file_name="limites_condenatorios_por_equipo.csv",
    mime="text/csv"
)


