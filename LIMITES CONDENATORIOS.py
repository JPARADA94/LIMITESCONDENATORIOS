# limites_condenatorios.py
# Autor: Javier Parada
# Entrada: Excel TAL CUAL viene de SmartAssintence (no se modifica el archivo)
#
# Cambios solicitados:
# ✅ 1) Los 3 filtros (Operación / Tipo equipo / Lubricante) son OPCIONALES (se activan con checkbox).
# ✅ 2) Selección de variables por “chulos” (checkboxes), no por lista.
# ✅ 3) Orden: FILTROS -> INVENTARIO -> seleccionar EQUIPOS -> seleccionar VARIABLES -> calcular.
# ✅ 4) Descargar resultado en EXCEL (.xlsx).
# ✅ 5) Inventario siempre visible (si hay datos).

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# =========================
# Configuración
# =========================
st.set_page_config(page_title="Límites Condenatorios - por Equipo", layout="wide")
st.title("📏 Límites Condenatorios por Equipo (Histórico)")

st.markdown("""
Esta app calcula **límites de Precaución y Condenatorio por equipo**, usando únicamente **variables “lógicas”** (las que vienen en pareja:
`<Variable>` y `<Variable> - Estado`).  
Opcionalmente puedes filtrar por **Operación**, **Tipo de equipo** y **Lubricante**.  
También puedes decidir si deseas **excluir del cálculo** los registros donde la variable esté en **ALERTA** (según su columna `- Estado`).
""")

# --------------------
# Utilidades
# --------------------
@st.cache_data(show_spinner=False)
def load_excel(file) -> pd.DataFrame:
    return pd.read_excel(file)

def normalize_estado_colname(s: str) -> str:
    return str(s).replace(" - Estado ", " - Estado").strip()

def clean_outliers_iqr(x: pd.Series) -> pd.Series:
    x = x.dropna()
    if len(x) < 4:
        return x
    q1, q3 = x.quantile(0.25), x.quantile(0.75)
    iqr = q3 - q1
    lo = q1 - 1.5 * iqr
    hi = q3 + 1.5 * iqr
    return x[(x >= lo) & (x <= hi)]

def to_excel_bytes(df_export: pd.DataFrame, sheet_name: str = "Limites") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

# --------------------
# Carga
# --------------------
archivo = st.file_uploader("📁 Sube tu Excel (.xlsx) de SmartAssintence (tal cual)", type=["xlsx"])
if not archivo:
    st.stop()

df = load_excel(archivo).copy()

# Validaciones mínimas
required_min = ["EQUIPO", "FECHA_INFORME"]
missing = [c for c in required_min if c not in df.columns]
if missing:
    st.error(f"❌ Falta(n) columna(s) mínima(s) requerida(s): {missing}")
    st.stop()

df["FECHA_INFORME"] = pd.to_datetime(df["FECHA_INFORME"], errors="coerce")
if df["FECHA_INFORME"].isna().all():
    st.error("❌ FECHA_INFORME no tiene fechas válidas (NaT). Revisa el Excel.")
    st.stop()

# Columnas opcionales (si existen)
col_op = "NOMBRE_OPERACION" if "NOMBRE_OPERACION" in df.columns else None
col_tipo = "TIPO_EQUIPO" if "TIPO_EQUIPO" in df.columns else None
col_lub = "PRODUCTO" if "PRODUCTO" in df.columns else ("Tested Lubricant" if "Tested Lubricant" in df.columns else None)

# --------------------
# Detectar variables lógicas (pares: Var y Var - Estado)
# --------------------
cols = [str(c) for c in df.columns]
estado_cols = [c for c in cols if " - Estado" in c]

logic_vars = []
var_to_estado = {}
for c_estado in estado_cols:
    c_norm = normalize_estado_colname(c_estado)
    base = c_norm.replace(" - Estado", "").strip()
    if base in df.columns:
        logic_vars.append(base)
        var_to_estado[base] = c_estado  # nombre real

# dedupe conservando orden
seen = set()
logic_vars = [v for v in logic_vars if not (v in seen or seen.add(v))]

if not logic_vars:
    st.error("❌ No se detectaron variables lógicas. Debe existir '<Variable>' y '<Variable> - Estado'.")
    st.stop()

# =========================
# 1) FILTROS (opcionales)
# =========================
st.markdown("## 1) Filtros (opcionales)")

df_f = df.copy()

# Filtro fechas (opcional)
use_dates = st.checkbox("Filtrar por rango de fechas", value=False)
if use_dates:
    min_d = df_f["FECHA_INFORME"].min()
    max_d = df_f["FECHA_INFORME"].max()
    if pd.isna(min_d) or pd.isna(max_d):
        st.warning("No se puede filtrar por fechas porque FECHA_INFORME tiene NaT.")
    else:
        d1, d2 = st.date_input(
            "Rango de fechas",
            value=[min_d.date(), max_d.date()],
            min_value=min_d.date(),
            max_value=max_d.date()
        )
        df_f = df_f[
            (df_f["FECHA_INFORME"] >= pd.to_datetime(d1)) &
            (df_f["FECHA_INFORME"] <= pd.to_datetime(d2))
        ].copy()

cA, cB, cC = st.columns(3)

with cA:
    use_op = st.checkbox("Filtrar por Operación", value=False, disabled=(col_op is None))
    if col_op is None:
        st.caption("No existe columna NOMBRE_OPERACION en el Excel.")
    ops_sel = []
    if use_op and col_op:
        ops_sel = st.multiselect("Operación(es)", sorted(df_f[col_op].dropna().unique()))
        if ops_sel:
            df_f = df_f[df_f[col_op].isin(ops_sel)].copy()

with cB:
    use_tipo = st.checkbox("Filtrar por Tipo de equipo", value=False, disabled=(col_tipo is None))
    if col_tipo is None:
        st.caption("No existe columna TIPO_EQUIPO en el Excel.")
    tipos_sel = []
    if use_tipo and col_tipo:
        tipos_sel = st.multiselect("Tipo(s) de equipo", sorted(df_f[col_tipo].dropna().unique()))
        if tipos_sel:
            df_f = df_f[df_f[col_tipo].isin(tipos_sel)].copy()

with cC:
    use_lub = st.checkbox("Filtrar por Lubricante", value=False, disabled=(col_lub is None))
    if col_lub is None:
        st.caption("No existe columna PRODUCTO / Tested Lubricant en el Excel.")
    lubs_sel = []
    if use_lub and col_lub:
        lubs_sel = st.multiselect("Lubricante(s)", sorted(df_f[col_lub].dropna().unique()))
        if lubs_sel:
            df_f = df_f[df_f[col_lub].isin(lubs_sel)].copy()

if df_f.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()

# =========================
# 2) INVENTARIO (siempre visible)
# =========================
st.markdown("## 2) Inventario de equipos (histórico disponible)")

# inventario con conteos válidos por variable lógica (para auditoría)
group_cols = ["EQUIPO"]
if col_op: group_cols.append(col_op)
if col_tipo: group_cols.append(col_tipo)
if col_lub: group_cols.append(col_lub)

inv = (
    df_f.groupby(group_cols, dropna=False)
        .agg(
            muestras_totales=("EQUIPO", "size"),
            fecha_min=("FECHA_INFORME", "min"),
            fecha_max=("FECHA_INFORME", "max"),
            ultima_muestra=("FECHA_INFORME", "max"),
        )
        .reset_index()
)

# Para no hacer el inventario enorme, mostramos n_validas solo de las primeras 10 variables por defecto
# (igual el cálculo luego se hace SOLO con las que el usuario seleccione)
max_vars_preview = 10
preview_vars = logic_vars[:max_vars_preview]

# Convertimos preview vars a numérico en memoria para conteo válido (sin cambiar archivo)
df_inv_tmp = df_f.copy()
for v in preview_vars:
    df_inv_tmp[v] = pd.to_numeric(
        df_inv_tmp[v].astype(str).str.replace(r"[^0-9\.\-]", "", regex=True),
        errors="coerce"
    )

for v in preview_vars:
    inv[f"n_validas_{v}"] = (
        df_inv_tmp.groupby(group_cols, dropna=False)[v]
                 .apply(lambda s: int(s.notna().sum()))
                 .values
    )

# Orden por historial general (muestras_totales)
inv = inv.sort_values(["muestras_totales"], ascending=False)

st.dataframe(inv, use_container_width=True)

# =========================
# 3) Selección de EQUIPOS
# =========================
st.markdown("## 3) Selecciona el/los equipo(s)")

equipos_disponibles = sorted(df_f["EQUIPO"].dropna().astype(str).unique())
equipos_sel = st.multiselect(
    "Equipo(s)",
    options=equipos_disponibles,
    default=equipos_disponibles[:1] if len(equipos_disponibles) > 0 else []
)

if not equipos_sel:
    st.warning("Selecciona al menos un equipo.")
    st.stop()

df_calc = df_f[df_f["EQUIPO"].astype(str).isin(set(map(str, equipos_sel)))].copy()
if df_calc.empty:
    st.warning("No hay registros para los equipos seleccionados con los filtros actuales.")
    st.stop()

# =========================
# 4) Selección de VARIABLES por CHULOS
# =========================
st.markdown("## 4) Selecciona variables a analizar (por chulos)")

# Buscador rápido
buscador = st.text_input("Buscar variable (opcional)", value="")
vars_filtradas = [v for v in logic_vars if buscador.strip().lower() in v.lower()]

# Botones select all / none
csel1, csel2, csel3 = st.columns([1, 1, 2])
if "vars_checked" not in st.session_state:
    st.session_state["vars_checked"] = set()

with csel1:
    if st.button("✅ Seleccionar todas (filtradas)"):
        st.session_state["vars_checked"].update(vars_filtradas)

with csel2:
    if st.button("🧹 Limpiar selección"):
        st.session_state["vars_checked"] = set()

# Checkboxes en columnas para que sea rápido
ncols = 3
cols_ui = st.columns(ncols)
for i, v in enumerate(vars_filtradas):
    col = cols_ui[i % ncols]
    with col:
        checked = v in st.session_state["vars_checked"]
        new_val = st.checkbox(v, value=checked, key=f"chk_{v}")
        if new_val:
            st.session_state["vars_checked"].add(v)
        else:
            st.session_state["vars_checked"].discard(v)

vars_sel = sorted(list(st.session_state["vars_checked"]))
if not vars_sel:
    st.warning("Selecciona al menos una variable.")
    st.stop()

st.info(f"Variables seleccionadas: {len(vars_sel)}")

# Convertir a numérico SOLO las seleccionadas (en memoria)
for v in vars_sel:
    df_calc[v] = pd.to_numeric(
        df_calc[v].astype(str).str.replace(r"[^0-9\.\-]", "", regex=True),
        errors="coerce"
    )

# =========================
# 5) Parámetros de cálculo + atípicos por estado
# =========================
st.markdown("## 5) Parámetros de cálculo")

min_n = st.number_input("Mínimo de datos válidos por equipo (por variable) para calcular límites", min_value=2, value=3, step=1)

c1, c2 = st.columns(2)
with c1:
    n_switch = st.number_input("Umbral n para usar percentiles (si n ≥ umbral)", min_value=3, value=10, step=1)
    p_prec = st.slider("Percentil Precaución (n ≥ umbral)", 50, 99, 90, 1)
    p_cond = st.slider("Percentil Condenatorio (n ≥ umbral)", 50, 99, 95, 1)
with c2:
    k_prec = st.number_input("k Precaución (n < umbral): μ + k·σ", min_value=0.0, value=2.0, step=0.5)
    k_cond = st.number_input("k Condenatorio (n < umbral): μ + k·σ", min_value=0.0, value=3.0, step=0.5)

st.markdown("### Filtro de atípicos por estado (según `<Variable> - Estado`)")

excluir_alertas = st.checkbox("Excluir del cálculo los registros en ALERTA (por variable)", value=True)
excluir_precaucion = st.checkbox("Excluir también los registros en PRECAUCIÓN (opcional)", value=False)
limpieza_iqr = st.checkbox("Aplicar limpieza adicional IQR (después del filtro por estado)", value=False)

def apply_estado_filter(g: pd.DataFrame, v: str) -> pd.Series:
    s = g[v]
    estado_col = var_to_estado.get(v)
    if not estado_col or estado_col not in g.columns:
        return s

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
            "n": n, "metodo": "Insuficiente",
            "precaucion": np.nan, "condenatorio": np.nan,
            "mean": np.nan, "std": np.nan, "median": np.nan,
            "min": np.nan, "max": np.nan
        }

    if limpieza_iqr:
        s2 = clean_outliers_iqr(s)
        n2 = int(len(s2))
        if n2 < min_n:
            return {
                "n": n2, "metodo": "Insuficiente (post-IQR)",
                "precaucion": np.nan, "condenatorio": np.nan,
                "mean": np.nan, "std": np.nan, "median": np.nan,
                "min": np.nan, "max": np.nan
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
        "n": n, "metodo": metodo,
        "precaucion": prec, "condenatorio": cond,
        "mean": mean, "std": std, "median": median,
        "min": vmin, "max": vmax
    }

# =========================
# 6) Calcular y mostrar
# =========================
st.markdown("## 6) Resultados (límites por equipo y variable)")

if st.button("🚀 Calcular límites"):
    rows = []
    for equipo, g_eq in df_calc.groupby(df_calc["EQUIPO"].astype(str)):
        for v in vars_sel:
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
                "fecha_min": g_eq["FECHA_INFORME"].min(),
                "fecha_max": g_eq["FECHA_INFORME"].max(),
                "excluye_ALERTA": "SI" if excluir_alertas else "NO",
                "excluye_PRECAUCION": "SI" if excluir_precaucion else "NO",
                "iqr": "SI" if limpieza_iqr else "NO",
            }

            # metadata (si existe)
            if col_op and col_op in g_eq.columns:
                row[col_op] = g_eq[col_op].dropna().iloc[0] if g_eq[col_op].notna().any() else None
            if col_tipo and col_tipo in g_eq.columns:
                row[col_tipo] = g_eq[col_tipo].dropna().iloc[0] if g_eq[col_tipo].notna().any() else None
            if col_lub and col_lub in g_eq.columns:
                row[col_lub] = g_eq[col_lub].dropna().iloc[0] if g_eq[col_lub].notna().any() else None

            rows.append(row)

    res = pd.DataFrame(rows)

    # Mostrar
    dec = st.number_input("Decimales para mostrar (solo visual)", min_value=0, value=0, step=1)
    res_show = res.copy()
    for c in ["precaucion", "condenatorio", "mean", "std", "median", "min", "max"]:
        res_show[c] = res_show[c].round(dec)

    # orden de columnas
    cols_order = []
    for c in ["EQUIPO", col_op, col_tipo, col_lub, "VARIABLE", "n", "metodo",
              "precaucion", "condenatorio", "mean", "std", "median", "min", "max",
              "fecha_min", "fecha_max", "excluye_ALERTA", "excluye_PRECAUCION", "iqr"]:
        if c and c in res_show.columns:
            cols_order.append(c)

    st.dataframe(res_show[cols_order].sort_values(["EQUIPO", "VARIABLE"]), use_container_width=True)

    # =========================
    # 7) Descargar en EXCEL
    # =========================
    st.markdown("## 7) Descargar resultados")
    xlsx_bytes = to_excel_bytes(res_show[cols_order], sheet_name="Limites")
    st.download_button(
        "⬇️ Descargar Excel (.xlsx)",
        data=xlsx_bytes,
        file_name="limites_condenatorios_por_equipo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Configura filtros, equipos y variables, luego presiona **🚀 Calcular límites**.")




