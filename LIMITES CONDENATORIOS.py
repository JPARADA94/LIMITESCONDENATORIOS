# limites_condenatorios.py
# Autor: Javier Parada
# Entrada: Excel TAL CUAL viene de SmartAssintence
#
# CAMBIO CLAVE:
# - Llave de análisis: COMPONENTE (NO EQUIPO)
#
# NUEVO (según tu solicitud):
# ✅ Selección de variables por “chulos” pero ORGANIZADAS por categorías:
#    - Desgaste
#    - Salud del aceite
#    - Contaminación
#    - Otras
#
# ✅ Sigue igual:
# - Filtros opcionales (Operación / Tipo / Lubricante + fechas)
# - Inventario por COMPONENTE
# - Selección de COMPONENTES
# - Excluir ALERTA (si existe <Variable> - Estado)
# - Descarga en Excel

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Límites Condenatorios - por COMPONENTE", layout="wide")
st.title("📏 Límites Condenatorios por COMPONENTE (Histórico)")

st.markdown("""
Esta app calcula **límites de Precaución y Condenatorio por COMPONENTE**, usando variables numéricas del Excel de SmartAssintence.
Puedes filtrar (opcional) por **Operación**, **Tipo de equipo** y **Lubricante**. Luego seleccionas los **COMPONENTES** y las **variables**
a evaluar. Si una variable tiene su columna **`<Variable> - Estado`**, puedes excluir del cálculo los registros en **ALERTA**.
""")

# --------------------
# Utilidades
# --------------------
@st.cache_data(show_spinner=False)
def load_excel(file) -> pd.DataFrame:
    return pd.read_excel(file)

def to_excel_bytes(df_export: pd.DataFrame, sheet_name: str = "Limites") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def clean_outliers_iqr(x: pd.Series) -> pd.Series:
    x = x.dropna()
    if len(x) < 4:
        return x
    q1, q3 = x.quantile(0.25), x.quantile(0.75)
    iqr = q3 - q1
    lo = q1 - 1.5 * iqr
    hi = q3 + 1.5 * iqr
    return x[(x >= lo) & (x <= hi)]

def convert_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(r"[^0-9\.\-]", "", regex=True),
        errors="coerce"
    )

def build_candidates(df: pd.DataFrame) -> tuple[list, dict]:
    """
    Variables “lógicas” = candidatas numéricas/convertibles que sí valen para límites.
    Excluye metadatos típicos y columnas '- Estado'. Si existe '<var> - Estado', se registra para filtro de ALERTA.
    """
    cols = list(df.columns)

    # Mapeo var -> columna "<var> - Estado"
    estado_cols = [c for c in cols if " - Estado" in str(c)]
    var_to_estado = {}
    for c_estado in estado_cols:
        base = str(c_estado).replace(" - Estado ", " - Estado").replace(" - Estado", "").strip()
        if base in df.columns:
            var_to_estado[base] = c_estado

    # Excluir metadatos típicos (ajusta aquí si tu SmartAssintence trae otros nombres)
    exclude_exact = {
        "COMPONENTE", "FECHA_INFORME",
        "NOMBRE_CLIENTE", "CLIENTE",
        "NOMBRE_OPERACION", "OPERACION",
        "TIPO_EQUIPO",
        "PRODUCTO", "Tested Lubricant",
        "ESTADO_REPORTE", "Report Status",
        "CORRELATIVO", "N_MUESTRA",
        "Sample Bottle ID", "Asset ID", "EQUIPO"
    }

    candidates = []
    for c in cols:
        cs = str(c)
        low = cs.lower()

        if c in exclude_exact:
            continue
        if " - Estado" in cs:
            continue

        # Excluir nombres que parezcan IDs / llaves / textos “administrativos”
        if any(k in low for k in ["id", "codigo", "código", "serial", "placa", "bottle", "sample", "coment", "observ"]):
            continue

        # Incluir numéricas
        if pd.api.types.is_numeric_dtype(df[c]):
            candidates.append(c)
            continue

        # Incluir convertibles (si la mayoría contiene dígitos)
        sample = df[c].dropna().astype(str).head(80)
        if sample.empty:
            continue
        looks_numeric = (sample.str.contains(r"\d", regex=True).mean() >= 0.6)
        if looks_numeric:
            candidates.append(c)

    # Prioriza variables que tienen "- Estado" (normalmente son “las importantes”)
    with_estado = [c for c in candidates if c in var_to_estado]
    without_estado = [c for c in candidates if c not in var_to_estado]
    ordered = with_estado + without_estado

    # Dedupe conservando orden
    seen = set()
    ordered = [x for x in ordered if not (x in seen or seen.add(x))]
    return ordered, var_to_estado

def categorize_variable(var_name: str) -> str:
    """
    Clasifica la variable en:
    - Desgaste
    - Salud del aceite
    - Contaminación
    - Otras

    (Heurística por palabras clave en ES/EN)
    """
    v = str(var_name).lower()

    # ---- Desgaste (metales típicos + PQ) ----
    wear_kw = [
        "fe", "iron", "cu", "copper", "pb", "lead", "sn", "tin",
        "cr", "chrom", "al", "alum", "ni", "nickel", "ag", "silver",
        "ti", "titan", "mo", "moly", "mn", "mangan", "zn", "zinc",
        "wear", "desgaste", "pq", "pq index"
    ]

    # ---- Contaminación ----
    contam_kw = [
        "water", "agua", "h2o", "%water", "vol%", "coolant", "refriger",
        "k (", "potassium", "na (", "sodium", "si (", "silicon", "glycol",
        "fuel", "diesel", "dilut", "dilution", "combustible",
        "soot", "hollin", "hollín", "dust", "polvo",
        "particle", "partícula", "particula", "iso", "4406", "cleanliness",
        ">4", ">6", ">14", "visc@40", "visc@100"  # viscosidad puede ser salud, pero a veces se analiza como contaminación/adulteración; la ponemos en salud abajo, y si coincide con salud gana salud.
    ]

    # ---- Salud del aceite ----
    health_kw = [
        "visc", "viscos", "tbn", "tan", "oxid", "nitr", "sulf", "sulph",
        "antioxid", "rul", "ftir", "ab/cm", "base number", "acid number",
        "additive", "aditivo", "aw", "antiwear", "zn", "phosph", "p (", "phos",
        "ca (", "calcium", "mg (", "magnesium", "b (", "boron", "ba (", "barium",
        "foam", "espuma"
    ]

    # Reglas de prioridad:
    # 1) Si parece salud (visc/TBN/TAN/oxid/nitr), salud gana.
    if any(k in v for k in health_kw):
        return "Salud del aceite"
    # 2) Si parece contaminación (agua, K, Na, Si, fuel, soot, ISO), contaminación.
    if any(k in v for k in contam_kw):
        return "Contaminación"
    # 3) Si parece desgaste (metales + PQ), desgaste.
    if any(k in v for k in wear_kw):
        return "Desgaste"
    # 4) Otras
    return "Otras"


# --------------------
# Carga
# --------------------
archivo = st.file_uploader("📁 Sube tu Excel (.xlsx) de SmartAssintence (tal cual)", type=["xlsx"])
if not archivo:
    st.stop()

df = load_excel(archivo).copy()

# Validaciones mínimas
required_min = ["COMPONENTE", "FECHA_INFORME"]
missing = [c for c in required_min if c not in df.columns]
if missing:
    st.error(f"❌ Falta(n) columna(s) mínima(s) requerida(s): {missing}")
    st.stop()

df["FECHA_INFORME"] = pd.to_datetime(df["FECHA_INFORME"], errors="coerce")
if df["FECHA_INFORME"].isna().all():
    st.error("❌ FECHA_INFORME no tiene fechas válidas (NaT). Revisa el Excel.")
    st.stop()

# Columnas opcionales para filtros
col_op = "NOMBRE_OPERACION" if "NOMBRE_OPERACION" in df.columns else None
col_tipo = "TIPO_EQUIPO" if "TIPO_EQUIPO" in df.columns else None
col_lub = "PRODUCTO" if "PRODUCTO" in df.columns else ("Tested Lubricant" if "Tested Lubricant" in df.columns else None)

# Variables candidatas “lógicas”
logic_candidates, var_to_estado = build_candidates(df)
if not logic_candidates:
    st.error("❌ No encontré variables numéricas/convertibles candidatas para límites.")
    st.stop()

# =========================
# 1) FILTROS (opcionales)
# =========================
st.markdown("## 1) Filtros (opcionales)")
df_f = df.copy()

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
        df_f = df_f[(df_f["FECHA_INFORME"] >= pd.to_datetime(d1)) &
                    (df_f["FECHA_INFORME"] <= pd.to_datetime(d2))].copy()

cA, cB, cC = st.columns(3)

with cA:
    use_op = st.checkbox("Filtrar por Operación", value=False, disabled=(col_op is None))
    if col_op is None:
        st.caption("No existe NOMBRE_OPERACION en el Excel.")
    if use_op and col_op:
        ops_sel = st.multiselect("Operación(es)", sorted(df_f[col_op].dropna().unique()))
        if ops_sel:
            df_f = df_f[df_f[col_op].isin(ops_sel)].copy()

with cB:
    use_tipo = st.checkbox("Filtrar por Tipo de equipo", value=False, disabled=(col_tipo is None))
    if col_tipo is None:
        st.caption("No existe TIPO_EQUIPO en el Excel.")
    if use_tipo and col_tipo:
        tipos_sel = st.multiselect("Tipo(s) de equipo", sorted(df_f[col_tipo].dropna().unique()))
        if tipos_sel:
            df_f = df_f[df_f[col_tipo].isin(tipos_sel)].copy()

with cC:
    use_lub = st.checkbox("Filtrar por Lubricante", value=False, disabled=(col_lub is None))
    if col_lub is None:
        st.caption("No existe PRODUCTO / Tested Lubricant en el Excel.")
    if use_lub and col_lub:
        lubs_sel = st.multiselect("Lubricante(s)", sorted(df_f[col_lub].dropna().unique()))
        if lubs_sel:
            df_f = df_f[df_f[col_lub].isin(lubs_sel)].copy()

if df_f.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()

# =========================
# 2) INVENTARIO (post-filtros)
# =========================
st.markdown("## 2) Inventario de COMPONENTES (post-filtros)")
st.caption("Consolida cuántas muestras tiene cada COMPONENTE (con los filtros aplicados).")

group_cols = ["COMPONENTE"]
if col_op: group_cols.append(col_op)
if col_tipo: group_cols.append(col_tipo)
if col_lub: group_cols.append(col_lub)

inv = (
    df_f.groupby(group_cols, dropna=False)
        .agg(
            muestras_totales=("COMPONENTE", "size"),
            fecha_min=("FECHA_INFORME", "min"),
            fecha_max=("FECHA_INFORME", "max"),
            ultima_muestra=("FECHA_INFORME", "max"),
        )
        .reset_index()
        .sort_values("muestras_totales", ascending=False)
)

st.dataframe(inv, use_container_width=True)

# =========================
# 3) Selección de COMPONENTES
# =========================
st.markdown("## 3) Selecciona el/los COMPONENTE(s)")

componentes_disponibles = sorted(df_f["COMPONENTE"].dropna().astype(str).unique())
componentes_sel = st.multiselect(
    "COMPONENTE(s)",
    options=componentes_disponibles,
    default=componentes_disponibles[:1] if componentes_disponibles else []
)

if not componentes_sel:
    st.warning("Selecciona al menos un COMPONENTE.")
    st.stop()

df_calc = df_f[df_f["COMPONENTE"].astype(str).isin(set(map(str, componentes_sel)))].copy()
if df_calc.empty:
    st.warning("No hay registros para los COMPONENTES seleccionados con los filtros actuales.")
    st.stop()

# =========================
# 4) Selección de VARIABLES por CHULOS (organizadas)
# =========================
st.markdown("## 4) Selecciona variables lógicas a analizar (por chulos y por categoría)")

# Estado de selección
if "vars_checked" not in st.session_state:
    st.session_state["vars_checked"] = set()

buscador = st.text_input("Buscar variable (opcional)", value="").strip().lower()

# Clasificar variables (post-buscador)
cats = {"Desgaste": [], "Salud del aceite": [], "Contaminación": [], "Otras": []}
for v in logic_candidates:
    if buscador and buscador not in str(v).lower():
        continue
    cat = categorize_variable(v)
    cats.setdefault(cat, []).append(v)

# Botones globales
g1, g2, _ = st.columns([1, 1, 3])
with g1:
    if st.button("✅ Seleccionar TODAS (según búsqueda)"):
        for cat_vars in cats.values():
            st.session_state["vars_checked"].update(cat_vars)
with g2:
    if st.button("🧹 Limpiar selección"):
        st.session_state["vars_checked"] = set()

st.caption("Tip: usa el buscador para filtrar rápido, y luego marca por categoría.")

def render_category(cat_name: str, vars_list: list):
    if not vars_list:
        return

    # Controles por categoría
    c1, c2, c3 = st.columns([1, 1, 3])
    with c1:
        if st.button(f"✅ Todo {cat_name}", key=f"all_{cat_name}"):
            st.session_state["vars_checked"].update(vars_list)
    with c2:
        if st.button(f"🧹 Limpiar {cat_name}", key=f"none_{cat_name}"):
            st.session_state["vars_checked"].difference_update(vars_list)

    # Checkboxes en 3 columnas
    ncols = 3
    cols_ui = st.columns(ncols)
    for i, v in enumerate(vars_list):
        col = cols_ui[i % ncols]
        with col:
            checked = v in st.session_state["vars_checked"]
            new_val = st.checkbox(str(v), value=checked, key=f"chk_{cat_name}_{v}")
            if new_val:
                st.session_state["vars_checked"].add(v)
            else:
                st.session_state["vars_checked"].discard(v)

# Mostrar categorías (puedes cambiar el orden si quieres)
with st.expander("🧱 Desgaste (metales / PQ)", expanded=True):
    render_category("Desgaste", cats["Desgaste"])

with st.expander("🧪 Salud del aceite (viscosidad / TAN / TBN / oxidación / nitración…)", expanded=True):
    render_category("Salud del aceite", cats["Salud del aceite"])

with st.expander("💧 Contaminación (agua / combustible / silicio / partículas / ISO…)", expanded=True):
    render_category("Contaminación", cats["Contaminación"])

with st.expander("📌 Otras", expanded=False):
    render_category("Otras", cats["Otras"])

vars_sel = sorted(list(st.session_state["vars_checked"]))
if not vars_sel:
    st.warning("Selecciona al menos una variable.")
    st.stop()

st.success(f"Variables seleccionadas: {len(vars_sel)}")

# Convertir a numérico SOLO las seleccionadas (en memoria)
for v in vars_sel:
    df_calc[v] = convert_numeric(df_calc[v])

# =========================
# 5) Parámetros + filtro por estado
# =========================
st.markdown("## 5) Parámetros de cálculo")

min_n = st.number_input("Mínimo de datos válidos por COMPONENTE (por variable) para calcular límites", min_value=2, value=3, step=1)

c1, c2 = st.columns(2)
with c1:
    n_switch = st.number_input("Umbral n para usar percentiles (si n ≥ umbral)", min_value=3, value=10, step=1)
    p_prec = st.slider("Percentil Precaución (n ≥ umbral)", 50, 99, 90, 1)
    p_cond = st.slider("Percentil Condenatorio (n ≥ umbral)", 50, 99, 95, 1)
with c2:
    k_prec = st.number_input("k Precaución (n < umbral): μ + k·σ", min_value=0.0, value=2.0, step=0.5)
    k_cond = st.number_input("k Condenatorio (n < umbral): μ + k·σ", min_value=0.0, value=3.0, step=0.5)

st.markdown("### Filtro de atípicos por estado (si existe `<Variable> - Estado`)")

excluir_alertas = st.checkbox("Excluir del cálculo los registros en ALERTA (por variable)", value=True)
excluir_precaucion = st.checkbox("Excluir también registros en PRECAUCIÓN (opcional)", value=False)
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
# 6) Calcular + Descargar Excel
# =========================
st.markdown("## 6) Resultados")

if st.button("🚀 Calcular límites"):
    rows = []
    for comp, g_comp in df_calc.groupby(df_calc["COMPONENTE"].astype(str)):
        for v in vars_sel:
            serie_filtrada = apply_estado_filter(g_comp, v)
            out = compute_limits(serie_filtrada)

            row = {
                "COMPONENTE": comp,
                "VARIABLE": v,
                "CATEGORIA": categorize_variable(v),
                "n": out["n"],
                "metodo": out["metodo"],
                "precaucion": out["precaucion"],
                "condenatorio": out["condenatorio"],
                "mean": out["mean"],
                "std": out["std"],
                "median": out["median"],
                "min": out["min"],
                "max": out["max"],
                "fecha_min": g_comp["FECHA_INFORME"].min(),
                "fecha_max": g_comp["FECHA_INFORME"].max(),
                "excluye_ALERTA": "SI" if excluir_alertas else "NO",
                "excluye_PRECAUCION": "SI" if excluir_precaucion else "NO",
                "iqr": "SI" if limpieza_iqr else "NO",
                "tiene_estado": "SI" if v in var_to_estado else "NO",
            }

            if col_op and col_op in g_comp.columns:
                row[col_op] = g_comp[col_op].dropna().iloc[0] if g_comp[col_op].notna().any() else None
            if col_tipo and col_tipo in g_comp.columns:
                row[col_tipo] = g_comp[col_tipo].dropna().iloc[0] if g_comp[col_tipo].notna().any() else None
            if col_lub and col_lub in g_comp.columns:
                row[col_lub] = g_comp[col_lub].dropna().iloc[0] if g_comp[col_lub].notna().any() else None

            rows.append(row)

    res = pd.DataFrame(rows)

    dec = st.number_input("Decimales para mostrar (solo visual)", min_value=0, value=0, step=1)
    res_show = res.copy()
    for c in ["precaucion", "condenatorio", "mean", "std", "median", "min", "max"]:
        res_show[c] = res_show[c].round(dec)

    cols_order = []
    for c in ["COMPONENTE", col_op, col_tipo, col_lub, "VARIABLE", "CATEGORIA",
              "tiene_estado", "n", "metodo", "precaucion", "condenatorio",
              "mean", "std", "median", "min", "max", "fecha_min", "fecha_max",
              "excluye_ALERTA", "excluye_PRECAUCION", "iqr"]:
        if c and c in res_show.columns:
            cols_order.append(c)

    st.dataframe(
        res_show[cols_order].sort_values(["COMPONENTE", "CATEGORIA", "VARIABLE"]),
        use_container_width=True
    )

    st.markdown("### Descargar en Excel (.xlsx)")
    xlsx_bytes = to_excel_bytes(res_show[cols_order], sheet_name="Limites")
    st.download_button(
        "⬇️ Descargar Excel (.xlsx)",
        data=xlsx_bytes,
        file_name="limites_condenatorios_por_componente.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Aplica filtros si quieres, selecciona COMPONENTES y variables, luego presiona **🚀 Calcular límites**.")








