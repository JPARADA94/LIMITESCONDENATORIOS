# limites_condenatorios.py
# Autor: Javier Parada
# Entrada: Excel TAL CUAL viene de SmartAssintence
#
# Llave de análisis: COMPONENTE
#
# Mejoras solicitadas:
# ✅ Categorías con “variables principales” (máx 15 por categoría: Desgaste, Salud del aceite, Contaminación)
# ✅ Títulos más profesionales y ordenados, evitando paréntesis innecesarios
# ✅ Párrafo inicial profesional y conciso, con paso a paso
# ✅ Mantiene selección por chulos, filtros opcionales, inventario, exclusión de ALERTA y descarga en Excel

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# =========================
# Configuración
# =========================
st.set_page_config(page_title="Límites Condenatorios por Componente", layout="wide")
st.title("Límites Condenatorios por Componente")

st.markdown("""
Esta herramienta calcula límites de precaución y condenatorio por componente a partir del histórico del archivo exportado desde SmartAssintence. 
Primero puedes aplicar filtros opcionales por operación, tipo de equipo, lubricante y fechas. Luego revisas el inventario de componentes y eliges 
cuáles deseas analizar. Después seleccionas las variables principales que quieres evaluar. La herramienta calcula los límites con dos métodos según el 
tamaño del histórico disponible: percentiles cuando hay suficiente información y media más desviación cuando el histórico es corto. Si existe una columna 
de estado asociada a la variable, puedes excluir del cálculo los registros que estén en alerta. Finalmente, descargas el consolidado en Excel.
""")

# =========================
# Utilidades
# =========================
@st.cache_data(show_spinner=False)
def load_excel(file) -> pd.DataFrame:
    return pd.read_excel(file)

def to_excel_bytes(df_export: pd.DataFrame, sheet_name: str = "Limites") -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def convert_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(r"[^0-9\.\-]", "", regex=True),
        errors="coerce"
    )

def clean_outliers_iqr(x: pd.Series) -> pd.Series:
    x = x.dropna()
    if len(x) < 4:
        return x
    q1, q3 = x.quantile(0.25), x.quantile(0.75)
    iqr = q3 - q1
    lo = q1 - 1.5 * iqr
    hi = q3 + 1.5 * iqr
    return x[(x >= lo) & (x <= hi)]

def build_candidates(df: pd.DataFrame) -> tuple[list, dict]:
    cols = list(df.columns)

    # Mapeo var -> columna estado
    estado_cols = [c for c in cols if " - Estado" in str(c)]
    var_to_estado = {}
    for c_estado in estado_cols:
        base = str(c_estado).replace(" - Estado ", " - Estado").replace(" - Estado", "").strip()
        if base in df.columns:
            var_to_estado[base] = c_estado

    # Excluir metadatos típicos
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
        if any(k in low for k in ["id", "codigo", "código", "serial", "placa", "bottle", "sample", "coment", "observ"]):
            continue

        if pd.api.types.is_numeric_dtype(df[c]):
            candidates.append(c)
            continue

        sample = df[c].dropna().astype(str).head(80)
        if sample.empty:
            continue
        if (sample.str.contains(r"\d", regex=True).mean() >= 0.6):
            candidates.append(c)

    # Dedupe conservando orden
    seen = set()
    candidates = [x for x in candidates if not (x in seen or seen.add(x))]
    return candidates, var_to_estado

# =========================
# Clasificación y “Top 15”
# =========================
# Lista “principal” por categoría (máx 15). Se usa como preferencia.
# Si alguna no existe en tu archivo, simplemente se ignora.
PRIMARY_WEAR = [
    "Fe", "Fe (Iron)", "Cu", "Cu (Copper)", "Pb", "Pb (Lead)", "Al", "Al (Aluminum)",
    "Cr", "Cr (Chromium)", "Ni", "Ni (Nickel)", "Sn", "Sn (Tin)", "Ag", "Ag (Silver)",
    "Ti", "Ti (Titanium)", "Mo", "Mo (Molybdenum)", "PQ Index"
][:15]

PRIMARY_HEALTH = [
    "Visc@40C (cSt)", "Visc@100C (cSt)", "TBN (mg KOH/g)", "TAN (mg KOH/g)",
    "Oxidation (Ab/cm)", "Nitration (Ab/cm)", "Sulfation (Ab/cm)",
    "AntiOx", "Antioxidant", "FTIR Oxidation", "FTIR Nitration"
][:15]

PRIMARY_CONTAM = [
    "Water (Vol%)", "Fuel Dilut. (Vol%)", "Si (Silicon)", "Na (Sodium)", "K (Potassium)",
    "Soot (Wt%)", "Glycol",
    "Particle Count  >4um", "Particle Count  >6um", "Particle Count>14um",
    "ISO 4406", "Cleanliness"
][:15]

def normalize_name(s: str) -> str:
    return str(s).strip().lower()

def categorize_variable(var_name: str) -> str:
    v = normalize_name(var_name)

    # Reglas claras y simples: salud gana sobre contaminación; contaminación gana sobre desgaste.
    health_kw = ["visc", "tbn", "tan", "oxid", "nitr", "sulf", "ftir", "antiox", "antioxid"]
    contam_kw = ["water", "agua", "fuel", "dilut", "soot", "holl", "silicon", "sodium", "potassium",
                 "particle", "iso", ">4", ">6", ">14", "glycol", "coolant", "refriger"]
    wear_kw = ["pq", "iron", "copper", "lead", "aluminum", "chrom", "nickel", "tin", "silver", "titan", "moly",
               "fe", "cu", "pb", "al", "cr", "ni", "sn", "ag", "ti", "mo"]

    if any(k in v for k in health_kw):
        return "Salud del aceite"
    if any(k in v for k in contam_kw):
        return "Contaminación"
    if any(k in v for k in wear_kw):
        return "Desgaste"
    return "Otras"

def top15_by_category(all_vars: list) -> dict:
    """
    Devuelve dict con máximo 15 por categoría:
    - Primero intenta usar las listas PRIMARY_* si están presentes en el archivo.
    - Luego completa con heurística, hasta 15.
    """
    # map de normalizado -> nombre real
    norm_map = {normalize_name(v): v for v in all_vars}

    def pick_from_primary(primary_list):
        picked = []
        for p in primary_list:
            key = normalize_name(p)
            if key in norm_map:
                picked.append(norm_map[key])
        return picked

    result = {"Desgaste": [], "Salud del aceite": [], "Contaminación": [], "Otras": []}

    # 1) preferidos
    result["Desgaste"] = pick_from_primary(PRIMARY_WEAR)
    result["Salud del aceite"] = pick_from_primary(PRIMARY_HEALTH)
    result["Contaminación"] = pick_from_primary(PRIMARY_CONTAM)

    # 2) completar con heurística hasta 15
    for v in all_vars:
        cat = categorize_variable(v)
        if cat in ["Desgaste", "Salud del aceite", "Contaminación"]:
            if v not in result[cat] and len(result[cat]) < 15:
                result[cat].append(v)

    # 3) otras (sin límite estricto, pero mostramos max 30 por usabilidad)
    for v in all_vars:
        cat = categorize_variable(v)
        if cat == "Otras":
            result["Otras"].append(v)

    result["Otras"] = result["Otras"][:30]
    return result

# =========================
# Carga de datos
# =========================
archivo = st.file_uploader("Cargar archivo Excel de SmartAssintence", type=["xlsx"])
if not archivo:
    st.stop()

df = load_excel(archivo).copy()

required_min = ["COMPONENTE", "FECHA_INFORME"]
missing = [c for c in required_min if c not in df.columns]
if missing:
    st.error(f"Faltan columnas requeridas: {missing}")
    st.stop()

df["FECHA_INFORME"] = pd.to_datetime(df["FECHA_INFORME"], errors="coerce")
if df["FECHA_INFORME"].isna().all():
    st.error("La columna FECHA_INFORME no tiene fechas válidas.")
    st.stop()

col_op = "NOMBRE_OPERACION" if "NOMBRE_OPERACION" in df.columns else None
col_tipo = "TIPO_EQUIPO" if "TIPO_EQUIPO" in df.columns else None
col_lub = "PRODUCTO" if "PRODUCTO" in df.columns else ("Tested Lubricant" if "Tested Lubricant" in df.columns else None)

logic_candidates, var_to_estado = build_candidates(df)
if not logic_candidates:
    st.error("No se encontraron variables numéricas candidatas para cálculo.")
    st.stop()

# =========================
# Sección 1. Filtros
# =========================
st.markdown("## 1. Filtros")

df_f = df.copy()

use_dates = st.checkbox("Activar filtro de fechas", value=False)
if use_dates:
    min_d = df_f["FECHA_INFORME"].min()
    max_d = df_f["FECHA_INFORME"].max()
    if not (pd.isna(min_d) or pd.isna(max_d)):
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
    use_op = st.checkbox("Activar filtro de operación", value=False, disabled=(col_op is None))
    if use_op and col_op:
        ops_sel = st.multiselect("Operación", sorted(df_f[col_op].dropna().unique()))
        if ops_sel:
            df_f = df_f[df_f[col_op].isin(ops_sel)].copy()

with cB:
    use_tipo = st.checkbox("Activar filtro de tipo de equipo", value=False, disabled=(col_tipo is None))
    if use_tipo and col_tipo:
        tipos_sel = st.multiselect("Tipo de equipo", sorted(df_f[col_tipo].dropna().unique()))
        if tipos_sel:
            df_f = df_f[df_f[col_tipo].isin(tipos_sel)].copy()

with cC:
    use_lub = st.checkbox("Activar filtro de lubricante", value=False, disabled=(col_lub is None))
    if use_lub and col_lub:
        lubs_sel = st.multiselect("Lubricante", sorted(df_f[col_lub].dropna().unique()))
        if lubs_sel:
            df_f = df_f[df_f[col_lub].isin(lubs_sel)].copy()

if df_f.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()

# =========================
# Sección 2. Inventario
# =========================
st.markdown("## 2. Inventario de componentes")

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
# Sección 3. Componentes
# =========================
st.markdown("## 3. Selección de componentes")

componentes_disponibles = sorted(df_f["COMPONENTE"].dropna().astype(str).unique())
componentes_sel = st.multiselect(
    "Componentes",
    options=componentes_disponibles,
    default=componentes_disponibles[:1] if componentes_disponibles else []
)

if not componentes_sel:
    st.warning("Selecciona al menos un componente.")
    st.stop()

df_calc = df_f[df_f["COMPONENTE"].astype(str).isin(set(map(str, componentes_sel)))].copy()
if df_calc.empty:
    st.warning("No hay registros para los componentes seleccionados.")
    st.stop()

# =========================
# Sección 4. Variables
# =========================
st.markdown("## 4. Variables para cálculo")

st.caption("Se muestran las variables principales por categoría. Usa el buscador si necesitas ubicar rápido una variable.")
buscador = st.text_input("Buscar variable", value="").strip().lower()

cats = top15_by_category(logic_candidates)

# Filtrar por buscador
if buscador:
    for k in list(cats.keys()):
        cats[k] = [v for v in cats[k] if buscador in normalize_name(v)]

# Estado de selección
if "vars_checked" not in st.session_state:
    st.session_state["vars_checked"] = set()

g1, g2, _ = st.columns([1, 1, 3])
with g1:
    if st.button("Seleccionar todo lo visible"):
        for lst in cats.values():
            st.session_state["vars_checked"].update(lst)
with g2:
    if st.button("Limpiar selección"):
        st.session_state["vars_checked"] = set()

def render_category_block(title: str, cat_key: str, expanded: bool):
    vars_list = cats.get(cat_key, [])
    if not vars_list:
        return
    with st.expander(title, expanded=expanded):
        c1, c2, _ = st.columns([1, 1, 3])
        with c1:
            if st.button(f"Seleccionar {cat_key}", key=f"all_{cat_key}"):
                st.session_state["vars_checked"].update(vars_list)
        with c2:
            if st.button(f"Limpiar {cat_key}", key=f"none_{cat_key}"):
                st.session_state["vars_checked"].difference_update(vars_list)

        cols_ui = st.columns(3)
        for i, v in enumerate(vars_list):
            col = cols_ui[i % 3]
            with col:
                checked = v in st.session_state["vars_checked"]
                new_val = st.checkbox(str(v), value=checked, key=f"chk_{cat_key}_{v}")
                if new_val:
                    st.session_state["vars_checked"].add(v)
                else:
                    st.session_state["vars_checked"].discard(v)

render_category_block("Metales de desgaste y PQ", "Desgaste", expanded=True)
render_category_block("Condición del aceite", "Salud del aceite", expanded=True)
render_category_block("Indicadores de contaminación", "Contaminación", expanded=True)
render_category_block("Otras variables disponibles", "Otras", expanded=False)

vars_sel = sorted(list(st.session_state["vars_checked"]))
if not vars_sel:
    st.warning("Selecciona al menos una variable.")
    st.stop()

st.success(f"Variables seleccionadas: {len(vars_sel)}")

for v in vars_sel:
    df_calc[v] = convert_numeric(df_calc[v])

# =========================
# Sección 5. Parámetros
# =========================
st.markdown("## 5. Parámetros de cálculo")

min_n = st.number_input("Mínimo de datos válidos por componente", min_value=2, value=3, step=1)

c1, c2 = st.columns(2)
with c1:
    n_switch = st.number_input("Umbral para percentiles", min_value=3, value=10, step=1)
    p_prec = st.slider("Percentil para precaución", 50, 99, 90, 1)
    p_cond = st.slider("Percentil para condenatorio", 50, 99, 95, 1)
with c2:
    k_prec = st.number_input("Factor para precaución en histórico corto", min_value=0.0, value=2.0, step=0.5)
    k_cond = st.number_input("Factor para condenatorio en histórico corto", min_value=0.0, value=3.0, step=0.5)

st.markdown("### Control de registros atípicos")

excluir_alertas = st.checkbox("Excluir registros en alerta cuando exista estado de la variable", value=True)
excluir_precaucion = st.checkbox("Excluir también registros en precaución", value=False)
limpieza_iqr = st.checkbox("Aplicar limpieza adicional por IQR", value=False)

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
        return {"n": n, "metodo": "Insuficiente", "prec": np.nan, "cond": np.nan,
                "mean": np.nan, "std": np.nan, "median": np.nan, "min": np.nan, "max": np.nan}

    if limpieza_iqr:
        s2 = clean_outliers_iqr(s)
        if len(s2) < min_n:
            return {"n": int(len(s2)), "metodo": "Insuficiente", "prec": np.nan, "cond": np.nan,
                    "mean": np.nan, "std": np.nan, "median": np.nan, "min": np.nan, "max": np.nan}
        s = s2
        n = int(len(s))

    mean = float(s.mean())
    std = float(s.std(ddof=1)) if n > 1 else 0.0
    median = float(s.median())
    vmin = float(s.min())
    vmax = float(s.max())

    if n >= n_switch:
        prec = float(s.quantile(p_prec / 100))
        cond = float(s.quantile(p_cond / 100))
        metodo = f"Percentiles P{p_prec} y P{p_cond}"
    else:
        prec = mean + (k_prec * std)
        cond = mean + (k_cond * std)
        metodo = "Media y desviación"

    return {"n": n, "metodo": metodo, "prec": prec, "cond": cond,
            "mean": mean, "std": std, "median": median, "min": vmin, "max": vmax}

# =========================
# Sección 6. Resultados
# =========================
st.markdown("## 6. Resultados y exportación")

if st.button("Calcular límites"):
    rows = []
    for comp, g_comp in df_calc.groupby(df_calc["COMPONENTE"].astype(str)):
        for v in vars_sel:
            serie = apply_estado_filter(g_comp, v)
            out = compute_limits(serie)

            row = {
                "COMPONENTE": comp,
                "VARIABLE": v,
                "CATEGORIA": categorize_variable(v),
                "n": out["n"],
                "metodo": out["metodo"],
                "precaucion": out["prec"],
                "condenatorio": out["cond"],
                "mean": out["mean"],
                "std": out["std"],
                "median": out["median"],
                "min": out["min"],
                "max": out["max"],
                "fecha_min": g_comp["FECHA_INFORME"].min(),
                "fecha_max": g_comp["FECHA_INFORME"].max(),
                "excluye_alerta": "SI" if excluir_alertas else "NO",
                "excluye_precaucion": "SI" if excluir_precaucion else "NO",
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

    dec = st.number_input("Decimales para visualización", min_value=0, value=0, step=1)
    res_show = res.copy()
    for c in ["precaucion", "condenatorio", "mean", "std", "median", "min", "max"]:
        res_show[c] = res_show[c].round(dec)

    cols_order = []
    for c in ["COMPONENTE", col_op, col_tipo, col_lub, "VARIABLE", "CATEGORIA",
              "tiene_estado", "n", "metodo", "precaucion", "condenatorio",
              "mean", "std", "median", "min", "max", "fecha_min", "fecha_max",
              "excluye_alerta", "excluye_precaucion", "iqr"]:
        if c and c in res_show.columns:
            cols_order.append(c)

    st.dataframe(res_show[cols_order].sort_values(["COMPONENTE", "CATEGORIA", "VARIABLE"]), use_container_width=True)

    xlsx_bytes = to_excel_bytes(res_show[cols_order], sheet_name="Limites")
    st.download_button(
        "Descargar Excel",
        data=xlsx_bytes,
        file_name="limites_condenatorios_por_componente.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Ajusta filtros, selecciona componentes y variables, y luego calcula los límites.")



