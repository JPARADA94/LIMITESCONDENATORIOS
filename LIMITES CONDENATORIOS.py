# limites_condenatorios.py
# Autor: Javier Parada
# Entrada: Excel TAL CUAL viene de SmartAssintence
#
# Llave de análisis: COMPONENTE
# Salida: Excel con límites por componente y variable

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
Puedes aplicar filtros opcionales por operación, tipo de equipo, lubricante y fechas. Luego revisas el inventario y seleccionas los componentes a analizar. 
Después eliges las variables principales organizadas por categorías. Los límites se calculan con dos métodos según el tamaño del histórico: percentiles cuando 
hay suficiente información y media más desviación cuando el histórico es corto. Si la variable cuenta con una columna de estado, puedes excluir del cálculo 
los registros que estén en alerta. Al final descargas el consolidado en Excel.
""")

# =========================
# Utilidades
# =========================
@st.cache_data(show_spinner=False)
def load_excel(file) -> pd.DataFrame:
    return pd.read_excel(file)

def to_excel_bytes(df_export: pd.DataFrame, sheet_name: str = "Resultados") -> bytes:
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

    # Mapeo variable -> columna de estado
    estado_cols = [c for c in cols if " - Estado" in str(c)]
    var_to_estado = {}
    for c_estado in estado_cols:
        base = str(c_estado).replace(" - Estado ", " - Estado").replace(" - Estado", "").strip()
        if base in df.columns:
            var_to_estado[base] = c_estado

    # Excluir columnas de contexto y administrativas
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

    seen = set()
    candidates = [x for x in candidates if not (x in seen or seen.add(x))]
    return candidates, var_to_estado

def _n(s: str) -> str:
    return str(s).strip().lower()

# =========================
# Clasificación alineada a SmartAssintence
# =========================
PRIMARY_WEAR = [
    "Plata (Ag) (mg/kg) ASTM D5185",
    "Aluminio (Al) (mg/kg) ASTM D5185",
    "Cromo (Cr) (mg/kg) ASTM D5185",
    "Cobre (Cu) (mg/kg) ASTM D5185",
    "Hierro (Fe) (mg/kg) ASTM D5185",
    "Índice PQ (PQI) (Adimensional) ASTM D8184",
    "Níquel (Ni) (mg/kg) ASTM D5185",
    "Plomo (Pb) (mg/kg) ASTM D5185",
    "Estaño (Sn) (mg/kg) ASTM D5185",
    "Titanio (Ti) (mg/kg) ASTM D5185",
][:15]

PRIMARY_PROPERTIES = [
    "Número Básico (BN) (mg KOH/g) ASTM D2896",
    "Número Ácido (AN) (mg KOH/g) ASTM D664",
    "Viscosidad a 40 °C (mm²/s) ASTM D445",
    "Viscosidad a 100 °C (mm²/s) ASTM D445",
    "Visc@40C (cSt)",
    "Visc@100C (cSt)",
    "TBN (mg KOH/g)",
    "TAN (mg KOH/g)",
    "Oxidación (Abs/cm) ASTM D7414",
    "Nitración (Abs/cm) ASTM D7624",
    "Oxidation (Ab/cm)",
    "Nitration (Ab/cm)",
][:15]

PRIMARY_CONTAM = [
    "Agua (IR) (% v/v) ASTM E2412",
    "Water (Vol%)",
    "Hollín (% w/w) ASTM D7844",
    "Soot (Wt%)",
    "Silicio (Si) (mg/kg) ASTM D5185",
    "Si (Silicon)",
    "Sodio (Na) (mg/kg) ASTM D5185",
    "Na (Sodium)",
    "Potasio (K) (mg/kg) ASTM D5185",
    "K (Potassium)",
    "Cadmio (Cd) (mg/kg) ASTM D5185",
    "Manganeso (Mn) (mg/kg) ASTM D5185",
    "Vanadio (V) (mg/kg) ASTM D5185",
    "Fuel Dilut. (Vol%)",
    "Glycol",
][:15]

PRIMARY_ADDITIVES = [
    "Calcio (Ca) (mg/kg) ASTM D5185",
    "Magnesio (Mg) (mg/kg) ASTM D5185",
    "Zinc (Zn) (mg/kg) ASTM D5185",
    "Fósforo (P) (mg/kg) ASTM D5185",
    "Boro (B) (mg/kg) ASTM D5185",
    "Molibdeno (Mo) (mg/kg) ASTM D5185",
    "Ca (Calcium)", "Mg (Magnesium)", "Zn (Zinc)", "P (Phosphorus)", "B (Boron)", "Mo (Molybdenum)"
][:15]

EXCLUDE_ALWAYS = {"Periodo uso aceite", "Unidad uso aceite"}

def categorize_variable(var_name: str) -> str:
    v = _n(var_name)

    if any(_n(x) in v for x in EXCLUDE_ALWAYS):
        return "No usar"

    prop_kw = ["visc", "viscos", "tbn", "tan", "bn", "an", "oxid", "nitr", "sulf", "ftir", "ab/cm", "acid", "base number"]
    if any(k in v for k in prop_kw):
        return "Propiedades del lubricante"

    contam_kw = [
        "agua", "water", "soot", "holl", "silic", "sodium", "sodio", "potassium", "potasio",
        "glycol", "fuel", "dilut", "particle", "iso 4406", "coolant", "refriger"
    ]
    if any(k in v for k in contam_kw):
        return "Contaminantes"

    wear_kw = [
        "hierro", "fe", "cobre", "cu", "plomo", "pb", "alumin", "al", "cromo", "cr",
        "niquel", "ni", "estaño", "sn", "plata", "ag", "titan", "ti", "pq", "pqi",
        "nickel", "chrom", "copper", "lead", "iron", "aluminum", "tin", "silver", "titanium"
    ]
    if any(k in v for k in wear_kw):
        return "Desgaste"

    add_kw = [
        "calcio", "ca", "magnes", "mg", "zinc", "zn", "fosfor", "phosph", "boro", "boron", "molyb", "molib"
    ]
    if any(k in v for k in add_kw):
        return "Aditivos"

    return "Otras variables"

def top15_by_category(all_vars: list) -> dict:
    norm_map = {_n(v): v for v in all_vars}

    def pick(primary):
        out = []
        for p in primary:
            key = _n(p)
            if key in norm_map:
                out.append(norm_map[key])
        return out

    cats = {
        "Desgaste": pick(PRIMARY_WEAR),
        "Propiedades del lubricante": pick(PRIMARY_PROPERTIES),
        "Contaminantes": pick(PRIMARY_CONTAM),
        "Aditivos": pick(PRIMARY_ADDITIVES),
        "Otras variables": []
    }

    for v in all_vars:
        c = categorize_variable(v)
        if c == "No usar":
            continue
        if c in ["Desgaste", "Propiedades del lubricante", "Contaminantes", "Aditivos"]:
            if v not in cats[c] and len(cats[c]) < 15:
                cats[c].append(v)

    for v in all_vars:
        if categorize_variable(v) == "Otras variables":
            cats["Otras variables"].append(v)
    cats["Otras variables"] = cats["Otras variables"][:30]
    return cats

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
# 1. Filtros
# =========================
st.markdown("## 1. Filtros de análisis")

df_f = df.copy()

use_dates = st.checkbox("Activar filtro por fechas", value=False)
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
    use_op = st.checkbox("Activar filtro por operación", value=False, disabled=(col_op is None))
    if use_op and col_op:
        ops_sel = st.multiselect("Operación", sorted(df_f[col_op].dropna().unique()))
        if ops_sel:
            df_f = df_f[df_f[col_op].isin(ops_sel)].copy()

with cB:
    use_tipo = st.checkbox("Activar filtro por tipo de equipo", value=False, disabled=(col_tipo is None))
    if use_tipo and col_tipo:
        tipos_sel = st.multiselect("Tipo de equipo", sorted(df_f[col_tipo].dropna().unique()))
        if tipos_sel:
            df_f = df_f[df_f[col_tipo].isin(tipos_sel)].copy()

with cC:
    use_lub = st.checkbox("Activar filtro por lubricante", value=False, disabled=(col_lub is None))
    if use_lub and col_lub:
        lubs_sel = st.multiselect("Lubricante", sorted(df_f[col_lub].dropna().unique()))
        if lubs_sel:
            df_f = df_f[df_f[col_lub].isin(lubs_sel)].copy()

if df_f.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()

# =========================
# 2. Inventario
# =========================
st.markdown("## 2. Inventario de componentes")
st.caption("Consolidado del histórico disponible por componente después de aplicar filtros.")

group_cols = ["COMPONENTE"]
if col_op: group_cols.append(col_op)
if col_tipo: group_cols.append(col_tipo)
if col_lub: group_cols.append(col_lub)

inventario = (
    df_f.groupby(group_cols, dropna=False)
        .agg(
            muestras=("COMPONENTE", "size"),
            primera_fecha=("FECHA_INFORME", "min"),
            ultima_fecha=("FECHA_INFORME", "max"),
        )
        .reset_index()
        .sort_values("muestras", ascending=False)
)

st.dataframe(inventario, use_container_width=True)

# =========================
# 3. Selección de componentes
# =========================
st.markdown("## 3. Selección de componentes")

componentes_disponibles = sorted(df_f["COMPONENTE"].dropna().astype(str).unique())
componentes_sel = st.multiselect(
    "Componentes a analizar",
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
# 4. Variables para cálculo
# =========================
st.markdown("## 4. Variables para el cálculo")

st.caption("Se muestran variables principales por categoría. Usa el buscador para ubicar una variable rápidamente.")
buscador = st.text_input("Buscar variable", value="").strip().lower()

cats = top15_by_category(logic_candidates)
if buscador:
    for k in list(cats.keys()):
        cats[k] = [v for v in cats[k] if buscador in _n(v)]

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

render_category_block("Desgaste", "Desgaste", expanded=True)
render_category_block("Propiedades del lubricante", "Propiedades del lubricante", expanded=True)
render_category_block("Contaminantes", "Contaminantes", expanded=True)
render_category_block("Aditivos", "Aditivos", expanded=False)
render_category_block("Otras variables", "Otras variables", expanded=False)

vars_sel = sorted(list(st.session_state["vars_checked"]))
if not vars_sel:
    st.warning("Selecciona al menos una variable.")
    st.stop()

st.success(f"Variables seleccionadas: {len(vars_sel)}")

for v in vars_sel:
    df_calc[v] = convert_numeric(df_calc[v])

# =========================
# 5. Parámetros de cálculo
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
# 6. Resultados y exportación
# =========================
st.markdown("## 6. Resultados y exportación")

if st.button("Calcular límites"):
    filas = []
    for comp, g_comp in df_calc.groupby(df_calc["COMPONENTE"].astype(str)):
        for v in vars_sel:
            serie = apply_estado_filter(g_comp, v)
            out = compute_limits(serie)

            fila = {
                "Componente": comp,
                "Variable": v,
                "Categoría": categorize_variable(v),
                "Datos válidos": out["n"],
                "Método": out["metodo"],
                "Límite de precaución": out["prec"],
                "Límite condenatorio": out["cond"],
                "Promedio": out["mean"],
                "Desviación estándar": out["std"],
                "Mediana": out["median"],
                "Mínimo": out["min"],
                "Máximo": out["max"],
                "Primera fecha": g_comp["FECHA_INFORME"].min(),
                "Última fecha": g_comp["FECHA_INFORME"].max(),
                "Excluye alerta": "Sí" if excluir_alertas else "No",
                "Excluye precaución": "Sí" if excluir_precaucion else "No",
                "Limpieza IQR": "Sí" if limpieza_iqr else "No",
                "Tiene estado": "Sí" if v in var_to_estado else "No",
            }

            if col_op and col_op in g_comp.columns:
                fila["Operación"] = g_comp[col_op].dropna().iloc[0] if g_comp[col_op].notna().any() else None
            if col_tipo and col_tipo in g_comp.columns:
                fila["Tipo de equipo"] = g_comp[col_tipo].dropna().iloc[0] if g_comp[col_tipo].notna().any() else None
            if col_lub and col_lub in g_comp.columns:
                fila["Lubricante"] = g_comp[col_lub].dropna().iloc[0] if g_comp[col_lub].notna().any() else None

            filas.append(fila)

    resultados = pd.DataFrame(filas)

    dec = st.number_input("Decimales para visualización", min_value=0, value=0, step=1)
    resultados_vista = resultados.copy()
    for c in ["Límite de precaución", "Límite condenatorio", "Promedio", "Desviación estándar", "Mediana", "Mínimo", "Máximo"]:
        resultados_vista[c] = pd.to_numeric(resultados_vista[c], errors="coerce").round(dec)

    orden = [
        "Operación", "Tipo de equipo", "Lubricante",
        "Componente", "Categoría", "Variable",
        "Datos válidos", "Método",
        "Límite de precaución", "Límite condenatorio",
        "Promedio", "Desviación estándar", "Mediana", "Mínimo", "Máximo",
        "Primera fecha", "Última fecha",
        "Excluye alerta", "Excluye precaución", "Limpieza IQR", "Tiene estado"
    ]
    columnas = [c for c in orden if c in resultados_vista.columns]

    st.dataframe(
        resultados_vista[columnas].sort_values(["Componente", "Categoría", "Variable"]),
        use_container_width=True
    )

    archivo_excel = to_excel_bytes(resultados_vista[columnas], sheet_name="Limites")
    st.download_button(
        "Descargar Excel",
        data=archivo_excel,
        file_name="limites_condenatorios_por_componente.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Ajusta filtros, selecciona componentes y variables, y luego calcula los límites.")





