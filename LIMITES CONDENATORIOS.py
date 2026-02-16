# limites_condenatorios.py
# Autor: Javier Parada
# Entrada: Excel TAL CUAL viene de SmartAssintence
# Llave de análisis: COMPONENTE
# Salida: Excel con límites por componente y variable

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re

# =========================
# Configuración
# =========================
st.set_page_config(page_title="Límites Condenatorios por Componente", layout="wide")
st.title("Límites Condenatorios por Componente")

st.markdown("""
Esta herramienta calcula límites de precaución y condenatorio por componente usando el histórico del archivo exportado desde SmartAssintence.
Puedes aplicar filtros opcionales por operación, tipo de equipo, lubricante y fechas. Luego revisas el inventario y seleccionas los componentes.
Después eliges las variables principales organizadas por categorías. Los límites se calculan con dos métodos según la cantidad de datos:
percentiles cuando hay histórico suficiente y media más desviación cuando el histórico es corto. Si la variable tiene una columna de estado,
puedes excluir del cálculo los registros en alerta. Al final descargas el consolidado en Excel.
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

def normalize_name(s: str) -> str:
    """Normaliza nombres para comparar: lower, sin dobles espacios, sin sufijos tipo ' - 20'."""
    txt = str(s).strip().lower()
    txt = re.sub(r"\s+", " ", txt)
    txt = re.sub(r"\s*-\s*\d+\s*$", "", txt)  # quita " - 20" al final
    return txt

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

# =========================
# Clasificación alineada a SmartAssintence (estricta)
# =========================
# Nota: se comparan con normalize_name(), y por eso no importa si en tu archivo vienen como "COBRE (CU) - 25"
PRIMARY_WEAR = [
    "plata (ag)", "aluminio (al)", "cromo (cr)", "cobre (cu)", "hierro (fe)",
    "índice pq", "indice pq", "pq index", "pqi",
    "níquel (ni)", "niquel (ni)", "plomo (pb)", "estaño (sn)", "estano (sn)", "titanio (ti)"
]

PRIMARY_PROPERTIES = [
    "número básico", "numero basico", "bn", "tbn",
    "número ácido", "numero acido", "an", "tan",
    "viscosidad a 40", "viscosidad a 100", "visc@40", "visc@100",
    "oxidación", "oxidacion", "nitración", "nitracion", "sulfatación", "sulfatacion",
    "oxidation", "nitration", "sulfation", "ab/cm"
]

PRIMARY_CONTAM = [
    "agua", "water",
    "hollín", "hollin", "soot",
    "silicio", "silicon", "si (",
    "sodio", "sodium", "na (",
    "potasio", "potassium", "k (",
    "glicol", "glycol",
    "dilución por combustible", "dilucion por combustible", "fuel dilut", "fuel dilution",
    "conteo partículas", "conteo particulas", "particle count", "iso 4406", "cleanliness"
]

PRIMARY_ADDITIVES = [
    "calcio", "ca (", "magnesio", "mg (", "zinc", "zn (", "fósforo", "fosforo", "p (", "boro", "b ("
]

# Columnas que jamás deberían entrar a “variables lógicas”
# (en tu captura estas se estaban colando)
NAME_EXCLUDE_PATTERNS = [
    "fecha", "ingreso", "recepcion", "recepción", "descriptor", "descrip", "descripcion", "descripción",
    "observ", "coment", "cliente", "operacion", "operación", "lubricante", "producto", "estado reporte",
    "componente",  # OJO: la llave ya la usamos aparte
]

def categorize_variable(var_name: str) -> str:
    v = normalize_name(var_name)

    # Exclusiones fuertes por nombre
    if any(p in v for p in NAME_EXCLUDE_PATTERNS):
        return "No usar"

    # Estricto por tipo de variable
    if any(k in v for k in PRIMARY_WEAR):
        return "Desgaste"

    if any(k in v for k in PRIMARY_PROPERTIES):
        return "Propiedades del lubricante"

    if any(k in v for k in PRIMARY_CONTAM):
        return "Contaminantes"

    if any(k in v for k in PRIMARY_ADDITIVES):
        return "Aditivos"

    return "Otras variables"

def build_candidates(df: pd.DataFrame) -> tuple[list, dict]:
    cols = list(df.columns)

    # Mapeo variable -> columna estado
    estado_cols = [c for c in cols if " - Estado" in str(c)]
    var_to_estado = {}
    for c_estado in estado_cols:
        base = str(c_estado).replace(" - Estado ", " - Estado").replace(" - Estado", "").strip()
        if base in df.columns:
            var_to_estado[base] = c_estado

    # Columnas de control que nunca deben ser variables
    hard_exclude = {
        "COMPONENTE", "FECHA_INFORME",
        "NOMBRE_CLIENTE", "CLIENTE",
        "NOMBRE_OPERACION", "OPERACION",
        "TIPO_EQUIPO",
        "PRODUCTO", "Tested Lubricant",
        "ESTADO_REPORTE", "Report Status",
        "CORRELATIVO", "N_MUESTRA",
        "Sample Bottle ID", "Asset ID", "EQUIPO",
        "DESCRIPTOR_COMPONENTE"
    }

    candidates = []
    for c in cols:
        cs = str(c)

        if c in hard_exclude:
            continue
        if " - Estado" in cs:
            continue

        # Exclusión fuerte por nombre (evita FECHA_MUESTREO, FECHA_RECEPCION, etc.)
        if any(p in normalize_name(cs) for p in NAME_EXCLUDE_PATTERNS):
            continue

        # Excluir columnas datetime
        if pd.api.types.is_datetime64_any_dtype(df[c]):
            continue

        # Aceptar numéricas puras
        if pd.api.types.is_numeric_dtype(df[c]):
            candidates.append(c)
            continue

        # Aceptar convertibles a numérico si la mayoría parecen números
        sample = df[c].dropna().astype(str).head(120)
        if sample.empty:
            continue

        # Si la mayoría no contiene dígitos, es texto -> fuera
        digit_rate = sample.str.contains(r"\d", regex=True).mean()
        if digit_rate < 0.6:
            continue

        # Si casi todo es único y largo, suele ser texto/ID -> fuera
        uniq_rate = sample.nunique() / max(1, len(sample))
        avg_len = sample.str.len().mean()
        if uniq_rate > 0.95 and avg_len > 18:
            continue

        candidates.append(c)

    # Dedupe conservando orden
    seen = set()
    candidates = [x for x in candidates if not (x in seen or seen.add(x))]
    return candidates, var_to_estado

def get_top_by_category(all_vars: list, max_each: int = 15) -> dict:
    """Categoriza y devuelve máximo 15 por categoría (estricto)."""
    cats = {
        "Desgaste": [],
        "Propiedades del lubricante": [],
        "Contaminantes": [],
        "Aditivos": [],
        "Otras variables": []
    }

    for v in all_vars:
        cat = categorize_variable(v)
        if cat == "No usar":
            continue
        if cat in cats and cat != "Otras variables":
            if len(cats[cat]) < max_each:
                cats[cat].append(v)

    # Otras (lista corta por usabilidad)
    for v in all_vars:
        cat = categorize_variable(v)
        if cat == "Otras variables":
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
            Muestras=("COMPONENTE", "size"),
            Primera_fecha=("FECHA_INFORME", "min"),
            Última_fecha=("FECHA_INFORME", "max"),
        )
        .reset_index()
        .sort_values("Muestras", ascending=False)
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
st.caption("Se muestran las variables principales por categoría, sin incluir fechas ni descriptores.")

buscador = st.text_input("Buscar variable", value="").strip().lower()

cats = get_top_by_category(logic_candidates, max_each=15)
if buscador:
    for k in list(cats.keys()):
        cats[k] = [v for v in cats[k] if buscador in normalize_name(v)]

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
                "Operación": None,
                "Tipo de equipo": None,
                "Lubricante": None,
                "Componente": comp,
                "Categoría": categorize_variable(v),
                "Variable": v,
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

            if col_op and col_op in g_comp.columns and g_comp[col_op].notna().any():
                fila["Operación"] = g_comp[col_op].dropna().iloc[0]
            if col_tipo and col_tipo in g_comp.columns and g_comp[col_tipo].notna().any():
                fila["Tipo de equipo"] = g_comp[col_tipo].dropna().iloc[0]
            if col_lub and col_lub in g_comp.columns and g_comp[col_lub].notna().any():
                fila["Lubricante"] = g_comp[col_lub].dropna().iloc[0]

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

    archivo_excel = to_excel_bytes(resultados_vista[columnas], sheet_name="Límites")
    st.download_button(
        "Descargar Excel",
        data=archivo_excel,
        file_name="limites_condenatorios_por_componente.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Ajusta filtros, selecciona componentes y variables, y luego calcula los límites.")









