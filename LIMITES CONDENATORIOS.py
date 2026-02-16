# limites_condenatorios.py
# Autor: Javier Parada
# Entrada: Excel TAL CUAL viene de SmartAssintence
#
# Requisitos implementados:
# 1) Solo se trabaja con las variables de tu guía (Desgaste, Propiedades del lubricante, Contaminantes, Aditivos)
#    y se excluyen "Periodo uso aceite" y "Unidad uso aceite".
# 2) Opción para mezclar el histórico de dos componentes (equipos) y calcular un único límite de precaución.
# 3) Se eliminan los botones de limpieza (no aparecen).
# 4) Mantiene filtros opcionales por operación, tipo de equipo, lubricante y fechas.
# 5) Descarga el resultado a Excel.

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re

# =========================
# Configuración
# =========================
st.set_page_config(page_title="Límites por Componente", layout="wide")
st.title("Límites por Componente")

st.markdown("""
Esta herramienta calcula límites de precaución por componente con base en el histórico del archivo exportado desde SmartAssintence.
Primero puedes aplicar filtros opcionales por operación, tipo de equipo, lubricante y fechas. Luego seleccionas el modo de cálculo:
límites individuales por componente o un límite único mezclando dos componentes. En ambos casos, el cálculo se realiza únicamente con las
variables definidas en la guía del reporte. Finalmente, descargas el consolidado en Excel.
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
    txt = str(s).strip().lower()
    txt = re.sub(r"\s+", " ", txt)
    txt = re.sub(r"\s*-\s*\d+\s*$", "", txt)  # quita sufijos tipo " - 20"
    return txt

def convert_numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype(str).str.replace(r"[^0-9\.\-]", "", regex=True),
        errors="coerce"
    )

def safe_first_non_null(s: pd.Series):
    s2 = s.dropna()
    return s2.iloc[0] if not s2.empty else None

# =========================
# Variables EXACTAS de la guía (lista fija)
# =========================
# Nota: se comparan con normalize_name() para soportar variantes como "COBRE (CU) - 25"

GUIDE_GROUPS = {
    "Desgaste": [
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
    ],
    "Propiedades del lubricante": [
        "Número Básico (BN) (mg KOH/g) ASTM D2896",
        "Viscosidad a 100 °C (mm²/s) ASTM D445",
        "Número Ácido (AN) (mg KOH/g) ASTM D664",
        "Oxidación (Abs/cm) ASTM D7414",
        "Nitración (Abs/cm) ASTM D7624",
    ],
    "Contaminantes": [
        "Cadmio (Cd) (mg/kg) ASTM D5185",
        "Potasio (K) (mg/kg) ASTM D5185",
        "Manganeso (Mn) (mg/kg) ASTM D5185",
        "Sodio (Na) (mg/kg) ASTM D5185",
        "Silicio (Si) (mg/kg) ASTM D5185",
        "Vanadio (V) (mg/kg) ASTM D5185",
        "Hollín (% w/w) ASTM D7844",
        "Agua (IR) (% v/v) ASTM E2412",
    ],
    "Aditivos": [
        "Boro (B) (mg/kg) ASTM D5185",
        "Bario (Ba) (mg/kg) ASTM D5185",
        "Calcio (Ca) (mg/kg) ASTM D5185",
        "Magnesio (Mg) (mg/kg) ASTM D5185",
        "Molibdeno (Mo) (mg/kg) ASTM D5185",
        "Fósforo (P) (mg/kg) ASTM D5185",
        "Zinc (Zn) (mg/kg) ASTM D5185",
    ],
}

EXCLUDE_ALWAYS = {"Periodo uso aceite", "Unidad uso aceite"}

def build_var_status_map(df: pd.DataFrame) -> dict:
    """Mapea variable -> columna de estado si existe '<variable> - Estado'."""
    var_to_estado = {}
    for c in df.columns:
        if " - Estado" in str(c):
            base = str(c).replace(" - Estado ", " - Estado").replace(" - Estado", "").strip()
            if base in df.columns:
                var_to_estado[base] = c
    return var_to_estado

def pick_existing_columns(df: pd.DataFrame) -> dict:
    """
    Devuelve dict {categoria: [columnas_reales_en_df]} usando coincidencia por nombre normalizado,
    soportando sufijos tipo ' - 20'.
    """
    df_cols = list(df.columns)
    norm_to_real = {normalize_name(c): c for c in df_cols}

    # Además, intentamos coincidencia "contiene" para robustez cuando el nombre viene con extras.
    def find_best_match(target: str):
        t = normalize_name(target)
        if t in norm_to_real:
            return norm_to_real[t]
        # búsqueda por contiene (segura)
        candidates = [c for c in df_cols if t in normalize_name(c)]
        return candidates[0] if candidates else None

    out = {}
    for cat, targets in GUIDE_GROUPS.items():
        out[cat] = []
        for t in targets:
            m = find_best_match(t)
            if m and m not in out[cat] and normalize_name(m) not in [normalize_name(x) for x in EXCLUDE_ALWAYS]:
                out[cat].append(m)
    return out

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

var_to_estado = build_var_status_map(df)

# Variables disponibles según guía
vars_by_cat = pick_existing_columns(df)
all_guide_vars = [v for lst in vars_by_cat.values() for v in lst]

if not all_guide_vars:
    st.error("No encontré en tu Excel ninguna de las variables de la guía. Revisa los nombres de columnas.")
    st.stop()

# =========================
# 1) Filtros opcionales
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
# 2) Inventario de componentes
# =========================
st.markdown("## 2. Inventario de componentes")
st.caption("Histórico disponible por componente después de aplicar filtros.")

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
# 3) Modo de cálculo
# =========================
st.markdown("## 3. Modo de cálculo")

modo = st.radio(
    "Selecciona el modo",
    options=[
        "Límites individuales por componente",
        "Límite único mezclando dos componentes"
    ],
    index=0
)

# =========================
# 4) Selección de componentes
# =========================
st.markdown("## 4. Selección de componentes")

componentes_disponibles = sorted(df_f["COMPONENTE"].dropna().astype(str).unique())

if modo == "Límites individuales por componente":
    componentes_sel = st.multiselect(
        "Componentes a analizar",
        options=componentes_disponibles,
        default=componentes_disponibles[:1] if componentes_disponibles else []
    )
    if not componentes_sel:
        st.warning("Selecciona al menos un componente.")
        st.stop()
    df_calc = df_f[df_f["COMPONENTE"].astype(str).isin(set(map(str, componentes_sel)))].copy()

else:
    c1, c2 = st.columns(2)
    with c1:
        comp_a = st.selectbox("Componente A", options=componentes_disponibles, index=0 if componentes_disponibles else None)
    with c2:
        comp_b = st.selectbox("Componente B", options=componentes_disponibles, index=1 if len(componentes_disponibles) > 1 else 0)
    if not comp_a or not comp_b:
        st.warning("Selecciona dos componentes.")
        st.stop()
    if str(comp_a) == str(comp_b):
        st.warning("Selecciona dos componentes diferentes.")
        st.stop()

    df_a = df_f[df_f["COMPONENTE"].astype(str) == str(comp_a)].copy()
    df_b = df_f[df_f["COMPONENTE"].astype(str) == str(comp_b)].copy()
    df_calc = pd.concat([df_a, df_b], ignore_index=True)

if df_calc.empty:
    st.warning("No hay registros para la selección actual.")
    st.stop()

# =========================
# 5) Selección de variables (solo guía)
# =========================
st.markdown("## 5. Variables para el cálculo")
st.caption("Solo se muestran variables de la guía del reporte.")

excluir_alertas = st.checkbox("Excluir registros en alerta cuando exista estado de la variable", value=True)

# selección por categorías (chulos)
if "vars_checked" not in st.session_state:
    st.session_state["vars_checked"] = set()

def render_category(cat_title: str, var_list: list, expanded: bool):
    if not var_list:
        return
    with st.expander(cat_title, expanded=expanded):
        cols_ui = st.columns(3)
        for i, v in enumerate(var_list):
            col = cols_ui[i % 3]
            with col:
                checked = v in st.session_state["vars_checked"]
                new_val = st.checkbox(str(v), value=checked, key=f"chk_{cat_title}_{v}")
                if new_val:
                    st.session_state["vars_checked"].add(v)
                else:
                    st.session_state["vars_checked"].discard(v)

render_category("Desgaste", vars_by_cat["Desgaste"][:15], expanded=True)
render_category("Propiedades del lubricante", vars_by_cat["Propiedades del lubricante"][:15], expanded=True)
render_category("Contaminantes", vars_by_cat["Contaminantes"][:15], expanded=True)
render_category("Aditivos", vars_by_cat["Aditivos"][:15], expanded=False)

vars_sel = sorted(list(st.session_state["vars_checked"]))
if not vars_sel:
    st.warning("Selecciona al menos una variable.")
    st.stop()

st.success(f"Variables seleccionadas: {len(vars_sel)}")

# Convertir variables seleccionadas a numérico
for v in vars_sel:
    df_calc[v] = convert_numeric(df_calc[v])

# =========================
# 6) Parámetros de cálculo
# =========================
st.markdown("## 6. Parámetros de cálculo")

min_n = st.number_input("Mínimo de datos válidos para calcular", min_value=2, value=3, step=1)
n_switch = st.number_input("Umbral para usar percentiles", min_value=3, value=10, step=1)
p_prec = st.slider("Percentil para precaución", 50, 99, 90, 1)

# Nota: En tu solicitud pides “solo un límite de precaución” al mezclar.
# Para modo individual dejamos también precaución únicamente, para mantener consistencia.

def apply_estado_filter(g: pd.DataFrame, v: str) -> pd.Series:
    s = g[v]
    estado_col = var_to_estado.get(v)
    if not excluir_alertas or not estado_col or estado_col not in g.columns:
        return s
    est = g[estado_col].astype(str).str.strip().str.upper()
    mask = (est != "ALERTA")
    return s[mask]

def calc_precaution(series: pd.Series) -> dict:
    s = series.dropna()
    n = int(len(s))
    if n < min_n:
        return {"n": n, "metodo": "Insuficiente", "prec": np.nan, "mean": np.nan, "std": np.nan, "median": np.nan}
    if n >= n_switch:
        prec = float(s.quantile(p_prec / 100))
        metodo = f"Percentil P{p_prec}"
    else:
        mean = float(s.mean())
        std = float(s.std(ddof=1)) if n > 1 else 0.0
        prec = mean + (2.0 * std)  # fijo a 2 sigma para histórico corto (puedes cambiarlo si quieres)
        metodo = "Media y desviación"
    return {"n": n, "metodo": metodo, "prec": prec, "mean": float(s.mean()), "std": float(s.std(ddof=1)) if n > 1 else 0.0, "median": float(s.median())}

def get_category(var_name: str) -> str:
    for cat, lst in vars_by_cat.items():
        if var_name in lst:
            return cat
    return "Otra"

# =========================
# 7) Calcular y exportar
# =========================
st.markdown("## 7. Resultados y descarga")

if st.button("Calcular límite de precaución"):
    filas = []

    if modo == "Límites individuales por componente":
        for comp, g_comp in df_calc.groupby(df_calc["COMPONENTE"].astype(str)):
            for v in vars_sel:
                serie = apply_estado_filter(g_comp, v)
                out = calc_precaution(serie)

                fila = {
                    "Operación": safe_first_non_null(g_comp[col_op]) if col_op else None,
                    "Tipo de equipo": safe_first_non_null(g_comp[col_tipo]) if col_tipo else None,
                    "Lubricante": safe_first_non_null(g_comp[col_lub]) if col_lub else None,
                    "Componente": comp,
                    "Categoría": get_category(v),
                    "Variable": v,
                    "Datos válidos": out["n"],
                    "Método": out["metodo"],
                    "Límite de precaución": out["prec"],
                    "Promedio": out["mean"],
                    "Desviación estándar": out["std"],
                    "Mediana": out["median"],
                    "Primera fecha": g_comp["FECHA_INFORME"].min(),
                    "Última fecha": g_comp["FECHA_INFORME"].max(),
                    "Excluye alerta": "Sí" if excluir_alertas else "No",
                }
                filas.append(fila)

    else:
        # Mezcla de dos componentes: un solo resultado por variable
        # Nota: se reportan también los dos componentes mezclados
        comps_in_mix = sorted(df_calc["COMPONENTE"].dropna().astype(str).unique().tolist())
        etiqueta = " + ".join(comps_in_mix) if comps_in_mix else "Mezcla"

        for v in vars_sel:
            serie = apply_estado_filter(df_calc, v)
            out = calc_precaution(serie)

            fila = {
                "Componentes mezclados": etiqueta,
                "Categoría": get_category(v),
                "Variable": v,
                "Datos válidos": out["n"],
                "Método": out["metodo"],
                "Límite de precaución": out["prec"],
                "Promedio": out["mean"],
                "Desviación estándar": out["std"],
                "Mediana": out["median"],
                "Primera fecha": df_calc["FECHA_INFORME"].min(),
                "Última fecha": df_calc["FECHA_INFORME"].max(),
                "Excluye alerta": "Sí" if excluir_alertas else "No",
            }
            filas.append(fila)

    resultados = pd.DataFrame(filas)

    # Visualización con redondeo
    dec = st.number_input("Decimales para visualización", min_value=0, value=0, step=1)
    resultados_vista = resultados.copy()

    num_cols = ["Límite de precaución", "Promedio", "Desviación estándar", "Mediana"]
    for c in num_cols:
        if c in resultados_vista.columns:
            resultados_vista[c] = pd.to_numeric(resultados_vista[c], errors="coerce").round(dec)

    # Orden de columnas
    if modo == "Límites individuales por componente":
        orden = [
            "Operación", "Tipo de equipo", "Lubricante",
            "Componente", "Categoría", "Variable",
            "Datos válidos", "Método", "Límite de precaución",
            "Promedio", "Desviación estándar", "Mediana",
            "Primera fecha", "Última fecha", "Excluye alerta"
        ]
    else:
        orden = [
            "Componentes mezclados", "Categoría", "Variable",
            "Datos válidos", "Método", "Límite de precaución",
            "Promedio", "Desviación estándar", "Mediana",
            "Primera fecha", "Última fecha", "Excluye alerta"
        ]

    columnas = [c for c in orden if c in resultados_vista.columns]
    st.dataframe(resultados_vista[columnas], use_container_width=True)

    archivo_excel = to_excel_bytes(resultados_vista[columnas], sheet_name="Límites")
    st.download_button(
        "Descargar Excel",
        data=archivo_excel,
        file_name="limite_precaucion_por_componente.xlsx" if modo == "Límites individuales por componente" else "limite_precaucion_mezcla_componentes.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Selecciona filtros, modo, componentes y variables, y luego calcula el límite de precaución.")











