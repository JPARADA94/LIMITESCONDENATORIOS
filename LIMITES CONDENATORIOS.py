# limites_condenatorios.py
# Autor: Javier Parada
# Entrada: Excel TAL CUAL viene de SmartAssintence (formato ARCHIVO 2)
# Llave de análisis: COMPONENTE
# Salida: Excel con límites de Precaución y Condenatorio (Alerta)

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# =========================
# Configuración
# =========================
st.set_page_config(page_title="Límites por componente", layout="wide")
st.title("Límites por componente")

st.markdown("""
Esta herramienta calcula límites de **precaución** y **condenatorio (alerta)** por componente con base en el histórico del archivo exportado desde SmartAssintence.
Puedes aplicar filtros opcionales por operación, tipo de equipo, lubricante y fechas. Luego eliges el modo de cálculo: límites por componente o un único límite
mezclando dos componentes. El cálculo se realiza únicamente con las variables definidas en la guía del reporte. Si existe la columna de estado de la variable,
puedes excluir del cálculo los registros en alerta. Finalmente, descargas los resultados en Excel.
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

def first_non_null(series: pd.Series):
    s = series.dropna()
    return s.iloc[0] if not s.empty else None

def build_var_to_status(df: pd.DataFrame) -> dict:
    """
    En ARCHIVO 2 las columnas de estado suelen venir como:
    'COBRE (CU) - 25 - Estado '
    """
    mapping = {}
    cols = set(df.columns)
    for c in df.columns:
        if isinstance(c, str) and " - Estado" in c:
            base = c.replace(" - Estado ", "").replace(" - Estado", "").strip()
            if base in cols:
                mapping[base] = c
    return mapping

# =========================
# Variables fijas: SOLO las de tu guía, con nombre EXACTO del export
# =========================
GUIDE_VARS = {
    "Desgaste": [
        "PLATA (AG) - 19",
        "ALUMINIO (AL) - 20",
        "CROMO (CR) - 24",
        "COBRE (CU) - 25",
        "HIERRO (FE) - 26",
        "ÍNDICE PQ (PQI) - 3",
        "NÍQUEL (NI) - 32",
        "PLOMO (PB) - 35",
        "ESTAÑO (SN) - 37",
        "TITANIO (TI) - 38",
    ],
    "Propiedades del lubricante": [
        "NÚMERO BÁSICO (BN) - 12",
        "VISCOSIDAD A 100 °C - 13",
        "NÚMERO ÁCIDO (AN) - 43",
        "OXIDACIÓN - 80",
        "NITRACIÓN - 82",
    ],
    "Contaminantes": [
        "CADMIO (CD) - 23",
        "POTASIO (K) - 27",
        "MANGANESO (MN) - 29",
        "SODIO (NA) - 31",
        "SILICIO (SI) - 36",
        "VANADIO (V) - 39",
        "HOLLÍN - 79",
        "AGUA (IR) - 74",
    ],
    "Aditivos": [
        "BORO (B) - 18",
        "BARIO (BA) - 21",
        "CALCIO (CA) - 22",
        "MAGNESIO (MG) - 28",
        "MOLIBDENO (MO) - 30",
        "FÓSFORO (P) - 34",
        "ZINC (ZN) - 40",
    ],
}

# =========================
# Carga de archivo
# =========================
archivo = st.file_uploader("Cargar archivo Excel de SmartAssintence", type=["xlsx"])
if not archivo:
    st.stop()

df = load_excel(archivo).copy()

required_cols = ["COMPONENTE", "FECHA_INFORME"]
missing = [c for c in required_cols if c not in df.columns]
if missing:
    st.error(f"Faltan columnas requeridas: {missing}")
    st.stop()

df["FECHA_INFORME"] = pd.to_datetime(df["FECHA_INFORME"], errors="coerce")
if df["FECHA_INFORME"].isna().all():
    st.error("La columna FECHA_INFORME no contiene fechas válidas.")
    st.stop()

col_operacion = "NOMBRE_OPERACION" if "NOMBRE_OPERACION" in df.columns else None
col_tipo_equipo = "TIPO_EQUIPO" if "TIPO_EQUIPO" in df.columns else None
col_lubricante = "PRODUCTO" if "PRODUCTO" in df.columns else None

var_to_status = build_var_to_status(df)

# Variables disponibles (según guía) que sí existen en el archivo
available = set(df.columns)
vars_by_cat = {cat: [v for v in vlist if v in available] for cat, vlist in GUIDE_VARS.items()}
total_exist = sum(len(v) for v in vars_by_cat.values())
if total_exist == 0:
    st.error(
        "No se encontró ninguna variable de la guía en el archivo. "
        "Verifica que sea el export estándar de SmartAssintence (formato ARCHIVO 2)."
    )
    st.stop()

with st.expander("Revisión de variables encontradas", expanded=False):
    for cat in vars_by_cat:
        st.write(f"**{cat}** — encontradas: {len(vars_by_cat[cat])} / {len(GUIDE_VARS[cat])}")
        miss = [v for v in GUIDE_VARS[cat] if v not in available]
        if miss:
            st.caption("No encontradas (puede variar si el export cambió):")
            st.write(miss)

# =========================
# 1) Filtros opcionales
# =========================
st.markdown("## 1. Filtros de análisis")

df_f = df.copy()

usar_fechas = st.checkbox("Filtrar por fechas", value=False)
if usar_fechas:
    fmin = df_f["FECHA_INFORME"].min()
    fmax = df_f["FECHA_INFORME"].max()
    if pd.notna(fmin) and pd.notna(fmax):
        ini, fin = st.date_input(
            "Rango de fechas",
            value=[fmin.date(), fmax.date()],
            min_value=fmin.date(),
            max_value=fmax.date()
        )
        df_f = df_f[(df_f["FECHA_INFORME"] >= pd.to_datetime(ini)) &
                    (df_f["FECHA_INFORME"] <= pd.to_datetime(fin))].copy()

c1, c2, c3 = st.columns(3)
with c1:
    usar_operacion = st.checkbox("Filtrar por operación", value=False, disabled=(col_operacion is None))
    if usar_operacion and col_operacion:
        ops = st.multiselect("Operación", sorted(df_f[col_operacion].dropna().unique()))
        if ops:
            df_f = df_f[df_f[col_operacion].isin(ops)].copy()

with c2:
    usar_tipo = st.checkbox("Filtrar por tipo de equipo", value=False, disabled=(col_tipo_equipo is None))
    if usar_tipo and col_tipo_equipo:
        tipos = st.multiselect("Tipo de equipo", sorted(df_f[col_tipo_equipo].dropna().unique()))
        if tipos:
            df_f = df_f[df_f[col_tipo_equipo].isin(tipos)].copy()

with c3:
    usar_lub = st.checkbox("Filtrar por lubricante", value=False, disabled=(col_lubricante is None))
    if usar_lub and col_lubricante:
        lubs = st.multiselect("Lubricante", sorted(df_f[col_lubricante].dropna().unique()))
        if lubs:
            df_f = df_f[df_f[col_lubricante].isin(lubs)].copy()

if df_f.empty:
    st.warning("No hay datos con los filtros seleccionados.")
    st.stop()

# =========================
# 2) Inventario
# =========================
st.markdown("## 2. Inventario de componentes")

group_cols = ["COMPONENTE"]
if col_operacion: group_cols.append(col_operacion)
if col_tipo_equipo: group_cols.append(col_tipo_equipo)
if col_lubricante: group_cols.append(col_lubricante)

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
    options=["Límites por componente", "Límite único mezclando dos componentes"],
    index=0
)

# =========================
# 4) Selección de componentes
# =========================
st.markdown("## 4. Selección de componentes")

componentes = sorted(df_f["COMPONENTE"].dropna().astype(str).unique())
if not componentes:
    st.warning("No hay componentes disponibles con los filtros actuales.")
    st.stop()

if modo == "Límites por componente":
    comps_sel = st.multiselect("Componentes a analizar", options=componentes, default=componentes[:1])
    if not comps_sel:
        st.warning("Selecciona al menos un componente.")
        st.stop()
    df_calc = df_f[df_f["COMPONENTE"].astype(str).isin(set(map(str, comps_sel)))].copy()
else:
    a, b = st.columns(2)
    with a:
        comp_a = st.selectbox("Componente A", options=componentes, index=0)
    with b:
        comp_b = st.selectbox("Componente B", options=componentes, index=1 if len(componentes) > 1 else 0)

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
# 5) Variables
# =========================
st.markdown("## 5. Variables para el cálculo")

excluir_alerta = st.checkbox("Excluir registros en alerta cuando exista el estado de la variable", value=True)

if "vars_checked" not in st.session_state:
    st.session_state["vars_checked"] = set()

def render_cat(title: str, vars_list: list, expanded: bool):
    if not vars_list:
        return
    with st.expander(title, expanded=expanded):
        cols = st.columns(3)
        for i, v in enumerate(vars_list):
            with cols[i % 3]:
                cur = v in st.session_state["vars_checked"]
                val = st.checkbox(v, value=cur, key=f"chk_{title}_{v}")
                if val:
                    st.session_state["vars_checked"].add(v)
                else:
                    st.session_state["vars_checked"].discard(v)

render_cat("Desgaste", vars_by_cat["Desgaste"], expanded=True)
render_cat("Propiedades del lubricante", vars_by_cat["Propiedades del lubricante"], expanded=True)
render_cat("Contaminantes", vars_by_cat["Contaminantes"], expanded=True)
render_cat("Aditivos", vars_by_cat["Aditivos"], expanded=False)

vars_sel = sorted(list(st.session_state["vars_checked"]))
if not vars_sel:
    st.warning("Selecciona al menos una variable.")
    st.stop()

st.success(f"Variables seleccionadas: {len(vars_sel)}")

for v in vars_sel:
    df_calc[v] = convert_numeric(df_calc[v])

# =========================
# 6) Parámetros de cálculo
# =========================
st.markdown("## 6. Parámetros de cálculo")

min_n = st.number_input("Mínimo de datos válidos para calcular", min_value=2, value=3, step=1)

cA, cB = st.columns(2)
with cA:
    n_switch = st.number_input("Umbral para usar percentiles", min_value=3, value=10, step=1)
    p_prec = st.slider("Percentil de precaución", 50, 99, 90, 1)
    p_alert = st.slider("Percentil de condenatorio (alerta)", 50, 99, 95, 1)
with cB:
    k_prec = st.number_input("Factor para precaución con histórico corto", min_value=0.0, value=2.0, step=0.5)
    k_alert = st.number_input("Factor para condenatorio con histórico corto", min_value=0.0, value=3.0, step=0.5)

def apply_status_filter(df_in: pd.DataFrame, var: str) -> pd.Series:
    s = df_in[var]
    if not excluir_alerta:
        return s
    status_col = var_to_status.get(var)
    if not status_col or status_col not in df_in.columns:
        return s
    estado = df_in[status_col].astype(str).str.strip().str.upper()
    return s[estado != "ALERTA"]

def calc_limits(series: pd.Series) -> dict:
    s = series.dropna()
    n = int(len(s))
    if n < min_n:
        return {
            "n": n,
            "metodo": "Insuficiente",
            "prec": np.nan,
            "alert": np.nan,
            "mean": np.nan,
            "std": np.nan,
            "median": np.nan
        }

    mean = float(s.mean())
    std = float(s.std(ddof=1)) if n > 1 else 0.0
    median = float(s.median())

    if n >= n_switch:
        prec = float(s.quantile(p_prec / 100))
        alert = float(s.quantile(p_alert / 100))
        metodo = f"Percentiles P{p_prec} y P{p_alert}"
    else:
        prec = mean + (k_prec * std)
        alert = mean + (k_alert * std)
        metodo = "Media y desviación"

    return {
        "n": n,
        "metodo": metodo,
        "prec": prec,
        "alert": alert,
        "mean": mean,
        "std": std,
        "median": median
    }

def get_category(var: str) -> str:
    for cat, lst in vars_by_cat.items():
        if var in lst:
            return cat
    return "Sin categoría"

# =========================
# 7) Resultados y descarga
# =========================
st.markdown("## 7. Resultados y descarga")

if st.button("Calcular límites"):
    filas = []

    if modo == "Límites por componente":
        for comp, g in df_calc.groupby(df_calc["COMPONENTE"].astype(str)):
            for var in vars_sel:
                serie = apply_status_filter(g, var)
                out = calc_limits(serie)

                filas.append({
                    "Operación": first_non_null(g[col_operacion]) if col_operacion else None,
                    "Tipo de equipo": first_non_null(g[col_tipo_equipo]) if col_tipo_equipo else None,
                    "Lubricante": first_non_null(g[col_lubricante]) if col_lubricante else None,
                    "Componente": comp,
                    "Categoría": get_category(var),
                    "Variable": var,
                    "Datos válidos": out["n"],
                    "Método": out["metodo"],
                    "Límite de precaución": out["prec"],
                    "Límite condenatorio": out["alert"],
                    "Promedio": out["mean"],
                    "Desviación estándar": out["std"],
                    "Mediana": out["median"],
                    "Primera fecha": g["FECHA_INFORME"].min(),
                    "Última fecha": g["FECHA_INFORME"].max(),
                    "Excluye alerta": "Sí" if excluir_alerta else "No",
                })

    else:
        comps_mix = sorted(df_calc["COMPONENTE"].dropna().astype(str).unique().tolist())
        etiqueta = " + ".join(comps_mix) if comps_mix else "Mezcla"

        for var in vars_sel:
            serie = apply_status_filter(df_calc, var)
            out = calc_limits(serie)

            filas.append({
                "Componentes mezclados": etiqueta,
                "Categoría": get_category(var),
                "Variable": var,
                "Datos válidos": out["n"],
                "Método": out["metodo"],
                "Límite de precaución": out["prec"],
                "Límite condenatorio": out["alert"],
                "Promedio": out["mean"],
                "Desviación estándar": out["std"],
                "Mediana": out["median"],
                "Primera fecha": df_calc["FECHA_INFORME"].min(),
                "Última fecha": df_calc["FECHA_INFORME"].max(),
                "Excluye alerta": "Sí" if excluir_alerta else "No",
            })

    resultados = pd.DataFrame(filas)

    dec = st.number_input("Decimales para visualización", min_value=0, value=0, step=1)
    vista = resultados.copy()
    for c in ["Límite de precaución", "Límite condenatorio", "Promedio", "Desviación estándar", "Mediana"]:
        if c in vista.columns:
            vista[c] = pd.to_numeric(vista[c], errors="coerce").round(dec)

    if modo == "Límites por componente":
        orden = [
            "Operación", "Tipo de equipo", "Lubricante",
            "Componente", "Categoría", "Variable",
            "Datos válidos", "Método",
            "Límite de precaución", "Límite condenatorio",
            "Promedio", "Desviación estándar", "Mediana",
            "Primera fecha", "Última fecha", "Excluye alerta"
        ]
    else:
        orden = [
            "Componentes mezclados", "Categoría", "Variable",
            "Datos válidos", "Método",
            "Límite de precaución", "Límite condenatorio",
            "Promedio", "Desviación estándar", "Mediana",
            "Primera fecha", "Última fecha", "Excluye alerta"
        ]

    columnas = [c for c in orden if c in vista.columns]
    st.dataframe(vista[columnas], use_container_width=True)

    bytes_xlsx = to_excel_bytes(vista[columnas], sheet_name="Límites")
    filename = "limites_por_componente.xlsx" if modo == "Límites por componente" else "limites_mezcla_componentes.xlsx"

    st.download_button(
        "Descargar Excel",
        data=bytes_xlsx,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Aplica filtros, selecciona el modo, elige componentes y variables, y luego calcula los límites.")







