# app.py
# Tablero Posventa — Macro → Micro + Cierre de Mes GAP
# Compatible Streamlit Cloud

import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

try:
    import gdown
except Exception:
    gdown = None


# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="Tablero Posventa",
    page_icon="🚘",
    layout="wide"
)

DRIVE_FILE_ID = "191JKfQWj3yehcnisKTPDs_KpWaOTyslhQ0g273Xvzjc"
LOCAL_FILE = "base_posventa.xlsx"


# ============================================================
# ESTILOS
# ============================================================

st.markdown("""
<style>
.main {background-color:#fafafa;}
.block-container {padding-top:1.5rem;}
.metric-card {
    background:white;
    padding:18px;
    border-radius:18px;
    box-shadow:0 2px 10px rgba(0,0,0,0.06);
    border:1px solid #eee;
}
.big-number {
    font-size:28px;
    font-weight:700;
}
.small-label {
    color:#666;
    font-size:14px;
}
.status-ok {color:#16803c;font-weight:700;}
.status-warn {color:#b8860b;font-weight:700;}
.status-bad {color:#b3261e;font-weight:700;}
</style>
""", unsafe_allow_html=True)


# ============================================================
# HELPERS
# ============================================================

def to_number(x):
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float, np.number)):
        return float(x)

    s = str(x).strip()
    s = s.replace("$", "").replace("%", "").replace(" ", "")
    s = s.replace("\xa0", "")

    # Formato argentino: 1.234.567,89
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")

    try:
        return float(s)
    except Exception:
        return 0.0


def format_money(x):
    try:
        return f"$ {x:,.0f}".replace(",", ".")
    except Exception:
        return "$ 0"


def format_qty(x):
    try:
        return f"{x:,.0f}".replace(",", ".")
    except Exception:
        return "0"


def format_pct(x):
    if pd.isna(x) or np.isinf(x):
        return "-"
    return f"{x:.1%}"


def week_number(value):
    if pd.isna(value):
        return np.nan
    m = re.search(r"(\d+)", str(value))
    return int(m.group(1)) if m else np.nan


def classify_estado(cump):
    if pd.isna(cump):
        return "⚪ Sin objetivo"
    if cump >= 1:
        return "🟢 Cumplido"
    if cump >= 0.90:
        return "🟡 Cerca"
    return "🔴 En riesgo"


def detect_col(df, options):
    cols = list(df.columns)
    normalized = {str(c).strip().lower(): c for c in cols}

    for op in options:
        key = op.strip().lower()
        if key in normalized:
            return normalized[key]

    for c in cols:
        c_low = str(c).strip().lower()
        for op in options:
            if op.strip().lower() in c_low:
                return c

    return None


# ============================================================
# DATA LOAD
# ============================================================

@st.cache_data(show_spinner=True)
def load_excel():
    if gdown is not None:
        try:
            url = f"https://drive.google.com/uc?id={DRIVE_FILE_ID}"
            gdown.download(url, LOCAL_FILE, quiet=True)
        except Exception:
            pass

    try:
        return pd.read_excel(LOCAL_FILE)
    except Exception as e:
        st.error("No pude cargar base_posventa.xlsx. Verificá que el archivo exista o que el ID de Drive sea correcto.")
        st.exception(e)
        st.stop()


def normalize_base(df_raw):
    df = df_raw.copy()

    col_semana = detect_col(df, ["Semana"])
    col_sucursal = detect_col(df, ["Sucursal"])
    col_kpi = detect_col(df, ["KPI"])
    col_categoria = detect_col(df, ["Categoria KPI", "Categoría KPI", "Categoria", "Categoría"])
    col_importe = detect_col(df, ["Importe FC", "Venta Neta", "Real_$", "Real $"])
    col_costo = detect_col(df, ["Costo FC", "Costo reposición", "Costo_$", "Costo $"])
    col_q = detect_col(df, ["Q", "Real_Q", "Real Q", "Cantidad"])
    col_obj_dinero = detect_col(df, ["Objetivo $", "Objetivo_$", "Obj $"])
    col_obj_q = detect_col(df, ["Objetivo Q", "Objetivo_Q", "Obj Q"])

    required = {
        "Semana": col_semana,
        "Sucursal": col_sucursal,
        "KPI": col_kpi,
    }

    missing = [k for k, v in required.items() if v is None]
    if missing:
        st.error(f"Faltan columnas obligatorias: {missing}")
        st.stop()

    out = pd.DataFrame()
    out["Semana"] = df[col_semana]
    out["Semana_Num"] = out["Semana"].apply(week_number)
    out["Sucursal"] = df[col_sucursal].astype(str).str.strip()
    out["KPI"] = df[col_kpi].astype(str).str.strip()

    if col_categoria:
        out["Categoria KPI"] = df[col_categoria].astype(str).str.strip()
    else:
        out["Categoria KPI"] = ""

    out["Real_$"] = df[col_importe].apply(to_number) if col_importe else 0.0
    out["Costo_$"] = df[col_costo].apply(to_number) if col_costo else 0.0
    out["Real_Q"] = df[col_q].apply(to_number) if col_q else 0.0
    out["Objetivo_$"] = df[col_obj_dinero].apply(to_number) if col_obj_dinero else 0.0
    out["Objetivo_Q"] = df[col_obj_q].apply(to_number) if col_obj_q else 0.0

    # Valor unificado para cumplimiento
    out["Real"] = np.where(out["Objetivo_$"] != 0, out["Real_$"], out["Real_Q"])
    out["Objetivo"] = np.where(out["Objetivo_$"] != 0, out["Objetivo_$"], out["Objetivo_Q"])

    out["Cumplimiento_%"] = np.where(
        out["Objetivo"] != 0,
        out["Real"] / out["Objetivo"],
        np.nan
    )

    out["GAP"] = out["Real"] - out["Objetivo"]
    out["Estado"] = out["Cumplimiento_%"].apply(classify_estado)

    return out


# ============================================================
# SIDEBAR
# ============================================================

df_raw = load_excel()
df = normalize_base(df_raw)

st.sidebar.title("⚙️ Filtros")

semanas = (
    df[["Semana", "Semana_Num"]]
    .dropna()
    .drop_duplicates()
    .sort_values("Semana_Num")
)

semana_labels = semanas["Semana"].astype(str).tolist()

semana_corte = st.sidebar.selectbox(
    "Semana de corte",
    options=semana_labels,
    index=len(semana_labels) - 1 if semana_labels else 0
)

semana_num = int(semanas.loc[semanas["Semana"].astype(str) == str(semana_corte), "Semana_Num"].iloc[0])

sucursales = sorted(df["Sucursal"].dropna().unique())
sucursales_sel = st.sidebar.multiselect(
    "Sucursal",
    options=sucursales,
    default=sucursales
)

modo_presentacion = st.sidebar.checkbox("Modo presentación", value=False)

df_corte = df[
    (df["Semana_Num"] <= semana_num) &
    (df["Sucursal"].isin(sucursales_sel))
].copy()


# ============================================================
# HEADER
# ============================================================

st.title("🚘 Tablero Posventa — Macro → Micro")
st.caption(f"Semana de corte: {semana_corte} | Vista acumulada hasta semana seleccionada")


# ============================================================
# TAB FUNCTIONS
# ============================================================

def resumen_kpis(df_view):
    total_real = df_view["Real"].sum()
    total_obj = df_view["Objetivo"].sum()
    cump = total_real / total_obj if total_obj else np.nan
    gap = total_real - total_obj

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Real acumulado", format_money(total_real))
    c2.metric("Objetivo acumulado", format_money(total_obj))
    c3.metric("Cumplimiento", format_pct(cump))
    c4.metric("GAP", format_money(gap))


def render_tab_resumen(df_view):
    st.subheader("📊 Resumen Ejecutivo")
    resumen_kpis(df_view)

    st.markdown("### Cumplimiento por sucursal")
    suc = (
        df_view.groupby("Sucursal", as_index=False)
        .agg(Real=("Real", "sum"), Objetivo=("Objetivo", "sum"))
    )
    suc["Cumplimiento_%"] = np.where(suc["Objetivo"] != 0, suc["Real"] / suc["Objetivo"], np.nan)
    suc["GAP"] = suc["Real"] - suc["Objetivo"]
    suc["Estado"] = suc["Cumplimiento_%"].apply(classify_estado)

    fig = px.bar(
        suc.sort_values("Cumplimiento_%", ascending=True),
        x="Cumplimiento_%",
        y="Sucursal",
        orientation="h",
        text=suc.sort_values("Cumplimiento_%", ascending=True)["Cumplimiento_%"].apply(format_pct),
        title="Cumplimiento acumulado por sucursal"
    )
    fig.update_layout(xaxis_tickformat=".0%")
    st.plotly_chart(fig, use_container_width=True)

    st.dataframe(
        suc.assign(
            Real=suc["Real"].apply(format_money),
            Objetivo=suc["Objetivo"].apply(format_money),
            GAP=suc["GAP"].apply(format_money),
            **{"Cumplimiento_%": suc["Cumplimiento_%"].apply(format_pct)}
        ),
        use_container_width=True,
        hide_index=True
    )


def render_tab_pnl(df_view, categoria_nombre):
    st.subheader(f"📈 P&L — {categoria_nombre}")

    df_cat = df_view[df_view["Categoria KPI"].astype(str).str.contains(categoria_nombre, case=False, na=False)].copy()

    if df_cat.empty:
        st.info(f"No hay datos para {categoria_nombre}.")
        return

    kpis = sorted(df_cat["KPI"].unique())

    kpis_sel = st.multiselect(
        f"Variables a incluir en {categoria_nombre}",
        options=kpis,
        default=kpis,
        key=f"pnl_{categoria_nombre}"
    )

    df_cat = df_cat[df_cat["KPI"].isin(kpis_sel)]

    resumen_kpis(df_cat)

    piv = (
        df_cat.groupby(["Sucursal", "KPI"], as_index=False)
        .agg(
            Real=("Real", "sum"),
            Objetivo=("Objetivo", "sum"),
            Costo=("Costo_$", "sum")
        )
    )

    piv["Margen_$"] = piv["Real"] - piv["Costo"]
    piv["Margen_%"] = np.where(piv["Real"] != 0, piv["Margen_$"] / piv["Real"], np.nan)
    piv["Cumplimiento_%"] = np.where(piv["Objetivo"] != 0, piv["Real"] / piv["Objetivo"], np.nan)
    piv["GAP"] = piv["Real"] - piv["Objetivo"]
    piv["Estado"] = piv["Cumplimiento_%"].apply(classify_estado)

    st.markdown("### Detalle por sucursal y KPI")

    st.dataframe(
        piv.assign(
            Real=piv["Real"].apply(format_money),
            Objetivo=piv["Objetivo"].apply(format_money),
            Costo=piv["Costo"].apply(format_money),
            **{
                "Margen_$": piv["Margen_$"].apply(format_money),
                "Margen_%": piv["Margen_%"].apply(format_pct),
                "Cumplimiento_%": piv["Cumplimiento_%"].apply(format_pct),
                "GAP": piv["GAP"].apply(format_money),
            }
        ).sort_values(["Sucursal", "KPI"]),
        use_container_width=True,
        hide_index=True
    )


def render_tab_otros(df_view):
    st.subheader("🔧 Otros KPIs")

    mask = (
        df_view["KPI"].astype(str).str.contains("CPU|Neum", case=False, na=False) |
        df_view["Categoria KPI"].astype(str).str.contains("CPU|Neum", case=False, na=False)
    )

    dfx = df_view[mask].copy()

    if dfx.empty:
        st.info("No hay datos para CPUs o Neumáticos.")
        return

    resumen = (
        dfx.groupby(["Sucursal", "KPI"], as_index=False)
        .agg(Real=("Real", "sum"), Objetivo=("Objetivo", "sum"))
    )

    resumen["Cumplimiento_%"] = np.where(resumen["Objetivo"] != 0, resumen["Real"] / resumen["Objetivo"], np.nan)
    resumen["GAP"] = resumen["Real"] - resumen["Objetivo"]
    resumen["Estado"] = resumen["Cumplimiento_%"].apply(classify_estado)

    st.dataframe(
        resumen.assign(
            Real=resumen["Real"].apply(format_qty),
            Objetivo=resumen["Objetivo"].apply(format_qty),
            GAP=resumen["GAP"].apply(format_qty),
            **{"Cumplimiento_%": resumen["Cumplimiento_%"].apply(format_pct)}
        ),
        use_container_width=True,
        hide_index=True
    )


def render_tab_cierre_gap(df_view):
    st.subheader("🎯 Cierre de Mes — GAP a Objetivo")
    st.caption("Pantalla de gestión para enfocar la última semana: qué falta, dónde falta y cuánto necesita cada sucursal.")

    df_work = df_view.copy()

    # -----------------------------
    # Selectores de variables
    # -----------------------------
    rep_mask = df_work["Categoria KPI"].astype(str).str.contains("Repuesto", case=False, na=False)
    serv_mask = df_work["Categoria KPI"].astype(str).str.contains("Servicio", case=False, na=False)

    rep_options = sorted(df_work.loc[rep_mask, "KPI"].dropna().unique())
    serv_options = sorted(df_work.loc[serv_mask, "KPI"].dropna().unique())

    c1, c2 = st.columns(2)

    with c1:
        rep_sel = st.multiselect(
            "Variables Repuestos a incluir",
            options=rep_options,
            default=rep_options,
            key="cierre_rep_sel"
        )

    with c2:
        serv_sel = st.multiselect(
            "Variables Servicios a incluir",
            options=serv_options,
            default=serv_options,
            key="cierre_serv_sel"
        )

    incluir_cpus = st.checkbox("Incluir CPUs", value=True)
    incluir_neumaticos = st.checkbox("Incluir Neumáticos", value=True)
    modo_cierre = st.checkbox("🔥 Modo cierre: mostrar solo pendientes", value=False)

    selected_kpis = rep_sel + serv_sel

    if incluir_cpus:
        selected_kpis += [k for k in df_work["KPI"].unique() if "cpu" in str(k).lower()]

    if incluir_neumaticos:
        selected_kpis += [k for k in df_work["KPI"].unique() if "neum" in str(k).lower()]

    selected_kpis = list(dict.fromkeys(selected_kpis))

    dfx = df_work[df_work["KPI"].isin(selected_kpis)].copy()

    if dfx.empty:
        st.warning("No hay datos para las variables seleccionadas.")
        return

    gap = (
        dfx.groupby(["Sucursal", "KPI"], as_index=False)
        .agg(
            Real=("Real", "sum"),
            Objetivo=("Objetivo", "sum")
        )
    )

    gap["Cumplimiento_%"] = np.where(gap["Objetivo"] != 0, gap["Real"] / gap["Objetivo"], np.nan)
    gap["GAP"] = gap["Real"] - gap["Objetivo"]
    gap["Falta"] = np.where(gap["GAP"] < 0, abs(gap["GAP"]), 0)
    gap["Estado"] = gap["Cumplimiento_%"].apply(classify_estado)

    if modo_cierre:
        gap = gap[gap["Cumplimiento_%"] < 1].copy()

    # -----------------------------
    # Resumen ejecutivo
    # -----------------------------
    total_real = gap["Real"].sum()
    total_obj = gap["Objetivo"].sum()
    total_gap = total_real - total_obj
    total_cump = total_real / total_obj if total_obj else np.nan
    suc_riesgo = gap[gap["Cumplimiento_%"] < 1]["Sucursal"].nunique()

    st.markdown("### 📌 Foto ejecutiva del cierre")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Cumplimiento total", format_pct(total_cump))
    m2.metric("GAP total", format_money(total_gap))
    m3.metric("Falta total", format_money(max(abs(total_gap), 0) if total_gap < 0 else 0))
    m4.metric("Sucursales con pendiente", f"{suc_riesgo}")

    # -----------------------------
    # Gap por sucursal
    # -----------------------------
    suc_gap = (
        gap.groupby("Sucursal", as_index=False)
        .agg(
            Real=("Real", "sum"),
            Objetivo=("Objetivo", "sum"),
            Falta=("Falta", "sum")
        )
    )

    suc_gap["GAP"] = suc_gap["Real"] - suc_gap["Objetivo"]
    suc_gap["Cumplimiento_%"] = np.where(suc_gap["Objetivo"] != 0, suc_gap["Real"] / suc_gap["Objetivo"], np.nan)
    suc_gap["Estado"] = suc_gap["Cumplimiento_%"].apply(classify_estado)

    st.markdown("### 📍 GAP acumulado por sucursal")

    fig_gap = px.bar(
        suc_gap.sort_values("GAP"),
        x="GAP",
        y="Sucursal",
        orientation="h",
        text=suc_gap.sort_values("GAP")["GAP"].apply(format_money),
        title="GAP vs objetivo acumulado"
    )
    fig_gap.add_vline(x=0, line_width=2, line_dash="dash")
    st.plotly_chart(fig_gap, use_container_width=True)

    # -----------------------------
    # Heatmap / tabla principal
    # -----------------------------
    st.markdown("### 🚦 Matriz de foco: Sucursal x KPI")

    tabla = gap.sort_values(["Cumplimiento_%", "GAP"], ascending=[True, True]).copy()

    tabla_show = tabla.assign(
        Real=tabla["Real"].apply(lambda x: format_money(x) if abs(x) > 1000 else format_qty(x)),
        Objetivo=tabla["Objetivo"].apply(lambda x: format_money(x) if abs(x) > 1000 else format_qty(x)),
        GAP=tabla["GAP"].apply(lambda x: format_money(x) if abs(x) > 1000 else format_qty(x)),
        Falta=tabla["Falta"].apply(lambda x: format_money(x) if abs(x) > 1000 else format_qty(x)),
        **{"Cumplimiento_%": tabla["Cumplimiento_%"].apply(format_pct)}
    )

    st.dataframe(
        tabla_show[["Sucursal", "KPI", "Real", "Objetivo", "Cumplimiento_%", "GAP", "Falta", "Estado"]],
        use_container_width=True,
        hide_index=True
    )

    # -----------------------------
    # Qué falta vender
    # -----------------------------
    st.markdown("### ⚔️ Prioridad de acción")

    pendientes = tabla[tabla["Falta"] > 0].copy()

    if pendientes.empty:
        st.success("Todas las variables seleccionadas están cumplidas o sin pendiente.")
    else:
        top = pendientes.sort_values("Falta", ascending=False).head(15)

        fig_top = px.bar(
            top,
            x="Falta",
            y="KPI",
            color="Sucursal",
            orientation="h",
            title="Top pendientes a recuperar"
        )
        st.plotly_chart(fig_top, use_container_width=True)

        st.markdown("#### Mensaje para gestión")
        st.info(
            "El foco de la reunión debe estar en las filas con mayor falta absoluta y menor cumplimiento. "
            "No se trata de revisar todo: se trata de atacar lo que todavía puede mover el cierre del mes."
        )

    # -----------------------------
    # Export
    # -----------------------------
    csv = tabla.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "⬇️ Descargar detalle GAP",
        data=csv,
        file_name=f"gap_cierre_semana_{semana_num}.csv",
        mime="text/csv"
    )


# ============================================================
# TABS
# ============================================================

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Resumen",
    "📦 Repuestos",
    "🛠️ Servicios",
    "🔧 CPUs / Neumáticos",
    "🎯 Cierre de Mes GAP"
])

with tab1:
    render_tab_resumen(df_corte)

with tab2:
    render_tab_pnl(df_corte, "Repuesto")

with tab3:
    render_tab_pnl(df_corte, "Servicio")

with tab4:
    render_tab_otros(df_corte)

with tab5:
    render_tab_cierre_gap(df_corte)


# ============================================================
# FOOTER
# ============================================================

st.caption("Tablero Posventa — versión con análisis acumulado y GAP de cierre.")
