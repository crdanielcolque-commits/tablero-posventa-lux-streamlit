import re
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px
import gdown

st.set_page_config(page_title="Tablero Posventa", layout="wide")

# ==========================
# CONFIG DRIVE (ACTUALIZADO)
# ==========================
DRIVE_FILE_ID = "191JKfQWj3yehcnisKTPDs_KpWaOTyslhQ0g273Xvzjc"
EXCEL_LOCAL = "base_posventa.xlsx"

# ==========================
# ESTILO (layout dirección)
# ==========================
st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
section[data-testid="stSidebar"] .block-container {padding-top: 1.2rem;}
h1 {margin-bottom: 0.2rem;}

.kpi-row {display:flex; gap:14px; flex-wrap:wrap; margin: 6px 0 10px 0;}
.kpi-card {
  background: #ffffff;
  border: 1px solid rgba(0,0,0,0.08);
  border-radius: 14px;
  padding: 14px 16px;
  box-shadow: 0 1px 10px rgba(0,0,0,0.04);
  min-width: 240px;
  flex: 1;
}
.kpi-title {font-size: 0.82rem; opacity: 0.72; margin-bottom: 6px;}
.kpi-value {font-size: 1.55rem; font-weight: 700; line-height: 1.2;}
.badge {
  display:inline-block; padding: 6px 10px; border-radius: 999px;
  font-size: 0.85rem; font-weight: 600; color: white;
}
.badge-red {background:#d64545;}
.badge-yellow {background:#d1a100;}
.badge-green {background:#2c9f6b;}
.badge-gray {background:#6c757d;}

.chips {display:flex; gap:10px; flex-wrap:wrap; margin-top: 8px;}
.chip {
  background: rgba(0,0,0,0.04);
  border: 1px solid rgba(0,0,0,0.06);
  border-radius: 999px;
  padding: 6px 10px;
  font-size: 0.9rem;
  white-space: nowrap;
}

.hr {height:1px; background:rgba(0,0,0,0.08); margin: 16px 0;}
</style>
""", unsafe_allow_html=True)

# ==========================
# Helpers
# ==========================
def _norm(s: str) -> str:
    """Normaliza: minus, sin acentos, sin espacios, sin símbolos raros."""
    s = str(s).strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = s.replace(" ", "").replace("-", "").replace(".", "").replace("/", "")
    s = s.replace("__", "_")
    return s

def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Devuelve el nombre real de la columna que matchea con algún candidato normalizado."""
    norm_map = {_norm(c): c for c in df.columns}
    for cand in candidates:
        key = _norm(cand)
        if key in norm_map:
            return norm_map[key]
    # match por contains (último recurso)
    for cand in candidates:
        key = _norm(cand)
        for k, real in norm_map.items():
            if key in k:
                return real
    return None

def parse_semana_num(x):
    if pd.isna(x):
        return None
    m = re.search(r"(\d+)", str(x))
    return int(m.group(1)) if m else None

def money_fmt(x):
    try:
        return f"${x:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(x)

def pct_fmt(x):
    if x is None or pd.isna(x):
        return "—"
    return f"{x*100:.1f}%"

def estado_por_umbral(cumpl, umbral_amar, umbral_verde):
    if cumpl is None or pd.isna(cumpl):
        return "—"
    if cumpl >= umbral_verde:
        return "Verde"
    if cumpl >= umbral_amar:
        return "Amarillo"
    return "Rojo"

def estado_global(serie_estados):
    s = set([x for x in serie_estados.dropna().tolist()])
    if "Rojo" in s: return "Rojo"
    if "Amarillo" in s: return "Amarillo"
    if "Verde" in s: return "Verde"
    return "—"

def badge_html(estado):
    if estado == "Rojo":
        return '<span class="badge badge-red">ROJO</span>'
    if estado == "Amarillo":
        return '<span class="badge badge-yellow">AMARILLO</span>'
    if estado == "Verde":
        return '<span class="badge badge-green">VERDE</span>'
    return '<span class="badge badge-gray">—</span>'

# ==========================
# Carga desde Google Sheets (export a XLSX)
# ==========================
@st.cache_data(show_spinner=True, ttl=300)
def load_from_drive():
    url = f"https://docs.google.com/spreadsheets/d/{DRIVE_FILE_ID}/export?format=xlsx"
    gdown.download(url, EXCEL_LOCAL, quiet=True, fuzzy=True)

    # Si el archivo nuevo cambió nombres de hojas, ajustá acá:
    df = pd.read_excel(EXCEL_LOCAL, sheet_name=0)  # primera hoja (más robusto)
    try:
        dim_kpi = pd.read_excel(EXCEL_LOCAL, sheet_name="DIM_KPI")
    except Exception:
        # si no existe DIM_KPI, creamos uno mínimo (umbrales por defecto)
        dim_kpi = pd.DataFrame(columns=["KPI", "Umbral_Amarillo", "Umbral_Verde"])

    # Normaliza headers (solo strip)
    df.columns = [str(c).strip() for c in df.columns]
    dim_kpi.columns = [str(c).strip() for c in dim_kpi.columns]

    return df, dim_kpi

# ==========================
# Resolver columnas del Excel nuevo
# ==========================
def resolve_schema(df: pd.DataFrame) -> dict:
    col = {}

    col["fecha"] = find_col(df, ["Fecha", "FECHA", "date"])
    col["semana"] = find_col(df, ["Semana", "SEMANA", "Semana corte", "Semana_Corte"])
    col["sucursal"] = find_col(df, ["Sucursal", "SUCURSAL", "Sede", "Localidad"])
    col["kpi"] = find_col(df, ["KPI", "Indicador", "Indicador_KPI"])
    col["tipo_kpi"] = find_col(df, ["Tipo_KPI", "Tipo KPI", "Tipo", "TipoIndicador", "Tipo_Indicador"])

    # Real y Objetivo (money y qty)
    col["real_$"] = find_col(df, ["Real_$", "Real $", "Real$", "Real_ARS", "Real_Ars", "Real_Pesos", "Real_Monto"])
    col["obj_$"]  = find_col(df, ["Objetivo_$", "Objetivo $", "Objetivo$", "Obj_$", "Obj $", "Objetivo_ARS", "Objetivo_Monto"])

    col["real_q"] = find_col(df, ["Real_Q", "Real Q", "Real_Unidades", "Real_UD", "Real_Cant", "Real_Cantidad"])
    col["obj_q"]  = find_col(df, ["Objetivo_Q", "Objetivo Q", "Obj_Q", "Objetivo_Unidades", "Objetivo_Cant", "Objetivo_Cantidad"])

    col["costo_$"]  = find_col(df, ["Costo_$", "Costo $", "Costo$", "Costo_Monto"])
    col["margen_$"] = find_col(df, ["Margen_$", "Margen $", "Margen$", "Margen_Monto"])

    missing = [k for k,v in col.items() if v is None and k in ["semana","sucursal","kpi","tipo_kpi","real_$","obj_$","real_q","obj_q","costo_$","margen_$"]]
    if missing:
        return {"ok": False, "missing": missing, "found_cols": list(df.columns), "col": col}
    return {"ok": True, "col": col}

# ==========================
# DIM KPI (umbrales)
# ==========================
def normalize_dim_kpi(dim_kpi: pd.DataFrame) -> pd.DataFrame:
    if dim_kpi is None or len(dim_kpi) == 0:
        return pd.DataFrame(columns=["KPI","Umbral_Amarillo","Umbral_Verde"])

    kpi_col = find_col(dim_kpi, ["KPI","Indicador","Indicador_KPI"])
    ua_col  = find_col(dim_kpi, ["Umbral_Amarillo","Umbral Amarillo","Amarillo"])
    uv_col  = find_col(dim_kpi, ["Umbral_Verde","Umbral Verde","Verde"])

    out = pd.DataFrame()
    out["KPI"] = dim_kpi[kpi_col] if kpi_col else None
    out["Umbral_Amarillo"] = dim_kpi[ua_col] if ua_col else 0.90
    out["Umbral_Verde"] = dim_kpi[uv_col] if uv_col else 1.00
    out["Umbral_Amarillo"] = pd.to_numeric(out["Umbral_Amarillo"], errors="coerce").fillna(0.90)
    out["Umbral_Verde"] = pd.to_numeric(out["Umbral_Verde"], errors="coerce").fillna(1.00)
    out = out.dropna(subset=["KPI"]).copy()
    return out

# ==========================
# Transformaciones
# ==========================
def build_kpi_week(df_raw: pd.DataFrame, schema: dict) -> pd.DataFrame:
    c = schema["col"]
    df = df_raw.copy()

    # Fecha (si existe)
    if c["fecha"]:
        df[c["fecha"]] = pd.to_datetime(df[c["fecha"]], errors="coerce")

    df["Semana_Num"] = df[c["semana"]].apply(parse_semana_num)

    # Asegurar tipos numéricos
    for k in ["real_$","obj_$","real_q","obj_q","costo_$","margen_$"]:
        df[c[k]] = pd.to_numeric(df[c[k]], errors="coerce")

    # Valores unificados
    df["Real_val"] = df.apply(lambda r: r[c["real_$"]] if r[c["tipo_kpi"]] == "$" else r[c["real_q"]], axis=1)
    df["Obj_val"]  = df.apply(lambda r: r[c["obj_$"]]  if r[c["tipo_kpi"]] == "$" else r[c["obj_q"]], axis=1)

    agg = df.groupby(
        ["Semana_Num", c["sucursal"], c["kpi"], c["tipo_kpi"]],
        as_index=False
    ).agg(
        Real_Sem=("Real_val", "sum"),
        Obj_Sem=("Obj_val", "max"),
        Costo_Sem=(c["costo_$"], "sum"),
        Margen_Sem=(c["margen_$"], "sum"),
    )

    # Renombres a estándar
    agg = agg.rename(columns={
        c["sucursal"]: "Sucursal",
        c["kpi"]: "KPI",
        c["tipo_kpi"]: "Tipo_KPI",
    })

    agg["Cumpl_Sem"] = agg["Real_Sem"] / agg["Obj_Sem"]

    agg = agg.sort_values(["Sucursal", "KPI", "Semana_Num"]).copy()
    agg["Real_Acum"] = agg.groupby(["Sucursal","KPI"])["Real_Sem"].cumsum()
    agg["Obj_Acum"]  = agg.groupby(["Sucursal","KPI"])["Obj_Sem"].cumsum()
    agg["Cumpl_Acum"] = agg["Real_Acum"] / agg["Obj_Acum"]

    agg["Margen_Acum"] = agg.groupby(["Sucursal","KPI"])["Margen_Sem"].cumsum()
    agg["MargenPct_Acum"] = (agg["Margen_Acum"] / agg["Real_Acum"]).where(agg["Real_Acum"] != 0)

    return agg

def aplicar_reglas(df_last, dim_kpi):
    out = df_last.merge(dim_kpi[["KPI","Umbral_Amarillo","Umbral_Verde"]], on="KPI", how="left")
    out["Umbral_Amarillo"] = out["Umbral_Amarillo"].fillna(0.90)
    out["Umbral_Verde"] = out["Umbral_Verde"].fillna(1.00)
    out["Estado_Acum"] = out.apply(
        lambda r: estado_por_umbral(r["Cumpl_Acum"], r["Umbral_Amarillo"], r["Umbral_Verde"]),
        axis=1
    )
    return out

def consolidar_todas(df_last_suc):
    cons = df_last_suc.groupby(["KPI","Tipo_KPI"], as_index=False).agg(
        Real_Acum=("Real_Acum","sum"),
        Obj_Acum=("Obj_Acum","sum"),
        Margen_Acum=("Margen_Acum","sum"),
        Real_Sem=("Real_Sem","sum"),
        Obj_Sem=("Obj_Sem","sum"),
        Margen_Sem=("Margen_Sem","sum"),
    )
    cons["Cumpl_Acum"] = cons["Real_Acum"] / cons["Obj_Acum"]
    cons["Cumpl_Sem"]  = cons["Real_Sem"] / cons["Obj_Sem"]
    cons["MargenPct_Acum"] = (cons["Margen_Acum"] / cons["Real_Acum"]).where(cons["Real_Acum"] != 0)
    return cons

# ==========================
# App
# ==========================
df_raw, dim_kpi_raw = load_from_drive()
schema = resolve_schema(df_raw)

if not schema["ok"]:
    st.error("⚠️ El archivo nuevo cambió nombres de columnas y faltan campos necesarios.")
    st.write("**Faltan (claves):**", schema["missing"])
    st.write("**Columnas encontradas en tu hoja:**")
    st.code("\n".join(schema["found_cols"]))
    st.stop()

dim_kpi = normalize_dim_kpi(dim_kpi_raw)

df_week = build_kpi_week(df_raw, schema)

# Sidebar filtros
st.sidebar.title("Filtros obligatorios")
semanas = sorted(df_week["Semana_Num"].dropna().unique())
sucursales = sorted(df_week["Sucursal"].dropna().unique())

semana_corte = st.sidebar.selectbox("Semana corte", semanas, index=len(semanas)-1)
sucursal = st.sidebar.selectbox("Sucursal", ["TODAS (Consolidado)"] + sucursales)

# Snapshots semana corte
df_last_suc = df_week[df_week["Semana_Num"] == semana_corte].copy()
df_last_suc = aplicar_reglas(df_last_suc, dim_kpi)

# Vista principal
if sucursal != "TODAS (Consolidado)":
    df_last = df_last_suc[df_last_suc["Sucursal"] == sucursal].copy()
else:
    df_last = consolidar_todas(df_last_suc)
    df_last = aplicar_reglas(df_last, dim_kpi)

# ==========================
# Header
# ==========================
st.title("Tablero Posventa — Semanal + Acumulado")
st.caption(f"Sucursal: **{sucursal}** | Corte semana **{semana_corte}**")

tab1, tab2, tab3 = st.tabs(["🏠 Resumen Ejecutivo", "📈 Seguimiento", "🧩 Gestión"])

# ==========================
# TAB 1: Resumen Ejecutivo
# ==========================
with tab1:
    econ = df_last[df_last["Tipo_KPI"] == "$"].copy()
    oper = df_last[df_last["Tipo_KPI"] == "Q"].copy()

    # Consolidados SUM/SUM
    econ_real = float(econ["Real_Acum"].sum()) if len(econ) else 0.0
    econ_obj  = float(econ["Obj_Acum"].sum())  if len(econ) else 0.0
    econ_cump = (econ_real / econ_obj) if econ_obj else None

    econ_margen = float(econ["Margen_Acum"].sum()) if ("Margen_Acum" in econ.columns and len(econ)) else 0.0
    econ_margen_pct = (econ_margen / econ_real) if econ_real else None

    oper_real = float(oper["Real_Acum"].sum()) if len(oper) else 0.0
    oper_obj  = float(oper["Obj_Acum"].sum())  if len(oper) else 0.0
    oper_cump = (oper_real / oper_obj) if oper_obj else None

    # Semáforo global + conteos
    econ_rojo = int((econ["Estado_Acum"] == "Rojo").sum()) if len(econ) else 0
    econ_amar = int((econ["Estado_Acum"] == "Amarillo").sum()) if len(econ) else 0
    econ_ver  = int((econ["Estado_Acum"] == "Verde").sum()) if len(econ) else 0

    oper_rojo = int((oper["Estado_Acum"] == "Rojo").sum()) if len(oper) else 0
    oper_amar = int((oper["Estado_Acum"] == "Amarillo").sum()) if len(oper) else 0
    oper_ver  = int((oper["Estado_Acum"] == "Verde").sum()) if len(oper) else 0

    glob_econ = estado_global(econ["Estado_Acum"]) if len(econ) else "—"
    glob_oper = estado_global(oper["Estado_Acum"]) if len(oper) else "—"

    st.markdown(f"""
    <div class="kpi-row">
      <div class="kpi-card">
        <div class="kpi-title">Estado Global ($)</div>
        <div class="kpi-value">{badge_html(glob_econ)}</div>
        <div class="chips">
          <div class="chip">Rojos: <b>{econ_rojo}</b></div>
          <div class="chip">Amarillos: <b>{econ_amar}</b></div>
          <div class="chip">Verdes: <b>{econ_ver}</b></div>
        </div>
      </div>
      <div class="kpi-card">
        <div class="kpi-title">Estado Global (Q)</div>
        <div class="kpi-value">{badge_html(glob_oper)}</div>
        <div class="chips">
          <div class="chip">Rojos: <b>{oper_rojo}</b></div>
          <div class="chip">Amarillos: <b>{oper_amar}</b></div>
          <div class="chip">Verdes: <b>{oper_ver}</b></div>
        </div>
      </div>
    </div>
    <div class="hr"></div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2, gap="large")

    with col1:
        st.subheader("🔵 Económico ($)")
        st.markdown(f"""
        <div class="kpi-row">
          <div class="kpi-card">
            <div class="kpi-title">Cumplimiento Acumulado ($)</div>
            <div class="kpi-value">{pct_fmt(econ_cump)}</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-title">Facturación Acumulada</div>
            <div class="kpi-value">{money_fmt(econ_real)}</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-title">Margen % Acumulado</div>
            <div class="kpi-value">{pct_fmt(econ_margen_pct)}</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        if len(econ):
            fig = px.bar(
                econ.sort_values("Cumpl_Acum"),
                x="Cumpl_Acum", y="KPI", orientation="h",
                text=econ["Cumpl_Acum"].apply(lambda x: f"{x*100:.1f}%")
            )
            fig.update_layout(xaxis_tickformat=".0%")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay KPIs económicos ($) en este corte.")

    with col2:
        st.subheader("🟢 Operativo (Q)")
        st.markdown(f"""
        <div class="kpi-row">
          <div class="kpi-card">
            <div class="kpi-title">Cumplimiento Acumulado (Q)</div>
            <div class="kpi-value">{pct_fmt(oper_cump)}</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-title">Real Acumulado (Q)</div>
            <div class="kpi-value">{f"{oper_real:,.0f}".replace(",", ".")}</div>
          </div>
          <div class="kpi-card">
            <div class="kpi-title">Objetivo Acumulado (Q)</div>
            <div class="kpi-value">{f"{oper_obj:,.0f}".replace(",", ".")}</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        if len(oper):
            fig2 = px.bar(
                oper.sort_values("Cumpl_Acum"),
                x="Cumpl_Acum", y="KPI", orientation="h",
                text=oper["Cumpl_Acum"].apply(lambda x: f"{x*100:.1f}%")
            )
            fig2.update_layout(xaxis_tickformat=".0%")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("No hay KPIs operativos (Q) en este corte.")

    if sucursal == "TODAS (Consolidado)":
        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)
        st.subheader("🔥 Heatmap KPI × Sucursal (Cumplimiento Acumulado)")

        heat = df_last_suc.pivot_table(
            index="KPI",
            columns="Sucursal",
            values="Cumpl_Acum",
            aggfunc="mean"
        ).reset_index().melt(id_vars="KPI", var_name="Sucursal", value_name="Cumpl_Acum")

        fig_h = px.density_heatmap(
            heat,
            x="Sucursal",
            y="KPI",
            z="Cumpl_Acum",
            histfunc="avg",
            title="Cumplimiento Acumulado por KPI y Sucursal",
            labels={"Cumpl_Acum": "Cumpl. Acum"}
        )
        fig_h.update_layout(coloraxis_colorbar=dict(tickformat=".0%"))
        st.plotly_chart(fig_h, use_container_width=True)

# ==========================
# TAB 2: Seguimiento
# ==========================
with tab2:
    st.subheader("Seguimiento por KPI (semanal vs acumulado)")

    if sucursal == "TODAS (Consolidado)":
        st.info("Para seguimiento semanal por KPI, elegí una sucursal (en consolidado mezclaría bases).")
    else:
        kpis = sorted(df_week[df_week["Sucursal"] == sucursal]["KPI"].dropna().unique())
        kpi_sel = st.selectbox("KPI", kpis)

        serie = df_week[(df_week["Sucursal"] == sucursal) & (df_week["KPI"] == kpi_sel)].sort_values("Semana_Num")

        c1, c2 = st.columns(2)
        with c1:
            fig1 = px.line(serie, x="Semana_Num", y="Cumpl_Sem", markers=True, title="Cumplimiento semanal")
            fig1.update_layout(yaxis_tickformat=".0%")
            st.plotly_chart(fig1, use_container_width=True)

        with c2:
            fig2 = px.line(serie, x="Semana_Num", y="Cumpl_Acum", markers=True, title="Cumplimiento acumulado")
            fig2.update_layout(yaxis_tickformat=".0%")
            st.plotly_chart(fig2, use_container_width=True)

# ==========================
# TAB 3: Gestión
# ==========================
with tab3:
    st.subheader("Gestión por desvíos (acumulado)")

    g = df_last.copy()
    g["Gap"] = g["Obj_Acum"] - g["Real_Acum"]

    order_map = {"Rojo": 0, "Amarillo": 1, "Verde": 2, "—": 9}
    g["OrdenEstado"] = g["Estado_Acum"].map(order_map).fillna(9)
    g = g.sort_values(["OrdenEstado", "Gap"], ascending=[True, False])

    def fmt_val(tipo, v):
        if pd.isna(v): return "—"
        return money_fmt(v) if tipo == "$" else f"{v:,.0f}".replace(",", ".")

    g["Cumpl_Acum_fmt"] = g["Cumpl_Acum"].apply(pct_fmt)
    g["Real_Acum_fmt"] = g.apply(lambda r: fmt_val(r["Tipo_KPI"], r["Real_Acum"]), axis=1)
    g["Obj_Acum_fmt"]  = g.apply(lambda r: fmt_val(r["Tipo_KPI"], r["Obj_Acum"]), axis=1)
    g["Gap_fmt"]       = g.apply(lambda r: fmt_val(r["Tipo_KPI"], r["Gap"]), axis=1)

    st.dataframe(
        g[["KPI","Tipo_KPI","Estado_Acum","Cumpl_Acum_fmt","Real_Acum_fmt","Obj_Acum_fmt","Gap_fmt"]],
        use_container_width=True,
        hide_index=True
    )
