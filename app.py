# APP COMPLETO - TABLERO POSVENTA + GAP + VOLUMEN SEPARADO
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

st.set_page_config(layout="wide")

def safe_ratio(a,b):
    return a/b if b!=0 else 0

def money(x):
    return f"$ {int(x):,}".replace(",",".") if not pd.isna(x) else "$ 0"

def qty(x):
    return f"{int(x):,}".replace(",",".") if not pd.isna(x) else "0"

def pct(x):
    return f"{x:.1%}" if not pd.isna(x) else "-"

@st.cache_data
def load():
    return pd.read_excel("base_posventa.xlsx")

df = load()

# Normalización
df["Real_$"] = pd.to_numeric(df.get("Real_$",0),errors="coerce").fillna(0)
df["Real_Q"] = pd.to_numeric(df.get("Real_Q",0),errors="coerce").fillna(0)
df["Objetivo_$"] = pd.to_numeric(df.get("Objetivo_$",0),errors="coerce").fillna(0)
df["Objetivo_Q"] = pd.to_numeric(df.get("Objetivo_Q",0),errors="coerce").fillna(0)

df["Real_val"] = np.where(df["Objetivo_$"]!=0, df["Real_$"], df["Real_Q"])
df["Obj_val"] = np.where(df["Objetivo_$"]!=0, df["Objetivo_$"], df["Objetivo_Q"])

# Filtros
st.sidebar.title("Filtros")
semana = st.sidebar.selectbox("Semana", sorted(df["Semana"].unique()))
df_cut = df[df["Semana"]<=semana]

sucs = st.sidebar.multiselect("Sucursal", df_cut["Sucursal"].unique(), default=df_cut["Sucursal"].unique())
df_cut = df_cut[df_cut["Sucursal"].isin(sucs)]

# Tabs
tab1, tab2 = st.tabs(["Resumen","🎯 GAP"])

# ---------------- RESUMEN ----------------
with tab1:
    st.title("Resumen")

    real = df_cut["Real_val"].sum()
    obj = df_cut["Obj_val"].sum()

    c1,c2,c3 = st.columns(3)
    c1.metric("Real", money(real))
    c2.metric("Objetivo", money(obj))
    c3.metric("Cumplimiento", pct(safe_ratio(real,obj)))

    st.markdown("### 🔍 Volumen (Q) separado")

    df_q = df_cut[df_cut["Objetivo_Q"]>0]

    cpu = df_q[df_q["KPI"].str.contains("CPU",case=False,na=False)]
    neum = df_q[df_q["KPI"].str.contains("NEUM",case=False,na=False)]
    otros = df_q[~df_q["KPI"].str.contains("CPU|NEUM",case=False,na=False)]

    def calc(d):
        return d["Real_Q"].sum(), d["Objetivo_Q"].sum()

    cr,co = calc(cpu)
    nr,no = calc(neum)
    or_,oo = calc(otros)

    c1,c2,c3 = st.columns(3)

    c1.metric("CPUs", pct(safe_ratio(cr,co)), f"{qty(cr)} / {qty(co)}")
    c2.metric("Neumáticos", pct(safe_ratio(nr,no)), f"{qty(nr)} / {qty(no)}")
    c3.metric("Otros", pct(safe_ratio(or_,oo)), f"{qty(or_)} / {qty(oo)}")

# ---------------- GAP ----------------
with tab2:
    st.title("🎯 Cierre GAP")

    gap = df_cut.groupby(["Sucursal","KPI"],as_index=False).agg(
        Real=("Real_val","sum"),
        Obj=("Obj_val","sum")
    )

    gap["GAP"] = gap["Real"]-gap["Obj"]
    gap["Cumpl"] = gap.apply(lambda r: safe_ratio(r["Real"],r["Obj"]),axis=1)
    gap["Falta"] = np.where(gap["GAP"]<0,abs(gap["GAP"]),0)

    tr = gap["Real"].sum()
    to = gap["Obj"].sum()

    c1,c2,c3 = st.columns(3)
    c1.metric("Cumplimiento", pct(safe_ratio(tr,to)))
    c2.metric("GAP", money(tr-to))
    c3.metric("Falta", money(abs(tr-to)))

    st.subheader("GAP por sucursal")

    suc = gap.groupby("Sucursal",as_index=False)["GAP"].sum().sort_values("GAP")
    fig = px.bar(suc,x="GAP",y="Sucursal",orientation="h")
    fig.add_vline(x=0)
    st.plotly_chart(fig,use_container_width=True)

    st.subheader("Detalle")
    show = gap.copy()
    show["Real"] = show["Real"].apply(money)
    show["Obj"] = show["Obj"].apply(money)
    show["GAP"] = show["GAP"].apply(money)
    show["Falta"] = show["Falta"].apply(money)
    show["Cumpl"] = show["Cumpl"].apply(pct)

    st.dataframe(show,use_container_width=True)

    st.subheader("Prioridades")
    top = gap.sort_values("Falta",ascending=False).head(10)
    fig2 = px.bar(top,x="Falta",y="KPI",color="Sucursal",orientation="h")
    st.plotly_chart(fig2,use_container_width=True)
