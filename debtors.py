"""Análisis de la foto de cartera de Deudores Postventa."""
import re
import unicodedata
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

BRANCHES = {1: "Salta Unidades", 2: "Jujuy", 3: "Taller Móvil", 4: "Salta",
            5: "Tartagal", 7: "Lajitas", 9: "Chapa y Pintura"}
MONEY = ["Crédito máximo", "Saldo", "A vencer", "Vencido", "0 a 30", "31 a 60", "61 a 90", "Más de 90"]
BANDS = ["0 a 30", "31 a 60", "61 a 90", "Más de 90"]


def normalized(value):
    return "".join(c for c in unicodedata.normalize("NFKD", str(value).lower()) if not unicodedata.combining(c)).strip()


def amount(value):
    if pd.isna(value) or str(value).strip() == "":
        return np.nan
    if isinstance(value, (int, float, np.number)):
        return float(value)
    text = str(value).strip().replace("$", "").replace(" ", "")
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return np.nan


def money(value):
    return "$" + f"{value:,.0f}".replace(",", ".")


def prepare_debtors(raw):
    aliases = {
        "codigo": "Código", "cliente": "Cliente", "c.u.i.t": "CUIT", "suc.": "Código sucursal",
        "tipo": "Tipo", "gestion": "Gestión", "dias": "Días", "cred. maximo": "Crédito máximo",
        "saldo": "Saldo", "a vencer": "A vencer", "vencido": "Vencido",
        "0 a 30": "0 a 30", "31 a 60": "31 a 60", "61 a 90": "61 a 90", "mas de 90 dias": "Más de 90",
    }
    d = raw.rename(columns={c: aliases.get(normalized(c), c) for c in raw.columns}).copy()
    required = list(aliases.values())
    missing = [c for c in required if c not in d]
    if missing:
        raise ValueError("Faltan columnas en Deudores Postventa: " + ", ".join(missing))
    d = d[d["Cliente"].notna()].copy()
    for c in ["Código", "Cliente", "CUIT", "Tipo", "Gestión"]:
        d[c] = d[c].fillna("").astype(str).str.strip()
    for c in MONEY + ["Días"]:
        d[c] = d[c].map(amount)
    # Los importes vacíos de composición son cero; saldo y límite desconocidos siguen vacíos.
    d[BANDS + ["A vencer", "Vencido"]] = d[BANDS + ["A vencer", "Vencido"]].fillna(0)
    codes = pd.to_numeric(d["Código sucursal"], errors="coerce")
    d["Sucursal"] = codes.map(BRANCHES)
    missing_branch = d["Sucursal"].isna()
    d.loc[missing_branch, "Sucursal"] = d.loc[missing_branch, "Código sucursal"].map(
        lambda c: f"Sucursal {c:g} — sin identificar" if isinstance(c, (int, float)) and pd.notna(c) else f"Sucursal {c} — sin identificar")
    d["Zona"] = np.select([codes.isin([2, 3, 4]), codes.isin([5, 7])],
                          ["Posventa Centro", "Posventa Interior"], default="Otras sucursales")
    cuit = d["CUIT"].str.replace(r"\D", "", regex=True)
    d["Cliente ID"] = np.where(cuit.ne(""), "CUIT:" + cuit, "COD:" + d["Código"])
    d["Deuda"] = d["Saldo"].clip(lower=0)
    d["Vencido positivo"] = d["Vencido"].clip(lower=0).where(d["Saldo"] > 0, 0)
    d["Mayor a 90 positivo"] = d["Más de 90"].clip(lower=0).where(d["Saldo"] > 0, 0)
    d["Diferencia saldo"] = d["Saldo"] - d["A vencer"] - d["Vencido"]
    d["Diferencia vencido"] = d["Vencido"] - d[BANDS].sum(axis=1)
    return d


def customer_credit(d):
    # El límite es del cliente: no se suma al repetirse entre sucursales.
    g = d.groupby("Cliente ID", as_index=False).agg(
        Cliente=("Cliente", "first"), Saldo=("Saldo", lambda s: s.sum(min_count=1)),
        Límite=("Crédito máximo", "max"),
        Límites_distintos=("Crédito máximo", "nunique"))
    g["Exceso"] = (g["Saldo"] - g["Límite"]).clip(lower=0).where(g["Límite"] > 0, np.nan)
    g["Uso crédito"] = (g["Saldo"] / g["Límite"]).where(g["Límite"] > 0)
    return g


def prioritize(d):
    out = d.copy()
    def action(r):
        if pd.isna(r["Saldo"]):
            return "Revisar", "Completar el saldo antes de gestionar."
        if r["Saldo"] <= 0:
            return "Sin deuda", "Conciliar saldo a favor o cuenta saldada."
        if "judicial" in normalized(r["Gestión"]):
            return "Judicial", "Coordinar seguimiento con el responsable de la gestión judicial."
        if r["Mayor a 90 positivo"] > 0:
            return "Alta", "Contactar y conciliar deuda mayor a 90 días; acordar fecha de pago y responsable."
        if r["61 a 90"] > 0:
            return "Alta", "Gestionar compromiso de pago antes de superar los 90 días."
        if r["Vencido positivo"] > 0:
            return "Media", "Confirmar recepción de facturas y obtener compromiso de pago del vencido."
        return "Preventiva", "Enviar estado de cuenta y recordar próximos vencimientos."
    pairs = out.apply(action, axis=1)
    out["Prioridad"] = [p[0] for p in pairs]
    out["Acción sugerida"] = [p[1] for p in pairs]
    rank = {"Judicial": 0, "Alta": 1, "Media": 2, "Preventiva": 3, "Revisar": 4, "Sin deuda": 5}
    out["_orden"] = out["Prioridad"].map(rank)
    return out.sort_values(["_orden", "Vencido positivo", "Deuda"], ascending=[True, False, False]).drop(columns="_orden")


def debt_table(data, key):
    formats = {c: st.column_config.NumberColumn(c, format="$ %.2f") for c in MONEY + ["Deuda", "Exceso", "Límite", "Vencido positivo", "Mayor a 90 positivo"] if c in data}
    st.dataframe(data, hide_index=True, use_container_width=True, column_config=formats, key=key)


def render_debtors_tab(raw, sheet_found):
    st.markdown("## Deudores de postventa")
    st.caption("Foto de cartera de la hoja de origen. Filtros propios: no depende del mes, semana ni garantías del P&L.")
    if not sheet_found:
        st.info("No se encontró la hoja Deudores Postventa en la base.")
        return
    if raw.empty:
        st.info("La hoja Deudores Postventa está vacía.")
        return
    try:
        all_data = prepare_debtors(raw)
    except ValueError as exc:
        st.error(str(exc))
        return
    if all_data.empty:
        st.info("No hay clientes en la hoja.")
        return
    with st.expander("Filtros de cartera", expanded=True):
        a, b, c = st.columns(3)
        branches = a.multiselect("Sucursales", sorted(all_data["Sucursal"].unique()), key="debt_branches")
        types = b.multiselect("Tipo de cliente", sorted(all_data["Tipo"].unique()), key="debt_types")
        management = c.multiselect("Gestión de cobranza", sorted(all_data["Gestión"].unique()), key="debt_management")
        a, b = st.columns(2)
        zone = a.selectbox("Zona", ["Todas", "Posventa Centro", "Posventa Interior", "Otras sucursales"], key="debt_zone")
        query = b.text_input("Buscar cliente, código o CUIT", key="debt_search").strip()
    d = all_data.copy()
    for col, selection in [("Sucursal", branches), ("Tipo", types), ("Gestión", management)]:
        if selection:
            d = d[d[col].isin(selection)]
    if zone != "Todas":
        d = d[d["Zona"] == zone]
    if query:
        match = d[["Cliente", "Código", "CUIT"]].apply(lambda s: s.map(normalized).str.contains(normalized(query), regex=False)).any(axis=1)
        d = d[match]
    if d.empty:
        st.info("No hay clientes para los filtros seleccionados.")
        return
    positive = d[d["Saldo"] > 0]
    total = positive["Saldo"].sum()
    overdue = positive["Vencido positivo"].sum()
    old = positive["Mayor a 90 positivo"].sum()
    credits = customer_credit(d)
    a, b, c, e = st.columns(4)
    a.metric("Saldo deudor", money(total))
    b.metric("Vencido de deudores", money(overdue))
    c.metric("Vencido mayor a 90 días", money(old))
    e.metric("Clientes deudores", positive["Cliente ID"].nunique())
    a, b, c = st.columns(3)
    a.metric("A vencer de deudores", money(positive["A vencer"].sum()))
    b.metric("Saldos a favor", money(-d.loc[d["Saldo"] < 0, "Saldo"].sum()))
    c.metric("Saldo neto de cartera", money(d["Saldo"].sum()))
    st.caption("Saldo deudor suma cuentas con saldo positivo por cliente y sucursal. Saldos a favor se muestran aparte. Los tramos conservan los ajustes de la fuente; un tramo positivo puede superar el saldo neto de una cuenta.")

    st.markdown("### Focos y acciones")
    if not positive.empty:
        top = positive.sort_values("Vencido positivo", ascending=False).iloc[0]
        if overdue > 0:
            st.info(f"Primer foco por importe vencido: {top['Cliente']} ({top['Sucursal']}), {money(top['Vencido positivo'])}. Confirmar documentación y acordar fecha de pago.")
            concentration = positive.groupby("Cliente ID")["Vencido positivo"].sum().nlargest(10).sum() / overdue
            st.write(f"Los 10 principales clientes concentran el {concentration:.1%} del vencido positivo.")
        if old > 0:
            st.write(f"Deuda mayor a 90 días: {money(old)}. Priorizar conciliación y compromisos de pago; coordinar los casos judiciales con su responsable.")
    over = credits[credits["Exceso"] > 0]
    if not over.empty:
        st.write(f"{len(over)} clientes superan el límite informado por un total de {money(over['Exceso'].sum())}. Revisar condiciones de crédito y acordar regularización.")
    missing_limit = credits[(credits["Saldo"] > 0) & (credits["Límite"].isna() | (credits["Límite"] <= 0))]
    st.write(f"{len(missing_limit)} clientes deudores no tienen un límite positivo informado. Validar el dato antes de evaluar exposición.")
    judicial = positive[positive["Gestión"].map(normalized).str.contains("judicial", regex=False)]
    if not judicial.empty:
        st.write(f"Gestión judicial: {judicial['Cliente ID'].nunique()} clientes, saldo deudor {money(judicial['Saldo'].sum())}.")

    st.markdown("### Cartera por sucursal")
    branch = d.groupby("Sucursal", as_index=False).agg(
        Deuda=("Deuda", "sum"), Vencido=("Vencido positivo", "sum"),
        Mayor_90=("Mayor a 90 positivo", "sum"))
    branch = branch.sort_values("Vencido", ascending=False)
    branch["Acción"] = np.where(branch["Mayor_90"] > 0, "Priorizar deuda mayor a 90 días",
                                  np.where(branch["Vencido"] > 0, "Acordar pagos del vencido", "Seguimiento preventivo"))
    a, b = st.columns(2)
    with a:
        fig = px.bar(branch, x="Sucursal", y=["Deuda", "Vencido"], barmode="group",
                     color_discrete_sequence=["#1d4ed8", "#d97706"], labels={"value": "Importe", "variable": "Concepto"})
        fig.update_layout(height=350, margin=dict(t=15, b=15))
        st.plotly_chart(fig, use_container_width=True, key="debt_branch_chart")
    with b:
        aging = pd.DataFrame({"Tramo": ["A vencer"] + BANDS,
                              "Importe": positive[["A vencer"] + BANDS].sum().values})
        fig = px.bar(aging, x="Tramo", y="Importe", color_discrete_sequence=["#1d4ed8"])
        fig.update_layout(height=350, margin=dict(t=15, b=15))
        st.plotly_chart(fig, use_container_width=True, key="debt_aging_chart")
    debt_table(branch.rename(columns={"Mayor_90": "Mayor a 90 positivo"}), "debt_branch_table")

    st.markdown("### Clientes y prioridades de cobranza")
    # Prioriza dentro de la selección; no mezcla créditos de distintos clientes.
    priority = prioritize(d)
    level = st.multiselect("Prioridades a mostrar", ["Judicial", "Alta", "Media", "Preventiva", "Revisar", "Sin deuda"], key="debt_priority")
    if level:
        priority = priority[priority["Prioridad"].isin(level)]
    cols = ["Prioridad", "Cliente", "Código", "CUIT", "Sucursal", "Tipo", "Gestión", "Días",
            "Saldo", "A vencer", "Vencido"] + BANDS + ["Crédito máximo", "Acción sugerida"]
    debt_table(priority[cols], "debt_priority_table")
    st.caption("Prioridad por antigüedad de importes vencidos: Alta si hay deuda de 61 días o más; Media si hay vencido; Preventiva si todo está a vencer. Dentro de cada nivel se ordena por importe vencido. Días conserva el dato de origen.")

    with st.expander("Crédito por cliente y tipo de cartera"):
        st.caption("El límite se toma una sola vez por cliente (CUIT, o código si falta CUIT). Si difiere entre registros se muestra el mayor y se señala para revisión. El saldo corresponde a la selección actual.")
        debt_table(credits, "debt_credit_table")
        if (credits["Límites_distintos"] > 1).any():
            st.warning("Hay clientes con límites distintos entre registros; validar el límite aplicable.")
        by_type = d.groupby("Tipo", as_index=False).agg(Deuda=("Deuda", "sum"), Vencido=("Vencido positivo", "sum"))
        debt_table(by_type.sort_values("Vencido", ascending=False), "debt_type_table")
    with st.expander("Control de datos"):
        st.write(f"Saldos sin informar: {int(d['Saldo'].isna().sum())}.")
        st.write(f"Cuentas con diferencia entre saldo y composición: {int((d['Diferencia saldo'].abs() > 1).sum())}.")
        st.write(f"Cuentas con diferencia entre vencido y tramos: {int((d['Diferencia vencido'].abs() > 1).sum())}.")
        st.caption("Tolerancia de conciliación: $1 por cuenta. Los vacíos en tramos se interpretan como cero; los saldos y límites vacíos quedan sin informar.")
        unknown = d[d["Sucursal"].str.contains("sin identificar", regex=False)]
        if not unknown.empty:
            st.warning("Hay sucursales sin identificar; sus registros se conservan en Otras sucursales.")
    st.download_button("Descargar cartera y acciones (CSV)", priority[cols].to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                       file_name="deudores_postventa.csv", mime="text/csv", key="debt_export")
