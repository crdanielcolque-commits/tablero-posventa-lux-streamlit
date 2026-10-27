"""Saldos acreedores y cruces de cuentas de clientes, sin compensaciones automáticas."""
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from debtors import prepare_debtors, normalized, money, BANDS


def credit_crosscheck(all_data, selected_negative):
    """Cruce global solo de los clientes con saldo negativo en la selección."""
    ids = selected_negative["Cliente ID"].unique()
    related = all_data[all_data["Cliente ID"].isin(ids)].copy()
    rows = []
    for customer_id, g in related.groupby("Cliente ID", sort=False):
        positive = g["Saldo"].clip(lower=0).sum()
        credit = -g["Saldo"].clip(upper=0).sum()
        missing = int(g["Saldo"].isna().sum())
        has_both = positive > 0 and credit > 0
        codes = g["Código"].nunique()
        branches = g["Sucursal"].nunique()
        selected = selected_negative[selected_negative["Cliente ID"] == customer_id]
        net = g["Saldo"].sum(min_count=1) if not missing else np.nan
        if missing:
            case = "Información incompleta"
            action = "Completar saldos faltantes antes de evaluar compensación o devolución."
        elif has_both:
            case = "Saldos de ambos signos"
            action = "Conciliar facturas y recibos entre cuentas y sucursales; validar titularidad y aplicación antes de proponer compensaciones."
        elif codes > 1 or branches > 1:
            case = "A favor en varias cuentas"
            action = "Confirmar que las cuentas pertenecen al mismo titular y revisar anticipos, notas de crédito y cobros sin aplicar."
        else:
            case = "A favor sin deuda positiva"
            action = "Identificar el comprobante de origen y confirmar si corresponde aplicar a futuras facturas o tramitar devolución según procedimiento."
        if "judicial" in " ".join(g["Gestión"].map(normalized)):
            action += " Coordinar con el responsable de gestión judicial."
        rows.append({
            "Cliente ID": customer_id, "Cliente": g["Cliente"].iloc[0],
            "CUIT": " / ".join(sorted(set(g["CUIT"]) - {""})),
            "Códigos": " / ".join(sorted(set(g["Código"]))),
            "Sucursales": " / ".join(sorted(set(g["Sucursal"]))),
            "Cantidad códigos": codes, "Cantidad sucursales": branches,
            "A favor seleccionado": -selected["Saldo"].sum(),
            "A favor global": credit, "Deuda positiva global": positive,
            "Saldo neto global": net,
            "Cruce potencial": min(positive, credit) if not missing else np.nan,
            "A favor neto potencial": max(-net, 0) if not missing else np.nan,
            "Deuda neta potencial": max(net, 0) if not missing else np.nan,
            "Saldos faltantes": missing, "Caso": case, "Acción sugerida": action,
        })
    columns = ["Cliente ID", "Cliente", "CUIT", "Códigos", "Sucursales", "Cantidad códigos",
               "Cantidad sucursales", "A favor seleccionado", "A favor global", "Deuda positiva global",
               "Saldo neto global", "Cruce potencial", "A favor neto potencial", "Deuda neta potencial",
               "Saldos faltantes", "Caso", "Acción sugerida"]
    return pd.DataFrame(rows, columns=columns).sort_values("A favor seleccionado", ascending=False)


def credit_table(d, key):
    monetary = ["Saldo", "A vencer", "Vencido"] + BANDS + [
        "A favor seleccionado", "A favor global", "Deuda positiva global", "Saldo neto global",
        "Cruce potencial", "A favor neto potencial", "Deuda neta potencial", "Saldo a favor", "Importe"]
    formats = {c: st.column_config.NumberColumn(c, format="$ %.2f") for c in monetary if c in d}
    st.dataframe(d, hide_index=True, use_container_width=True, column_config=formats, key=key)


def render_credit_balances_tab(raw, sheet_found):
    st.markdown("## Saldos a favor y conciliación")
    st.caption("Origen: Deudores Postventa. Foto de cartera con filtros propios; no depende del mes, semana ni garantías del P&L.")
    st.info("Un saldo negativo indica una posición acreedora en esta base. La hoja no permite confirmar si proviene de anticipos, cobros sin aplicar, notas de crédito o errores, ni si corresponde devolver dinero. La conciliación requiere revisar los comprobantes.")
    if not sheet_found or raw.empty:
        st.info("No hay una hoja Deudores Postventa con datos disponibles.")
        return
    try:
        all_data = prepare_debtors(raw)
    except ValueError as exc:
        st.error(str(exc))
        return
    negative = all_data[all_data["Saldo"] < 0].copy()
    if negative.empty:
        st.info("No hay cuentas con saldo negativo en la hoja.")
        return
    with st.expander("Filtros de saldos a favor", expanded=True):
        a, b, c = st.columns(3)
        branches = a.multiselect("Sucursales", sorted(negative["Sucursal"].unique()), key="credit_branches")
        types = b.multiselect("Tipo de cliente", sorted(negative["Tipo"].unique()), key="credit_types")
        management = c.multiselect("Gestión", sorted(negative["Gestión"].unique()), key="credit_management")
        a, b = st.columns(2)
        zone = a.selectbox("Zona", ["Todas", "Posventa Centro", "Posventa Interior", "Otras sucursales"], key="credit_zone")
        query = b.text_input("Buscar cliente, código o CUIT", key="credit_search").strip()
    d = negative.copy()
    for col, chosen in [("Sucursal", branches), ("Tipo", types), ("Gestión", management)]:
        if chosen:
            d = d[d[col].isin(chosen)]
    if zone != "Todas":
        d = d[d["Zona"] == zone]
    if query:
        mask = d[["Cliente", "Código", "CUIT"]].apply(lambda col: col.map(normalized).str.contains(normalized(query), regex=False)).any(axis=1)
        d = d[mask]
    if d.empty:
        st.info("No hay saldos negativos para los filtros seleccionados.")
        return
    d["Saldo a favor"] = -d["Saldo"]
    cross = credit_crosscheck(all_data, d)
    a, b, c, e = st.columns(4)
    a.metric("Saldos a favor seleccionados", money(d["Saldo a favor"].sum()))
    b.metric("Clientes con saldo a favor", d["Cliente ID"].nunique())
    c.metric("Registros cliente / sucursal", len(d))
    e.metric("Clientes con ambos signos", int(((cross["A favor global"] > 0) & (cross["Deuda positiva global"] > 0)).sum()))
    st.caption("Los importes a favor se muestran positivos para facilitar la lectura. En el detalle se conserva también el saldo negativo original.")

    st.markdown("### Qué revisar primero")
    first = cross.iloc[0]
    st.write(f"Mayor importe seleccionado: {first['Cliente']}, {money(first['A favor seleccionado'])}. Revisar comprobantes y aplicación de fondos.")
    potential = cross["Cruce potencial"].sum()
    st.write(f"Cruce potencial con deuda del mismo cliente: {money(potential)}. Es una simulación por titular, no una compensación realizada ni autorizada.")
    multiple = cross[(cross["Cantidad códigos"] > 1) | (cross["Cantidad sucursales"] > 1)]
    st.write(f"{len(multiple)} clientes aparecen en varios códigos o sucursales. Verificar su unificación y evitar cobrar o devolver importes sin conciliar.")
    no_cuit = d[d["CUIT"].str.replace(r"\D", "", regex=True).eq("")]
    if not no_cuit.empty:
        st.write(f"{no_cuit['Cliente ID'].nunique()} clientes seleccionados no tienen CUIT informado. El cruce usa su código; no se agrupan por similitud del nombre.")
    st.caption("Se priorizan los mayores importes a favor. La antigüedad informada corresponde a la cuenta de origen y no acredita la fecha de un anticipo.")

    st.markdown("### Distribución de saldos a favor")
    branch = d.groupby("Sucursal", as_index=False).agg(**{"Saldo a favor": ("Saldo a favor", "sum"), "Registros": ("Cliente ID", "size"), "Clientes": ("Cliente ID", "nunique")}).sort_values("Saldo a favor", ascending=False)
    by_type = d.groupby("Tipo", as_index=False)["Saldo a favor"].sum().sort_values("Saldo a favor", ascending=False)
    a, b = st.columns(2)
    with a:
        fig = px.bar(branch, x="Sucursal", y="Saldo a favor", color_discrete_sequence=["#1d4ed8"])
        fig.update_layout(height=350, margin=dict(t=15, b=15))
        st.plotly_chart(fig, use_container_width=True, key="credit_branch_chart")
    with b:
        top = cross.head(15).sort_values("A favor seleccionado")
        fig = px.bar(top, x="A favor seleccionado", y="Cliente", orientation="h", color_discrete_sequence=["#0f766e"])
        fig.update_layout(height=400, margin=dict(t=15, b=15))
        st.plotly_chart(fig, use_container_width=True, key="credit_customer_chart")
    credit_table(branch, "credit_branch_table")
    with st.expander("Saldos a favor por tipo de cliente"):
        credit_table(by_type, "credit_type_table")

    st.markdown("### Cruce global por cliente")
    st.caption("Los filtros seleccionan cuentas negativas. Para esos mismos clientes, este cruce consulta TODAS sus cuentas en la hoja, incluso fuera de las sucursales o tipos filtrados. CUIT como clave; código cuando falta CUIT. No se compensan clientes distintos.")
    show = cross.drop(columns="Cliente ID")
    credit_table(show, "credit_cross_table")
    st.caption("Cruce potencial = menor entre deuda positiva global y saldo a favor global del mismo cliente. Saldo neto global = deuda positiva menos saldo a favor. No confirma que la compensación sea procedente; los casos con saldos faltantes quedan sin neto.")
    related = all_data[all_data["Cliente ID"].isin(cross["Cliente ID"])]
    with st.expander("Cuentas relacionadas de los clientes seleccionados"):
        credit_table(related[["Cliente", "Código", "CUIT", "Sucursal", "Tipo", "Gestión", "Saldo", "A vencer", "Vencido"] + BANDS], "credit_related_table")
    with st.expander("Detalle de cuentas con saldo negativo", expanded=False):
        credit_table(d[["Cliente", "Código", "CUIT", "Sucursal", "Tipo", "Gestión", "Días", "Saldo", "Saldo a favor", "A vencer", "Vencido"] + BANDS].sort_values("Saldo a favor", ascending=False), "credit_detail_table")
    with st.expander("Conciliación y datos a validar"):
        positive_total = all_data.loc[all_data["Saldo"] > 0, "Saldo"].sum()
        negative_total = -all_data.loc[all_data["Saldo"] < 0, "Saldo"].sum()
        st.write(f"Hoja completa: deuda positiva {money(positive_total)} menos saldos a favor {money(negative_total)} = saldo neto {money(positive_total-negative_total)}.")
        st.caption("Este puente es global de toda la hoja, no solo de la selección. El tab Deudores mantiene la deuda positiva sin descontar saldos a favor de otros clientes.")
        st.write(f"Registros seleccionados con diferencias entre saldo y composición: {int((d['Diferencia saldo'].abs() > 1).sum())}. Tolerancia: $1 por registro.")
        duplicate = all_data.duplicated(subset=["Código", "Código sucursal"], keep=False)
        st.write(f"Registros repetidos por código y sucursal en la hoja: {int(duplicate.sum())}. Revisar si son movimientos o duplicados antes de interpretar los totales.")
        st.write("Para confirmar el origen: extracto de cuenta corriente, recibos, facturas y notas de crédito, fecha de cada comprobante y su imputación.")
    a, b = st.columns(2)
    with a:
        st.download_button("Descargar cruces y acciones (CSV)", show.to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                           file_name="saldos_a_favor_cruces.csv", mime="text/csv", key="credit_export_cross")
    with b:
        st.download_button("Descargar cuentas negativas (CSV)", d.drop(columns=["Cliente ID"]).to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig"),
                           file_name="saldos_negativos_detalle.csv", mime="text/csv", key="credit_export_detail")
