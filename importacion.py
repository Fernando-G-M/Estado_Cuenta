import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Image, Spacer, KeepTogether
)
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from datetime import datetime

st.title("Estado de Cuenta")

# -------------------------------
# FUNCIONES
# -------------------------------
def to_number(valor):
    if pd.isna(valor):
        return 0
    s = str(valor).replace("$", "").replace(",", "").strip()
    try:
        return float(s)
    except:
        return 0

def convertir_fecha(valor):
    if pd.isna(valor):
        return None

    meses = {
        "ene": "01","feb": "02","mar": "03","abr": "04",
        "may": "05","jun": "06","jul": "07","ago": "08",
        "sep": "09","oct": "10","nov": "11","dic": "12"
    }

    texto = str(valor).lower()

    for mes_txt, mes_num in meses.items():
        if mes_txt in texto:
            texto = texto.replace(mes_txt, mes_num)

    return pd.to_datetime(texto, dayfirst=True, errors='coerce')

def color_deposito(val):
    return "color: green" if val > 0 else ""

def color_retiro(val):
    return "color: red" if val > 0 else ""

# -------------------------------
# PDF
# -------------------------------
def generar_pdf(df, fecha, total_dep, total_ret, saldo_final):
    doc = SimpleDocTemplate("estado_cuenta.pdf", pagesize=letter)
    elementos = []
    styles = getSampleStyleSheet()

    # HEADER
    try:
        barra = Table([[""]], colWidths=[6.5*inch])
        barra.setStyle([
            ("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#003366")),
        ])
        elementos.append(barra)

        logo = Image("logo.png", width=130, height=80)

        info_header = Paragraph(
            "<b><font size=14 color='#003366'>Informacion del estado de cuenta</font></b>",
            styles["Normal"]
        )

        header = Table([[logo, info_header]], colWidths=[2*inch, 4.5*inch])
        elementos.append(header)

    except:
        elementos.append(Paragraph("ESTADO DE CUENTA", styles["Title"]))

    elementos.append(Spacer(1, 12))

    # INFO
    fecha_actual = datetime.now().strftime("%d/%m/%Y")

    info = Table([
        ["Propietario:", "Luis Pascual Martinez Ochoa"],
        ["Fecha emisión:", fecha_actual],
        ["Fecha corte:", fecha.strftime("%d/%m/%Y")]
    ])

    elementos.append(info)
    elementos.append(Spacer(1, 15))

    # RESUMEN
    resumen = Table([
        ["Depósitos", "Retiros", "Saldo"],
        [f"${total_dep:,.2f}", f"${total_ret:,.2f}", f"${saldo_final:,.2f}"]
    ])

    resumen.setStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#003366")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("GRID", (0,0), (-1,-1), 1, colors.black),
    ])

    elementos.append(resumen)
    elementos.append(Spacer(1, 20))

    # FORMATO PARA PDF
    df_pdf = df.copy()
    for col in ["DEPÓSITO", "RETIRO", "SALDO"]:
        df_pdf[col] = df_pdf[col].apply(lambda x: f"{x:,.2f}")

    data = [df_pdf.columns.tolist()] + df_pdf.values.tolist()

    tabla = Table(data, repeatRows=1)

    tabla.setStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#003366")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("ALIGN", (2,1), (-1,-1), "RIGHT"),
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("ROWBACKGROUNDS", (0,1), (-1,-1),
         [colors.white, colors.HexColor("#f4f6f7")]),
        ("FONTSIZE", (0,0), (-1,-1), 8),
    ])

    elementos.append(tabla)
    elementos.append(Spacer(1, 15))

    # -------------------------------
    # OBSERVACIONES (INTELIGENTE)
    # -------------------------------
    titulo_obs = Paragraph(
        "<b>Observaciones de transferencia</b>",
        styles["Heading3"]
    )

    data_obs = [["Fecha", "Descripción", "Depósitos", "Observaciones"]]

    for _ in range(10):
        data_obs.append([
            fecha.strftime("%d/%m/%Y"),
            "",
            "",
            ""
        ])

    tabla_obs = Table(
        data_obs,
        colWidths=[1.2*inch, 2.5*inch, 1.2*inch, 1.6*inch],
        repeatRows=1
    )

    tabla_obs.setStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#003366")),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("ALIGN", (2,1), (2,-1), "RIGHT"),
        ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
        ("ROWBACKGROUNDS", (0,1), (-1,-1),
         [colors.white, colors.HexColor("#f9f9f9")]),
        ("FONTSIZE", (0,0), (-1,-1), 8),
    ])

    bloque_obs = [
        titulo_obs,
        Spacer(1, 10),
        tabla_obs
    ]

    # 👇 CLAVE: usa espacio restante inteligentemente
    elementos.append(KeepTogether(bloque_obs))

    doc.build(elementos)

# -------------------------------
# UI
# -------------------------------
fecha_seleccionada = st.date_input("Selecciona la fecha")
archivo = st.file_uploader("Sube tu Excel", type=["xlsx"])

if archivo is not None:

    df = pd.read_excel(archivo)

    df["FECHA_OK"] = df.iloc[:, 0].apply(convertir_fecha)
    df = df[df["FECHA_OK"].notna()]

    df_filtrado = df[df["FECHA_OK"].dt.date == fecha_seleccionada].copy()

    if df_filtrado.empty:
        st.warning("No hay datos para esa fecha")
    else:
        resultado = []
        saldo = 0
        total_dep = 0
        total_ret = 0

        for _, row in df_filtrado.iterrows():
            fecha = row["FECHA_OK"]
            desc = row["Descripción"]

            dep = to_number(row["Depósitos"])
            ret = to_number(row["Retiros"])

            total_dep += dep
            total_ret += ret
            saldo += (dep - ret)

            resultado.append({
                "FECHA": fecha.strftime("%d/%m/%Y"),
                "DESCRIPCIÓN": desc,
                "DEPÓSITO": dep,
                "RETIRO": ret,
                "SALDO": saldo
            })

        df_resultado = pd.DataFrame(resultado)

        # RESUMEN
        st.subheader("Resumen")
        c1, c2, c3 = st.columns(3)
        c1.metric("Depósitos", f"${total_dep:,.2f}")
        c2.metric("Retiros", f"${total_ret:,.2f}")
        c3.metric("Saldo", f"${saldo:,.2f}")

        # TABLA CON COLORES
        styled = df_resultado.style \
            .map(color_deposito, subset=["DEPÓSITO"]) \
            .map(color_retiro, subset=["RETIRO"]) \
            .format({
                "DEPÓSITO": "${:,.2f}",
                "RETIRO": "${:,.2f}",
                "SALDO": "${:,.2f}"
            })

        st.dataframe(styled, hide_index=True, use_container_width=True)

        # GENERAR ARCHIVOS
        generar_pdf(df_resultado, fecha_seleccionada, total_dep, total_ret, saldo)

        wb = Workbook()
        ws = wb.active

        for r in dataframe_to_rows(df_resultado, index=False, header=True):
            ws.append(r)

        wb.save("estado.xlsx")

        col1, col2 = st.columns(2)

        with open("estado.xlsx", "rb") as f:
            col1.download_button("📥 Descargar Excel", f, file_name="estado.xlsx")

        with open("estado_cuenta.pdf", "rb") as f:
            col2.download_button("📥 Descargar PDF", f, file_name="estado_cuenta.pdf")

        st.success("Reporte listo automáticamente")