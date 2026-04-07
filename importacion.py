import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Image, Spacer, KeepTogether, HRFlowable
)
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
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

def es_pago_terceros(desc):
    return "pago a terceros" in str(desc).lower()

# -------------------------------
# PDF HORIZONTAL MEJORADO
# -------------------------------
def generar_pdf(df, fecha, total_dep, total_ret, saldo_final):
    AZUL_OSCURO = colors.HexColor("#0D2B55")
    AZUL_MEDIO  = colors.HexColor("#1A4A8A")
    AZUL_CLARO  = colors.HexColor("#E8F0FB")
    GRIS_LINEA  = colors.HexColor("#CBD5E1")
    VERDE       = colors.HexColor("#16A34A")
    ROJO        = colors.HexColor("#DC2626")
    BLANCO      = colors.white
    GRIS_TEXTO  = colors.HexColor("#374151")

    doc = SimpleDocTemplate(
        "estado_cuenta.pdf",
        pagesize=landscape(letter),
        rightMargin=0.6*inch,
        leftMargin=0.6*inch,
        topMargin=0.5*inch,
        bottomMargin=0.5*inch
    )

    elementos = []

    estilo_label_w = ParagraphStyle("lw", fontSize=7,  textColor=colors.HexColor("#A8C4E8"), fontName="Helvetica")
    estilo_valor_w = ParagraphStyle("vw", fontSize=10, textColor=BLANCO, fontName="Helvetica-Bold")
    estilo_titulo  = ParagraphStyle("ti", fontSize=18, textColor=BLANCO, fontName="Helvetica-Bold", leading=22)
    estilo_sub     = ParagraphStyle("su", fontSize=9,  textColor=colors.HexColor("#A8C4E8"), fontName="Helvetica")
    estilo_seccion = ParagraphStyle("se", fontSize=9,  textColor=AZUL_OSCURO, fontName="Helvetica-Bold", spaceBefore=4)

    fecha_actual = datetime.now().strftime("%d/%m/%Y")

    # HEADER
    try:
        logo_img = Image("logo.png", width=90, height=55)
    except:
        logo_img = Paragraph("<b><font color='white' size=12>EC</font></b>",
                             ParagraphStyle("x", textColor=BLANCO))

    col_logo  = Table([[logo_img]], colWidths=[1.2*inch])
    col_texto = Table([
        [Paragraph("Estado de Cuenta", estilo_titulo)],
        [Paragraph("Resumen de movimientos bancarios", estilo_sub)]
    ], colWidths=[5*inch])

    info_right = Table([
        [Paragraph("Propietario",               estilo_label_w)],
        [Paragraph("Luis Pascual Martinez Ochoa", estilo_valor_w)],
        [Spacer(1, 4)],
        [Paragraph("Fecha de emision",          estilo_label_w)],
        [Paragraph(fecha_actual,                estilo_valor_w)],
        [Spacer(1, 4)],
        [Paragraph("Fecha de corte",            estilo_label_w)],
        [Paragraph(fecha.strftime("%d/%m/%Y"),  estilo_valor_w)],
    ], colWidths=[2.5*inch])

    header_inner = Table(
        [[col_logo, col_texto, info_right]],
        colWidths=[1.3*inch, 5.7*inch, 2.8*inch]
    )
    header_inner.setStyle([
        ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING",  (0,0), (-1,-1), 8),
        ("RIGHTPADDING", (0,0), (-1,-1), 8),
    ])

    header_outer = Table([[header_inner]], colWidths=[9.8*inch])
    header_outer.setStyle([
        ("BACKGROUND",    (0,0), (-1,-1), AZUL_OSCURO),
        ("TOPPADDING",    (0,0), (-1,-1), 12),
        ("BOTTOMPADDING", (0,0), (-1,-1), 12),
        ("LEFTPADDING",   (0,0), (-1,-1), 0),
        ("RIGHTPADDING",  (0,0), (-1,-1), 0),
    ])

    elementos.append(header_outer)
    elementos.append(Spacer(1, 14))

    # TARJETAS RESUMEN
    def tarjeta(titulo, valor, color_val):
        t = Table([
            [Paragraph(titulo, ParagraphStyle("ct", fontSize=7, textColor=colors.HexColor("#6B7280"), fontName="Helvetica"))],
            [Paragraph(valor,  ParagraphStyle("cv", fontSize=13, textColor=color_val, fontName="Helvetica-Bold"))],
        ], colWidths=[3*inch])
        t.setStyle([
            ("BACKGROUND",    (0,0), (-1,-1), BLANCO),
            ("BOX",           (0,0), (-1,-1), 0.5, GRIS_LINEA),
            ("TOPPADDING",    (0,0), (-1,-1), 8),
            ("BOTTOMPADDING", (0,0), (-1,-1), 8),
            ("LEFTPADDING",   (0,0), (-1,-1), 12),
            ("RIGHTPADDING",  (0,0), (-1,-1), 12),
        ])
        return t

    resumen_row = Table([[
        tarjeta("DEPOSITOS",  f"${total_dep:,.2f}",    VERDE),
        tarjeta("RETIROS",    f"${total_ret:,.2f}",    ROJO),
        tarjeta("SALDO",      f"${saldo_final:,.2f}",  AZUL_MEDIO),
    ]], colWidths=[3.26*inch, 3.26*inch, 3.26*inch])
    resumen_row.setStyle([
        ("LEFTPADDING",  (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
    ])

    elementos.append(resumen_row)
    elementos.append(Spacer(1, 14))

    # TABLA PRINCIPAL
    elementos.append(Paragraph("Detalle de movimientos", estilo_seccion))
    elementos.append(Spacer(1, 5))

    df_pdf = df.copy()
    for col in ["DEPÓSITO", "RETIRO", "SALDO"]:
        df_pdf[col] = df_pdf[col].apply(lambda x: f"${x:,.2f}")

    encabezados = [
        Paragraph(c, ParagraphStyle("th", fontSize=8, textColor=BLANCO,
                                    fontName="Helvetica-Bold", alignment=TA_CENTER))
        for c in df_pdf.columns.tolist()
    ]

    data = [encabezados] + df_pdf.values.tolist()

    row_colors = []
    for i in range(1, len(data)):
        bg = BLANCO if i % 2 == 0 else AZUL_CLARO
        row_colors.append(("BACKGROUND", (0, i), (-1, i), bg))

    tabla = Table(
        data,
        colWidths=[1*inch, 2.9*inch, 0.8*inch, 0.8*inch, 0.8*inch, 2.9*inch],
        repeatRows=1
    )
    tabla.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),  (-1,0),  AZUL_OSCURO),
        ("TEXTCOLOR",     (0,0),  (-1,0),  BLANCO),
        ("ALIGN",         (0,0),  (-1,0),  "CENTER"),
        ("ALIGN",         (2,1),  (-1,-1), "RIGHT"),
        ("ALIGN",         (0,1),  (1,-1),  "LEFT"),
        ("FONTNAME",      (0,0),  (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0),  (-1,-1), 8),
        ("ROWHEIGHT",     (0,0),  (-1,-1), 16),
        ("GRID",          (0,0),  (-1,-1), 0.5, GRIS_LINEA),
        ("LINEBELOW",     (0,0),  (-1,0),  1.5, AZUL_MEDIO),
        ("LEFTPADDING",   (0,0),  (-1,-1), 6),
        ("RIGHTPADDING",  (0,0),  (-1,-1), 6),
        ("TOPPADDING",    (0,0),  (-1,-1), 3),
        ("BOTTOMPADDING", (0,0),  (-1,-1), 3),
        ("VALIGN",        (0,0),  (-1,-1), "MIDDLE"),
        ("BOX",           (0,0),  (-1,-1), 1, colors.HexColor("#0D2B55")),
    ] + row_colors))

    elementos.append(tabla)
    elementos.append(Spacer(1, 16))

    # TABLA OBSERVACIONES
    elementos.append(HRFlowable(width="100%", thickness=0.5, color=GRIS_LINEA))
    elementos.append(Spacer(1, 8))
    elementos.append(Paragraph("Observaciones de transferencia", estilo_seccion))
    elementos.append(Spacer(1, 6))

    enc_obs = [
        Paragraph(c, ParagraphStyle("tho", fontSize=8, textColor=BLANCO,
                                    fontName="Helvetica-Bold", alignment=TA_CENTER))
        for c in ["Fecha", "Descripcion", "Depositos", "Observaciones"]
    ]

    data_obs = [enc_obs]
    for _ in range(10):
        data_obs.append([fecha.strftime("%d/%m/%Y"), "", "", ""])

    obs_row_colors = []
    for i in range(1, len(data_obs)):
        bg = BLANCO if i % 2 == 0 else AZUL_CLARO
        obs_row_colors.append(("BACKGROUND", (0, i), (-1, i), bg))

    tabla_obs = Table(
        data_obs,
        colWidths=[1*inch, 4*inch, 1.1*inch, 3.5*inch],
        repeatRows=1
    )
    tabla_obs.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),  (-1,0),  AZUL_OSCURO),
        ("TEXTCOLOR",     (0,0),  (-1,0),  BLANCO),
        ("ALIGN",         (0,0),  (-1,0),  "CENTER"),
        ("ALIGN",         (2,1),  (2,-1),  "RIGHT"),
        ("FONTNAME",      (0,0),  (-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0),  (-1,-1), 8),
        ("ROWHEIGHT",     (0,0),  (-1,-1), 16),
        ("GRID",          (0,0),  (-1,-1), 0.5, GRIS_LINEA),
        ("LINEBELOW",     (0,0),  (-1,0),  1.5, AZUL_MEDIO),
        ("LEFTPADDING",   (0,0),  (-1,-1), 6),
        ("RIGHTPADDING",  (0,0),  (-1,-1), 6),
        ("TOPPADDING",    (0,0),  (-1,-1), 3),
        ("BOTTOMPADDING", (0,0),  (-1,-1), 3),
        ("VALIGN",        (0,0),  (-1,-1), "MIDDLE"),
        ("BOX",           (0,0),  (-1,-1), 1, colors.HexColor("#0D2B55")),
    ] + obs_row_colors))

    elementos.append(KeepTogether([tabla_obs]))
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

    # FILTRAR PAGO A TERCEROS
    df_filtrado = df_filtrado[
        ~df_filtrado["Descripción"].apply(es_pago_terceros)
    ].copy()

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
                "DESCRIPCIÓN": str(desc)[:40] + "..." if len(str(desc)) > 40 else str(desc),
                "DEPÓSITO": dep,
                "RETIRO": ret,
                "SALDO": saldo
            })

        df_resultado = pd.DataFrame(resultado)

        # RESUMEN
        st.subheader("Resumen")
        c1, c2, c3 = st.columns(3)
        c1.metric("Depósitos", f"${total_dep:,.2f}")
        c2.metric("Retiros",   f"${total_ret:,.2f}")
        c3.metric("Saldo",     f"${saldo:,.2f}")

        # TABLA CON COLORES
        st.subheader("Detalle de movimientos")
        styled = df_resultado.style \
            .map(color_deposito, subset=["DEPÓSITO"]) \
            .map(color_retiro,   subset=["RETIRO"]) \
            .format({
                "DEPÓSITO": "${:,.2f}",
                "RETIRO":   "${:,.2f}",
                "SALDO":    "${:,.2f}"
            })
        st.dataframe(styled, hide_index=True, use_container_width=True)

        # Para el PDF agregamos columna OBSERVACIONES vacia
        df_pdf = df_resultado.copy()
        df_pdf["OBSERVACIONES"] = ""

        # GENERAR ARCHIVOS
        generar_pdf(df_pdf, fecha_seleccionada, total_dep, total_ret, saldo)

        wb = Workbook()
        ws = wb.active
        for r in dataframe_to_rows(df_resultado, index=False, header=True):
            ws.append(r)
        wb.save("estado.xlsx")

        col1, col2 = st.columns(2)
        with open("estado.xlsx", "rb") as f:
            col1.download_button("Descargar Excel", f, file_name="estado.xlsx")
        with open("estado_cuenta.pdf", "rb") as f:
            col2.download_button("Descargar PDF", f, file_name="estado_cuenta.pdf")

        st.success("Reporte listo automaticamente")