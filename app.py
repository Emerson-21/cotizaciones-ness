import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image

LOGO_PATH = "logo_ness.png"

# ---------- Datos fijos de Confecciones Ness (lo que SIEMPRE va igual) ----------
EMPRESA_NOMBRE = "CONFECCIONES NESS"
EMPRESA_CONTACTO = "Nidia Suarez Silva."
EMPRESA_DIRECCION = "CLL14C 24b-59 Riohacha - La Guajira"
EMPRESA_TEL = "Tel: 3013662419"

_styles = getSampleStyleSheet()
ESTILO_NORMAL = ParagraphStyle("normal_ness", parent=_styles["Normal"], fontSize=11, leading=15)
ESTILO_NEGRITA = ParagraphStyle("bold_ness", parent=ESTILO_NORMAL, fontName="Helvetica-Bold")


def formato_pesos(valor: int) -> str:
    """170000 -> $170.000"""
    return "$" + f"{int(valor):,}".replace(",", ".")


def frase_resumen_automatica(n_prendas: int, valor_total: int) -> str:
    plural_prenda = "prendas" if n_prendas != 1 else "prenda"
    texto_num = {
        1: "una", 2: "dos", 3: "tres", 4: "cuatro", 5: "cinco",
        6: "seis", 7: "siete", 8: "ocho", 9: "nueve", 10: "diez",
    }.get(n_prendas, str(n_prendas))
    return (
        f"El uniforme de promoción consta de {texto_num} {plural_prenda}, "
        f"su valor total por estudiante es de {formato_pesos(valor_total)} pesos."
    )


def generar_cotizacion(fecha_str, titulo, institucion, prendas_df, correo_destino, frase_resumen=""):
    """
    prendas_df: DataFrame con columnas Prenda, Especificaciones, Valor
    frase_resumen: si viene vacía, se calcula automática. Si el usuario escribió algo, se usa tal cual.
    Devuelve (bytes_pdf, nombre_archivo, valor_total)
    """
    valor_total = int(prendas_df["Valor"].sum())
    n_prendas = len(prendas_df)

    texto_frase = frase_resumen.strip() if frase_resumen and frase_resumen.strip() else frase_resumen_automatica(n_prendas, valor_total)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        topMargin=0.6 * inch,
        bottomMargin=0.6 * inch,
        leftMargin=0.7 * inch,
        rightMargin=0.7 * inch,
    )

    elementos = []

    # ---- Encabezado: texto a la izquierda, logo a la derecha ----
    texto_izquierda = [
        Paragraph(fecha_str, ESTILO_NORMAL),
        Paragraph(titulo, ESTILO_NEGRITA),
        Paragraph(institucion.upper(), ESTILO_NEGRITA),
        Paragraph(EMPRESA_NOMBRE, ESTILO_NEGRITA),
        Paragraph(EMPRESA_CONTACTO, ESTILO_NORMAL),
        Paragraph(EMPRESA_DIRECCION, ESTILO_NORMAL),
        Paragraph(EMPRESA_TEL, ESTILO_NORMAL),
    ]

    try:
        logo = Image(LOGO_PATH, width=1.3 * inch, height=1.38 * inch)
    except Exception:
        logo = Paragraph("", ESTILO_NORMAL)

    header_table = Table([[texto_izquierda, logo]], colWidths=[4.9 * inch, 1.6 * inch])
    header_table.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (1, 0), (1, 0), "CENTER"),
            ]
        )
    )
    elementos.append(header_table)
    elementos.append(Spacer(1, 0.2 * inch))

    # ---- Frase resumen (automática o escrita por el usuario) ----
    elementos.append(Paragraph(texto_frase, ESTILO_NORMAL))
    elementos.append(Spacer(1, 0.15 * inch))

    # ---- Tabla de prendas ----
    datos_tabla = [["Prenda", "Especificaciones", "Valor unitario"]]
    for _, fila in prendas_df.iterrows():
        datos_tabla.append(
            [
                str(fila["Prenda"]),
                Paragraph(str(fila["Especificaciones"]), ESTILO_NORMAL),
                formato_pesos(fila["Valor"]),
            ]
        )
    datos_tabla.append(["Valor total", "", formato_pesos(valor_total)])

    tabla = Table(datos_tabla, colWidths=[1.3 * inch, 3.3 * inch, 1.6 * inch])
    ultima_fila = len(datos_tabla) - 1
    tabla.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.75, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9D9D9")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("ALIGN", (0, 1), (0, ultima_fila - 1), "CENTER"),
                ("ALIGN", (2, 1), (2, ultima_fila), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("SPAN", (0, ultima_fila), (1, ultima_fila)),
                ("FONTNAME", (0, ultima_fila), (-1, ultima_fila), "Helvetica-Bold"),
            ]
        )
    )
    elementos.append(tabla)
    elementos.append(Spacer(1, 0.25 * inch))

    # ---- Nota ----
    elementos.append(
        Paragraph(
            "<b>Nota:</b> La fecha de entrega, esta estimada en 30 días calendario "
            "a partir del abono inicial del 50%.",
            ESTILO_NORMAL,
        )
    )

    doc.build(elementos)
    buffer.seek(0)

    nombre_archivo = f"Cotizacion_{institucion.strip().replace(' ', '_')}.pdf"
    return buffer, nombre_archivo, valor_total


def enviar_correo(destinatario, asunto, cuerpo, archivo_bytes, nombre_archivo):
    remitente = st.secrets["EMAIL_ADDRESS"]
    clave_app = st.secrets["EMAIL_APP_PASSWORD"]

    msg = MIMEMultipart()
    msg["From"] = remitente
    msg["To"] = destinatario
    msg["Subject"] = asunto
    msg.attach(MIMEText(cuerpo, "plain"))

    part = MIMEBase("application", "pdf")
    part.set_payload(archivo_bytes.getvalue())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename={nombre_archivo}")
    msg.attach(part)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(remitente, clave_app)
        server.sendmail(remitente, destinatario, msg.as_string())


# ================= INTERFAZ =================

st.set_page_config(page_title="Cotizaciones Ness", page_icon="🧵", layout="centered")

st.image(LOGO_PATH, width=110)
st.title("Generador de cotizaciones")
st.caption("Confecciones Ness — llena los datos y se genera y envía el documento automáticamente a tu correo.")

with st.form("form_cotizacion"):
    col1, col2 = st.columns(2)
    with col1:
        fecha = st.date_input("Fecha", value=date.today(), format="DD/MM/YYYY")
    with col2:
        titulo = st.text_input("Título de la cotización", value="COTIZACION – UNIFORME DE PROMOCION")

    institucion = st.text_input("Institución / Cliente", placeholder="Ej: INSTITUCION EDUCATIVA SANTA MARIA GORETTI")

    st.markdown("**Prendas de la cotización**")
    st.caption("Agrega, edita o borra filas. El valor total se calcula solo.")
    prendas_default = pd.DataFrame(
        [
            {"Prenda": "Camisa", "Especificaciones": "", "Valor": 0},
        ]
    )
    prendas_df = st.data_editor(
        prendas_default,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Valor": st.column_config.NumberColumn("Valor unitario", min_value=0, step=1000, format="%d"),
        },
        key="editor_prendas",
    )

    st.markdown("**Frase resumen (la que va justo antes de la tabla)**")
    frase_resumen = st.text_area(
        "Escríbela como la necesites para esta cotización. Si la dejas vacía, se genera sola.",
        placeholder='Ej: "El uniforme de promoción consta de dos prendas, su valor total por estudiante es de $120.000 pesos."',
        height=80,
    )

    enviar = st.form_submit_button("Generar y enviar", type="primary", use_container_width=True)

if enviar:
    correo_destino = st.secrets["EMAIL_ADDRESS"]

    prendas_validas = prendas_df.dropna(subset=["Prenda"])
    prendas_validas = prendas_validas[prendas_validas["Prenda"].astype(str).str.strip() != ""]

    if not institucion.strip():
        st.error("Falta el nombre de la institución o cliente.")
    elif prendas_validas.empty:
        st.error("Agrega al menos una prenda con su valor.")
    else:
        fecha_str = fecha.strftime("%d/%m/%Y")
        buffer, nombre_archivo, total = generar_cotizacion(
            fecha_str, titulo, institucion, prendas_validas, correo_destino, frase_resumen
        )
        st.success(f"Cotización generada. Valor total: {formato_pesos(total)}")

        st.download_button(
            "Descargar PDF",
            data=buffer,
            file_name=nombre_archivo,
            mime="application/pdf",
            use_container_width=True,
        )

        try:
            buffer.seek(0)
            enviar_correo(
                destinatario=correo_destino,
                asunto=f"Cotización — {institucion}",
                cuerpo=(
                    f"Buen día,\n\nAdjunto la cotización para {institucion} "
                    f"por un valor total de {formato_pesos(total)} pesos por estudiante.\n\n"
                    f"Confecciones Ness\n{EMPRESA_CONTACTO}\n{EMPRESA_TEL}"
                ),
                archivo_bytes=buffer,
                nombre_archivo=nombre_archivo,
            )
            st.success(f"Correo enviado a {correo_destino} ✅")
        except Exception as e:
            st.warning(
                "El documento se generó bien, pero no se pudo enviar el correo automáticamente. "
                "Puedes descargarlo arriba y adjuntarlo manualmente."
            )
            st.caption(f"Detalle técnico: {e}")