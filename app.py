import json
import uuid
from datetime import date, datetime
from io import BytesIO

import pandas as pd
import streamlit as st
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

import gspread
from google.oauth2.service_account import Credentials

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image

LOGO_PATH = "logo_ness.png"
NOMBRE_HOJA = "Historial Cotizaciones Ness"

# ---------- Datos fijos de Confecciones Ness (lo que SIEMPRE va igual) ----------
EMPRESA_NOMBRE = "CONFECCIONES NESS"
EMPRESA_CONTACTO = "Nidia Suarez Silva."
EMPRESA_DIRECCION = "CLL14C 24b-59 Riohacha - La Guajira"
EMPRESA_TEL = "Tel: 3013662419"

_styles = getSampleStyleSheet()
ESTILO_NORMAL = ParagraphStyle("normal_ness", parent=_styles["Normal"], fontSize=11, leading=15)
ESTILO_NEGRITA = ParagraphStyle("bold_ness", parent=ESTILO_NORMAL, fontName="Helvetica-Bold")


# ================= UTILIDADES =================

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


# ================= HISTORIAL (Google Sheets) =================

@st.cache_resource
def conectar_hoja():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    info_credenciales = dict(st.secrets["gcp_service_account"])
    credenciales = Credentials.from_service_account_info(info_credenciales, scopes=scopes)
    cliente = gspread.authorize(credenciales)
    return cliente.open(NOMBRE_HOJA).sheet1


def guardar_en_historial(fecha_str, titulo, institucion, prendas_df, frase_resumen, valor_total):
    hoja = conectar_hoja()
    nuevo_id = uuid.uuid4().hex[:8]
    prendas_json = prendas_df.to_dict(orient="records")
    fila = [
        nuevo_id,
        fecha_str,
        titulo,
        institucion,
        json.dumps(prendas_json, ensure_ascii=False),
        frase_resumen,
        valor_total,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    ]
    hoja.append_row(fila)
    return nuevo_id


@st.cache_data(ttl=30)
def obtener_historial():
    hoja = conectar_hoja()
    registros = hoja.get_all_records()
    return list(reversed(registros))  # más reciente primero


def cargar_cotizacion_en_formulario(registro):
    """Prepara los datos de un registro del historial para precargar el formulario."""
    try:
        prendas = pd.DataFrame(json.loads(registro["Prendas"]))
    except Exception:
        prendas = pd.DataFrame([{"Prenda": "Camisa", "Especificaciones": "", "Valor": 0}])

    try:
        fecha_valor = datetime.strptime(registro["Fecha"], "%d/%m/%Y").date()
    except Exception:
        fecha_valor = date.today()

    st.session_state["campo_fecha"] = fecha_valor
    st.session_state["campo_titulo"] = registro.get("Titulo", "")
    st.session_state["campo_institucion"] = registro.get("Institucion", "")
    st.session_state["campo_prendas"] = prendas
    st.session_state["campo_frase"] = registro.get("FraseResumen", "")


# ================= GENERAR PDF =================

def generar_cotizacion(fecha_str, titulo, institucion, prendas_df, frase_resumen=""):
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

    elementos.append(Paragraph(texto_frase, ESTILO_NORMAL))
    elementos.append(Spacer(1, 0.15 * inch))

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
    return buffer, nombre_archivo, valor_total, texto_frase


def enviar_correo(destinatario, asunto, cuerpo, archivo_bytes, nombre_archivo):
    remitente = st.secrets["EMAIL_ADDRESS"]
    clave_app = st.secrets["EMAIL_APP_PASSWORD"]

    msg = MIMEMultipart()
    msg["From"] = remitente
    msg["To"] = ", ".join(destinatario)
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


# ================= VALORES POR DEFECTO DEL FORMULARIO =================

if "campo_fecha" not in st.session_state:
    st.session_state["campo_fecha"] = date.today()
if "campo_titulo" not in st.session_state:
    st.session_state["campo_titulo"] = "COTIZACION – UNIFORME DE PROMOCION"
if "campo_institucion" not in st.session_state:
    st.session_state["campo_institucion"] = ""
if "campo_prendas" not in st.session_state:
    st.session_state["campo_prendas"] = pd.DataFrame([{"Prenda": "Camisa", "Especificaciones": "", "Valor": 0}])
if "campo_frase" not in st.session_state:
    st.session_state["campo_frase"] = ""


# ================= INTERFAZ =================

st.set_page_config(page_title="Cotizaciones Ness", page_icon="🧵", layout="centered")

st.image(LOGO_PATH, width=110)
st.title("Generador de cotizaciones")
st.caption("Confecciones Ness — llena los datos y se genera y envía el documento automáticamente a tu correo.")

# ---- Historial ----
with st.expander("📋 Historial de cotizaciones (ver o editar una anterior)"):
    try:
        historial = obtener_historial()
    except Exception as e:
        historial = []
        st.warning("No se pudo cargar el historial todavía.")
        st.caption(f"Detalle técnico: {e}")

    if not historial:
        st.caption("Todavía no hay cotizaciones guardadas.")
    else:
        opciones = {
            f'{r.get("Institucion", "(sin nombre)")} — {r.get("Fecha", "")} — {r.get("ID", "")}': r
            for r in historial
        }
        seleccion = st.selectbox("Elige una cotización anterior", list(opciones.keys()))
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("Cargar esta cotización para editar", use_container_width=True):
                cargar_cotizacion_en_formulario(opciones[seleccion])
                st.rerun()
        with col_b:
            if st.button("Empezar una cotización nueva", use_container_width=True):
                st.session_state["campo_fecha"] = date.today()
                st.session_state["campo_titulo"] = "COTIZACION – UNIFORME DE PROMOCION"
                st.session_state["campo_institucion"] = ""
                st.session_state["campo_prendas"] = pd.DataFrame([{"Prenda": "Camisa", "Especificaciones": "", "Valor": 0}])
                st.session_state["campo_frase"] = ""
                st.rerun()

# ---- Formulario ----
with st.form("form_cotizacion"):
    col1, col2 = st.columns(2)
    with col1:
        fecha = st.date_input("Fecha", format="DD/MM/YYYY", key="campo_fecha")
    with col2:
        titulo = st.text_input("Título de la cotización", key="campo_titulo")

    institucion = st.text_input(
        "Institución / Cliente",
        placeholder="Ej: INSTITUCION EDUCATIVA SANTA MARIA GORETTI",
        key="campo_institucion",
    )

    st.markdown("**Prendas de la cotización**")
    st.caption("Agrega, edita o borra filas. El valor total se calcula solo.")
    prendas_df = st.data_editor(
        st.session_state["campo_prendas"],
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
        key="campo_frase",
    )

    enviar = st.form_submit_button("Generar y enviar", type="primary", use_container_width=True)

if enviar:
    correo_destino = st.secrets["EMAIL_ADDRESS"]
    correo_destino_2 = st.secrets.get("CORREO_DESTINO_2", "")
    lista_destinatarios = [correo_destino] + ([correo_destino_2] if correo_destino_2 else [])

    prendas_validas = prendas_df.dropna(subset=["Prenda"])
    prendas_validas = prendas_validas[prendas_validas["Prenda"].astype(str).str.strip() != ""]

    if not institucion.strip():
        st.error("Falta el nombre de la institución o cliente.")
    elif prendas_validas.empty:
        st.error("Agrega al menos una prenda con su valor.")
    else:
        fecha_str = fecha.strftime("%d/%m/%Y")
        buffer, nombre_archivo, total, frase_final = generar_cotizacion(
            fecha_str, titulo, institucion, prendas_validas, frase_resumen
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
                destinatario=lista_destinatarios,
                asunto=f"Cotización — {institucion}",
                cuerpo=(
                    f"Buen día,\n\nAdjunto la cotización para {institucion} "
                    f"por un valor total de {formato_pesos(total)} pesos por estudiante.\n\n"
                    f"Confecciones Ness\n{EMPRESA_CONTACTO}\n{EMPRESA_TEL}"
                ),
                archivo_bytes=buffer,
                nombre_archivo=nombre_archivo,
            )
            st.success(f"Correo enviado a {', '.join(lista_destinatarios)} ✅")
        except Exception as e:
            st.warning(
                "El documento se generó bien, pero no se pudo enviar el correo automáticamente. "
                "Puedes descargarlo arriba y adjuntarlo manualmente."
            )
            st.caption(f"Detalle técnico: {e}")

        try:
            guardar_en_historial(fecha_str, titulo, institucion, prendas_validas, frase_final, total)
            obtener_historial.clear()
            st.success("Guardada en el historial 📋")
        except Exception as e:
            st.warning("No se pudo guardar en el historial (la cotización y el correo sí se generaron bien).")
            st.caption(f"Detalle técnico: {e}")
