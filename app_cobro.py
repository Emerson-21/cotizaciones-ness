import json
import re
import smtplib
import unicodedata
import uuid
from datetime import date, datetime
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO
from xml.sax.saxutils import escape

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image

# set_page_config debe ser la primera instrucción de Streamlit
st.set_page_config(page_title="Cuentas de cobro Ness", page_icon="🧵", layout="centered")

LOGO_PATH = "logo_ness.png"
NOMBRE_HOJA = "Historial Cuentas de Cobro Ness"

# ---------- Datos fijos (lo que SIEMPRE va igual) ----------
COBRADOR_NOMBRE = "Nidia Suarez Silva"
COBRADOR_NIT = "22.591.062-1"
EMPRESA_TEL = "Tel: 3013662419"
CORREO_FACTURACION = "tallerness@gmail.com"
NOTA_POR_DEFECTO = f"Se prefiere la facturación de forma electrónica al correo {CORREO_FACTURACION}."
MEDIOS_PAGO = ["Nequi", "Daviplata", "Cuenta de ahorros", "Cuenta corriente", "Otro"]
ESTADOS = ["Pendiente", "Abonada", "Pagada"]

_styles = getSampleStyleSheet()
ESTILO_NORMAL = ParagraphStyle("normal_ness", parent=_styles["Normal"], fontSize=11, leading=15)
ESTILO_NEGRITA = ParagraphStyle("bold_ness", parent=ESTILO_NORMAL, fontName="Helvetica-Bold")
ESTILO_CENTRO = ParagraphStyle("centro_ness", parent=ESTILO_NORMAL, alignment=TA_CENTER)


# ================= UTILIDADES =================

def formato_pesos(valor) -> str:
    """170000 -> $ 170.000"""
    return "$ " + f"{int(valor):,}".replace(",", ".")


def texto(valor) -> str:
    """Convierte una celda (que puede venir vacía/NaN) en texto limpio."""
    return "" if pd.isna(valor) else str(valor).strip()


def a_entero(valor) -> int:
    """Convierte lo que venga de la hoja en entero (0 si está vacío o no se puede)."""
    try:
        return int(float(valor))
    except (TypeError, ValueError):
        return 0


def limpiar_nombre(t: str) -> str:
    """Deja solo letras, números y guiones bajos (para nombres de archivo)."""
    t = unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode()
    return re.sub(r"[^A-Za-z0-9_-]+", "_", t).strip("_")


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


def siguiente_consecutivo() -> int:
    """Lee la hoja y devuelve el último consecutivo + 1. Si no puede leerla, devuelve 1."""
    try:
        hoja = conectar_hoja()
        encabezados = hoja.row_values(1)
        columna = encabezados.index("Consecutivo") + 1
        numeros = [int(v) for v in hoja.col_values(columna)[1:] if str(v).strip().isdigit()]
        return max(numeros, default=0) + 1
    except Exception:
        return 1


def guardar_en_historial(consecutivo, fecha_str, cliente, concepto, items, abono, total, pagos):
    hoja = conectar_hoja()
    encabezados = hoja.row_values(1)
    if not encabezados:
        raise ValueError("La hoja no tiene encabezados en la fila 1.")

    datos = {
        "ID": "C" + uuid.uuid4().hex[:7],
        "Consecutivo": int(consecutivo),
        "Fecha": fecha_str,
        "Cliente": cliente,
        "Concepto": concepto,
        "Prendas": json.dumps(items.to_dict(orient="records"), ensure_ascii=False, default=int),
        "Abono": int(abono),
        "ValorTotal": int(total),
        "DatosPago": json.dumps(
            [{"Medio": m, "Detalle": d} for m, d in pagos], ensure_ascii=False
        ),
        "Estado": "Pendiente",
        "Creado": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    # Se arma la fila según el orden de los encabezados de la hoja
    hoja.append_row([datos.get(h, "") for h in encabezados])


@st.cache_data(ttl=30)
def obtener_historial():
    hoja = conectar_hoja()
    registros = hoja.get_all_records()
    return list(reversed(registros))  # más reciente primero


def actualizar_estado(id_registro, nuevo_estado):
    hoja = conectar_hoja()
    encabezados = hoja.row_values(1)
    col_id = encabezados.index("ID") + 1
    col_estado = encabezados.index("Estado") + 1
    celda = hoja.find(str(id_registro), in_column=col_id)
    if celda is None:
        raise ValueError("No se encontró esa cuenta de cobro en la hoja.")
    hoja.update_cell(celda.row, col_estado, nuevo_estado)


def etiqueta_registro(r) -> str:
    try:
        numero = f"{int(r.get('Consecutivo')):03d}"
    except (TypeError, ValueError):
        numero = str(r.get("Consecutivo", "?"))
    return (
        f"No. {numero} — {r.get('Cliente', '(sin nombre)')} — "
        f"{r.get('Fecha', '')} — {r.get('Estado', '')} — {r.get('ID', '')}"
    )


# ================= CORREO =================

def enviar_correo(destinatarios, asunto, cuerpo, pdf_bytes, nombre_archivo):
    remitente = st.secrets["EMAIL_ADDRESS"]
    clave_app = st.secrets["EMAIL_APP_PASSWORD"]

    msg = MIMEMultipart()
    msg["From"] = remitente
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = Header(asunto, "utf-8")
    msg.attach(MIMEText(cuerpo, "plain", "utf-8"))

    parte = MIMEBase("application", "pdf")
    parte.set_payload(pdf_bytes)
    encoders.encode_base64(parte)
    parte.add_header("Content-Disposition", "attachment", filename=nombre_archivo)
    msg.attach(parte)

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(remitente, clave_app)
        server.sendmail(remitente, destinatarios, msg.as_string())


# ================= GENERAR PDF =================

def generar_cuenta_cobro(consecutivo, fecha_str, cliente, concepto, items_df, abono, pagos, nota):
    """
    items_df: columnas Prenda, Cantidad, Valor unitario (ya limpias)
    pagos: lista de tuplas (medio, detalle)
    """
    items = items_df.copy()
    items["Total"] = items["Cantidad"] * items["Valor unitario"]
    total = int(items["Total"].sum())
    saldo = total - int(abono)

    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        topMargin=0.6 * inch,
        bottomMargin=0.6 * inch,
        leftMargin=0.8 * inch,
        rightMargin=0.8 * inch,
    )
    elementos = []

    # --- Encabezado: fecha a la izquierda, logo en el centro ---
    try:
        ancho, alto = ImageReader(LOGO_PATH).getSize()
        logo = Image(LOGO_PATH, width=1.3 * inch, height=1.3 * inch * alto / ancho)
    except Exception:
        logo = Paragraph("", ESTILO_NORMAL)

    encabezado = Table(
        [[Paragraph(f"Fecha: {fecha_str}", ESTILO_NORMAL), logo, ""]],
        colWidths=[2.0 * inch, 2.9 * inch, 2.0 * inch],
    )
    encabezado.setStyle(
        TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("ALIGN", (1, 0), (1, 0), "CENTER")])
    )
    elementos.append(encabezado)
    elementos.append(Spacer(1, 0.3 * inch))

    # --- Destinatario y título ---
    elementos.append(Paragraph(f"Señores {escape(cliente)}", ESTILO_NEGRITA))
    elementos.append(Spacer(1, 0.08 * inch))
    elementos.append(Paragraph(f"Cuenta de cobro No. {int(consecutivo):03d}", ESTILO_NEGRITA))
    if concepto.strip():
        elementos.append(Spacer(1, 0.08 * inch))
        elementos.append(Paragraph(f"Concepto: {escape(concepto.strip())}", ESTILO_NORMAL))
    elementos.append(Spacer(1, 0.2 * inch))

    # --- Tabla ---
    datos = [["PRENDA", "CANTIDAD", "V. UNITARIO", "TOTAL"]]
    for _, fila in items.iterrows():
        datos.append(
            [
                Paragraph(escape(str(fila["Prenda"])), ESTILO_CENTRO),
                str(int(fila["Cantidad"])),
                formato_pesos(fila["Valor unitario"]),
                formato_pesos(fila["Total"]),
            ]
        )
    primera_fila_total = len(datos)
    datos.append(["TOTAL", "", "", formato_pesos(total)])
    if abono > 0:
        datos.append(["Abono recibido", "", "", formato_pesos(abono)])
        datos.append(["SALDO A PAGAR", "", "", formato_pesos(saldo)])

    tabla = Table(datos, colWidths=[2.3 * inch, 1.3 * inch, 1.6 * inch, 1.7 * inch])
    tabla.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.75, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.black),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, primera_fila_total), (-1, -1), "Helvetica-Bold"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    elementos.append(tabla)
    elementos.append(Spacer(1, 0.3 * inch))

    # --- Datos de pago (editables en el formulario) ---
    if pagos:
        elementos.append(Paragraph("Datos de pago", ESTILO_NEGRITA))
        for medio, detalle in pagos:
            elementos.append(Paragraph(f"<b>{escape(medio)}:</b> {escape(detalle)}", ESTILO_NORMAL))
        elementos.append(Spacer(1, 0.2 * inch))

    # --- Nota sobre facturación ---
    if nota.strip():
        elementos.append(Paragraph(escape(nota.strip()), ESTILO_NORMAL))
        elementos.append(Spacer(1, 0.3 * inch))

    # --- Quien cobra ---
    elementos.append(Paragraph(COBRADOR_NOMBRE, ESTILO_NEGRITA))
    elementos.append(Paragraph(f"NIT {COBRADOR_NIT}", ESTILO_NORMAL))

    doc.build(elementos)
    buffer.seek(0)

    nombre_archivo = f"Cuenta_de_cobro_{int(consecutivo):03d}_{limpiar_nombre(cliente)}.pdf"
    return buffer, nombre_archivo, total


# ================= VALORES DEL FORMULARIO =================

def items_por_defecto():
    return pd.DataFrame([{"Prenda": "", "Cantidad": 1, "Valor unitario": 0}])


def pagos_por_defecto():
    return pd.DataFrame([{"Medio": "Nequi", "Número / detalle": ""}])


def valores_iniciales():
    return {
        "campo_consecutivo": siguiente_consecutivo(),
        "campo_fecha": date.today(),
        "campo_cliente": "",
        "campo_concepto": "",
        "campo_items": items_por_defecto(),
        "campo_abono": 0,
        "campo_pagos": pagos_por_defecto(),
        "campo_nota": NOTA_POR_DEFECTO,
    }


def nueva_cuenta():
    """Deja el formulario limpio y con el siguiente consecutivo."""
    for clave, valor in valores_iniciales().items():
        st.session_state[clave] = valor
    st.session_state.pop("editor_items", None)
    st.session_state.pop("editor_pagos", None)


def cargar_cuenta_en_formulario(registro):
    """Usa una cuenta anterior como base de una cuenta NUEVA (nuevo consecutivo y fecha de hoy)."""
    try:
        items = pd.DataFrame(json.loads(registro["Prendas"]))[["Prenda", "Cantidad", "Valor unitario"]]
    except Exception:
        items = items_por_defecto()

    try:
        pagos = pd.DataFrame(
            [
                {"Medio": p.get("Medio", ""), "Número / detalle": p.get("Detalle", "")}
                for p in json.loads(registro["DatosPago"])
            ]
        )
        if pagos.empty:
            pagos = pagos_por_defecto()
    except Exception:
        pagos = pagos_por_defecto()

    st.session_state["campo_consecutivo"] = siguiente_consecutivo()
    st.session_state["campo_fecha"] = date.today()
    st.session_state["campo_cliente"] = str(registro.get("Cliente", ""))
    st.session_state["campo_concepto"] = str(registro.get("Concepto", ""))
    st.session_state["campo_items"] = items
    st.session_state["campo_abono"] = a_entero(registro.get("Abono"))
    st.session_state["campo_pagos"] = pagos
    st.session_state["campo_nota"] = NOTA_POR_DEFECTO
    st.session_state.pop("editor_items", None)
    st.session_state.pop("editor_pagos", None)


if "campo_cliente" not in st.session_state:
    for clave, valor in valores_iniciales().items():
        st.session_state[clave] = valor


# ================= INTERFAZ =================

st.image(LOGO_PATH, width=110)
st.title("Generador de cuentas de cobro")
st.caption("Confecciones Ness — llena los datos; se genera el PDF, se envía a los correos y se guarda en el historial.")

# ---- Historial ----
with st.expander("📋 Historial de cuentas de cobro (reutilizar una o cambiar su estado)"):
    try:
        historial = obtener_historial()
    except Exception as e:
        historial = []
        st.warning("No se pudo cargar el historial todavía.")
        st.caption(f"Detalle técnico: {e}")

    if not historial:
        st.caption("Todavía no hay cuentas de cobro guardadas.")
    else:
        opciones = {etiqueta_registro(r): r for r in historial}
        seleccion = st.selectbox("Elige una cuenta de cobro", list(opciones.keys()), key="sel_historial")
        registro = opciones[seleccion]

        total_r = a_entero(registro.get("ValorTotal"))
        abono_r = a_entero(registro.get("Abono"))
        st.caption(
            f"Total {formato_pesos(total_r)} · Abono {formato_pesos(abono_r)} · "
            f"Saldo {formato_pesos(total_r - abono_r)}"
        )

        st.button(
            "Usar como base para una cuenta nueva",
            on_click=cargar_cuenta_en_formulario,
            args=(registro,),
            width="stretch",
            key="btn_cargar",
        )

        estado_actual = registro.get("Estado", "Pendiente")
        nuevo_estado = st.selectbox(
            "Estado",
            ESTADOS,
            index=ESTADOS.index(estado_actual) if estado_actual in ESTADOS else 0,
            key=f"estado_{registro.get('ID')}",
        )
        if st.button("Guardar estado", width="stretch", key="btn_estado"):
            try:
                actualizar_estado(registro.get("ID"), nuevo_estado)
                obtener_historial.clear()
                st.success(f"Estado actualizado a «{nuevo_estado}».")
            except Exception as e:
                st.warning("No se pudo cambiar el estado.")
                st.caption(f"Detalle técnico: {e}")

# ---- Formulario ----
with st.form("form_cobro"):
    col1, col2 = st.columns(2)
    with col1:
        consecutivo = st.number_input("Número de cuenta de cobro", min_value=1, step=1, key="campo_consecutivo")
    with col2:
        fecha = st.date_input("Fecha", format="DD/MM/YYYY", key="campo_fecha")

    cliente = st.text_input(
        "Cliente (Señores…)",
        placeholder="Ej: IPS Eiteraa Jawapia",
        key="campo_cliente",
    )
    concepto = st.text_input(
        "Concepto (opcional)",
        placeholder="Ej: Confección de forros de mesa",
        key="campo_concepto",
    )

    st.markdown("**Prendas**")
    st.caption("Agrega, edita o borra filas. El total se calcula solo.")
    items_df = st.data_editor(
        st.session_state["campo_items"],
        num_rows="dynamic",
        width="stretch",
        column_config={
            "Cantidad": st.column_config.NumberColumn("Cantidad", min_value=1, step=1, format="%d"),
            "Valor unitario": st.column_config.NumberColumn("Valor unitario", min_value=0, step=1000, format="%d"),
        },
        key="editor_items",
    )

    abono = st.number_input(
        "Abono recibido (deja en 0 si no hubo)", min_value=0, step=10000, key="campo_abono"
    )

    st.markdown("**Datos de pago**")
    st.caption("Elige el medio y escribe el número. Puedes agregar varias filas.")
    pagos_df = st.data_editor(
        st.session_state["campo_pagos"],
        num_rows="dynamic",
        width="stretch",
        column_config={
            "Medio": st.column_config.SelectboxColumn("Medio", options=MEDIOS_PAGO, required=True),
            "Número / detalle": st.column_config.TextColumn("Número / detalle"),
        },
        key="editor_pagos",
    )

    nota = st.text_area("Nota al pie", height=80, key="campo_nota")

    generar = st.form_submit_button("Generar y enviar", type="primary", width="stretch")

if generar:
    # Limpiar prendas
    items = items_df.copy()
    items["Prenda"] = items["Prenda"].apply(texto)
    items = items[items["Prenda"] != ""]
    items["Cantidad"] = pd.to_numeric(items["Cantidad"], errors="coerce").fillna(0).astype(int)
    items["Valor unitario"] = pd.to_numeric(items["Valor unitario"], errors="coerce").fillna(0).astype(int)

    # Limpiar datos de pago
    pagos = []
    for _, fila in pagos_df.iterrows():
        medio = texto(fila.get("Medio"))
        detalle = texto(fila.get("Número / detalle"))
        if medio and detalle:
            pagos.append((medio, detalle))

    total_previo = int((items["Cantidad"] * items["Valor unitario"]).sum()) if not items.empty else 0

    if not cliente.strip():
        st.error("Falta el nombre del cliente.")
    elif items.empty:
        st.error("Agrega al menos una prenda.")
    elif (items["Cantidad"] <= 0).any() or (items["Valor unitario"] <= 0).any():
        st.error("Revisa que todas las prendas tengan cantidad y valor unitario mayores a 0.")
    elif abono > total_previo:
        st.error("El abono no puede ser mayor que el total.")
    else:
        fecha_str = fecha.strftime("%d/%m/%Y")
        buffer, nombre_archivo, total = generar_cuenta_cobro(
            consecutivo, fecha_str, cliente, concepto, items, abono, pagos, nota
        )
        pdf_bytes = buffer.getvalue()
        st.success(f"Cuenta de cobro generada. Total: {formato_pesos(total)}")

        st.download_button(
            "Descargar PDF",
            data=pdf_bytes,
            file_name=nombre_archivo,
            mime="application/pdf",
            width="stretch",
        )

        # ---- Correo ----
        try:
            destinatarios = [st.secrets["EMAIL_ADDRESS"]]
            correo_2 = st.secrets.get("CORREO_DESTINO_2", "")
            if correo_2:
                destinatarios.append(correo_2)

            enviar_correo(
                destinatarios=destinatarios,
                asunto=f"Cuenta de cobro No. {int(consecutivo):03d} — {cliente}",
                cuerpo=(
                    f"Buen día,\n\nAdjunto la cuenta de cobro No. {int(consecutivo):03d} "
                    f"para {cliente} por un valor total de {formato_pesos(total)} pesos.\n\n"
                    f"Confecciones Ness\n{COBRADOR_NOMBRE}\n{EMPRESA_TEL}"
                ),
                pdf_bytes=pdf_bytes,
                nombre_archivo=nombre_archivo,
            )
            st.success(f"Correo enviado a {', '.join(destinatarios)} ✅")
        except Exception as e:
            st.warning(
                "El documento se generó bien, pero no se pudo enviar el correo automáticamente. "
                "Puedes descargarlo arriba y adjuntarlo manualmente."
            )
            st.caption(f"Detalle técnico: {e}")

        # ---- Historial ----
        try:
            guardar_en_historial(consecutivo, fecha_str, cliente, concepto, items, abono, total, pagos)
            obtener_historial.clear()
            st.success("Guardada en el historial 📋")
        except Exception as e:
            st.warning("No se pudo guardar en el historial (el PDF y el correo no dependen de esto).")
            st.caption(f"Detalle técnico: {e}")

st.divider()
st.button("🆕 Empezar una cuenta nueva", on_click=nueva_cuenta, width="stretch")
