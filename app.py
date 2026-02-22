import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from fpdf import FPDF
import os
import base64
import json
import io
from datetime import datetime
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from pdf2image import convert_from_bytes
from streamlit_image_coordinates import streamlit_image_coordinates

# ============================================================================
# CONFIGURACIÓN GOOGLE DRIVE (Service Account)
# ============================================================================

SCOPES = ['https://www.googleapis.com/auth/drive']

@st.cache_resource
def get_drive_service():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["google_service_account"],
        scopes=SCOPES
    )
    return build('drive', 'v3', credentials=creds)

drive_service = get_drive_service()
DRIVE_FOLDER_ID = st.secrets["google_service_account"]["drive_folder_id"]
DRIVE_FILE_NAME = "proyectos_retilap.json"

# Descargar proyectos desde Drive al inicio
if "proyectos" not in st.session_state:
    file_list = drive_service.files().list(
        q=f"name='{DRIVE_FILE_NAME}' and '{DRIVE_FOLDER_ID}' in parents and trashed=false",
        fields="files(id, name)"
    ).execute()
    files = file_list.get('files', [])
    if files:
        file_id = files[0]['id']
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        fh.seek(0)
        st.session_state.proyectos = json.load(fh)
    else:
        st.session_state.proyectos = {}

if "proyecto_actual" not in st.session_state:
    st.session_state.proyecto_actual = None if not st.session_state.proyectos else list(st.session_state.proyectos.keys())[0]

if "logo_bytes" not in st.session_state:
    st.session_state.logo_bytes = None

# Guardar en Drive
def guardar_en_drive():
    content = json.dumps(st.session_state.proyectos, ensure_ascii=False, indent=4).encode("utf-8")
    file_list = drive_service.files().list(
        q=f"name='{DRIVE_FILE_NAME}' and '{DRIVE_FOLDER_ID}' in parents and trashed=false",
        fields="files(id)"
    ).execute()
    files = file_list.get('files', [])
    if files:
        file_id = files[0]['id']
        media = MediaIoBaseUpload(io.BytesIO(content), mimetype='application/json', resumable=True)
        drive_service.files().update(fileId=file_id, media_body=media).execute()
    else:
        file_metadata = {
            'name': DRIVE_FILE_NAME,
            'parents': [DRIVE_FOLDER_ID]
        }
        media = MediaIoBaseUpload(io.BytesIO(content), mimetype='application/json', resumable=True)
        drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

# Tabla RETILAP completa
RETILAP_REFERENCIA = {
    "Oficinas - Escritura y lectura detallada": {"Em": 500, "Uo": 0.60},
    "Oficinas - Trabajo administrativo general": {"Em": 300, "Uo": 0.40},
    "Oficinas - Recepción y áreas de espera": {"Em": 200, "Uo": 0.40},
    "Educación - Aulas y laboratorios": {"Em": 500, "Uo": 0.60},
    "Educación - Bibliotecas y salas de lectura": {"Em": 500, "Uo": 0.60},
    "Educación - Pasillos y escaleras": {"Em": 100, "Uo": 0.40},
    "Comercio - Ventas y exhibición general": {"Em": 300, "Uo": 0.40},
    "Comercio - Cajas y áreas de pago": {"Em": 500, "Uo": 0.60},
    "Salud - Consultorios y habitaciones pacientes": {"Em": 300, "Uo": 0.60},
    "Salud - Quirófanos y salas de cirugía": {"Em": 1000, "Uo": 0.70},
    "Salud - Pasillos hospitales": {"Em": 100, "Uo": 0.40},
    "Industria - Tareas de precisión alta": {"Em": 1500, "Uo": 0.70},
    "Industria - Montaje e inspección fina": {"Em": 750, "Uo": 0.60},
    "Industria - Trabajo ordinario": {"Em": 300, "Uo": 0.40},
    "Almacenes y depósitos": {"Em": 200, "Uo": 0.40},
    "Restaurantes y cafeterías - Áreas de mesas": {"Em": 200, "Uo": 0.40},
    "Restaurantes - Cocinas": {"Em": 500, "Uo": 0.60},
    "Hoteles - Habitaciones": {"Em": 200, "Uo": 0.40},
    "Hoteles - Pasillos y escaleras": {"Em": 100, "Uo": 0.40},
    "Vías M1 - Alta velocidad / tránsito pesado": {"Em": 50, "Uo": 0.40},
    "Vías M2 - Velocidad media / tránsito mixto": {"Em": 30, "Uo": 0.35},
    "Vías M3 - Velocidad moderada / zonas urbanas": {"Em": 20, "Uo": 0.35},
    "Vías M4 - Zonas residenciales / colectoras": {"Em": 10, "Uo": 0.30},
    "Vías M5 - Peatonal / ciclorrutas principales": {"Em": 7.5, "Uo": 0.25},
    "Vías M6 - Peatonal secundaria / andenes": {"Em": 5, "Uo": 0.20},
    "Estacionamientos exteriores": {"Em": 30, "Uo": 0.30},
    "Parques y plazas públicas": {"Em": 2, "Uo": 0.20},
    "Áreas deportivas - Fútbol / canchas grandes": {"Em": 100, "Uo": 0.40},
    "Áreas deportivas - Tenis / básquet": {"Em": 200, "Uo": 0.50},
    "Estaciones de servicio": {"Em": 50, "Uo": 0.40},
    "Túneles - Zona de acceso (noche)": {"Em": 30, "Uo": 0.40},
    "Túneles - Zona interior": {"Em": 5, "Uo": 0.40},
}

# Subir logo (sidebar)
st.sidebar.header("Configuración")
logo_file = st.sidebar.file_uploader("Logo HISIG CONSULTORIA (PNG/JPG)", type=["png", "jpg", "jpeg"], key="logo_upload")
if logo_file is not None:
    st.session_state.logo_bytes = logo_file.read()
    st.sidebar.success("Logo cargado")
    st.sidebar.image(logo_file, width=150)

# Gestión de proyectos
st.subheader("Gestión de Proyectos")
col1, col2 = st.columns([3, 1])
with col1:
    nombre_proyecto = st.text_input("Nombre del proyecto (ej. Auditoría Centro XYZ - 2026)", key="nombre_proyecto_input")
with col2:
    if st.button("Crear / Seleccionar Proyecto", key="crear_proyecto_btn"):
        if nombre_proyecto.strip():
            if nombre_proyecto not in st.session_state.proyectos:
                st.session_state.proyectos[nombre_proyecto] = {
                    "general": {
                        "numero_orden": "",
                        "nombre_empresa": "",
                        "caracteristicas_sistema": "",
                        "estructura_entorno": "",
                        "tipo_area": list(RETILAP_REFERENCIA.keys())[0],
                    },
                    "planos": {}
                }
                guardar_en_drive()
            st.session_state.proyecto_actual = nombre_proyecto
            st.rerun()
        else:
            st.warning("Ingresa un nombre para el proyecto")

# Seleccionar proyecto
if st.session_state.proyectos:
    proyecto_actual = st.selectbox("Proyecto actual", list(st.session_state.proyectos.keys()), key="proyecto_select")
    st.session_state.proyecto_actual = proyecto_actual

    proyecto_data = st.session_state.proyectos[proyecto_actual]
    general = proyecto_data["general"]

    st.subheader(f"Datos generales - {proyecto_actual}")
    general["numero_orden"] = st.text_input("Número de Orden", value=general.get("numero_orden", ""), key=f"orden_{proyecto_actual}")
    general["nombre_empresa"] = st.text_input("Nombre de la empresa / cliente", value=general.get("nombre_empresa", ""), key=f"empresa_{proyecto_actual}")
    general["caracteristicas_sistema"] = st.text_area("Características del sistema de iluminación", value=general.get("caracteristicas_sistema", ""), height=80, key=f"sistema_{proyecto_actual}")
    general["estructura_entorno"] = st.text_area("Estructura del entorno", value=general.get("estructura_entorno", ""), height=80, key=f"entorno_{proyecto_actual}")

    # Guardar cambios generales
    if any([
        general["numero_orden"] != st.session_state.get(f"orden_{proyecto_actual}", ""),
        general["nombre_empresa"] != st.session_state.get(f"empresa_{proyecto_actual}", ""),
        general["caracteristicas_sistema"] != st.session_state.get(f"sistema_{proyecto_actual}", ""),
        general["estructura_entorno"] != st.session_state.get(f"entorno_{proyecto_actual}", ""),
    ]):
        guardar_en_drive()

    # Planos del proyecto
    st.subheader("Planos del proyecto")
    plano_nombre = st.text_input("Nombre del plano", key=f"plano_nombre_{proyecto_actual}")
    uploaded_plano = st.file_uploader(f"Subir plano '{plano_nombre}' (JPG o PDF)", type=["jpg", "jpeg", "pdf"], key=f"upload_plano_{proyecto_actual}")

    if plano_nombre and uploaded_plano and plano_nombre not in proyecto_data["planos"]:
        try:
            if uploaded_plano.type == "application/pdf":
                images = convert_from_bytes(uploaded_plano.read())
                img = images[0]
            else:
                img = Image.open(uploaded_plano)
            
            proyecto_data["planos"][plano_nombre] = {
                "img": img,
                "puntos": [],
                "data": [],
                "fotos": {}
            }
            guardar_en_drive()
            st.success(f"Plano '{plano_nombre}' agregado")
            st.rerun()
        except Exception as e:
            st.error(f"Error al cargar plano: {e}")

    # Seleccionar plano actual
    if proyecto_data["planos"]:
        plano_actual = st.selectbox("Plano actual", list(proyecto_data["planos"].keys()), key=f"plano_select_{proyecto_actual}")
        plano_data = proyecto_data["planos"][plano_actual]
        plano_img = plano_data.get("img")

        if plano_img is None:
            st.warning(f"El plano '{plano_actual}' no tiene imagen guardada. Sube el plano de nuevo.")
            uploaded_plano = st.file_uploader(f"Subir plano '{plano_actual}' nuevamente", type=["jpg", "jpeg", "pdf"], key=f"reupload_{proyecto_actual}_{plano_actual}")
            if uploaded_plano is not None:
                try:
                    if uploaded_plano.type == "application/pdf":
                        images = convert_from_bytes(uploaded_plano.read())
                        img = images[0]
                    else:
                        img = Image.open(uploaded_plano)
                    plano_data["img"] = img
                    guardar_en_drive()
                    st.success("Plano recargado")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al recargar plano: {e}")
        else:
            st.image(plano_img, caption=f"Editando: {plano_actual} - Haz clic para marcar puntos", use_column_width=True)

        tipo_area = st.selectbox("Tipo de área según RETILAP", list(RETILAP_REFERENCIA.keys()), index=list(RETILAP_REFERENCIA.keys()).index(general["tipo_area"]), key=f"tipo_area_{proyecto_actual}_{plano_actual}")
        general["tipo_area"] = tipo_area

        valores = RETILAP_REFERENCIA[tipo_area]
        em_sugerido = valores["Em"]
        uo_min = valores["Uo"]

        st.success(f"Iluminancia mantenida sugerida (Em): **{em_sugerido} lx** | Uniformidad mínima sugerida (Uo): **{uo_min}**")

        # Selección de puntos
        clicked = streamlit_image_coordinates(
            plano_img,
            key=f"clicker_{proyecto_actual}_{plano_actual}",
            height=plano_img.height,
            width=plano_img.width,
        )

        if clicked is not None:
            x, y = clicked["x"], clicked["y"]
            if not any(abs(px - x) < 12 and abs(py - y) < 12 for px, py in plano_data["puntos"]):
                plano_data["puntos"].append((x, y))
                guardar_en_drive()
                st.rerun()

        st.write(f"**Puntos en este plano:** {len(plano_data['puntos'])}")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Eliminar último punto", key=f"eliminar_ultimo_{proyecto_actual}_{plano_actual}"):
                if plano_data["puntos"]:
                    plano_data["puntos"].pop()
                    guardar_en_drive()
                    st.rerun()
        with col2:
            if st.button("Limpiar puntos de este plano", key=f"limpiar_puntos_{proyecto_actual}_{plano_actual}"):
                plano_data["puntos"] = []
                guardar_en_drive()
                st.rerun()

        # Ingreso de mediciones, foto y notas
        plano_data["data"] = []
        for i, (x, y) in enumerate(plano_data["puntos"]):
            st.subheader(f"Punto {i+1} ({int(x)}, {int(y)})")

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                med1 = st.number_input("Med 1", min_value=0.0, step=0.1, key=f"m1_{i}")
            with col2:
                med2 = st.number_input("Med 2", min_value=0.0, step=0.1, key=f"m2_{i}")
            with col3:
                med3 = st.number_input("Med 3", min_value=0.0, step=0.1, key=f"m3_{i}")
            with col4:
                med4 = st.number_input("Med 4", min_value=0.0, step=0.1, key=f"m4_{i}")

            foto_subida = st.file_uploader(f"Foto del punto {i+1} (opcional)", type=["jpg", "jpeg", "png"], key=f"foto_{i}")
            if foto_subida is not None:
                plano_data["fotos"][i+1] = foto_subida.read()
                guardar_en_drive()

            nota = st.text_area("Notas / observaciones", height=100, key=f"nota_{i}")
            if nota.strip() != plano_data.get("data", [{}])[i].get("Nota", "") if plano_data["data"] else "":
                guardar_en_drive()

            if all(v > 0 for v in [med1, med2, med3, med4]):
                promedio = (med1 + med2 + med3 + med4) / 4
                conforme = promedio >= em_sugerido
                color = "green" if conforme else "red"
                resultado = "Conforme" if conforme else "No conforme"

                plano_data["data"].append({
                    "Número": i+1,
                    "Coordenadas": f"({int(x)}, {int(y)})",
                    "Med1": med1,
                    "Med2": med2,
                    "Med3": med3,
                    "Med4": med4,
                    "Promedio": round(promedio, 1),
                    "Resultado": resultado,
                    "Color": color,
                    "Nota": nota.strip(),
                    "Foto": foto_subida is not None
                })
                guardar_en_drive()

        # Vista previa mapa
        if plano_data["data"] and "img" in plano_data:
            df_plano = pd.DataFrame(plano_data["data"])
            draw_img = plano_data["img"].copy()
            draw = ImageDraw.Draw(draw_img)
            font = ImageFont.load_default()

            for _, r in df_plano.iterrows():
                x, y = map(int, r["Coordenadas"].strip("()").split(", "))
                color = r["Color"]
                draw.ellipse((x - 18, y - 18, x + 18, y + 18), fill=color, outline="black", width=3)

                texto = str(r["Número"])
                bbox = font.getbbox(texto)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                text_x = x - text_width // 2
                text_y = y - text_height // 2

                text_color = "white" if color == "red" else "black"
                draw.text((text_x, text_y), texto, fill=text_color, font=font)

            st.image(draw_img, caption=f"Mapa - {plano_actual}")

        # Botones para generar PDF
        st.markdown("---")
        st.subheader("Generar Reportes PDF")
        col1, col2 = st.columns(2)
        with col1:
            if st.button(f"Descargar PDF SOLO de '{proyecto_actual}'", key=f"pdf_individual_{proyecto_actual}"):
                pdf = generar_pdf_proyecto(proyecto_data, proyecto_actual, st.session_state.logo_bytes)
                pdf_bytes = pdf.output(dest="S")
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="Reporte_{proyecto_actual.replace(" ", "_")}.pdf">Descargar PDF de este proyecto</a>'
                st.markdown(href, unsafe_allow_html=True)
        with col2:
            if st.button("Generar PDF con TODOS los proyectos", key="pdf_todos_global"):
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", "B", 16)
                pdf.cell(0, 10, "Reporte Auditoría Iluminación RETILAP 2024 - Todos los proyectos", ln=1, align="C")
                pdf.set_font("Arial", size=12)
                pdf.cell(0, 8, f"Fecha y hora: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=1)

                for proyecto_nombre, proyecto_data in st.session_state.proyectos.items():
                    pdf = generar_pdf_proyecto(proyecto_data, proyecto_nombre, st.session_state.logo_bytes)

                pdf_bytes = pdf.output(dest="S")
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="Reporte_Todos_Proyectos_RETILAP.pdf">Descargar PDF Completo</a>'
                st.markdown(href, unsafe_allow_html=True)

# Función para generar PDF de un proyecto individual
def generar_pdf_proyecto(proyecto_data, proyecto_nombre, logo_bytes=None):
    pdf = FPDF()
    pdf.add_page()

    # Logo en portada
    if logo_bytes is not None:
        logo_img = Image.open(io.BytesIO(logo_bytes))
        temp_logo = "temp_logo.png"
        logo_img.save(temp_logo, format="PNG")
        pdf.image(temp_logo, x=75, y=10, w=60)
        os.remove(temp_logo)

    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 40, f"Reporte Auditoría Iluminación RETILAP 2024 - {proyecto_nombre}", ln=1, align="C")
    pdf.set_font("Arial", size=11)
    pdf.cell(0, 8, f"Número de Orden: {proyecto_data['general']['numero_orden'] or 'No especificado'}", ln=1)
    pdf.cell(0, 8, f"Empresa: {proyecto_data['general']['nombre_empresa'] or 'No especificada'}", ln=1)
    pdf.cell(0, 8, f"Fecha y hora: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=1)
    pdf.ln(5)

    pdf.multi_cell(190, 5, f"Características del sistema:\n{proyecto_data['general']['caracteristicas_sistema'] or 'No especificadas'}")
    pdf.ln(5)
    pdf.multi_cell(190, 5, f"Estructura del entorno:\n{proyecto_data['general']['estructura_entorno'] or 'No especificada'}")
    pdf.ln(10)

    for plano_nombre, plano_info in proyecto_data["planos"].items():
        if not plano_info["data"]:
            continue

        pdf.add_page()
        pdf.set_font("Arial", "B", 12)
        pdf.cell(0, 10, f"Plano: {plano_nombre}", ln=1)

        df_plano = pd.DataFrame(plano_info["data"])

        # Tabla con mediciones separadas
        pdf.set_font("Arial", "B", 9)
        pdf.cell(15, 8, "Punto", border=1)
        pdf.cell(30, 8, "Coords", border=1)
        pdf.cell(18, 8, "Med1", border=1)
        pdf.cell(18, 8, "Med2", border=1)
        pdf.cell(18, 8, "Med3", border=1)
        pdf.cell(18, 8, "Med4", border=1)
        pdf.cell(25, 8, "Promedio", border=1)
        pdf.cell(40, 8, "Resultado", border=1)
        pdf.ln()

        pdf.set_font("Arial", size=9)
        for _, r in df_plano.iterrows():
            pdf.cell(15, 8, str(r["Número"]), border=1)
            pdf.cell(30, 8, r["Coordenadas"], border=1)
            pdf.cell(18, 8, f"{r['Med1']}", border=1)
            pdf.cell(18, 8, f"{r['Med2']}", border=1)
            pdf.cell(18, 8, f"{r['Med3']}", border=1)
            pdf.cell(18, 8, f"{r['Med4']}", border=1)
            pdf.cell(25, 8, f"{r['Promedio']:.1f}", border=1)

            if r["Resultado"] == "Conforme":
                pdf.set_text_color(0, 128, 0)
            else:
                pdf.set_text_color(255, 0, 0)

            pdf.cell(40, 8, r["Resultado"], border=1)
            pdf.set_text_color(0, 0, 0)
            pdf.ln()

        # Mapa (solo si hay imagen)
        if "img" in plano_info:
            pdf.add_page()
            scale = min(190 / plano_info["img"].width, 270 / plano_info["img"].height)
            pdf_w = plano_info["img"].width * scale
            pdf_h = plano_info["img"].height * scale

            temp_plano = f"temp_{plano_nombre}.png"
            plano_info["img"].save(temp_plano, format="PNG")
            pdf.image(temp_plano, x=10, y=10, w=pdf_w, h=pdf_h)
            os.remove(temp_plano)

            for _, r in df_plano.iterrows():
                x = int(r["Coordenadas"].split(",")[0].strip("() ")) * scale + 10
                y = int(r["Coordenadas"].split(",")[1].strip("() ")) * scale + 10

                if r["Color"] == "green":
                    pdf.set_fill_color(0, 255, 0)
                else:
                    pdf.set_fill_color(255, 0, 0)
                pdf.circle(x, y, 5, style="FD")

                texto = str(r["Número"])
                pdf.set_xy(x - 2, y - 3)
                pdf.set_text_color(0, 0, 0)
                pdf.set_font("Arial", size=8)
                pdf.cell(4, 4, texto, align="C")

        # Fotos y notas
        if plano_info["fotos"] or df_plano["Nota"].str.strip().any():
            pdf.add_page()
            pdf.set_font("Arial", "B", 12)
            pdf.cell(0, 10, f"Fotos y notas - {plano_nombre}", ln=1, align="C")
            pdf.ln(5)

            y_pos = 25
            for i, row in df_plano.iterrows():
                nota = row["Nota"].strip()
                foto_bytes = plano_info["fotos"].get(row["Número"])

                if foto_bytes or nota:
                    pdf.set_font("Arial", "B", 10)
                    pdf.cell(0, 8, f"Punto {row['Número']}", ln=1)
                    pdf.set_font("Arial", size=10)

                    if foto_bytes:
                        foto_img = Image.open(io.BytesIO(foto_bytes))
                        temp_foto = f"temp_foto_{row['Número']}.png"
                        foto_img.save(temp_foto, format="PNG")
                        pdf.image(temp_foto, x=10, y=y_pos, w=80, h=60)
                        os.remove(temp_foto)

                        pdf.set_xy(95, y_pos)
                        pdf.multi_cell(100, 5, f"Nota: {nota}" if nota else "Sin nota adicional", align="L")

                        y_pos += 70
                    else:
                        pdf.multi_cell(0, 5, f"Nota: {nota}")
                        y_pos += 15

                    y_pos += 10
                    if y_pos > 250:
                        pdf.add_page()
                        y_pos = 25

    return pdf

# Botón global para PDF de todos los proyectos
if st.session_state.proyectos:
    st.markdown("---")
    st.subheader("Generar Reportes PDF")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Descargar PDF SOLO de proyecto actual", key="pdf_individual_global"):
            pdf = generar_pdf_proyecto(st.session_state.proyectos[st.session_state.proyecto_actual], st.session_state.proyecto_actual, st.session_state.logo_bytes)
            pdf_bytes = pdf.output(dest="S")
            b64 = base64.b64encode(pdf_bytes).decode()
            href = f'<a href="data:application/pdf;base64,{b64}" download="Reporte_{st.session_state.proyecto_actual.replace(" ", "_")}.pdf">Descargar PDF de este proyecto</a>'
            st.markdown(href, unsafe_allow_html=True)
    with col2:
        if st.button("Generar PDF con TODOS los proyectos", key="pdf_todos_global"):
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", "B", 16)
            pdf.cell(0, 10, "Reporte Auditoría Iluminación RETILAP 2024 - Todos los proyectos", ln=1, align="C")
            pdf.set_font("Arial", size=12)
            pdf.cell(0, 8, f"Fecha y hora: {datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=1)

            for proyecto_nombre, proyecto_data in st.session_state.proyectos.items():
                pdf = generar_pdf_proyecto(proyecto_data, proyecto_nombre, st.session_state.logo_bytes)

            pdf_bytes = pdf.output(dest="S")
            b64 = base64.b64encode(pdf_bytes).decode()
            href = f'<a href="data:application/pdf;base64,{b64}" download="Reporte_Todos_Proyectos_RETILAP.pdf">Descargar PDF Completo</a>'
            st.markdown(href, unsafe_allow_html=True)
