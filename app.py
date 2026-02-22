import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import base64
import json
from pdf2image import convert_from_bytes
from streamlit_image_coordinates import streamlit_image_coordinates
import io
from datetime import datetime
import urllib.parse
import requests

# ============================================================================
# CONFIGURACIÓN Y CONSTANTES
# ============================================================================

DRIVE_FOLDER_NAME = "Auditoría RETILAP"
DRIVE_FILE_NAME   = "proyectos_retilap.json"
GOOGLE_AUTH_URL   = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN_URL  = "https://oauth2.googleapis.com/token"
GOOGLE_USERINFO   = "https://www.googleapis.com/oauth2/v3/userinfo"
DRIVE_API         = "https://www.googleapis.com/drive/v3"
DRIVE_UPLOAD      = "https://www.googleapis.com/upload/drive/v3"

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

# ============================================================================
# AUTENTICACIÓN GOOGLE OAuth 2.0
# ============================================================================

def get_google_secrets():
    """Lee las credenciales de Streamlit secrets."""
    try:
        return (
            st.secrets["google"]["client_id"],
            st.secrets["google"]["client_secret"],
            st.secrets["google"]["redirect_uri"],
        )
    except Exception:
        st.error("❌ Faltan las credenciales de Google en Streamlit Secrets. "
                 "Ve a Settings → Secrets y agrega [google] client_id, client_secret y redirect_uri.")
        st.stop()


def get_auth_url():
    """Genera la URL de autorización de Google."""
    client_id, _, redirect_uri = get_google_secrets()
    params = {
        "client_id": client_id,
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": "openid email profile https://www.googleapis.com/auth/drive.file",
        "access_type": "offline",
        "prompt": "consent",
    }
    return GOOGLE_AUTH_URL + "?" + urllib.parse.urlencode(params)


def exchange_code_for_token(code):
    """Intercambia el código de autorización por un access_token."""
    client_id, client_secret, redirect_uri = get_google_secrets()
    resp = requests.post(GOOGLE_TOKEN_URL, data={
        "code": code,
        "client_id": client_id,
        "client_secret": client_secret,
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
    })
    if resp.status_code == 200:
        return resp.json()
    return None


def get_user_info(access_token):
    """Obtiene el perfil del usuario de Google."""
    resp = requests.get(GOOGLE_USERINFO, headers={"Authorization": f"Bearer {access_token}"})
    if resp.status_code == 200:
        return resp.json()
    return None


def refresh_access_token():
    """Refresca el access_token usando el refresh_token guardado en session."""
    client_id, client_secret, _ = get_google_secrets()
    refresh_token = st.session_state.get("refresh_token")
    if not refresh_token:
        return False
    resp = requests.post(GOOGLE_TOKEN_URL, data={
        "refresh_token": refresh_token,
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "refresh_token",
    })
    if resp.status_code == 200:
        data = resp.json()
        st.session_state.access_token = data["access_token"]
        return True
    return False


def auth_headers():
    return {"Authorization": f"Bearer {st.session_state.access_token}"}


def pagina_login():
    """Pantalla de inicio de sesión."""
    st.set_page_config(page_title="Auditoría Iluminación RETILAP", layout="centered")
    st.markdown("""
        <div style='text-align:center; padding: 60px 0 20px 0;'>
            <h1>💡 Auditoría de Iluminación</h1>
            <p style='font-size:1.2rem; color:#888;'>Norma RETILAP 2024</p>
        </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("Inicia sesión con tu cuenta de Google para acceder a la app. "
                "Tus proyectos se guardarán automáticamente en tu Google Drive.")
        auth_url = get_auth_url()
        st.markdown(f"""
            <a href="{auth_url}" target="_self">
                <button style="
                    background:#4285F4; color:white; border:none; padding:12px 28px;
                    font-size:1rem; border-radius:6px; cursor:pointer; width:100%;
                    display:flex; align-items:center; justify-content:center; gap:10px;
                ">
                    🔐 Iniciar sesión con Google
                </button>
            </a>
        """, unsafe_allow_html=True)


def manejar_callback_oauth():
    """
    Detecta si hay un código OAuth en la URL y lo intercambia por tokens.
    Retorna True si el usuario quedó autenticado.
    """
    params = st.query_params
    code = params.get("code")
    if not code:
        return False

    # Limpiar el código de la URL para no procesarlo dos veces
    st.query_params.clear()

    with st.spinner("🔐 Autenticando con Google..."):
        token_data = exchange_code_for_token(code)
        if not token_data or "access_token" not in token_data:
            st.error("❌ Error al autenticar con Google. Intenta de nuevo.")
            return False

        st.session_state.access_token  = token_data["access_token"]
        st.session_state.refresh_token = token_data.get("refresh_token", "")

        user_info = get_user_info(token_data["access_token"])
        if not user_info:
            st.error("❌ No se pudo obtener la información del usuario.")
            return False

        st.session_state.user_email = user_info.get("email", "")
        st.session_state.user_name  = user_info.get("name", "Usuario")
        st.session_state.user_pic   = user_info.get("picture", "")
        st.session_state.autenticado = True

    return True


# ============================================================================
# GOOGLE DRIVE — FUNCIONES
# ============================================================================

def drive_get_or_create_folder():
    """Busca o crea la carpeta 'Auditoría RETILAP' en Drive del usuario."""
    # Buscar carpeta existente
    q = f"name='{DRIVE_FOLDER_NAME}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    resp = requests.get(f"{DRIVE_API}/files", headers=auth_headers(),
                        params={"q": q, "fields": "files(id,name)"})
    if resp.status_code == 401:
        refresh_access_token()
        resp = requests.get(f"{DRIVE_API}/files", headers=auth_headers(),
                            params={"q": q, "fields": "files(id,name)"})

    files = resp.json().get("files", [])
    if files:
        return files[0]["id"]

    # Crear carpeta
    resp = requests.post(f"{DRIVE_API}/files", headers={**auth_headers(), "Content-Type": "application/json"},
                         json={"name": DRIVE_FOLDER_NAME,
                               "mimeType": "application/vnd.google-apps.folder"})
    return resp.json().get("id")


def drive_get_file_id(folder_id):
    """Busca el archivo proyectos_retilap.json dentro de la carpeta."""
    q = f"name='{DRIVE_FILE_NAME}' and '{folder_id}' in parents and trashed=false"
    resp = requests.get(f"{DRIVE_API}/files", headers=auth_headers(),
                        params={"q": q, "fields": "files(id,name)"})
    files = resp.json().get("files", [])
    return files[0]["id"] if files else None


def drive_cargar_proyectos():
    """Descarga y deserializa proyectos desde Drive."""
    try:
        folder_id = drive_get_or_create_folder()
        st.session_state.drive_folder_id = folder_id
        file_id = drive_get_file_id(folder_id)
        st.session_state.drive_file_id = file_id

        if not file_id:
            return {}  # Primera vez, no existe archivo aún

        resp = requests.get(f"{DRIVE_API}/files/{file_id}",
                            headers=auth_headers(),
                            params={"alt": "media"})
        if resp.status_code != 200:
            return {}

        proyectos_data = resp.json()
        proyectos = {}

        for p_name, p_data in proyectos_data.items():
            proyectos[p_name] = {"general": p_data["general"], "planos": {}}
            for plano_name, plano_info in p_data["planos"].items():
                plano_dict = {
                    "puntos": plano_info["puntos"],
                    "data": plano_info["data"],
                    "fotos": {}
                }
                if "img_base64" in plano_info:
                    try:
                        img_bytes = base64.b64decode(plano_info["img_base64"])
                        img = Image.open(io.BytesIO(img_bytes))
                        if img.mode != "RGB":
                            img = img.convert("RGB")
                        plano_dict["img"] = img
                    except Exception:
                        plano_dict["img"] = None
                else:
                    plano_dict["img"] = None
                proyectos[p_name]["planos"][plano_name] = plano_dict

        return proyectos
    except Exception as e:
        st.error(f"❌ Error al cargar desde Drive: {e}")
        return {}


def drive_guardar_proyectos(proyectos):
    """Serializa y sube proyectos a Drive (crea o actualiza el archivo)."""
    try:
        serializable = {}
        for p_name, p_data in proyectos.items():
            serializable[p_name] = {"general": p_data["general"].copy(), "planos": {}}
            for plano_name, plano_info in p_data["planos"].items():
                plano_dict = {
                    "puntos": plano_info["puntos"].copy() if isinstance(plano_info["puntos"], list) else [],
                    "data": [row.copy() for row in plano_info["data"]] if isinstance(plano_info["data"], list) else [],
                    "fotos": {}
                }
                if plano_info.get("img") is not None:
                    try:
                        buf = io.BytesIO()
                        plano_info["img"].save(buf, format="PNG")
                        buf.seek(0)
                        plano_dict["img_base64"] = base64.b64encode(buf.read()).decode()
                    except Exception as e:
                        st.warning(f"⚠️ No se pudo guardar imagen de '{plano_name}': {e}")
                serializable[p_name]["planos"][plano_name] = plano_dict

        content = json.dumps(serializable, ensure_ascii=False, indent=2).encode("utf-8")
        folder_id = st.session_state.get("drive_folder_id")
        file_id   = st.session_state.get("drive_file_id")

        if file_id:
            # Actualizar archivo existente
            resp = requests.patch(
                f"{DRIVE_UPLOAD}/files/{file_id}",
                headers={**auth_headers(), "Content-Type": "application/json"},
                params={"uploadType": "media"},
                data=content,
            )
        else:
            # Crear nuevo archivo con metadata
            boundary = "retilap_boundary"
            metadata = json.dumps({"name": DRIVE_FILE_NAME, "parents": [folder_id]}).encode()
            body = (
                f"--{boundary}\r\nContent-Type: application/json\r\n\r\n".encode()
                + metadata
                + f"\r\n--{boundary}\r\nContent-Type: application/json\r\n\r\n".encode()
                + content
                + f"\r\n--{boundary}--".encode()
            )
            resp = requests.post(
                f"{DRIVE_UPLOAD}/files",
                headers={**auth_headers(),
                         "Content-Type": f"multipart/related; boundary={boundary}"},
                params={"uploadType": "multipart"},
                data=body,
            )
            if resp.status_code in (200, 201):
                st.session_state.drive_file_id = resp.json().get("id")

        if resp.status_code not in (200, 201, 204):
            st.error(f"❌ Error al guardar en Drive: {resp.text}")

    except Exception as e:
        st.error(f"❌ Error al guardar en Drive: {e}")


# ============================================================================
# FUNCIONES DE REPORTE
# ============================================================================

def generar_reporte_csv(proyecto_data, proyecto_nombre):
    reportes = []
    for plano_nombre, plano_info in proyecto_data["planos"].items():
        if not plano_info["data"]:
            continue
        for row in plano_info["data"]:
            reportes.append({
                "Proyecto": proyecto_nombre,
                "Plano": plano_nombre,
                "Punto": row["Número"],
                "Coordenadas": row["Coordenadas"],
                "Tipo de Área": row.get("TipoArea", "N/A"),
                "Em requerida (lx)": row.get("Em_req", "N/A"),
                "Uo mínima": row.get("Uo_min", "N/A"),
                "Med1": row["Med1"],
                "Med2": row["Med2"],
                "Med3": row["Med3"],
                "Med4": row["Med4"],
                "Promedio": row["Promedio"],
                "Resultado": row["Resultado"],
                "Nota": row["Nota"],
            })
    if reportes:
        return pd.DataFrame(reportes).to_csv(index=False).encode("utf-8")
    return None


def generar_reporte_pdf(proyecto_data, proyecto_nombre):
    try:
        from reportlab.lib.pagesizes import letter
        from reportlab.lib import colors
        from reportlab.lib.units import inch, cm
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER
        from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                        Table, TableStyle, PageBreak,
                                        Image as RLImage, HRFlowable)

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter,
                                rightMargin=2*cm, leftMargin=2*cm,
                                topMargin=2*cm, bottomMargin=2*cm)
        styles = getSampleStyleSheet()
        story  = []

        estilo_titulo = ParagraphStyle("Titulo", parent=styles["Title"],
                                       fontSize=20, textColor=colors.HexColor("#1a3a5c"),
                                       spaceAfter=6, alignment=TA_CENTER)
        estilo_subtitulo = ParagraphStyle("Sub", parent=styles["Normal"],
                                          fontSize=12, textColor=colors.HexColor("#2c6fad"),
                                          spaceAfter=4, alignment=TA_CENTER)
        estilo_seccion = ParagraphStyle("Sec", parent=styles["Heading2"],
                                        fontSize=13, textColor=colors.HexColor("#1a3a5c"),
                                        spaceBefore=16, spaceAfter=8)
        estilo_normal = ParagraphStyle("Norm", parent=styles["Normal"],
                                       fontSize=10, spaceAfter=4)
        estilo_pie = ParagraphStyle("Pie", parent=styles["Normal"],
                                    fontSize=8, textColor=colors.grey, alignment=TA_CENTER)

        general = proyecto_data.get("general", {})

        # Portada
        story.append(Spacer(1, 1.5*inch))
        story.append(Paragraph("💡 REPORTE DE AUDITORÍA DE ILUMINACIÓN", estilo_titulo))
        story.append(Paragraph("Norma RETILAP 2024", estilo_subtitulo))
        story.append(Spacer(1, 0.3*inch))
        story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor("#2c6fad")))
        story.append(Spacer(1, 0.3*inch))

        info_data = [
            ["Proyecto:", proyecto_nombre],
            ["Orden:", general.get("numero_orden", "N/A")],
            ["Empresa:", general.get("nombre_empresa", "N/A")],
            ["Sede:", general.get("sede", "N/A")],
            ["Fecha:", general.get("fecha", "N/A")],
            ["Generado por:", st.session_state.get("user_email", "N/A")],
            ["Fecha reporte:", datetime.now().strftime("%d/%m/%Y %H:%M")],
        ]
        tabla_info = Table(info_data, colWidths=[3.5*cm, 12*cm])
        tabla_info.setStyle(TableStyle([
            ("FONTNAME", (0,0), (0,-1), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,-1), 11),
            ("TEXTCOLOR", (0,0), (0,-1), colors.HexColor("#1a3a5c")),
            ("ROWBACKGROUNDS", (0,0), (-1,-1), [colors.HexColor("#f0f4f8"), colors.white]),
            ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#ccddee")),
            ("TOPPADDING", (0,0), (-1,-1), 6),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
            ("LEFTPADDING", (0,0), (-1,-1), 10),
        ]))
        story.append(tabla_info)
        story.append(Spacer(1, 0.4*inch))

        # Resumen
        total, conformes, no_conformes = 0, 0, 0
        for pi in proyecto_data["planos"].values():
            for row in pi.get("data", []):
                total += 1
                if "✅" in str(row.get("Resultado","")):
                    conformes += 1
                else:
                    no_conformes += 1

        if total > 0:
            pct = round(conformes / total * 100, 1)
            c_pct = colors.HexColor("#27ae60") if pct >= 80 else colors.HexColor("#e74c3c")
            res_data = [["Puntos medidos","Conformes","No conformes","% Conformidad"],
                        [str(total), str(conformes), str(no_conformes), f"{pct}%"]]
            t_res = Table(res_data, colWidths=[4*cm]*4)
            t_res.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#1a3a5c")),
                ("TEXTCOLOR",(0,0),(-1,0), colors.white),
                ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),
                ("FONTSIZE",(0,0),(-1,-1),11),
                ("ALIGN",(0,0),(-1,-1),"CENTER"),
                ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                ("ROWHEIGHTS",(0,0),(-1,-1),28),
                ("GRID",(0,0),(-1,-1),0.5, colors.HexColor("#ccddee")),
                ("TEXTCOLOR",(3,1),(3,1), c_pct),
                ("BACKGROUND",(0,1),(-1,1), colors.HexColor("#f0f4f8")),
            ]))
            story.append(t_res)

        story.append(PageBreak())

        # Sección por plano
        for plano_nombre, plano_info in proyecto_data["planos"].items():
            data_rows = plano_info.get("data", [])
            plano_img = plano_info.get("img")

            story.append(Paragraph(f"📐 Plano: {plano_nombre}", estilo_seccion))
            story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor("#ccddee")))
            story.append(Spacer(1, 0.15*inch))

            if plano_img and data_rows:
                try:
                    draw_img = plano_img.copy()
                    if draw_img.width > 1400:
                        r = 1400 / draw_img.width
                        draw_img = draw_img.resize((1400, int(draw_img.height*r)), Image.LANCZOS)
                    draw = ImageDraw.Draw(draw_img)
                    font = ImageFont.load_default()
                    for row in data_rows:
                        try:
                            coords = row["Coordenadas"]
                            x, y = map(int, coords.strip("()").split(", "))
                            clr = row.get("Color","gray")
                            draw.ellipse((x-18,y-18,x+18,y+18), fill=clr, outline="black", width=3)
                            txt = str(row["Número"])
                            bb = font.getbbox(txt)
                            tc = "white" if clr == "red" else "black"
                            draw.text((x-(bb[2]-bb[0])//2, y-(bb[3]-bb[1])//2), txt, fill=tc, font=font)
                        except Exception:
                            pass
                    ib = io.BytesIO()
                    draw_img.save(ib, format="PNG")
                    ib.seek(0)
                    pw = letter[0] - 4*cm
                    ph = min(pw * draw_img.height / draw_img.width, 5*inch)
                    story.append(RLImage(ib, width=pw, height=ph))
                    story.append(Spacer(1, 0.2*inch))
                except Exception as e:
                    story.append(Paragraph(f"(Error al renderizar mapa: {e})", estilo_normal))

            if data_rows:
                story.append(Paragraph("Tabla de mediciones:", estilo_normal))
                enc = ["#","Tipo de Área","Em req.","Med1","Med2","Med3","Med4","Promedio","Resultado","Nota"]
                filas = [enc]
                for row in data_rows:
                    filas.append([
                        str(row.get("Número","")),
                        str(row.get("TipoArea","N/A"))[:35],
                        f"{row.get('Em_req','N/A')} lx",
                        str(row.get("Med1","")), str(row.get("Med2","")),
                        str(row.get("Med3","")), str(row.get("Med4","")),
                        str(row.get("Promedio","")),
                        "Conforme" if "✅" in str(row.get("Resultado","")) else "No conforme",
                        str(row.get("Nota",""))[:40],
                    ])
                cw = [0.8*cm,4.5*cm,1.8*cm,1.4*cm,1.4*cm,1.4*cm,1.4*cm,2*cm,2.2*cm,3*cm]
                t = Table(filas, colWidths=cw, repeatRows=1)
                es = [
                    ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#1a3a5c")),
                    ("TEXTCOLOR",(0,0),(-1,0), colors.white),
                    ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),
                    ("FONTSIZE",(0,0),(-1,-1),8),
                    ("ALIGN",(0,0),(-1,-1),"CENTER"),
                    ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
                    ("GRID",(0,0),(-1,-1),0.4, colors.HexColor("#aaccee")),
                    ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#f0f4f8")]),
                    ("TOPPADDING",(0,0),(-1,-1),4),
                    ("BOTTOMPADDING",(0,0),(-1,-1),4),
                ]
                for idx2, row in enumerate(data_rows, start=1):
                    clr2 = colors.HexColor("#27ae60") if "✅" in str(row.get("Resultado","")) else colors.HexColor("#e74c3c")
                    es.append(("TEXTCOLOR",(8,idx2),(8,idx2), clr2))
                    es.append(("FONTNAME",(8,idx2),(8,idx2),"Helvetica-Bold"))
                t.setStyle(TableStyle(es))
                story.append(t)
            else:
                story.append(Paragraph("Sin mediciones en este plano.", estilo_normal))

            story.append(Spacer(1, 0.3*inch))
            story.append(PageBreak())

        story.append(HRFlowable(width="100%", thickness=1, color=colors.grey))
        story.append(Spacer(1, 0.1*inch))
        story.append(Paragraph(
            f"Reporte generado por {st.session_state.get('user_email','N/A')} · "
            f"Auditoría Iluminación RETILAP 2024 · {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            estilo_pie))

        doc.build(story)
        buffer.seek(0)
        return buffer.getvalue()

    except ImportError:
        st.error("❌ Instala reportlab en requirements.txt")
        return None
    except Exception as e:
        st.error(f"❌ Error al generar PDF: {e}")
        return None


# ============================================================================
# INICIALIZACIÓN
# ============================================================================

def inicializar_session_state():
    if "pagina" not in st.session_state:
        st.session_state.pagina = "inicio"
    if "proyecto_actual" not in st.session_state:
        st.session_state.proyecto_actual = None
    if "autenticado" not in st.session_state:
        st.session_state.autenticado = False

    # Cargar proyectos desde Drive solo la primera vez tras login
    if st.session_state.autenticado and "proyectos" not in st.session_state:
        with st.spinner("☁️ Cargando proyectos desde Google Drive..."):
            st.session_state.proyectos = drive_cargar_proyectos()
            if not isinstance(st.session_state.proyectos, dict):
                st.session_state.proyectos = {}


def sidebar_usuario():
    """Muestra info del usuario y botón de cerrar sesión en el sidebar."""
    with st.sidebar:
        st.markdown("---")
        pic = st.session_state.get("user_pic", "")
        name = st.session_state.get("user_name", "Usuario")
        email = st.session_state.get("user_email", "")
        if pic:
            st.image(pic, width=50)
        st.markdown(f"**{name}**")
        st.caption(email)
        st.markdown("☁️ Datos en Google Drive")
        st.markdown("---")
        if st.button("🚪 Cerrar sesión"):
            for key in ["autenticado","access_token","refresh_token",
                        "user_email","user_name","user_pic","proyectos",
                        "drive_folder_id","drive_file_id","pagina","proyecto_actual"]:
                st.session_state.pop(key, None)
            st.rerun()


# ============================================================================
# PÁGINAS DE LA APLICACIÓN
# ============================================================================

def pagina_inicio():
    st.set_page_config(page_title="Auditoría Iluminación RETILAP", layout="centered")
    st.markdown("""
        <div style='text-align:center; padding: 60px 0 20px 0;'>
            <h1>💡 Auditoría de Iluminación</h1>
            <p style='font-size:1.2rem; color:#888;'>Norma RETILAP 2024</p>
        </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("Inicia sesión con tu cuenta de Google para acceder a la app. "
                "Tus proyectos se guardarán automáticamente en tu Google Drive.")
        auth_url = get_auth_url()
        st.markdown(f"""
            <a href="{auth_url}" target="_self">
                <button style="
                    background:#4285F4; color:white; border:none; padding:12px 28px;
                    font-size:1rem; border-radius:6px; cursor:pointer; width:100%;
                    display:flex; align-items:center; justify-content:center; gap:10px;
                ">
                    🔐 Iniciar sesión con Google
                </button>
            </a>
        """, unsafe_allow_html=True)


def pagina_nuevo_proyecto():
    st.title("➕ Nuevo Proyecto")
    if st.button("← Volver", key="btn_volver_nuevo"):
        st.session_state.pagina = "inicio"
        st.rerun()

    st.subheader("📝 Datos del Proyecto")
    col1, col2 = st.columns(2)
    with col1:
        numero_orden   = st.text_input("Número de Orden", key="input_numero_orden")
        nombre_empresa = st.text_input("Nombre de la Empresa", key="input_nombre_empresa")
    with col2:
        sede  = st.text_input("Sede", key="input_sede")
        fecha = st.date_input("Fecha", key="input_fecha")

    if st.button("✅ Crear Proyecto", key="btn_crear_proyecto"):
        if numero_orden and nombre_empresa and sede:
            proyecto_nombre = f"{nombre_empresa} - {sede} ({fecha.strftime('%Y-%m-%d')})"
            if proyecto_nombre not in st.session_state.proyectos:
                st.session_state.proyectos[proyecto_nombre] = {
                    "general": {
                        "numero_orden": numero_orden,
                        "nombre_empresa": nombre_empresa,
                        "sede": sede,
                        "fecha": fecha.strftime("%d/%m/%Y"),
                        "tipo_area": list(RETILAP_REFERENCIA.keys())[0],
                    },
                    "planos": {}
                }
                drive_guardar_proyectos(st.session_state.proyectos)
                st.session_state.proyecto_actual = proyecto_nombre
                st.session_state.pagina = "editar_proyecto"
                st.success("✅ Proyecto creado y guardado en Drive")
                st.rerun()
            else:
                st.error("❌ Este proyecto ya existe")
        else:
            st.error("❌ Por favor completa todos los campos")


def pagina_editar_proyecto():
    proyecto_actual = st.session_state.proyecto_actual
    proyecto_data   = st.session_state.proyectos[proyecto_actual]
    general         = proyecto_data["general"]

    for campo, default in [("sede",""),("fecha",""),("tipo_area", list(RETILAP_REFERENCIA.keys())[0])]:
        if campo not in general:
            general[campo] = default

    st.title(f"✏️ Editar Proyecto: {proyecto_actual}")
    if st.button("← Volver", key="btn_volver_editar"):
        st.session_state.pagina = "inicio"
        st.rerun()

    st.subheader("📋 Información del Proyecto")
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Orden:** {general.get('numero_orden','N/A')}")
        st.write(f"**Empresa:** {general.get('nombre_empresa','N/A')}")
    with col2:
        st.write(f"**Sede:** {general.get('sede','N/A')}")
        st.write(f"**Fecha:** {general.get('fecha','N/A')}")

    st.divider()
    st.subheader("📐 Planos")
    st.write("**Agregar nuevo plano:**")
    col1, col2 = st.columns([2, 1])
    with col1:
        plano_nombre = st.text_input("Nombre del plano", key="input_plano_nombre")
    with col2:
        uploaded_plano = st.file_uploader("Archivo (JPG o PDF)", type=["jpg","jpeg","pdf"], key="upload_plano")

    if plano_nombre and uploaded_plano:
        if plano_nombre not in proyecto_data["planos"]:
            if st.button("✅ Agregar Plano", key="btn_agregar_plano"):
                try:
                    if uploaded_plano.type == "application/pdf":
                        img = convert_from_bytes(uploaded_plano.read())[0]
                    else:
                        img = Image.open(uploaded_plano)
                    if img.mode != "RGB":
                        img = img.convert("RGB")
                    if img.width > 1920:
                        r = 1920 / img.width
                        img = img.resize((1920, int(img.height*r)), Image.LANCZOS)
                    proyecto_data["planos"][plano_nombre] = {"img":img,"puntos":[],"data":[],"fotos":{}}
                    drive_guardar_proyectos(st.session_state.proyectos)
                    st.success(f"✅ Plano '{plano_nombre}' agregado y guardado en Drive")
                except Exception as e:
                    st.error(f"❌ Error al cargar plano: {e}")
        else:
            st.warning(f"⚠️ El plano '{plano_nombre}' ya existe")

    st.divider()
    if proyecto_data["planos"]:
        st.write("**Planos disponibles:**")
        for pn in proyecto_data["planos"]:
            c1, c2 = st.columns([3,1])
            with c1:
                st.write(f"📄 {pn}")
            with c2:
                if st.button("📍 Editar", key=f"btn_editar_plano_{pn}"):
                    st.session_state.plano_actual = pn
                    st.session_state.pagina = "editar_plano"
                    st.rerun()
    else:
        st.info("ℹ️ Agrega un plano para comenzar a marcar puntos")


def pagina_editar_plano():
    if "plano_actual" not in st.session_state:
        st.session_state.pagina = "inicio"
        st.rerun()

    proyecto_actual = st.session_state.proyecto_actual
    plano_actual    = st.session_state.plano_actual
    proyecto_data   = st.session_state.proyectos[proyecto_actual]
    plano_data      = proyecto_data["planos"][plano_actual]
    general         = proyecto_data["general"]
    plano_img       = plano_data.get("img")

    st.title(f"📍 Editar Plano: {plano_actual}")
    if st.button("← Volver", key="btn_volver_plano"):
        st.session_state.pagina = "editar_proyecto"
        st.rerun()

    if plano_img is None:
        st.error("⚠️ La imagen del plano no se pudo cargar. Vuelve a subirla.")
        return

    st.image(plano_img, caption=f"Plano: {plano_actual} — Haz clic para marcar puntos",
             use_container_width=True)

    clicked = streamlit_image_coordinates(
        plano_img,
        key=f"clicker_{proyecto_actual}_{plano_actual}",
        height=plano_img.height,
        width=plano_img.width,
    )
    if clicked is not None:
        x, y = clicked["x"], clicked["y"]
        if not any(abs(px-x) < 12 and abs(py-y) < 12 for px, py in plano_data["puntos"]):
            plano_data["puntos"].append((x, y))
            drive_guardar_proyectos(st.session_state.proyectos)
            st.rerun()

    st.write(f"**Puntos en este plano:** {len(plano_data['puntos'])}")

    col1, col2 = st.columns([1,1])
    with col1:
        if st.button("🗑️ Eliminar último punto", key=f"eliminar_ultimo_{proyecto_actual}_{plano_actual}"):
            if plano_data["puntos"]:
                ultimo_num = len(plano_data["puntos"])
                plano_data["puntos"].pop()
                plano_data["data"] = [d for d in plano_data["data"] if d["Número"] != ultimo_num]
                drive_guardar_proyectos(st.session_state.proyectos)
                st.rerun()
    with col2:
        if st.button("🧹 Limpiar todos los puntos", key=f"limpiar_puntos_{proyecto_actual}_{plano_actual}"):
            plano_data["puntos"] = []
            plano_data["data"]   = []
            drive_guardar_proyectos(st.session_state.proyectos)
            st.rerun()

    st.divider()

    if plano_data["puntos"]:
        st.subheader("📊 Mediciones por punto")
        for i, (x, y) in enumerate(plano_data["puntos"]):
            existing = next((d for d in plano_data["data"] if d["Número"] == i+1), {})

            with st.expander(f"Punto {i+1} ({int(x)}, {int(y)})", expanded=False):

                # Eliminar punto individual
                if st.button(f"🗑️ Eliminar punto {i+1}",
                             key=f"del_punto_{proyecto_actual}_{plano_actual}_{i}"):
                    plano_data["puntos"].pop(i)
                    plano_data["data"] = [d for d in plano_data["data"] if d["Número"] != i+1]
                    for d in plano_data["data"]:
                        if d["Número"] > i+1:
                            d["Número"] -= 1
                    drive_guardar_proyectos(st.session_state.proyectos)
                    st.rerun()

                # Tipo de área por punto
                tipos = list(RETILAP_REFERENCIA.keys())
                ta_guardado = existing.get("TipoArea", general.get("tipo_area", tipos[0]))
                ta_idx = tipos.index(ta_guardado) if ta_guardado in tipos else 0
                tipo_area_punto = st.selectbox("🏷️ Tipo de área según RETILAP", tipos,
                                               index=ta_idx,
                                               key=f"tipo_area_{proyecto_actual}_{plano_actual}_{i}")
                vals = RETILAP_REFERENCIA[tipo_area_punto]
                em_sugerido = vals["Em"]
                uo_min      = vals["Uo"]
                st.info(f"Em requerida: **{em_sugerido} lx** · Uo mínima: **{uo_min}**")

                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    med1 = st.number_input("Med 1 (lx)", min_value=0.0, step=0.1,
                                           value=float(existing.get("Med1",0.0)),
                                           key=f"m1_{proyecto_actual}_{plano_actual}_{i}")
                with c2:
                    med2 = st.number_input("Med 2 (lx)", min_value=0.0, step=0.1,
                                           value=float(existing.get("Med2",0.0)),
                                           key=f"m2_{proyecto_actual}_{plano_actual}_{i}")
                with c3:
                    med3 = st.number_input("Med 3 (lx)", min_value=0.0, step=0.1,
                                           value=float(existing.get("Med3",0.0)),
                                           key=f"m3_{proyecto_actual}_{plano_actual}_{i}")
                with c4:
                    med4 = st.number_input("Med 4 (lx)", min_value=0.0, step=0.1,
                                           value=float(existing.get("Med4",0.0)),
                                           key=f"m4_{proyecto_actual}_{plano_actual}_{i}")

                foto_subida = st.file_uploader(f"Foto punto {i+1} (opcional)",
                                               type=["jpg","jpeg","png"],
                                               key=f"foto_{proyecto_actual}_{plano_actual}_{i}")
                if foto_subida:
                    plano_data["fotos"][i+1] = foto_subida.read()

                nota = st.text_area("Notas / observaciones", height=80,
                                    value=existing.get("Nota",""),
                                    key=f"nota_{proyecto_actual}_{plano_actual}_{i}")

                if all(v > 0 for v in [med1, med2, med3, med4]):
                    promedio  = (med1+med2+med3+med4) / 4
                    conforme  = promedio >= em_sugerido
                    color     = "green" if conforme else "red"
                    resultado = "✅ Conforme" if conforme else "❌ No conforme"

                    if conforme:
                        st.success(f"Promedio: **{round(promedio,1)} lx** → {resultado}")
                    else:
                        st.error(f"Promedio: **{round(promedio,1)} lx** → {resultado} (requiere ≥ {em_sugerido} lx)")

                    entrada = {
                        "Número": i+1,
                        "Coordenadas": f"({int(x)}, {int(y)})",
                        "TipoArea": tipo_area_punto,
                        "Em_req": em_sugerido,
                        "Uo_min": uo_min,
                        "Med1": med1, "Med2": med2, "Med3": med3, "Med4": med4,
                        "Promedio": round(promedio, 1),
                        "Resultado": resultado,
                        "Color": color,
                        "Nota": nota.strip(),
                        "Foto": foto_subida is not None,
                    }
                    idx_e = next((j for j,d in enumerate(plano_data["data"]) if d["Número"]==i+1), None)
                    if idx_e is not None:
                        plano_data["data"][idx_e] = entrada
                    else:
                        plano_data["data"].append(entrada)
                    drive_guardar_proyectos(st.session_state.proyectos)

        st.divider()

        if plano_data["data"]:
            st.subheader("🗺️ Mapa de puntos")
            df_plano  = pd.DataFrame(plano_data["data"])
            draw_img  = plano_img.copy()
            draw      = ImageDraw.Draw(draw_img)
            font      = ImageFont.load_default()
            for _, r in df_plano.iterrows():
                x, y = map(int, r["Coordenadas"].strip("()").split(", "))
                clr  = r["Color"]
                draw.ellipse((x-18,y-18,x+18,y+18), fill=clr, outline="black", width=3)
                txt  = str(r["Número"])
                bb   = font.getbbox(txt)
                tc   = "white" if clr == "red" else "black"
                draw.text((x-(bb[2]-bb[0])//2, y-(bb[3]-bb[1])//2), txt, fill=tc, font=font)
            st.image(draw_img, caption=f"Mapa — {plano_actual}", use_container_width=True)

            st.subheader("📊 Tabla de Resultados")
            cols = ["Número","Coordenadas","TipoArea","Em_req","Med1","Med2","Med3","Med4","Promedio","Resultado"]
            cols_ok = [c for c in cols if c in df_plano.columns]
            st.dataframe(df_plano[cols_ok].rename(columns={
                "TipoArea":"Tipo de Área","Em_req":"Em req. (lx)"}),
                use_container_width=True)


# ============================================================================
# FUNCIÓN PRINCIPAL
# ============================================================================

def main():
    st.set_page_config(page_title="Auditoría Iluminación RETILAP", layout="wide",
                       page_icon="💡")

    # 1. Manejar callback OAuth (cuando Google redirige de vuelta)
    if not st.session_state.get("autenticado"):
        if manejar_callback_oauth():
            st.rerun()

    # 2. Si no está autenticado → mostrar login
    if not st.session_state.get("autenticado"):
        pagina_login()
        return

    # 3. Inicializar estado (incluye carga de Drive)
    inicializar_session_state()

    # 4. Sidebar con info de usuario
    sidebar_usuario()

    # 5. Navegar a la página correspondiente
    pagina = st.session_state.pagina
    if pagina == "inicio":
        pagina_inicio()
    elif pagina == "nuevo_proyecto":
        pagina_nuevo_proyecto()
    elif pagina == "editar_proyecto":
        pagina_editar_proyecto()
    elif pagina == "editar_plano":
        pagina_editar_plano()


if __name__ == "__main__":
    main()

