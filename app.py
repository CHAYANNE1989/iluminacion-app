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

# ============================================================================
# CONFIGURACIÓN Y CONSTANTES
# ============================================================================

PROYECTOS_FILE = "proyectos.json"

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
# FUNCIONES DE PERSISTENCIA
# ============================================================================

def cargar_proyectos():
    """Carga proyectos desde JSON y reconstruye imágenes desde base64"""
    if os.path.exists(PROYECTOS_FILE):
        try:
            with open(PROYECTOS_FILE, "r", encoding="utf-8") as f:
                proyectos_data = json.load(f)
                proyectos = {}
                
                for p_name, p_data in proyectos_data.items():
                    proyectos[p_name] = {
                        "general": p_data["general"],
                        "planos": {}
                    }
                    
                    for plano_name, plano_info in p_data["planos"].items():
                        plano_dict = {
                            "puntos": plano_info["puntos"],
                            "data": plano_info["data"],
                            "fotos": plano_info.get("fotos", {})
                        }
                        
                        if "img_base64" in plano_info:
                            try:
                                img_bytes = base64.b64decode(plano_info["img_base64"])
                                img = Image.open(io.BytesIO(img_bytes))
                                if img.mode != 'RGB':
                                    img = img.convert('RGB')
                                plano_dict["img"] = img
                            except Exception as e:
                                plano_dict["img"] = None
                        else:
                            plano_dict["img"] = None
                        
                        proyectos[p_name]["planos"][plano_name] = plano_dict
                
                return proyectos
        except Exception as e:
            st.error(f"Error al cargar proyectos: {e}")
            return {}
    return {}


def guardar_proyectos(proyectos):
    """Guarda proyectos a JSON, convirtiendo imágenes a base64"""
    try:
        serializable = {}
        
        for p_name, p_data in proyectos.items():
            serializable[p_name] = {
                "general": p_data["general"].copy(),
                "planos": {}
            }
            
            for plano_name, plano_info in p_data["planos"].items():
                plano_dict = {
                    "puntos": plano_info["puntos"].copy() if isinstance(plano_info["puntos"], list) else [],
                    "data": [row.copy() for row in plano_info["data"]] if isinstance(plano_info["data"], list) else [],
                    "fotos": {}
                }
                
                if plano_info.get("img") is not None:
                    try:
                        img_bytes = io.BytesIO()
                        plano_info["img"].save(img_bytes, format="PNG")
                        img_bytes.seek(0)
                        plano_dict["img_base64"] = base64.b64encode(img_bytes.read()).decode()
                    except Exception as e:
                        st.warning(f"⚠️ No se pudo guardar la imagen del plano '{plano_name}': {e}")
                
                serializable[p_name]["planos"][plano_name] = plano_dict
        
        with open(PROYECTOS_FILE, "w", encoding="utf-8") as f:
            json.dump(serializable, f, ensure_ascii=False, indent=4)
    except Exception as e:
        st.error(f"Error al guardar proyectos: {e}")


def generar_reporte_csv(proyecto_data, proyecto_nombre):
    """Genera un reporte en CSV"""
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
                "Nota": row["Nota"]
            })
    
    if reportes:
        df = pd.DataFrame(reportes)
        return df.to_csv(index=False).encode('utf-8')
    return None


def generar_reporte_pdf(proyecto_data, proyecto_nombre):
    """
    Genera un reporte profesional en PDF usando reportlab.
    Incluye portada, tabla de mediciones por plano y mapa de puntos anotado.
    """
    try:
        from reportlab.lib.pagesizes import letter, landscape
        from reportlab.lib import colors
        from reportlab.lib.units import inch, cm
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
        from reportlab.platypus import (
            SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
            PageBreak, Image as RLImage, HRFlowable
        )
        from reportlab.platypus.flowables import KeepTogether

        buffer = io.BytesIO()
        doc = SimpleDocTemplate(
            buffer,
            pagesize=letter,
            rightMargin=2*cm,
            leftMargin=2*cm,
            topMargin=2*cm,
            bottomMargin=2*cm
        )

        styles = getSampleStyleSheet()
        story = []

        # ---- Estilos personalizados ----
        estilo_titulo = ParagraphStyle(
            'Titulo',
            parent=styles['Title'],
            fontSize=20,
            textColor=colors.HexColor('#1a3a5c'),
            spaceAfter=6,
            alignment=TA_CENTER,
        )
        estilo_subtitulo = ParagraphStyle(
            'Subtitulo',
            parent=styles['Normal'],
            fontSize=12,
            textColor=colors.HexColor('#2c6fad'),
            spaceAfter=4,
            alignment=TA_CENTER,
        )
        estilo_seccion = ParagraphStyle(
            'Seccion',
            parent=styles['Heading2'],
            fontSize=13,
            textColor=colors.HexColor('#1a3a5c'),
            spaceBefore=16,
            spaceAfter=8,
            borderPad=4,
        )
        estilo_normal = ParagraphStyle(
            'Normal2',
            parent=styles['Normal'],
            fontSize=10,
            spaceAfter=4,
        )
        estilo_pie = ParagraphStyle(
            'Pie',
            parent=styles['Normal'],
            fontSize=8,
            textColor=colors.grey,
            alignment=TA_CENTER,
        )

        general = proyecto_data.get("general", {})

        # ================================================================
        # PORTADA
        # ================================================================
        story.append(Spacer(1, 1.5*inch))
        story.append(Paragraph("💡 REPORTE DE AUDITORÍA DE ILUMINACIÓN", estilo_titulo))
        story.append(Paragraph("Norma RETILAP 2024", estilo_subtitulo))
        story.append(Spacer(1, 0.3*inch))
        story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#2c6fad')))
        story.append(Spacer(1, 0.3*inch))

        # Tabla de datos del proyecto
        info_data = [
            ["Proyecto:", proyecto_nombre],
            ["Orden de trabajo:", general.get("numero_orden", "N/A")],
            ["Empresa:", general.get("nombre_empresa", "N/A")],
            ["Sede:", general.get("sede", "N/A")],
            ["Fecha:", general.get("fecha", "N/A")],
            ["Tipo de área:", general.get("tipo_area", "N/A")],
            ["Generado:", datetime.now().strftime("%d/%m/%Y %H:%M")],
        ]

        tabla_info = Table(info_data, colWidths=[3.5*cm, 12*cm])
        tabla_info.setStyle(TableStyle([
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('TEXTCOLOR', (0, 0), (0, -1), colors.HexColor('#1a3a5c')),
            ('ROWBACKGROUNDS', (0, 0), (-1, -1), [colors.HexColor('#f0f4f8'), colors.white]),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#ccddee')),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ]))
        story.append(tabla_info)
        story.append(Spacer(1, 0.5*inch))

        # Resumen general de conformidad
        total_puntos = 0
        conformes = 0
        no_conformes = 0
        for plano_info in proyecto_data["planos"].values():
            for row in plano_info.get("data", []):
                total_puntos += 1
                if "✅" in str(row.get("Resultado", "")):
                    conformes += 1
                else:
                    no_conformes += 1

        if total_puntos > 0:
            pct = round(conformes / total_puntos * 100, 1)
            color_pct = colors.HexColor('#27ae60') if pct >= 80 else colors.HexColor('#e74c3c')

            resumen_data = [
                ["Total puntos medidos", "Conformes", "No conformes", "% Conformidad"],
                [str(total_puntos), str(conformes), str(no_conformes), f"{pct}%"],
            ]
            tabla_resumen = Table(resumen_data, colWidths=[4*cm]*4)
            tabla_resumen.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a3a5c')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 11),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('ROWHEIGHTS', (0, 0), (-1, -1), 28),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#ccddee')),
                ('TEXTCOLOR', (3, 1), (3, 1), color_pct),
                ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#f0f4f8')),
            ]))
            story.append(tabla_resumen)

        story.append(PageBreak())

        # ================================================================
        # SECCIÓN POR PLANO
        # ================================================================
        for plano_nombre, plano_info in proyecto_data["planos"].items():
            data_rows = plano_info.get("data", [])
            plano_img = plano_info.get("img")

            story.append(Paragraph(f"📐 Plano: {plano_nombre}", estilo_seccion))
            story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#ccddee')))
            story.append(Spacer(1, 0.15*inch))

            # Mapa de puntos anotado
            if plano_img is not None and data_rows:
                try:
                    draw_img = plano_img.copy()

                    # Escalar si es muy grande
                    max_w = 1400
                    if draw_img.width > max_w:
                        ratio = max_w / draw_img.width
                        draw_img = draw_img.resize(
                            (max_w, int(draw_img.height * ratio)), Image.LANCZOS
                        )

                    draw = ImageDraw.Draw(draw_img)
                    try:
                        font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 28)
                    except Exception:
                        font = ImageFont.load_default(size=28)

                    for row in data_rows:
                        try:
                            coords = row["Coordenadas"]
                            x, y = map(int, str(coords).strip("()").split(", "))
                            clr = row.get("Color","gray")
                            draw.ellipse((x-24, y-24, x+24, y+24), fill=clr)
                            txt = str(row["Número"])
                            bb = font.getbbox(txt)
                            tw, th = bb[2]-bb[0], bb[3]-bb[1]
                            for dx, dy in [(-1,-1),(1,-1),(-1,1),(1,1)]:
                                draw.text((x-tw//2+dx, y-th//2+dy), txt, fill="black", font=font)
                            draw.text((x-tw//2, y-th//2), txt, fill="white", font=font)
                        except Exception:
                            pass

                    # Insertar imagen en el PDF
                    img_buffer = io.BytesIO()
                    draw_img.save(img_buffer, format="PNG")
                    img_buffer.seek(0)

                    page_w = letter[0] - 4*cm
                    aspect = draw_img.height / draw_img.width
                    img_h = min(page_w * aspect, 5*inch)

                    rl_img = RLImage(img_buffer, width=page_w, height=img_h)
                    story.append(rl_img)
                    story.append(Spacer(1, 0.2*inch))
                except Exception as e:
                    story.append(Paragraph(f"(No se pudo renderizar el mapa: {e})", estilo_normal))

            # Tabla de mediciones
            if data_rows:
                story.append(Paragraph("Tabla de mediciones:", estilo_normal))

                encabezado = ["#", "Tipo de Área", "Em req.", "Med1", "Med2", "Med3", "Med4", "Promedio", "Resultado", "Nota"]
                filas = [encabezado]

                for row in data_rows:
                    filas.append([
                        str(row.get("Número", "")),
                        str(row.get("TipoArea", "N/A"))[:35],
                        f"{row.get('Em_req', 'N/A')} lx",
                        str(row.get("Med1", "")),
                        str(row.get("Med2", "")),
                        str(row.get("Med3", "")),
                        str(row.get("Med4", "")),
                        str(row.get("Promedio", "")),
                        "Conforme" if "✅" in str(row.get("Resultado", "")) else "No conforme",
                        str(row.get("Nota", ""))[:40],
                    ])

                col_widths = [0.8*cm, 4.5*cm, 1.8*cm, 1.5*cm, 1.5*cm, 1.5*cm, 1.5*cm, 2*cm, 2.2*cm, 3*cm]
                tabla = Table(filas, colWidths=col_widths, repeatRows=1)

                estilos_tabla = [
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a3a5c')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#aaccee')),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f4f8')]),
                    ('TOPPADDING', (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ]

                # Colorear columna Resultado
                for i, row in enumerate(data_rows, start=1):
                    if "✅" in str(row.get("Resultado", "")):
                        estilos_tabla.append(('TEXTCOLOR', (7, i), (7, i), colors.HexColor('#27ae60')))
                        estilos_tabla.append(('FONTNAME', (7, i), (7, i), 'Helvetica-Bold'))
                    else:
                        estilos_tabla.append(('TEXTCOLOR', (7, i), (7, i), colors.HexColor('#e74c3c')))
                        estilos_tabla.append(('FONTNAME', (7, i), (7, i), 'Helvetica-Bold'))

                tabla.setStyle(TableStyle(estilos_tabla))
                story.append(tabla)
            else:
                story.append(Paragraph("Sin mediciones registradas en este plano.", estilo_normal))

            story.append(Spacer(1, 0.3*inch))
            story.append(PageBreak())

        # ================================================================
        # PIE DE PÁGINA / NOTA LEGAL
        # ================================================================
        story.append(Spacer(1, 0.5*inch))
        story.append(HRFlowable(width="100%", thickness=1, color=colors.grey))
        story.append(Spacer(1, 0.1*inch))
        story.append(Paragraph(
            f"Reporte generado automáticamente · Auditoría Iluminación RETILAP 2024 · "
            f"{datetime.now().strftime('%d/%m/%Y %H:%M')}",
            estilo_pie
        ))

        doc.build(story)
        buffer.seek(0)
        return buffer.getvalue()

    except ImportError:
        st.error("❌ Falta instalar reportlab. Agrega 'reportlab' a requirements.txt")
        return None
    except Exception as e:
        st.error(f"❌ Error al generar PDF: {e}")
        return None


# ============================================================================
# INICIALIZACIÓN DE SESSION STATE
# ============================================================================

def inicializar_session_state():
    """Inicializa las variables de sesión necesarias"""
    try:
        if "proyectos" not in st.session_state:
            st.session_state.proyectos = cargar_proyectos()
            if not isinstance(st.session_state.proyectos, dict):
                st.session_state.proyectos = {}

        if "pagina" not in st.session_state:
            st.session_state.pagina = "inicio"

        if "proyecto_actual" not in st.session_state:
            st.session_state.proyecto_actual = None
    except Exception as e:
        st.error(f"Error al inicializar: {str(e)}")
        st.session_state.proyectos = {}
        st.session_state.pagina = "inicio"
        st.session_state.proyecto_actual = None


# ============================================================================
# PÁGINAS DE LA APLICACIÓN
# ============================================================================

def pagina_inicio():
    """Página inicial con lista de proyectos"""
    st.title("💡 Auditoría Iluminación RETILAP 2024")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        st.subheader("📋 Proyectos Guardados")
    
    with col2:
        if st.button("➕ Nuevo Proyecto", key="btn_nuevo_proyecto"):
            st.session_state.pagina = "nuevo_proyecto"
            st.rerun()
    
    if st.session_state.proyectos:
        proyectos_list = list(st.session_state.proyectos.items())
        for idx, (proyecto_nombre, proyecto_data) in enumerate(proyectos_list):
            with st.container(border=True):
                col_info, col_buttons = st.columns([3, 1])
                
                with col_info:
                    st.write(f"**{proyecto_nombre}**")
                    st.caption(f"Orden: {proyecto_data['general'].get('numero_orden', 'N/A')}")
                    st.caption(f"Empresa: {proyecto_data['general'].get('nombre_empresa', 'N/A')}")
                    st.caption(f"Sede: {proyecto_data['general'].get('sede', 'N/A')}")
                    st.caption(f"Fecha: {proyecto_data['general'].get('fecha', 'N/A')}")
                
                with col_buttons:
                    # Botón Editar
                    if st.button("✏️ Editar", key=f"btn_editar_{idx}_{proyecto_nombre}"):
                        st.session_state.proyecto_actual = proyecto_nombre
                        st.session_state.pagina = "editar_proyecto"
                        st.rerun()

                    # Descarga CSV — siempre visible (fix bug download_button anidado)
                    csv_data = generar_reporte_csv(proyecto_data, proyecto_nombre)
                    if csv_data:
                        st.download_button(
                            label="📊 Descargar CSV",
                            data=csv_data,
                            file_name=f"Reporte_{proyecto_nombre.replace(' ', '_')}.csv",
                            mime="text/csv",
                            key=f"download_csv_{idx}_{proyecto_nombre}"
                        )

                    # Descarga PDF
                    pdf_data = generar_reporte_pdf(proyecto_data, proyecto_nombre)
                    if pdf_data:
                        st.download_button(
                            label="📄 Descargar PDF",
                            data=pdf_data,
                            file_name=f"Reporte_{proyecto_nombre.replace(' ', '_')}.pdf",
                            mime="application/pdf",
                            key=f"download_pdf_{idx}_{proyecto_nombre}"
                        )

                    # Eliminar
                    if st.button("🗑️ Eliminar", key=f"btn_eliminar_{idx}_{proyecto_nombre}"):
                        if proyecto_nombre in st.session_state.proyectos:
                            del st.session_state.proyectos[proyecto_nombre]
                            guardar_proyectos(st.session_state.proyectos)
                            st.success(f"✅ Proyecto '{proyecto_nombre}' eliminado")
                            st.rerun()
    else:
        st.info("ℹ️ No hay proyectos guardados. Crea uno nuevo para comenzar.")


def pagina_nuevo_proyecto():
    """Página para crear un nuevo proyecto"""
    st.title("➕ Nuevo Proyecto")
    
    if st.button("← Volver", key="btn_volver_nuevo"):
        st.session_state.pagina = "inicio"
        st.rerun()
    
    st.subheader("📝 Datos del Proyecto")
    
    col1, col2 = st.columns(2)
    
    with col1:
        numero_orden = st.text_input("Número de Orden", key="input_numero_orden")
        nombre_empresa = st.text_input("Nombre de la Empresa", key="input_nombre_empresa")
    
    with col2:
        sede = st.text_input("Sede", key="input_sede")
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
                        "fecha": fecha.strftime('%d/%m/%Y'),
                        "tipo_area": list(RETILAP_REFERENCIA.keys())[0],
                    },
                    "planos": {}
                }
                guardar_proyectos(st.session_state.proyectos)
                st.session_state.proyecto_actual = proyecto_nombre
                st.session_state.pagina = "editar_proyecto"
                st.success("✅ Proyecto creado exitosamente")
                st.rerun()
            else:
                st.error("❌ Este proyecto ya existe")
        else:
            st.error("❌ Por favor completa todos los campos")


def pagina_editar_proyecto():
    """Página para editar un proyecto y sus planos"""
    proyecto_actual = st.session_state.proyecto_actual
    proyecto_data = st.session_state.proyectos[proyecto_actual]
    general = proyecto_data["general"]
    
    if "sede" not in general:
        general["sede"] = ""
    if "fecha" not in general:
        general["fecha"] = ""
    if "tipo_area" not in general:
        general["tipo_area"] = list(RETILAP_REFERENCIA.keys())[0]
    
    st.title(f"✏️ Editando: {proyecto_actual}")
    
    if st.button("← Volver", key="btn_volver_editar"):
        st.session_state.pagina = "inicio"
        st.rerun()
    
    st.subheader("📋 Información del Proyecto")
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Orden:** {general.get('numero_orden', 'N/A')}")
        st.write(f"**Empresa:** {general.get('nombre_empresa', 'N/A')}")
    with col2:
        st.write(f"**Sede:** {general.get('sede', 'N/A')}")
        st.write(f"**Fecha:** {general.get('fecha', 'N/A')}")
    
    st.divider()
    
    st.subheader("📐 Planos")
    
    st.write("**Agregar nuevo plano:**")
    col1, col2 = st.columns([2, 1])
    
    with col1:
        plano_nombre = st.text_input("Nombre del plano", key="input_plano_nombre")
    
    with col2:
        uploaded_plano = st.file_uploader("Seleccionar archivo (JPG o PDF)", type=["jpg", "jpeg", "pdf"], key="upload_plano")
    
    if plano_nombre and uploaded_plano:
        if plano_nombre not in proyecto_data["planos"]:
            if st.button("✅ Agregar Plano", key="btn_agregar_plano"):
                try:
                    if uploaded_plano.type == "application/pdf":
                        images = convert_from_bytes(uploaded_plano.read())
                        img = images[0]
                    else:
                        img = Image.open(uploaded_plano)
                    
                    if img.mode != 'RGB':
                        img = img.convert('RGB')

                    # Redimensionar si es muy grande
                    MAX_WIDTH = 1920
                    if img.width > MAX_WIDTH:
                        ratio = MAX_WIDTH / img.width
                        img = img.resize((MAX_WIDTH, int(img.height * ratio)), Image.LANCZOS)
                    
                    proyecto_data["planos"][plano_nombre] = {
                        "img": img,
                        "puntos": [],
                        "data": [],
                        "fotos": {}
                    }
                    guardar_proyectos(st.session_state.proyectos)
                    st.success(f"✅ Plano '{plano_nombre}' agregado correctamente")
                    st.session_state.plano_agregado = True
                except Exception as e:
                    st.error(f"❌ Error al cargar plano: {str(e)}")
        else:
            st.warning(f"⚠️ El plano '{plano_nombre}' ya existe")
    
    st.divider()
    
    if proyecto_data["planos"]:
        st.write("**Planos disponibles:**")
        
        for plano_nombre in proyecto_data["planos"].keys():
            col1, col2 = st.columns([3, 1])
            
            with col1:
                st.write(f"📄 {plano_nombre}")
            
            with col2:
                if st.button("📍 Editar", key=f"btn_editar_plano_{plano_nombre}"):
                    st.session_state.plano_actual = plano_nombre
                    st.session_state.pagina = "editar_plano"
                    st.rerun()
    else:
        st.info("ℹ️ Agrega un plano para comenzar a marcar puntos")


def pagina_editar_plano():
    """Página para editar puntos en un plano"""
    # Validación defensiva
    if "plano_actual" not in st.session_state:
        st.session_state.pagina = "inicio"
        st.rerun()

    proyecto_actual = st.session_state.proyecto_actual
    plano_actual = st.session_state.plano_actual
    
    proyecto_data = st.session_state.proyectos[proyecto_actual]
    plano_data = proyecto_data["planos"][plano_actual]
    general = proyecto_data["general"]
    plano_img = plano_data.get("img")
    
    st.title(f"📍 Editar Plano: {plano_actual}")
    
    if st.button("← Volver", key="btn_volver_plano"):
        st.session_state.pagina = "editar_proyecto"
        st.rerun()
    
    if plano_img is None:
        st.error(f"⚠️ La imagen del plano no se pudo cargar. Por favor, vuelve a subir el plano.")
        return
    
    st.image(plano_img, caption=f"Plano: {plano_actual} - Haz clic para marcar puntos", use_container_width=True)

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
            guardar_proyectos(st.session_state.proyectos)
            st.rerun()
    
    st.write(f"**Puntos en este plano:** {len(plano_data['puntos'])}")
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🗑️ Eliminar último punto", key=f"eliminar_ultimo_{proyecto_actual}_{plano_actual}"):
            if plano_data["puntos"]:
                ultimo_num = len(plano_data["puntos"])
                plano_data["puntos"].pop()
                plano_data["data"] = [d for d in plano_data["data"] if d["Número"] != ultimo_num]
                guardar_proyectos(st.session_state.proyectos)
                st.rerun()
    with col2:
        if st.button("🧹 Limpiar todos los puntos", key=f"limpiar_puntos_{proyecto_actual}_{plano_actual}"):
            plano_data["puntos"] = []
            plano_data["data"] = []
            guardar_proyectos(st.session_state.proyectos)
            st.rerun()
    
    st.divider()
    
    if plano_data["puntos"]:
        st.subheader("📊 Mediciones por punto")
        
        for i, (x, y) in enumerate(plano_data["puntos"]):
            # Cargar valores existentes para que no se pierdan en reruns
            existing = next((d for d in plano_data["data"] if d["Número"] == i + 1), {})

            with st.expander(f"Punto {i+1} ({int(x)}, {int(y)})", expanded=False):

                # Botón eliminar este punto específico
                if st.button(f"🗑️ Eliminar punto {i+1}", key=f"del_punto_{proyecto_actual}_{plano_actual}_{i}"):
                    plano_data["puntos"].pop(i)
                    # Eliminar de data y renumerar
                    plano_data["data"] = [d for d in plano_data["data"] if d["Número"] != i + 1]
                    for d in plano_data["data"]:
                        if d["Número"] > i + 1:
                            d["Número"] -= 1
                    guardar_proyectos(st.session_state.proyectos)
                    st.rerun()

                # Tipo de área individual por punto
                tipos_area_list = list(RETILAP_REFERENCIA.keys())
                tipo_area_guardado = existing.get("TipoArea", general.get("tipo_area", tipos_area_list[0]))
                tipo_area_idx = tipos_area_list.index(tipo_area_guardado) if tipo_area_guardado in tipos_area_list else 0

                tipo_area_punto = st.selectbox(
                    "🏷️ Tipo de área según RETILAP",
                    tipos_area_list,
                    index=tipo_area_idx,
                    key=f"tipo_area_{proyecto_actual}_{plano_actual}_{i}"
                )
                valores_punto = RETILAP_REFERENCIA[tipo_area_punto]
                em_sugerido = valores_punto["Em"]
                uo_min = valores_punto["Uo"]
                st.info(f"Em requerida: **{em_sugerido} lx** · Uo mínima: **{uo_min}**")

                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    med1 = st.number_input("Med 1 (lx)", min_value=0.0, step=0.1,
                                           value=float(existing.get("Med1", 0.0)),
                                           key=f"m1_{proyecto_actual}_{plano_actual}_{i}")
                with col2:
                    med2 = st.number_input("Med 2 (lx)", min_value=0.0, step=0.1,
                                           value=float(existing.get("Med2", 0.0)),
                                           key=f"m2_{proyecto_actual}_{plano_actual}_{i}")
                with col3:
                    med3 = st.number_input("Med 3 (lx)", min_value=0.0, step=0.1,
                                           value=float(existing.get("Med3", 0.0)),
                                           key=f"m3_{proyecto_actual}_{plano_actual}_{i}")
                with col4:
                    med4 = st.number_input("Med 4 (lx)", min_value=0.0, step=0.1,
                                           value=float(existing.get("Med4", 0.0)),
                                           key=f"m4_{proyecto_actual}_{plano_actual}_{i}")

                foto_subida = st.file_uploader(f"Foto del punto {i+1} (opcional)", type=["jpg", "jpeg", "png"], key=f"foto_{proyecto_actual}_{plano_actual}_{i}")
                if foto_subida is not None:
                    plano_data["fotos"][i+1] = foto_subida.read()
                    guardar_proyectos(st.session_state.proyectos)

                nota = st.text_area("Notas / observaciones", height=80,
                                    value=existing.get("Nota", ""),
                                    key=f"nota_{proyecto_actual}_{plano_actual}_{i}")

                if all(v > 0 for v in [med1, med2, med3, med4]):
                    promedio = (med1 + med2 + med3 + med4) / 4
                    conforme = promedio >= em_sugerido
                    color = "green" if conforme else "red"
                    resultado = "✅ Conforme" if conforme else "❌ No conforme"

                    if conforme:
                        st.success(f"Promedio: **{round(promedio, 1)} lx** → {resultado}")
                    else:
                        st.error(f"Promedio: **{round(promedio, 1)} lx** → {resultado} (requiere ≥ {em_sugerido} lx)")

                    entrada = {
                        "Número": i + 1,
                        "Coordenadas": f"({int(x)}, {int(y)})",
                        "TipoArea": tipo_area_punto,
                        "Em_req": em_sugerido,
                        "Uo_min": uo_min,
                        "Med1": med1,
                        "Med2": med2,
                        "Med3": med3,
                        "Med4": med4,
                        "Promedio": round(promedio, 1),
                        "Resultado": resultado,
                        "Color": color,
                        "Nota": nota.strip(),
                        "Foto": foto_subida is not None
                    }

                    idx_existing = next((j for j, d in enumerate(plano_data["data"]) if d["Número"] == i + 1), None)
                    if idx_existing is not None:
                        plano_data["data"][idx_existing] = entrada
                    else:
                        plano_data["data"].append(entrada)

                    guardar_proyectos(st.session_state.proyectos)
        
        st.divider()
        
        if plano_data["data"]:
            st.subheader("🗺️ Mapa de puntos")
            df_plano = pd.DataFrame(plano_data["data"])
            draw_img = plano_img.copy()
            draw = ImageDraw.Draw(draw_img)
            # Fuente grande para los números
            try:
                font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 28)
            except Exception:
                font = ImageFont.load_default(size=28)

            for _, r in df_plano.iterrows():
                x, y = map(int, r["Coordenadas"].strip("()").split(", "))
                color = r["Color"]
                draw.ellipse((x - 24, y - 24, x + 24, y + 24), fill=color)

                texto = str(r["Número"])
                bbox = font.getbbox(texto)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                text_x = x - text_width // 2
                text_y = y - text_height // 2

                # Sombra para que el número resalte más
                for dx, dy in [(-1,-1),(1,-1),(-1,1),(1,1)]:
                    draw.text((text_x+dx, text_y+dy), texto, fill="black", font=font)
                draw.text((text_x, text_y), texto, fill="white", font=font)
            
            st.image(draw_img, caption=f"Mapa - {plano_actual}", use_container_width=True)
            
            st.subheader("📊 Tabla de Resultados")
            cols_mostrar = ["Número", "Coordenadas", "TipoArea", "Em_req", "Med1", "Med2", "Med3", "Med4", "Promedio", "Resultado"]
            cols_existentes = [c for c in cols_mostrar if c in df_plano.columns]
            st.dataframe(df_plano[cols_existentes].rename(columns={
                "TipoArea": "Tipo de Área",
                "Em_req": "Em req. (lx)"
            }), use_container_width=True)


# ============================================================================
# FUNCIÓN PRINCIPAL
# ============================================================================

def main():
    """Función principal de la aplicación"""
    st.set_page_config(page_title="Auditoría Iluminación RETILAP", layout="wide")
    
    inicializar_session_state()
    
    if st.session_state.pagina == "inicio":
        pagina_inicio()
    elif st.session_state.pagina == "nuevo_proyecto":
        pagina_nuevo_proyecto()
    elif st.session_state.pagina == "editar_proyecto":
        pagina_editar_proyecto()
    elif st.session_state.pagina == "editar_plano":
        pagina_editar_plano()


if __name__ == "__main__":
    main()
