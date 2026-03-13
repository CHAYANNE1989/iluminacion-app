import streamlit as st
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import base64
import json
from pdf2image import convert_from_bytes
from generar_word import generar_informe_word
from streamlit_image_coordinates import streamlit_image_coordinates
import io
from datetime import datetime
import anthropic

# ============================================================================
# CONFIGURACIÓN Y CONSTANTES
# ============================================================================

PROYECTOS_DIR = "dispositivos"

def get_device_id():
    params = st.query_params
    device_id = params.get("device_id", "default")
    device_id = "".join(c for c in str(device_id) if c.isalnum() or c in "-_")
    return device_id if device_id else "default"

def get_proyectos_file():
    os.makedirs(PROYECTOS_DIR, exist_ok=True)
    return os.path.join(PROYECTOS_DIR, f"proyectos_{get_device_id()}.json")

# ============================================================================
# ÁREAS RETILAP 2024 — tomadas del archivo oficial Excel
# Formato: "Categoría – Actividad": {"Em": lx, "Uo": valor}
# ============================================================================

RETILAP_REFERENCIA = {
    # Imprentas
    "Imprentas – Corte, estampado, grabado, máquinas de impresión": {"Em": 500, "Uo": 0.6},
    "Imprentas – Clasificación de papel e impresión a mano": {"Em": 500, "Uo": 0.6},
    # Oficinas
    "Oficinas – Escritura, mecanografía, lectura, procesamiento de datos": {"Em": 500, "Uo": 0.6},
    "Oficinas – Oficinas de tipo general, mecanografía y computación": {"Em": 300, "Uo": 0.19},
    "Oficinas – Oficinas abiertas": {"Em": 500, "Uo": 0.19},
    "Oficinas – Oficinas de dibujo": {"Em": 500, "Uo": 0.16},
    "Oficinas – Salas de conferencia": {"Em": 300, "Uo": 0.19},
    # Procesos químicos
    "Procesos químicos – Procesos automáticos": {"Em": 50, "Uo": 0.0},
    "Procesos químicos – Intervención ocasional": {"Em": 100, "Uo": 0.28},
    "Procesos químicos – Áreas generales en interior de fábricas": {"Em": 200, "Uo": 0.25},
    "Procesos químicos – Cuartos de control, laboratorios": {"Em": 300, "Uo": 0.19},
    "Procesos químicos – Industria farmacéutica": {"Em": 300, "Uo": 0.22},
    "Procesos químicos – Inspección": {"Em": 500, "Uo": 0.19},
    "Procesos químicos – Balanceo de colores": {"Em": 750, "Uo": 0.16},
    "Procesos químicos – Fabricación de llantas de caucho": {"Em": 300, "Uo": 0.22},
    # Confecciones
    "Confecciones – Costura": {"Em": 500, "Uo": 0.22},
    "Confecciones – Inspección": {"Em": 750, "Uo": 0.16},
    "Confecciones – Prensado": {"Em": 300, "Uo": 0.22},
    # Industria eléctrica
    "Industria eléctrica – Fabricación de cables": {"Em": 200, "Uo": 0.25},
    "Industria eléctrica – Ensamble de aparatos telefónicos": {"Em": 300, "Uo": 0.19},
    "Industria eléctrica – Ensamble de devanados": {"Em": 500, "Uo": 0.19},
    "Industria eléctrica – Ensamble aparatos de radio y TV": {"Em": 750, "Uo": 0.19},
    "Industria eléctrica – Ensamble componentes electrónicos ultra precisión": {"Em": 1000, "Uo": 0.16},
    # Industria alimenticia
    "Industria alimenticia – Áreas generales de trabajo": {"Em": 200, "Uo": 0.25},
    "Industria alimenticia – Procesos automáticos": {"Em": 150, "Uo": 0.0},
    "Industria alimenticia – Decoración manual, inspección": {"Em": 300, "Uo": 0.16},
    # Fundición
    "Fundición – Pozos de fundición": {"Em": 150, "Uo": 0.25},
    "Fundición – Moldeado basto, elaboración de machos": {"Em": 200, "Uo": 0.25},
    "Fundición – Moldeo fino, inspección": {"Em": 300, "Uo": 0.22},
    # Vidrio y cerámica
    "Vidrio y cerámica – Zona de hornos": {"Em": 100, "Uo": 0.25},
    "Vidrio y cerámica – Mezcla, moldeo, conformado y estufas": {"Em": 200, "Uo": 0.25},
    "Vidrio y cerámica – Terminado, esmaltado, envidriado": {"Em": 300, "Uo": 0.19},
    "Vidrio y cerámica – Pintura y decoración": {"Em": 500, "Uo": 0.16},
    "Vidrio y cerámica – Afilado, lentes y cristalería, trabajo fino": {"Em": 750, "Uo": 0.19},
    # Hierro y acero
    "Hierro y acero – Sin intervención manual": {"Em": 50, "Uo": 0.0},
    "Hierro y acero – Intervención ocasional": {"Em": 100, "Uo": 0.28},
    "Hierro y acero – Puestos permanentes en plantas de producción": {"Em": 200, "Uo": 0.25},
    "Hierro y acero – Plataformas de control e inspección": {"Em": 300, "Uo": 0.22},
    # Industria del cuero
    "Industria del cuero – Áreas generales de trabajo": {"Em": 200, "Uo": 0.25},
    "Industria del cuero – Prensado, corte, costura, producción de calzado": {"Em": 500, "Uo": 0.22},
    "Industria del cuero – Clasificación, adaptación y control de calidad": {"Em": 750, "Uo": 0.19},
    # Taller mecánica
    "Taller mecánica – Trabajo ocasional": {"Em": 150, "Uo": 0.25},
    "Taller mecánica – Trabajo basto en banca y maquinado, soldadura": {"Em": 200, "Uo": 0.22},
    "Taller mecánica – Maquinado y trabajo de media precisión": {"Em": 300, "Uo": 0.22},
    "Taller mecánica – Maquinado fino, inspección y ensayos": {"Em": 500, "Uo": 0.19},
    "Taller mecánica – Trabajo muy fino, calibración partes pequeñas": {"Em": 1000, "Uo": 0.09},
    # Pintura
    "Pintura – Inmersión, rociado basto": {"Em": 200, "Uo": 0.25},
    "Pintura – Pintura ordinaria, rociado y terminado": {"Em": 300, "Uo": 0.22},
    "Pintura – Pintura fina, rociado y terminado": {"Em": 500, "Uo": 0.19},
    "Pintura – Retoque y balanceo de colores": {"Em": 750, "Uo": 0.16},
    # Papel
    "Fábricas de papel – Elaboración de papel y cartón": {"Em": 200, "Uo": 0.25},
    "Fábricas de papel – Procesos automáticos": {"Em": 150, "Uo": 0.0},
    "Fábricas de papel – Inspección y clasificación": {"Em": 300, "Uo": 0.22},
    # Impresión y encuadernación
    "Impresión – Recintos con máquinas de impresión": {"Em": 300, "Uo": 0.19},
    "Impresión – Cuartos de composición y lecturas de prueba": {"Em": 500, "Uo": 0.19},
    "Impresión – Pruebas de precisión, retoque y grabado": {"Em": 750, "Uo": 0.16},
    "Impresión – Reproducción del color e impresión": {"Em": 1000, "Uo": 0.19},
    "Impresión – Grabado con acero y cobre": {"Em": 1500, "Uo": 0.16},
    "Impresión – Encuadernación": {"Em": 300, "Uo": 0.22},
    "Impresión – Decoración y estampado": {"Em": 500, "Uo": 0.19},
    # Textil
    "Industria textil – Rompimiento de paca, cardado, hilado": {"Em": 200, "Uo": 0.25},
    "Industria textil – Giro, embobinado, peinado, tintura": {"Em": 300, "Uo": 0.22},
    "Industria textil – Balanceo, rotación, entretejido, tejido": {"Em": 500, "Uo": 0.22},
    "Industria textil – Costura, desmonte e inspección": {"Em": 750, "Uo": 0.19},
    # Madera y muebles
    "Madera y muebles – Aserraderos": {"Em": 150, "Uo": 0.25},
    "Madera y muebles – Trabajo en banco y montaje": {"Em": 200, "Uo": 0.25},
    "Madera y muebles – Maquinado de madera": {"Em": 300, "Uo": 0.19},
    "Madera y muebles – Terminado e inspección final": {"Em": 500, "Uo": 0.19},
    # Salas (hospitales)
    "Salas hospitalarias – Iluminación general": {"Em": 50, "Uo": 0.22},
    "Salas hospitalarias – Examen": {"Em": 200, "Uo": 0.19},
    "Salas hospitalarias – Lectura": {"Em": 150, "Uo": 0.16},
    "Salas hospitalarias – Circulación nocturna": {"Em": 3, "Uo": 0.22},
    # Salas de examen
    "Salas de examen – Iluminación general": {"Em": 300, "Uo": 0.19},
    "Salas de examen – Inspección local": {"Em": 750, "Uo": 0.19},
    # Terapia intensiva
    "Terapia intensiva – Cabecera de la cama": {"Em": 30, "Uo": 0.19},
    "Terapia intensiva – Observación": {"Em": 200, "Uo": 0.19},
    "Terapia intensiva – Estación de enfermería": {"Em": 200, "Uo": 0.19},
    # Salas de operación
    "Salas de operación – Iluminación general": {"Em": 500, "Uo": 0.19},
    "Salas de operación – Iluminación local": {"Em": 10000, "Uo": 0.19},
    # Autopsia
    "Salas de autopsia – Iluminación general": {"Em": 500, "Uo": 0.19},
    "Salas de autopsia – Iluminación local": {"Em": 5000, "Uo": 0.0},
    # Consultorios
    "Consultorios – Iluminación general": {"Em": 300, "Uo": 0.19},
    "Consultorios – Iluminación local": {"Em": 500, "Uo": 0.19},
    # Farmacia y laboratorios
    "Farmacia y laboratorios – Iluminación general": {"Em": 300, "Uo": 0.19},
    "Farmacia y laboratorios – Iluminación local": {"Em": 500, "Uo": 0.19},
    # Comercio
    "Comercio – Grandes centros comerciales": {"Em": 500, "Uo": 0.19},
    "Comercio – Locales en cualquier parte": {"Em": 300, "Uo": 0.22},
    "Comercio – Supermercados": {"Em": 500, "Uo": 0.19},
    # Educación
    "Educación – Salones de clase (iluminación general)": {"Em": 300, "Uo": 0.19},
    "Educación – Tableros": {"Em": 300, "Uo": 0.19},
    "Educación – Elaboración de planos": {"Em": 500, "Uo": 0.16},
    "Educación – Salas de conferencias (iluminación general)": {"Em": 300, "Uo": 0.22},
    "Educación – Tableros en salas de conferencias": {"Em": 500, "Uo": 0.19},
    "Educación – Bancos de demostración": {"Em": 500, "Uo": 0.19},
    "Educación – Laboratorios": {"Em": 300, "Uo": 0.19},
    "Educación – Salas de arte": {"Em": 300, "Uo": 0.19},
    "Educación – Talleres": {"Em": 300, "Uo": 0.19},
    "Educación – Salas de asamblea": {"Em": 150, "Uo": 0.22},
}

# ============================================================================
# CSS GLOBAL
# ============================================================================

def aplicar_estilos():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .main-header {
        background: linear-gradient(135deg, #1a569a 0%, #0d3461 100%);
        color: white; padding: 1.1rem 1.4rem; border-radius: 12px;
        margin-bottom: 1.2rem; display: flex; align-items: center; gap: 14px;
        box-shadow: 0 4px 16px rgba(26,86,154,0.25);
    }
    .main-header h1 { margin: 0; font-size: 1.4rem; font-weight: 700; }
    .main-header p  { margin: 0; font-size: 0.82rem; opacity: 0.75; }
    .badge-ok  { background:#d1fae5; color:#065f46; border-radius:20px; padding:2px 10px; font-size:0.78rem; font-weight:600; }
    .badge-err { background:#fee2e2; color:#991b1b; border-radius:20px; padding:2px 10px; font-size:0.78rem; font-weight:600; }
    .badge-nd  { background:#f1f5f9; color:#475569; border-radius:20px; padding:2px 10px; font-size:0.78rem; font-weight:600; }
    .em-box { background:linear-gradient(90deg,#eff6ff,#f0f9ff); border:1px solid #bfdbfe;
              border-radius:8px; padding:0.45rem 1rem; margin:0.3rem 0 0.7rem;
              font-size:0.86rem; color:#1e40af; }
    .recomendacion-box { background:#f0fdf4; border:1px solid #bbf7d0; border-radius:8px;
                         padding:0.8rem 1rem; margin-top:0.5rem; font-size:0.87rem; color:#14532d; }
    div[data-testid="stButton"] > button { border-radius: 8px; font-weight: 500; }
    #MainMenu, footer { visibility: hidden; }
    </style>
    """, unsafe_allow_html=True)


# ============================================================================
# IA — GENERAR RECOMENDACIONES
# ============================================================================

def generar_recomendaciones_ia(puntos_data):
    """
    Recibe lista de dicts con info de los puntos.
    Retorna texto con recomendaciones generadas por Claude.
    """
    try:
        client = anthropic.Anthropic()

        # Construir resumen de puntos para el prompt
        resumen = []
        for p in puntos_data:
            estado = "CONFORME" if "✅" in str(p.get("Resultado","")) else "DEFICIENTE"
            resumen.append(
                f"- Punto {p['Número']} ({p.get('TipoArea','')}):"
                f" Promedio={p.get('Promedio',0)} lx,"
                f" Em requerida={p.get('Em_req',0)} lx,"
                f" Tipo iluminación={p.get('TipoIluminacion','')},"
                f" Lámpara={p.get('TipoLampara','')},"
                f" Observación: {p.get('Nota','Sin observación')}."
                f" Estado: {estado}."
            )

        prompt = f"""Eres un experto en higiene y seguridad industrial especializado en iluminación según la norma RETILAP 2024 de Colombia.

Analiza los siguientes puntos de medición de iluminancia y genera recomendaciones técnicas:

{chr(10).join(resumen)}

Instrucciones:
- Si varios puntos tienen el mismo problema o condiciones similares, agrupa la recomendación en una sola recomendación general.
- Si hay puntos con problemas diferentes, genera una recomendación específica para cada caso.
- Sé concreto y técnico: menciona acciones como "aumentar densidad de luminarias", "reemplazar lámparas por LED de mayor flujo luminoso", "instalar luminarias localizadas", etc.
- Si hay puntos conformes, menciona brevemente que se deben mantener las condiciones.
- Usa lenguaje técnico pero claro. Máximo 250 palabras.
- Responde SOLO con las recomendaciones, sin introducción ni conclusión."""

        message = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=500,
            messages=[{"role": "user", "content": prompt}]
        )
        return message.content[0].text

    except Exception as e:
        return f"⚠️ No se pudo generar la recomendación automática: {e}"


# ============================================================================
# PERSISTENCIA
# ============================================================================

def cargar_proyectos():
    if not os.path.exists(get_proyectos_file()):
        return {}
    try:
        with open(get_proyectos_file(), "r", encoding="utf-8") as f:
            data = json.load(f)
        proyectos = {}
        for p_name, p_data in data.items():
            proyectos[p_name] = {"general": p_data["general"], "planos": {}}
            for pl_name, pl_info in p_data["planos"].items():
                pd_ = {
                    "puntos": pl_info["puntos"],
                    "data":   pl_info["data"],
                    "fotos":  pl_info.get("fotos", {})
                }
                if "img_base64" in pl_info:
                    try:
                        img_b = base64.b64decode(pl_info["img_base64"])
                        img   = Image.open(io.BytesIO(img_b))
                        if img.mode != 'RGB': img = img.convert('RGB')
                        pd_["img"] = img
                    except:
                        pd_["img"] = None
                else:
                    pd_["img"] = None
                proyectos[p_name]["planos"][pl_name] = pd_
        return proyectos
    except Exception as e:
        st.error(f"Error al cargar: {e}")
        return {}


def guardar_proyectos(proyectos):
    try:
        serial = {}
        for p_name, p_data in proyectos.items():
            serial[p_name] = {"general": p_data["general"].copy(), "planos": {}}
            for pl_name, pl_info in p_data["planos"].items():
                pd_ = {
                    "puntos": pl_info["puntos"].copy() if isinstance(pl_info["puntos"], list) else [],
                    "data":   [r.copy() for r in pl_info["data"]] if isinstance(pl_info["data"], list) else [],
                    "fotos":  {}
                }
                for k, v in pl_info.get("fotos", {}).items():
                    pd_["fotos"][str(k)] = (base64.b64encode(v).decode() if isinstance(v, bytes) else v)
                if pl_info.get("img"):
                    try:
                        buf = io.BytesIO()
                        pl_info["img"].save(buf, format="PNG")
                        pd_["img_base64"] = base64.b64encode(buf.getvalue()).decode()
                    except Exception as e:
                        st.warning(f"⚠️ No se guardó imagen '{pl_name}': {e}")
                serial[p_name]["planos"][pl_name] = pd_
        with open(get_proyectos_file(), "w", encoding="utf-8") as f:
            json.dump(serial, f, ensure_ascii=False, indent=4)
    except Exception as e:
        st.error(f"Error al guardar: {e}")


def cargar_foto_punto(plano_info, num):
    v = plano_info.get("fotos", {}).get(str(num)) or plano_info.get("fotos", {}).get(num)
    if isinstance(v, str):
        try:    return base64.b64decode(v)
        except: return None
    return v if isinstance(v, bytes) else None


# ============================================================================
# HELPER: dibujar puntos proporcionales
# ============================================================================

def dibujar_puntos(img, data_rows):
    draw_img = img.copy()
    if not data_rows:
        return draw_img
    draw  = ImageDraw.Draw(draw_img)
    lado  = min(draw_img.width, draw_img.height)
    radio = max(10, min(22, int(lado * 0.012)))
    fsize = max(9,  min(16, radio - 1))
    try:
        font = ImageFont.truetype(
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", fsize)
    except:
        font = ImageFont.load_default()
    for row in data_rows:
        try:
            raw = str(row["Coordenadas"]).strip("()").split(", ")
            cx, cy = float(raw[0]), float(raw[1])
            x = int(cx * draw_img.width)  if cx <= 1.0 else int(cx)
            y = int(cy * draw_img.height) if cy <= 1.0 else int(cy)
            clr = row.get("Color", "gray")
            draw.ellipse((x-radio-1, y-radio-1, x+radio+1, y+radio+1), fill="white")
            draw.ellipse((x-radio,   y-radio,   x+radio,   y+radio),   fill=clr)
            txt = str(row["Número"])
            bb  = font.getbbox(txt)
            tw, th = bb[2]-bb[0], bb[3]-bb[1]
            tx, ty = x - tw//2, y - th//2 - 1
            for dx, dy in [(-1,-1),(1,-1),(-1,1),(1,1)]:
                draw.text((tx+dx, ty+dy), txt, fill="black", font=font)
            draw.text((tx, ty), txt, fill="white", font=font)
        except:
            pass
    return draw_img


# ============================================================================
# GRÁFICA DE CONFORMIDAD — solo % adecuados vs deficientes
# ============================================================================

def grafica_conformidad(data_rows, titulo=""):
    """Genera gráfica de torta simple: % conformes vs deficientes."""
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt

        total     = len(data_rows)
        conformes = sum(1 for r in data_rows if "✅" in str(r.get("Resultado","")))
        deficientes = total - conformes

        if total == 0:
            return None

        pct_conf = round(conformes / total * 100, 1)
        pct_def  = round(deficientes / total * 100, 1)

        fig, ax = plt.subplots(figsize=(4.5, 4.5), facecolor='#f8fafc')
        ax.set_facecolor('#f8fafc')

        valores = [pct_conf, pct_def] if deficientes > 0 else [100]
        labels  = [f"Adecuados\n{pct_conf}%", f"Deficientes\n{pct_def}%"] if deficientes > 0 else [f"Adecuados\n100%"]
        colores = ['#22c55e', '#ef4444'] if deficientes > 0 else ['#22c55e']
        explotar= [0.04, 0.04] if deficientes > 0 else [0]

        wedges, texts = ax.pie(
            valores, labels=labels, colors=colores,
            explode=explotar, startangle=90,
            textprops={'fontsize': 11, 'fontweight': 'bold'},
            wedgeprops={'linewidth': 2, 'edgecolor': 'white'}
        )

        ax.set_title(titulo or "Conformidad RETILAP",
                     fontsize=12, fontweight='bold', color='#1a3a5c', pad=12)

        # Texto central
        ax.text(0, 0, f"{total}\npuntos",
                ha='center', va='center',
                fontsize=10, color='#475569', fontweight='bold')

        buf = io.BytesIO()
        plt.savefig(buf, format='PNG', bbox_inches='tight', dpi=130,
                    facecolor='#f8fafc')
        plt.close(fig)
        buf.seek(0)
        return buf.getvalue()
    except Exception as e:
        st.warning(f"No se pudo generar la gráfica: {e}")
        return None


# ============================================================================
# REPORTES CSV Y PDF
# ============================================================================

def generar_reporte_csv(proyecto_data, proyecto_nombre):
    rows = []
    for pln, pi in proyecto_data["planos"].items():
        for r in pi.get("data", []):
            rows.append({
                "Proyecto": proyecto_nombre, "Plano": pln,
                "Punto": r["Número"], "Tipo de Área": r.get("TipoArea",""),
                "Em req.(lx)": r.get("Em_req",""), "Med1": r.get("Med1",""),
                "Med2": r.get("Med2",""), "Med3": r.get("Med3",""),
                "Med4": r.get("Med4",""), "Promedio": r.get("Promedio",""),
                "Resultado": r.get("Resultado",""), "Nota": r.get("Nota","")
            })
    if rows:
        return pd.DataFrame(rows).to_csv(index=False).encode('utf-8')
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
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=letter,
                                rightMargin=2*cm, leftMargin=2*cm,
                                topMargin=2*cm, bottomMargin=2*cm)
        S   = getSampleStyleSheet()
        eTi = ParagraphStyle('T', parent=S['Title'], fontSize=18,
                              textColor=colors.HexColor('#1a3a5c'), alignment=TA_CENTER)
        eSu = ParagraphStyle('S', parent=S['Normal'], fontSize=11,
                              textColor=colors.HexColor('#2c6fad'), alignment=TA_CENTER)
        eSe = ParagraphStyle('H', parent=S['Heading2'], fontSize=12,
                              textColor=colors.HexColor('#1a3a5c'), spaceBefore=10, spaceAfter=4)
        eNo = ParagraphStyle('N', parent=S['Normal'], fontSize=9, spaceAfter=3)
        ePi = ParagraphStyle('P', parent=S['Normal'], fontSize=7,
                              textColor=colors.grey, alignment=TA_CENTER)

        g     = proyecto_data.get("general", {})
        story = [
            Spacer(1, 0.8*inch),
            Paragraph("REPORTE DE AUDITORÍA DE ILUMINACIÓN", eTi),
            Paragraph("Norma RETILAP 2024", eSu),
            Spacer(1, 0.2*inch),
            HRFlowable(width="100%", thickness=2, color=colors.HexColor('#2c6fad')),
            Spacer(1, 0.2*inch),
        ]
        info = [
            ["Proyecto:", proyecto_nombre],
            ["Orden:", g.get("numero_orden","N/A")],
            ["Empresa:", g.get("nombre_empresa","N/A")],
            ["NIT:", g.get("nit","N/A")],
            ["Dirección:", g.get("direccion","N/A")],
            ["Ciudad:", g.get("sede","N/A")],
            ["Fecha:", g.get("fecha","N/A")],
            ["Higienista:", g.get("responsable_higienista","N/A")],
            ["Resolución SST:", g.get("resolucion","N/A")],
            ["Generado:", datetime.now().strftime("%d/%m/%Y %H:%M")],
        ]
        tI = Table(info, colWidths=[3.2*cm, 12.8*cm])
        tI.setStyle(TableStyle([
            ('FONTNAME',(0,0),(0,-1),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),9),
            ('TEXTCOLOR',(0,0),(0,-1),colors.HexColor('#1a3a5c')),
            ('ROWBACKGROUNDS',(0,0),(-1,-1),[colors.HexColor('#f0f4f8'),colors.white]),
            ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#ccddee')),
            ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
            ('LEFTPADDING',(0,0),(-1,-1),8),
        ]))
        story.append(tI); story.append(Spacer(1, 0.25*inch))

        # Resumen general
        tot = conf = 0
        for pi in proyecto_data["planos"].values():
            for r in pi.get("data",[]):
                tot += 1
                if "✅" in str(r.get("Resultado","")): conf += 1
        if tot > 0:
            pct = round(conf/tot*100, 1)
            rD  = [["Total","Conformes","Deficientes","% Conformidad"],
                   [str(tot), str(conf), str(tot-conf), f"{pct}%"]]
            tR  = Table(rD, colWidths=[4*cm]*4)
            tR.setStyle(TableStyle([
                ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1a3a5c')),
                ('TEXTCOLOR',(0,0),(-1,0),colors.white),
                ('FONTNAME',(0,0),(-1,-1),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),10),
                ('ALIGN',(0,0),(-1,-1),'CENTER'),
                ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#ccddee')),
                ('BACKGROUND',(0,1),(-1,1),colors.HexColor('#f0f4f8')),
                ('TEXTCOLOR',(3,1),(3,1),
                 colors.HexColor('#27ae60') if pct>=80 else colors.HexColor('#e74c3c')),
            ]))
            story.append(tR)

            # Gráfica de conformidad en el PDF
            grafica_bytes = grafica_conformidad(
                [r for pi in proyecto_data["planos"].values() for r in pi.get("data",[])],
                "Conformidad General")
            if grafica_bytes:
                story.append(Spacer(1, 0.2*inch))
                story.append(RLImage(io.BytesIO(grafica_bytes), width=3.5*inch, height=3.5*inch))

        story.append(PageBreak())

        for pln, pi in proyecto_data["planos"].items():
            drows = pi.get("data",[])
            pimg  = pi.get("img")
            story.append(Paragraph(f"Plano: {pln}", eSe))
            story.append(HRFlowable(width="100%",thickness=1,color=colors.HexColor('#ccddee')))
            story.append(Spacer(1,0.08*inch))

            if pimg and drows:
                try:
                    an = dibujar_puntos(pimg, drows)
                    if an.width > 1400:
                        r = 1400/an.width
                        an = an.resize((1400, int(an.height*r)), Image.LANCZOS)
                    b = io.BytesIO(); an.save(b, format="PNG"); b.seek(0)
                    pw = letter[0]-4*cm
                    ph = min(pw*an.height/an.width, 5*inch)
                    story += [RLImage(b, width=pw, height=ph), Spacer(1,0.12*inch)]
                except Exception as e:
                    story.append(Paragraph(f"(Error mapa: {e})", eNo))

            if drows:
                enc = ["#","Área","Em req.","Med1","Med2","Med3","Med4","Prom.","Result.","Obs."]
                fls = [enc]+[[
                    str(r.get("Número","")), str(r.get("TipoArea",""))[:28],
                    f"{r.get('Em_req','?')} lx", str(r.get("Med1","")),
                    str(r.get("Med2","")), str(r.get("Med3","")), str(r.get("Med4","")),
                    str(r.get("Promedio","")),
                    "✓ OK" if "✅" in str(r.get("Resultado","")) else "✗ NO",
                    str(r.get("Nota",""))[:30]
                ] for r in drows]
                cw  = [0.7*cm,3.8*cm,1.7*cm,1.3*cm,1.3*cm,1.3*cm,1.3*cm,1.7*cm,1.8*cm,2.7*cm]
                tab = Table(fls, colWidths=cw, repeatRows=1)
                ts  = [
                    ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1a3a5c')),
                    ('TEXTCOLOR',(0,0),(-1,0),colors.white),
                    ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                    ('FONTSIZE',(0,0),(-1,-1),7.5),
                    ('ALIGN',(0,0),(-1,-1),'CENTER'),
                    ('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#aaccee')),
                    ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f0f4f8')]),
                    ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
                ]
                for i, r in enumerate(drows, 1):
                    c = (colors.HexColor('#27ae60') if "✅" in str(r.get("Resultado",""))
                         else colors.HexColor('#e74c3c'))
                    ts += [('TEXTCOLOR',(8,i),(8,i),c),('FONTNAME',(8,i),(8,i),'Helvetica-Bold')]
                tab.setStyle(TableStyle(ts))
                story.append(tab)
            else:
                story.append(Paragraph("Sin mediciones.", eNo))
            story += [Spacer(1,0.15*inch), PageBreak()]

        story += [HRFlowable(width="100%",thickness=1,color=colors.grey),
                  Spacer(1,0.08*inch),
                  Paragraph(f"RETILAP 2024 · {datetime.now().strftime('%d/%m/%Y %H:%M')}", ePi)]
        doc.build(story)
        buf.seek(0)
        return buf.getvalue()
    except ImportError:
        st.error("❌ Instala reportlab en requirements.txt"); return None
    except Exception as e:
        st.error(f"❌ Error PDF: {e}"); return None


# ============================================================================
# SESSION STATE
# ============================================================================

def inicializar_session_state():
    try:
        if "proyectos" not in st.session_state:
            st.session_state.proyectos = cargar_proyectos()
        if "pagina" not in st.session_state:
            st.session_state.pagina = "inicio"
        if "proyecto_actual" not in st.session_state:
            st.session_state.proyecto_actual = None
    except Exception as e:
        st.error(f"Error al inicializar: {e}")
        st.session_state.proyectos = {}
        st.session_state.pagina = "inicio"
        st.session_state.proyecto_actual = None


# ============================================================================
# PÁGINA: INICIO
# ============================================================================

def pagina_inicio():
    st.markdown("""
    <div class="main-header">
      <span style="font-size:2.2rem">💡</span>
      <div><h1>LuxOMeter PRO</h1>
      <p>Auditoría de Iluminación · Norma RETILAP 2024</p></div>
    </div>""", unsafe_allow_html=True)

    c1, c2 = st.columns([3,1])
    with c1: st.subheader("📋 Proyectos")
    with c2:
        if st.button("➕ Nuevo Proyecto", use_container_width=True, key="btn_np"):
            st.session_state.pagina = "nuevo_proyecto"; st.rerun()

    if not st.session_state.proyectos:
        st.info("ℹ️ No hay proyectos guardados. Crea uno nuevo para comenzar.")
        return

    for idx, (pnombre, pdata) in enumerate(st.session_state.proyectos.items()):
        g = pdata["general"]
        tot = conf = 0
        all_data = []
        for pi in pdata["planos"].values():
            for d in pi.get("data",[]):
                tot += 1; all_data.append(d)
                if "✅" in str(d.get("Resultado","")): conf += 1

        with st.container(border=True):
            ci, cb = st.columns([3,1])
            with ci:
                st.markdown(f"**{g.get('nombre_empresa','Sin nombre')}**")
                st.caption(f"📋 OT: {g.get('numero_orden','N/A')}  |  "
                           f"📍 {g.get('sede','N/A')}  |  📅 {g.get('fecha','N/A')}")
                if tot > 0:
                    pct   = round(conf/tot*100)
                    badge = "ok" if pct>=80 else "err"
                    icono = "✅" if pct>=80 else "⚠️"
                    st.markdown(f"<span class='badge-{badge}'>{icono} {conf}/{tot} "
                                f"puntos conformes ({pct}%)</span>", unsafe_allow_html=True)

                    # Gráfica de torta
                    graf = grafica_conformidad(all_data)
                    if graf:
                        st.image(graf, width=220)
                else:
                    st.markdown("<span class='badge-nd'>Sin mediciones</span>",
                                unsafe_allow_html=True)

            with cb:
                if st.button("✏️ Editar", key=f"ed_{idx}", use_container_width=True):
                    st.session_state.proyecto_actual = pnombre
                    st.session_state.pagina = "editar_proyecto"; st.rerun()

                csv_d = generar_reporte_csv(pdata, pnombre)
                if csv_d:
                    st.download_button("📊 CSV", data=csv_d,
                        file_name=f"RETILAP_{pnombre[:18].replace(' ','_')}.csv",
                        mime="text/csv", key=f"csv_{idx}", use_container_width=True)

                pdf_d = generar_reporte_pdf(pdata, pnombre)
                if pdf_d:
                    st.download_button("📄 PDF", data=pdf_d,
                        file_name=f"RETILAP_{pnombre[:18].replace(' ','_')}.pdf",
                        mime="application/pdf", key=f"pdf_{idx}", use_container_width=True)

                if st.button("📝 Word", key=f"word_{idx}", use_container_width=True):
                    with st.spinner("Generando Word..."):
                        try:
                            todas_med = []
                            for pln, pi in pdata.get("planos",{}).items():
                                for d in pi.get("data",[]):
                                    todas_med.append({
                                        "num": d.get("Número",0),
                                        "area": d.get("TipoArea",""),
                                        "tipo_iluminacion": d.get("TipoIluminacion",""),
                                        "tipo_lampara": d.get("TipoLampara",""),
                                        "ubicacion_luminaria": d.get("UbicacionLuminaria",""),
                                        "med1": d.get("Med1",0),"med2": d.get("Med2",0),
                                        "med3": d.get("Med3",0),"med4": d.get("Med4",0),
                                        "promedio": d.get("Promedio",0),
                                        "em_req": d.get("Em_req",0),
                                        "resultado": d.get("Resultado",""),
                                        "nota": d.get("Nota",""),
                                        "recomendacion": d.get("Recomendacion",""),
                                    })
                            plano_imgs = {
                                pln: dibujar_puntos(pi["img"], pi["data"])
                                for pln, pi in pdata.get("planos",{}).items()
                                if pi.get("img") and pi.get("data")
                            }
                            word_buf = generar_informe_word(g, todas_med, plano_imgs)
                            fname = (f"Informe_RETILAP_"
                                     f"{g.get('nombre_empresa','').replace(' ','_')}"
                                     f"_{datetime.now().strftime('%Y%m%d')}.docx")
                            st.download_button("⬇️ Descargar Word", data=word_buf,
                                file_name=fname,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"dlw_{idx}")
                            st.success("✅ Listo.")
                        except Exception as e:
                            st.error(f"Error: {e}")

                if st.button("🗑️ Eliminar", key=f"del_{idx}", use_container_width=True):
                    del st.session_state.proyectos[pnombre]
                    guardar_proyectos(st.session_state.proyectos); st.rerun()


# ============================================================================
# PÁGINA: NUEVO PROYECTO
# ============================================================================

def pagina_nuevo_proyecto():
    st.markdown('<div class="main-header"><span style="font-size:2rem">➕</span>'
                '<div><h1>Nuevo Proyecto</h1></div></div>', unsafe_allow_html=True)
    if st.button("← Volver", key="volver_np"):
        st.session_state.pagina = "inicio"; st.rerun()

    with st.form("form_nuevo"):
        st.subheader("🏢 Datos de la Empresa")
        c1, c2 = st.columns(2)
        with c1:
            num_orden = st.text_input("N° Orden de Trabajo *")
            nom_emp   = st.text_input("Razón Social *")
            nit       = st.text_input("NIT")
            direccion = st.text_input("Dirección")
        with c2:
            sede    = st.text_input("Ciudad de Ejecución *")
            tel     = st.text_input("Teléfono")
            fecha   = st.date_input("Fecha de Actividad")
            resp_e  = st.text_input("Responsable (empresa)")
        st.subheader("👷 Datos del Higienista")
        c3, c4 = st.columns(2)
        with c3: resp_h = st.text_input("Nombre del Higienista")
        with c4: resol  = st.text_input("N° Resolución Licencia SST")
        ok = st.form_submit_button("✅ Crear Proyecto", use_container_width=True)

    if ok:
        if num_orden and nom_emp and sede:
            pnombre = f"{nom_emp} - {sede} ({fecha.strftime('%Y-%m-%d')})"
            if pnombre in st.session_state.proyectos:
                st.error("❌ Ya existe un proyecto con ese nombre")
            else:
                st.session_state.proyectos[pnombre] = {
                    "general": {
                        "numero_orden": num_orden, "nombre_empresa": nom_emp,
                        "nit": nit, "direccion": direccion, "sede": sede,
                        "telefono": tel, "fecha": fecha.strftime('%d/%m/%Y'),
                        "mes_anio": fecha.strftime('%B de %Y').capitalize(),
                        "responsable_empresa": resp_e,
                        "responsable_higienista": resp_h,
                        "resolucion": resol,
                        "tipo_area": list(RETILAP_REFERENCIA.keys())[0],
                    },
                    "planos": {}
                }
                guardar_proyectos(st.session_state.proyectos)
                st.session_state.proyecto_actual = pnombre
                st.session_state.pagina = "editar_proyecto"; st.rerun()
        else:
            st.error("❌ Completa los campos obligatorios (*)")


# ============================================================================
# PÁGINA: EDITAR PROYECTO
# ============================================================================

def pagina_editar_proyecto():
    pnombre = st.session_state.proyecto_actual
    pdata   = st.session_state.proyectos[pnombre]
    g       = pdata["general"]
    for k in ["nit","direccion","telefono","responsable_empresa",
              "responsable_higienista","resolucion","sede","fecha"]:
        g.setdefault(k, "")

    st.markdown(
        f'<div class="main-header"><span style="font-size:2rem">✏️</span>'
        f'<div><h1>{g.get("nombre_empresa","Proyecto")}</h1>'
        f'<p>{g.get("sede","")} · {g.get("fecha","")}</p></div></div>',
        unsafe_allow_html=True)

    cv, ce = st.columns([1,1])
    with cv:
        if st.button("← Volver", key="volver_ep"):
            st.session_state.pagina = "inicio"; st.rerun()
    with ce:
        if st.button("⚙️ Editar datos del proyecto", key="toggle_edit"):
            st.session_state["_show_edit"] = not st.session_state.get("_show_edit", False)
            st.rerun()

    if st.session_state.get("_show_edit", False):
        with st.expander("📝 Editar información del proyecto", expanded=True):
            with st.form("form_edit_gral"):
                c1, c2 = st.columns(2)
                with c1:
                    v_or = st.text_input("N° Orden",     value=g.get("numero_orden",""))
                    v_em = st.text_input("Razón Social", value=g.get("nombre_empresa",""))
                    v_ni = st.text_input("NIT",          value=g.get("nit",""))
                    v_di = st.text_input("Dirección",    value=g.get("direccion",""))
                with c2:
                    v_se = st.text_input("Ciudad",       value=g.get("sede",""))
                    v_te = st.text_input("Teléfono",     value=g.get("telefono",""))
                    v_re = st.text_input("Responsable empresa",
                                         value=g.get("responsable_empresa",""))
                c3, c4 = st.columns(2)
                with c3: v_hi = st.text_input("Higienista",     value=g.get("responsable_higienista",""))
                with c4: v_rs = st.text_input("N° Resolución SST", value=g.get("resolucion",""))
                if st.form_submit_button("💾 Guardar cambios", use_container_width=True):
                    g.update({"numero_orden":v_or,"nombre_empresa":v_em,"nit":v_ni,
                              "direccion":v_di,"sede":v_se,"telefono":v_te,
                              "responsable_empresa":v_re,"responsable_higienista":v_hi,
                              "resolucion":v_rs})
                    guardar_proyectos(st.session_state.proyectos)
                    st.session_state["_show_edit"] = False
                    st.success("✅ Datos actualizados"); st.rerun()

    st.divider()
    st.subheader("📐 Planos")

    with st.expander("➕ Agregar plano", expanded=not bool(pdata["planos"])):
        c1, c2 = st.columns([2,2])
        with c1: plano_nombre = st.text_input("Nombre del plano", key="inp_pnombre")
        with c2: up_plano     = st.file_uploader("Archivo JPG o PDF",
                                                  type=["jpg","jpeg","pdf"], key="up_plano")
        if plano_nombre and up_plano:
            if st.button("✅ Agregar", key="btn_add_plano"):
                if plano_nombre in pdata["planos"]:
                    st.warning("⚠️ Ese plano ya existe")
                else:
                    try:
                        img = (convert_from_bytes(up_plano.read())[0]
                               if up_plano.type == "application/pdf"
                               else Image.open(up_plano))
                        if img.mode != 'RGB': img = img.convert('RGB')
                        if img.width > 1920:
                            r = 1920/img.width
                            img = img.resize((1920, int(img.height*r)), Image.LANCZOS)
                        pdata["planos"][plano_nombre] = {
                            "img": img, "puntos": [], "data": [], "fotos": {}}
                        guardar_proyectos(st.session_state.proyectos)
                        st.success(f"✅ Plano '{plano_nombre}' agregado"); st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error: {e}")

    if pdata["planos"]:
        for pln, pi in list(pdata["planos"].items()):
            n_pts  = len(pi.get("data",[]))
            n_conf = sum(1 for d in pi.get("data",[]) if "✅" in str(d.get("Resultado","")))
            c1, c2, c3 = st.columns([3,1,1])
            with c1:
                res = f"  ({n_conf} conformes)" if n_pts else ""
                st.write(f"📄 **{pln}** — {n_pts} punto{'s' if n_pts!=1 else ''}{res}")
            with c2:
                if st.button("📍 Editar puntos", key=f"ep_{pln}"):
                    st.session_state.plano_actual = pln
                    st.session_state.pagina = "editar_plano"; st.rerun()
            with c3:
                if st.button("🗑️ Eliminar", key=f"delp_{pln}"):
                    del pdata["planos"][pln]
                    guardar_proyectos(st.session_state.proyectos); st.rerun()
    else:
        st.info("ℹ️ Agrega un plano para comenzar")


# ============================================================================
# PÁGINA: EDITAR PLANO
# ============================================================================

def pagina_editar_plano():
    if "plano_actual" not in st.session_state:
        st.session_state.pagina = "inicio"; st.rerun()

    pnombre   = st.session_state.proyecto_actual
    pl_nombre = st.session_state.plano_actual
    pdata     = st.session_state.proyectos[pnombre]
    pl_data   = pdata["planos"][pl_nombre]
    g         = pdata["general"]
    plano_img = pl_data.get("img")

    st.markdown(
        f'<div class="main-header"><span style="font-size:2rem">📍</span>'
        f'<div><h1>{pl_nombre}</h1>'
        f'<p>{g.get("nombre_empresa","")} · {g.get("sede","")}</p></div></div>',
        unsafe_allow_html=True)

    if st.button("← Volver al proyecto", key="volver_pl"):
        st.session_state.pagina = "editar_proyecto"; st.rerun()

    if plano_img is None:
        st.error("⚠️ La imagen no pudo cargarse. Sube el plano nuevamente.")
        return

    img_mostrar = dibujar_puntos(plano_img, pl_data["data"]) if pl_data["data"] else plano_img
    st.image(img_mostrar, caption="Haz clic sobre el plano para agregar un punto",
             use_container_width=True)

    clicked = streamlit_image_coordinates(
        plano_img,
        key=f"clicker_{pnombre}_{pl_nombre}",
        height=plano_img.height,
        width=plano_img.width,
    )
    if clicked is not None:
        xn = clicked["x"] / plano_img.width
        yn = clicked["y"] / plano_img.height
        if not any(abs(px-xn)<0.01 and abs(py-yn)<0.01
                   for px, py in pl_data["puntos"]):
            pl_data["puntos"].append((xn, yn))
            guardar_proyectos(st.session_state.proyectos); st.rerun()

    cm1, cm2, cm3 = st.columns(3)
    with cm1: st.metric("Puntos marcados", len(pl_data["puntos"]))
    with cm2:
        if st.button("🗑️ Eliminar último", key=f"del_ul_{pl_nombre}"):
            if pl_data["puntos"]:
                n = len(pl_data["puntos"])
                pl_data["puntos"].pop()
                pl_data["data"] = [d for d in pl_data["data"] if d["Número"] != n]
                guardar_proyectos(st.session_state.proyectos); st.rerun()
    with cm3:
        if st.button("🧹 Limpiar todos", key=f"limpiar_{pl_nombre}"):
            pl_data["puntos"] = []; pl_data["data"] = []
            guardar_proyectos(st.session_state.proyectos); st.rerun()

    st.divider()

    if not pl_data["puntos"]:
        st.info("Haz clic sobre el plano para marcar el primer punto.")
        return

    st.subheader("📊 Mediciones por punto")

    TIPOS = list(RETILAP_REFERENCIA.keys())
    ILUM  = ["Natural","Artificial","Mixta"]
    LAMP  = ["LED","Fluorescente","Incandescente","Halógeno","Otro"]
    UBIC  = ["Localizado","Lateral","Frontal","Trasera","Cenital"]

    for i, (xn, yn) in enumerate(pl_data["puntos"]):
        x  = int(xn * plano_img.width)
        y  = int(yn * plano_img.height)
        ex = next((d for d in pl_data["data"] if d["Número"] == i+1), {})
        r_actual = ex.get("Resultado","")
        icono = "✅" if "✅" in r_actual else ("❌" if "❌" in r_actual else "⏳")

        with st.expander(f"{icono} Punto {i+1}  ·  ({x}, {y})", expanded=False):

            if st.button(f"🗑️ Eliminar punto {i+1}",
                         key=f"delpt_{pnombre}_{pl_nombre}_{i}"):
                pl_data["puntos"].pop(i)
                pl_data["data"] = [d for d in pl_data["data"] if d["Número"] != i+1]
                for d in pl_data["data"]:
                    if d["Número"] > i+1: d["Número"] -= 1
                guardar_proyectos(st.session_state.proyectos); st.rerun()

            ta_g = ex.get("TipoArea", TIPOS[0])
            tipo_area = st.selectbox("🏷️ Tipo de área RETILAP", TIPOS,
                index=TIPOS.index(ta_g) if ta_g in TIPOS else 0,
                key=f"ta_{pnombre}_{pl_nombre}_{i}")
            em_req = RETILAP_REFERENCIA[tipo_area]["Em"]
            uo_min = RETILAP_REFERENCIA[tipo_area]["Uo"]
            st.markdown(f'<div class="em-box">⚡ Em requerida: <strong>{em_req} lx</strong>'
                        f'&nbsp;·&nbsp; Uo mínima: <strong>{uo_min}</strong></div>',
                        unsafe_allow_html=True)

            c1, c2, c3, c4 = st.columns(4)
            with c1: med1=st.number_input("Med 1 (lx)",min_value=0.0,step=1.0,
                value=float(ex.get("Med1",0)),key=f"m1_{pnombre}_{pl_nombre}_{i}")
            with c2: med2=st.number_input("Med 2 (lx)",min_value=0.0,step=1.0,
                value=float(ex.get("Med2",0)),key=f"m2_{pnombre}_{pl_nombre}_{i}")
            with c3: med3=st.number_input("Med 3 (lx)",min_value=0.0,step=1.0,
                value=float(ex.get("Med3",0)),key=f"m3_{pnombre}_{pl_nombre}_{i}")
            with c4: med4=st.number_input("Med 4 (lx)",min_value=0.0,step=1.0,
                value=float(ex.get("Med4",0)),key=f"m4_{pnombre}_{pl_nombre}_{i}")

            ca, cb, cc = st.columns(3)
            with ca:
                ti_v = ex.get("TipoIluminacion","Artificial")
                tipo_ilum = st.selectbox("Tipo iluminación", ILUM,
                    index=ILUM.index(ti_v) if ti_v in ILUM else 1,
                    key=f"ilum_{pnombre}_{pl_nombre}_{i}")
            with cb:
                tl_v = ex.get("TipoLampara","LED")
                tipo_lamp = st.selectbox("Tipo lámpara", LAMP,
                    index=LAMP.index(tl_v) if tl_v in LAMP else 0,
                    key=f"lamp_{pnombre}_{pl_nombre}_{i}")
            with cc:
                ul_v = ex.get("UbicacionLuminaria","Localizado")
                ubic_lum = st.selectbox("Ubicación luminaria", UBIC,
                    index=UBIC.index(ul_v) if ul_v in UBIC else 0,
                    key=f"ubic_{pnombre}_{pl_nombre}_{i}")

            # Foto del punto
            st.markdown("**📷 Foto del punto**")
            foto_bytes = cargar_foto_punto(pl_data, i+1)
            cf1, cf2 = st.columns([1,2])
            with cf1:
                if foto_bytes:
                    st.image(foto_bytes, caption=f"Foto {i+1}", width=140)
            with cf2:
                foto_up = st.file_uploader("Subir / cambiar foto",
                    type=["jpg","jpeg","png"],
                    key=f"foto_{pnombre}_{pl_nombre}_{i}")
                if foto_up:
                    pl_data["fotos"][i+1] = foto_up.read()
                    guardar_proyectos(st.session_state.proyectos)
                    st.success("✅ Foto guardada"); st.rerun()

            nota = st.text_area("Observaciones", height=60,
                value=ex.get("Nota",""), key=f"nota_{pnombre}_{pl_nombre}_{i}")

            # Recomendación manual o IA
            recom_guardada = ex.get("Recomendacion","")
            recom = st.text_area("Recomendaciones", height=80,
                value=recom_guardada,
                key=f"recom_{pnombre}_{pl_nombre}_{i}",
                help="Escribe tu recomendación o usa el botón IA para generarla automáticamente")

            if st.button(f"🤖 Generar recomendación con IA",
                         key=f"ia_recom_{pnombre}_{pl_nombre}_{i}"):
                with st.spinner("Generando recomendación..."):
                    # Solo este punto para recomendación individual
                    punto_temp = ex.copy() if ex else {}
                    punto_temp.update({
                        "Número": i+1, "TipoArea": tipo_area,
                        "Em_req": em_req, "Promedio": round((med1+med2+med3+med4)/4, 1) if all(v>0 for v in [med1,med2,med3,med4]) else ex.get("Promedio",0),
                        "Resultado": ex.get("Resultado","Sin medición"),
                        "TipoIluminacion": tipo_ilum, "TipoLampara": tipo_lamp,
                        "Nota": nota
                    })
                    recom_ia = generar_recomendaciones_ia([punto_temp])
                    st.session_state[f"recom_ia_{pnombre}_{pl_nombre}_{i}"] = recom_ia
                    st.rerun()

            # Mostrar recomendación IA generada
            if f"recom_ia_{pnombre}_{pl_nombre}_{i}" in st.session_state:
                recom_ia_texto = st.session_state[f"recom_ia_{pnombre}_{pl_nombre}_{i}"]
                st.markdown(f'<div class="recomendacion-box">🤖 <strong>Recomendación IA:</strong><br>{recom_ia_texto}</div>',
                            unsafe_allow_html=True)
                if st.button("✅ Usar esta recomendación",
                             key=f"usar_ia_{pnombre}_{pl_nombre}_{i}"):
                    # Guardar en el dato del punto
                    idx_ex = next((j for j,d in enumerate(pl_data["data"])
                                   if d["Número"]==i+1), None)
                    if idx_ex is not None:
                        pl_data["data"][idx_ex]["Recomendacion"] = recom_ia_texto
                    del st.session_state[f"recom_ia_{pnombre}_{pl_nombre}_{i}"]
                    guardar_proyectos(st.session_state.proyectos)
                    st.rerun()

            if all(v > 0 for v in [med1, med2, med3, med4]):
                promedio  = (med1+med2+med3+med4)/4
                conforme  = promedio >= em_req
                resultado = "✅ Conforme" if conforme else "❌ No conforme"
                color_res = "green" if conforme else "red"
                if conforme:
                    st.success(f"Promedio: **{round(promedio,1)} lx** → {resultado}")
                else:
                    st.error(f"Promedio: **{round(promedio,1)} lx** → {resultado} "
                             f"(requiere ≥ {em_req} lx)")

                entrada = {
                    "Número": i+1, "Coordenadas": f"({xn:.6f}, {yn:.6f})",
                    "TipoArea": tipo_area, "Em_req": em_req, "Uo_min": uo_min,
                    "Med1": med1,"Med2": med2,"Med3": med3,"Med4": med4,
                    "Promedio": round(promedio,1), "Resultado": resultado,
                    "Color": color_res, "TipoIluminacion": tipo_ilum,
                    "TipoLampara": tipo_lamp, "UbicacionLuminaria": ubic_lum,
                    "Nota": nota.strip(), "Recomendacion": recom.strip(),
                    "Foto": foto_bytes is not None
                }
                idx_ex = next((j for j,d in enumerate(pl_data["data"])
                               if d["Número"]==i+1), None)
                if idx_ex is not None: pl_data["data"][idx_ex] = entrada
                else:                  pl_data["data"].append(entrada)
                guardar_proyectos(st.session_state.proyectos)

    # ── Gráfica de conformidad del plano + botón IA general ──────────────────
    st.divider()
    if pl_data["data"]:
        col_graf, col_ia = st.columns([1,1])

        with col_graf:
            st.subheader("📊 Conformidad del Plano")
            graf = grafica_conformidad(pl_data["data"], pl_nombre)
            if graf:
                st.image(graf, width=300)

        with col_ia:
            st.subheader("🤖 Recomendaciones IA — Plano completo")
            if st.button("🤖 Generar recomendaciones para todo el plano",
                         key=f"ia_plano_{pnombre}_{pl_nombre}",
                         use_container_width=True):
                with st.spinner("Analizando todos los puntos con IA..."):
                    recom_general = generar_recomendaciones_ia(pl_data["data"])
                    st.session_state[f"recom_plano_{pnombre}_{pl_nombre}"] = recom_general
                    st.rerun()

            if f"recom_plano_{pnombre}_{pl_nombre}" in st.session_state:
                texto = st.session_state[f"recom_plano_{pnombre}_{pl_nombre}"]
                st.markdown(f'<div class="recomendacion-box">{texto}</div>',
                            unsafe_allow_html=True)

        st.subheader("📋 Tabla de Resultados")
        df   = pd.DataFrame(pl_data["data"])
        cols = ["Número","TipoArea","Em_req","Med1","Med2","Med3","Med4","Promedio","Resultado"]
        cex  = [c for c in cols if c in df.columns]
        st.dataframe(df[cex].rename(columns={
            "TipoArea":"Tipo de Área","Em_req":"Em req.(lx)"}),
            use_container_width=True)


# ============================================================================
# MAIN
# ============================================================================

def main():
    st.set_page_config(
        page_title="LuxOMeter PRO · RETILAP",
        page_icon="💡",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    aplicar_estilos()
    inicializar_session_state()

    pagina = st.session_state.pagina
    if   pagina == "inicio":          pagina_inicio()
    elif pagina == "nuevo_proyecto":  pagina_nuevo_proyecto()
    elif pagina == "editar_proyecto": pagina_editar_proyecto()
    elif pagina == "editar_plano":    pagina_editar_plano()


if __name__ == "__main__":
    main()
