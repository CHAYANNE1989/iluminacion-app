"""
generar_word.py  —  LuxOMeter PRO / RETILAP 2024
Genera el informe Word usando INFORME_PREFORMA.docx como plantilla base,
reemplazando los campos destacados en amarillo con los datos reales del proyecto.
"""
import io
import copy
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

# ── Ruta de la plantilla ─────────────────────────────────────────────────────
PLANTILLA_PATH = "INFORME_PREFORMA.docx"

# ── Colores ──────────────────────────────────────────────────────────────────
AZ_OSC = RGBColor(0x1A, 0x3A, 0x5C)
AZ_CLA = RGBColor(0xD6, 0xE4, 0xF0)
VERDE  = RGBColor(0x27, 0xAE, 0x60)
ROJO   = RGBColor(0xE7, 0x4C, 0x3C)
BLANCO = RGBColor(0xFF, 0xFF, 0xFF)
NEGRO  = RGBColor(0x00, 0x00, 0x00)
GRIS   = RGBColor(0xF0, 0xF4, 0xF8)
AMARILLO_RGB = RGBColor(0xFF, 0xC0, 0x00)


def _bg(cell, color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    hx = f"{color[0]:02X}{color[1]:02X}{color[2]:02X}"
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hx)
    tcPr.append(shd)


def _txt(cell, text, bold=False, sz=8, color=NEGRO,
         align=WD_ALIGN_PARAGRAPH.CENTER, italic=False):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.clear()
    p.alignment = align
    r = p.add_run(str(text) if text else "")
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(sz)
    r.font.color.rgb = color


def _borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    b = OxmlElement('w:tblBorders')
    for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        e = OxmlElement(f'w:{s}')
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), '4')
        e.set(qn('w:space'), '0')
        e.set(qn('w:color'), 'AACCEE')
        b.append(e)
    tblPr.append(b)


def _is_yellow(run):
    """Detecta si un run tiene resaltado amarillo."""
    try:
        return run.font.highlight_color is not None
    except:
        return False


def _reemplazar_run_amarillo(para, nuevo_texto):
    """Reemplaza el texto del primer run amarillo de un párrafo."""
    for run in para.runs:
        if _is_yellow(run) and run.text.strip():
            run.text = nuevo_texto
            run.font.highlight_color = None  # Quitar amarillo
            return True
    # Si todos los runs amarillos están vacíos, buscar el run con más texto
    for run in para.runs:
        if _is_yellow(run):
            run.text = nuevo_texto
            run.font.highlight_color = None
            return True
    return False


def _mes_texto(fecha_str):
    """Convierte 'dd/mm/yyyy' a 'MES - YYYY'."""
    meses = {
        '01':'ENERO','02':'FEBRERO','03':'MARZO','04':'ABRIL',
        '05':'MAYO','06':'JUNIO','07':'JULIO','08':'AGOSTO',
        '09':'SEPTIEMBRE','10':'OCTUBRE','11':'NOVIEMBRE','12':'DICIEMBRE'
    }
    try:
        partes = fecha_str.split('/')
        mes = meses.get(partes[1], partes[1])
        anio = partes[2]
        return f"{mes} - {anio}"
    except:
        return fecha_str


def _img_buf(pil_img):
    buf = io.BytesIO()
    pil_img.save(buf, format="PNG")
    buf.seek(0)
    return buf


def generar_informe_word(general: dict, mediciones: list,
                         plano_imgs: dict = None) -> bytes:
    """
    Genera informe Word sobre la plantilla INFORME_PREFORMA.docx.
    Reemplaza campos amarillos y agrega la tabla de resultados.
    """
    # ── Cargar plantilla ────────────────────────────────────────────────────
    if os.path.exists(PLANTILLA_PATH):
        doc = Document(PLANTILLA_PATH)
    else:
        # Si no existe la plantilla, generar documento básico
        return _generar_sin_plantilla(general, mediciones, plano_imgs)

    empresa    = general.get("nombre_empresa", "").upper()
    ciudad     = general.get("sede", "").upper()
    fecha_str  = general.get("fecha", datetime.now().strftime('%d/%m/%Y'))
    mes_anio   = _mes_texto(fecha_str).upper()
    higienista = general.get("responsable_higienista", "")
    resolucion = general.get("resolucion", "")
    nit        = general.get("nit", "")
    direccion  = general.get("direccion", "")
    orden      = general.get("numero_orden", "")

    # ── Reemplazar campos amarillos en párrafos ─────────────────────────────
    # Mapa: índice de párrafo → qué reemplazar
    # P6  → Nombre empresa (portada)
    # P7  → Ciudad
    # P30 → Ciudad + mes/año
    # P65 → Nombre empresa (objetivo)

    paras = doc.paragraphs

    for i, para in enumerate(paras):
        full = para.text.strip()

        # P6 — nombre empresa portada
        if i == 6 and any(_is_yellow(r) for r in para.runs):
            for run in para.runs:
                if _is_yellow(run):
                    if run.text.strip():
                        run.text = empresa
                        run.font.highlight_color = None
                    else:
                        run.text = ""
                        run.font.highlight_color = None

        # P7 — ciudad portada
        elif i == 7 and any(_is_yellow(r) for r in para.runs):
            for run in para.runs:
                if _is_yellow(run):
                    if run.text.strip():
                        run.text = ciudad
                        run.font.highlight_color = None
                    else:
                        run.text = ""
                        run.font.highlight_color = None

        # P30 — ciudad, mes/año (contraportada)
        elif i == 30 and any(_is_yellow(r) for r in para.runs):
            ciudad_puesto = True
            for run in para.runs:
                if _is_yellow(run):
                    if run.text.strip() not in ('', ',', ', '):
                        if ciudad_puesto:
                            run.text = ciudad
                            ciudad_puesto = False
                        else:
                            run.text = mes_anio
                        run.font.highlight_color = None
                    else:
                        run.font.highlight_color = None

        # P65 — nombre empresa en objetivo general
        elif i == 65 and any(_is_yellow(r) for r in para.runs):
            for run in para.runs:
                if _is_yellow(run):
                    if run.text.strip():
                        run.text = empresa
                        run.font.highlight_color = None
                    else:
                        run.text = ""
                        run.font.highlight_color = None

        # Reemplazar nombre empresa en cualquier párrafo que lo contenga
        elif 'INDEPENDIENTE SANTA FE' in para.text:
            for run in para.runs:
                if 'INDEPENDIENTE SANTA FE' in run.text:
                    run.text = run.text.replace('INDEPENDIENTE SANTA FE', empresa)
                if 'BOGOTA' in run.text and i not in (6, 7, 30):
                    run.text = run.text.replace('BOGOTA', ciudad)

        # Reemplazar en párrafos de análisis con números hardcoded
        elif 'seis (6)' in para.text or 'cinco (5)' in para.text:
            tot = len(mediciones)
            conf = sum(1 for m in mediciones if "✅" in str(m.get("resultado", "")))
            defic = tot - conf
            # Convertir números a texto
            nums = {0:'cero',1:'uno',2:'dos',3:'tres',4:'cuatro',5:'cinco',
                    6:'seis',7:'siete',8:'ocho',9:'nueve',10:'diez'}
            for run in para.runs:
                for n_orig, n_nuevo in [(6, tot), (5, defic), (1, conf)]:
                    txt_orig = f'{_num_txt(n_orig, nums)} ({n_orig})'
                    txt_nuevo = f'{_num_txt(n_nuevo, nums)} ({n_nuevo})'
                    if txt_orig in run.text:
                        run.text = run.text.replace(txt_orig, txt_nuevo)

    # ── Gráfica de conformidad ───────────────────────────────────────────────
    # Buscar párrafo P196 (Grafica 1) e insertar imagen debajo
    for i, para in enumerate(paras):
        if 'Grafica 1' in para.text or 'Conformidad' in para.text:
            # Actualizar texto con nombre real
            for run in para.runs:
                if 'INDEPENDIENTE SANTA FE' in run.text:
                    run.text = run.text.replace('INDEPENDIENTE SANTA FE', empresa)
                if _is_yellow(run):
                    run.text = empresa
                    run.font.highlight_color = None
            break

    # ── Actualizar Tabla 2 (niveles RETILAP) con áreas reales ──────────────
    # La tabla doc.tables[1] tiene las áreas — reemplazar con las del proyecto
    if len(doc.tables) > 1 and mediciones:
        _actualizar_tabla2_retilap(doc.tables[1], mediciones)

    # ── Insertar tabla de resultados en P189 (Tabla 4) ───────────────────────
    # Buscar exactamente el párrafo "Tabla 4." para insertar DESPUÉS de él
    idx_tabla4 = None
    for i, para in enumerate(paras):
        if para.text.strip().startswith('Tabla 4'):
            idx_tabla4 = i
            break

    if idx_tabla4 is not None and mediciones:
        _insertar_tabla_resultados(doc, paras[idx_tabla4], mediciones, general)

    # ── Insertar gráfica en P196 (Grafica 1) ────────────────────────────────
    for i, para in enumerate(paras):
        if para.text.strip().startswith('Grafica 1'):
            try:
                graf_bytes = _generar_grafica_bytes(mediciones)
                if graf_bytes:
                    # Insertar imagen justo después del párrafo Grafica 1
                    new_p = OxmlElement('w:p')
                    para._p.addnext(new_p)
                    # Crear párrafo centrado con la imagen
                    p_graf = doc.add_paragraph()
                    p_graf.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run_g = p_graf.add_run()
                    run_g.add_picture(io.BytesIO(graf_bytes), width=Cm(14))
                    # Mover ese párrafo justo después de Grafica 1
                    para._p.addnext(p_graf._p)
            except Exception as e:
                pass
            break

    # ── Insertar planos ──────────────────────────────────────────────────────
    if plano_imgs:
        for pln_nombre, pln_img in plano_imgs.items():
            if pln_img:
                try:
                    p = doc.add_paragraph()
                    run = p.add_run(f"Plano: {pln_nombre}")
                    run.bold = True
                    run.font.color.rgb = AZ_OSC
                    doc.add_picture(_img_buf(pln_img), width=Cm(22))
                    doc.add_paragraph()
                except Exception as e:
                    doc.add_paragraph(f"(Error al insertar plano: {e})")

    # ── Guardar y retornar ────────────────────────────────────────────────────
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def _num_txt(n, nums):
    return nums.get(n, str(n))


def _generar_grafica_bytes(mediciones):
    """Genera la gráfica de barras de conformidad y retorna bytes PNG."""
    try:
        import matplotlib
        matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        import io as _io

        total      = len(mediciones)
        conformes  = sum(1 for m in mediciones if "✅" in str(m.get("resultado", "")))
        deficientes = total - conformes
        if total == 0:
            return None

        pct_conf = round(conformes / total * 100, 1)
        pct_def  = round(deficientes / total * 100, 1)

        fig, ax = plt.subplots(figsize=(7, 3), facecolor='#f8fafc')
        ax.set_facecolor('#f8fafc')

        bars = ax.barh([1, 0], [pct_conf, pct_def],
                       color=['#27ae60', '#e74c3c'], height=0.5,
                       edgecolor='white', linewidth=1.5)
        for bar, val, n in zip(bars, [pct_conf, pct_def], [conformes, deficientes]):
            ax.text(val/2, bar.get_y() + bar.get_height()/2,
                    f"{val}%  ({n} pts)",
                    ha='center', va='center',
                    fontsize=11, fontweight='bold', color='white')

        ax.set_yticks([1, 0])
        ax.set_yticklabels(['Adecuados', 'Deficientes'],
                           fontsize=11, fontweight='bold', color='#1a3a5c')
        ax.set_xlim(0, 115)
        ax.set_xlabel('Porcentaje (%)', fontsize=9, color='#475569')
        ax.set_title(f'Conformidad Lumínica — {total} puntos evaluados',
                     fontsize=11, fontweight='bold', color='#1a3a5c', pad=10)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.tick_params(axis='x', colors='#94a3b8')
        ax.tick_params(axis='y', left=False)
        ax.xaxis.grid(True, linestyle='--', alpha=0.4, color='#cbd5e1')
        ax.set_axisbelow(True)
        plt.tight_layout()

        buf = _io.BytesIO()
        plt.savefig(buf, format='PNG', bbox_inches='tight', dpi=140, facecolor='#f8fafc')
        plt.close(fig)
        buf.seek(0)
        return buf.getvalue()
    except Exception as e:
        return None


def _actualizar_tabla2_retilap(tabla, mediciones):
    """
    Rellena la Tabla 2 de la plantilla con las áreas RETILAP
    realmente usadas en los puntos evaluados (sin duplicados).
    """
    from docx.oxml import OxmlElement as OE
    from docx.oxml.ns import qn

    # Obtener áreas únicas usadas, preservando orden
    areas_usadas = []
    vistas = set()
    for m in mediciones:
        area  = m.get("area", "")
        em    = m.get("em_req", "")
        uo    = m.get("uo_min", m.get("uo_calc", ""))
        clave = area
        if area and clave not in vistas:
            vistas.add(clave)
            areas_usadas.append((area, str(em), str(uo) if uo else ""))

    if not areas_usadas:
        return

    # Guardar el formato de la fila de encabezado (fila 0)
    # Eliminar todas las filas de datos existentes (fila 1 en adelante)
    tbl = tabla._tbl
    filas = tbl.findall(qn('w:tr'))
    for fila in filas[1:]:   # conservar encabezado (fila 0)
        tbl.remove(fila)

    # Agregar una fila por cada área única
    for area, em, uo in areas_usadas:
        # Copiar estructura de la fila de encabezado como base
        tr = OE('w:tr')
        for ci, val in enumerate([area, f"{em} lx" if em else "", uo]):
            tc = OE('w:tc')
            tcPr = OE('w:tcPr')
            # Ancho de celda igual al original
            try:
                orig_w = filas[0].findall(qn('w:tc'))[ci].find(qn('w:tcPr')).find(qn('w:tcW'))
                if orig_w is not None:
                    new_w = OE('w:tcW')
                    new_w.set(qn('w:w'), orig_w.get(qn('w:w')))
                    new_w.set(qn('w:type'), orig_w.get(qn('w:type')))
                    tcPr.append(new_w)
            except: pass
            tc.append(tcPr)
            p = OE('w:p')
            pPr = OE('w:pPr')
            jc = OE('w:jc')
            jc.set(qn('w:val'), 'center' if ci > 0 else 'left')
            pPr.append(jc)
            p.append(pPr)
            r = OE('w:r')
            rPr = OE('w:rPr')
            sz = OE('w:sz'); sz.set(qn('w:val'), '18')
            rPr.append(sz)
            r.append(rPr)
            t_elem = OE('w:t')
            t_elem.text = val
            r.append(t_elem)
            p.append(r)
            tc.append(p)
            tr.append(tc)
        tbl.append(tr)


def _insertar_tabla_resultados(doc, para_ref, mediciones, general):
    """Inserta la tabla de resultados después del párrafo de referencia."""
    HEADS = [
        "N°\nMed",
        "Puesto de trabajo\no Área evaluada",
        "Descripción",
        "E\nMIN\n(lx)",
        "E\nMAX\n(lx)",
        "Promedio\nmedido\n(lx)",
        "Valor\nUo",
        "Interp.\nUo",
        "Tipo de Área\nRETILAP",
        "Em\nrec.\n(lx)",
        "Interpretación\ndel Nivel de\nIluminancia",
        "Observaciones /\nRecomendaciones",
    ]
    CW = [0.9, 3.2, 2.8, 1.1, 1.1, 1.3, 1.1, 1.1, 3.2, 1.1, 2.0, 3.5]
    NC = len(HEADS)

    # Insertar párrafo título antes de la tabla
    p_tit = OxmlElement('w:p')
    para_ref._p.addnext(p_tit)

    tbl_elem = doc.add_table(rows=1, cols=NC)
    tbl_elem.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(tbl_elem)

    # Encabezado
    hr = tbl_elem.rows[0]
    for ci, (h, cw) in enumerate(zip(HEADS, CW)):
        cell = hr.cells[ci]
        cell.width = Cm(cw)
        _bg(cell, AZ_OSC)
        _txt(cell, h, bold=True, color=BLANCO, sz=7)

    # Filas de datos
    for idx_m, m in enumerate(mediciones):
        conf_m = "✅" in str(m.get("resultado", ""))
        rbg = GRIS if idx_m % 2 == 0 else BLANCO

        m1 = m.get("med1", 0) or 0
        m2 = m.get("med2", 0) or 0
        m3 = m.get("med3", 0) or 0
        m4 = m.get("med4", 0) or 0
        vals = [v for v in [m1, m2, m3, m4] if v > 0]
        e_min   = m.get("e_min")   or (round(min(vals), 1)         if vals else "")
        e_max   = m.get("e_max")   or (round(max(vals), 1)         if vals else "")

        desc = (f"Tipo Ilum.: {m.get('tipo_iluminacion','')}\n"
                f"Lámpara: {m.get('tipo_lampara','')}\n"
                f"Ubic.: {m.get('ubicacion_luminaria','')}\n"
                f"Ctrl. Luz Nat.: {m.get('control_luz_natural','')}\n"
                f"Altura (m): {m.get('altura_luminaria','')}")
        obs = (f"Obs.: {m.get('nota','')}\n"
               f"Rec.: {m.get('recomendacion','')}")

        vals_row = [
            str(m.get("num", "")),
            str(m.get("puesto_evaluado", "") or m.get("area", "")),
            desc,
            str(e_min), str(e_max),
            str(m.get("promedio", "")),
            str(m.get("uo_calc", "")),
            str(m.get("interpretacion_uo", "")),
            str(m.get("area", "")),
            str(m.get("em_req", "")),
            "ADECUADO" if conf_m else "DEFICIENTE",
            obs,
        ]

        dr = tbl_elem.add_row()
        for ci, (val, cw) in enumerate(zip(vals_row, CW)):
            cell = dr.cells[ci]
            cell.width = Cm(cw)
            if ci == NC - 2:  # Interpretación
                _bg(cell, VERDE if conf_m else ROJO)
                _txt(cell, val, bold=True, color=BLANCO, sz=7)
            else:
                _bg(cell, rbg)
                al = (WD_ALIGN_PARAGRAPH.LEFT
                      if ci in (1, 2, 8, NC - 1) else WD_ALIGN_PARAGRAPH.CENTER)
                _txt(cell, val, sz=7, align=al)

    # Mover tabla al lugar correcto (después del párrafo de referencia)
    para_ref._p.addnext(tbl_elem._tbl)


def _generar_sin_plantilla(general, mediciones, plano_imgs):
    """Fallback: genera Word básico si no existe la plantilla."""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width  = Cm(35.56)
    sec.page_height = Cm(21.59)
    sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = Cm(1.5)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ESTUDIO DE LUXOMETRÍA – RETILAP 2024")
    r.bold = True; r.font.size = Pt(16); r.font.color.rgb = AZ_OSC

    empresa = general.get("nombre_empresa", "")
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(empresa)
    r2.bold = True; r2.font.size = Pt(13)
    doc.add_paragraph()

    # Tabla básica de resultados
    HEADS = ["N° Med","Puesto/Área","Descripción","E Min","E Max",
             "Promedio","Uo","Interp. Uo","Tipo Área","Em rec.","Resultado","Obs/Rec"]
    CW = [0.9,3.2,2.8,1.1,1.1,1.3,1.1,1.1,3.2,1.1,2.0,3.5]
    NC = len(HEADS)

    tbl = doc.add_table(rows=1, cols=NC)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _borders(tbl)
    hr = tbl.rows[0]
    for ci,(h,cw) in enumerate(zip(HEADS,CW)):
        cell=hr.cells[ci]; cell.width=Cm(cw)
        _bg(cell,AZ_OSC); _txt(cell,h,bold=True,color=BLANCO,sz=7)

    for idx_m,m in enumerate(mediciones):
        conf_m="✅" in str(m.get("resultado",""))
        rbg=GRIS if idx_m%2==0 else BLANCO
        m1=m.get("med1",0) or 0; m2=m.get("med2",0) or 0
        m3=m.get("med3",0) or 0; m4=m.get("med4",0) or 0
        vals=[v for v in[m1,m2,m3,m4] if v>0]
        e_min=m.get("e_min") or (round(min(vals),1) if vals else "")
        e_max=m.get("e_max") or (round(max(vals),1) if vals else "")
        desc=(f"Tipo: {m.get('tipo_iluminacion','')} | Lámpara: {m.get('tipo_lampara','')}\n"
              f"Ubic.: {m.get('ubicacion_luminaria','')} | Alt.: {m.get('altura_luminaria','')}")
        obs=f"Obs.: {m.get('nota','')} Rec.: {m.get('recomendacion','')}"
        vr=[str(m.get("num","")),str(m.get("puesto_evaluado","") or m.get("area","")),
            desc,str(e_min),str(e_max),str(m.get("promedio","")),
            str(m.get("uo_calc","")),str(m.get("interpretacion_uo","")),
            str(m.get("area","")),str(m.get("em_req","")),
            "ADECUADO" if conf_m else "DEFICIENTE",obs]
        dr=tbl.add_row()
        for ci,(val,cw) in enumerate(zip(vr,CW)):
            cell=dr.cells[ci]; cell.width=Cm(cw)
            if ci==NC-2:
                _bg(cell,VERDE if conf_m else ROJO)
                _txt(cell,val,bold=True,color=BLANCO,sz=7)
            else:
                _bg(cell,rbg)
                al=WD_ALIGN_PARAGRAPH.LEFT if ci in(1,2,8,NC-1) else WD_ALIGN_PARAGRAPH.CENTER
                _txt(cell,val,sz=7,align=al)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.getvalue()
