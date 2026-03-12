"""
Generador de informe Word RETILAP
Toma datos directamente de la app (sin Excel externo).
"""
import io
from datetime import datetime
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import base64
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ─── HELPERS ─────────────────────────────────────────────────────────────────

def _set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def _celda(cell, text, size=9, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, bg=None, color=None):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    run = p.add_run(str(text))
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = 'Arial'
    if color:
        run.font.color.rgb = RGBColor(*bytes.fromhex(color))
    if bg:
        _set_cell_bg(cell, bg)

def _bordes(table):
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcB = OxmlElement('w:tcBorders')
            for side in ['top','left','bottom','right']:
                b = OxmlElement(f'w:{side}')
                b.set(qn('w:val'), 'single')
                b.set(qn('w:sz'), '4')
                b.set(qn('w:color'), 'AAAAAA')
                tcB.append(b)
            tcPr.append(tcB)

# ─── GRÁFICA ─────────────────────────────────────────────────────────────────

def generar_grafica(mediciones):
    labels = []
    for m in mediciones:
        area = m.get('area', '')
        labels.append(f"P{m['num']}\n{area[:12]}..." if len(area)>12 else f"P{m['num']}\n{area}")

    promedios = [m['promedio'] or 0 for m in mediciones]
    em_reqs   = [m['em_req'] or 0 for m in mediciones]
    colores   = ['#27AE60' if 'Conforme' in m.get('resultado','') else '#E74C3C' for m in mediciones]

    fig, ax = plt.subplots(figsize=(max(10, len(mediciones)*1.8), 5))
    x = np.arange(len(labels))
    w = 0.35

    bars1 = ax.bar(x - w/2, promedios, w, color=colores, label='Promedio Medido (lx)', edgecolor='white')
    bars2 = ax.bar(x + w/2, em_reqs,   w, color='#2C6FAD', alpha=0.7, label='Em Requerida (lx)', edgecolor='white')

    for bar in bars1:
        h = bar.get_height()
        ax.annotate(f'{h:.0f}', xy=(bar.get_x()+bar.get_width()/2, h), xytext=(0,3),
                    textcoords='offset points', ha='center', va='bottom', fontsize=7)
    for bar in bars2:
        h = bar.get_height()
        ax.annotate(f'{h:.0f}', xy=(bar.get_x()+bar.get_width()/2, h), xytext=(0,3),
                    textcoords='offset points', ha='center', va='bottom', fontsize=7)

    ax.set_xlabel('Puntos de Medición', fontsize=10)
    ax.set_ylabel('Iluminancia (lx)', fontsize=10)
    ax.set_title('Niveles de Iluminancia Medidos vs Requeridos', fontsize=12, fontweight='bold')
    ax.set_xticks(x); ax.set_xticklabels(labels, fontsize=7)
    verde = mpatches.Patch(color='#27AE60', label='Adecuado')
    rojo  = mpatches.Patch(color='#E74C3C', label='Deficiente')
    azul  = mpatches.Patch(color='#2C6FAD', alpha=0.7, label='Em Requerida')
    ax.legend(handles=[verde, rojo, azul], fontsize=8)
    ax.grid(axis='y', alpha=0.3)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    buf = io.BytesIO()
    plt.tight_layout()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    plt.close(); buf.seek(0)
    return buf

# ─── GENERADOR PRINCIPAL ──────────────────────────────────────────────────────

def generar_informe_word(general, mediciones, plano_imgs=None):
    """
    general: dict con datos del proyecto (de st.session_state)
    mediciones: lista de dicts con datos de cada punto
    plano_imgs: dict {nombre: PIL.Image con puntos dibujados}
    """
    doc = Document()

    # Márgenes
    for section in doc.sections:
        section.top_margin    = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin   = Cm(2.5)
        section.right_margin  = Cm(2.5)

    doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(11)

    # ── PORTADA ──────────────────────────────────────────────────────────────
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('INFORME DE EVALUACIONES OCUPACIONALES\nNIVELES DE ILUMINACIÓN')
    run.bold = True; run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)

    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run('EMPRESA').bold = True

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(general.get('nombre_empresa', ''))
    run.bold = True; run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"NIT: {general.get('nit', '')}").font.size = Pt(12)

    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run('Elaborado Por').font.size = Pt(11)

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{general.get('responsable_higienista','')}  |  Licencia en SST Res. No. {general.get('resolucion','')}")
    run.bold = True; run.font.size = Pt(11)

    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(general.get('mes_anio', '')).font.size = Pt(12)

    doc.add_page_break()

    # ── DATOS DE LA EMPRESA ───────────────────────────────────────────────────
    h = doc.add_heading('DATOS DE LA EMPRESA', level=1)
    h.runs[0].font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)

    campos_emp = [
        ('NIT',                              general.get('nit', '')),
        ('RAZÓN SOCIAL',                     general.get('nombre_empresa', '')),
        ('FECHA Y HORA DE ACTIVIDAD',        general.get('fecha', '')),
        ('DIRECCIÓN',                        general.get('direccion', '')),
        ('TELÉFONO',                         general.get('telefono', '')),
        ('RESPONSABLE DE ATENDER LA ASESORÍA', general.get('responsable_empresa', '')),
        ('CIUDAD DE EJECUCIÓN',              general.get('sede', '')),
        ('ORDEN DE TRABAJO N°',              general.get('numero_orden', '')),
    ]
    t_emp = doc.add_table(rows=len(campos_emp), cols=2)
    t_emp.style = 'Table Grid'
    for i, (lbl, val) in enumerate(campos_emp):
        _celda(t_emp.rows[i].cells[0], lbl, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, bg='D5E8F0')
        _celda(t_emp.rows[i].cells[1], val, align=WD_ALIGN_PARAGRAPH.LEFT)
    _bordes(t_emp)
    doc.add_paragraph()

    # ── INTRODUCCIÓN ─────────────────────────────────────────────────────────
    doc.add_page_break()
    h = doc.add_heading('INTRODUCCIÓN', level=1)
    h.runs[0].font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)
    intro = (
        f"La iluminación industrial es uno de los factores ambientales primordiales que contribuye a la correcta "
        f"y adecuada realización de las tareas, facilitando la visualización de las cosas dentro de su contexto "
        f"espacial, de modo que el trabajo se pueda realizar en unas condiciones aceptables de comodidad, seguridad "
        f"y eficacia. La empresa {general.get('nombre_empresa','')} ubicada en {general.get('sede','')} solicitó "
        f"la evaluación de los niveles de iluminación en sus instalaciones, con el fin de verificar el cumplimiento "
        f"de los requisitos establecidos en el Reglamento Técnico de Iluminación y Alumbrado Público – RETILAP – "
        f"Resolución 40150 del 03 de mayo de 2024."
    )
    p = doc.add_paragraph(intro); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # ── TABLA DE RESULTADOS ───────────────────────────────────────────────────
    doc.add_page_break()
    h = doc.add_heading('RESULTADOS DE MEDICIÓN', level=1)
    h.runs[0].font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)
    p = doc.add_paragraph('Tabla 2. NIVELES DE ILUMINANCIA Y UNIFORMIDAD OBTENIDOS')
    p.runs[0].bold = True; p.runs[0].font.size = Pt(10)

    hdrs = ['N°', 'Área / Puesto', 'Tipo Ilum.', 'Med1\n(lx)', 'Med2\n(lx)', 'Med3\n(lx)', 'Med4\n(lx)',
            'Prom.\n(lx)', 'Em Req.\n(lx)', 'Resultado']
    anch = [0.8, 3.5, 2, 1.3, 1.3, 1.3, 1.3, 1.4, 1.4, 2.2]

    t_res = doc.add_table(rows=1+len(mediciones), cols=len(hdrs))
    t_res.style = 'Table Grid'
    for j, (h_txt, a) in enumerate(zip(hdrs, anch)):
        c = t_res.rows[0].cells[j]
        c.width = Cm(a)
        _celda(c, h_txt, bold=True, color='FFFFFF', bg='1A569A', size=8)

    for i, m in enumerate(mediciones):
        row = t_res.rows[i+1]
        es_def = 'Conforme' not in m.get('resultado', '')
        bg_r = 'FDDEDE' if es_def else 'D5F5E3'
        vals = [
            str(m['num']),
            m.get('area', ''),
            m.get('tipo_iluminacion', ''),
            str(m.get('med1', '')),
            str(m.get('med2', '')),
            str(m.get('med3', '')),
            str(m.get('med4', '')),
            str(m.get('promedio', '')),
            str(m.get('em_req', '')),
            '✓ ADECUADO' if not es_def else '✗ DEFICIENTE',
        ]
        bgs = [None]*9 + [bg_r]
        for j, (v, bg) in enumerate(zip(vals, bgs)):
            _celda(row.cells[j], v, size=8, bg=bg)
    _bordes(t_res)
    doc.add_paragraph()

    # ── DETALLE POR PUNTO ─────────────────────────────────────────────────────
    h = doc.add_heading('DETALLE POR PUNTO DE MEDICIÓN', level=2)
    h.runs[0].font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)

    for m in mediciones:
        p = doc.add_paragraph()
        p.add_run(f"Punto {m['num']} – {m.get('area', '')}").bold = True

        t_obs = doc.add_table(rows=3, cols=2)
        t_obs.style = 'Table Grid'
        obs_data = [
            ('Tipo de Iluminación', m.get('tipo_iluminacion', '')),
            ('Tipo de Lámpara',     m.get('tipo_lampara', '')),
            ('Ubicación Luminaria', m.get('ubicacion_luminaria', '')),
        ]
        for k, (lbl, val) in enumerate(obs_data):
            _celda(t_obs.rows[k].cells[0], lbl, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, bg='EBF5FB', size=8)
            _celda(t_obs.rows[k].cells[1], val, align=WD_ALIGN_PARAGRAPH.LEFT, size=8)
        _bordes(t_obs)

        if m.get('nota'):
            p2 = doc.add_paragraph()
            p2.add_run('Observaciones: ').bold = True
            p2.add_run(m['nota'])
            p2.runs[0].font.size = Pt(9)
        if m.get('recomendacion'):
            p3 = doc.add_paragraph()
            p3.add_run('Recomendaciones: ').bold = True
            p3.add_run(m['recomendacion'])
            p3.runs[0].font.size = Pt(9)
        doc.add_paragraph()

    # ── GRÁFICA ───────────────────────────────────────────────────────────────
    if mediciones:
        doc.add_page_break()
        h = doc.add_heading('ANÁLISIS GRÁFICO DE RESULTADOS', level=1)
        h.runs[0].font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)
        graf = generar_grafica(mediciones)
        doc.add_picture(graf, width=Cm(16))
        doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ── RESUMEN CONFORMIDAD ───────────────────────────────────────────────────
    doc.add_paragraph()
    h = doc.add_heading('RESUMEN DE CONFORMIDAD', level=2)
    h.runs[0].font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)

    total     = len(mediciones)
    adecuados = sum(1 for m in mediciones if 'Conforme' in m.get('resultado',''))
    defic     = total - adecuados
    pct       = (adecuados/total*100) if total > 0 else 0

    t_sum = doc.add_table(rows=4, cols=2)
    t_sum.style = 'Table Grid'
    sum_data = [
        ('Total puntos evaluados', str(total)),
        ('Puntos ADECUADOS',       f"{adecuados} ({pct:.1f}%)"),
        ('Puntos DEFICIENTES',     f"{defic} ({100-pct:.1f}%)"),
        ('Conformidad General',    'CONFORME' if pct >= 80 else 'NO CONFORME'),
    ]
    bgs_s = ['EBF5FB','D5F5E3','FDDEDE','D5F5E3' if pct>=80 else 'FDDEDE']
    for k, ((lbl, val), bg) in enumerate(zip(sum_data, bgs_s)):
        _celda(t_sum.rows[k].cells[0], lbl, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, bg='D5E8F0', size=10)
        _celda(t_sum.rows[k].cells[1], val, bold=True, size=10, bg=bg)
    _bordes(t_sum)

    # ── PLANOS CON PUNTOS ─────────────────────────────────────────────────────
    if plano_imgs:
        doc.add_page_break()
        h = doc.add_heading('PLANOS CON PUNTOS DE MEDICIÓN', level=1)
        h.runs[0].font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)
        for nombre_plano, img_pil in plano_imgs.items():
            p = doc.add_paragraph(f'Plano: {nombre_plano}')
            p.runs[0].bold = True
            buf_img = io.BytesIO()
            img_pil.save(buf_img, format='PNG'); buf_img.seek(0)
            doc.add_picture(buf_img, width=Cm(16))
            doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph()

    # ── CONCLUSIONES ──────────────────────────────────────────────────────────
    doc.add_page_break()
    h = doc.add_heading('ANÁLISIS DE RESULTADOS Y CONCLUSIONES', level=1)
    h.runs[0].font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)
    p = doc.add_paragraph('[Complete aquí las conclusiones del informe]')
    p.runs[0].font.color.rgb = RGBColor(0x99,0x99,0x99)
    p.runs[0].font.italic = True

    doc.add_paragraph()
    h = doc.add_heading('RECOMENDACIONES', level=1)
    h.runs[0].font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)
    p = doc.add_paragraph('[Complete aquí las recomendaciones generales]')
    p.runs[0].font.color.rgb = RGBColor(0x99,0x99,0x99)
    p.runs[0].font.italic = True

    # ── BIBLIOGRAFÍA ──────────────────────────────────────────────────────────
    doc.add_page_break()
    h = doc.add_heading('BIBLIOGRAFÍA', level=1)
    h.runs[0].font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)
    for b in [
        'Reglamento Técnico de Iluminación y Alumbrado Público – RETILAP. Ministerio de Minas y Energía, Colombia, 2010.',
        'Resolución 40150 del 03 de mayo de 2024. Ministerio de Minas y Energía.',
        'Guía Técnica Colombiana GTC-8: 1994. ICONTEC.',
        'NTC GTC 8 de 1994. Principios de Ergonomía Visual.',
    ]:
        p = doc.add_paragraph(b, style='List Bullet')
        p.runs[0].font.size = Pt(10)

    buf = io.BytesIO()
    doc.save(buf); buf.seek(0)
    return buf
