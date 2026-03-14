"""
generar_word.py  —  LuxOMeter PRO / RETILAP 2024
Tabla de luxometría con la estructura exacta del formato RETILAP.
"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
from datetime import datetime

AZ_OSC = RGBColor(0x1A,0x3A,0x5C)
AZ_CLA = RGBColor(0xD6,0xE4,0xF0)
VERDE  = RGBColor(0x27,0xAE,0x60)
ROJO   = RGBColor(0xE7,0x4C,0x3C)
BLANCO = RGBColor(0xFF,0xFF,0xFF)
NEGRO  = RGBColor(0x00,0x00,0x00)
GRIS   = RGBColor(0xF0,0xF4,0xF8)

def _bg(cell, color):
    tc=cell._tc; tcPr=tc.get_or_add_tcPr()
    shd=OxmlElement('w:shd')
    hx=f"{color[0]:02X}{color[1]:02X}{color[2]:02X}"
    shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),hx)
    tcPr.append(shd)

def _txt(cell, text, bold=False, sz=7, color=NEGRO,
         align=WD_ALIGN_PARAGRAPH.CENTER, italic=False):
    cell.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
    p=cell.paragraphs[0]; p.clear(); p.alignment=align
    r=p.add_run(str(text) if text is not None else "")
    r.bold=bold; r.italic=italic
    r.font.size=Pt(sz); r.font.color.rgb=color

def _borders(table):
    tbl=table._tbl; tblPr=tbl.tblPr
    b=OxmlElement('w:tblBorders')
    for s in('top','left','bottom','right','insideH','insideV'):
        e=OxmlElement(f'w:{s}')
        e.set(qn('w:val'),'single'); e.set(qn('w:sz'),'4')
        e.set(qn('w:space'),'0');    e.set(qn('w:color'),'AACCEE')
        b.append(e)
    tblPr.append(b)

def _img_buf(pil_img):
    buf=io.BytesIO(); pil_img.save(buf,format="PNG"); buf.seek(0); return buf

def generar_informe_word(general:dict, mediciones:list, plano_imgs:dict=None) -> bytes:
    doc=Document()
    sec=doc.sections[0]
    sec.page_width=Cm(35.56); sec.page_height=Cm(21.59)
    sec.left_margin=sec.right_margin=sec.top_margin=sec.bottom_margin=Cm(1.5)

    # Título
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r=p.add_run("ESTUDIO DE LUXOMETRÍA")
    r.bold=True; r.font.size=Pt(16); r.font.color.rgb=AZ_OSC

    p2=doc.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
    r2=p2.add_run("Auditoría de Iluminación en el Lugar de Trabajo – Norma RETILAP 2024")
    r2.font.size=Pt(9); r2.font.color.rgb=RGBColor(0x2C,0x6F,0xAD)
    doc.add_paragraph()

    # Ficha empresa
    ficha=doc.add_table(rows=3,cols=6); ficha.alignment=WD_TABLE_ALIGNMENT.CENTER; _borders(ficha)
    fd=[
        ["Empresa:",general.get("nombre_empresa",""),"NIT:",general.get("nit",""),"N° Orden:",general.get("numero_orden","")],
        ["Dirección:",general.get("direccion",""),"Ciudad:",general.get("sede",""),"Fecha:",general.get("fecha","")],
        ["Higienista:",general.get("responsable_higienista",""),"Lic. SST:",general.get("resolucion",""),"Responsable:",general.get("responsable_empresa","")],
    ]
    cw_f=[2.2,5.5,2.2,5.5,2.2,5.5]
    for ri,fila in enumerate(fd):
        bg=GRIS if ri%2==0 else BLANCO
        for ci,val in enumerate(fila):
            cell=ficha.cell(ri,ci); cell.width=Cm(cw_f[ci])
            _bg(cell,bg)
            is_lbl=(ci%2==0)
            _txt(cell,val,bold=is_lbl,sz=8,color=AZ_OSC if is_lbl else NEGRO,align=WD_ALIGN_PARAGRAPH.LEFT)
    doc.add_paragraph()

    # Resumen
    tot=len(mediciones); conf=sum(1 for m in mediciones if "✅" in str(m.get("resultado",""))); defic=tot-conf
    pct=round(conf/tot*100,1) if tot>0 else 0
    rs=doc.add_table(rows=2,cols=4); rs.alignment=WD_TABLE_ALIGNMENT.CENTER; _borders(rs)
    rh=["Total puntos","Adecuados","Deficientes","% Adecuados"]
    rv=[str(tot),str(conf),str(defic),f"{pct}%"]
    for ci,h in enumerate(rh):
        cell=rs.cell(0,ci); cell.width=Cm(4); _bg(cell,AZ_OSC); _txt(cell,h,bold=True,color=BLANCO,sz=9)
    for ci,v in enumerate(rv):
        cell=rs.cell(1,ci); cell.width=Cm(4); _bg(cell,GRIS)
        c=(VERDE if pct>=80 else ROJO) if ci==3 else NEGRO
        _txt(cell,v,bold=(ci==3),color=c,sz=9)
    doc.add_page_break()

    # Planos
    for pln_nombre,pln_img in (plano_imgs or {}).items():
        p_t=doc.add_paragraph(); rt=p_t.add_run(f"Plano: {pln_nombre}")
        rt.bold=True; rt.font.size=Pt(11); rt.font.color.rgb=AZ_OSC
        if pln_img:
            try: doc.add_picture(_img_buf(pln_img),width=Cm(32))
            except Exception as e: doc.add_paragraph(f"(Error imagen: {e})")
        doc.add_paragraph()

    # Tabla principal
    HEADS=[
        "N°\nMed","Puesto de trabajo\no Área evaluada","Ubicación",
        "Descripción",
        "E\nMIN\n(lx)","E\nMAX\n(lx)","E\nMEDIO\n(lx)",
        "Promedio\nmedido\n(lx)","Valor\nUo","Interp.\nUo",
        "Tipo de Área\nRETILAP","Em\nrec.\n(lx)",
        "Interpretación\ndel Nivel de\nIluminancia",
        "Observaciones /\nRecomendaciones",
    ]
    CW=[0.9,3.0,1.8,2.2,1.1,1.1,1.1,1.3,1.1,1.1,2.8,1.1,2.0,3.2]
    NC=len(HEADS)

    if mediciones:
        tbl=doc.add_table(rows=1,cols=NC); tbl.alignment=WD_TABLE_ALIGNMENT.CENTER; _borders(tbl)
        hr=tbl.rows[0]
        for ci,(h,cw) in enumerate(zip(HEADS,CW)):
            cell=hr.cells[ci]; cell.width=Cm(cw); _bg(cell,AZ_OSC)
            _txt(cell,h,bold=True,color=BLANCO,sz=6.5)

        for idx,m in enumerate(mediciones):
            conf_m="✅" in str(m.get("resultado",""))
            rbg=GRIS if idx%2==0 else BLANCO
            m1=m.get("med1",0) or 0; m2=m.get("med2",0) or 0
            m3=m.get("med3",0) or 0; m4=m.get("med4",0) or 0
            vals=[v for v in[m1,m2,m3,m4] if v>0]
            e_min  =m.get("e_min")  or (round(min(vals),1)         if vals else "")
            e_max  =m.get("e_max")  or (round(max(vals),1)         if vals else "")
            e_medio=m.get("e_medio")or (round(sum(vals)/len(vals),1) if vals else "")
            desc=(f"Tipo Ilum.: {m.get('tipo_iluminacion','')}\n"
                  f"Lámpara: {m.get('tipo_lampara','')}\n"
                  f"Ubic.: {m.get('ubicacion_luminaria','')}\n"
                  f"Ctrl. Luz Nat.: {m.get('control_luz_natural','')}\n"
                  f"Altura (m): {m.get('altura_luminaria','')}")
            obs=(f"Obs.: {m.get('nota','')}\n"
                 f"Rec.: {m.get('recomendacion','')}")
            vals_row=[
                str(m.get("num","")),
                str(m.get("puesto_evaluado","")) or str(m.get("area","")),
                str(m.get("ubicacion_luminaria","")),
                desc,
                str(e_min),str(e_max),str(e_medio),
                str(m.get("promedio","")),
                str(m.get("uo_calc","")),
                str(m.get("interpretacion_uo","")),
                str(m.get("area","")),
                str(m.get("em_req","")),
                "ADECUADO" if conf_m else "DEFICIENTE",
                obs,
            ]
            dr=tbl.add_row()
            for ci,(val,cw) in enumerate(zip(vals_row,CW)):
                cell=dr.cells[ci]; cell.width=Cm(cw)
                if ci==NC-1:  # Interpretación final
                    _bg(cell,VERDE if conf_m else ROJO)
                    _txt(cell,val,bold=True,color=BLANCO,sz=6.5)
                else:
                    _bg(cell,rbg)
                    al=(WD_ALIGN_PARAGRAPH.LEFT
                        if ci in(1,3,4,15) else WD_ALIGN_PARAGRAPH.CENTER)
                    _txt(cell,val,sz=6.5,align=al)

    # Pie
    doc.add_paragraph()
    pf=doc.add_paragraph(); pf.alignment=WD_ALIGN_PARAGRAPH.CENTER
    rf=pf.add_run(f"RETILAP 2024  ·  Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    rf.font.size=Pt(7); rf.font.color.rgb=RGBColor(0x80,0x80,0x80)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0); return buf.getvalue()
