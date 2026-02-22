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
                        
                        # Reconstruir imagen desde base64
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
                
                # Guardar imagen como base64
                if plano_info.get("img") is not None:
                    try:
                        img_bytes = io.BytesIO()
                        plano_info["img"].save(img_bytes, format="PNG")
                        img_bytes.seek(0)
                        plano_dict["img_base64"] = base64.b64encode(img_bytes.read()).decode()
                    except Exception as e:
                        pass
                
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
        # Mostrar proyectos en tarjetas
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
                    btn_col1, btn_col2, btn_col3 = st.columns(3)
                    
                    with btn_col1:
                        if st.button("✏️", key=f"btn_editar_{idx}_{proyecto_nombre}", help="Editar"):
                            st.session_state.proyecto_actual = proyecto_nombre
                            st.session_state.pagina = "editar_proyecto"
                            st.rerun()
                    
                    with btn_col2:
                        if st.button("📊", key=f"btn_csv_{idx}_{proyecto_nombre}", help="Descargar CSV"):
                            csv_data = generar_reporte_csv(proyecto_data, proyecto_nombre)
                            if csv_data:
                                st.download_button(
                                    label="Descargar CSV",
                                    data=csv_data,
                                    file_name=f"Reporte_{proyecto_nombre.replace(' ', '_')}.csv",
                                    mime="text/csv",
                                    key=f"download_csv_{idx}"
                                )
                    
                    with btn_col3:
                        if st.button("🗑️", key=f"btn_eliminar_{idx}_{proyecto_nombre}", help="Eliminar"):
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
    
    # Asegurar que existen todos los campos necesarios
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
    
    # Información del proyecto
    st.subheader("📋 Información del Proyecto")
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Orden:** {general.get('numero_orden', 'N/A')}")
        st.write(f"**Empresa:** {general.get('nombre_empresa', 'N/A')}")
    with col2:
        st.write(f"**Sede:** {general.get('sede', 'N/A')}")
        st.write(f"**Fecha:** {general.get('fecha', 'N/A')}")
    
    st.divider()
    
    # Sección de planos
    st.subheader("📐 Planos")
    
    # Subir nuevo plano
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
    
    # Listar planos existentes
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
    
    # Validar imagen
    if plano_img is None:
        st.error(f"⚠️ La imagen del plano no se pudo cargar. Por favor, vuelve a subir el plano.")
        return
    
    # Mostrar imagen
    st.image(plano_img, caption=f"Plano: {plano_actual} - Haz clic para marcar puntos", use_column_width=True)
    
    # Seleccionar tipo de área
    tipo_area = st.selectbox("Tipo de área según RETILAP", list(RETILAP_REFERENCIA.keys()), 
                            index=list(RETILAP_REFERENCIA.keys()).index(general["tipo_area"]), 
                            key=f"tipo_area_{proyecto_actual}_{plano_actual}")
    general["tipo_area"] = tipo_area
    
    valores = RETILAP_REFERENCIA[tipo_area]
    em_sugerido = valores["Em"]
    uo_min = valores["Uo"]
    
    st.success(f"💡 Iluminancia mantenida sugerida (Em): **{em_sugerido} lx** | Uniformidad mínima sugerida (Uo): **{uo_min}**")
    
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
            guardar_proyectos(st.session_state.proyectos)
            st.rerun()
    
    st.write(f"**Puntos en este plano:** {len(plano_data['puntos'])}")
    
    # Botones de gestión de puntos
    col1, col2 = st.columns(2)
    with col1:
        if st.button("🗑️ Eliminar último punto", key=f"eliminar_ultimo_{proyecto_actual}_{plano_actual}"):
            if plano_data["puntos"]:
                plano_data["puntos"].pop()
                guardar_proyectos(st.session_state.proyectos)
                st.rerun()
    with col2:
        if st.button("🧹 Limpiar todos los puntos", key=f"limpiar_puntos_{proyecto_actual}_{plano_actual}"):
            plano_data["puntos"] = []
            guardar_proyectos(st.session_state.proyectos)
            st.rerun()
    
    st.divider()
    
    # Ingreso de mediciones
    if plano_data["puntos"]:
        st.subheader("📊 Mediciones por punto")
        
        for i, (x, y) in enumerate(plano_data["puntos"]):
            with st.expander(f"Punto {i+1} ({int(x)}, {int(y)})", expanded=False):
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    med1 = st.number_input("Med 1", min_value=0.0, step=0.1, key=f"m1_{proyecto_actual}_{plano_actual}_{i}")
                with col2:
                    med2 = st.number_input("Med 2", min_value=0.0, step=0.1, key=f"m2_{proyecto_actual}_{plano_actual}_{i}")
                with col3:
                    med3 = st.number_input("Med 3", min_value=0.0, step=0.1, key=f"m3_{proyecto_actual}_{plano_actual}_{i}")
                with col4:
                    med4 = st.number_input("Med 4", min_value=0.0, step=0.1, key=f"m4_{proyecto_actual}_{plano_actual}_{i}")
                
                foto_subida = st.file_uploader(f"Foto del punto {i+1} (opcional)", type=["jpg", "jpeg", "png"], key=f"foto_{proyecto_actual}_{plano_actual}_{i}")
                if foto_subida is not None:
                    plano_data["fotos"][i+1] = foto_subida.read()
                    guardar_proyectos(st.session_state.proyectos)
                
                nota = st.text_area("Notas / observaciones", height=80, key=f"nota_{proyecto_actual}_{plano_actual}_{i}")
                
                # Actualizar datos
                if all(v > 0 for v in [med1, med2, med3, med4]):
                    promedio = (med1 + med2 + med3 + med4) / 4
                    conforme = promedio >= em_sugerido
                    color = "green" if conforme else "red"
                    resultado = "✅ Conforme" if conforme else "❌ No conforme"
                    
                    if len(plano_data["data"]) > i:
                        plano_data["data"][i] = {
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
                        }
                    else:
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
                    guardar_proyectos(st.session_state.proyectos)
        
        st.divider()
        
        # Mostrar mapa
        if plano_data["data"]:
            st.subheader("🗺️ Mapa de puntos")
            df_plano = pd.DataFrame(plano_data["data"])
            draw_img = plano_img.copy()
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
            
            # Tabla de resultados
            st.subheader("📊 Tabla de Resultados")
            st.dataframe(df_plano[["Número", "Coordenadas", "Med1", "Med2", "Med3", "Med4", "Promedio", "Resultado"]], use_container_width=True)


# ============================================================================
# FUNCIÓN PRINCIPAL
# ============================================================================

def main():
    """Función principal de la aplicación"""
    st.set_page_config(page_title="Auditoría Iluminación RETILAP", layout="wide")
    
    # Inicializar session state
    inicializar_session_state()
    
    # Mostrar página según el estado
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

