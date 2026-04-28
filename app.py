import streamlit as st
import pandas as pd
import io

st.image("Logo.png", width=200)

# 1. CONFIGURACIÓN DE LA PÁGINA
st.set_page_config(
    page_title="Transformación de Datos HR",
    page_icon="🔄",
    layout="wide"
)

color_verde = "#009A3F"
color_naranja = "#F39200"
color_gris = "#9D9D9C"

# =========================================================
# FUNCIONES MODULARES (Preparación para futura lógica)
# =========================================================

# Estas funciones albergarán la lógica de limpieza y transformación por país.
# Por ahora usarán el generador dummy, pero la estructura ya está lista.
def procesar_el_salvador(f_inc, f_mae, f_req):
    try:
        # --- CARGA DE DATOS POR PESTAÑA ---
        # Leemos las dos pestañas específicas del archivo de requerimientos (f_req)
        df_org = pd.read_excel(f_req, sheet_name="Datos organizativos")
        df_cand = pd.read_excel(f_req, sheet_name="Candidatos seleccionados")
        # El maestro tiene una sola pestaña
        df_mae = pd.read_excel(f_mae)

        # Limpieza de nombres de columnas (quitar espacios invisibles)
        for df in [df_org, df_cand, df_mae]:
            df.columns = [str(c).strip() for c in df.columns]

        # 1️⃣ Requerimiento – Datos Organizativos
        # Mantenemos solo las columnas solicitadas
        cols_mantener_org = ["RQ", "PLANILLA", "TIPO DE CONTRATO", "MOTIVO", "BONO ANUAL"]
        req_org = df_org[[c for c in cols_mantener_org if c in df_org.columns]].copy()

        # 2️⃣ Requerimiento – Candidatos seleccionados
        # Mantenemos solo las columnas solicitadas y eliminamos duplicados por RQ
        cols_mantener_cand = ["RQ", "ERF - BÁSICO", "DNI"]
        req_cand = df_cand[[c for c in cols_mantener_cand if c in df_cand.columns]].copy()
        req_cand = req_cand.drop_duplicates(subset=["RQ"])

        # Estandarización de tipo de dato para el primer JOIN (RQ)
        req_org["RQ"] = req_org["RQ"].astype(str).str.strip()
        req_cand["RQ"] = req_cand["RQ"].astype(str).str.strip()

        # 3️⃣ Combinación Base_Requerimiento
        base_req = pd.merge(req_org, req_cand, on="RQ", how="inner")

        # 4️⃣ Preparación del maestro de personal
        cols_mantener_mae = [
            "COD PERSONAL", "COD. POSICION", "ESTADO", "FECHA DE BAJA", 
            "FECHA FIN DE CONTRATO", "FECHA DE INICIO DE CONTRATO", 
            "FECHA DE INGRESO AL GRUPO", "N° DOCUMENTO"
        ]
        base_mae = df_mae[[c for c in cols_mantener_mae if c in df_mae.columns]].copy()

        # Estandarización de tipo de dato para el segundo JOIN (DNI / N° DOCUMENTO)
        base_mae["N° DOCUMENTO"] = base_mae["N° DOCUMENTO"].astype(str).str.strip()
        base_req["DNI"] = base_req["DNI"].astype(str).str.strip()

        # 5️⃣ Generación final de la plantilla “Empleos”
        df_final = pd.merge(base_mae, base_req, left_on="N° DOCUMENTO", right_on="DNI", how="inner")

        # --- SECCIÓN DE CAMPOS NUEVOS ---
        # Estos campos se inicializan vacíos según requerimiento
        df_final["Gastos de Representacion"] = ""
        df_final["Numero Horas por Mes"] = ""
        df_final["Jornada"] = ""
        # --------------------------------

        # 6️⃣ Reordenamiento y Renombre de columnas
        mapeo_nombres = {
            "COD PERSONAL": "Código De Empleado",
            "COD. POSICION": "Codigo Plaza",
            "ESTADO": "Estado",
            "PLANILLA": "Tipo de Planilla",
            "FECHA DE INGRESO AL GRUPO": "Fecha Ingreso",
            "FECHA DE BAJA": "Fecha de Retiro",
            "MOTIVO": "Motivo de Retiro",
            "TIPO DE CONTRATO": "Tipo de Contrato",
            "ERF - BÁSICO": "Salario Mensual",
            "FECHA DE INICIO DE CONTRATO": "Fecha Inicio Contrato",
            "FECHA FIN DE CONTRATO": "Fecha Fin Contrato",
            "BONO ANUAL": "Bonificacion Decreto"
        }
        
        df_final = df_final.rename(columns=mapeo_nombres)

        orden_final = [
            "Código De Empleado", "Codigo Plaza", "Estado", "Tipo de Planilla", "Fecha Ingreso",
            "Fecha de Retiro", "Motivo de Retiro", "Tipo de Contrato", "Jornada", "Salario Mensual",
            "Fecha Inicio Contrato", "Fecha Fin Contrato", "Bonificacion Decreto", 
            "Gastos de Representacion", "Numero Horas por Mes", "N° DOCUMENTO"
        ]
        
        # Reindexamos para asegurar el orden y creación de columnas si alguna faltase
        df_final = df_final.reindex(columns=orden_final)

        # 7️⃣ Limpieza final
        df_final = df_final.drop(columns=["N° DOCUMENTO"])

        # --- GENERACIÓN DE EXCEL CON AUTOAJUSTE DE COLUMNAS ---
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Empleos')
            
            # Lógica de autoajuste de ancho de celdas
            worksheet = writer.sheets['Empleos']
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column].width = adjusted_width
        
        return output.getvalue()

    except Exception as e:
        st.error(f"Error en la lógica de El Salvador (Pestañas): {e}")
        return None
    

def procesar_guatemala(f_inc, f_mae, f_req):
    pass

def procesar_honduras(f_inc, f_mae, f_req):
    pass

# Generador temporal de Excel para pruebas de botones
def generar_excel_dummy(pais_origen, tipo_plantilla):
    df_dummy = pd.DataFrame({
        "País": [pais_origen],
        "Plantilla": [tipo_plantilla],
        "Estado": ["En construcción"],
        "Mensaje": ["Aquí se insertará la data transformada posteriormente."]
    })
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_dummy.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()


# =========================================================
# INTERFAZ DE USUARIO (FRONTEND)
# =========================================================

st.markdown(f"<h1 style='text-align: center; color: {color_naranja};'>Automatización de Plantillas de Personal</h1>", unsafe_allow_html=True)
st.markdown(f"<h4 style='text-align: center; color: {color_gris};'>Generador estandarizado de Empleos y Expediente</h4>", unsafe_allow_html=True)
st.divider()

# --- PASO 1: SELECCIÓN DE PAÍS ---
st.markdown(f"<h3 style='color: {color_verde};'>1. Selección de País</h3>", unsafe_allow_html=True)
st.write("Selecciona el país correspondiente para aplicar las reglas de transformación adecuadas.")

pais_seleccionado = st.selectbox(
    "País de origen de los datos:",
    options=["El Salvador", "Guatemala", "Honduras"],
    index=None, # Muestra el selectbox vacío por defecto
    placeholder="Elige un país de la lista..."
)

# --- PASO 2: CARGA DE ARCHIVOS (Visible solo tras seleccionar país) ---
if pais_seleccionado:
    st.divider()
    st.markdown(f"<h3 style='color: {color_verde};'>2. Adjuntar Archivos Base - {pais_seleccionado}</h3>", unsafe_allow_html=True)
    st.write("Asegúrate de cargar los archivos con los prefijos correctos.")

    col_up1, col_up2, col_up3 = st.columns(3)

    with col_up1:
        st.markdown("**Incorporación de personal**")
        st.caption("Prefijo: `Datos_Web_...`")
        file_incorporacion = st.file_uploader("Subir archivo 1", type=["xlsx", "xls"], key="up_inc")

    with col_up2:
        st.markdown("**Maestro de personal**")
        st.caption("Prefijo: `Maestro_de_personal_...`")
        file_maestro = st.file_uploader("Subir archivo 2", type=["xlsx", "xls"], key="up_mae")

    with col_up3:
        st.markdown("**Requerimientos de personal**")
        st.caption("Prefijo: `Datos_Web_rq_...`")
        file_requerimientos = st.file_uploader("Subir archivo 3", type=["xlsx", "xls"], key="up_req")

    # --- PASO 3: LÓGICA DE RUTEO Y DESCARGA ---
    if file_incorporacion and file_maestro and file_requerimientos:
        st.divider()
        st.markdown(f"<h3 style='color: {color_verde};'>3. Resultados y Descarga</h3>", unsafe_allow_html=True)
        st.success(f"✅ Archivos cargados exitosamente. Motor de transformación listo para {pais_seleccionado}.")
        
    # Aquí es donde el código se bifurca según el país seleccionado
        if pais_seleccionado == "El Salvador":
            # 1. Obtenemos el Excel real procesado para Empleos
            excel_empleos = procesar_el_salvador(file_incorporacion, file_maestro, file_requerimientos)
            # 2. Expediente sigue usando el archivo temporal "dummy" por ahora
            excel_expediente = generar_excel_dummy(pais_seleccionado, "Expediente")
            
        elif pais_seleccionado == "Guatemala":
            excel_empleos = generar_excel_dummy(pais_seleccionado, "Empleos")
            excel_expediente = generar_excel_dummy(pais_seleccionado, "Expediente")
            
        elif pais_seleccionado == "Honduras":
            excel_empleos = generar_excel_dummy(pais_seleccionado, "Empleos")
            excel_expediente = generar_excel_dummy(pais_seleccionado, "Expediente")
        
        # Renderizado estandarizado de los botones 
        # (Solo se muestran si el Excel se generó correctamente, evitando errores si algo falla en la lógica)
        if excel_empleos and excel_expediente:
            col_down1, col_down2 = st.columns(2)
            
            with col_down1:
                st.download_button(
                    label="📥 Descargar Plantilla 'Empleos'",
                    data=excel_empleos, # Aquí le pasamos la data real
                    file_name=f"Empleos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            with col_down2:
                st.download_button(
                    label="📥 Descargar Plantilla 'Expediente'",
                    data=excel_expediente,
                    file_name=f"Expediente.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
else:
    # Mensaje de ayuda mientras no se seleccione un país
    st.info("👆 Por favor, selecciona un país en el menú superior para habilitar la carga de archivos.")