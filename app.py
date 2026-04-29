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

def procesar_expediente_el_salvador(f_inc, f_mae):
    try:
        # 1️⃣ Carga del Maestro
        df_mae = pd.read_excel(f_mae)
        df_mae.columns = [str(c).strip() for c in df_mae.columns]

        # 2️⃣ Carga Inteligente del archivo de Incorporación (Buscador de Pestañas)
        hojas_inc = pd.read_excel(f_inc, sheet_name=None)
        df_inc_per = None
        df_inc_fam = None
        
        for nombre_hoja, df in hojas_inc.items():
            df.columns = [str(c).strip() for c in df.columns]
            if "Nro. Documento" in df.columns and "Primer apellido" in df.columns:
                df_inc_per = df
            if "Parentesco" in df.columns and "Número de Documento" in df.columns:
                df_inc_fam = df
                
        if df_inc_per is None or df_inc_fam is None:
            st.error("⚠️ No se encontraron las estructuras esperadas (Personas y Familiares) en el Excel de Incorporación.")
            return None

        # ========================================================
        # 🧾 PESTAÑA 1: DATOS GENERALES
        # ========================================================
        
        # Preparar Maestro
        cols_mae_gen = ["PAIS", "COD PERSONAL", "N° DOCUMENTO", "SEXO"]
        base_mae_gen = df_mae[[c for c in cols_mae_gen if c in df_mae.columns]].copy()
        
        # Preparar Incorporación (Personas)
        cols_inc_gen = ["Nro. Documento", "Primer nombre", "Segundo nombre", "Primer apellido", "Segundo apellido", 
                        "Apellido de casada", "Estado civil", "Fecha de nacimiento", "Nacionalidad", "País de nacimiento", 
                        "Profesión", "Correo personal", "Número de celular", "Teléfono fijo", "Departamento residencia", 
                        "Municipio residencia", "Dirección completa", "Banco", "Número de cta bancaria", 
                        "Sistema pensionario", "NUP", "Número de ISSS"]
        base_inc_per = df_inc_per[[c for c in cols_inc_gen if c in df_inc_per.columns]].copy()
        
        # Estandarización y Join
        base_mae_gen["N° DOCUMENTO"] = base_mae_gen["N° DOCUMENTO"].astype(str).str.strip()
        base_inc_per["Nro. Documento"] = base_inc_per["Nro. Documento"].astype(str).str.strip()
        df_gen = pd.merge(base_mae_gen, base_inc_per, left_on="N° DOCUMENTO", right_on="Nro. Documento", how="inner")
        
        # Renombrado de Columnas
        map_gen = {
            "PAIS": "Pais Empresa", "COD PERSONAL": "Código De Empleado", "Primer nombre": "Primer Nombre",
            "Segundo nombre": "Segundo Nombre", "Primer apellido": "Primer Apellido", "Segundo apellido": "Segundo Apellido",
            "Apellido de casada": "Apellido Casada", "SEXO": "Genero", "Estado civil": "EstadoCivil",
            "Nacionalidad": "Pais Nacionalidad", "Número de cta bancaria": "Cuenta Banco", "Dirección completa": "Dirección",
            "Departamento residencia": "Departamento Residencia", "Municipio residencia": "Municipio Residencia",
            "Teléfono fijo": "Telefono", "Número de celular": "Celular", "Correo personal": "eMail Interno",
            "Profesión": "Profesión (100 caracteres)", "Nro. Documento": "No Identificacion", "Sistema pensionario": "AFP",
            "País de nacimiento": "Pais Nacimiento", "Número de ISSS": "No Seguro Social"
        }
        df_gen = df_gen.rename(columns=map_gen)
        
        # Agregar columnas nuevas
        df_gen["Otros Nombres"] = ""
        df_gen["Tipo Cuenta Banco"] = "Ahorro"
        df_gen["No Identificacion Tributaria"] = ""
        df_gen["Digito Verificador"] = ""
        
        # Orden y limpieza final
        orden_gen = ["Pais Empresa", "Código De Empleado", "Primer Nombre", "Segundo Nombre", "Otros Nombres", "Primer Apellido", "Segundo Apellido", "Apellido Casada", "Genero", "EstadoCivil", "Pais Nacimiento", "Fecha Nacimiento", "Pais Nacionalidad", "Cuenta Banco", "Tipo Cuenta Banco", "Banco", "Dirección", "Departamento Residencia", "Municipio Residencia", "Telefono", "Celular", "eMail Interno", "Profesión (100 caracteres)", "No Identificacion", "No Seguro Social", "No Identificacion Tributaria", "AFP", "NUP", "Digito Verificador"]
        df_gen = df_gen.reindex(columns=orden_gen)

        # ========================================================
        # 👨‍👩‍👧 PESTAÑA 2: DEPENDIENTES
        # ========================================================
        
        cols_mae_dep = ["COD PERSONAL", "N° DOCUMENTO"]
        base_mae_dep = df_mae[[c for c in cols_mae_dep if c in df_mae.columns]].copy()
        
        # Búsqueda dinámica de la columna de dependencia económica
        col_dependencia = next((c for c in df_inc_fam.columns if "¿Cuéntas con familiares que dependen económicamente?" in str(c)), None)
        
        cols_inc_fam = ["Número de Documento", "Parentesco", "N° de doc. de derechohabiente", "Fecha de nacimiento", "Nombres completos", "Sexo", "Nivel educativo"]
        if col_dependencia: cols_inc_fam.append(col_dependencia)
            
        base_inc_fam = df_inc_fam[[c for c in cols_inc_fam if c in df_inc_fam.columns]].copy()
        
        # Estandarización y Join
        base_mae_dep["N° DOCUMENTO"] = base_mae_dep["N° DOCUMENTO"].astype(str).str.strip()
        base_inc_fam["Número de Documento"] = base_inc_fam["Número de Documento"].astype(str).str.strip()
        df_dep = pd.merge(base_mae_dep, base_inc_fam, left_on="N° DOCUMENTO", right_on="Número de Documento", how="inner")
        
        # Renombrado de Columnas
        map_dep = {
            "COD PERSONAL": "CodigoEmpleado", "Nombres completos": "nombre", "Sexo": "Género",
            "N° de doc. de derechohabiente": "No Documento", "Fecha de nacimiento": "FechaNacimiento", "Nivel educativo": "NivelEstudio"
        }
        if col_dependencia: map_dep[col_dependencia] = "DependenciaEconomica"
            
        df_dep = df_dep.rename(columns=map_dep)
        
        # Agregar columnas nuevas
        for col in ["EstadoCivil", "Trabaja", "LugarTrabajo", "Estudia", "LugarEstudio"]:
            df_dep[col] = ""
            
        # Orden final
        orden_dep = ["CodigoEmpleado", "Parentesco", "nombre", "Género", "EstadoCivil", "DependenciaEconomica", "FechaNacimiento", "No Documento", "Trabaja", "LugarTrabajo", "Estudia", "NivelEstudio", "LugarEstudio"]
        df_dep = df_dep.reindex(columns=orden_dep)

        # ========================================================
        # 📥 GENERACIÓN DEL EXCEL MULTI-HOJA
        # ========================================================
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_gen.to_excel(writer, index=False, sheet_name='Datos Generales')
            df_dep.to_excel(writer, index=False, sheet_name='Dependientes')
            
            # Autoajuste de celdas para ambas hojas
            for sheet_name in ['Datos Generales', 'Dependientes']:
                worksheet = writer.sheets[sheet_name]
                for col in worksheet.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
                        except: pass
                    worksheet.column_dimensions[column].width = (max_length + 2)
                    
        return output.getvalue()

    except Exception as e:
        st.error(f"Error crítico en la plantilla de Expediente (El Salvador): {e}")
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

# --- PASO 1: SELECCIÓN DE SOCIEDAD ---
st.markdown(f"<h3 style='color: {color_verde};'>1. Selección de Sociedad</h3>", unsafe_allow_html=True)
st.write("Selecciona la sociedad correspondiente para aplicar las reglas de transformación adecuadas.")

# Diccionario de mapeo: { "Nombre a mostrar en pantalla" : "País lógico para el código" }
mapeo_sociedades = {
    "OPERADORES LOGISTICOS RANSA S.A. DE C.V. (EL SALVADOR)": "El Salvador",
    "ALMACENES GENERALES DE DEPÓSITOS DE OCCIDENTE S.A. (AGDOSA)": "El Salvador",
    "OPERADORES LOGISTICOS RANSA S.A. (GUATEMALA)": "Guatemala",
    "OPERADORES LOGISTICOS RANSA S.A. DE C.V. (HONDURAS)": "Honduras"
}

sociedad_seleccionada = st.selectbox(
    "Sociedad de origen de los datos:",
    options=list(mapeo_sociedades.keys()), # Mostramos solo las llaves (nombres largos)
    index=None, 
    placeholder="Elige una sociedad de la lista..."
)

# --- PASO 2: CARGA DE ARCHIVOS (Visible solo tras seleccionar sociedad) ---
if sociedad_seleccionada:
    # "Traducimos" la sociedad al país correspondiente usando el diccionario
    pais_mapeado = mapeo_sociedades[sociedad_seleccionada]
    
    st.divider()
    st.markdown(f"<h3 style='color: {color_verde};'>2. Adjuntar Archivos Base</h3>", unsafe_allow_html=True)
    st.info(f"🏢 Sociedad activa: **{sociedad_seleccionada}** (Lógica aplicada: {pais_mapeado})")

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
        st.success("✅ Archivos cargados exitosamente. Procesando información...")
        
    
    # Aquí bifurcamos basándonos en el PAÍS MAPEADO, no en el nombre de la sociedad
        if pais_mapeado == "El Salvador":
            excel_empleos = procesar_el_salvador(file_incorporacion, file_maestro, file_requerimientos)
            # NUEVO: Reemplazamos el dummy por la función real
            excel_expediente = procesar_expediente_el_salvador(file_incorporacion, file_maestro)
            
        elif pais_mapeado == "Guatemala":
            excel_empleos = generar_excel_dummy(pais_mapeado, "Empleos")
            excel_expediente = generar_excel_dummy(pais_mapeado, "Expediente")
            
        elif pais_mapeado == "Honduras":
            excel_empleos = generar_excel_dummy(pais_mapeado, "Empleos")
            excel_expediente = generar_excel_dummy(pais_mapeado, "Expediente")
        
        # Renderizado estandarizado de los botones
        if excel_empleos and excel_expediente:
            col_down1, col_down2 = st.columns(2)
            
            with col_down1:
                st.download_button(
                    label="📥 Descargar Plantilla 'Empleos'",
                    data=excel_empleos,
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
    # Mensaje de ayuda mientras no se seleccione una sociedad
    st.info("👆 Por favor, selecciona una sociedad en el menú superior para habilitar la carga de archivos.")