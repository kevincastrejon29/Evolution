import streamlit as st
import pandas as pd
import io

#col_izq, col_centro, col_der = st.columns([1.5, 1, 1.5,])
#with col_centro:
st.image("Logo.png", width=200)

# 1. CONFIGURACIÓN DE LA PÁGINA
st.set_page_config(
    page_title="ETL-Evolution",
    page_icon="Icono.png",
    layout="wide"
)

# --- BARRA LATERAL (SIDEBAR): CENTRO DE AYUDA ---
with st.sidebar:
    st.markdown("<h2 style='color: #009A3F;'>💡 Centro de Ayuda</h2>", unsafe_allow_html=True)
    st.write("Bienvenido al asistente de automatización HR. Haz clic en las preguntas para resolver tus dudas:")
    
    with st.expander("🤔 ¿Para qué sirve esta plataforma?"):
        st.write("Esta herramienta automatiza el cruce de datos entre el Maestro de Personal y las hojas de Requerimientos e Incorporación, generando las plantillas 'Empleos' y 'Expediente' en segundos y sin errores manuales.")
        
    with st.expander("🛠️ ¿Cómo debo utilizarla?"):
        st.markdown("""
        1. Selecciona tu **Sociedad** en el panel principal.
        2. Arrastra los **3 archivos Excel** originales.
        3. Haz clic en **Descargar** para obtener tus plantillas formateadas.
        """)
        
    with st.expander("⚠️ ¿Qué hago si me sale un error?"):
        st.write("Asegúrate de que los archivos estén en formato Excel (.xlsx o .xls) y que no se hayan modificado los nombres de las pestañas originales ('Datos organizativos', 'Candidatos seleccionados', etc.).")

    with st.expander("🔒 ¿Qué pasa con los datos que subo?"):
        st.write("Toda la información se procesa de forma temporal en la memoria. No se guarda ningún archivo ni dato sensible en servidores externos. Una vez que descargas tu Excel o cierras la página, los datos desaparecen.")

    st.divider()
    st.caption("Desarrollado para la mejora de procesos. v1.0")

color_verde = "#009A3F"
color_naranja = "#F39200"
color_gris = "#9D9D9C"

# Estas funciones albergarán la lógica de limpieza y transformación por país.
# Por ahora usarán el generador dummy, pero la estructura ya está lista.fdef
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

        # 1️ Requerimiento – Datos Organizativos
        # Mantenemos solo las columnas solicitadas
        cols_mantener_org = ["RQ", "CLASE DE CONTRATO", "MOTIVO", "BONO ANUAL"]
        req_org = df_org[[c for c in cols_mantener_org if c in df_org.columns]].copy()

        # 2️ Requerimiento – Candidatos seleccionados
        # Mantenemos solo las columnas solicitadas y eliminamos duplicados por RQ
        cols_mantener_cand = ["RQ", "ERF - BÁSICO", "DNI"]
        req_cand = df_cand[[c for c in cols_mantener_cand if c in df_cand.columns]].copy()
        req_cand = req_cand.drop_duplicates(subset=["RQ"])

        # Estandarización de tipo de dato para el primer JOIN (RQ)
        req_org["RQ"] = req_org["RQ"].astype(str).str.strip()
        req_cand["RQ"] = req_cand["RQ"].astype(str).str.strip()

        # 3️ Combinación Base_Requerimiento
        base_req = pd.merge(req_org, req_cand, on="RQ", how="inner")

        # 4️ Preparación del maestro de personal
        cols_mantener_mae = [
            "COD PERSONAL", "COD POSICION", "ESTADO", "FECHA DE BAJA", 
            "FECHA FIN CONTRATO", "FECHA INICIO CONTRATO", 
            "FECHA INGRESO AL GRUPO", "N° DOCUMENTO", "CLASE CONTRATO"
        ]
        base_mae = df_mae[[c for c in cols_mantener_mae if c in df_mae.columns]].copy()

        # Estandarización de tipo de dato para el segundo JOIN (DNI / N° DOCUMENTO)
        base_mae["N° DOCUMENTO"] = base_mae["N° DOCUMENTO"].astype(str).str.strip()
        base_req["DNI"] = base_req["DNI"].astype(str).str.strip()

        # 5️ Generación final de la plantilla “Empleos”
        df_final = pd.merge(base_mae, base_req, left_on="N° DOCUMENTO", right_on="DNI", how="inner")

        # --- SECCIÓN DE CAMPOS NUEVOS ---
        # Estos campos se inicializan vacíos según requerimiento
        df_final["Gastos de Representacion"] = ""
        df_final["Numero Horas por Mes"] = ""
        df_final["Jornada"] = ""
        # --------------------------------

        # 6️ Reordenamiento y Renombre de columnas
        mapeo_nombres = {
            "COD PERSONAL": "Código De Empleado",
            "COD POSICION": "Codigo Plaza",
            "ESTADO": "Estado",
            "CLASE DE CONTRATO": "Tipo de Planilla",
            "FECHA INGRESO AL GRUPO": "Fecha Ingreso",
            "FECHA DE BAJA": "Fecha de Retiro",
            "MOTIVO": "Motivo de Retiro",
            "CLASE CONTRATO": "Tipo de Contrato",
            "ERF - BÁSICO": "Salario Mensual",
            "FECHA INICIO CONTRATO": "Fecha Inicio Contrato",
            "FECHA FIN CONTRATO": "Fecha Fin Contrato",
            "BONO ANUAL": "Bonificacion Decreto"
        }
        
        df_final = df_final.rename(columns=mapeo_nombres)

        orden_final = [
            "Código De Empleado", "Codigo Plaza", "Estado", "Tipo de Planilla", "Fecha Ingreso",
            "Fecha de Retiro", "Motivo de Retiro", "Tipo de Contrato", "Jornada", "Salario Mensual",
            "Fecha Inicio Contrato", "Fecha Fin Contrato", "Bonificacion Decreto", 
            "Gastos de Representacion", "Numero Horas por Mes", "N° DOCUMENTO"
        ]
        
        #--------------
        #--- AJUSTES FINALES: ESTADO ---
        if "Estado" in df_final.columns:
            # Reemplazamos los valores exactos
            map_estado = {"ACTIVO": "Activo", "CESADO": "Retirado"}
            df_final["Estado"] = df_final["Estado"].replace(map_estado)
        #--------------

        # Reindexamos para asegurar el orden y creación de columnas si alguna faltase
        df_final = df_final.reindex(columns=orden_final)

        # 7️ Limpieza final
        df_final = df_final.drop(columns=["N° DOCUMENTO"])

        # --- GENERACIÓN DE EXCEL CON AUTOAJUSTE DE COLUMNAS ---
        
        # Limpieza Global: Reemplazar guiones "-" por vacío
        df_final = df_final.replace("-", "", regex=False)

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
            "PAIS": "Pais Empresa", 
            "COD PERSONAL": "Código De Empleado", 
            "Primer nombre": "Primer Nombre",
            "Segundo nombre": "Segundo Nombre", "Primer apellido": "Primer Apellido", "Segundo apellido": "Segundo Apellido",
            "Apellido de casada": "Apellido Casada", "SEXO": "Genero", "Estado civil": "EstadoCivil",
            "Nacionalidad": "Pais Nacionalidad", "Número de cta bancaria": "Cuenta Banco", "Dirección completa": "Dirección",
            "Departamento residencia": "Departamento Residencia", "Municipio residencia": "Municipio Residencia",
            "Teléfono fijo": "Telefono", "Número de celular": "Celular", "Correo personal": "eMail Interno",
            "Profesión": "Profesión (100 caracteres)", "Nro. Documento": "No Identificacion", "Sistema pensionario": "AFP",
            "País de nacimiento": "Pais Nacimiento", "Número de ISSS": "No Seguro Social"
        }
        df_gen = df_gen.rename(columns=map_gen)
        # Formateo de Fecha de Nacimiento (DD/MM/YYYY) para Datos Generales
        if "Fecha Nacimiento" in df_gen.columns:
            df_gen["Fecha Nacimiento"] = pd.to_datetime(df_gen["Fecha Nacimiento"], errors='coerce').dt.strftime('%d/%m/%Y').fillna("")
        
        # Agregar columnas nuevas
        df_gen["Otros Nombres"] = ""
        df_gen["Tipo Cuenta Banco"] = "Ahorro"
        df_gen["No Identificacion Tributaria"] = ""
        df_gen["Digito Verificador"] = ""
        
        #-----------------
        # --- MAPEOS GLOBALES: DATOS GENERALES ---
        
        # 1. Mapeo de Género
        if "Genero" in df_gen.columns:
            map_genero = {"MASCULINO": "M", "FEMENINO": "F"}
            df_gen["Genero"] = df_gen["Genero"].str.upper().map(map_genero).fillna(df_gen["Genero"])

        # 2. Mapeo de Países (Nacimiento y Nacionalidad) - CORREGIDO
        map_paises = {
            "el salvador": "sv", 
            "guatemala": "gt", 
            "nicaragua": "ni",
            "honduras": "hn", 
            "panama": "pa", 
            "costa rica": "cr", 
            "estados unidos": "us"
        }
        for col in ["Pais Nacimiento", "Pais Nacionalidad"]:
            if col in df_gen.columns:
                # Convertimos lo que viene del Excel a minúsculas y sin espacios para que haga "match" perfecto
                df_gen[col] = df_gen[col].astype(str).str.lower().str.strip().map(map_paises).fillna(df_gen[col])

        # 3. Mapeo de Tipo de Cuenta
        if "Tipo Cuenta Banco" in df_gen.columns:
            map_cuentas = {"Ahorro": "A", "Cuenta Corriente": "C"}
            df_gen["Tipo Cuenta Banco"] = df_gen["Tipo Cuenta Banco"].map(map_cuentas).fillna(df_gen["Tipo Cuenta Banco"])
        
        # 4. Mapeo de Estado Civil (Datos Generales)
        if "EstadoCivil" in df_gen.columns:
            map_estado_civil = {
                "SOLTERO(A)": "S",
                "CASADO(A)": "C",
                "UNION LIBRE": "A",
                "DIVORCIADO(A)": "D",
                "VIUDO(A)": "V"
            }
            # Convertimos a mayúsculas y quitamos espacios para asegurar el match
            df_gen["EstadoCivil"] = df_gen["EstadoCivil"].str.upper().str.strip().map(map_estado_civil).fillna(df_gen["EstadoCivil"])
        #-----------------
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

        #---------------
        # --- MAPEOS GLOBALES: DEPENDIENTES ---
        
        # 1. Mapeo de Género (ya lo tenías)
        if "Género" in df_dep.columns:
            map_genero_dep = {"MASCULINO": "M", "FEMENINO": "F"}
            df_dep["Género"] = df_dep["Género"].str.upper().map(map_genero_dep).fillna(df_dep["Género"])

        # 2. Mapeo de Estado Civil (Nuevo)
        if "EstadoCivil" in df_dep.columns:
            map_civil_dep = {
                "SOLTERO(A)": "S",
                "CASADO(A)": "C",
                "UNION LIBRE": "A",
                "DIVORCIADO(A)": "D",
                "VIUDO(A)": "V"
            }
            df_dep["EstadoCivil"] = df_dep["EstadoCivil"].str.upper().str.strip().map(map_civil_dep).fillna(df_dep["EstadoCivil"])
        # -------------- 

        # Formateo de Fecha de Nacimiento (DD/MM/YYYY) para Dependientes
        if "FechaNacimiento" in df_dep.columns:
            df_dep["FechaNacimiento"] = pd.to_datetime(df_dep["FechaNacimiento"], errors='coerce').dt.strftime('%d/%m/%Y').fillna("")
        
        # Agregar columnas nuevas
        for col in ["EstadoCivil", "Trabaja", "LugarTrabajo", "Estudia", "LugarEstudio"]:
            df_dep[col] = ""
            
        # Orden final
        orden_dep = ["CodigoEmpleado", "Parentesco", "nombre", "Género", "EstadoCivil", "DependenciaEconomica", "FechaNacimiento", "No Documento", "Trabaja", "LugarTrabajo", "Estudia", "NivelEstudio", "LugarEstudio"]
        df_dep = df_dep.reindex(columns=orden_dep)

        # ========================================================
        # 📥 GENERACIÓN DEL EXCEL MULTI-HOJA
        # ========================================================
        
        # Limpieza Global: Reemplazar guiones "-" por vacío
        df_gen = df_gen.replace("-", "", regex=False)
        df_dep = df_dep.replace("-", "", regex=False)

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
    try:
        # --- CARGA DE DATOS POR PESTAÑA ---
        df_org = pd.read_excel(f_req, sheet_name="Datos organizativos")
        df_cand = pd.read_excel(f_req, sheet_name="Candidatos seleccionados")
        df_mae = pd.read_excel(f_mae)

        # Limpieza de nombres de columnas
        for df in [df_org, df_cand, df_mae]:
            df.columns = [str(c).strip() for c in df.columns]

        # 1️⃣ Requerimiento – Datos Organizativos
        cols_mantener_org = ["RQ", "CLASE DE CONTRATO", "MOTIVO", "BONO ANUAL"]
        req_org = df_org[[c for c in cols_mantener_org if c in df_org.columns]].copy()

        # 2️⃣ Requerimiento – Candidatos seleccionados
        cols_mantener_cand = ["RQ", "ERF - BÁSICO", "DNI"]
        req_cand = df_cand[[c for c in cols_mantener_cand if c in df_cand.columns]].copy()
        req_cand = req_cand.drop_duplicates(subset=["RQ"])

        # Estandarización de tipos para el primer JOIN (RQ)
        req_org["RQ"] = req_org["RQ"].astype(str).str.strip()
        req_cand["RQ"] = req_cand["RQ"].astype(str).str.strip()

        # 3️⃣ Combinación Base_Requerimiento
        base_req = pd.merge(req_org, req_cand, on="RQ", how="inner")

        # 4️⃣ Preparación del maestro de personal
        cols_mantener_mae = [
            "COD PERSONAL", "COD POSICION", "ESTADO", "FECHA DE BAJA", 
            "FECHA FIN CONTRATO", "FECHA INICIO CONTRATO", 
            "FECHA INGRESO AL GRUPO", "N° DOCUMENTO", "CLASE CONTRATO"
        ]
        base_mae = df_mae[[c for c in cols_mantener_mae if c in df_mae.columns]].copy()

        # Estandarización de tipos para el segundo JOIN
        base_mae["N° DOCUMENTO"] = base_mae["N° DOCUMENTO"].astype(str).str.strip()
        base_req["DNI"] = base_req["DNI"].astype(str).str.strip()

        # 5️⃣ Generación final de la plantilla “Empleos”
        df_final = pd.merge(base_mae, base_req, left_on="N° DOCUMENTO", right_on="DNI", how="inner")

        # --- SECCIÓN DE CAMPOS NUEVOS (GUATEMALA) ---
        # Se agregan vacíos según requerimiento para identificación posterior
        df_final["Gastos de Representacion"] = ""
        df_final["Numero Horas por Mes"] = ""
        
        # Lógica de autocompletado para Jornada
        # Como las columnas anteriores se crean en este paso, asignamos el valor predeterminado
        df_final["Jornada"] = "Jornada Administrativa"
        # --------------------------------------------

        # 6️⃣ Reordenamiento y Renombre de columnas
        mapeo_nombres = {
            "COD PERSONAL": "Código De Empleado",
            "COD POSICION": "Codigo Plaza",
            "ESTADO": "Estado",
            "CLASE DE CONTRATO": "Tipo de Planilla",
            "FECHA INGRESO AL GRUPO": "Fecha Ingreso",
            "FECHA DE BAJA": "Fecha de Retiro",
            "MOTIVO": "Motivo de Retiro",
            "CLASE CONTRATO": "Tipo de Contrato",
            "ERF - BÁSICO": "Salario Mensual",
            "FECHA INICIO CONTRATO": "Fecha Inicio Contrato",
            "FECHA FIN CONTRATO": "Fecha Fin Contrato",
            "BONO ANUAL": "Bonificacion Decreto"
        }
        
        df_final = df_final.rename(columns=mapeo_nombres)

        orden_final = [
            "Código De Empleado", "Codigo Plaza", "Estado", "Tipo de Planilla", "Fecha Ingreso",
            "Fecha de Retiro", "Motivo de Retiro", "Tipo de Contrato", "Jornada", "Salario Mensual",
            "Fecha Inicio Contrato", "Fecha Fin Contrato", "Bonificacion Decreto", 
            "Gastos de Representacion", "Numero Horas por Mes", "N° DOCUMENTO"
        ]
        
        #-----------------
        # --- AJUSTES FINALES: ESTADO ---
        if "Estado" in df_final.columns:
            # Reemplazamos los valores exactos
            map_estado = {"ACTIVO": "Activo", "CESADO": "Retirado"}
            df_final["Estado"] = df_final["Estado"].replace(map_estado)
        #-----------------

        df_final = df_final.reindex(columns=orden_final)

        # 7️⃣ Limpieza final
        df_final = df_final.drop(columns=["N° DOCUMENTO"])

        # Generación de Excel con autoajuste

        # Limpieza Global: Reemplazar guiones "-" por vacío
        df_final = df_final.replace("-", "", regex=False)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Empleos')
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
                worksheet.column_dimensions[column].width = (max_length + 2)
        
        return output.getvalue()

    except Exception as e:
        st.error(f"Error en la lógica de Guatemala: {e}")
        return None

def procesar_expediente_guatemala(f_inc, f_mae):
    try:
        # 1️ Carga del Maestro
        df_mae = pd.read_excel(f_mae)
        df_mae.columns = [str(c).strip() for c in df_mae.columns]

        # 2️ Carga Inteligente del archivo de Incorporación (Buscador de Pestañas)
        hojas_inc = pd.read_excel(f_inc, sheet_name=None)
        df_inc_per = None
        df_inc_fam = None
        
        for nombre_hoja, df in hojas_inc.items():
            df.columns = [str(c).strip() for c in df.columns]
            # Identificamos la pestaña "Personas" buscando columnas únicas de Guatemala
            if "Nro. Documento" in df.columns and "Apellido paterno" in df.columns:
                df_inc_per = df
            # Identificamos la pestaña "Beneficiarios"
            if "Vinculo" in df.columns and "Número de Documento" in df.columns:
                df_inc_fam = df
                
        if df_inc_per is None or df_inc_fam is None:
            st.error("⚠️ No se encontraron las estructuras esperadas en el Excel de Incorporación de Guatemala.")
            return None

        # ========================================================
        # PESTAÑA 1: DATOS GENERALES
        # ========================================================
        
        # Preparar Maestro
        cols_mae_gen = ["PAIS", "COD PERSONAL", "N° DOCUMENTO", "SEXO"]
        base_mae_gen = df_mae[[c for c in cols_mae_gen if c in df_mae.columns]].copy()
        
        # Preparar Incorporación (Personas)
        cols_inc_gen = [
            "Nro. Documento", "Nombres completos", "Apellido paterno", "Apellido materno", 
            "Estado civil", "Fecha de nacimiento", "Nacionalidad", "País de nacimiento", 
            "Correo personal", "Número de celular", "Teléfono fijo", "Cuenta con cta bancaria", 
            "Tipo de banco para cta bancaria", "Numero de cta bancaria", "N° Documento IGSS", 
            "N° Documento NIT", "Nombre de carrera", "Departamento residencia", "Municipio residencia", "Dirección","Tiene IGSS"
        ]
        base_inc_per = df_inc_per[[c for c in cols_inc_gen if c in df_inc_per.columns]].copy()
        
        # Estandarización y Join
        base_mae_gen["N° DOCUMENTO"] = base_mae_gen["N° DOCUMENTO"].astype(str).str.strip()
        base_inc_per["Nro. Documento"] = base_inc_per["Nro. Documento"].astype(str).str.strip()
        df_gen = pd.merge(base_mae_gen, base_inc_per, left_on="N° DOCUMENTO", right_on="Nro. Documento", how="inner")
        
        #Separación de la columna "Nombres completos"
        if "Nombres completos" in df_gen.columns:
            # n=2 permite separar hasta 3 partes (Primer, Segundo y el resto)
            split_names = df_gen["Nombres completos"].astype(str).str.split(" ", n=2, expand=True)
            
            df_gen["Primer Nombre"] = split_names[0] if 0 in split_names.columns else ""
            df_gen["Segundo Nombre"] = split_names[1] if 1 in split_names.columns else ""
            # Si existe una tercera parte, se asigna a Otros Nombres; si no, queda vacío
            df_gen["Otros Nombres"] = split_names[2] if 2 in split_names.columns else ""
        else:
            df_gen["Primer Nombre"] = ""
            df_gen["Segundo Nombre"] = ""
            df_gen["Otros Nombres"] = ""
            
        # Renombrado de Columnas
        map_gen = {
            "PAIS": "Pais Empresa", "COD PERSONAL": "Código De Empleado", "SEXO": "Genero",
            "Apellido paterno": "Primer Apellido", "Apellido materno": "Segundo Apellido",
            "Estado civil": "EstadoCivil", "País de nacimiento": "Pais Nacimiento",
            "Fecha de nacimiento": "Fecha Nacimiento", "Nacionalidad": "Pais Nacionalidad",
            "Numero de cta bancaria": "Cuenta Banco", "Departamento residencia": "Departamento Residencia",
            "Municipio residencia": "Municipio Residencia", "Teléfono fijo": "Telefono",
            "Número de celular": "Celular", "Correo personal": "eMail Interno",
            "Nombre de carrera": "Profesión (100 caracteres)", 
            "Nro. Documento": "No Identificacion",          # <-- CAMBIO 1
            "N° Documento IGSS": "No Seguro Social",        # <-- CAMBIO 2
            "N° Documento NIT": "No Identificacion Tributaria"
        }
        # Lógica condicional para Banco y Tipo Cuenta Banco (Guatemala)
        # 1. Por defecto, Banco viene de "Tipo de banco para cta bancaria"
        df_gen["Banco"] = df_gen["Tipo de banco para cta bancaria"]
        
        # 2. Identificamos registros donde no hay cuenta (basado en la columna original)
        mask_no_cuenta = df_gen["Cuenta con cta bancaria"].astype(str).str.upper() == "NO"
        
        # 3. Aplicamos condiciones: si es NO, limpiar. Si no, poner "Ahorro"
        df_gen.loc[mask_no_cuenta, "Banco"] = ""
        df_gen.loc[mask_no_cuenta, "Tipo Cuenta Banco"] = ""
        df_gen.loc[~mask_no_cuenta, "Tipo Cuenta Banco"] = "Ahorro"

        df_gen = df_gen.rename(columns=map_gen)
        # Formateo de Fecha de Nacimiento (DD/MM/YYYY) para Datos Generales
        if "Fecha Nacimiento" in df_gen.columns:
            df_gen["Fecha Nacimiento"] = pd.to_datetime(df_gen["Fecha Nacimiento"], errors='coerce').dt.strftime('%d/%m/%Y').fillna("")
        
        # Agregar columnas nuevas
        #df_gen["Otros Nombres"] = ""
        df_gen["Apellido Casada"] = ""
        #df_gen["No Seguro Social"] = ""
        #df_gen["AFP"] = ""
        df_gen["NUP"] = ""
        df_gen["Digito Verificador"] = ""

#------------------------------
#------------------------------
        
        # --- MAPEOS GLOBALES: DATOS GENERALES ---
        
        # 1. Mapeo de Género
        if "Genero" in df_gen.columns:
            map_genero = {"MASCULINO": "M", "FEMENINO": "F"}
            df_gen["Genero"] = df_gen["Genero"].str.upper().map(map_genero).fillna(df_gen["Genero"])

        # 2. Mapeo de Países (Nacimiento y Nacionalidad) - CORREGIDO
        map_paises = {
            "el salvador": "sv", 
            "guatemala": "gt", 
            "nicaragua": "ni",
            "honduras": "hn", 
            "panama": "pa", 
            "costa rica": "cr", 
            "estados unidos": "us"
        }
        for col in ["Pais Nacimiento", "Pais Nacionalidad"]:
            if col in df_gen.columns:
                # Convertimos lo que viene del Excel a minúsculas y sin espacios para que haga "match" perfecto
                df_gen[col] = df_gen[col].astype(str).str.lower().str.strip().map(map_paises).fillna(df_gen[col])

        # 3. Mapeo de Tipo de Cuenta
        if "Tipo Cuenta Banco" in df_gen.columns:
            map_cuentas = {"Ahorro": "A", "Cuenta Corriente": "C"}
            df_gen["Tipo Cuenta Banco"] = df_gen["Tipo Cuenta Banco"].map(map_cuentas).fillna(df_gen["Tipo Cuenta Banco"])
        # 4. Mapeo de Estado Civil (Datos Generales)
        if "EstadoCivil" in df_gen.columns:
            map_estado_civil = {
                "SOLTERO(A)": "S",
                "CASADO(A)": "C",
                "UNION LIBRE": "A",
                "DIVORCIADO(A)": "D",
                "VIUDO(A)": "V"
            }
            # Convertimos a mayúsculas y quitamos espacios para asegurar el match
            df_gen["EstadoCivil"] = df_gen["EstadoCivil"].str.upper().str.strip().map(map_estado_civil).fillna(df_gen["EstadoCivil"])

#------------------------------
#------------------------------

        df_gen["AFP"] = "" # Por defecto vacío
        if "Tiene IGSS" in df_gen.columns:
            # Validamos si es "SI" (limpiando posibles tildes o espacios para evitar errores)
            mask_igss = df_gen["Tiene IGSS"].astype(str).str.upper().str.replace("Í", "I").str.strip() == "SI"
            df_gen.loc[mask_igss, "AFP"] = "IGSS"
        
        # Orden y limpieza final
        orden_gen = [
            "Pais Empresa", "Código De Empleado", "Primer Nombre", "Segundo Nombre", "Otros Nombres", 
            "Primer Apellido", "Segundo Apellido", "Apellido Casada", "Genero", "EstadoCivil", 
            "Pais Nacimiento", "Fecha Nacimiento", "Pais Nacionalidad", "Cuenta Banco", "Tipo Cuenta Banco", 
            "Banco", "Dirección", "Departamento Residencia", "Municipio Residencia", "Telefono", 
            "Celular", "eMail Interno", "Profesión (100 caracteres)", "No Identificacion", "No Seguro Social", 
            "No Identificacion Tributaria", "AFP", "NUP", "Digito Verificador"
        ]
        df_gen = df_gen.reindex(columns=orden_gen)
        
        if "N° DOCUMENTO" in df_gen.columns:
            df_gen = df_gen.drop(columns=["N° DOCUMENTO"])

        # ========================================================
        # PESTAÑA 2: DEPENDIENTES
        # ========================================================
        
        cols_mae_dep = ["COD PERSONAL", "N° DOCUMENTO"]
        base_mae_dep = df_mae[[c for c in cols_mae_dep if c in df_mae.columns]].copy()
        
        cols_inc_fam = [
            "Número de Documento", "Vinculo", "N° de doc. de derechohabiente", 
            "Fecha de nacimiento", "Nombres completos", "Apellido paterno", # <-- Agregada
            "Apellido materno", "Sexo", "Nivel educativo", "Cargo y área en la que trabaja" # <-- Agregada
        ]
        base_inc_fam = df_inc_fam[[c for c in cols_inc_fam if c in df_inc_fam.columns]].copy()
        
        # Estandarización y Join
        base_mae_dep["N° DOCUMENTO"] = base_mae_dep["N° DOCUMENTO"].astype(str).str.strip()
        base_inc_fam["Número de Documento"] = base_inc_fam["Número de Documento"].astype(str).str.strip()
        df_dep = pd.merge(base_mae_dep, base_inc_fam, left_on="N° DOCUMENTO", right_on="Número de Documento", how="inner")
        
        # Renombrado de Columnas
        map_dep = {
            "COD PERSONAL": "CodigoEmpleado", 
            "Vinculo": "Parentesco", 
            # SE ELIMINÓ "Nombres completos": "nombre" PORQUE YA LA CREAMOS ARRIBA
            "Sexo": "Género", 
            "Fecha de nacimiento": "FechaNacimiento", 
            "N° de doc. de derechohabiente": "No Documento", 
            "Cargo y área en la que trabaja": "LugarTrabajo", 
            "Nivel educativo": "NivelEstudio"
        }

        # Concatenación de nombre completo para Dependientes (Guatemala)
        for col in ["Nombres completos", "Apellido paterno", "Apellido materno"]:
            if col in df_dep.columns:
                df_dep[col] = df_dep[col].astype(str).replace("-", "").str.strip()
        
        df_dep["nombre"] = (df_dep["Nombres completos"] + " " + 
                            df_dep["Apellido paterno"] + " " + 
                            df_dep["Apellido materno"]).str.replace(r'\s+', ' ', regex=True).str.strip()

        df_dep = df_dep.rename(columns=map_dep)

#----------
        # --- MAPEOS GLOBALES: DEPENDIENTES ---
        
        # 1. Mapeo de Género (ya lo tenías)
        if "Género" in df_dep.columns:
            map_genero_dep = {"MASCULINO": "M", "FEMENINO": "F"}
            df_dep["Género"] = df_dep["Género"].str.upper().map(map_genero_dep).fillna(df_dep["Género"])

        # 2. Mapeo de Estado Civil (Nuevo)
        if "EstadoCivil" in df_dep.columns:
            map_civil_dep = {
                "SOLTERO(A)": "S",
                "CASADO(A)": "C",
                "UNION LIBRE": "A",
                "DIVORCIADO(A)": "D",
                "VIUDO(A)": "V"
            }
            df_dep["EstadoCivil"] = df_dep["EstadoCivil"].str.upper().str.strip().map(map_civil_dep).fillna(df_dep["EstadoCivil"])
#------------
        # Formateo de Fecha de Nacimiento (DD/MM/YYYY) para Dependientes
        if "FechaNacimiento" in df_dep.columns:
            df_dep["FechaNacimiento"] = pd.to_datetime(df_dep["FechaNacimiento"], errors='coerce').dt.strftime('%d/%m/%Y').fillna("")
        
        # Agregar columnas vacías
        df_dep["EstadoCivil"] = ""
        df_dep["DependenciaEconomica"] = ""
        df_dep["Trabaja"] = ""
        df_dep["Estudia"] = ""
        df_dep["LugarEstudio"] = ""
            
        # Orden final
        orden_dep = [
            "CodigoEmpleado", "Parentesco", "nombre", "Género", "EstadoCivil", 
            "DependenciaEconomica", "FechaNacimiento", "No Documento", "Trabaja", 
            "LugarTrabajo", "Estudia", "NivelEstudio", "LugarEstudio"
        ]
        df_dep = df_dep.reindex(columns=orden_dep)

        # ========================================================
        # GENERACIÓN DEL EXCEL MULTI-HOJA
        # ========================================================
        # Limpieza Global: Reemplazar guiones "-" por vacío
        df_gen = df_gen.replace("-", "", regex=False)
        df_dep = df_dep.replace("-", "", regex=False)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_gen.to_excel(writer, index=False, sheet_name='Datos Generales')
            df_dep.to_excel(writer, index=False, sheet_name='Dependientes')
            
            # Autoajuste de celdas
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
        st.error(f"Error crítico en la plantilla de Expediente (Guatemala): {e}")
        return None

def procesar_honduras(f_inc, f_mae, f_req):
    try:
        # --- CARGA DE DATOS POR PESTAÑA ---
        df_org = pd.read_excel(f_req, sheet_name="Datos organizativos")
        df_cand = pd.read_excel(f_req, sheet_name="Candidatos seleccionados")
        df_mae = pd.read_excel(f_mae)

        # Limpieza inicial de nombres de columnas (quitar espacios invisibles)
        for df in [df_org, df_cand, df_mae]:
            df.columns = [str(c).strip() for c in df.columns]

        # 1️ Requerimiento – Datos Organizativos
        # Nota: Aquí usamos "TIPO DE CONV" específico de Honduras
        cols_mantener_org = ["RQ", "CLASE DE CONTRATO", "MOTIVO", "BONO ANUAL"]
        req_org = df_org[[c for c in cols_mantener_org if c in df_org.columns]].copy()

        # 2️ Requerimiento – Candidatos seleccionados
        cols_mantener_cand = ["RQ", "DNI", "ERF - BÁSICO"]
        req_cand = df_cand[[c for c in cols_mantener_cand if c in df_cand.columns]].copy()
        
        # Eliminación de duplicados por RQ
        if "RQ" in req_cand.columns:
            req_cand = req_cand.drop_duplicates(subset=["RQ"])

        # Estandarización de tipos de dato para el primer JOIN
        if "RQ" in req_org.columns: req_org["RQ"] = req_org["RQ"].astype(str).str.strip()
        if "RQ" in req_cand.columns: req_cand["RQ"] = req_cand["RQ"].astype(str).str.strip()

        # 3️ Combinación Base_Requerimiento
        base_req = pd.merge(req_org, req_cand, on="RQ", how="inner")

        # 4️ Preparación del maestro de personal
        cols_mantener_mae = [
            "COD PERSONAL", "COD POSICION", "ESTADO", "FECHA DE BAJA", 
            "FECHA FIN CONTRATO", "FECHA INICIO CONTRATO", 
            "FECHA INGRESO AL GRUPO", "N° DOCUMENTO", "CLASE CONTRATO"
        ]
        base_mae = df_mae[[c for c in cols_mantener_mae if c in df_mae.columns]].copy()

        # Estandarización de tipos de dato para el segundo JOIN
        if "N° DOCUMENTO" in base_mae.columns: base_mae["N° DOCUMENTO"] = base_mae["N° DOCUMENTO"].astype(str).str.strip()
        if "DNI" in base_req.columns: base_req["DNI"] = base_req["DNI"].astype(str).str.strip()

        # 5️ Generación final de la plantilla “Empleos”
        df_final = pd.merge(base_mae, base_req, left_on="N° DOCUMENTO", right_on="DNI", how="inner")

        # Renombrado de columnas (incluyendo TIPO DE CONV)
        mapeo_nombres = {
            "COD PERSONAL": "Código De Empleado",
            "COD POSICION": "Codigo Plaza",
            "ESTADO": "Estado",
            "CLASE DE CONTRATO": "Tipo de Planilla",
            "FECHA INGRESO AL GRUPO": "Fecha Ingreso",
            "FECHA DE BAJA": "Fecha de Retiro",
            "MOTIVO": "Motivo de Retiro",
            "CLASE CONTRATO": "Tipo de Contrato",
            "ERF - BÁSICO": "Salario Mensual",
            "FECHA INICIO CONTRATO": "Fecha Inicio Contrato",
            "FECHA FIN CONTRATO": "Fecha Fin Contrato",
            "BONO ANUAL": "Bonificacion Decreto"
        }
        df_final = df_final.rename(columns=mapeo_nombres)

        # Agregar campos vacíos
        df_final["Gastos de Representacion"] = ""
        df_final["Numero Horas por Mes"] = ""
        
        # Autocompletado de Jornada
        df_final["Jornada"] = "Jornada Administrativa"

        # 6️ Reordenamiento de columnas
        orden_final = [
            "Código De Empleado", "Codigo Plaza", "Estado", "Tipo de Planilla", "Fecha Ingreso",
            "Fecha de Retiro", "Motivo de Retiro", "Tipo de Contrato", "Jornada", "Salario Mensual",
            "Fecha Inicio Contrato", "Fecha Fin Contrato", "Bonificacion Decreto", 
            "Gastos de Representacion", "Numero Horas por Mes", "N° DOCUMENTO"
        ]

#---------------
        # --- AJUSTES FINALES: ESTADO ---
        if "Estado" in df_final.columns:
            # Reemplazamos los valores exactos
            map_estado = {"ACTIVO": "Activo", "CESADO": "Retirado"}
            df_final["Estado"] = df_final["Estado"].replace(map_estado)
#---------------
        
        df_final = df_final.reindex(columns=orden_final)

        # 7️ Limpieza final
        if "N° DOCUMENTO" in df_final.columns:
            df_final = df_final.drop(columns=["N° DOCUMENTO"])

        # Generación de Excel con autoajuste de celdas
        
        
        # Limpieza Global: Reemplazar guiones "-" por vacío
        df_final = df_final.replace("-", "", regex=False)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Empleos')
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
                worksheet.column_dimensions[column].width = (max_length + 2)
        
        return output.getvalue()

    except Exception as e:
        st.error(f"Error en la lógica de Honduras (Empleos): {e}")
        return None

def procesar_expediente_honduras(f_inc, f_mae):
    try:
        # 1️ Carga del Maestro
        df_mae = pd.read_excel(f_mae)
        df_mae.columns = [str(c).strip() for c in df_mae.columns]

        # 2️ Carga Inteligente del archivo de Incorporación (Buscador de Pestañas)
        hojas_inc = pd.read_excel(f_inc, sheet_name=None)
        df_inc_per = None
        df_inc_fam = None
        
        for nombre_hoja, df in hojas_inc.items():
            df.columns = [str(c).strip() for c in df.columns]
            # Identificamos la pestaña "Personas" buscando columnas clave de Honduras
            if "Nro. Documento" in df.columns and "Tienes cta. en bancaria en BAC" in df.columns:
                df_inc_per = df
            # Identificamos la pestaña "Beneficiarios"
            if "N° de doc. de beneficiario" in df.columns and "Número de Documento" in df.columns:
                df_inc_fam = df
                
        if df_inc_per is None or df_inc_fam is None:
            st.error("⚠️ No se encontraron las estructuras esperadas en el Excel de Incorporación de Honduras.")
            return None

        # ========================================================
        # PESTAÑA 1: DATOS GENERALES
        # ========================================================
        
        # Preparar Maestro
        cols_mae_gen = ["PAIS", "COD PERSONAL", "N° DOCUMENTO", "SEXO"]
        base_mae_gen = df_mae[[c for c in cols_mae_gen if c in df_mae.columns]].copy()
        
        # Preparar Incorporación (Personas)
        cols_inc_gen = [
            "Nro. Documento", "Nombres completos", "Apellido paterno", "Apellido materno", 
            "Estado civil", "Fecha de nacimiento", "Nacionalidad", "País de nacimiento", 
            "Correo personal", "Número de celular", "Teléfono fijo", "Tienes cta. en bancaria en BAC", 
            "Numero de cta. bancaria", "Tipo de banco para apertura cta. bancaria", 
            "Nombre de carrera", "Departamento residencia", "Municipio residencia", "Dirección exacta"
        ]
        base_inc_per = df_inc_per[[c for c in cols_inc_gen if c in df_inc_per.columns]].copy()
        
        # Estandarización y Join
        base_mae_gen["N° DOCUMENTO"] = base_mae_gen["N° DOCUMENTO"].astype(str).str.strip()
        base_inc_per["Nro. Documento"] = base_inc_per["Nro. Documento"].astype(str).str.strip()
        df_gen = pd.merge(base_mae_gen, base_inc_per, left_on="N° DOCUMENTO", right_on="Nro. Documento", how="inner")
        
        # Separación de la columna "Nombres completos" (SplitTextByDelimiter)
        if "Nombres completos" in df_gen.columns:
            # Separamos por el primer espacio encontrado
            split_names = df_gen["Nombres completos"].astype(str).str.split(" ", n=1, expand=True)
            df_gen["Primer Nombre"] = split_names[0] if 0 in split_names.columns else ""
            df_gen["Segundo Nombre"] = split_names[1] if 1 in split_names.columns else ""
        else:
            df_gen["Primer Nombre"] = ""
            df_gen["Segundo Nombre"] = ""
            
        # Renombrado de Columnas
        map_gen = {
            "PAIS": "Pais Empresa", "COD PERSONAL": "Código De Empleado", "SEXO": "Genero",
            "Apellido paterno": "Primer Apellido", "Apellido materno": "Segundo Apellido",
            "Estado civil": "EstadoCivil", "País de nacimiento": "Pais Nacimiento",
            "Fecha de nacimiento": "Fecha Nacimiento", "Nacionalidad": "Pais Nacionalidad",
            "Tipo de banco para apertura cta. bancaria": "Tipo Cuenta Banco",
            "Dirección exacta": "Dirección", "Departamento residencia": "Departamento Residencia",
            "Municipio residencia": "Municipio Residencia", "Teléfono fijo": "Telefono",
            "Número de celular": "Celular", "Correo personal": "eMail Interno",
            "Nombre de carrera": "Profesión (100 caracteres)", "Nro. Documento": "No Identificacion",
            "Numero de cta. bancaria": "Cuenta Banco", "Tienes cta. en bancaria en BAC": "Banco"
        }
        df_gen = df_gen.rename(columns=map_gen)
        # Formateo de Fecha de Nacimiento (DD/MM/YYYY) para Datos Generales
        if "Fecha Nacimiento" in df_gen.columns:
            df_gen["Fecha Nacimiento"] = pd.to_datetime(df_gen["Fecha Nacimiento"], errors='coerce').dt.strftime('%d/%m/%Y').fillna("")
        
        # Agregar columnas nuevas/vacías
        df_gen["Otros Nombres"] = ""
        df_gen["Apellido Casada"] = ""
        df_gen["No Seguro Social"] = ""
        df_gen["No Identificacion Tributaria"] = ""
        df_gen["AFP"] = ""
        df_gen["NUP"] = ""
        df_gen["Digito Verificador"] = ""
        
        #--------------------------
        # --- MAPEOS GLOBALES: DATOS GENERALES ---
        
        # 1. Mapeo de Género
        if "Genero" in df_gen.columns:
            map_genero = {"MASCULINO": "M", "FEMENINO": "F"}
            df_gen["Genero"] = df_gen["Genero"].str.upper().map(map_genero).fillna(df_gen["Genero"])

        # 2. Mapeo de Países (Nacimiento y Nacionalidad) - CORREGIDO
        map_paises = {
            "el salvador": "sv", 
            "guatemala": "gt", 
            "nicaragua": "ni",
            "honduras": "hn", 
            "panama": "pa", 
            "costa rica": "cr", 
            "estados unidos": "us"
        }
        for col in ["Pais Nacimiento", "Pais Nacionalidad"]:
            if col in df_gen.columns:
                # Convertimos lo que viene del Excel a minúsculas y sin espacios para que haga "match" perfecto
                df_gen[col] = df_gen[col].astype(str).str.lower().str.strip().map(map_paises).fillna(df_gen[col])

        # 3. Mapeo de Tipo de Cuenta
        if "Tipo Cuenta Banco" in df_gen.columns:
            map_cuentas = {"Ahorro": "A", "Cuenta Corriente": "C"}
            df_gen["Tipo Cuenta Banco"] = df_gen["Tipo Cuenta Banco"].map(map_cuentas).fillna(df_gen["Tipo Cuenta Banco"])
        # 4. Mapeo de Estado Civil (Datos Generales)
        if "EstadoCivil" in df_gen.columns:
            map_estado_civil = {
                "SOLTERO(A)": "S",
                "CASADO(A)": "C",
                "UNION LIBRE": "A",
                "DIVORCIADO(A)": "D",
                "VIUDO(A)": "V"
            }
            # Convertimos a mayúsculas y quitamos espacios para asegurar el match
            df_gen["EstadoCivil"] = df_gen["EstadoCivil"].str.upper().str.strip().map(map_estado_civil).fillna(df_gen["EstadoCivil"])
        
        #--------------------------

        # Orden y limpieza final (Asegurando las 29 columnas obligatorias)
        orden_gen = [
            "Pais Empresa", "Código De Empleado", "Primer Nombre", "Segundo Nombre", "Otros Nombres", 
            "Primer Apellido", "Segundo Apellido", "Apellido Casada", "Genero", "EstadoCivil", 
            "Pais Nacimiento", "Fecha Nacimiento", "Pais Nacionalidad", "Cuenta Banco", "Tipo Cuenta Banco", 
            "Banco", "Dirección", "Departamento Residencia", "Municipio Residencia", "Telefono", 
            "Celular", "eMail Interno", "Profesión (100 caracteres)", "No Identificacion", "No Seguro Social", 
            "No Identificacion Tributaria", "AFP", "NUP", "Digito Verificador"
        ]
        df_gen = df_gen.reindex(columns=orden_gen)
        
        if "N° DOCUMENTO" in df_gen.columns:
            df_gen = df_gen.drop(columns=["N° DOCUMENTO"])

        # ========================================================
        # PESTAÑA 2: DEPENDIENTES
        # ========================================================
        
        # Preparar Maestro
        cols_mae_dep = ["COD PERSONAL", "N° DOCUMENTO"]
        base_mae_dep = df_mae[[c for c in cols_mae_dep if c in df_mae.columns]].copy()
        
        # Preparar Incorporación (Beneficiarios) - Agregando 'Número de Documento' para el join
        cols_inc_fam = [
            "Número de Documento", "N° de doc. de beneficiario", "Fecha de nacimiento", 
            "Nombres completos", "Sexo", "Nivel educativo"
        ]
        base_inc_fam = df_inc_fam[[c for c in cols_inc_fam if c in df_inc_fam.columns]].copy()
        
        # Estandarización y Join
        base_mae_dep["N° DOCUMENTO"] = base_mae_dep["N° DOCUMENTO"].astype(str).str.strip()
        base_inc_fam["Número de Documento"] = base_inc_fam["Número de Documento"].astype(str).str.strip()
        df_dep = pd.merge(base_mae_dep, base_inc_fam, left_on="N° DOCUMENTO", right_on="Número de Documento", how="inner")
        
        # Renombrado de Columnas
        map_dep = {
            "COD PERSONAL": "CodigoEmpleado", "Nombres completos": "nombre", 
            "Sexo": "Género", "Fecha de nacimiento": "FechaNacimiento", 
            "N° de doc. de beneficiario": "No Documento", "Nivel educativo": "NivelEstudio"
        }
        df_dep = df_dep.rename(columns=map_dep)

        #-----------------------
        # --- MAPEOS GLOBALES: DEPENDIENTES ---
        
        # 1. Mapeo de Género (ya lo tenías)
        if "Género" in df_dep.columns:
            map_genero_dep = {"MASCULINO": "M", "FEMENINO": "F"}
            df_dep["Género"] = df_dep["Género"].str.upper().map(map_genero_dep).fillna(df_dep["Género"])

        # 2. Mapeo de Estado Civil (Nuevo)
        if "EstadoCivil" in df_dep.columns:
            map_civil_dep = {
                "SOLTERO(A)": "S",
                "CASADO(A)": "C",
                "UNION LIBRE": "A",
                "DIVORCIADO(A)": "D",
                "VIUDO(A)": "V"
            }
            df_dep["EstadoCivil"] = df_dep["EstadoCivil"].str.upper().str.strip().map(map_civil_dep).fillna(df_dep["EstadoCivil"])
        #-----------------------

        # Formateo de Fecha de Nacimiento (DD/MM/YYYY) para Dependientes
        if "FechaNacimiento" in df_dep.columns:
            df_dep["FechaNacimiento"] = pd.to_datetime(df_dep["FechaNacimiento"], errors='coerce').dt.strftime('%d/%m/%Y').fillna("")
        
        # Agregar columnas vacías
        for col in ["Parentesco", "EstadoCivil", "DependenciaEconomica", "Trabaja", "LugarTrabajo", "Estudia", "LugarEstudio"]:
            df_dep[col] = ""
            
        # Orden final (Asegurando las 13 columnas obligatorias)
        orden_dep = [
            "CodigoEmpleado", "Parentesco", "nombre", "Género", "EstadoCivil", 
            "DependenciaEconomica", "FechaNacimiento", "No Documento", "Trabaja", 
            "LugarTrabajo", "Estudia", "NivelEstudio", "LugarEstudio"
        ]
        df_dep = df_dep.reindex(columns=orden_dep)

        # ========================================================
        # GENERACIÓN DEL EXCEL MULTI-HOJA
        # ========================================================
        
        # Limpieza Global: Reemplazar guiones "-" por vacío
        df_gen = df_gen.replace("-", "", regex=False)
        df_dep = df_dep.replace("-", "", regex=False)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_gen.to_excel(writer, index=False, sheet_name='Datos Generales')
            df_dep.to_excel(writer, index=False, sheet_name='Dependientes')
            
            # Autoajuste de celdas
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
        st.error(f"Error crítico en la plantilla de Expediente (Honduras): {e}")
        return None

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


# INTERFAZ DE USUARIO (FRONTEND)

st.markdown(f"<h1 style='text-align: center; color: {color_naranja};'>Automatización de Plantillas de Personal</h1>", unsafe_allow_html=True)
#st.markdown(f"<h4 style='text-align: center; color: {color_gris};'>Generador estandarizado de Empleos y Expediente</h4>", unsafe_allow_html=True)
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
            excel_expediente = procesar_expediente_el_salvador(file_incorporacion, file_maestro)
            
        elif pais_mapeado == "Guatemala":
            excel_empleos = procesar_guatemala(file_incorporacion, file_maestro, file_requerimientos)
            excel_expediente = procesar_expediente_guatemala(file_incorporacion, file_maestro)
            
        elif pais_mapeado == "Honduras":
            excel_empleos = procesar_honduras(file_incorporacion, file_maestro, file_requerimientos)
            excel_expediente = procesar_expediente_honduras(file_incorporacion, file_maestro)
        
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


# --- SECCIÓN DE CONTACTO / SOPORTE AL FINAL DE LA PÁGINA ---
st.divider() # Línea separadora sutil
# 1. Franja verde corporativa
st.markdown("""
    <div style="background-color: #009A3F; padding: 10px 20px; border-radius: 5px; margin-bottom: 25px;">
        <h4 style="color: white; margin: 0; font-family: sans-serif; font-size: 18px;">Si tienes alguna consulta, no dudes en contactarnos</h4>
    </div>
""", unsafe_allow_html=True)

# 2. Función para crear el diseño de cada perfil
def crear_tarjeta_perfil(nombre, cargo, correo, url_imagen):
    return f"""
    <div style="display: flex; align-items: center; margin-bottom: 10px;">
        <img src="{url_imagen}" style="border-radius: 50%; width: 50px; height: 50px; object-fit: cover; margin-right: 15px; border: 1px solid #ddd;">
        <div style="line-height: 1.2;">
            <a href="mailto:{correo}" style="text-decoration: none; color: #333; font-weight: bold; font-size: 14px;">{nombre}</a><br>
            <span style="font-size: 11px; color: #666; text-transform: uppercase;">{cargo}</span>
        </div>
    </div>
    """

# 3. Creamos las 3 columnas en Streamlit
col_soporte1, col_soporte2, col_soporte3 = st.columns(3)

# 4. Llenamos cada columna con la información
with col_soporte1:
    st.markdown(crear_tarjeta_perfil(
        nombre="Jampierre Balabarca Nicasio", 
        cargo="Líder de People Data & Analytics", 
        correo="LBalabarcaN@ransa.net", 
        # Usamos un generador automático de iniciales si no tienes la URL de la foto real a la mano
        url_imagen="https://ui-avatars.com/api/?name=Jampierre+Balabarca&background=009A3F&color=fff" 
    ), unsafe_allow_html=True)

with col_soporte2:
    st.markdown(crear_tarjeta_perfil(
        nombre="Fabian Martin Alvarado Vargas", 
        cargo="Analista JR Desarrollo de Soluciones", 
        correo="fFalvaradoV@ransa.net", 
        url_imagen="https://ui-avatars.com/api/?name=Fabian+Alvarado&background=009A3F&color=fff"
    ), unsafe_allow_html=True)

with col_soporte3:
    st.markdown(crear_tarjeta_perfil(
        nombre="Kevin Yago Castrejon Sosa", 
        cargo="Analista JR Mejora de Procesos", 
        correo="KCastrejonS@ransa.net", 
        url_imagen="https://ui-avatars.com/api/?name=Kevin+Castrejon&background=009A3F&color=fff"
    ), unsafe_allow_html=True)