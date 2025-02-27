import streamlit as st
import pandas as pd
from io import BytesIO

# --- Sidebar ---
with st.sidebar:
    opcion = st.radio("Menú Principal", ["Grupo Cerrado", "Grupo Abierto"])

# --- Opción Grupo Abierto ---
if opcion == "Grupo Abierto":
    st.header("Grupo Abierto")
    
    with st.expander("Ver Instrucciones"):
        st.write("""
**Instrucciones - Grupo Abierto:**

1. Cargue los archivos de ASISTENCIAS y luego los archivos de CALIFICACIONES.
2. Ingrese el número de ciclo.
3. Presione **Procesar Grupo Abierto** para generar el archivo consolidado.

*Salida:*  
- Hoja 1: Consolidado de ASISTENCIAS.  
- Hojas siguientes: Datos de CALIFICACIONES.  
- Archivo nombrado como "Seguimiento CLEVEL Ciclo X" (X = número de ciclo).
""")
    
    st.subheader("Carga de archivos de ASISTENCIAS")
    uploaded_asistencias_abierto = st.file_uploader(
        "Seleccione archivos de Excel de ASISTENCIAS", 
        type=["xlsx"], 
        accept_multiple_files=True, 
        key="asistencias_abierto"
    )
    
    st.subheader("Carga de archivos de CALIFICACIONES")
    uploaded_calificaciones_abierto = st.file_uploader(
        "Seleccione archivos de Excel de CALIFICACIONES", 
        type=["xlsx"], 
        accept_multiple_files=True, 
        key="calificaciones_abierto"
    )
    
    ciclo_abierto = st.number_input("Ingrese el número de ciclo", min_value=1, step=1, key="ciclo_abierto")

    if st.button("Procesar Grupo Abierto"):
        if not uploaded_asistencias_abierto:
            st.error("Debe cargar los archivos de ASISTENCIAS (Grupo Abierto).")
        elif not uploaded_calificaciones_abierto:
            st.error("Debe cargar los archivos de CALIFICACIONES (Grupo Abierto).")
        else:
            # --- Procesamiento de ASISTENCIAS (Grupo Abierto) ---
            lista_dfs_asist = []
            for file in uploaded_asistencias_abierto:
                try:
                    file.seek(0)
                    df_cabecera = pd.read_excel(file, header=None, nrows=1)
                    valor_b1 = df_cabecera.iloc[0, 1]
                    
                    file.seek(0)
                    df = pd.read_excel(file, header=3)
                    df.sort_values(by="Porcentaje", ascending=True, inplace=True)
                    df["Curso"] = valor_b1
                    
                    nombre_archivo = file.name
                    codigo = nombre_archivo.split("_", 1)[0]
                    df["Codigo"] = codigo
                    
                    # Reordenar para que las últimas columnas sean "Porcentaje", "Curso" y "Codigo"
                    ultimas = ["Porcentaje", "Curso", "Codigo"]
                    otras = [col for col in df.columns if col not in ultimas]
                    df = df[otras + ultimas]
                    
                    # Reemplazar celdas vacías
                    df.fillna("No aplica", inplace=True)
                    
                    lista_dfs_asist.append(df)
                except Exception as e:
                    st.error(f"Error en ASISTENCIAS: {file.name}: {e}")
            
            if not lista_dfs_asist:
                st.error("No se pudo procesar ningún archivo de ASISTENCIAS.")
            else:
                df_consolidado_abierto = pd.concat(lista_dfs_asist, ignore_index=True)
                # Eliminar duplicados en columnas y reordenar
                df_consolidado_abierto = df_consolidado_abierto.loc[:, ~df_consolidado_abierto.columns.duplicated()]
                ultimas = ["Porcentaje", "Curso", "Codigo"]
                otras = [col for col in df_consolidado_abierto.columns if col not in ultimas]
                df_consolidado_abierto = df_consolidado_abierto[otras + ultimas]
            
            # --- Crear mapeos a partir del consolidado de ASISTENCIAS ---
            mapping_num_id_to_grupos = df_consolidado_abierto.drop_duplicates(subset="Número de ID").set_index("Número de ID")["Grupos"].to_dict()
            mapping_codigo_curso = df_consolidado_abierto.drop_duplicates(subset="Codigo").set_index("Codigo")["Curso"].to_dict()
            
            # --- Procesamiento de CALIFICACIONES (Grupo Abierto) ---
            dfs_calificaciones_abierto = {}
            for file in uploaded_calificaciones_abierto:
                try:
                    file.seek(0)
                    df = pd.read_excel(file)
                    
                    # Forzar conversión de "Total del curso (Real)" a numérico
                    df["Total del curso (Real)"] = pd.to_numeric(df["Total del curso (Real)"], errors='coerce')
                    
                    # Agregar la columna 'Grupos' mapeando desde "Número de ID"
                    df["Grupos"] = df["Número de ID"].map(mapping_num_id_to_grupos)
                    
                    # Ordenar de forma ascendente por "Total del curso (Real)"
                    df.sort_values(by="Total del curso (Real)", ascending=True, inplace=True)
                    
                    nombre_archivo = file.name
                    codigo = nombre_archivo.split(" ", 1)[0]
                    df["Codigo"] = codigo
                    
                    df["Curso"] = mapping_codigo_curso.get(codigo, None)
                    
                    ultimas_cal = ["Grupos", "Codigo", "Curso"]
                    otras_cal = [col for col in df.columns if col not in ultimas_cal]
                    df = df[otras_cal + ultimas_cal]
                    
                    dfs_calificaciones_abierto[codigo] = df
                except Exception as e:
                    st.error(f"Error en CALIFICACIONES: {file.name}: {e}")
            
            # --- Consolidar en un único archivo Excel ---
            output_final_abierto = BytesIO()
            try:
                with pd.ExcelWriter(output_final_abierto, engine='xlsxwriter') as writer:
                    # Primera hoja: consolidado de ASISTENCIAS
                    df_consolidado_abierto.to_excel(writer, sheet_name="Asistencias", index=False)
                    # Siguientes hojas: archivos de CALIFICACIONES
                    for sheet_name, df_sheet in dfs_calificaciones_abierto.items():
                        df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                output_final_abierto.seek(0)
                
                file_name_abierto = f"Seguimiento CLEVEL Abierto Ciclo {int(ciclo_abierto)}.xlsx"
                st.download_button(
                    label="Descargar archivo consolidado",
                    data=output_final_abierto,
                    file_name=file_name_abierto,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("Procesamiento de Grupo Abierto completado.")
            except Exception as e:
                st.error(f"Error al generar el archivo consolidado: {e}")


# --- Opción Grupo Cerrado ---
if opcion == "Grupo Cerrado":
    st.header("Grupo Cerrado")
    
    # Instrucciones en un expander (por defecto contraído)
    with st.expander("Ver Instrucciones"):
        st.write("""
**Instrucciones - Grupo Cerrado:**

1. Cargue los archivos de ASISTENCIAS y luego los archivos de CALIFICACIONES.
2. Ingrese el número de ciclo.
3. Presione **Procesar** para generar el archivo consolidado.

*Salida:*  
- Hoja 1: Consolidado de ASISTENCIAS.  
- Hojas siguientes: Datos de CALIFICACIONES.  
- Archivo nombrado como "Seguimiento CLEVEL Ciclo X" (X = número de ciclo).
""")

    # --- Sección para cargar archivos de ASISTENCIAS ---
    st.subheader("Carga de archivos de ASISTENCIAS")
    uploaded_asistencias = st.file_uploader(
        "Seleccione archivos de Excel de ASISTENCIAS", 
        type=["xlsx"], 
        accept_multiple_files=True, 
        key="asistencias"
    )
    
    # --- Sección para cargar archivos de CALIFICACIONES ---
    st.subheader("Carga de archivos de CALIFICACIONES")
    uploaded_calificaciones = st.file_uploader(
        "Seleccione archivos de Excel de CALIFICACIONES", 
        type=["xlsx"], 
        accept_multiple_files=True, 
        key="calificaciones"
    )
    
    # --- Input para número de ciclo ---
    ciclo = st.number_input("Ingrese el número de ciclo", min_value=1, step=1)
    
    # --- Botón para iniciar el procesamiento ---
    if st.button("Procesar Grupo Cerrado"):
        if not uploaded_asistencias:
            st.error("Debe cargar los archivos de ASISTENCIAS.")
        elif not uploaded_calificaciones:
            st.error("Debe cargar los archivos de CALIFICACIONES.")
        else:
            # Procesamiento de ASISTENCIAS
            lista_dfs = []
            for file in uploaded_asistencias:
                try:
                    file.seek(0)
                    df_cabecera = pd.read_excel(file, header=None, nrows=1)
                    valor_b1 = df_cabecera.iloc[0, 1]  # Valor de la celda B1
                    
                    file.seek(0)
                    df = pd.read_excel(file, header=3)
                    df.sort_values(by="Porcentaje", ascending=True, inplace=True)
                    df["Curso"] = valor_b1
                    
                    # Obtener el nombre del archivo y extraer la parte anterior al primer guion bajo
                    nombre_archivo = file.name
                    nombre_modificado = nombre_archivo.split("_", 1)[0]
                    df["Codigo"] = nombre_modificado
                    
                    lista_dfs.append(df)
                except Exception as e:
                    st.error(f"Error procesando el archivo {file.name} en ASISTENCIAS: {e}")
            
            if not lista_dfs:
                st.error("No se pudo procesar ningún archivo de ASISTENCIAS.")
            else:
                try:
                    df_consolidado = pd.concat(lista_dfs, ignore_index=True)
                except Exception as e:
                    st.error(f"Error al consolidar archivos de ASISTENCIAS: {e}")
                    df_consolidado = None
            
            if df_consolidado is not None:
                # Crear mapeo: Codigo -> Curso a partir del consolidado
                mapping_codigo_curso = df_consolidado.drop_duplicates(subset="Codigo").set_index("Codigo")["Curso"].to_dict()
                
                # Procesamiento de CALIFICACIONES
                dfs_calificaciones = {}
                for file in uploaded_calificaciones:
                    try:
                        file.seek(0)
                        df = pd.read_excel(file)

                        df["Total del curso (Real)"] = pd.to_numeric(df["Total del curso (Real)"], errors='coerce')
                        
                        nombre_archivo = file.name
                        # Extraer el código usando el primer espacio como separador
                        codigo = nombre_archivo.split(" ", 1)[0]
                        df["Codigo"] = codigo
                        
                        # Asignar la columna Curso utilizando el mapeo generado
                        df["Curso"] = mapping_codigo_curso.get(codigo, None)
                        
                        # Ordenar en forma descendente por "Total del curso (Real)"
                        df.sort_values(by="Total del curso (Real)", ascending=True, inplace=True)
                        
                        # Usar el mismo código para asignar el nombre de la hoja
                        hoja_nombre = codigo
                        dfs_calificaciones[hoja_nombre] = df
                    except Exception as e:
                        st.error(f"Error procesando el archivo {file.name} en CALIFICACIONES: {e}")
                
                # Consolidar en un único archivo Excel con múltiples hojas:
                # La primera hoja contendrá el df_consolidado de ASISTENCIAS
                output_final = BytesIO()
                try:
                    with pd.ExcelWriter(output_final, engine='xlsxwriter') as writer:
                        # Hoja 1: Asistencias
                        df_consolidado.to_excel(writer, sheet_name="Asistencias", index=False)
                        
                        # Hojas siguientes: Calificaciones
                        for sheet_name, df_sheet in dfs_calificaciones.items():
                            df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    output_final.seek(0)
                    
                    # Nombre del archivo de salida basado en el ciclo ingresado
                    file_name = f"Seguimiento CLEVEL Cerrado Ciclo {int(ciclo)}.xlsx"
                    
                    st.download_button(
                        label="Descargar archivo consolidado",
                        data=output_final,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("Procesamiento completado.")
                except Exception as e:
                    st.error(f"Error al generar el archivo consolidado: {e}")
