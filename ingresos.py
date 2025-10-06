import streamlit as st
import pandas as pd
import io

# Título de la aplicación
st.title("Procesador de Archivos Excel")

# Cargar archivo Excel
uploaded_file = st.file_uploader("Carga tu archivo Excel", type=["xlsx", "xls"])

if uploaded_file is not None:
    # Leer el archivo Excel, comenzando desde la fila 23 (índice 22 en Pandas)
    df = pd.read_excel(uploaded_file, skiprows=22)
    
    # Mostrar las columnas del archivo para diagnóstico
    st.subheader("Columnas encontradas en el archivo Excel")
    st.write(df.columns.tolist())
    
    # Visualizar los datos originales
    st.subheader("Vista previa de los datos originales")
    st.dataframe(df)
    
    # Definir las columnas esperadas
    expected_columns = ['Local', 'Nombre o ID', 'Costo del servicio', 'IVA del costo', 'Monto neto']
    # Agregar ambas variantes de 'Monto Bruto'
    monto_bruto_variants = ['Monto Bruto', 'Monto bruto']
    
    # Verificar si alguna variante de 'Monto Bruto' está presente
    monto_bruto_col = None
    for variant in monto_bruto_variants:
        if variant in df.columns:
            monto_bruto_col = variant
            break
    
    # Verificar si todas las columnas esperadas están presentes
    missing_columns = [col for col in expected_columns if col not in df.columns]
    if monto_bruto_col is None:
        missing_columns.append("Monto Bruto o Monto bruto")
    
    if missing_columns:
        st.error(f"No se encontraron las siguientes columnas en el archivo: {missing_columns}. Por favor, verifica los nombres de las columnas.")
    else:
        # Renombrar la columna 'Monto Bruto' o 'Monto bruto' a 'Monto Bruto' para consistencia
        if monto_bruto_col != 'Monto Bruto':
            df = df.rename(columns={monto_bruto_col: 'Monto Bruto'})
        
        # Filtrar el local a ignorar
        ignored_local = "DIPLOMATURA DE TRANSF EDUCATIVAS CON IA"
        df_filtered = df[df['Local'] != ignored_local]
        
        # Limpiar y convertir columnas numéricas a numérico (maneja errores)
        numeric_columns = ['Monto neto', 'Costo del servicio', 'IVA del costo', 'Monto Bruto']
        for col in numeric_columns:
            if col in df_filtered.columns:
                df_filtered[col] = pd.to_numeric(df_filtered[col], errors='coerce').fillna(0)
        
        # Normalizar 'Nombre o ID': mayúsculas y strip espacios
        df_filtered['Nombre o ID'] = df_filtered['Nombre o ID'].astype(str).str.upper().str.strip()
        
        # Tabla 1: Agrupar por 'Local' y 'Nombre o ID', y sumar las columnas especificadas
        columns_to_sum = ['Monto Bruto', 'Costo del servicio', 'IVA del costo', 'Monto neto']
        grouped = df_filtered.groupby(['Local', 'Nombre o ID'])[columns_to_sum].sum().reset_index()
        
        # Calcular la suma total de las columnas numéricas para Tabla 1
        totals = grouped[columns_to_sum].sum()
        totals_row = pd.DataFrame([['Total', '', *totals]], columns=['Local', 'Nombre o ID', *columns_to_sum])
        grouped_with_totals = pd.concat([grouped, totals_row], ignore_index=True)
        
        # Visualizar Tabla 1
        st.subheader("Resultados procesados: Sumas por Local e ID")
        st.dataframe(grouped_with_totals)
        
        # DIAGNÓSTICO: Mostrar IDs únicos en el archivo (normalizados)
        unique_ids = df_filtered['Nombre o ID'].unique()
        st.subheader("Diagnóstico: IDs únicos encontrados en 'Nombre o ID' (normalizados a mayúsculas)")
        st.write(unique_ids.tolist())
        
        # Definir los conceptos y sus IDs (normalizados)
        concepts = {
            "GP2032 - Concepto 286 (Comedor SL)": ["COMEDOR CAJA 3", "L40020037", "PYB005762", "PYB005808", "PYB004172"],
            "GP2028 - Concepto 282 (Deportes SL)": ["CAJA DEPORTE SL 2", "PYB004313"],
            "GP2034 - Concepto 288 (Deportes VM)": ["CAJA DEPORTE VM 1", "DEPORTES VARIOS"],
            "GP2036 - Concepto 290 (Comedor VM)": ["COMEDOR CAJA 1 VM", "COMEDOR CAJA 2 VM", "PYB005420"]
        }
        normalized_concepts = {}
        for concept, ids in concepts.items():
            normalized_ids = [id_.upper().strip() for id_ in ids]
            normalized_concepts[concept] = normalized_ids
        
        # DIAGNÓSTICO: IDs no registrados
        all_registered_ids = []
        for ids in normalized_concepts.values():
            all_registered_ids.extend(ids)
        unregistered_ids = [id_ for id_ in unique_ids if id_ not in all_registered_ids]
        if unregistered_ids:
            st.warning(f"Se encontraron IDs no registrados (no asociados a ningún concepto): {unregistered_ids}")
        else:
            st.info("Todos los IDs están registrados en algún concepto.")
        
        # Tabla 2: Discriminación por Conceptos (Montos Netos)
        concept_data_neto = []
        st.subheader("Diagnóstico detallado por concepto (Montos Netos)")
        for concept, ids in normalized_concepts.items():
            # Filtrar los datos para los IDs del concepto
            matching_mask = df_filtered['Nombre o ID'].isin(ids)
            concept_df = df_filtered[matching_mask]
            
            # Desglose: IDs encontrados y sus sumas
            found_ids = concept_df['Nombre o ID'].unique().tolist()
            st.write(f"**{concept}:** IDs buscados: {ids}")
            st.write(f"IDs encontrados: {found_ids} (de {len(found_ids)}/{len(ids)})")
            
            if len(concept_df) > 0:
                # Mostrar desglose por ID
                breakdown = concept_df.groupby('Nombre o ID')['Monto neto'].sum().round(2)
                st.write("Desglose por ID:", breakdown.to_dict())
            else:
                st.warning(f"No se encontraron datos para {concept}. Verifica los nombres de IDs.")
            
            # Sumar el Monto neto
            monto_neto_total = concept_df['Monto neto'].sum()
            # Calcular IIBB (Monto neto × 0.0001)
            iibb = monto_neto_total * 0.0001
            # Calcular TOTAL INGRESO (Monto neto - IIBB)
            total_ingreso = monto_neto_total - iibb
            concept_data_neto.append([concept, monto_neto_total, iibb, total_ingreso])
        
        # Crear DataFrame para la tabla de montos netos
        concepts_df_neto = pd.DataFrame(concept_data_neto, columns=['Concepto', 'Monto Neto Total', 'IIBB', 'TOTAL INGRESO'])
        
        # Calcular la fila de totales para la tabla de montos netos
        totals_neto = concepts_df_neto[['Monto Neto Total', 'IIBB', 'TOTAL INGRESO']].sum()
        totals_neto_row = pd.DataFrame([['Total', *totals_neto]], columns=['Concepto', 'Monto Neto Total', 'IIBB', 'TOTAL INGRESO'])
        concepts_df_neto_with_totals = pd.concat([concepts_df_neto, totals_neto_row], ignore_index=True)
        
        # Formatear las columnas numéricas a 2 decimales
        concepts_df_neto_with_totals['Monto Neto Total'] = concepts_df_neto_with_totals['Monto Neto Total'].round(2)
        concepts_df_neto_with_totals['IIBB'] = concepts_df_neto_with_totals['IIBB'].round(2)
        concepts_df_neto_with_totals['TOTAL INGRESO'] = concepts_df_neto_with_totals['TOTAL INGRESO'].round(2)
        
        # Visualizar Tabla 2 (Montos Netos)
        st.subheader("Discriminación por Conceptos (Montos Netos)")
        st.dataframe(concepts_df_neto_with_totals)
        
        # Tabla 3: Discriminación por Conceptos (Costos)
        st.subheader("Diagnóstico detallado por concepto (Costos)")
        concept_data_costos = []
        for concept, ids in normalized_concepts.items():
            # Filtrar los datos para los IDs del concepto
            matching_mask = df_filtered['Nombre o ID'].isin(ids)
            concept_df = df_filtered[matching_mask]
            
            # Desglose: IDs encontrados y sus sumas (para costos)
            found_ids = concept_df['Nombre o ID'].unique().tolist()
            st.write(f"**{concept}:** IDs buscados: {ids}")
            st.write(f"IDs encontrados: {found_ids} (de {len(found_ids)}/{len(ids)})")
            
            if len(concept_df) > 0:
                # Mostrar desglose por ID para 'Costo del servicio'
                breakdown_costo = concept_df.groupby('Nombre o ID')['Costo del servicio'].sum().round(2)
                st.write("Desglose por ID (Costo del servicio):", breakdown_costo.to_dict())
                
                # Mostrar desglose por ID para 'IVA del costo'
                breakdown_iva = concept_df.groupby('Nombre o ID')['IVA del costo'].sum().round(2)
                st.write("Desglose por ID (IVA del costo):", breakdown_iva.to_dict())
            else:
                st.warning(f"No se encontraron datos para {concept}. Verifica los nombres de IDs.")
            
            # Sumar 'Costo del servicio' y 'IVA del costo'
            costo_servicio_total = concept_df['Costo del servicio'].sum()
            iva_costo_total = concept_df['IVA del costo'].sum()
            # Calcular 'Costo total'
            costo_total = costo_servicio_total + iva_costo_total
            concept_data_costos.append([concept, costo_servicio_total, iva_costo_total, costo_total])
        
        # Crear DataFrame para la tabla de costos
        concepts_df_costos = pd.DataFrame(concept_data_costos, columns=['Concepto', 'Costo del servicio', 'IVA del costo', 'Costo total'])
        
        # Calcular la fila de totales para la tabla de costos
        totals_costos = concepts_df_costos[['Costo del servicio', 'IVA del costo', 'Costo total']].sum()
        totals_costos_row = pd.DataFrame([['Total', *totals_costos]], columns=['Concepto', 'Costo del servicio', 'IVA del costo', 'Costo total'])
        concepts_df_costos_with_totals = pd.concat([concepts_df_costos, totals_costos_row], ignore_index=True)
        
        # Formatear las columnas numéricas a 2 decimales
        concepts_df_costos_with_totals['Costo del servicio'] = concepts_df_costos_with_totals['Costo del servicio'].round(2)
        concepts_df_costos_with_totals['IVA del costo'] = concepts_df_costos_with_totals['IVA del costo'].round(2)
        concepts_df_costos_with_totals['Costo total'] = concepts_df_costos_with_totals['Costo total'].round(2)
        
        # Visualizar Tabla 3 (Costos)
        st.subheader("Discriminación por Conceptos (Costos)")
        st.dataframe(concepts_df_costos_with_totals)
        
        # Preparar el archivo Excel para exportar las tres tablas
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            grouped_with_totals.to_excel(writer, index=False, sheet_name='Resultados')
            concepts_df_neto_with_totals.to_excel(writer, index=False, sheet_name='Conceptos Netos')
            concepts_df_costos_with_totals.to_excel(writer, index=False, sheet_name='Conceptos Costos')
        output.seek(0)
        
        # Botón para descargar el Excel procesado
        st.download_button(
            label="Descargar resultados en Excel",
            data=output,
            file_name="resultados_procesados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("Por favor, carga un archivo Excel para continuar.")