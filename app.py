import streamlit as st
import pandas as pd

# ... (Funciones de limpieza previas se mantienen igual) ...

if st.button("Ejecutar Análisis ABC y Generar Muestra"):
    # 1. Asegurar que los datos sean numéricos
    df[col_stock] = pd.to_numeric(df[col_stock], errors='coerce').fillna(0)
    df[col_costo] = pd.to_numeric(df[col_costo], errors='coerce').fillna(0)
    
    # 2. Cálculo del ABC
    df['Valor_Total'] = df[col_stock] * df[col_costo]
    df = df.sort_values(by='Valor_Total', ascending=False)
    df['Pct_Acumulado'] = df['Valor_Total'].cumsum() / df['Valor_Total'].sum()
    
    def asignar_abc(pct):
        if pct <= 0.80: return 'A'
        elif pct <= 0.95: return 'B'
        else: return 'C'
    
    df['Categoria'] = df['Pct_Acumulado'].apply(asignar_abc)

    # 3. GENERACIÓN DE MUESTRA SEGÚN SOLICITUD (85A, 10B, 5C)
    # Usamos min() por seguridad en caso de que una categoría tenga menos artículos de los pedidos
    m_a = df[df['Categoria'] == 'A'].sample(n=min(85, len(df[df['Categoria'] == 'A'])))
    m_b = df[df['Categoria'] == 'B'].sample(n=min(10, len(df[df['Categoria'] == 'B'])))
    m_c = df[df['Categoria'] == 'C'].sample(n=min(5, len(df[df['Categoria'] == 'C'])))
    
    muestra_final = pd.concat([m_a, m_b, m_c])

    # 4. VISUALIZACIÓN
    st.success(f"Muestra generada con éxito: {len(muestra_final)} artículos seleccionados.")
    
    # Definimos las columnas que queremos ver (incluyendo Locación y Descripción)
    # Buscamos los nombres de columnas que mapeaste anteriormente
    columnas_a_mostrar = [col_art, 'Descripción', 'Locación', col_stock, 'Categoria']
    
    # Verificamos que existan en el df para no tener errores de visualización
    cols_existentes = [c for c in columnas_a_mostrar if c in muestra_final.columns]
    
    st.dataframe(muestra_final[cols_existentes])

    # Guardar en el estado de la sesión para las siguientes pestañas
    st.session_state['muestra_completa'] = muestra_final
