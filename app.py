import streamlit as st
import pandas as pd
import numpy as np

# Configuración de página
st.set_page_config(page_title="Auditoría Grupo Cenoa", layout="wide")

def limpiar_y_detectar_columnas(df):
    """Busca la fila que contiene los encabezados y limpia el DataFrame"""
    # Buscamos la fila que contenga la palabra 'Artículo' o 'Stock'
    for i in range(len(df)):
        fila = df.iloc[i].astype(str).tolist()
        if any('Articulo' in x or 'Artículo' in x or 'Stock' in x for x in fila):
            df.columns = fila
            df = df.iloc[i+1:].reset_index(drop=True)
            break
    return df

def clasificar_abc(df, col_stock, col_costo):
    # Convertir a numérico por seguridad
    df[col_stock] = pd.to_numeric(df[col_stock], errors='coerce').fillna(0)
    df[col_costo] = pd.to_numeric(df[col_costo], errors='coerce').fillna(0)
    
    # Calcular Valor Total
    df['Valor_Total'] = df[col_stock] * df[col_costo]
    df = df.sort_values(by='Valor_Total', ascending=False)
    
    # Calcular % Acumulado
    df['Pct_Acumulado'] = df['Valor_Total'].cumsum() / df['Valor_Total'].sum()
    
    def categorizar(pct):
        if pct <= 0.80: return 'A'
        elif pct <= 0.95: return 'B'
        else: return 'C'
        
    df['Categoria'] = df['Pct_Acumulado'].apply(categorizar)
    return df

st.title("📦 Control de Inventarios Rotativos - Grupo Cenoa")

# Pestañas para las etapas del proceso
tab1, tab2, tab3, tab4 = st.tabs(["1. Carga y ABC", "2. Conteo Físico", "3. Justificaciones", "4. Informe Final"])

with tab1:
    archivo = st.file_uploader("Subir Reporte de Stock", type=['xlsx', 'csv'])
    
    if archivo:
        # Carga inicial (leemos todo como texto para no perder datos en la limpieza)
        raw_df = pd.read_excel(archivo) if archivo.name.endswith('xlsx') else pd.read_csv(archivo)
        
        # Limpieza automática de encabezados
        df_limpio = limpiar_y_detectar_columnas(raw_df)
        
        st.subheader("Configuración de Columnas")
        col1, col2, col3, col4 = st.columns(4)
        
        # Selectores flexibles: El auditor elige qué columna es cual
        columnas_disponibles = df_limpio.columns.tolist()
        
        # Intentamos pre-seleccionar si coinciden los nombres
        with col1:
            c_art = st.selectbox("Columna de Artículo", columnas_disponibles, 
                                 index=columnas_disponibles.index('Artículo') if 'Artículo' in columnas_disponibles else 0)
        with col2:
            c_loc = st.selectbox("Columna de Ubicación/Locación", columnas_disponibles,
                                 index=columnas_disponibles.index('Locación') if 'Locación' in columnas_disponibles else 0)
        with col3:
            c_stock = st.selectbox("Columna de Stock Sistema", columnas_disponibles,
                                   index=columnas_disponibles.index('Stock') if 'Stock' in columnas_disponibles else 0)
        with col4:
            c_costo = st.selectbox("Columna de Costo Reposición", columnas_disponibles,
                                   index=columnas_disponibles.index('Cto.Rep.') if 'Cto.Rep.' in columnas_disponibles else 0)

        if st.button("Generar Clasificación ABC y Muestra"):
            # Procesamos el ABC
            df_abc = clasificar_abc(df_limpio, c_stock, c_costo)
            
            # Selección aleatoria (80A, 15B, 5C)
            m_a = df_abc[df_abc['Categoria'] == 'A'].sample(n=min(80, len(df_abc[df_abc['Categoria'] == 'A'])))
            m_b = df_abc[df_abc['Categoria'] == 'B'].sample(n=min(15, len(df_abc[df_abc['Categoria'] == 'B'])))
            m_c = df_abc[df_abc['Categoria'] == 'C'].sample(n=min(5, len(df_abc[df_abc['Categoria'] == 'C'])))
            
            muestra_final = pd.concat([m_a, m_b, m_c])
            
            # Guardamos en la sesión de Streamlit (luego lo conectaremos a la base de datos)
            st.session_state['muestra'] = muestra_final
            st.success(f"Muestra generada: {len(muestra_final)} artículos.")
            st.dataframe(muestra_final[[c_art, c_loc, c_stock, 'Categoria', 'Valor_Total']])

with tab2:
    st.header("Toma de Inventario")
    if 'muestra' in st.session_state:
        st.write("Cargue los resultados del conteo físico abajo:")
        # Aquí se implementará la tabla editable
        df_conteo = st.data_editor(st.session_state['muestra'], 
                                   column_order=(c_art, c_loc, c_stock, "Conteo_Fisico"),
                                   num_rows="fixed")
    else:
        st.warning("Primero genera la muestra en la pestaña 1.")
