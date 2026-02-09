import streamlit as st
import pandas as pd

def limpiar_datos_cenoa(df_raw):
    # 1. Buscamos la fila exacta donde están los títulos reales
    # En tu archivo de Jujuy, los títulos reales están donde aparece 'Artículo'
    for i in range(len(df_raw)):
        fila_actual = df_raw.iloc[i].astype(str).tolist()
        if 'Artículo' in fila_actual or 'Articulo' in fila_actual:
            df_limpio = df_raw.iloc[i+1:].copy()
            df_limpio.columns = fila_actual
            return df_limpio.reset_index(drop=True)
    return df_raw

st.title("📦 Auditoría Interna - Grupo Cenoa")

archivo = st.file_uploader("Subir Reporte de Stock Jujuy", type=['xlsx'])

if archivo:
    # Leemos sin encabezados inicialmente para no perder ninguna fila
    df_input = pd.read_excel(archivo, header=None)
    df = limpiar_datos_cenoa(df_input)
    
    # Mapeo de columnas basado en tu archivo real
    # Usamos nombres exactos detectados: 'Locación', 'Artículo', 'Descripción', 'Stock', 'Cto.Rep.'
    col_art = 'Artículo'
    col_loc = 'Locación'
    col_desc = 'Descripción'
    col_stock = 'Stock'
    col_costo = 'Cto.Rep.'

    if st.button("Ejecutar Análisis y Muestra"):
        # Limpieza de números (importante para evitar errores de cálculo)
        df[col_stock] = pd.to_numeric(df[col_stock], errors='coerce').fillna(0)
        df[col_costo] = pd.to_numeric(df[col_costo], errors='coerce').fillna(0)
        
        # Lógica ABC
        df['Valor_Total'] = df[col_stock] * df[col_costo]
        df = df.sort_values(by='Valor_Total', ascending=False)
        df['Pct_Acumulado'] = df['Valor_Total'].cumsum() / df['Valor_Total'].sum()
        
        def categorizar(pct):
            if pct <= 0.80: return 'A'
            elif pct <= 0.95: return 'B'
            else: return 'C'
        
        df['Categoria'] = df['Pct_Acumulado'].apply(categorizar)

        # MUESTRA SOLICITADA: 85A, 10B, 5C
        m_a = df[df['Categoria'] == 'A'].sample(n=min(85, len(df[df['Categoria'] == 'A'])))
        m_b = df[df['Categoria'] == 'B'].sample(n=min(10, len(df[df['Categoria'] == 'B'])))
        m_c = df[df['Categoria'] == 'C'].sample(n=min(5, len(df[df['Categoria'] == 'C'])))
        
        muestra_final = pd.concat([m_a, m_b, m_c])

        st.success(f"Muestra generada: {len(muestra_final)} artículos")
        
        # Mostramos la tabla con las columnas que pediste
        columnas_visibles = [col_loc, col_art, col_desc, col_stock, 'Categoria']
        st.dataframe(muestra_final[columnas_visibles])
        
        # Guardamos en sesión para el siguiente paso (Conteo)
        st.session_state['muestra_final'] = muestra_final
